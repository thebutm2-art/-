# -*- coding: utf-8 -*-
"""
손익 보고서 생성기 (GUI)
- 매장 선택 -> 월/가게배달건수 지정 -> 손익 보고서 생성 -> 바탕화면/{매장}/{YYMM 매장}.xlsx 저장
- 토더 데이터는 '토더 자동수집' 체크 시 Playwright로 자동 추출
- 배민 정산서(downloads 폴더), 쿠팡 정산내역서, 식부자재_{월}월.xlsx 는 미리 준비되어 있어야 함

실행: 이 파일 더블클릭 (또는 pythonw 손익생성기.pyw)
"""
import sys, io, os, threading, traceback
from pathlib import Path
from datetime import date

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

BASE = Path(__file__).parent
os.chdir(BASE)

from store_master import load_stores, MASTER_V2
import run_report


class TextRedirector(io.TextIOBase):
    """rich/print 출력을 텍스트 위젯으로 리다이렉트."""
    def __init__(self, widget):
        self.widget = widget
    def write(self, s):
        try:
            self.widget.after(0, lambda: (self.widget.insert("end", s),
                                          self.widget.see("end")))
        except Exception:
            pass
        return len(s)
    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        root.title("배달 손익 보고서 생성기")
        root.geometry("720x620")

        try:
            self.stores = load_stores()
        except Exception as e:
            messagebox.showerror("마스터 오류",
                                 f"점포마스터_v2.xlsx 로딩 실패:\n{e}\n\n"
                                 f"import_accounts.py 로 마스터를 먼저 생성하세요.")
            self.stores = []
        self.by_name = {s["name"]: s for s in self.stores}

        pad = {"padx": 8, "pady": 4}

        top = ttk.Frame(root); top.pack(fill="x", **pad)
        ttk.Label(top, text="① 매장 선택 (한 개만 선택하면 아래 매장정보 편집 가능)").pack(anchor="w")
        self.lb = tk.Listbox(top, selectmode="extended", height=8, exportselection=False)
        for s in self.stores:
            self.lb.insert("end", s["name"])
        self.lb.pack(fill="x")
        self.lb.bind("<<ListboxSelect>>", self._on_select)
        ttk.Button(top, text="전체 선택", command=self._select_all).pack(anchor="e", pady=2)

        # ── 선택 매장 정보(배민파트너명·파일암호) — 마스터에 저장됨 ──
        info = ttk.LabelFrame(root, text="② 선택 매장 정보 (정산서용 — 입력하면 마스터에 저장)")
        info.pack(fill="x", padx=8, pady=4)
        ttk.Label(info, text="배민 파트너명:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.partner = ttk.Entry(info, width=16)
        self.partner.grid(row=0, column=1, sticky="w")
        ttk.Label(info, text="(정산명세서의 파트너 이름, 예: 장승환)", foreground="#888")\
            .grid(row=0, column=2, sticky="w", padx=6)
        ttk.Label(info, text="파일암호:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.filepw = ttk.Entry(info, width=16)
        self.filepw.grid(row=1, column=1, sticky="w")
        ttk.Label(info, text="(정산서 엑셀 잠금해제용 — 생년월일6자리 / 법인번호 뒷7자리)",
                  foreground="#888").grid(row=1, column=2, sticky="w", padx=6)

        opt = ttk.Frame(root); opt.pack(fill="x", **pad)
        ym = date.today()
        y, m = (ym.year, ym.month - 1) if ym.month > 1 else (ym.year - 1, 12)
        ttk.Label(opt, text="③ 대상 월 (YYYY-MM):").grid(row=0, column=0, sticky="w")
        self.month = ttk.Entry(opt, width=12); self.month.insert(0, f"{y}-{m:02d}")
        self.month.grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(opt, text="가게배달 건수(선택):").grid(row=0, column=2, sticky="w")
        self.gagae = ttk.Entry(opt, width=8)
        self.gagae.grid(row=0, column=3, sticky="w", padx=6)

        # Gmail 앱 비밀번호 (메일 수신용) — 파일암호와 다름!
        ttk.Label(opt, text="④ Gmail 앱 비밀번호:").grid(row=1, column=0, sticky="w", pady=6)
        self.mailpw = ttk.Entry(opt, width=24, show="●")
        self.mailpw.grid(row=1, column=1, columnspan=2, sticky="w", padx=6)
        ttk.Label(opt, text="(정산서 메일 수신용 16자리 · 파일암호와 다름)", foreground="#c00")\
            .grid(row=1, column=3, sticky="w")

        self.dosend = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="배민셀프 정산명세서 자동발송(메일)", variable=self.dosend)\
            .grid(row=2, column=0, columnspan=4, sticky="w", pady=2)
        self.fetch = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="자동수집(배민 메일 수신 + 토더/쿠팡)", variable=self.fetch)\
            .grid(row=3, column=0, columnspan=4, sticky="w", pady=2)

        self.run_btn = ttk.Button(root, text="손익 보고서 생성", command=self._run)
        self.run_btn.pack(**pad)

        ttk.Label(root, text="진행 로그").pack(anchor="w", padx=8)
        self.log = scrolledtext.ScrolledText(root, height=16)
        self.log.pack(fill="both", expand=True, padx=8, pady=4)

    def _select_all(self):
        self.lb.select_set(0, "end")
        self._on_select()

    def _on_select(self, event=None):
        """한 개만 선택 시 그 매장의 파트너명/파일암호를 불러와 편집 가능."""
        sel = self.lb.curselection()
        self.partner.delete(0, "end"); self.filepw.delete(0, "end")
        if len(sel) == 1:
            s = self.by_name[self.lb.get(sel[0])]
            self.partner.insert(0, s.get("partner", "") or "")
            self.filepw.insert(0, s.get("file_pw", "") or "")
            self.partner.config(state="normal"); self.filepw.config(state="normal")
        else:
            self.partner.config(state="disabled"); self.filepw.config(state="disabled")

    def _save_store_info(self, name, partner, filepw):
        """배민파트너명/파일암호를 마스터(점포마스터_v2.xlsx)에 저장 + 메모리 갱신."""
        import openpyxl
        from store_master import MASTER_V2
        wb = openpyxl.load_workbook(MASTER_V2); ws = wb.active
        hdr = {c.value: i + 1 for i, c in enumerate(ws[1])}
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, hdr["점포명"]).value == name:
                if "배민파트너명" in hdr: ws.cell(r, hdr["배민파트너명"]).value = partner
                if "파일암호" in hdr:     ws.cell(r, hdr["파일암호"]).value = filepw
                break
        wb.save(MASTER_V2)
        self.by_name[name]["partner"] = partner
        self.by_name[name]["file_pw"] = filepw

    def _run(self):
        sel = [self.lb.get(i) for i in self.lb.curselection()]
        if not sel:
            messagebox.showwarning("선택 없음", "매장을 1개 이상 선택하세요.")
            return
        try:
            year, month = map(int, self.month.get().strip().split("-"))
        except Exception:
            messagebox.showerror("월 오류", "대상 월을 YYYY-MM 형식으로 입력하세요.")
            return
        gagae = self.gagae.get().strip()
        gagae_val = int(gagae) if gagae.isdigit() else None
        fetch = self.fetch.get()
        mail_pw = self.mailpw.get().strip()
        if fetch and not mail_pw:
            if not messagebox.askyesno("메일 암호 없음",
                    "메일 암호가 비어 있어 배민 정산서를 자동 수신하지 못합니다.\n"
                    "downloads 폴더에 정산서가 이미 있다면 계속 진행할 수 있습니다.\n계속할까요?"):
                return

        # 단일 매장 선택 시, 입력한 파트너명/파일암호를 마스터에 저장
        if len(sel) == 1:
            p = self.partner.get().strip(); f = self.filepw.get().strip().replace("-", "")
            if p or f:
                try:
                    self._save_store_info(sel[0], p, f)
                except Exception as e:
                    messagebox.showwarning("저장 경고", f"매장정보 저장 실패: {e}")

        do_send = self.dosend.get()
        self.run_btn.config(state="disabled")
        self.log.delete("1.0", "end")
        threading.Thread(target=self._worker,
                         args=(sel, year, month, gagae_val, fetch, mail_pw, do_send),
                         daemon=True).start()

    def _worker(self, names, year, month, gagae_val, fetch, mail_pw, do_send):
        import time, asyncio
        import config
        old = sys.stdout
        sys.stdout = TextRedirector(self.log)
        done, fail = [], []
        sel_stores = [self.by_name[n] for n in names]

        # ① 배민셀프 정산명세서 자동발송 (메일)
        if do_send:
            from step_baemin_self_send import send_settlement_mail
            for s in sel_stores:
                print(f"[{s['name']}] 배민셀프 정산명세서 발송 중...")
                try:
                    asyncio.run(send_settlement_mail(s, year, month, config.MAIL_USER))
                except Exception as e:
                    print(f"  발송 실패: {repr(e)[:100]}")

        # ② 배민 정산서 Gmail 수신 (폴링: 메일 도착까지 최대 ~8분)
        if mail_pw:
            from step_baemin_mail import fetch_all as fetch_mail
            print("배민 정산서 메일 수신 대기 중... (최대 8분)")
            need = {s["name"] for s in sel_stores if s.get("partner")}
            got = set()
            for attempt in range(16):
                try:
                    saved = fetch_mail(sel_stores, year, month, mail_pass=mail_pw)
                    got |= set(saved.keys())
                except Exception as e:
                    print(f"  수신 시도 실패: {repr(e)[:80]}")
                if need and need <= got:
                    break
                if attempt < 15:
                    time.sleep(30)
        try:
            sik_bu = run_report._load_sikbu(month)
        except Exception as e:
            print(f"[오류] 식부자재_{month}월.xlsx 필요: {e}")
            sys.stdout = old
            self._finish(done, fail)
            return
        for name in names:
            store = dict(self.by_name[name])
            if gagae_val is not None:
                store["gagae_count"] = gagae_val
            try:
                r = run_report.process_store(store, year, month, sik_bu, fetch)
                (done if r else fail).append(name)
            except Exception as e:
                print(f"[오류] {name}: {e}")
                traceback.print_exc()
                fail.append(name)
        sys.stdout = old
        self._finish(done, fail)

    def _finish(self, done, fail):
        def ui():
            self.run_btn.config(state="normal")
            msg = f"완료: {len(done)}개 성공"
            if fail:
                msg += f", {len(fail)}개 실패\n실패: {', '.join(fail)}"
            if done:
                msg += f"\n\n저장 위치: 바탕화면\\{{매장명}}\\"
            messagebox.showinfo("생성 완료", msg)
        self.root.after(0, ui)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
