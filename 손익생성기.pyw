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
        ttk.Label(top, text="매장 선택 (복수 선택 가능)").pack(anchor="w")
        self.lb = tk.Listbox(top, selectmode="extended", height=10, exportselection=False)
        for s in self.stores:
            self.lb.insert("end", s["name"])
        self.lb.pack(fill="x")
        ttk.Button(top, text="전체 선택", command=self._select_all).pack(anchor="e", pady=2)

        opt = ttk.Frame(root); opt.pack(fill="x", **pad)
        ym = date.today()
        y, m = (ym.year, ym.month - 1) if ym.month > 1 else (ym.year - 1, 12)
        ttk.Label(opt, text="대상 월 (YYYY-MM):").grid(row=0, column=0, sticky="w")
        self.month = ttk.Entry(opt, width=12); self.month.insert(0, f"{y}-{m:02d}")
        self.month.grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(opt, text="가게배달 건수(선택):").grid(row=0, column=2, sticky="w")
        self.gagae = ttk.Entry(opt, width=8)
        self.gagae.grid(row=0, column=3, sticky="w", padx=6)

        self.fetch = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="토더 자동수집(Playwright)", variable=self.fetch)\
            .grid(row=1, column=0, columnspan=2, sticky="w", pady=4)

        self.run_btn = ttk.Button(root, text="손익 보고서 생성", command=self._run)
        self.run_btn.pack(**pad)

        ttk.Label(root, text="진행 로그").pack(anchor="w", padx=8)
        self.log = scrolledtext.ScrolledText(root, height=16)
        self.log.pack(fill="both", expand=True, padx=8, pady=4)

    def _select_all(self):
        self.lb.select_set(0, "end")

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

        self.run_btn.config(state="disabled")
        self.log.delete("1.0", "end")
        threading.Thread(target=self._worker,
                         args=(sel, year, month, gagae_val, fetch), daemon=True).start()

    def _worker(self, names, year, month, gagae_val, fetch):
        old = sys.stdout
        sys.stdout = TextRedirector(self.log)
        done, fail = [], []
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
