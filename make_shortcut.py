"""
바탕화면에 '배달 손익 생성기' 바로가기(.lnk) 생성 (Windows 전용)
실행: python make_shortcut.py
"""
import sys, os, subprocess
from pathlib import Path

BASE = Path(__file__).resolve().parent


def find_pythonw() -> str:
    # 같은 파이썬의 pythonw.exe 우선
    cand = Path(sys.executable).with_name("pythonw.exe")
    if cand.exists():
        return str(cand)
    return sys.executable


def main():
    pyw = find_pythonw()
    desktop = Path(os.path.join(os.path.expanduser("~"), "Desktop"))
    lnk = desktop / "배달 손익 생성기.lnk"
    ps = f'''
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut("{lnk}")
$sc.TargetPath = "{pyw}"
$sc.Arguments = '"손익생성기.pyw"'
$sc.WorkingDirectory = "{BASE}"
$sc.IconLocation = "shell32.dll,21"
$sc.Description = "배달 손익 보고서 생성기"
$sc.Save()
'''
    subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True)
    print(f"[OK] 바로가기 생성: {lnk}")
    print(f"     실행기: {pyw}")
    print(f"     작업폴더: {BASE}")


if __name__ == "__main__":
    main()
