# -*- coding: utf-8 -*-
"""
DBC 影子中台 桌面啟動器
─────────────────────────────────
按鍵點擊執行常用作業，免再開 CMD 貼指令
"""

import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import scrolledtext, font as tkfont
from datetime import datetime
import webbrowser

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
PYTHON       = sys.executable   # 用同一個 python.exe
PORTAL_URL   = 'https://dbcjimmyoklay.github.io/dbc-base/portal/'
DBC_BASE_URL = 'https://docs.google.com/spreadsheets/d/12p71jgMErzZYO4toU2LVELDPfwklHCyDARclldYImCc/'
COACH_VER_PATH = os.path.join(BASE_DIR, 'portal', 'coach_version.json')

# ── 顏色（與 Portal 一致）──
BG0  = '#000b18'
BG1  = '#001529'
BG2  = '#001f3d'
BG3  = '#002a52'
BD   = '#0a3260'
BD2  = '#0d3d75'
TX   = '#ffffff'
TX2  = '#8aadcc'
TX3  = '#4a7299'
AC   = '#1a6fb5'
AC2  = '#2489d6'
OK_C = '#10b981'
NG_C = '#ef4444'


class Launcher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('DBC 影子中台 啟動器')
        self.geometry('640x600')
        self.configure(bg=BG0)
        self.minsize(560, 520)

        # 字型
        self.f_title = tkfont.Font(family='Microsoft JhengHei UI', size=14, weight='bold')
        self.f_sub   = tkfont.Font(family='Microsoft JhengHei UI', size=10)
        self.f_btn   = tkfont.Font(family='Microsoft JhengHei UI', size=11, weight='bold')
        self.f_log   = tkfont.Font(family='Consolas', size=9)

        self.busy = False
        self._build_ui()

    def _build_ui(self):
        # Title bar
        top = tk.Frame(self, bg=BG1, height=58)
        top.pack(side='top', fill='x')
        top.pack_propagate(False)
        tk.Label(top, text='🏔  DBC 影子中台', font=self.f_title,
                 bg=BG1, fg=TX).pack(side='left', padx=18, pady=12)
        tk.Label(top, text='啟動器', font=self.f_sub,
                 bg=BG1, fg=TX3).pack(side='left', pady=14)

        # Buttons section
        btns = tk.Frame(self, bg=BG0)
        btns.pack(side='top', fill='x', padx=14, pady=10)
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)

        self._mkbtn(btns, '🚀  跑每日流程',
                    '清洗 → 寫入 Google Sheets → 銷量分析 → 推 Portal',
                    self.run_daily, 0, 0, colspan=2, primary=True)

        self._mkbtn(btns, '🌐  開啟 Portal',
                    '在瀏覽器開啟 DBC Portal',
                    self.open_portal, 1, 0)

        self._mkbtn(btns, '📊  開啟 DBC Base',
                    '開啟主 Google 試算表',
                    self.open_dbc_base, 1, 1)

        self._mkbtn(btns, '🎿  班表更新',
                    '通知 Portal 強制刷新教練班表',
                    self.publish_coach, 2, 0)

        self._mkbtn(btns, '🏨  訂房表更新',
                    '比對野澤訂房表 → 更新入住單與課程安排',
                    self.run_checkin, 2, 1)

        # Status
        self.status = tk.Label(self, text='閒置 ✓', font=self.f_sub,
                               bg=BG0, fg=TX3, anchor='w')
        self.status.pack(side='top', fill='x', padx=16, pady=(4, 2))

        # Log area
        log_frame = tk.Frame(self, bg=BG0)
        log_frame.pack(side='top', fill='both', expand=True, padx=14, pady=(0, 12))

        tk.Label(log_frame, text='輸出', font=self.f_sub,
                 bg=BG0, fg=TX3, anchor='w').pack(side='top', fill='x', pady=(0, 4))

        self.log = scrolledtext.ScrolledText(log_frame, font=self.f_log,
                                             bg='#000610', fg=TX2,
                                             insertbackground=TX,
                                             relief='flat', bd=1,
                                             highlightthickness=1,
                                             highlightbackground=BD,
                                             highlightcolor=BD2,
                                             wrap='word')
        self.log.pack(side='top', fill='both', expand=True)
        self.log.config(state='disabled')

        # Tags
        self.log.tag_config('ok',  foreground=OK_C)
        self.log.tag_config('ng',  foreground=NG_C)
        self.log.tag_config('hdr', foreground=AC2)
        self.log.tag_config('dim', foreground=TX3)

        self._log('=== DBC 啟動器準備就緒 ===\n', 'hdr')

    def _mkbtn(self, parent, title, sub, cmd, r, c, colspan=1, primary=False):
        f = tk.Frame(parent, bg=BG2 if not primary else AC,
                     highlightthickness=1,
                     highlightbackground=BD2 if not primary else AC2,
                     cursor='hand2')
        f.grid(row=r, column=c, columnspan=colspan, sticky='nsew',
               padx=4, pady=4, ipady=8)
        title_lbl = tk.Label(f, text=title, font=self.f_btn,
                             bg=BG2 if not primary else AC, fg=TX,
                             cursor='hand2')
        title_lbl.pack(side='top', padx=12, pady=(8, 2))
        sub_lbl = tk.Label(f, text=sub, font=self.f_sub,
                           bg=BG2 if not primary else AC,
                           fg=TX2 if not primary else '#cfe8ff',
                           cursor='hand2')
        sub_lbl.pack(side='top', padx=12, pady=(0, 6))
        for w in (f, title_lbl, sub_lbl):
            w.bind('<Button-1>', lambda e, c=cmd: c())
            w.bind('<Enter>',    lambda e, fr=f, p=primary: self._hover(fr, True, p))
            w.bind('<Leave>',    lambda e, fr=f, p=primary: self._hover(fr, False, p))

    def _hover(self, frame, on, primary):
        if self.busy: return
        if primary:
            color = AC2 if on else AC
        else:
            color = BG3 if on else BG2
        frame.config(bg=color)
        for child in frame.winfo_children():
            child.config(bg=color)

    def _log(self, txt, tag=None):
        self.log.config(state='normal')
        if tag: self.log.insert('end', txt, tag)
        else:   self.log.insert('end', txt)
        self.log.see('end')
        self.log.config(state='disabled')
        self.update_idletasks()

    def _set_status(self, txt, color=TX3):
        self.status.config(text=txt, fg=color)
        self.update_idletasks()

    def _run_cmd_async(self, label, args, cwd=BASE_DIR, then=None, final=False):
        """final=True 時整個流程結束顯示「全部完成」"""
        if self.busy:
            self._log('⚠ 已有任務執行中，請稍候\n', 'ng')
            return
        self.busy = True
        self._set_status(f'⏳ 執行中：{label}', AC2)
        self._log(f'\n─── {label} ─── {datetime.now().strftime("%H:%M:%S")}\n', 'hdr')

        def worker():
            try:
                # ── Windows：用 UTF-8 強迫 Python 子程序輸出 utf-8 ──
                env = os.environ.copy()
                env['PYTHONIOENCODING'] = 'utf-8'
                env['PYTHONUTF8']       = '1'

                if os.name == 'nt':
                    si = subprocess.STARTUPINFO()
                    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                else:
                    si = None

                proc = subprocess.Popen(
                    args, cwd=cwd, env=env,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding='utf-8', errors='replace',
                    startupinfo=si, bufsize=1,
                )
                for line in proc.stdout:
                    self._log(line)
                proc.wait()
                ok = (proc.returncode == 0)
                if ok:
                    self._log(f'✓ {label} 完成\n', 'ok')
                    self._set_status(f'✓ {label} 完成', OK_C)
                else:
                    self._log(f'✗ {label} 失敗（exit code {proc.returncode}）\n', 'ng')
                    self._set_status(f'✗ {label} 失敗', NG_C)

                # 最後一步：全部完成標記
                if final and ok:
                    self._log('\n' + '═' * 40 + '\n', 'hdr')
                    self._log(f'🎉 全部完成！{datetime.now().strftime("%H:%M:%S")}\n', 'ok')
                    self._log('═' * 40 + '\n', 'hdr')
                    self._set_status('🎉 全部完成', OK_C)
                elif final and not ok:
                    self._log('\n' + '═' * 40 + '\n', 'ng')
                    self._log('⚠️ 流程中止（有步驟失敗）\n', 'ng')
                    self._log('═' * 40 + '\n', 'ng')
            except Exception as e:
                self._log(f'✗ 例外：{e}\n', 'ng')
                self._set_status('✗ 執行例外', NG_C)
                ok = False

            # 先釋放 busy，再呼叫 then；下一步的 _run_cmd_async 才能正常啟動
            self.busy = False
            if then: then(success=ok)

        threading.Thread(target=worker, daemon=True).start()

    # ── 各按鈕對應的動作 ──

    def _chain_push(self, msg):
        """git add → commit → push 的鏈式執行，最後標記 final"""
        def do_push(success):
            if not success:
                return
            self._run_cmd_async('git push', ['git', 'push'], final=True)
        def do_commit(success):
            if not success:
                return
            self._run_cmd_async('git commit',
                ['git', 'commit', '-m', msg], then=do_push)
        self._run_cmd_async('git add', ['git', 'add', '-A'], then=do_commit)

    def run_daily(self):
        """跑每日流程 → 自動 push"""
        def after_main(success):
            if not success:
                # 報告失敗
                self._set_status('✗ main.py 失敗，停止', NG_C)
                return
            self._chain_push(f'每日資料更新 {datetime.now():%Y-%m-%d %H:%M}')
        self._run_cmd_async('main.py --all',
            [PYTHON, 'main.py', '--all'], then=after_main)

    def run_checkin(self):
        self._run_cmd_async('main.py --checkin',
            [PYTHON, 'main.py', '--checkin'], final=True)

    def open_portal(self):
        webbrowser.open(PORTAL_URL)
        self._log(f'🌐 已開啟 Portal\n', 'dim')

    def open_dbc_base(self):
        webbrowser.open(DBC_BASE_URL)
        self._log(f'📊 已開啟 DBC Base 試算表\n', 'dim')

    def publish_coach(self):
        """產生新的版本標記，git push 後 Portal 端會自動清快取重抓"""
        import json
        ts = datetime.now().strftime('%Y%m%d-%H%M%S')
        try:
            os.makedirs(os.path.dirname(COACH_VER_PATH), exist_ok=True)
            with open(COACH_VER_PATH, 'w', encoding='utf-8') as f:
                json.dump({'version': ts, 'updated': datetime.now().isoformat(timespec='minutes')},
                          f, ensure_ascii=False)
            self._log(f'✓ 已產生新版本標記 {ts}\n', 'ok')
        except Exception as e:
            self._log(f'✗ 無法寫入 coach_version.json：{e}\n', 'ng')
            return
        self._chain_push(f'班表更新 {ts}')


if __name__ == '__main__':
    app = Launcher()
    app.mainloop()
