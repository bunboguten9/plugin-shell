# -*- coding: utf-8 -*-
"""
app_shell.py

後からプラグインを追加したいときの想定:
  - ./plugins/ ディレクトリにプラグインを配置
  - プラグインは "class Plugin(PluginBase)" を1つエクスポートする（名前は 'Plugin' 固定）
  - アプリ内の「プラグインを再読み込み」から反映
"""
import os
import sys
import importlib.util
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
import time
import platform
import webbrowser
import subprocess

APP_NAME = "Plugin Shell"
APP_DESC = "モード切替型プラグイン・アプリケーションのシェル"
PLUGINS_DIRNAME = "plugins"

# ---------------------------
# プラグイン基底クラス（仕様）
# ---------------------------
class PluginBase:
    """
    すべてのプラグインはこの基底クラスを継承し、以下を実装/設定する想定。
      - name:     一意の表示名（左リボンボタン・ヘッダーに表示）
      - icon:     絵文字や短いテキスト（左ボタンに表示）
      - mount(parent): 選択されたとき、作業領域 parent(Frame) にUIを構築する
      - unmount(): プラグインが切り替わる際に呼ばれる（後始末）
    """
    name: str = "Unnamed"
    icon: str = "🔧"

    def __init__(self, shell_context: dict | None = None) -> None:
        self.shell_context = shell_context or {}

    def mount(self, parent: tk.Frame) -> None:
        raise NotImplementedError

    def unmount(self) -> None:
        pass


# ---------------------------
# シェル（UIフレーム）
# ---------------------------
class ShellApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("1100x720")
        self.root.minsize(980, 640)
        self.root.configure(bg="#eef1f7")

        # パス・状態
        self.base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
        self.plugins_dir = self.base_dir / PLUGINS_DIRNAME
        self.plugins_dir.mkdir(exist_ok=True)
        self.plugins: list[PluginBase] = []
        self.current_plugin: PluginBase | None = None

        # スタイル
        self._init_style()

        # 全体レイアウト  ┌─────────────── header ────────────────┐
        #                │                                       │
        #                ├─ ribbon (left) ─┬─ work area (right) ┤
        #                │                  │                    │
        #                └───────────────────────────────────────┘
        self._build_header()
        self._build_body()

        # 空状態の表示
        self._mount_empty_state()

        # メニューバー（ヘッダー右上の小さなボタン群）
        self._build_header_actions()

        # 初期: プラグイン読み込み
        self.reload_plugins()

        # ステータスバー更新タイマー
        self._tick_clock()

    # ---------- STYLE ----------
    def _init_style(self):
        style = ttk.Style()
        try:
            # Windows 推奨テーマ
            style.theme_use("vista")
        except Exception:
            pass

        # 色
        self.c_bg = "#eef1f7"        # アプリ背景
        self.c_hdr_top = "#3b82f6"   # ヘッダー上グラデ
        self.c_hdr_bot = "#2563eb"   # ヘッダー下グラデ
        self.c_hdr_text = "#ffffff"  # ヘッダーテキスト
        self.c_ribbon_bg = "#f7f9fc" # リボン背景
        self.c_card = "#ffffff"      # カード背景
        self.c_muted = "#6b7280"     # 説明テキスト
        self.c_acc = "#10b981"       # アクセント（緑）
        self.c_warn = "#ef4444"      # 警告（赤）

        # ラベル系
        style.configure("HdrTitle.TLabel", font=("Yu Gothic UI", 18, "bold"), foreground=self.c_hdr_text, background=self.c_hdr_bot)
        style.configure("HdrSub.TLabel", font=("Yu Gothic UI", 9), foreground="#e6f0ff", background=self.c_hdr_bot)
        style.configure("HdrMode.TLabel", font=("Consolas", 10, "bold"), foreground="#e6f0ff", background=self.c_hdr_bot)

        style.configure("Ribbon.TFrame", background=self.c_ribbon_bg)
        style.configure("RibbonTitle.TLabel", font=("Yu Gothic UI", 11, "bold"), foreground="#1f2937", background=self.c_ribbon_bg)
        style.configure("RibbonMuted.TLabel", font=("Yu Gothic UI", 9), foreground=self.c_muted, background=self.c_ribbon_bg)

        style.configure("Card.TFrame", background=self.c_card)
        style.configure("CardTitle.TLabel", font=("Yu Gothic UI", 13, "bold"), background=self.c_card, foreground="#111827")
        style.configure("CardText.TLabel", font=("Yu Gothic UI", 10), background=self.c_card, foreground=self.c_muted)

        style.configure("BigAction.TButton", font=("Yu Gothic UI", 11, "bold"))
        style.configure("Ghost.TButton", font=("Yu Gothic UI", 10))

        # リボンボタン（丸みのあるフラット）
        style.configure("Ribbon.TButton", font=("Yu Gothic UI", 11, "bold"), padding=10)
        style.map("Ribbon.TButton",
                  background=[("active", "#e0ebff")],
                  relief=[("pressed", "sunken"), ("!pressed", "flat")])

    # ---------- HEADER ----------
    def _build_header(self):
        # グラデーション風ヘッダー（Canvasで段階塗り）
        self.header = tk.Canvas(self.root, height=84, bd=0, highlightthickness=0, relief="ridge")
        self.header.pack(fill="x", side="top")
        self._draw_header_gradient()

        # タイトル＆モード名
        self.hdr_title = tk.Label(self.header, text=f"✨ {APP_NAME}", font=("Yu Gothic UI", 18, "bold"), fg=self.c_hdr_text, bg=self.c_hdr_bot)
        self.hdr_desc  = tk.Label(self.header, text=APP_DESC, font=("Yu Gothic UI", 9), fg="#e6f0ff", bg=self.c_hdr_bot)
        self.hdr_mode  = tk.Label(self.header, text="CURRENT MODE: —", font=("Consolas", 10, "bold"), fg="#e6f0ff", bg=self.c_hdr_bot)

        # 右上: ヘルプ／プラグインフォルダ／再読み込み
        self.hdr_btns = tk.Frame(self.header, bg=self.c_hdr_bot)

        # 位置
        self.header.create_window(20, 18, window=self.hdr_title, anchor="nw")
        self.header.create_window(20, 50, window=self.hdr_desc, anchor="nw")
        self.header.create_window(320, 20, window=self.hdr_mode, anchor="nw")
        self.header.create_window(self.root.winfo_width() - 20, 18, window=self.hdr_btns, anchor="ne")
        self.header.bind("<Configure>", lambda e: self._reposition_header_buttons())

    def _draw_header_gradient(self):
        self.header.delete("grad")
        w = max(self.header.winfo_width(), 600)
        h = self.header.winfo_height() or 84
        steps = 60
        for i in range(steps):
            r = i / max(steps - 1, 1)
            # 簡易グラデーション（上: c_hdr_top → 下: c_hdr_bot）
            color = self._lerp_color(self.c_hdr_top, self.c_hdr_bot, r)
            self.header.create_rectangle(0, int(i * h / steps), w, int((i + 1) * h / steps), outline="", fill=color, tags="grad")

    def _reposition_header_buttons(self):
        self.header.delete("btns")
        self.header.create_window(self.header.winfo_width() - 20, 18, window=self.hdr_btns, anchor="ne", tags="btns")

    @staticmethod
    def _hex_to_rgb(hexstr: str):
        hexstr = hexstr.lstrip("#")
        return tuple(int(hexstr[i:i+2], 16) for i in (0, 2, 4))

    @staticmethod
    def _rgb_to_hex(rgb: tuple[int,int,int]):
        return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"

    def _lerp_color(self, c1: str, c2: str, t: float) -> str:
        r1,g1,b1 = self._hex_to_rgb(c1)
        r2,g2,b2 = self._hex_to_rgb(c2)
        r = int(r1 + (r2 - r1) * t)
        g = int(g1 + (g2 - g1) * t)
        b = int(b1 + (b2 - b1) * t)
        return self._rgb_to_hex((r,g,b))

    def _build_header_actions(self):
        for w in self.hdr_btns.winfo_children():
            w.destroy()

        def btn(text, cmd):
            b = ttk.Button(self.hdr_btns, text=text, style="Ghost.TButton", command=cmd)
            b.pack(side="left", padx=4)
            return b

        btn("🧩 プラグインフォルダを開く", self.open_plugins_folder)
        btn("🔄 プラグインを再読み込み", self.reload_plugins)
        btn("❔ ヘルプ", self.open_help)

        # ステータス（右下）
        self.statusbar = tk.Label(self.root, text="—", anchor="e", bg=self.c_bg, fg="#4b5563", font=("Consolas", 9))
        self.statusbar.pack(fill="x", side="bottom", padx=14, pady=(0,6))

    # ---------- BODY ----------
    def _build_body(self):
        body = tk.Frame(self.root, bg=self.c_bg)
        body.pack(fill="both", expand=True)

        # リボン（左）
        self.ribbon = tk.Frame(body, bg=self.c_ribbon_bg, width=220, bd=0, highlightthickness=0)
        self.ribbon.pack(side="left", fill="y")
        self._build_ribbon()

        # 作業エリア（右）カード
        work_wrap = tk.Frame(body, bg=self.c_bg)
        work_wrap.pack(side="right", fill="both", expand=True)

        self.work_card = tk.Frame(work_wrap, bg=self.c_card, bd=0, highlightthickness=0)
        self.work_card.pack(fill="both", expand=True, padx=18, pady=18)

        # 角丸っぽい影（簡易）
        self.work_card.configure(highlightbackground="#dbe3f2", highlightcolor="#dbe3f2", highlightthickness=1)

    def _build_ribbon(self):
        # タイトル
        tk.Label(self.ribbon, text="モード（プラグイン）", bg=self.c_ribbon_bg, fg="#111827",
                 font=("Yu Gothic UI", 12, "bold")).pack(anchor="w", padx=16, pady=(16, 2))
        ttk.Label(self.ribbon, text="左のボタンからモードを切替えます", style="RibbonMuted.TLabel").pack(anchor="w", padx=16, pady=(0, 12))

        # スクロール可能なボタン領域
        container = tk.Frame(self.ribbon, bg=self.c_ribbon_bg)
        container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.ribbon_canvas = tk.Canvas(container, bg=self.c_ribbon_bg, bd=0, highlightthickness=0)
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.ribbon_canvas.yview)
        self.ribbon_body = tk.Frame(self.ribbon_canvas, bg=self.c_ribbon_bg)

        self.ribbon_body.bind("<Configure>", lambda e: self.ribbon_canvas.configure(scrollregion=self.ribbon_canvas.bbox("all")))
        self.ribbon_canvas.create_window((0,0), window=self.ribbon_body, anchor="nw")
        self.ribbon_canvas.configure(yscrollcommand=vsb.set)

        self.ribbon_canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # “空”の時のプレースホルダ
        self.ribbon_empty = None
        self._mount_ribbon_empty()

        # 下部固定の操作
        bottom = tk.Frame(self.ribbon, bg=self.c_ribbon_bg)
        bottom.pack(side="bottom", fill="x", padx=10, pady=12)
        ttk.Button(bottom, text="🧩 プラグインフォルダ", style="Ghost.TButton", command=self.open_plugins_folder).pack(side="left")
        ttk.Button(bottom, text="🔄 再読み込み", style="Ghost.TButton", command=self.reload_plugins).pack(side="right")

    def _mount_ribbon_empty(self):
        # 既存のプレースホルダがあれば破棄（壊れた参照対策）
        if getattr(self, "ribbon_empty", None) is not None:
            try:
                if self.ribbon_empty.winfo_exists():
                    self.ribbon_empty.destroy()
            except Exception:
                pass
            self.ribbon_empty = None

        # 新しく作り直す
        self.ribbon_empty = tk.Frame(self.ribbon_body, bg=self.c_ribbon_bg)
        self.ribbon_empty.pack(fill="both", expand=True, padx=8, pady=8)

        card = tk.Frame(self.ribbon_empty, bg="#ffffff", highlightbackground="#e5e7eb", highlightthickness=1)
        card.pack(fill="x", padx=4, pady=8)

        tk.Label(card, text="プラグインがありません", bg="#ffffff", fg="#111827",
                 font=("Yu Gothic UI", 11, "bold")).pack(anchor="w", padx=10, pady=(10, 2))
        tk.Label(card, text="plugins フォルダにプラグインを追加すると\nここにモードボタンが増えます。",
                 bg="#ffffff", fg=self.c_muted, font=("Yu Gothic UI", 9), justify="left").pack(anchor="w", padx=10, pady=(0, 10))

    def _clear_ribbon_buttons(self):
        for child in list(self.ribbon_body.winfo_children()):
            # 空カードは残さず全部消す
            child.destroy()
        self.ribbon_empty = None

    def _add_ribbon_button(self, plugin: PluginBase):
        btn = ttk.Button(self.ribbon_body,
                         text=f"{getattr(plugin, 'icon', '🔹')}  {getattr(plugin, 'name', 'Unnamed')}",
                         style="Ribbon.TButton",
                         command=lambda p=plugin: self.switch_mode(p))
        btn.pack(fill="x", padx=8, pady=6)

    # ---------- CONTENT (EMPTY STATE) ----------
    def _mount_empty_state(self):
        for w in self.work_card.winfo_children():
            w.destroy()

        # 右側の空状態ビュー
        header = tk.Frame(self.work_card, bg=self.c_card)
        header.pack(fill="x", padx=18, pady=(18, 8))

        ttk.Label(header, text="ようこそ！", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(header, text="このアプリはプラグインを追加して機能拡張する“シェル”です。\n左のリボンからモードを選択できますが、現状はプラグインがありません。",
                  style="CardText.TLabel").pack(anchor="w", pady=(4, 6))

        # アクションライン
        actions = tk.Frame(self.work_card, bg=self.c_card)
        actions.pack(fill="x", padx=18, pady=(2, 18))

        ttk.Button(actions, text="🧩 プラグインフォルダを開く", style="BigAction.TButton", command=self.open_plugins_folder).pack(side="left")
        ttk.Button(actions, text="🔄 プラグインを再読み込み", style="BigAction.TButton", command=self.reload_plugins).pack(side="left", padx=10)

        # 視覚カード（ヒント）
        tips = tk.Frame(self.work_card, bg=self.c_card)
        tips.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        tip_card = tk.Frame(tips, bg="#f8fafc", highlightbackground="#e5e7eb", highlightthickness=1)
        tip_card.pack(fill="both", expand=True)

        tk.Label(tip_card, text="プラグインの作り方（ざっくり）", bg="#f8fafc", fg="#0f172a",
                 font=("Yu Gothic UI", 11, "bold")).pack(anchor="w", padx=14, pady=(12, 6))

        msg = (
            "1) plugins/ に Python ファイルを置き、PluginBase を継承した class Plugin を実装\n"
            "2) name / icon / mount(parent) / unmount() を定義\n"
            "3) 保存後、左下の [再読み込み] を押すとボタンが追加されます\n"
        )
        tk.Label(tip_card, text=msg, bg="#f8fafc", fg=self.c_muted, font=("Yu Gothic UI", 9), justify="left").pack(anchor="w", padx=14, pady=(0, 14))

        self._set_mode_label("—")

    # ---------- PLUGIN MGMT ----------
    def reload_plugins(self):
        # 現プラグインのアンマウント
        if self.current_plugin:
            try:
                self.current_plugin.unmount()
            except Exception:
                pass
        self.current_plugin = None

        # 既存プラグイン破棄
        self.plugins.clear()
        self._clear_ribbon_buttons()

        # 読込
        found = 0
        for file in sorted(self.plugins_dir.glob("*.py")):
            try:
                spec = importlib.util.spec_from_file_location(file.stem, file)
                if not spec or not spec.loader:
                    continue
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)  # type: ignore
                if hasattr(module, "Plugin"):
                    cls = getattr(module, "Plugin")
                    if issubclass(cls, PluginBase):
                        instance = cls(shell_context={"base_dir": str(self.base_dir)})
                        self.plugins.append(instance)
                        self._add_ribbon_button(instance)
                        found += 1
            except Exception as e:
                print(f"[WARN] failed to load plugin {file.name}: {e}")

        if found == 0:
            self._mount_ribbon_empty()
            self._mount_empty_state()
        else:
            # 最初のプラグインを自動選択してもいいが、今回は空状態のままにする
            # self.switch_mode(self.plugins[0])
            pass

        self._set_status(f"プラグイン: {found} 個")
        self._draw_header_gradient()  # サイズ変化時に再描画

    def switch_mode(self, plugin: PluginBase):
        # 前のプラグインを降ろす
        if self.current_plugin:
            try:
                self.current_plugin.unmount()
            except Exception:
                pass

        # 作業領域クリア
        for w in self.work_card.winfo_children():
            w.destroy()

        # 新プラグインをマウント
        try:
            plugin.mount(self.work_card)
            self.current_plugin = plugin
            self._set_mode_label(getattr(plugin, "name", "—"))
            self._set_status(f"現在のモード: {getattr(plugin, 'name', '')}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"プラグインのマウントに失敗しました。\n{e}")
            self._mount_empty_state()

    # ---------- ACTIONS ----------
    def open_plugins_folder(self):
        p = str(self.plugins_dir.resolve())
        try:
            if platform.system() == "Windows":
                os.startfile(p)  # type: ignore
            elif platform.system() == "Darwin":
                subprocess.run(["open", p])
            else:
                subprocess.run(["xdg-open", p])
        except Exception:
            messagebox.showinfo(APP_NAME, f"プラグインフォルダ: {p}")

    def open_help(self):
        # 任意のヘルプ先（ここでは簡易メッセージ）
        messagebox.showinfo(APP_NAME,
                            "このアプリはプラグインを追加して機能を拡張します。\n"
                            "plugins フォルダに Python ファイルを追加し、PluginBase を継承した class Plugin を実装してください。")

    # ---------- HEADER/STATUS ----------
    def _set_mode_label(self, mode: str):
        self.hdr_mode.configure(text=f"CURRENT MODE: {mode}")

    def _set_status(self, text: str):
        self.statusbar.configure(text=f"{time.strftime('%H:%M:%S')}  {text}")

    def _tick_clock(self):
        # 時計（と軽い“生きてる”感）の更新
        now = time.strftime("%H:%M:%S")
        cur = self.statusbar.cget("text")
        if "  " in cur:
            suffix = cur.split("  ", 1)[1]
        else:
            suffix = "Ready"
        self.statusbar.configure(text=f"{now}  {suffix}")
        self.root.after(1000, self._tick_clock)


def main():
    root = tk.Tk()
    app = ShellApp(root)
    # ウィンドウサイズ変更に合わせてヘッダーグラデを再描画
    root.bind("<Configure>", lambda e: app._draw_header_gradient() if e.widget is root else None)
    root.mainloop()


if __name__ == "__main__":
    main()
