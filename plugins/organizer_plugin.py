# plugins/organizer_plugin.py
# -*- coding: utf-8 -*-
"""
Excel 整形プラグイン（Plugin Shell 用）
- app_shell.py の PluginBase を継承して UI をマウントします
- シェルの base_dir（shell_context["base_dir"]）にある以下のファイルを使用します：
    - Attendee_format_original.xlsx  … テンプレ配布の原本
    - romaji_mapping.json            … ローマ字変換
    - company_replacements.json      … 会社名置換ルール
- 実行には Windows + Microsoft Excel（pywin32/win32com）が必要です
- Web 版（Streamlit/Render）の場合は web_mount() で“非対応”の案内を表示します
"""

from __future__ import annotations
import sys
import json
import unicodedata
import shutil
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---- PluginBase をシェルと“同一オブジェクト”で継承するための対策 ----
#   - デスクトップ(app_shell.pyを __main__ として実行)： sys.modules["__main__"].PluginBase
#   - Web(app_web.py から import)： from app_shell import PluginBase
try:
    from app_shell import PluginBase  # Web側（app_web.py）経由を想定
except Exception:
    PluginBase = sys.modules["__main__"].PluginBase  # デスクトップ側

# ---- Excel COM ----
try:
    import win32com.client as win32  # type: ignore
except Exception:
    win32 = None  # Excel が無い環境対策

# ==== 定数（シェルの assets と合わせる）====
TEMPLATE_XLSX_ORIGINAL = "Attendee_format_original.xlsx"
ROMAJI_JSON = "romaji_mapping.json"
COMPANY_JSON = "company_replacements.json"

DATA_SHEET = "DATA"
OUTPUTS_SHEET = "Outputs"

# ==== ドメイン処理（excel_organizer 相当の要点）====
def _to_zen_katakana(s):
    if s is None:
        return None
    t = unicodedata.normalize("NFKC", str(s))
    res = []
    for ch in t:
        code = ord(ch)
        if 0x3041 <= code <= 0x3096:  # ひらがな → カタカナ
            res.append(chr(code + 0x60))
        else:
            res.append(ch)
    return "".join(res)

def _kata_to_romaji(text, digraphs, mono):
    if text is None:
        return None
    s = _to_zen_katakana(text)

    def double_consonant(roma_next: str) -> str:
        if not roma_next:
            return ""
        if roma_next.startswith("ch"): return "c"
        if roma_next.startswith("sh"): return "s"
        if roma_next.startswith("j"):  return "j"
        if roma_next.startswith("ts"): return "t"
        return roma_next[0]

    def prolong(prev: str) -> str:
        if not prev:
            return ""
        for v in ("a","i","u","e","o"):
            if prev.endswith(v): return v
        return ""

    res = []
    i = 0
    while i < len(s):
        ch = s[i]
        code = ord(ch)
        is_katakana = (0x30A0 <= code <= 0x30FF) or ch == "ー"

        if not is_katakana:
            res.append(ch); i += 1; continue

        if ch == "ッ":
            if i + 1 < len(s):
                two = s[i+1:i+3]
                roma_next = digraphs.get(two) if len(two) == 2 else None
                if roma_next is None:
                    roma_next = mono.get(s[i+1], "")
                res.append(double_consonant(roma_next))
            i += 1; continue

        if ch == "ー":
            res.append(prolong("".join(res))); i += 1; continue

        two = s[i:i+2]
        if len(two) == 2 and two in digraphs:
            res.append(digraphs[two]); i += 2; continue

        roma = mono.get(ch)
        if roma:
            if ch == "ン":
                nxt = s[i+1] if i + 1 < len(s) else ""
                nxt_roma = digraphs.get(s[i+1:i+3]) if i + 2 < len(s) else None
                if nxt_roma is None:
                    nxt_roma = mono.get(nxt, "")
                res.append("n'" if nxt_roma[:1] in ("a","i","u","e","o","y") else "n")
            else:
                res.append(roma)
        else:
            res.append(ch)
        i += 1

    out = "".join(res).replace("-", "")
    romaji = out.lower()
    return romaji.capitalize()

def _run_excel_pipeline(input_path: Path, base: Path, output_path: Path):
    """
    入力Excel（DATA→Outputs）に対して整形処理を実行し、output_path に保存する。
    - バックアップは作らない（出力は別ファイル）
    """
    if win32 is None:
        raise RuntimeError("pywin32 / win32com が使用できません。Excel と pywin32 を確認してください。")

    # JSON 読み込み
    with open(base / ROMAJI_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)
    digraphs, mono = data.get("digraphs", {}), data.get("mono", {})

    with open(base / COMPANY_JSON, "r", encoding="utf-8") as f:
        company_rules = json.load(f)

    FILE_PATH = Path(input_path)
    SHEET_SRC = DATA_SHEET
    SHEET_DST = OUTPUTS_SHEET

    # Excel const
    xlValues = -4163
    xlWhole  = 1
    xlPart   = 2
    xlByRows = 1

    def find_header(ws, name: str):
        found = ws.Cells.Find(What=name, LookIn=xlValues, LookAt=xlWhole)
        if found is None:
            return None, None
        return found.Row, found.Column

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(str(FILE_PATH))

        # ① DATA→Outputs
        try:
            ws_src = wb.Worksheets(SHEET_SRC)
        except Exception:
            raise RuntimeError("シート 'DATA' が見つかりません。")

        try:
            ws_dst = wb.Worksheets(SHEET_DST)
        except Exception:
            ws_dst = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws_dst.Name = SHEET_DST

        ws_dst.Cells.Clear()
        ws_src.UsedRange.Copy()
        ws_dst.Range("A1").PasteSpecial(-4104)  # xlPasteAll

        # 列幅/行高
        src_ur = ws_src.UsedRange
        for i in range(1, src_ur.Columns.Count + 1):
            ws_dst.Columns(i).ColumnWidth = ws_src.Columns(i).ColumnWidth
        for r in range(1, src_ur.Rows.Count + 1):
            ws_dst.Rows(r).RowHeight = ws_src.Rows(r).RowHeight

        # ページ設定（主要）
        try:
            sps, dps = ws_src.PageSetup, ws_dst.PageSetup
            for a in ("Orientation","Zoom","FitToPagesWide","FitToPagesTall",
                      "LeftMargin","RightMargin","TopMargin","BottomMargin",
                      "HeaderMargin","FooterMargin","CenterHorizontally","CenterVertically",
                      "PrintTitleRows","PrintTitleColumns"):
                try: setattr(dps, a, getattr(sps, a))
                except Exception: pass
            try:
                pa = sps.PrintArea
                if pa and "!" in pa: pa = pa.split("!",1)[1]
                if pa: dps.PrintArea = pa
            except Exception: pass
        except Exception:
            pass

        used = ws_dst.UsedRange
        last_row = used.Row + used.Rows.Count - 1

        # ② かな正規化
        for col_name in ("Kana_First_Orig", "Kana_Last_Orig"):
            hr, hc = find_header(ws_dst, col_name)
            if hr is None:
                continue
            for r in range(hr + 1, last_row + 1):
                cell = ws_dst.Cells(r, hc)
                v = cell.Value
                nv = _to_zen_katakana(v) if v is not None else v
                if nv != v:
                    cell.Value = nv

        # ③ ローマ字生成
        targets = {"Romaji_First_Orig": "Kana_First_Orig", "Romaji_Last_Orig": "Kana_Last_Orig"}
        for romaji_col, kana_col in targets.items():
            hr_rom, hc_rom = find_header(ws_dst, romaji_col)
            if hr_rom is None:
                continue
            hr_kana, hc_kana = find_header(ws_dst, kana_col)
            src_hc = hc_kana if hc_kana is not None else hc_rom
            for r in range(hr_rom + 1, last_row + 1):
                src_val = ws_dst.Cells(r, src_hc).Value
                if src_val is None or str(src_val).strip() == "":
                    continue
                roma = _kata_to_romaji(src_val, digraphs, mono)
                cur  = ws_dst.Cells(r, hc_rom).Value
                if roma is not None and roma != cur:
                    ws_dst.Cells(r, hc_rom).Value = roma

        # ④ 会社略記→正式表記（部分一致）
        for rule in company_rules:
            patterns = rule.get("patterns", [])
            replacement = rule.get("replacement", "")
            for old in patterns:
                ws_dst.Cells.Replace(
                    What=old,
                    Replacement=replacement,
                    LookAt=xlPart,
                    SearchOrder=xlByRows,
                    MatchCase=False,
                    SearchFormat=False,
                    ReplaceFormat=False
                )

        wb.SaveCopyAs(str(output_path))

    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass

# ==== プラグイン本体 ====
class Plugin(PluginBase):
    name = "Excel整形"
    icon = "🧹"

    def __init__(self, shell_context: dict | None = None) -> None:
        super().__init__(shell_context)
        self.base = Path(self.shell_context.get("base_dir", "."))  # シェルから渡される base_dir
        self.selected_file: Optional[Path] = None

        # UI要素（後で参照するもの）
        self.root: Optional[tk.Frame] = None
        self.lbl_status: Optional[ttk.Label] = None
        self.btn_run: Optional[ttk.Button] = None
        self.drop_label: Optional[ttk.Label] = None
        self.log: Optional[tk.Text] = None

    # UI構築（デスクトップ/Tk）
    def mount(self, parent: tk.Frame) -> None:
        self.root = tk.Frame(parent, bg="#ffffff")
        self.root.pack(fill="both", expand=True)

        # セクション: ヘッダ
        header = ttk.Frame(self.root, padding=(18, 14))
        header.pack(fill="x")
        ttk.Label(header, text="Excel データ整形", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(header, text="テンプレ配布 → DATA を記入 → 整形して Outputs を生成（別名保存）",
                  style="CardText.TLabel").pack(anchor="w", pady=(2, 0))

        # ステータス（テンプレ/JSONの有無）
        status = ttk.Frame(self.root, padding=(18, 0))
        status.pack(fill="x")
        self.lbl_status = ttk.Label(status, text=self._status_text(), style="CardText.TLabel")
        self.lbl_status.pack(anchor="e")

        # カード：テンプレ配布
        card1 = ttk.Frame(self.root, style="Card.TFrame", padding=14)
        card1.pack(fill="x", padx=18, pady=(12, 8))
        ttk.Label(card1, text="テンプレート", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(card1, text="同フォルダの Attendee_format_original.xlsx をコピーして配布します。",
                  style="CardText.TLabel").pack(anchor="w", pady=(2, 8))
        ttk.Button(card1, text="📄 テンプレートを保存 (Attendee_format.xlsx)", command=self._on_export_template).pack(anchor="w")

        # カード：ファイル選択
        card2 = ttk.Frame(self.root, style="Card.TFrame", padding=14)
        card2.pack(fill="x", padx=18, pady=(8, 8))
        ttk.Label(card2, text="入力ファイル", style="CardTitle.TLabel").pack(anchor="w")
        drop = ttk.Frame(card2, padding=16)
        drop.pack(fill="x", pady=(6, 6))
        drop.configure(style="Card.TFrame")
        self.drop_label = ttk.Label(drop, text="クリックして Excel（.xlsx / DATAシート）を選択", style="CardText.TLabel")
        self.drop_label.pack(fill="x")
        drop.bind("<Button-1>", lambda e: self._on_browse_file())
        ttk.Button(card2, text="🔎 ファイルを選択", command=self._on_browse_file).pack(anchor="e")
        self.btn_run = ttk.Button(card2, text="▶ データ整形を実行", command=self._on_run, state="disabled")
        self.btn_run.pack(anchor="e", pady=(6, 0))

        # カード：ログ
        card3 = ttk.Frame(self.root, style="Card.TFrame", padding=10)
        card3.pack(fill="both", expand=True, padx=18, pady=(8, 18))
        ttk.Label(card3, text="ログ", style="CardTitle.TLabel").pack(anchor="w")
        self.log = tk.Text(card3, height=10, relief="flat", bg="#ffffff")
        self.log.pack(fill="both", expand=True, pady=(6, 0))
        self._log("プラグイン起動", "ready")

    # Web用（Streamlit）：Render等での案内
    def web_mount(self, st):
        base = Path(self.shell_context.get("base_dir", "."))
        st.info("このプラグインは **Windows の Excel COM(pywin32)** を使用するため、"
                "ブラウザ実行（Render/Streamlit）では処理本体は動作しません。")
        st.write("デスクトップ版での利用手順：")
        st.markdown(
            "- app_shell.py と同じフォルダに以下のファイルを置く：\n"
            "  - `Attendee_format_original.xlsx`\n"
            "  - `romaji_mapping.json`\n"
            "  - `company_replacements.json`\n"
            "- `pip install pywin32 openpyxl`\n"
            "- プラグインから Excel ファイルを選択 → 整形 → 別名保存"
        )
        exists = {
            "Attendee_format_original.xlsx": (base / "Attendee_format_original.xlsx").exists(),
            "romaji_mapping.json": (base / "romaji_mapping.json").exists(),
            "company_replacements.json": (base / "company_replacements.json").exists(),
        }
        st.subheader("配置チェック")
        st.json(exists)

    def unmount(self) -> None:
        if self.root and self.root.winfo_exists():
            self.root.destroy()
        self.root = None

    # ====== イベント ======
    def _on_export_template(self):
        src = self.base / TEMPLATE_XLSX_ORIGINAL
        if not src.exists():
            messagebox.showerror(self.name, f"{TEMPLATE_XLSX_ORIGINAL} が見つかりません。")
            return
        dest = filedialog.asksaveasfilename(
            title="テンプレートの保存先",
            defaultextension=".xlsx",
            initialfile="Attendee_format.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not dest:
            return
        try:
            shutil.copyfile(src, dest)
            self._log(f"テンプレ保存: {dest}", "ok")
            messagebox.showinfo(self.name, "テンプレートを保存しました。")
        except Exception as e:
            self._log(f"テンプレ保存に失敗: {e}", "error")
            messagebox.showerror(self.name, f"保存に失敗しました。\n{e}")

    def _on_browse_file(self):
        f = filedialog.askopenfilename(
            title="Excelファイルを選択（DATAシート）",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")]
        )
        if not f:
            return
        p = Path(f)
        # 最低限 DATA シート存在チェック
        try:
            import openpyxl  # type: ignore
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            if DATA_SHEET not in wb.sheetnames:
                messagebox.showwarning(self.name, f"シート '{DATA_SHEET}' が見つかりません。")
                return
        except Exception as e:
            messagebox.showerror(self.name, f"Excel の読み込みに失敗しました。\n{e}")
            return

        self.selected_file = p
        if self.drop_label:
            self.drop_label.configure(text=f"選択中: {p.name}")
        if self.btn_run:
            self.btn_run.configure(state="normal")
        self._log(f"選択: {p}", "ok")

    def _on_run(self):
        if self.selected_file is None:
            return
        # 必須ファイル
        missing = [name for name in (ROMAJI_JSON, COMPANY_JSON) if not (self.base / name).exists()]
        if missing:
            messagebox.showerror(self.name, f"必要ファイルが見つかりません：{', '.join(missing)}")
            return
        if win32 is None:
            messagebox.showerror(self.name, "pywin32 が必要です。\n pip install pywin32")
            return

        # 保存先
        default_name = f"{self.selected_file.stem}_organized.xlsx"
        out = filedialog.asksaveasfilename(
            title="整形後ファイルの保存先",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not out:
            self._log("保存先の指定をキャンセル", "info")
            return

        try:
            if self.btn_run: self.btn_run.configure(state="disabled")
            if self.root: self.root.config(cursor="watch")
            self._log("データ整形を開始...", "info")
            _run_excel_pipeline(self.selected_file, self.base, Path(out))
            self._log(f"完了: {out}", "ok")
            messagebox.showinfo(self.name, f"整形が完了しました。\n\n{out}")
        except Exception as e:
            self._log(f"整形に失敗: {e}", "error")
            messagebox.showerror(self.name, f"整形に失敗しました。\n{e}")
        finally:
            if self.root: self.root.config(cursor="")
            if self.btn_run: self.btn_run.configure(state="normal")

    # ===== ユーティリティ =====
    def _status_text(self) -> str:
        tmpl = (self.base / TEMPLATE_XLSX_ORIGINAL).exists()
        r_ok = (self.base / ROMAJI_JSON).exists()
        c_ok = (self.base / COMPANY_JSON).exists()
        return f"TEMPLATE={'OK' if tmpl else 'NG'} / JSON: romaji={'OK' if r_ok else 'NG'} / company={'OK' if c_ok else 'NG'}"

    def _log(self, msg: str, level: str = "info"):
        if not self.log:
            return
        from datetime import datetime
        now = datetime.now().strftime("%H:%M:%S")
        tags = {"info": "[i]", "ok": "[✓]", "error": "[!]", "ready": "[•]"}
        self.log.insert("end", f"{now} {tags.get(level,'[ ]')} {msg}\n")
        self.log.see("end")
