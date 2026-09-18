from __future__ import annotations

import re
import threading
import os
import shutil
import tempfile
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


MONTH_FILE_RE = re.compile(r"(?P<yy>\d{2})년\s*(?P<mm>\d{1,2})월\s*검수현황", re.IGNORECASE)
APP_DATA_DIR = Path(__file__).resolve().parent / "data"
SUMMARY_FILE = APP_DATA_DIR / "검수종합_현황.xlsx"


def extract_year_month(path: Path) -> tuple[str, str]:
    """Return (YYYY-MM, display label) from a Korean monthly filename."""
    match = MONTH_FILE_RE.search(path.stem)
    if not match:
        raise ValueError("파일명에서 'YY년 MM월 검수현황' 패턴을 찾을 수 없습니다.")
    year = 2000 + int(match.group("yy"))
    month = int(match.group("mm"))
    if not 1 <= month <= 12:
        raise ValueError(f"월 값이 올바르지 않습니다: {month}")
    return f"{year:04d}-{month:02d}", f"{year}년 {month:02d}월"


def find_header_row(path: Path, sheet_name=0, scan_rows: int = 30) -> int:
    """Find the first row that looks like a column header."""
    preview = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=scan_rows)
    keywords = ("검수", "품번", "품명", "부서", "일자", "날짜", "업체", "구분")
    best_row, best_score = 0, -1
    for idx, row in preview.iterrows():
        values = [str(v).strip() for v in row.tolist() if pd.notna(v) and str(v).strip()]
        if not values:
            continue
        score = sum(any(k in value for k in keywords) for value in values)
        score += min(len(values), 10) / 10
        if score > best_score:
            best_row, best_score = idx, score
    return int(best_row)


def canonical_column_name(value) -> str:
    """Make monthly Excel headers comparable despite spaces/newlines or NBSP."""
    text = str(value).replace("\u00a0", " ").replace("\n", " ").strip()
    compact = re.sub(r"\s+", "", text)
    aliases = {
        "수량누계": "수량누계",
        "금액누계": "금액누계",
    }
    return aliases.get(compact, text)


def coalesce_duplicate_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Merge columns that only differ by header spacing, preserving all values."""
    result = pd.DataFrame(index=frame.index)
    seen = set()
    for name in frame.columns:
        if name in seen:
            continue
        seen.add(name)
        matching = frame.loc[:, [column == name for column in frame.columns]]
        merged = matching.iloc[:, 0]
        for column_index in range(1, matching.shape[1]):
            merged = merged.combine_first(matching.iloc[:, column_index])
        result[name] = merged
    return result


def read_one_file(path: Path) -> pd.DataFrame:
    excel = pd.ExcelFile(path)
    sheet_name = excel.sheet_names[0]
    header_row = find_header_row(path, sheet_name)
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
    df = df.dropna(how="all").copy()
    df.columns = [canonical_column_name(c) if not str(c).startswith("Unnamed:") else "" for c in df.columns]
    df = df.loc[:, [c != "" for c in df.columns]]
    df = df.dropna(axis=1, how="all")
    df = coalesce_duplicate_columns(df)
    existing_period_col = next((c for c in df.columns if str(c).strip() == "해당년월"), None)
    if existing_period_col:
        # Already consolidated files can be used as input without relying on
        # their filename (for example, 검수종합_추가_20260917.xlsx).
        period = None
    else:
        period, _ = extract_year_month(path)
    # Exclude source summary rows so monthly totals are not counted as data.
    summary_pattern = r"합계|총계|소계|grand\s*total|subtotal|total"
    text_view = df.astype("string").fillna("")
    summary_rows = text_view.apply(
        lambda row: row.str.contains(summary_pattern, case=False, regex=True).any(), axis=1
    )
    df = df.loc[~summary_rows].copy()
    part_col = next((c for c in df.columns if re.sub(r"\s+", "", str(c)) == "품번"), None)
    if part_col:
        # In these source sheets, rows with no part number are monthly subtotal rows.
        df = df[df[part_col].notna() & (df[part_col].astype("string").str.strip() != "")].copy()
    # Keep only the sortable YYYY-MM period column requested by the user.
    if not existing_period_col:
        df.insert(0, "해당년월", period)
    return df


def collect_files(folders: list[Path]) -> list[Path]:
    files: list[Path] = []
    for folder in folders:
        for path in folder.iterdir():
            if path.is_file() and path.suffix.lower() in {".xls", ".xlsx"} and MONTH_FILE_RE.search(path.stem):
                files.append(path)
    return sorted(files, key=lambda p: (extract_year_month(p)[0], str(p).lower()))


def normalize_customer_name(value) -> str:
    """Normalize punctuation and spacing for customer-name comparison."""
    if pd.isna(value):
        return ""
    return re.sub(r"[^0-9A-Za-z가-힣]", "", str(value)).lower()


def merge_files(folders: list[Path], output: Path, progress_callback=None, selected_files: list[Path] | None = None) -> tuple[int, int, list[str]]:
    files = selected_files if selected_files is not None else collect_files(folders)
    if not files:
        raise ValueError("조건에 맞는 월별 검수현황 파일이 없습니다.")
    frames: list[pd.DataFrame] = []
    errors: list[str] = []
    for i, path in enumerate(files, 1):
        try:
            frames.append(read_one_file(path))
        except Exception as exc:
            errors.append(f"{path.name}: {exc}")
        if progress_callback:
            progress_callback(i, len(files), path.name)
    if not frames:
        raise ValueError("읽을 수 있는 파일이 없습니다.\n" + "\n".join(errors))
    merged = pd.concat(frames, ignore_index=True, sort=False)

    customer_col = next((c for c in merged.columns if str(c).strip() == "거래처명"), None)
    if customer_col and "해당년월" in merged.columns:
        # Process newest records first so the first representative is the
        # latest spelling used in the source data.
        ordered = merged.sort_values("해당년월", ascending=False, kind="stable")
        representatives: list[tuple[str, str]] = []
        replacements: dict[str, str] = {}
        for value in ordered[customer_col].dropna():
            original = str(value).strip()
            key = normalize_customer_name(original)
            if not key or key in replacements:
                continue
            match = None
            for rep_key, rep_value in representatives:
                similarity = SequenceMatcher(None, key, rep_key).ratio()
                if similarity >= 0.90 or (min(len(key), len(rep_key)) >= 8 and (key in rep_key or rep_key in key) and similarity >= 0.82):
                    match = rep_value
                    break
            if match is None:
                representatives.append((key, original))
                replacements[key] = original
            else:
                replacements[key] = match
        merged[customer_col] = merged[customer_col].map(
            lambda value: replacements.get(normalize_customer_name(value), value)
        )

    # Normalize the subjective vehicle-model entry. For each part number,
    # use the non-empty vehicle-model value from the latest year-month for
    # every row having that same part number.
    part_col = next((c for c in merged.columns if str(c).strip() == "품번"), None)
    vehicle_col = next((c for c in merged.columns if str(c).strip() == "차종"), None)
    period_col = "해당년월"
    if part_col and vehicle_col and period_col in merged.columns:
        work = merged[[part_col, vehicle_col, period_col]].copy()
        work[part_col] = work[part_col].astype("string").str.strip()
        work[vehicle_col] = work[vehicle_col].astype("string").str.strip()
        work[period_col] = work[period_col].astype("string")
        valid = work[work[part_col].notna() & (work[part_col] != "") & work[vehicle_col].notna() & (work[vehicle_col] != "")]
        if not valid.empty:
            latest = valid.sort_values(period_col).drop_duplicates(part_col, keep="last")
            latest_map = latest.set_index(part_col)[vehicle_col].to_dict()
            merged[vehicle_col] = work[part_col].map(latest_map).fillna(merged[vehicle_col])

    output.parent.mkdir(parents=True, exist_ok=True)
    merged.to_excel(output, index=False, engine="openpyxl")
    return len(files), len(merged), errors


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("검수현황 엑셀 파일 합치기")
        self.geometry("760x520")
        self.minsize(680, 460)
        self.folders: list[Path] = []
        self.input_files: list[Path] = []
        self.output_var = tk.StringVar(value=str(Path.cwd() / "검수현황_통합.xlsx"))
        self.status_var = tk.StringVar(value="월별 검수 폴더를 추가하세요.")
        self._build_ui()

    def _build_menu(self):
        menu = tk.Menu(self)
        upload = tk.Menu(menu, tearoff=False)
        upload.add_command(label="기본 DATA 등록", command=self.upload_initial_data)
        upload.add_command(label="매월 DATA 추가", command=self.append_monthly_data)
        menu.add_cascade(label="검수종합 현황 업로드", menu=upload)
        self.config(menu=menu)

    def _build_ui(self):
        main = ttk.Frame(self, padding=16)
        main.pack(fill="both", expand=True)
        title_row = ttk.Frame(main)
        title_row.pack(fill="x")
        ttk.Label(title_row, text="검수현황 엑셀 파일 합치기", font=("맑은 고딕", 16, "bold")).pack(side="left")
        ttk.Button(title_row, text="검수종합 현황 다운로드", command=self.download_summary).pack(side="right", padx=(8, 0))
        ttk.Button(title_row, text="검수종합 현황 업로드", command=self.show_upload_menu).pack(side="right")
        ttk.Label(main, text="파일명에서 년월을 읽어 '해당년월' 컬럼을 추가한 뒤 하나의 xlsx로 저장합니다.").pack(anchor="w", pady=(4, 14))

        folder_frame = ttk.LabelFrame(main, text="입력 폴더")
        folder_frame.pack(fill="both", expand=True)
        self.folder_list = tk.Listbox(folder_frame, height=8)
        self.folder_list.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        buttons = ttk.Frame(folder_frame)
        buttons.pack(side="right", fill="y", padx=8, pady=8)
        ttk.Button(buttons, text="여러 폴더 선택", command=self.add_multiple_folders).pack(fill="x", pady=2)
        ttk.Button(buttons, text="여러 파일 선택", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(buttons, text="폴더 하나 추가", command=self.add_folder).pack(fill="x", pady=2)
        ttk.Button(buttons, text="선택 삭제", command=self.remove_folder).pack(fill="x", pady=2)
        ttk.Button(buttons, text="전체 삭제", command=self.clear_folders).pack(fill="x", pady=2)

        ttk.Label(main, text="병합 결과는 임시 파일로 생성한 뒤 Microsoft Excel에서 자동으로 엽니다.", foreground="#555").pack(anchor="w", pady=(8, 12))

        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.pack(fill="x", pady=(4, 6))
        ttk.Label(main, textvariable=self.status_var).pack(anchor="w")
        self.run_button = ttk.Button(main, text="파일 합치기 실행", command=self.start_merge)
        self.run_button.pack(anchor="e", pady=14)

        note = ttk.Label(main, text="대상 파일명 예: 21년 01월 검수현황.xls\n파일명 패턴이 다른 파일은 자동으로 제외됩니다.", foreground="#555")
        note.pack(anchor="w", side="bottom")

    def show_upload_menu(self):
        menu = tk.Menu(self, tearoff=False)
        menu.add_command(label="기본 DATA 등록", command=self.upload_initial_data)
        menu.add_command(label="매월 DATA 추가", command=self.append_monthly_data)
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery() + 8
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def download_summary(self):
        if not SUMMARY_FILE.exists():
            messagebox.showwarning("DATA 없음", "먼저 검수종합 현황 DATA를 등록해 주세요.")
            return
        try:
            os.startfile(SUMMARY_FILE)
        except Exception as exc:
            messagebox.showerror("Excel 열기 오류", str(exc))

    def upload_initial_data(self):
        files = filedialog.askopenfilenames(
            title="기본 DATA로 등록할 파일 선택",
            filetypes=[("Excel 파일", "*.xls *.xlsx")],
        )
        if not files:
            return
        self.run_button.configure(state="disabled")
        threading.Thread(
            target=self._upload_worker,
            args=([Path(p) for p in files], SUMMARY_FILE, False),
            daemon=True,
        ).start()

    def append_monthly_data(self):
        if not SUMMARY_FILE.exists():
            messagebox.showwarning("기본 DATA 필요", "먼저 '기본 DATA 등록'으로 검수종합 현황을 만들어 주세요.")
            return
        files = filedialog.askopenfilenames(
            title="추가할 월별 DATA 파일 선택",
            filetypes=[("Excel 파일", "*.xls *.xlsx")],
        )
        if not files:
            return
        self.run_button.configure(state="disabled")
        threading.Thread(
            target=self._upload_worker,
            args=([Path(p) for p in files], SUMMARY_FILE, True),
            daemon=True,
        ).start()

    def _upload_worker(self, files: list[Path], output: Path, append: bool):
        temp_output = Path(tempfile.gettempdir()) / f"검수종합_추가_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        try:
            APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
            merge_files([], temp_output, self._progress, selected_files=files)
            new_data = pd.read_excel(temp_output)
            if append and output.exists():
                old_data = pd.read_excel(output)
                old_part_col = next((c for c in old_data.columns if re.sub(r"\s+", "", str(c)) == "품번"), None)
                if old_part_col:
                    old_data = old_data[old_data[old_part_col].notna() & (old_data[old_part_col].astype("string").str.strip() != "")].copy()
                result = pd.concat([old_data, new_data], ignore_index=True, sort=False)
            else:
                result = new_data
            result.to_excel(output, index=False, engine="openpyxl")
            self.after(0, lambda: self._upload_done(output, len(new_data), append))
        except Exception as exc:
            self.after(0, self._failed, str(exc))

    def _upload_done(self, output: Path, added_rows: int, append: bool):
        self.run_button.configure(state="normal")
        self.status_var.set(f"검수종합 현황 업데이트 완료: {added_rows:,}개 행")
        try:
            os.startfile(output)
            action = "추가" if append else "등록"
            messagebox.showinfo("업로드 완료", f"{action} 완료되었습니다.\n추가된 행: {added_rows:,}개\nExcel에서 파일을 열었습니다.")
        except OSError as exc:
            messagebox.showerror("Excel 열기 오류", f"파일은 저장되었지만 Excel을 열 수 없습니다.\n{exc}\n\n파일: {output}")

    def add_folder(self):
        folder = filedialog.askdirectory(title="검수현황 폴더 선택")
        if folder:
            self.input_files.clear()
            self.folders.clear()
            self.folder_list.delete(0, tk.END)
            path = Path(folder)
            if path not in self.folders:
                self.folders.append(path)
                self.folder_list.insert(tk.END, str(path))
                self.status_var.set(f"폴더 {len(self.folders)}개가 선택되었습니다.")

    def add_multiple_folders(self):
        """Choose a common parent and add all of its immediate subfolders."""
        parent = filedialog.askdirectory(title="검수 폴더들이 들어 있는 상위 폴더 선택")
        if not parent:
            return
        self.input_files.clear()
        self.folders.clear()
        self.folder_list.delete(0, tk.END)
        parent_path = Path(parent)
        candidates = sorted([p for p in parent_path.iterdir() if p.is_dir()], key=lambda p: p.name.lower())

        # Windows' folder picker returns the highlighted folder, even when
        # the user is browsing its parent. If a year folder was highlighted
        # (for example, 21년 검수), use its parent and collect sibling year
        # folders instead of incorrectly searching inside that year folder.
        if re.match(r"^\d{2}년\s*검수$", parent_path.name) or any(
            MONTH_FILE_RE.search(p.stem) for p in parent_path.iterdir() if p.is_file()
        ):
            parent_path = parent_path.parent
            candidates = sorted([p for p in parent_path.iterdir() if p.is_dir()], key=lambda p: p.name.lower())

        candidates = [p for p in candidates if re.match(r"^\d{2}년\s*검수$", p.name)]
        if not candidates:
            messagebox.showinfo("검수 폴더 없음", "선택한 위치에서 'YY년 검수' 폴더를 찾지 못했습니다.")
            return

        dialog = tk.Toplevel(self)
        dialog.title("필요한 검수 폴더 선택")
        dialog.geometry("620x420")
        dialog.transient(self)
        dialog.grab_set()
        ttk.Label(dialog, text=f"상위 폴더: {parent_path}").pack(anchor="w", padx=12, pady=(12, 4))
        ttk.Label(dialog, text="병합할 폴더만 체크한 뒤 '선택한 폴더 추가'를 누르세요.").pack(anchor="w", padx=12)
        checks_frame = ttk.Frame(dialog)
        checks_frame.pack(fill="both", expand=True, padx=12, pady=8)
        canvas = tk.Canvas(checks_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(checks_frame, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=e.width))
        selected_vars = []
        for path in candidates:
            var = tk.BooleanVar(value=False)
            selected_vars.append((path, var))
            ttk.Checkbutton(inner, text=path.name, variable=var).pack(anchor="w", padx=8, pady=3)

        actions = ttk.Frame(dialog)
        actions.pack(fill="x", padx=12, pady=(0, 12))

        def select_folders():
            selected = [path for path, var in selected_vars if var.get()]
            if not selected:
                messagebox.showwarning("폴더 선택 필요", "추가할 폴더를 하나 이상 체크하세요.", parent=dialog)
                return
            for path in selected:
                if path not in self.folders:
                    self.folders.append(path)
                    self.folder_list.insert(tk.END, str(path))
            self.status_var.set(f"폴더 {len(self.folders)}개가 선택되었습니다.")
            dialog.destroy()

        ttk.Button(actions, text="선택한 폴더 추가", command=select_folders).pack(side="right")
        ttk.Button(actions, text="취소", command=dialog.destroy).pack(side="right", padx=8)

    def remove_folder(self):
        selected = list(self.folder_list.curselection())
        for index in reversed(selected):
            self.folder_list.delete(index)
            if self.input_files:
                self.input_files.pop(index)
            else:
                self.folders.pop(index)

    def clear_folders(self):
        self.folders.clear()
        self.input_files.clear()
        self.folder_list.delete(0, tk.END)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="합칠 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xls *.xlsx"), ("모든 파일", "*.*")],
        )
        if not files:
            return
        self.folders.clear()
        self.input_files = [Path(p) for p in files if Path(p).suffix.lower() in {".xls", ".xlsx"}]
        self.folder_list.delete(0, tk.END)
        for path in self.input_files:
            self.folder_list.insert(tk.END, str(path))
        self.status_var.set(f"엑셀 파일 {len(self.input_files)}개가 선택되었습니다.")

    def choose_output(self):
        output = filedialog.asksaveasfilename(
            title="통합 파일 저장", defaultextension=".xlsx",
            filetypes=[("Excel 통합 문서", "*.xlsx")], initialfile="검수현황_통합.xlsx")
        if output:
            self.output_var.set(output)

    def start_merge(self):
        if not self.folders and not self.input_files:
            messagebox.showwarning("입력 필요", "폴더 또는 엑셀 파일을 하나 이상 추가하세요.")
            return
        output = Path(tempfile.gettempdir()) / f"검수현황_통합_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        self.run_button.configure(state="disabled")
        self.progress.configure(value=0)
        self.status_var.set("파일을 읽는 중입니다...")
        threading.Thread(target=self._worker, args=(output,), daemon=True).start()

    def _worker(self, output: Path):
        try:
            result = merge_files(self.folders, output, self._progress, self.input_files or None)
            self.after(0, self._done, output, result)
        except Exception as exc:
            self.after(0, self._failed, str(exc))

    def _progress(self, current, total, name):
        self.after(0, lambda: (self.progress.configure(maximum=total, value=current), self.status_var.set(f"{current}/{total} 처리 중: {name}")))

    def _done(self, output, result):
        files, rows, errors = result
        self.run_button.configure(state="normal")
        message = f"완료되었습니다.\n파일 {files}개, 행 {rows:,}개\nExcel에서 결과를 열었습니다."
        if errors:
            message += f"\n\n읽기 실패 {len(errors)}개:\n" + "\n".join(errors[:10])
        self.status_var.set(f"완료: {rows:,}개 행")
        try:
            os.startfile(output)
        except OSError as exc:
            message += f"\n\nExcel 자동 열기에 실패했습니다:\n{exc}\n파일: {output}"
        messagebox.showinfo("병합 완료", message)

    def _failed(self, error):
        self.run_button.configure(state="normal")
        self.status_var.set("오류가 발생했습니다.")
        messagebox.showerror("병합 오류", error)


if __name__ == "__main__":
    App().mainloop()
