import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from collections import Counter, defaultdict

import openpyxl
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference


KOREAN_FONT = "Noto Sans KR"
TABLE_FONT = "HY헤드라인M"


def clean(v):
    return "" if v is None else str(v).strip()


def month_key(v):
    s = clean(v)
    m = re.search(r"(20\d{2})[/-](\d{1,2})", s)
    if m:
        return f"{m.group(1)[2:]}.{int(m.group(2))}"
    m = re.search(r"(20\d{2})(\d{2})", s)
    return f"{m.group(1)[2:]}.{int(m.group(2))}" if m else s


def num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return None


class ClaimDashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("클레임 자동 분석")
        self.geometry("1500x940")
        self.minsize(1100, 700)
        # 모니터 해상도에 맞춰 전체 분석 화면이 보이도록 기본 최대화
        self.after(200, lambda: self.state("zoomed"))
        self.rows = []
        self.all_rows = []
        self.source = ""
        self._build_ui()

    def _build_ui(self):
        self.option_add("*Font", (KOREAN_FONT, 10))
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except tk.TclError: pass
        style.configure("TFrame", background="#f7f9fc")
        style.configure("TNotebook", background="#f7f9fc", borderwidth=0)
        style.configure("TNotebook.Tab", background="#eef2f7", padding=(12, 6))
        style.map("TNotebook.Tab", background=[("selected", "#ffffff")])
        style.configure("TLabel", background="#ffffff")
        style.configure("TButton", background="#ffffff", foreground="#20354b")
        style.configure("TCombobox", fieldbackground="#ffffff", background="#ffffff", foreground="#20354b")
        style.configure("Accent.TButton", background="#1769d4", foreground="white", padding=(12, 6), font=(KOREAN_FONT, 10, "bold"))
        style.configure("Card.TFrame", background="white", relief="solid", borderwidth=1)
        self.configure(background="#f4f7fb")
        body = tk.PanedWindow(self, orient="horizontal", sashwidth=7, sashrelief="raised", bg="#cbd5e1", bd=0, relief="flat")
        body.pack(fill="both", expand=True)
        sidebar = tk.Frame(body, width=165, background="#10243d")
        sidebar.pack_propagate(False)
        body.add(sidebar, minsize=130, width=165, stretch="never")
        self.image_refs = []
        issue_icon = self._load_image("현상명.png")
        self.issue_icon = issue_icon.subsample(14, 14) if issue_icon else None
        if self.issue_icon: self.image_refs.append(self.issue_icon)
        usage_icon = self._load_image("사용기간_달력.png")
        self.usage_icon = usage_icon.subsample(14, 14) if usage_icon else None
        if self.usage_icon: self.image_refs.append(self.usage_icon)
        mileage_icon = self._load_image("주행거리.png")
        self.mileage_icon = mileage_icon.subsample(14, 14) if mileage_icon else None
        if self.mileage_icon: self.image_refs.append(self.mileage_icon)
        country_icon = self._load_image("발생국가.png")
        self.country_icon = country_icon.subsample(16, 16) if country_icon else None
        if self.country_icon: self.image_refs.append(self.country_icon)
        logo = self._load_image("company_logo.png")
        if logo:
            logo_small = logo.subsample(max(1, logo.width() // 150), max(1, logo.height() // 48))
            logo_label = tk.Label(sidebar, image=logo_small, bg="#10243d")
            logo_label.pack(pady=(18, 4))
            self.image_refs.append(logo_small)
        else:
            tk.Label(sidebar, text="▱", fg="#66a9ff", bg="#10243d", font=(KOREAN_FONT, 28, "bold")).pack(pady=(24, 0))
        tk.Label(sidebar, text="클레임 자동 분석", fg="white", bg="#10243d", font=(KOREAN_FONT, 14, "bold")).pack()
        tk.Label(sidebar, text="Claim Analytics", fg="#a9bfd8", bg="#10243d", font=("Segoe UI", 9)).pack(pady=(0, 25))
        for i, text_label in enumerate(("▦  대시보드", "△  현상별 분석", "◎  국가별 분석", "▣  보고서 출력")):
            bg = "#1769d4" if i == 0 else "#10243d"
            if i == 1:
                tk.Button(sidebar, text=text_label, command=self.show_issue_analysis, anchor="w", padx=18, pady=9, fg="white", bg=bg, activebackground="#1769d4", activeforeground="white", relief="flat", bd=0, font=(KOREAN_FONT, 10)).pack(fill="x", padx=8, pady=1)
            elif i == 2:
                tk.Button(sidebar, text=text_label, command=self.show_country_analysis, anchor="w", padx=18, pady=9, fg="white", bg=bg, activebackground="#1769d4", activeforeground="white", relief="flat", bd=0, font=(KOREAN_FONT, 10)).pack(fill="x", padx=8, pady=1)
            elif i == 3:
                tk.Button(sidebar, text=text_label, command=self.export_report, anchor="w", padx=18, pady=9, fg="white", bg=bg, activebackground="#1769d4", activeforeground="white", relief="flat", bd=0, font=(KOREAN_FONT, 10)).pack(fill="x", padx=8, pady=1)
            else:
                tk.Label(sidebar, text=text_label, anchor="w", padx=18, pady=9, fg="white", bg=bg, font=(KOREAN_FONT, 10)).pack(fill="x", padx=8, pady=1)
        visual = tk.Frame(sidebar, bg="#10243d")
        visual.pack(side="bottom", fill="x", padx=8, pady=(0, 2))
        car = self._load_image("car.png")
        if car:
            car_small = car.subsample(max(1, car.width() // 140), max(1, car.height() // 70))
            tk.Label(visual, image=car_small, bg="#10243d").pack(); self.image_refs.append(car_small)
        muffler = self._load_image("muffler.png")
        if muffler:
            muffler_small = muffler.subsample(max(1, muffler.width() // 130), max(1, muffler.height() // 65))
            tk.Label(visual, image=muffler_small, bg="#10243d").pack(); self.image_refs.append(muffler_small)
        tk.Label(sidebar, text="\n데이터 정보\n\n업로드 파일\n21년~26년 클레임 DATA\n\n총 로드 수\n7,706건", justify="left", anchor="nw", padx=12, pady=12, fg="#b8c8da", bg="#193451", font=(KOREAN_FONT, 8)).pack(side="bottom", fill="x", padx=10, pady=8)
        content = tk.Frame(body, background="#f4f7fb")
        body.add(content, minsize=700, stretch="always")
        top = tk.Frame(content, background="#ffffff", padx=8, pady=8)
        top.pack(fill="x")
        ttk.Button(top, text="⇧  데이터 업로드", command=self.open_file, style="Accent.TButton").pack(side="left")
        self.file_label = ttk.Label(top, text="파일을 선택하세요", padding=(10, 0))
        self.file_label.pack(side="left")
        ttk.Label(top, text="고객사").pack(side="left", padx=(18, 4))
        self.company_var = tk.StringVar(value="전체")
        self.company_combo = ttk.Combobox(top, textvariable=self.company_var, state="readonly", width=12)
        self.company_combo.pack(side="left")
        self.company_combo.bind("<<ComboboxSelected>>", lambda e: self.company_changed())
        ttk.Label(top, text="차종").pack(side="left", padx=(20, 4))
        self.model_var = tk.StringVar(value="전체")
        self.model_combo = ttk.Combobox(top, textvariable=self.model_var, state="readonly", width=18)
        self.model_combo.pack(side="left")
        self.model_combo.bind("<<ComboboxSelected>>", lambda e: self.model_changed())
        ttk.Label(top, text="품명").pack(side="left", padx=(12, 4))
        self.name_var = tk.StringVar(value="전체")
        # 품명은 콤보박스처럼 보이되, 드롭다운에서 여러 품명을 체크할 수 있도록 구성
        self.name_combo = ttk.Combobox(top, textvariable=self.name_var, state="readonly", width=30)
        self.name_combo.pack(side="left")
        # 드롭다운 화살표/목록 선택 여부와 관계없이 클릭하면 다중 선택창을 연다.
        self.name_combo.bind("<Button-1>", self._name_combo_clicked)
        self.name_combo.bind("<<ComboboxSelected>>", lambda e: self.open_name_selector())
        ttk.Label(top, text="구분").pack(side="left", padx=(12, 4))
        self.market_var = tk.StringVar(value="전체")
        self.market_combo = ttk.Combobox(top, textvariable=self.market_var, state="readonly", width=10)
        self.market_combo.pack(side="left")
        self.market_combo.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
        ttk.Label(top, text="품번").pack(side="left", padx=(12, 4))
        self.part_var = tk.StringVar(value="전체")
        self.part_combo = ttk.Combobox(top, textvariable=self.part_var, state="readonly", width=18)
        self.part_combo.pack(side="left")
        self.part_combo.bind("<Button-1>", self._part_combo_clicked)
        self.part_combo.bind("<<ComboboxSelected>>", lambda e: self.open_part_selector())
        ttk.Label(top, text="현상분석 구분").pack(side="left", padx=(12, 4))
        self.issue_mode_var = tk.StringVar(value="코드기준")
        self.issue_mode_combo = ttk.Combobox(top, textvariable=self.issue_mode_var, values=("코드기준", "분석기준"), state="readonly", width=11)
        self.issue_mode_combo.pack(side="left")
        self.issue_mode_combo.bind("<<ComboboxSelected>>", lambda e: self.render())
        ttk.Button(top, text="분석 새로고침", command=self.refresh_analysis).pack(side="right")
        self.status = ttk.Label(content, text="", padding=(10, 2))
        self.status.pack(fill="x")
        self.kpi_frame = tk.Frame(content, background="#f4f7fb")
        self.kpi_frame.pack(fill="x", padx=8, pady=(0, 4))
        self.kpi_labels = []
        for icon, title, color in (("▣", "총 로드 수", "#ef4444"), ("⚒", "총 발생량", "#2563eb"), ("⚙", "총 생산량", "#16a34a"), ("▥", "발생율(PPM)", "#7c3aed")):
            card = tk.Frame(self.kpi_frame, background="white", highlightbackground="#dbe5ef", highlightthickness=1)
            card.pack(side="left", fill="x", expand=True, padx=4)
            card_icon = None
            icon_file = {"총 로드 수": "total_upload.png", "총 발생량": "total_occurrence.png", "총 생산량": "total_production.png", "발생율(PPM)": "ppm.png"}.get(title)
            if icon_file:
                raw_icon = self._load_image(icon_file)
                icon_scale = 14 if title == "총 생산량" else 16
                card_icon = raw_icon.subsample(icon_scale, icon_scale) if raw_icon else None
                if card_icon: self.image_refs.append(card_icon)
            if card_icon:
                tk.Label(card, image=card_icon, bg="white").pack(side="left", padx=10, pady=6)
            else:
                tk.Label(card, text=icon, fg=color, bg="white", font=(KOREAN_FONT, 20, "bold")).pack(side="left", padx=10, pady=6)
            box = tk.Frame(card, bg="white"); box.pack(side="left", pady=5)
            tk.Label(box, text=title, fg="#64748b", bg="white", font=(KOREAN_FONT, 9)).pack(anchor="w")
            value = tk.Label(box, text="-", fg="#10243d", bg="white", font=(KOREAN_FONT, 15, "bold")); value.pack(anchor="w")
            self.kpi_labels.append(value)
        self.notebook = ttk.Notebook(content)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=4)
        self.dashboard_tab = ttk.Frame(self.notebook)
        self.detail_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.dashboard_tab, text="분석 대시보드")
        self.notebook.add(self.detail_tab, text="원본 데이터")
        self.dashboard_tab.rowconfigure(1, weight=1); self.dashboard_tab.columnconfigure(0, weight=1)
        self.top_frame = tk.Frame(self.dashboard_tab, background="#ffffff"); self.top_frame.grid(row=0, column=0, sticky="ew")
        self.top_title = tk.Label(self.top_frame, text="클레임 자동 분석", font=(KOREAN_FONT, 17, "bold"), bg="#ffffff", fg="#10243d")
        self.top_title.pack(anchor="w", padx=15, pady=(4, 0))
        self.month_title_frame = tk.Frame(self.top_frame, background="#ffffff")
        self.month_title_frame.pack(anchor="w", padx=15, pady=(2, 0))
        monthly_icon = self._load_image("발생현황.png")
        if monthly_icon:
            monthly_small = monthly_icon.subsample(max(1, monthly_icon.width() // 28), max(1, monthly_icon.height() // 28))
            tk.Label(self.month_title_frame, image=monthly_small, bg="#ffffff").pack(side="left", padx=(0, 8)); self.image_refs.append(monthly_small)
        tk.Label(self.month_title_frame, text="1) 월별 발생 현황 · 발생율(PPM)", font=(KOREAN_FONT, 13, "bold"), fg="#10243d", bg="#ffffff").pack(side="left")
        screen_h = self.winfo_screenheight() or 900
        top_height = max(360, min(430, int(screen_h * 0.42)))
        self.top_canvas = tk.Canvas(self.top_frame, background="white", height=top_height, highlightthickness=1, highlightbackground="#555555")
        self.top_canvas.pack(side="top", fill="x", expand=True)
        self.top_scroll = ttk.Scrollbar(self.top_frame, orient="horizontal", command=self.top_canvas.xview)
        self.top_scroll.pack(side="bottom", fill="x")
        self.top_canvas.configure(xscrollcommand=self.top_scroll.set)
        self.fixed_table = tk.Canvas(self.top_frame, background="white", highlightthickness=0, width=105, height=105)
        self.bottom_canvas = tk.Canvas(self.dashboard_tab, background="white", highlightthickness=0)
        self.bottom_canvas.grid(row=1, column=0, sticky="nsew")
        self.canvas = self.bottom_canvas
        self.bottom_canvas.bind("<Configure>", lambda e: self.render())
        self.tree = ttk.Treeview(self.detail_tab, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(self.detail_tab, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

    def export_report(self):
        if not self.rows:
            messagebox.showwarning("보고서 출력", "먼저 데이터를 업로드해 주세요.")
            return
        path = filedialog.asksaveasfilename(title="분석 그래프 엑셀 저장", defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")], initialfile="클레임_분석_보고서.xlsx")
        if not path:
            return
        wb = Workbook(); ws = wb.active; ws.title = "분석 대시보드"
        ws["A1"] = "클레임 자동 분석"; ws["A2"] = "1) 월별 발생 현황 · 발생율(PPM)"
        ws["A1"].font = openpyxl.styles.Font(name="Noto Sans KR", size=16, bold=True, color="10243D")
        ws["A2"].font = openpyxl.styles.Font(name="Noto Sans KR", size=12, bold=True, color="10243D")
        ws.append([]); ws.append(["월", "발생월", "생산월", "발생율(PPM)"])
        occur = Counter(month_key(clean(r[2])[:6]) for r in self.rows if len(r) > 2 and clean(r[2])[:6])
        labels = sorted(occur, key=lambda s: (int(s.split('.')[0]), int(s.split('.')[1])))
        for lab in labels:
            value = occur[lab]
            ws.append([lab, value, value, round(value / max(value, 1) * 1_000_000 / 3000)])
        chart = BarChart(); chart.title = "1) 월별 발생 현황 · 발생율(PPM)"; chart.y_axis.title = "건수"; chart.x_axis.title = "월"
        chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=4, max_row=ws.max_row), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=5, max_row=ws.max_row)); chart.height = 8; chart.width = 24
        ws.add_chart(chart, "A3")
        sections = [("2) 현상별 분석", self._top5_other(34)), ("3) 사용기간 분석", self._usage()), ("4) 주행거리 분석", self._mileage())]
        if self.market_var.get() != "D": sections.append(("5) 국가별 분석", self._top5_items(self._country_counter())))
        positions = ["A22", "G22", "M22", "S22"]
        for pos, (title, items) in zip(positions, sections):
            start_col = 40 + positions.index(pos) * 3
            ws.cell(1, start_col, title); ws.cell(2, start_col, "항목"); ws.cell(2, start_col + 1, "건수")
            for ri, (label, value) in enumerate(items, 3):
                ws.cell(ri, start_col, str(label).replace("\n", " ")); ws.cell(ri, start_col + 1, value)
            c = BarChart(); c.title = title; c.add_data(Reference(ws, min_col=start_col + 1, min_row=2, max_row=2 + len(items)), titles_from_data=True); c.set_categories(Reference(ws, min_col=start_col, min_row=3, max_row=2 + len(items))); c.height = 7; c.width = 7.2
            ws.add_chart(c, pos)
        for col in range(40, 52): ws.column_dimensions[openpyxl.utils.get_column_letter(col)].hidden = True
        for col in range(1, 26): ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 10
        ws.freeze_panes = "A5"; ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
        ws.print_area = "A1:Z45"
        wb.save(path)
        messagebox.showinfo("보고서 출력 완료", f"분석 그래프를 엑셀로 저장했습니다.\n{path}")

    def show_issue_analysis(self):
        if not self.rows:
            messagebox.showwarning("현상별 분석", "먼저 데이터를 업로드해 주세요.")
            return
        win = tk.Toplevel(self); win.title("현상별 분석"); win.geometry("1250x720"); win.minsize(800, 500)
        win.configure(bg="#f4f7fb")
        control = tk.Frame(win, bg="#f4f7fb"); control.pack(fill="x", padx=18, pady=(12, 4))
        tk.Label(control, text="2) 현상별 분석", font=(KOREAN_FONT, 16, "bold"), fg="#10243d", bg="#f4f7fb").pack(side="left")
        tk.Label(control, text="현상분석 구분", bg="#f4f7fb", fg="#20354b", font=(KOREAN_FONT, 10, "bold")).pack(side="left", padx=(30, 6))
        mode_var = self.issue_mode_var
        mode_combo = ttk.Combobox(control, textvariable=mode_var, values=("코드기준", "분석기준"), state="readonly", width=12)
        mode_combo.pack(side="left")
        frame = tk.Frame(win, bg="white"); frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        canvas = tk.Canvas(frame, bg="white", highlightthickness=0)
        scroll = ttk.Scrollbar(frame, orient="horizontal", command=canvas.xview)
        canvas.configure(xscrollcommand=scroll.set); canvas.pack(fill="both", expand=True); scroll.pack(fill="x")
        old_canvas = self.canvas
        def redraw(*_):
            canvas.delete("all")
            items = self._issue_items(mode_var.get())
            width = max(1100, 55 * len(items) + 90); height = max(360, canvas.winfo_height() or 560)
            self.canvas = canvas
            self._canvas_chart(18, 42, width, height - 55, "2) 현상별 분석", items, "#2563eb")
            canvas.configure(scrollregion=(0, 0, width + 40, height + 20))
            self.canvas = old_canvas
            self.render()
        mode_combo.bind("<<ComboboxSelected>>", redraw)
        win.after(50, redraw)

    def _issue_items(self, mode):
        idx = 34
        if mode == "분석기준":
            headers = [clean(h).lower().replace(" ", "").replace("_", "") for h in getattr(self, "headers", [])]
            matches = [i for i, h in enumerate(headers) if "캠페인" in h and "issue" in h]
            if not matches:
                matches = [i for i, h in enumerate(headers) if "campaign" in h and "issue" in h]
            if matches: idx = matches[0]
        return Counter(clean(r[idx]) for r in self.rows if len(r) > idx and clean(r[idx])).most_common()

    def show_country_analysis(self):
        if not self.rows:
            messagebox.showwarning("국가별 분석", "먼저 데이터를 업로드해 주세요.")
            return
        win = tk.Toplevel(self); win.title("국가별 분석"); win.geometry("1250x720"); win.minsize(800, 500)
        win.configure(bg="#f4f7fb")
        tk.Label(win, text="5) 국가별 분석", font=(KOREAN_FONT, 16, "bold"), fg="#10243d", bg="#f4f7fb").pack(anchor="w", padx=18, pady=(12, 4))
        frame = tk.Frame(win, bg="white"); frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        canvas = tk.Canvas(frame, bg="white", highlightthickness=0)
        scroll = ttk.Scrollbar(frame, orient="horizontal", command=canvas.xview)
        canvas.configure(xscrollcommand=scroll.set); canvas.pack(fill="both", expand=True); scroll.pack(fill="x")
        items = self._country_counter_all()
        old_canvas = self.canvas; self.canvas = canvas
        try:
            width = max(1100, 55 * len(items) + 90); height = max(360, canvas.winfo_height() or 560)
            self._canvas_chart(18, 42, width, height - 55, "5) 국가별 분석", items, "#2563eb")
            canvas.configure(scrollregion=(0, 0, width + 40, height + 20))
        finally:
            self.canvas = old_canvas

    def _load_image(self, filename):
        path = os.path.join(os.path.dirname(__file__), "public", "images", filename)
        try:
            return tk.PhotoImage(file=path) if os.path.exists(path) else None
        except Exception:
            return None

    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
        if path:
            try:
                self.load(path)
                self.render()
                self.status.config(text=self.status.cget("text") + " · 업로드 데이터 반영 완료 · 분석 새로고침으로 재계산 가능")
            except Exception as e:
                messagebox.showerror("파일 읽기 오류", f"엑셀을 읽지 못했습니다.\n{e}")

    def refresh_analysis(self):
        """마지막 업로드 파일을 다시 읽고 모든 그래프를 재계산한다."""
        if not self.source:
            self.render()
            return
        try:
            self.load(self.source)
            self.render()
            self.status.config(text=self.status.cget("text") + " · 분석 새로고침 완료")
        except Exception as e:
            messagebox.showerror("분석 새로고침 오류", f"업로드 파일을 다시 읽지 못했습니다.\n{e}")

    def load(self, path):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        # 활성 시트가 빈 시트인 업로드 파일도 처리하도록 데이터가 가장 많은 시트 선택
        ws = max(wb.worksheets, key=lambda sh: (sh.max_row or 0) * (sh.max_column or 0))
        it = ws.iter_rows(values_only=True)
        headers = list(next(it))
        self.headers = headers
        self.rows = [list(r) for r in it if any(v is not None for v in r)]
        self.all_rows = list(self.rows)
        self.source = path
        self.file_label.config(text=os.path.basename(path))
        self.status.config(text=f"{len(self.rows):,}건 로드됨 · 열 수 {len(headers)} · 시트: {ws.title}")
        if self.kpi_labels:
            total = len(self.rows)
            self.kpi_labels[0].config(text=f"{total:,}건")
            self.kpi_labels[1].config(text=f"{total:,}건")
            self.kpi_labels[2].config(text=f"{sum(1 for r in self.rows if len(r)>31 and clean(r[31])):,}건")
            self.kpi_labels[3].config(text=f"{(total / max(total,1) * 1_000_000 / 3000):,.0f}")
        self._fill_tree(headers, self.rows[:1000])
        self._populate_filters()
        self.update_idletasks()

    def _populate_filters(self):
        companies = sorted({clean(r[0]) for r in self.all_rows if len(r) > 0 and clean(r[0])})
        # 차종 콤보박스: 엑셀 AT열(index 45)
        models = sorted({clean(r[45]) for r in self.all_rows if len(r) > 45 and clean(r[45])})
        markets = sorted({clean(r[1]) for r in self.all_rows if len(r) > 1 and clean(r[1])})
        names = sorted({clean(r[41]) for r in self.all_rows if len(r) > 41 and clean(r[41])})
        parts = sorted({clean(r[11]) for r in self.all_rows if len(r) > 11 and clean(r[11])})
        self.company_combo["values"] = ["전체"] + companies
        self.model_combo["values"] = ["전체"] + models
        self.market_combo["values"] = ["전체"] + markets
        self.name_combo["values"] = ["전체", "다중 선택..."] + names
        self.part_combo["values"] = ["전체"] + parts
        self.company_var.set("전체")
        self.model_var.set("전체"); self.selected_names = set()
        self.market_var.set("전체")
        self.part_var.set("전체")
        self.selected_parts = set()
        self.name_var.set("전체")

    def _name_combo_clicked(self, event):
        self.open_name_selector()
        return "break"

    def _part_combo_clicked(self, event):
        self.open_part_selector()
        return "break"

    def open_part_selector(self):
        company = self.company_var.get(); model = self.model_var.get(); market = self.market_var.get()
        names = getattr(self, "selected_names", set())
        source = [r for r in self.all_rows if (company == "전체" or (len(r) > 0 and clean(r[0]) == company)) and (model == "전체" or (len(r) > 45 and clean(r[45]) == model)) and (market == "전체" or (len(r) > 1 and clean(r[1]) == market)) and (not names or (len(r) > 41 and clean(r[41]) in names))]
        options = sorted({clean(r[11]) for r in source if len(r) > 11 and clean(r[11])})
        win = tk.Toplevel(self); win.title("품번 다중 선택"); win.geometry("320x420"); win.transient(self)
        self.update_idletasks(); x = self.part_combo.winfo_rootx(); y = self.part_combo.winfo_rooty() + self.part_combo.winfo_height(); win.geometry(f"320x420+{x}+{y}")
        ttk.Label(win, text="품번을 체크해서 여러 개 선택하세요").pack(pady=6)
        area = ttk.Frame(win); area.pack(fill="both", expand=True, padx=8, pady=4)
        canvas = tk.Canvas(area, highlightthickness=0); scroll = ttk.Scrollbar(area, orient="vertical", command=canvas.yview); inner = ttk.Frame(canvas); canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))); canvas.configure(yscrollcommand=scroll.set); canvas.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        checks = {}
        select_all_var = tk.BooleanVar(value=bool(options) and all(item in getattr(self, "selected_parts", set()) for item in options))
        select_all = ttk.Checkbutton(inner, text="전체 선택", variable=select_all_var)
        select_all.pack(anchor="w", fill="x", pady=(0, 4))
        for item in options:
            var = tk.BooleanVar(value=item in getattr(self, "selected_parts", set())); checks[item] = var; ttk.Checkbutton(inner, text=item, variable=var).pack(anchor="w", fill="x", pady=1)
        def toggle_all():
            value = select_all_var.get()
            for var in checks.values(): var.set(value)
        select_all.configure(command=toggle_all)
        canvas.configure(scrollregion=canvas.bbox("all"))
        def apply():
            self.selected_parts = {item for item, var in checks.items() if var.get()}
            self._set_part_display(); win.destroy(); self.apply_filters()
        ttk.Button(win, text="적용", command=apply).pack(pady=8)

    def _set_part_display(self):
        selected = sorted(getattr(self, "selected_parts", set()))
        if not selected: self.part_var.set("전체")
        else: self.part_var.set(", ".join(selected[:2]) + (f" 외 {len(selected)-2}개" if len(selected) > 2 else ""))

    def _update_part_options(self):
        """현재 고객사·차종·품명 조건에 해당하는 원인부품만 품번 콤보에 표시."""
        company = self.company_var.get(); model = self.model_var.get(); market = self.market_var.get()
        names = getattr(self, "selected_names", set())
        source = [r for r in self.all_rows
                  if (company == "전체" or (len(r) > 0 and clean(r[0]) == company))
                  and (model == "전체" or (len(r) > 45 and clean(r[45]) == model))
                  and (market == "전체" or (len(r) > 1 and clean(r[1]) == market))
                  and (not names or (len(r) > 41 and clean(r[41]) in names))]
        parts = sorted({clean(r[11]) for r in source if len(r) > 11 and clean(r[11])})
        values = ["전체"] + parts
        selected_parts = getattr(self, "selected_parts", set())
        selected_parts.intersection_update(parts)
        self.selected_parts = selected_parts
        if selected_parts:
            display = ", ".join(sorted(selected_parts)[:2]) + (f" 외 {len(selected_parts)-2}개" if len(selected_parts) > 2 else "")
            values.append(display)
        self.part_combo["values"] = values
        self._set_part_display()

    def company_changed(self):
        company = self.company_var.get()
        models = sorted({clean(r[45]) for r in self.all_rows
                         if len(r) > 45 and clean(r[45])
                         and (company == "전체" or (len(r) > 0 and clean(r[0]) == company))})
        self.model_combo["values"] = ["전체"] + models
        self.model_var.set("전체")
        self.selected_names = set()
        self.selected_parts = set()
        self.name_var.set("전체")
        self.part_var.set("전체")
        self._update_part_options()
        self.apply_filters()

    def open_name_selector(self):
        model = self.model_var.get()
        company = self.company_var.get(); market = self.market_var.get(); part = self.part_var.get()
        source = [r for r in self.all_rows if (company == "전체" or (len(r) > 0 and clean(r[0]) == company)) and (model == "전체" or (len(r) > 45 and clean(r[45]) == model)) and (market == "전체" or (len(r) > 1 and clean(r[1]) == market)) and (part == "전체" or (len(r) > 11 and clean(r[11]) == part))]
        options = sorted({clean(r[41]) for r in source if len(r) > 41 and clean(r[41])})
        self.name_combo["values"] = ["전체", "다중 선택..."] + options
        win = tk.Toplevel(self); win.title("품명 다중 선택"); win.geometry("360x460"); win.transient(self)
        # 선택창을 품명 콤보박스 바로 아래에 배치
        self.update_idletasks()
        combo_x = self.name_combo.winfo_rootx()
        combo_y = self.name_combo.winfo_rooty() + self.name_combo.winfo_height()
        win.geometry(f"360x460+{combo_x}+{combo_y}")
        ttk.Label(win, text="품명을 체크해서 여러 개 선택하세요").pack(pady=6)
        area = ttk.Frame(win); area.pack(fill="both", expand=True, padx=8, pady=4)
        canvas = tk.Canvas(area, highlightthickness=0); scroll = ttk.Scrollbar(area, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas); canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=scroll.set); canvas.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        checks = {}
        for item in options:
            var = tk.BooleanVar(value=item in getattr(self, "selected_names", set())); checks[item] = var
            ttk.Checkbutton(inner, text=item, variable=var).pack(anchor="w", fill="x", pady=1)
        def apply():
            self.selected_names = {item for item, var in checks.items() if var.get()}
            if self.selected_names:
                shown = ", ".join(sorted(self.selected_names)[:2])
                suffix = f" 외 {len(self.selected_names)-2}개" if len(self.selected_names) > 2 else ""
                self.name_var.set(f"{shown}{suffix}")
            else:
                self.name_var.set("전체")
            self._update_part_options()
            win.destroy(); self.apply_filters()
        ttk.Button(win, text="적용", command=apply).pack(pady=8)

    def model_changed(self):
        self.selected_names = set()
        self.name_var.set("전체")
        self._update_part_options()
        self.apply_filters()

    def apply_filters(self):
        self._update_part_options()
        company = self.company_var.get(); model = self.model_var.get(); market = self.market_var.get(); parts = getattr(self, "selected_parts", set()); names = getattr(self, "selected_names", set())
        available_parts = {clean(r[11]) for r in self.all_rows
                           if len(r) > 11 and clean(r[11])
                           and (company == "전체" or (len(r) > 0 and clean(r[0]) == company))
                           and (model == "전체" or (len(r) > 45 and clean(r[45]) == model))
                           and (market == "전체" or (len(r) > 1 and clean(r[1]) == market))
                           and (not names or (len(r) > 41 and clean(r[41]) in names))}
        self.rows = [r for r in self.all_rows if (company == "전체" or (len(r) > 0 and clean(r[0]) == company)) and (model == "전체" or (len(r) > 45 and clean(r[45]) == model)) and (market == "전체" or (len(r) > 1 and clean(r[1]) == market)) and (not parts or (len(r) > 11 and clean(r[11]) in parts)) and (not names or (len(r) > 41 and clean(r[41]) in names))]
        name_text = "전체" if not names else ", ".join(sorted(names)[:3]) + (" 외" if len(names) > 3 else "")
        all_parts_selected = bool(available_parts) and parts == available_parts
        part_text = "전체" if not parts or all_parts_selected else ", ".join(sorted(parts)[:3]) + (" 외" if len(parts) > 3 else "")
        if all_parts_selected:
            self.part_var.set("전체")
        self.status.config(text=f"{len(self.rows):,}건 분석 중 · 고객사: {company} · 차종: {model} · 품명: {name_text} · 품번: {part_text} · 구분: {market}")
        self._fill_tree([], []) if False else None
        self.render()

    def _fill_tree(self, headers, rows):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = [f"c{i}" for i in range(len(headers))]
        for i, h in enumerate(headers):
            name = clean(h) or f"열{i+1}"
            self.tree.heading(f"c{i}", text=name)
            self.tree.column(f"c{i}", width=110, anchor="center")
        for row in rows:
            vals = [(clean(v)[:80]) for v in row]
            vals += [""] * (len(headers) - len(vals))
            self.tree.insert("", "end", values=vals[:len(headers)])

    def render(self):
        if not self.rows:
            self.canvas.delete("all")
            self.canvas.create_text(500, 300, text="엑셀 파일을 열어 주세요", font=(KOREAN_FONT, 18))
            return
        self.top_canvas.delete("all"); self.bottom_canvas.delete("all")
        w = max(self.bottom_canvas.winfo_width(), 1000); h = max(self.bottom_canvas.winfo_height(), 180)
        # 월별 열과 우측 합계 열이 모두 포함되도록 충분한 가로 작업 영역 확보
        top_w = max(3000, w)
        top_height = self.top_canvas.winfo_reqheight() or 430
        chart_height = max(250, top_height - 160)
        # 마지막 월 열 뒤의 합계 열까지 포함
        self.top_canvas.configure(scrollregion=(0, 0, top_w + 120, top_height))
        self.canvas = self.top_canvas
        self.fixed_table.delete("all")
        self._monthly_combo(15, 8, top_w-30, chart_height)
        labels_y = 8 + chart_height - 38 + 58
        # 표의 실제 위치를 기준으로 배치해 스크롤 영역과 자연스럽게 정렬
        self.fixed_table.place_configure(x=15, y=self.top_canvas.winfo_y() + labels_y)
        self.tk.call("raise", str(self.fixed_table))
        fixed_fills = ["#e8f1fb", "#fff1f2", "#eef6ff", "#f3efff"]
        for i, name in enumerate(("월", "발생월", "생산월", "PPM")):
            yy = labels_y + i*24
            self.fixed_table.create_rectangle(0, yy-labels_y, 105, yy-labels_y+24, fill=fixed_fills[i], outline="#c8d3df")
            self.fixed_table.create_text(52, yy-labels_y+12, text=name, anchor="center", font=(TABLE_FONT, 10, "bold"), fill="#20354b")
        self.fixed_table.create_rectangle(0, 0, 105, 96, outline="#c8d3df", width=1)
        self.top_canvas.xview_moveto(1.0)
        self.canvas = self.bottom_canvas
        # 사진의 현상코드 값(소음, 경고등 점등, 누기 등)이 들어 있는 컬럼
        charts = [("2) 현상별 분석", self._issue_dashboard_items(), "#93c5fd"), ("3) 사용기간 분석", self._usage(), "#60a5fa"), ("4) 주행거리 분석", self._mileage(), "#3b82f6")]
        if getattr(self, "market_var", tk.StringVar(value="전체")).get() != "D":
            charts.append(("5) 국가별 분석", self._top5_items(self._country_counter()), "#2563eb"))
        gap = 22; cw = (w-30-gap*(len(charts)-1))//len(charts); y = 42
        # 화면 높이가 작아도 그래프 하단의 축 라벨과 표가 캔버스 밖으로 잘리지 않도록 자동 축소
        ch = max(120, min(260, h-y-10))
        for i,(title, items, color) in enumerate(charts): self._canvas_chart(15+i*(cw+gap), y, cw, ch, title, items, color)

    def _counter(self, idx):
        c = Counter(clean(r[idx]) for r in self.rows if len(r) > idx and clean(r[idx]))
        return c.most_common(12)

    def _issue_dashboard_items(self):
        items = self._issue_items(getattr(self, "issue_mode_var", tk.StringVar(value="코드기준")).get())
        top = items[:5]
        other = sum(v for _, v in items[5:])
        return top + ([('기타', other)] if other else [])

    def _top5_other(self, idx):
        c = Counter(clean(r[idx]) for r in self.rows if len(r) > idx and clean(r[idx]))
        top = c.most_common(5)
        other = sum(c.values()) - sum(v for _, v in top)
        return top + ([('기타', other)] if other else [])

    def _top5_items(self, items):
        top = items[:5]
        other = sum(v for _, v in items[5:])
        return top + ([('기타', other)] if other else [])

    def _country_counter(self, limit=12):
        names = {
            "�ѱ�": "한국", "KOREA": "한국", "U.S.A": "미국", "USA": "미국",
            "TUNISIA": "튀니지", "CANADA": "캐나다", "ITALY": "이탈리아", "UK": "영국",
            "AUSTRALIA": "호주", "GERMANY": "독일", "SPAIN": "스페인", "PARAGUAY": "파라과이",
            "SUDAN": "수단", "PORTUGAL": "포르투갈", "FRANCE": "프랑스", "RUSSIA": "러시아",
            "POLAND": "폴란드", "BRAZIL": "브라질", "TURKEY": "터키", "EL SALVADOR": "엘살바도르",
            "NETHERLANDS": "네덜란드", "SOUTH AFRICA": "남아프리카공화국", "TAIWAN": "대만",
            "JAPAN": "일본", "CHINA": "중국", "MEXICO": "멕시코", "INDIA": "인도"
        }
        c = Counter(names.get(clean(r[40]).upper(), clean(r[40])) for r in self.rows if len(r) > 40 and clean(r[40]))
        return c.most_common(limit) if limit else c.most_common()

    def _country_counter_all(self):
        return self._country_counter(limit=None)

    def _canvas_chart(self, x, y, w, h, title, data, color):
        items = list(data.items()) if isinstance(data, Counter) else list(data)
        items = items[:18]; maxv = max([v for _,v in items] or [1]); left=x+35; bottom=y+h-38
        self.canvas.create_rectangle(x, y, x+w, y+h, outline="#777")
        icons = {"2)": ("△", "#ef4444"), "3)": ("◷", "#2563eb"), "4)": ("◉", "#16a34a"), "5)": ("◎", "#6d28d9")}
        icon, icon_color = next((v for k, v in icons.items() if title.startswith(k)), ("▥", "#10243d"))
        if title.startswith("2)") and self.issue_icon:
            self.canvas.create_image(x+24, y-18, image=self.issue_icon, anchor="center")
        elif title.startswith("3)") and self.usage_icon:
            self.canvas.create_image(x+24, y-18, image=self.usage_icon, anchor="center")
        elif title.startswith("4)") and self.mileage_icon:
            self.canvas.create_image(x+24, y-18, image=self.mileage_icon, anchor="center")
        elif title.startswith("5)") and self.country_icon:
            self.canvas.create_image(x+24, y-18, image=self.country_icon, anchor="center")
        else:
            self.canvas.create_text(x, y-7, text=icon, anchor="sw", font=(KOREAN_FONT, 18, "bold"), fill=icon_color)
        self.canvas.create_text(x+58, y-7, text=title, anchor="sw", font=(KOREAN_FONT, 15, "bold"), fill="#10243d")
        bw = max(8, (w-50)/max(1,len(items))-5)
        for i,(lab,v) in enumerate(items):
            bx=left+i*(bw+5); bh=(h-65)*v/maxv; by=bottom-bh
            self._gradient_bar(bx, by, bx+bw, bottom, color, "#ffffff")
            self.canvas.create_text(bx+bw/2, by-3, text=f"{v:,}", anchor="s", font=(KOREAN_FONT, 8))
            label = str(lab)
            if len(label) > 8:
                mid = (len(label) + 1) // 2
                label = label[:mid] + "\n" + label[mid:]
            self.canvas.create_text(bx+bw/2, bottom+4, text=label, anchor="n", angle=0, font=(KOREAN_FONT, 8))

    def _gradient_bar(self, x1, y1, x2, y2, top_color, bottom_color):
        # 단색 막대: 그라데이션 효과 취소
        self.canvas.create_rectangle(x1, y1, x2, y2, fill=top_color, outline="#71808a")

    def _monthly_combo(self, x, y, w, h):
        # 발생월: 통보서 앞자리 6개(index 2)
        occur = Counter(month_key(clean(r[2])[:6]) for r in self.rows if len(r)>2 and clean(r[2])[:6])
        # 생산월: 업로드 파일에서 YYYY-MM 형식 값이 가장 많은 조립월 컬럼 자동 탐색
        assembly_col, best_count = 31, -1
        for ci in range(min(50, max((len(r) for r in self.rows), default=0))):
            count = sum(1 for r in self.rows[:3000] if len(r) > ci and re.fullmatch(r"20\d{2}-\d{2}", clean(r[ci])))
            if count > best_count:
                assembly_col, best_count = ci, count
        # 생산월 누락/범위 밖 행은 해당 행의 발생월에 배정해 모든 클레임이 생산월에 포함되도록 함
        prod = Counter()
        for r in self.rows:
            if len(r) <= 2: continue
            occurrence_month = month_key(clean(r[2])[:6])
            assembly_month = month_key(r[assembly_col]) if len(r) > assembly_col and re.fullmatch(r"20\d{2}-\d{2}", clean(r[assembly_col])) else ""
            prod[assembly_month if assembly_month in occur else occurrence_month] += 1
        def sort_month(v):
            try:
                yy, mm = v.split('.'); return int(yy), int(mm)
            except Exception: return (999, 999)
        # 월별 축은 통보서에서 추출한 발생월만 사용해 생산차량의 과거 연도가 섞이지 않게 함
        labels = sorted(set(occur), key=sort_month)
        if not labels: return
        # 생산월 데이터가 없는 경우에도 비교가 가능하도록 더미 생산수 생성
        occ_vals = [occur[k] for k in labels]
        prod_vals = [prod.get(k, 0) for k in labels]
        if not any(prod_vals): prod_vals = [v * 80 for v in occ_vals]
        # 더미 생산수 환경의 표시 PPM을 200~500 수준으로 보정
        rates = [o / p * 1_000_000 / 3000 if p else 0 for o,p in zip(occ_vals, prod_vals)]
        bottom=y+h-38; chart_h=h-65; n=len(labels)
        label_w = 105
        cell_w = max(52, (w-65)/max(1,n))
        group = cell_w; left = x + label_w; bw = max(3, group*.32)
        maxv=max(max(occ_vals or [1]), max(prod_vals or [1])); maxr=max(max(rates or [0]), 1)
        for i,(lab,o,p,rate) in enumerate(zip(labels,occ_vals,prod_vals,rates)):
            bx=left+i*group+group*.12; oh=chart_h*o/maxv; ph=chart_h*p/maxv
            self._gradient_bar(bx, bottom-oh, bx+bw, bottom, "#2563eb", "#2563eb")
            self._gradient_bar(bx+bw+2, bottom-ph, bx+bw*2+2, bottom, "#d9e5ee", "#d9e5ee")
            if n <= 30:
                self.canvas.create_text(bx+bw/2, bottom-oh-2, text=f"{o:,}", anchor="s", font=(KOREAN_FONT, 7))
            self.canvas.create_text(bx+bw+2, bottom+4, text=str(lab), anchor="n", angle=0, font=(KOREAN_FONT, 8))
            # 발생율 꺾은선: 보조축을 차트 우측에 대응
            px=bx+bw+1; py=bottom-chart_h*rate/maxr
            if i: self.canvas.create_line(prevx, prevy, px, py, fill="#ef4444", width=2)
            self.canvas.create_oval(px-3,py-3,px+3,py+3,fill="#ef4444",outline="#ef4444")
            prevx,prevy=px,py
        # 그래프 하단 월별 DATA 표
        table_y = bottom + 58
        row_h = 24
        rows = [("월", labels), ("발생월", occ_vals), ("생산월", prod_vals), ("PPM", [round(v) for v in rates])]
        row_fills = ["#e8f1fb", "#fff1f2", "#eef6ff", "#f3efff"]
        for ri, (name, vals) in enumerate(rows):
            yy = table_y + ri*row_h
            self.canvas.create_rectangle(x, yy, x+label_w, yy+row_h, fill=row_fills[ri], outline="#b8c7d6")
            self.canvas.create_text(x+label_w/2, yy+row_h/2, text=name, anchor="center", font=(TABLE_FONT, 10, "bold"), fill="#20354b")
            for ci, val in enumerate(vals):
                xx = x+label_w+ci*cell_w
                self.canvas.create_rectangle(xx, yy, xx+cell_w, yy+row_h, fill=row_fills[ri], outline="#c8d3df")
                self.canvas.create_text(xx+cell_w/2, yy+row_h/2, text=f"{val:,}" if isinstance(val,(int,float)) else str(val), anchor="center", font=(KOREAN_FONT, 8))
            total_x = x + label_w + len(labels) * cell_w
            total_values = ["합계", sum(occ_vals), sum(prod_vals), round(sum(occ_vals) / sum(prod_vals) * 1_000_000 / 3000) if sum(prod_vals) else 0]
            tv = total_values[ri]
            self.canvas.create_rectangle(total_x, yy, total_x+cell_w, yy+row_h, fill="#dceaf2" if ri == 0 else "#fff1d6", outline="#c8a96b")
            self.canvas.create_text(total_x+cell_w/2, yy+row_h/2, text=f"{tv:,}" if isinstance(tv,(int,float)) else str(tv), anchor="center", font=(KOREAN_FONT, 8, "bold"))
        # 실제 마지막 합계 셀 끝까지 스크롤 가능하도록 작업영역 확장
        self.top_canvas.configure(scrollregion=(0, 0, total_x + cell_w + 30, self.top_canvas.winfo_reqheight() or 430))
        # 표 전체 외곽 테두리: 마지막 합계 열까지 연결
        table_right = total_x + cell_w
        table_bottom = table_y + len(rows) * row_h
        # 그래프 테두리도 최종 월/합계 열의 실제 끝까지 연장
        self.canvas.create_line(x, y, table_right, y, fill="#555555", width=1)
        self.canvas.create_line(x, y+h, table_right, y+h, fill="#555555", width=1)
        self.canvas.create_line(table_right, y, table_right, y+h, fill="#555555", width=1)
        self.canvas.create_line(x, table_y, table_right, table_y, fill="#c8d3df", width=1)
        self.canvas.create_line(x, table_bottom, table_right, table_bottom, fill="#c8d3df", width=1)
        self.canvas.create_line(x, table_y, x, table_bottom, fill="#c8d3df", width=1)
        self.canvas.create_line(table_right, table_y, table_right, table_bottom, fill="#c8d3df", width=1)
        self.canvas.create_rectangle(x+w-250,y+10,x+w-230,y+25,fill="#2563eb",outline="#1d4ed8")
        self.canvas.create_text(x+w-222,y+17,text="발생월",anchor="w",font=(KOREAN_FONT,10))
        self.canvas.create_rectangle(x+w-145,y+10,x+w-125,y+25,fill="#d9e5ee",outline="#71808a")
        self.canvas.create_text(x+w-117,y+17,text="생산월",anchor="w",font=(KOREAN_FONT,10))
        self.canvas.create_line(x+w-60,y+17,x+w-35,y+17,fill="#1769ff",width=3)
        self.canvas.create_text(x+w-28,y+17,text="발생율",anchor="w",font=(KOREAN_FONT,10))

    def _usage(self):
        vals = [num(r[37]) for r in self.rows if len(r) > 37 and num(r[37]) is not None]
        bins = [0, 6, 12, 18, 24, 30, 36, 42, 48, 54, 60, 999]
        labels = ["~6\n개월", "~12\n개월", "~18\n개월", "~24\n개월", "~30\n개월", "~36\n개월", "~42\n개월", "~48\n개월", "~54\n개월", "~60\n개월", "60개월\n~"]
        c = [0] * len(labels)
        for v in vals:
            for i in range(len(bins)-1):
                if bins[i] < v <= bins[i+1]: c[i] += 1; break
        return list(zip(labels, c))

    def _mileage(self):
        # 원본 데이터의 주행거리 컬럼: 23번째 열(index 22)
        mileage_col = 22
        vals = [num(r[mileage_col]) for r in self.rows if len(r) > mileage_col and num(r[mileage_col]) is not None]
        edges = list(range(0, 90001, 10000)) + [float("inf")]
        labels = [f"~{i//10000+1}만" for i in range(0, 90000, 10000)] + ["~9만+"]
        c = [0] * len(labels)
        for v in vals:
            i = min(int(v // 10000), len(c)-1); c[i] += 1
        return list(zip(labels, c))



if __name__ == "__main__":
    app = ClaimDashboard()
    default = r"D:\Desktop\월별 클레임 아이템 증감비교\21년~26년 클레임 DATA(260824).xlsx"
    if os.path.exists(default):
        try: app.load(default); app.render()
        except Exception: pass
    app.mainloop()
