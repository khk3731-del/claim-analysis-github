import os
import json
import re
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from collections import Counter, defaultdict

import openpyxl
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from claim_engine import ClaimDataEngine
from dashboard_view import DashboardView


KOREAN_FONT = "Malgun Gothic"
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


class VirtualCostTable(tk.Frame):
    """현재 화면에 보이는 셀만 그리는 비용현황 가상화 표."""
    def __init__(self, parent):
        super().__init__(parent, bg="white", highlightbackground="#AAB8C8", highlightthickness=1)
        self.canvas = tk.Canvas(self, bg="white", highlightthickness=0)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self._yview)
        self.hbar = ttk.Scrollbar(self, orient="horizontal", command=self._xview)
        self.canvas.grid(row=0, column=0, sticky="nsew"); self.vbar.grid(row=0, column=1, sticky="ns"); self.hbar.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1); self.columnconfigure(0, weight=1)
        self.canvas.configure(yscrollcommand=self.vbar.set, xscrollcommand=self.hbar.set)
        self.canvas.bind("<Configure>", lambda e: self.redraw())
        self.headers=[]; self.rows=[]; self.widths=[]; self.row_h=30; self.head_h=38

    def set_data(self, headers, rows):
        self.headers, self.rows = headers, rows
        self.widths=[140 if i == 0 else (180 if i == 1 else 120) for i in range(len(headers))]
        self.canvas.configure(scrollregion=(0, 0, sum(self.widths), self.head_h + len(rows)*self.row_h)); self.redraw()

    def _xview(self, *args):
        self.canvas.xview_scroll(int(args[1])*20, args[2]) if args and args[0] == "scroll" else self.canvas.xview(*args); self.redraw()

    def _yview(self, *args):
        self.canvas.yview(*args); self.redraw()

    def redraw(self):
        if not self.headers: return
        self.canvas.delete("all"); x0=self.canvas.canvasx(0); y0=self.canvas.canvasy(0); cw=self.canvas.winfo_width(); ch=self.canvas.winfo_height()
        first=max(0, int(max(0,y0-self.head_h)//self.row_h)); last=min(len(self.rows), int(max(0,y0+ch-self.head_h)//self.row_h)+2)
        x=0
        for c, header in enumerate(self.headers):
            w=self.widths[c]
            if x+w >= x0 and x <= x0+cw:
                self.canvas.create_rectangle(x,0,x+w,self.head_h,fill="#173F6B",outline="#AAB8C8")
                self.canvas.create_text(x+w/2,self.head_h/2,text=str(header),fill="white",font=(KOREAN_FONT,10,"bold"))
            x += w
        # 원본 보고서처럼 구분 열의 연속된 행을 하나의 병합 셀로 표시합니다.
        spans = []
        start = None
        current_value = ""
        for index, row in enumerate(self.rows + [["", ""]]):
            value = str(row[0]) if row else ""
            if value != current_value:
                if start is not None:
                    spans.append((start, index, current_value))
                current_value = value
                start = index if value else None
        if start is not None:
            spans.append((start, len(self.rows), current_value))
        for start, end, value in spans:
            y1 = self.head_h + start * self.row_h; y2 = self.head_h + end * self.row_h
            self.canvas.create_rectangle(0, y1, self.widths[0], y2, fill="#F7FAFC", outline="#AAB8C8")
            self.canvas.create_text(self.widths[0] / 2, (y1 + y2) / 2, text=str(value), fill="#243B53", font=(KOREAN_FONT, 10), width=self.widths[0] - 12)
        for r in range(first,last):
            y=self.head_h+r*self.row_h; x=0; fill="#F7FAFC" if r%2==0 else "white"
            for c,value in enumerate(self.rows[r]):
                if c == 0:
                    x += self.widths[0]
                    continue
                w=self.widths[c]
                if x+w >= x0 and x <= x0+cw:
                    self.canvas.create_rectangle(x,y,x+w,y+self.row_h,fill=fill,outline="#D5DEE8")
                    self.canvas.create_text(x+w/2,y+self.row_h/2,text=str(value),fill="#243B53",font=(KOREAN_FONT,10))
                x += w
        for start, end, value in spans:
            y1 = self.head_h + start * self.row_h; y2 = self.head_h + end * self.row_h
            self.canvas.create_rectangle(0, y1, self.widths[0], y2, fill="#F7FAFC", outline="#AAB8C8")
            self.canvas.create_text(self.widths[0] / 2, (y1 + y2) / 2, text=str(value), fill="#243B53", font=(KOREAN_FONT, 10), width=self.widths[0] - 12)


class ClaimDashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("클레임 자동 분석")
        self.geometry("1700x980")
        self.minsize(1100, 700)
        # 모니터 해상도에 맞춰 전체 분석 화면이 보이도록 기본 최대화
        self.after(200, lambda: self.state("zoomed"))
        self.rows = []
        self.all_rows = []
        self.engine = ClaimDataEngine()
        self.source = ""
        self.cost_source = ""
        self._restore_cost_source()
        self._inspection_assembly_cache = None
        self._render_job = None
        self._build_ui()
        self._restore_saved_data()
        # 초기 화면을 막지 않고 집계 캐시를 즉시 백그라운드에서 준비한다.
        self.after(0, self._preload_customer_assembly)

    def _preload_customer_assembly(self):
        """Warm common customer assembly caches without blocking the UI thread."""
        def worker():
            for company in ("전체", "WIA", "기아", "현대", "HMC"):
                try:
                    self._inspection_assembly_by_month(company)
                except Exception:
                    pass
        threading.Thread(target=worker, daemon=True).start()

    @property
    def _saved_data_path(self):
        # GitHub에 함께 커밋할 수 있는 프로젝트 내 영구 저장 파일
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "claim_dashboard_data.json")

    @property
    def _saved_cost_source_path(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "cost_dashboard_source.json")

    def _restore_cost_source(self):
        try:
            with open(self._saved_cost_source_path, "r", encoding="utf-8") as f:
                path = json.load(f).get("path", "")
            if path and os.path.exists(path):
                self.cost_source = path
                return
        except Exception:
            pass
        # 이전 버전에서 업로드해 경로 저장이 없었던 경우에도 사용자가 제공한
        # 비용현황 원본을 자동으로 찾아 기존 DATA를 이어서 사용합니다.
        fallback_dir = r"W:\김훈기\김훈기\2.클레임 업무\★클레임 자료"
        explicit = os.path.join(fallback_dir, "클레임 금액 종합현황(수정)REV13.xlsx")
        if os.path.exists(explicit):
            candidates = [explicit]
        else:
            candidates = []
        if os.path.isdir(fallback_dir):
            candidates += [os.path.join(fallback_dir, name) for name in os.listdir(fallback_dir)
                          if name.startswith("클레임 금액 종합현황") and name.lower().endswith((".xlsx", ".xlsm", ".xls"))]
            if candidates:
                self.cost_source = max(candidates, key=os.path.getmtime)
                try:
                    with open(self._saved_cost_source_path, "w", encoding="utf-8") as f:
                        json.dump({"path": self.cost_source}, f, ensure_ascii=False, indent=2)
                except OSError:
                    pass

    def _save_data(self):
        """업로드 데이터를 프로젝트 파일에 저장해 재실행/PC 이동 후 복원한다."""
        if not self.headers or not self.rows:
            return
        payload = {
            "version": 1,
            "source_name": os.path.basename(self.source) if self.source else "업로드 데이터",
            "headers": [clean(v) for v in self.headers],
            "rows": [[clean(v) for v in row] for row in self.rows],
        }
        temp_path = self._saved_data_path + ".tmp"
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
        os.replace(temp_path, self._saved_data_path)

    def _restore_saved_data(self):
        """저장된 업로드 데이터가 있으면 앱 시작 시 자동으로 대시보드에 반영한다."""
        if not os.path.exists(self._saved_data_path):
            return
        try:
            with open(self._saved_data_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            headers = payload.get("headers") or []
            rows = payload.get("rows") or []
            if not headers or not rows:
                return
            self.headers = headers
            self.rows = rows
            self.all_rows = list(rows)
            self.engine.set_data(headers, rows)
            self.source = payload.get("source_name", "저장된 업로드 데이터")
            self.file_label.config(text=f"저장 데이터 · {self.source}")
            self.status.config(text=f"{len(rows):,}건 복원됨 · 저장 데이터 자동 불러오기 완료")
            self._update_kpis()
            self._fill_tree(headers, rows[:1000])
            self._populate_filters()
            self._schedule_render()
        except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
            self.status.config(text=f"저장 데이터 복원 실패: {exc}")

    def _build_ui(self):
        """Create the replacement dashboard shell; calculations remain on ClaimDashboard."""
        DashboardView(self).build()

    def _build_legacy_ui(self):
        self.option_add("*Font", (KOREAN_FONT, 10))
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except tk.TclError: pass
        style.configure("TFrame", background="#eaf4ff")
        style.configure("TNotebook", background="#eaf4ff", borderwidth=0)
        style.configure("TNotebook.Tab", background="#e8f1fb", padding=(14, 8), font=(KOREAN_FONT, 10, "bold"))
        style.map("TNotebook.Tab", background=[("selected", "#ffffff")])
        style.configure("TLabel", background="#ffffff")
        style.configure("TButton", background="#ffffff", foreground="#20354b")
        style.configure("TCombobox", fieldbackground="#ffffff", background="#ffffff", foreground="#20354b")
        style.configure("Accent.TButton", background="#1769d4", foreground="white", padding=(14, 8), font=(KOREAN_FONT, 10, "bold"))
        style.configure("Card.TFrame", background="white", relief="solid", borderwidth=1)
        self.configure(background="#eaf4ff")
        body = tk.PanedWindow(self, orient="horizontal", sashwidth=7, sashrelief="raised", bg="#cbd5e1", bd=0, relief="flat")
        body.pack(fill="both", expand=True)
        sidebar = tk.Frame(body, width=215, background="#082b52")
        sidebar.pack_propagate(False)
        body.add(sidebar, minsize=190, width=215, stretch="never")
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
        def load_menu_icon(filename, size=22):
            path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "public", "images", filename)
            try:
                raw = tk.PhotoImage(file=path)
                return raw.subsample(max(1, raw.width() // size), max(1, raw.height() // size))
            except Exception:
                return None

        self.analysis_icon = load_menu_icon("free-icon-graph-2848907.png")
        self.data_icon = load_menu_icon("free-icon-database-16778425.png")
        if self.analysis_icon: self.image_refs.append(self.analysis_icon)
        if self.data_icon: self.image_refs.append(self.data_icon)
        logo = self._load_image("company_logo.png")
        if logo:
            logo_small = logo.subsample(max(1, logo.width() // 150), max(1, logo.height() // 48))
            logo_label = tk.Label(sidebar, image=logo_small, bg="#082b52")
            logo_label.pack(pady=(18, 4))
            self.image_refs.append(logo_small)
        else:
            tk.Label(sidebar, text="▱", fg="#66a9ff", bg="#082b52", font=(KOREAN_FONT, 28, "bold")).pack(pady=(24, 0))
        tk.Label(sidebar, text="클레임 자동 분석", fg="white", bg="#082b52", font=(KOREAN_FONT, 17, "bold")).pack()
        tk.Label(sidebar, text="Claim Analytics", fg="#a9bfd8", bg="#082b52", font=("Segoe UI", 10)).pack(pady=(0, 25))
        menu_style = ttk.Style(self)
        menu_style.configure("Sidebar.Treeview", background="#082b52", fieldbackground="#082b52", foreground="white", borderwidth=0, rowheight=34, font=(KOREAN_FONT, 11))
        menu_style.map("Sidebar.Treeview", background=[("selected", "#1769d4")], foreground=[("selected", "white")])
        menu = ttk.Treeview(sidebar, show="tree", selectmode="browse", style="Sidebar.Treeview", height=5)
        menu.pack(fill="x", padx=8, pady=(0, 10))
        root = menu.insert("", "end", text="클레임 분석", image=self.analysis_icon, open=True, iid="dashboard")
        menu.insert(root, "end", text="1. 현상별 분석", iid="issue")
        menu.insert(root, "end", text="2. 국가별 분석", iid="country")
        menu.insert(root, "end", text="3. 보고서 출력", iid="report")
        menu.selection_set(root)
        menu.bind("<<TreeviewSelect>>", self._on_sidebar_select)
        self.sidebar_menu = menu
        customer_menu = ttk.Treeview(sidebar, show="tree", selectmode="browse", style="Sidebar.Treeview", height=2)
        customer_menu.pack(fill="x", padx=8, pady=(4, 10))
        customer_root = customer_menu.insert("", "end", text="고객사별 DATA 자동", image=self.data_icon, open=True, iid="customer_root")
        customer_menu.insert(customer_root, "end", text="DATA 업로드", iid="customer_upload")
        customer_menu.bind("<<TreeviewSelect>>", self._on_customer_menu_select)
        self.customer_menu = customer_menu
        merge_menu = ttk.Treeview(sidebar, show="tree", selectmode="browse", style="Sidebar.Treeview", height=2)
        merge_menu.pack(fill="x", padx=8, pady=(4, 10))
        merge_root = merge_menu.insert("", "end", text="검수폴더 병합", image=self.data_icon, open=True, iid="merge_root")
        merge_menu.insert(merge_root, "end", text="검수폴더 병합앱", iid="merge_app")
        merge_menu.bind("<<TreeviewSelect>>", self._on_merge_menu_select)
        self.merge_menu = merge_menu
        visual = tk.Frame(sidebar, bg="#082b52")
        visual.pack(side="bottom", fill="x", padx=8, pady=(0, 2))
        car = self._load_image("car.png")
        if car:
            car_small = car.subsample(max(1, car.width() // 140), max(1, car.height() // 70))
            tk.Label(visual, image=car_small, bg="#082b52").pack(); self.image_refs.append(car_small)
        muffler = self._load_image("muffler.png")
        if muffler:
            muffler_small = muffler.subsample(max(1, muffler.width() // 130), max(1, muffler.height() // 65))
            tk.Label(visual, image=muffler_small, bg="#082b52").pack(); self.image_refs.append(muffler_small)
        tk.Label(sidebar, text="\n데이터 정보\n\n업로드 파일\n21년~26년 클레임 DATA\n\n총 로드 수\n7,706건", justify="left", anchor="nw", padx=12, pady=12, fg="#b8c8da", bg="#123e68", font=(KOREAN_FONT, 9)).pack(side="bottom", fill="x", padx=10, pady=8)
        content = tk.Frame(body, background="#eaf4ff")
        body.add(content, minsize=700, stretch="always")
        top = tk.Frame(content, background="#ffffff", padx=16, pady=14, highlightbackground="#d7e6f5", highlightthickness=1)
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
        self.status = ttk.Label(content, text="", padding=(14, 5), background="#eaf4ff", foreground="#58718e")
        self.status.pack(fill="x")
        self.kpi_frame = tk.Frame(content, background="#eaf4ff")
        self.kpi_frame.pack(fill="x", padx=12, pady=(4, 10))
        self.kpi_labels = []
        for icon, title, color in (("▣", "총 로드 수", "#ef4444"), ("⚒", "총 발생량", "#2563eb"), ("⚙", "총 생산량", "#16a34a"), ("▥", "발생율(PPM)", "#7c3aed")):
            card = tk.Frame(self.kpi_frame, background="white", highlightbackground="#d7e6f5", highlightthickness=1, padx=8, pady=6)
            card.pack(side="left", fill="x", expand=True, padx=7)
            card_icon = None
            icon_file = {"총 로드 수": "free-icon-data-10139543.png", "총 발생량": "free-icon-bar-chart-8696653.png", "총 생산량": "free-icon-gear-8680172.png", "발생율(PPM)": "free-icon-percent-3097292.png"}.get(title)
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
        self.notebook.pack(fill="both", expand=True, padx=12, pady=6)
        self.dashboard_tab = ttk.Frame(self.notebook)
        self.detail_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.dashboard_tab, text="분석 대시보드")
        self.notebook.add(self.detail_tab, text="원본 데이터")
        self.dashboard_tab.rowconfigure(1, weight=1); self.dashboard_tab.columnconfigure(0, weight=1)
        self.top_frame = tk.Frame(self.dashboard_tab, background="#ffffff", highlightbackground="#d7e6f5", highlightthickness=1); self.top_frame.grid(row=0, column=0, sticky="ew", padx=2, pady=2)
        self.top_title = tk.Label(self.top_frame, text="클레임 분석", font=(KOREAN_FONT, 17, "bold"), bg="#ffffff", fg="#10243d")
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
        self.fixed_table = tk.Canvas(self.top_frame, background="white", highlightthickness=0, width=105, height=125)
        self.bottom_canvas = tk.Canvas(self.dashboard_tab, background="#eaf4ff", highlightthickness=0)
        self.bottom_canvas.grid(row=1, column=0, sticky="nsew")
        self.canvas = self.bottom_canvas
        self.bottom_canvas.bind("<Configure>", lambda e: self._schedule_render())
        self.bottom_canvas.bind("<Button-1>", self._on_chart_detail_click)
        # 원본 DATA는 열 수가 많으므로 표 전용 프레임에 가로·세로 스크롤을 모두 둡니다.
        detail_frame = ttk.Frame(self.detail_tab)
        detail_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(detail_frame, show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")
        detail_frame.rowconfigure(0, weight=1)
        detail_frame.columnconfigure(0, weight=1)
        vertical_sb = ttk.Scrollbar(detail_frame, orient="vertical", command=self.tree.yview)
        horizontal_sb = ttk.Scrollbar(detail_frame, orient="horizontal", command=self.tree.xview)
        vertical_sb.grid(row=0, column=1, sticky="ns")
        horizontal_sb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=vertical_sb.set, xscrollcommand=horizontal_sb.set)

    def _on_sidebar_select(self, _event=None):
        selected = self.sidebar_menu.selection()
        if not selected:
            return
        actions = {
            "dashboard": lambda: self.notebook.select(self.dashboard_tab),
            "issue": self.show_issue_analysis,
            "country": self.show_country_analysis,
            "report": self.export_report,
        }
        action = actions.get(selected[0])
        if action:
            action()

    def _on_cost_menu_select(self, _event=None):
        selected = self.cost_menu.selection()
        if selected == ("cost_upload",):
            self.open_cost_data()
        elif selected in (("cost_view",), ("cost_root",)):
            self.show_cost_analysis()

    def open_cost_data(self):
        path = filedialog.askopenfilename(
            title="클레임 비용현황 DATA 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xlsm *.xls"), ("모든 파일", "*.*")],
        )
        if path:
            self.cost_source = path
            try:
                with open(self._saved_cost_source_path, "w", encoding="utf-8") as f:
                    json.dump({"path": path}, f, ensure_ascii=False, indent=2)
            except OSError:
                pass
            self.show_cost_analysis()

    def show_cost_analysis(self):
        if not getattr(self, "cost_source", ""):
            messagebox.showinfo("클레임 비용현황", "등록된 비용 DATA가 없습니다. 사이드바의 DATA 업로드를 눌러 한 번 등록해 주세요.")
            return
        win = tk.Toplevel(self); win.title("클레임 비용현황"); win.geometry("1500x900"); win.configure(bg="#EEF6FF")
        top = tk.Frame(win, bg="white", highlightbackground="#D7E6F5", highlightthickness=1); top.pack(fill="x", padx=14, pady=14)
        tk.Label(top, text="클레임 비용현황", bg="white", fg="#102A4C", font=(KOREAN_FONT, 18, "bold")).pack(side="left", padx=14, pady=10)
        ttk.Button(top, text="DATA 업로드", command=lambda: (win.destroy(), self.open_cost_data())).pack(side="right", padx=10)
        ttk.Button(top, text="새로고침", command=lambda: self._render_cost_view(win)).pack(side="right")
        self._render_cost_view(win)

    def _render_cost_view(self, win):
        for child in win.winfo_children()[1:]: child.destroy()
        try:
            # 수식 검증은 원본 수식 파일로 수행하고, 화면에는 Excel이 계산한
            # cached result만 표시해 수식 원문이 보이지 않도록 합니다.
            formula_wb = openpyxl.load_workbook(self.cost_source, read_only=True, data_only=False)
            wb = openpyxl.load_workbook(self.cost_source, read_only=True, data_only=True)
            control = tk.Frame(win, bg="#EEF6FF"); control.pack(fill="x", padx=14, pady=(0, 8))
            formula_count = 0; error_count = 0; total_rows = 0; total_cells = 0
            for ws in formula_wb.worksheets:
                total_rows += ws.max_row or 0; total_cells += (ws.max_row or 0) * (ws.max_column or 0)
                # 수식 검증은 화면 로딩을 막지 않도록 사용 영역의 앞부분만 빠르게 점검합니다.
                for row in ws.iter_rows(max_row=min(ws.max_row or 0, 300), max_col=min(ws.max_column or 0, 80)):
                    for cell in row:
                        if isinstance(cell.value, str) and cell.value.startswith("="): formula_count += 1
                        if isinstance(cell.value, str) and cell.value.startswith("#"): error_count += 1
            tk.Label(control, text=f"전체 DATA 등록 완료 · 시트 {len(wb.sheetnames)}개 · 행 {total_rows:,} · 계산 결과 표시 · 수식 검증 {formula_count:,}개", bg="#EEF6FF", fg="#102A4C", font=(KOREAN_FONT, 10, "bold")).pack(side="left")
            tk.Label(control, text="시트", bg="#EEF6FF", fg="#102A4C", font=(KOREAN_FONT, 10, "bold")).pack(side="right", padx=(12, 4))
            sheet_var = tk.StringVar(value="전체")
            sheet_combo = ttk.Combobox(control, textvariable=sheet_var, values=["전체"] + list(wb.sheetnames), state="readonly", width=24)
            sheet_combo.pack(side="right")
            body = tk.Frame(win, bg="#EEF6FF"); body.pack(fill="both", expand=True, padx=14, pady=8)
            table_frame = tk.Frame(body, bg="white"); table_frame.pack(fill="both", expand=True, pady=(10, 0))
            virtual_table = VirtualCostTable(table_frame); virtual_table.pack(fill="both", expand=True)
            self._cost_widgets = (virtual_table, wb, sheet_var)
            self._update_cost_view()
            sheet_combo.bind("<<ComboboxSelected>>", lambda e: self._update_cost_view())
        except Exception as exc:
            messagebox.showerror("클레임 비용현황 오류", str(exc))

    def _update_cost_view(self):
        tree, wb, sheet_var = self._cost_widgets
        selected_sheet = sheet_var.get()
        worksheets = wb.worksheets if selected_sheet == "전체" else [wb[selected_sheet]]
        # 서식만 남은 빈 열은 제외해 가로 이동 시 불필요한 렌더링을 줄입니다.
        max_columns = 0
        for ws in worksheets:
            for row in ws.iter_rows(max_row=min(ws.max_row or 0, 500)):
                for index, cell in enumerate(row, start=1):
                    if cell.value not in (None, ""):
                        max_columns = max(max_columns, index)
        max_columns = max(5, max_columns)
        # 원본 보고서의 첫 번째 월 헤더 행을 찾아 월 이름을 그대로 사용합니다.
        month_headers = []
        month_positions = []
        for ws in worksheets:
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 8), values_only=True):
                positions = [i for i, v in enumerate(row) if i >= 4 and clean(v) and ("월" in clean(v) or "합계" in clean(v))]
                candidates = [clean(row[i]) for i in positions]
                if sum("월" in v or "합계" in v for v in candidates) >= 3:
                    month_headers = candidates
                    month_positions = positions
                    break
            if month_headers: break
        if not month_headers:
            month_headers = [f"열{i}" for i in range(5, max_columns + 1)]
            month_positions = list(range(4, max_columns))
        headers = ["구분", "항목"] + month_headers
        output_rows = []
        current_group = ""
        display_row = 0
        for ws in worksheets:
            current_group = ""
            for row_no, row in enumerate(ws.iter_rows(values_only=True), start=1):
                raw = [clean(v) for v in row]
                row_months = [clean(row[i]) for i in month_positions if i < len(row) and clean(row[i])]
                # 원본의 월 헤더 행은 이미 파란색 표 머리글로 표시했으므로
                # 데이터 영역에 중복 삽입하지 않습니다.
                matching_months = sum(a == b for a, b in zip(row_months, month_headers))
                if matching_months >= max(3, len(month_headers) // 2):
                    continue
                item = next((v for v in raw[:5] if v), "")
                if len(raw) > 1 and raw[1] and (len(raw) > 2 and raw[2] or "KMC" in raw[1].upper() or "HMC" in raw[1].upper() or "WIA" in raw[1].upper() or "HMB" in raw[1].upper() or "MOBIS" in raw[1].upper()):
                    current_group = raw[1]
                group = current_group
                label = raw[2] if len(raw) > 2 else item
                is_sales = "매출액" in " ".join(raw[:5])
                divisor = 1000000 if is_sales else 1000
                formatted = []
                for source_index in month_positions:
                    value = row[source_index] if source_index < len(row) else None
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        formatted.append(f"{value / divisor:,.0f}")
                    else:
                        formatted.append(clean(value))
                # 제목/구분만 있고 월별 값이 없는 장식 행은 제외합니다.
                if not any(v not in ("", None) for v in formatted) and not label:
                    continue
                values = [group, label] + formatted
                values += [""] * (len(headers) - len(values))
                if any(v not in ("", None) for v in values):
                    output_rows.append(values[:len(headers)])
                    display_row += 1
        tree.set_data(headers, output_rows)

    def _on_customer_menu_select(self, _event=None):
        if self.customer_menu.selection() == ("customer_upload",):
            self.customer_data_aggregate()

    def _on_merge_menu_select(self, _event=None):
        if self.merge_menu.selection() == ("merge_app",):
            self.open_inspection_merge_app()

    def open_inspection_merge_app(self):
        # 최신 검수폴더 병합 프로젝트를 그대로 실행한다. 기존에 이 파일에
        # 일부 기능만 복사해 두던 방식은 최신 app.py의 체크박스 폴더 선택,
        # DATA 업로드/추가, 정규화/필터링 로직이 누락되는 원인이었다.
        latest_app = os.path.join(os.path.dirname(__file__), "inspection_merge_app.py")
        if os.path.exists(latest_app):
            try:
                subprocess.Popen([sys.executable, latest_app], cwd=os.path.dirname(latest_app))
                return
            except Exception as exc:
                messagebox.showerror("검수폴더 병합앱 실행 오류", str(exc))
                return
        messagebox.showerror("검수폴더 병합앱 없음", f"최신 프로젝트를 찾을 수 없습니다.\n{latest_app}")
        return

        # 아래 구현은 최신 프로젝트가 없는 환경에서의 예비 코드로 유지합니다.
        win = tk.Toplevel(self)
        win.title("검수현황 엑셀 파일 합치기")
        win.geometry("780x560")
        win.minsize(680, 480)
        win.configure(bg="#f4f4f4")
        tk.Label(win, text="검수현황 엑셀 파일 합치기", font=(KOREAN_FONT, 18, "bold"), bg="#f4f4f4").pack(anchor="w", padx=20, pady=(18, 2))
        tk.Label(win, text="파일명에서 년월을 읽어 '해당년월' 컬럼을 추가한 뒤 하나의 xlsx로 저장합니다.", fg="#444", bg="#f4f4f4").pack(anchor="w", padx=20)
        controls = tk.Frame(win, bg="#f4f4f4"); controls.pack(fill="x", padx=20, pady=(14, 6))
        tk.Button(controls, text="검수종합 현황 업로드", command=lambda: self._merge_pick_files(listbox)).pack(side="right", padx=3)
        tk.Button(controls, text="검수종합 현황 다운로드", command=lambda: self._merge_run(listbox, win)).pack(side="right", padx=3)
        area = tk.Frame(win, bg="#f4f4f4"); area.pack(fill="both", expand=True, padx=20)
        listbox = tk.Listbox(area, selectmode="extended", font=("Segoe UI", 10), bg="white")
        listbox.pack(side="left", fill="both", expand=True)
        buttons = tk.Frame(area, bg="#f4f4f4"); buttons.pack(side="right", fill="y", padx=(12, 0))
        tk.Button(buttons, text="여러 폴더 선택", command=lambda: self._merge_pick_folders(listbox)).pack(fill="x", pady=3)
        tk.Button(buttons, text="여러 파일 선택", command=lambda: self._merge_pick_files(listbox)).pack(fill="x", pady=3)
        tk.Button(buttons, text="폴더 하나 추가", command=lambda: self._merge_pick_folder(listbox)).pack(fill="x", pady=3)
        tk.Button(buttons, text="선택 삭제", command=lambda: self._merge_delete_selected(listbox)).pack(fill="x", pady=3)
        tk.Button(buttons, text="전체 삭제", command=lambda: listbox.delete(0, "end")).pack(fill="x", pady=3)
        self._merge_status = tk.StringVar(value="월별 검수 폴더를 추가하세요.")
        tk.Label(win, textvariable=self._merge_status, anchor="w", bg="#f4f4f4", fg="#555").pack(fill="x", padx=20, pady=10)

    def _merge_add_paths(self, listbox, paths):
        existing = set(listbox.get(0, "end"))
        for path in paths:
            if path not in existing: listbox.insert("end", path); existing.add(path)
        self._merge_status.set(f"선택 항목 {listbox.size()}개")

    def _merge_pick_folder(self, listbox):
        path = filedialog.askdirectory(title="검수 폴더 선택")
        if path: self._merge_add_paths(listbox, [path])

    def _merge_pick_folders(self, listbox):
        while True:
            path = filedialog.askdirectory(title="검수 폴더 선택 (취소하면 종료)")
            if not path: break
            self._merge_add_paths(listbox, [path])

    def _merge_pick_files(self, listbox):
        paths = filedialog.askopenfilenames(title="검수현황 Excel 파일 선택", filetypes=[("Excel 파일", "*.xlsx *.xlsm")])
        self._merge_add_paths(listbox, paths)

    def _merge_delete_selected(self, listbox):
        for i in reversed(listbox.curselection()): listbox.delete(i)

    def _merge_run(self, listbox, win):
        inputs = list(listbox.get(0, "end"))
        files = []
        for path in inputs:
            if os.path.isdir(path):
                for root, _dirs, names in os.walk(path):
                    files.extend(sorted(os.path.join(root, f) for f in names if f.lower().endswith((".xlsx", ".xlsm"))))
            elif os.path.isfile(path): files.append(path)
        files = list(dict.fromkeys(files))
        if not files: return messagebox.showwarning("검수현황 병합", "폴더 또는 Excel 파일을 먼저 추가하세요.", parent=win)
        save = filedialog.asksaveasfilename(parent=win, title="병합 결과 저장", defaultextension=".xlsx", initialfile="검수종합_현황_병합.xlsx", filetypes=[("Excel 파일", "*.xlsx")])
        if not save: return
        try:
            out = Workbook(); ws = out.active; ws.title = "검수종합 현황"; header_written = False; total = 0
            for path in files:
                src = openpyxl.load_workbook(path, read_only=True, data_only=True)
                for sheet in src.worksheets:
                    rows = list(sheet.iter_rows(values_only=True))
                    if not rows: continue
                    headers = [clean(v) for v in rows[0]]
                    if not header_written:
                        ws.append(headers + ["해당년월"]); header_written = True
                    month = self._merge_month_from_name(os.path.basename(path))
                    for row in rows[1:]:
                        if any(v not in (None, "") for v in row): ws.append(list(row) + [month]); total += 1
                src.close()
            if not header_written: raise ValueError("읽을 수 있는 데이터가 없습니다.")
            ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
            out.save(save); self._merge_status.set(f"완료: {total:,}건 저장")
            messagebox.showinfo("병합 완료", f"{total:,}건을 저장했습니다.\n{save}", parent=win)
            try: os.startfile(save)
            except OSError: pass
        except Exception as exc:
            messagebox.showerror("병합 오류", str(exc), parent=win)

    def _merge_month_from_name(self, name):
        m = re.search(r"((?:20)?\d{2})\s*[-_.년/ ]\s*(\d{1,2})\s*월?", name)
        if not m: return ""
        year = int(m.group(1)); year += 2000 if year < 100 else 0
        return f"{year:04d}-{int(m.group(2)):02d}"

    def export_report(self):
        if not self.rows:
            messagebox.showwarning("보고서 출력", "먼저 데이터를 업로드해 주세요.")
            return
        # 사용자가 별도 저장 위치를 선택하지 않고, 대시보드와 동일한 내용을
        # 임시 Excel 파일로 만들어 바로 연다.
        path = os.path.join(tempfile.gettempdir(), "클레임_분석_대시보드_미리보기.xlsx")
        wb = Workbook(); ws = wb.active; ws.title = "분석 대시보드"
        ws["A1"] = "클레임 자동 분석"; ws["A2"] = "1) 월별 발생 현황 · 발생율(PPM)"
        ws["A1"].font = openpyxl.styles.Font(name="Noto Sans KR", size=16, bold=True, color="10243D")
        ws["A2"].font = openpyxl.styles.Font(name="Noto Sans KR", size=12, bold=True, color="10243D")
        occur = Counter(month_key(clean(r[2])[:6]) for r in self.rows if len(r) > 2 and clean(r[2])[:6])
        assembly_col = 32  # 업로드 DATA의 Excel AG열
        prod = Counter()
        for r in self.rows:
            if len(r) <= 2:
                continue
            occurrence_month = month_key(clean(r[2])[:6])
            assembly_text = clean(r[assembly_col]) if len(r) > assembly_col else ""
            assembly_month = month_key(assembly_text) if re.search(r"20\d{2}", assembly_text) else ""
            prod[assembly_month or occurrence_month] += 1
        visible_months = set(occur) | set(prod)
        labels = sorted((s for s in visible_months if (int(s.split('.')[0]), int(s.split('.')[1])) >= (21, 1)), key=lambda s: (int(s.split('.')[0]), int(s.split('.')[1])))
        official_assembly = self._inspection_assembly_by_month(self.company_var.get())
        # 대시보드와 동일한 가로형 월별 표
        table_header_row = 21
        table_rows = [("생산월", [prod.get(lab, 0) for lab in labels]),
                      ("발생월", [occur.get(lab, 0) for lab in labels]),
                      ("조립수", [round(official_assembly.get(lab, 0)) for lab in labels]),
                      ("PPM", [round(prod.get(lab, 0) / official_assembly.get(lab, 0) * 1_000_000) if official_assembly.get(lab, 0) else 0 for lab in labels])]
        for col, value in enumerate(["구분"] + labels, 1):
            cell = ws.cell(table_header_row, col, value)
            cell.font = openpyxl.styles.Font(name="Malgun Gothic", size=9, bold=True, color="20354B")
            cell.fill = openpyxl.styles.PatternFill("solid", fgColor="E8F1FB")
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        for ri, (name, values) in enumerate(table_rows, table_header_row + 1):
            ws.cell(ri, 1, name)
            for ci, value in enumerate(values, 2): ws.cell(ri, ci, value)
            for ci in range(1, len(labels) + 2):
                cell = ws.cell(ri, ci); cell.fill = openpyxl.styles.PatternFill("solid", fgColor="FFFFFF")
                cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style="thin", color="C8D3DF"), right=openpyxl.styles.Side(style="thin", color="C8D3DF"), top=openpyxl.styles.Side(style="thin", color="C8D3DF"), bottom=openpyxl.styles.Side(style="thin", color="C8D3DF"))
                cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        # 월별 표 자체를 차트 원본으로 사용해 화면의 표와 그래프가 항상 같은 DATA를 사용하게 한다.
        first_data_col = 2; last_data_col = len(labels) + 1
        chart = BarChart(); chart.title = "1) 월별 클레임 발생현황"; chart.y_axis.title = "건수"; chart.x_axis.title = "월"; chart.height = 8; chart.width = 24
        chart.add_data(Reference(ws, min_col=1, max_col=last_data_col, min_row=table_header_row + 1, max_row=table_header_row + 2), from_rows=True, titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=first_data_col, max_col=last_data_col, min_row=table_header_row)); chart.style = 10
        line = LineChart(); line.add_data(Reference(ws, min_col=1, max_col=last_data_col, min_row=table_header_row + 4, max_row=table_header_row + 4), from_rows=True, titles_from_data=True)
        line.set_categories(Reference(ws, min_col=first_data_col, max_col=last_data_col, min_row=table_header_row)); line.y_axis.axId = 200; line.y_axis.title = "PPM"; line.y_axis.crosses = "max"; line.height = 8; line.width = 24
        chart += line
        ws.add_chart(chart, "A3")
        sections = [("2) 현상별 분석", self._top5_other(34)), ("3) 사용기간 분석", self._usage()), ("4) 주행거리 분석", self._mileage())]
        if self.market_var.get() != "D": sections.append(("5) 국가별 분석", self._top5_items(self._country_counter())))
        positions = ["A30", "G30", "M30", "S30"]
        for pos, (title, items) in zip(positions, sections):
            start_col = 110 + positions.index(pos) * 3
            ws.cell(1, start_col, title); ws.cell(2, start_col, "항목"); ws.cell(2, start_col + 1, "건수")
            for ri, (label, value) in enumerate(items, 3):
                ws.cell(ri, start_col, str(label).replace("\n", " ")); ws.cell(ri, start_col + 1, value)
            c = BarChart(); c.title = title; c.add_data(Reference(ws, min_col=start_col + 1, min_row=2, max_row=2 + len(items)), titles_from_data=True); c.set_categories(Reference(ws, min_col=start_col, min_row=3, max_row=2 + len(items))); c.height = 7; c.width = 7.2
            ws.add_chart(c, pos)
        for col in range(110, 126):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 2
        # Excel은 숨김 열의 데이터를 차트에서 제외할 수 있으므로, 보조 데이터 열은 숨기지 않고
        # 인쇄 영역 밖에 유지한다. 열 폭만 줄여 보고서 본문에서는 보이지 않게 한다.
        for col in range(1, max(26, len(labels) + 2)): ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 10
        ws.freeze_panes = "B22"; ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
        ws.print_area = "A1:Z55"
        wb.save(path)
        try:
            os.startfile(path)
        except OSError as exc:
            messagebox.showerror("Excel 열기 오류", f"보고서 미리보기는 생성되었지만 Excel을 열 수 없습니다.\n{exc}")

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

    def customer_data_aggregate(self):
        """기아 시트의 컬럼 순서를 기준으로 모든 고객사 시트를 하나로 합칩니다."""
        source = filedialog.askopenfilename(
            title="고객사별 DATA 파일 선택",
            initialfile="7월 클레임 DATA(AI편집용).xlsx",
            filetypes=[("Excel 파일", "*.xlsx *.xlsm")],
        )
        if not source:
            return
        try:
            wb = openpyxl.load_workbook(source, read_only=True, data_only=False)
            names = wb.sheetnames
            kia_name = next((n for n in names if "kia" in n.lower() or "기아" in n), None)
            if kia_name is None:
                messagebox.showerror("고객사별 DATA 자동", "기아(KIA) 기준 시트를 찾을 수 없습니다.")
                return
            kia_ws = wb[kia_name]
            kia_rows = list(kia_ws.iter_rows(values_only=True))
            if not kia_rows:
                messagebox.showerror("고객사별 DATA 자동", "기아 기준 시트에 데이터가 없습니다.")
                return
            headers = list(kia_rows[0])
            header_keys = [clean(h).replace("\n", " ").strip().lower() for h in headers]
            target_index = {key: idx for idx, key in enumerate(header_keys) if key}
            occurrence_idx = next((i for i, key in enumerate(header_keys) if "발생구분" in key), None)
            corporation_idx = next((i for i, key in enumerate(header_keys) if "법인" in key), None)
            ro_month_idx = next((i for i, key in enumerate(header_keys) if "ro" in key and "년월" in key), None)
            oem_notice_idx = next((i for i, key in enumerate(header_keys) if "oem" in key and "통보" in key), None)
            if oem_notice_idx is None:
                oem_notice_idx = next((i for i, key in enumerate(header_keys) if "통보서번호" in key), None)
            notice_idx = next((i for i, key in enumerate(header_keys) if "통보서" in key and "oem" not in key), None)
            seq_idx = next((i for i, key in enumerate(header_keys) if key == "seq" or key.endswith(" seq")), None)
            campaign_issue_target_indices = [i for i, key in enumerate(header_keys) if "캠페인" in key and "issue" in key.replace(" ", "")]
            kia_notice_month_idx = next((i for i, key in enumerate(header_keys) if "통보서월" in key), None)
            kia_notice_month = next((row[kia_notice_month_idx] for row in kia_rows[1:] if kia_notice_month_idx is not None and kia_notice_month_idx < len(row) and row[kia_notice_month_idx] not in (None, "")), None)
            kia_year_idx = next((i for i, key in enumerate(header_keys) if key in ("년도", "연도")), None)
            kia_quarter_idx = next((i for i, key in enumerate(header_keys) if "분기" in key), None)
            kia_year_value = next((row[kia_year_idx] for row in kia_rows[1:] if kia_year_idx is not None and kia_year_idx < len(row) and row[kia_year_idx] not in (None, "")), None)
            kia_quarter_value = next((row[kia_quarter_idx] for row in kia_rows[1:] if kia_quarter_idx is not None and kia_quarter_idx < len(row) and row[kia_quarter_idx] not in (None, "")), None)
            output = Workbook()
            ws = output.active
            ws.title = "고객사별 통합"
            ws.append(headers)
            total = 0
            sheet_counts = []
            for name in names:
                # 모비스는 일반 모비스 시트가 아니라 사용자가 지정한 모비스OEM 시트만 사용합니다.
                if (name.strip() == "모비스" or name.strip().lower() == "mobis") and not ("oem" in name.lower() or "oem" in name):
                    continue
                rows = list(wb[name].iter_rows(values_only=True))
                if not rows:
                    continue
                source_headers = [clean(h).replace("\n", " ").strip().lower() for h in rows[0]]
                positions = {}
                for idx, key in enumerate(source_headers):
                    if key and key not in positions:
                        positions[key] = idx
                count = 0
                for row in rows[1:]:
                    if not any(v not in (None, "") for v in row):
                        continue
                    mapped = [row[positions[key]] if key in positions and positions[key] < len(row) else None for key in header_keys]
                    # 일부 시트의 끝부분에는 VIN·날짜 등 몇 개 값만 남은 보조 행이 있습니다.
                    # 이런 행은 실제 클레임 레코드가 아니므로 통합 대상에서 제외합니다.
                    if sum(value not in (None, "") for value in mapped) < 5:
                        continue
                    # 발생구분은 고객사 원본의 한글 표현과 관계없이 공통 코드로 통일합니다.
                    if occurrence_idx is not None:
                        occurrence_value = " ".join(clean(value).lower() for value in row if value is not None)
                        if "해외" in occurrence_value or "overseas" in occurrence_value:
                            mapped[occurrence_idx] = "E"
                        elif "국내" in occurrence_value or "domestic" in occurrence_value:
                            mapped[occurrence_idx] = "D"
                    # 모비스 데이터의 통보서 번호는 기아 기준의 OEM통보서번호로 이동합니다.
                    if oem_notice_idx is not None:
                        mobis_notice_idx = next((i for i, key in enumerate(source_headers) if "oem" in key and "통보" in key), None)
                        if mobis_notice_idx is None:
                            mobis_notice_idx = next((i for i, key in enumerate(source_headers) if "통보서번호" in key), None)
                        if mobis_notice_idx is not None and mobis_notice_idx < len(row):
                            mapped[oem_notice_idx] = row[mobis_notice_idx]
                            if notice_idx is not None:
                                mapped[notice_idx] = row[mobis_notice_idx]
                    if "모비스" in name or "mobis" in name.lower() or any("oem" in key and "통보" in key for key in source_headers):
                        mobis_hk_idx = next((i for i, key in enumerate(source_headers) if "oem" in key and "구분" in key), None)
                        hk_target_idx = next((i for i, key in enumerate(header_keys) if key in ("h/k", "hk", "h / k")), None)
                        if mobis_hk_idx is not None and hk_target_idx is not None and mobis_hk_idx < len(row):
                            mapped[hk_target_idx] = row[mobis_hk_idx]
                    # 모비스 법인은 고정값으로 입력합니다.
                    if corporation_idx is not None and ("모비스" in name or "mobis" in name.lower()):
                        mapped[corporation_idx] = "모비스"
                    # RO년월은 원본의 RO년월을 우선 사용하고, 없으면 RO일자에서 계산합니다.
                    if ro_month_idx is not None and ("위아" in name or "모비스" in name or "wia" in name.lower() or "mobis" in name.lower()):
                        source_ro_idx = next((i for i, key in enumerate(source_headers) if "ro" in key and "년월" in key), None)
                        ro_value = row[source_ro_idx] if source_ro_idx is not None and source_ro_idx < len(row) else None
                        if ro_value in (None, ""):
                            date_idx = next((i for i, key in enumerate(source_headers) if "ro" in key and ("일자" in key or "접수" in key or "date" in key)), None)
                            ro_value = row[date_idx] if date_idx is not None and date_idx < len(row) else None
                            if ro_value not in (None, ""):
                                ro_text = clean(ro_value).replace("/", "-")
                                match = re.search(r"(20\d{2})[- ]?(\d{1,2})", ro_text)
                                ro_value = f"{match.group(1)}{int(match.group(2)):02d}" if match else ro_value
                        if ro_value not in (None, ""):
                            mapped[ro_month_idx] = ro_value
                    # 위아 원본의 전용 컬럼을 기아 기준 컬럼에 맞춰 보정합니다.
                    # 시트명이 달라도 원본 전용 컬럼으로 고객사를 식별합니다.
                    is_wia = ("위아" in name or "wia" in name.lower() or
                              any("고객" in key and "통보" in key for key in source_headers) or
                              any("차종" in key and "표준" in key for key in source_headers) or
                              (name != kia_name and any("캠페인" in key for key in source_headers) and any("vin" in key for key in source_headers)))
                    is_mobis = (("모비스" in name or "mobis" in name.lower() or
                                 ("oem" in name.lower() and "기아" not in name and "kia" not in name.lower())) and name != kia_name) or \
                               any("oem통보서번호" in key.replace(" ", "") or "oem구분" in key.replace(" ", "") or ("oem법인" in key.replace(" ", "") and "통보서번호" in " ".join(source_headers)) for key in source_headers) or \
                               (any("차종명" in key.replace(" ", "") for key in source_headers) and any("통보서번호" in key.replace(" ", "") for key in source_headers))
                    if is_mobis:
                        is_wia = False
                    if is_wia:
                        if corporation_idx is not None:
                            mapped[corporation_idx] = "WIA"
                        wia_hk_source_idx = next((i for i, key in enumerate(source_headers) if key in ("h/k", "hk", "oem구분") or ("h/k" in key and "구분" in key)), None)
                        wia_hk_target_idx = next((i for i, key in enumerate(header_keys) if key in ("h/k", "hk", "h / k")), None)
                        if wia_hk_source_idx is not None and wia_hk_target_idx is not None and wia_hk_source_idx < len(row):
                            wia_hk_value = clean(row[wia_hk_source_idx]).upper()
                            if "KIA" in wia_hk_value:
                                mapped[wia_hk_target_idx] = "K"
                            elif "HMC" in wia_hk_value:
                                mapped[wia_hk_target_idx] = "H"
                        campaign_issue_indices = [i for i, key in enumerate(header_keys) if "캠페인" in key and "issue" in key.replace(" ", "")]
                        for campaign_issue_idx in campaign_issue_indices:
                            mapped[campaign_issue_idx] = None
                        vin_source_idx = next((i for i, key in enumerate(source_headers) if key.replace(" ", "") in ("vinno", "vin번호")), None)
                        vin_target_idx = next((i for i, key in enumerate(header_keys) if key in ("vin", "vin번호", "vin no")), None)
                        if vin_source_idx is not None and vin_target_idx is not None and vin_source_idx < len(row):
                            mapped[vin_target_idx] = row[vin_source_idx]
                        model_source_idx = next((i for i, key in enumerate(source_headers) if "차종" in key and "표준" in key), None)
                        model_target_idx = next((i for i, key in enumerate(header_keys) if "차종" in key), None)
                        if model_source_idx is not None and model_target_idx is not None and model_source_idx < len(row):
                            mapped[model_target_idx] = row[model_source_idx]
                        part_source_idx = next((i for i, key in enumerate(source_headers) if key in ("품번", "part no", "partno")), None)
                        name_source_idx = next((i for i, key in enumerate(source_headers) if "품명" in key), None)
                        cause_part_idx = next((i for i, key in enumerate(header_keys) if "원인부품" in key), None)
                        cause_name_idx = next((i for i, key in enumerate(header_keys) if "원인품명" in key), None)
                        if part_source_idx is not None and part_source_idx < len(row):
                            if cause_part_idx is not None:
                                mapped[cause_part_idx] = row[part_source_idx]
                        if name_source_idx is not None and name_source_idx < len(row) and cause_name_idx is not None:
                            mapped[cause_name_idx] = row[name_source_idx]
                        product_class_idx = next((i for i, key in enumerate(header_keys) if "품명구분" in key), None)
                        if product_class_idx is not None and cause_name_idx is not None:
                            cause_name = clean(mapped[cause_name_idx])
                            is_converter = any(word in cause_name.upper() for word in ("CONVERTER", "CATALYTIC"))
                            mapped[product_class_idx] = "WCC" if is_converter else cause_name
                            product_major_idx = next((i for i, key in enumerate(header_keys) if "품명대구분" in key), None)
                            if product_major_idx is not None:
                                mapped[product_major_idx] = "컨버터" if is_converter else "머플러"
                            converter_spec_idx = next((i for i, key in enumerate(header_keys) if "컨버터" in key and "사양" in key), None)
                            if converter_spec_idx is not None:
                                mapped[converter_spec_idx] = "카파(WIA)" if is_converter else None
                            vehicle2_idx = next((i for i, key in enumerate(header_keys) if "차종2" in key.replace(" ", "")), None)
                            source_vehicle_name_idx = next((i for i, key in enumerate(source_headers) if "차종명" in key), None)
                            if vehicle2_idx is not None:
                                mapped[vehicle2_idx] = "카파(WIA)" if is_converter else (row[source_vehicle_name_idx] if source_vehicle_name_idx is not None and source_vehicle_name_idx < len(row) else None)
                        campaign_source_idx = next((i for i, key in enumerate(source_headers) if key in ("캠페인", "캠페인명", "캠페인 명") or "캠페인" in key), None)
                        campaign_target_idx = next((i for i, key in enumerate(header_keys) if "캠페인" in key), None)
                        if campaign_source_idx is not None and campaign_target_idx is not None and campaign_source_idx < len(row):
                            mapped[campaign_target_idx] = row[campaign_source_idx]
                        date_pairs = (("생산일자", ("생신일", "생산일")), ("수리일자", ("수리일",)), ("판매일자", ("판매일",)))
                        if is_wia or is_mobis:
                            payment_idx = next((i for i, key in enumerate(header_keys) if "납입률" in key), None)
                            if payment_idx is not None:
                                mapped[payment_idx] = 100
                        if is_wia:
                            proposed_rate_idx = next((i for i, key in enumerate(source_headers) if "제안분담율" in key), None)
                            for target_word in ("분담률", "적용률"):
                                target_rate_idx = next((i for i, key in enumerate(header_keys) if target_word in key), None)
                                if proposed_rate_idx is not None and target_rate_idx is not None and proposed_rate_idx < len(row):
                                    mapped[target_rate_idx] = row[proposed_rate_idx]
                            cost_pairs = (("부품비", "변제부품비"), ("공임비", "변제공임비"), ("외주비", "변제외주비"))
                            for target_word, source_word in cost_pairs:
                                target_cost_idx = next((i for i, key in enumerate(header_keys) if target_word in key), None)
                                source_cost_idx = next((i for i, key in enumerate(source_headers) if source_word in key), None)
                                if target_cost_idx is not None and source_cost_idx is not None and source_cost_idx < len(row):
                                    mapped[target_cost_idx] = row[source_cost_idx]
                        for target_word, source_words in date_pairs:
                            target_idx = next((i for i, key in enumerate(header_keys) if target_word in key), None)
                            source_idx = next((i for i, key in enumerate(source_headers) if any(word in key for word in source_words)), None)
                            if target_idx is not None and source_idx is not None and source_idx < len(row) and row[source_idx] not in (None, ""):
                                date_text = clean(row[source_idx]).replace("-", "/").replace(".", "/")
                                date_match = re.search(r"(20\d{2})/(\d{1,2})/(\d{1,2})", date_text)
                                compact_match = re.fullmatch(r"(20\d{2})(\d{2})(\d{2})", date_text.strip())
                                if date_match:
                                    mapped[target_idx] = f"{date_match.group(1)}/{int(date_match.group(2)):02d}/{int(date_match.group(3)):02d}"
                                elif compact_match:
                                    mapped[target_idx] = f"{compact_match.group(1)}/{compact_match.group(2)}/{compact_match.group(3)}"
                                else:
                                    mapped[target_idx] = row[source_idx]
                        mileage_source_idx = next((i for i, key in enumerate(source_headers) if "주행거리" in key and ("km" in key or "㎞" in key)), None)
                        mileage_target_idx = next((i for i, key in enumerate(header_keys) if "주행거리" in key), None)
                        if mileage_source_idx is not None and mileage_target_idx is not None and mileage_source_idx < len(row):
                            mileage_text = clean(row[mileage_source_idx]).replace(",", "")
                            try:
                                mapped[mileage_target_idx] = float(mileage_text) if "." in mileage_text else int(mileage_text)
                            except (TypeError, ValueError):
                                mapped[mileage_target_idx] = row[mileage_source_idx]
                        company_idx = next((i for i, key in enumerate(header_keys) if "업체" in key), None)
                        if company_idx is not None and mapped[company_idx] in (None, ""):
                            mapped[company_idx] = "R151"
                        customer_notice_idx = next((i for i, key in enumerate(source_headers) if "고객" in key and "통보" in key), None)
                        notice_target_idx = next((i for i, key in enumerate(header_keys) if "통보서" in key and "oem" not in key), None)
                        if customer_notice_idx is not None and notice_target_idx is not None and customer_notice_idx < len(row):
                            mapped[notice_target_idx] = row[customer_notice_idx]
                        claim_ro_idx = next((i for i, key in enumerate(source_headers) if "ro" in key and "클레임" in key), None)
                        ro_target_idx = next((i for i, key in enumerate(header_keys) if ("r/o" in key or key.startswith("ro")) and "번호" in key), None)
                        if claim_ro_idx is not None and ro_target_idx is not None and claim_ro_idx < len(row):
                            mapped[ro_target_idx] = row[claim_ro_idx]
                        ctype_idx = next((i for i, key in enumerate(header_keys) if "c/type" in key or "ctype" in key), None)
                        if ctype_idx is not None and clean(mapped[ctype_idx]).replace(" ", "") in ("W:일반클레임", "W:일반클레임", "W:일반클레임"):
                            mapped[ctype_idx] = "W"
                    if is_mobis:
                        mobis_company_idx = next((i for i, key in enumerate(header_keys) if "업체" in key), None)
                        if mobis_company_idx is not None:
                            mapped[mobis_company_idx] = "R151"
                        mobis_field_pairs = (
                            (("r/o번호", "ro번호", "ro 번호"), ("r/o번호", "ro번호")),
                            (("vin", "vin번호"), ("vin", "vin번호")),
                            (("차종명",), ("차종명", "차종(표준)", "차종")),
                            (("c/type", "클레임타입"), ("c/type", "ctype")),
                            (("원인부품",), ("품번",)),
                            (("원인품명",), ("품명",)),
                            (("원인",), ("원인코드",)),
                            (("현상",), ("현상코드",)),
                            (("납입률",), ("업체책임율",)),
                            (("분담률",), ("업체분담율",)),
                            (("적용률",), ("업체적용율",)),
                            (("부품비",), ("업체부품비",)),
                            (("공임비",), ("업체공임비",)),
                            (("외주비",), ("업체외주비",)),
                            (("변제합계",), ("금액",)),
                        )
                        for target_names, source_names in mobis_field_pairs:
                            normalized_targets = {value.replace(" ", "").lower() for value in target_names}
                            normalized_sources = {value.replace(" ", "").lower() for value in source_names}
                            target_idx = next((i for i, key in enumerate(header_keys) if key.replace(" ", "").lower() in normalized_targets), None)
                            if target_names == ("차종명",):
                                source_idx = next((i for i, key in enumerate(source_headers) if key.replace(" ", "").lower() == "차종명"), None)
                                if source_idx is None:
                                    source_idx = next((i for i, key in enumerate(source_headers) if key.replace(" ", "").lower() in normalized_sources), None)
                            else:
                                source_idx = next((i for i, key in enumerate(source_headers) if key.replace(" ", "").lower() in normalized_sources), None)
                            if target_idx is not None and source_idx is not None and source_idx < len(row):
                                mapped[target_idx] = row[source_idx]
                        # 모비스OEM의 원인코드·현상코드는 설명문(원인/현상)으로 변환하지 않고
                        # 원본 DATA의 코드값을 그대로 사용합니다. 이 대입을 마지막에 수행해
                        # 앞선 일반 매핑이나 코드-설명 매핑이 값을 덮어쓰지 못하게 합니다.
                        for code_name in ("원인코드", "현상코드"):
                            source_code_idx = next(
                                (i for i, key in enumerate(source_headers)
                                 if key.replace(" ", "").lower() == code_name),
                                None,
                            )
                            target_code_idx = next(
                                (i for i, key in enumerate(header_keys)
                                 if key.replace(" ", "").lower() == code_name),
                                None,
                            )
                            if (source_code_idx is not None and target_code_idx is not None
                                    and source_code_idx < len(row)):
                                mapped[target_code_idx] = row[source_code_idx]
                        # 모비스OEM 원본의 실제 기준 열은 AU=원인코드, AV=현상코드이며,
                        # 통합 결과의 좌측 AH·AI(원인코드·현상코드)로 옮깁니다.
                        # 원본 열의 설명문/헤더명과 무관하게 사용자가 지정한 열 위치를
                        # 최우선으로 적용해 다른 매핑이 값을 덮어쓰지 않게 합니다.
                        mobis_code_columns = ((33, 46), (34, 47))  # AH<-AU, AI<-AV
                        for target_code_idx, source_code_idx in mobis_code_columns:
                            if (target_code_idx < len(mapped) and source_code_idx < len(row)
                                    and row[source_code_idx] not in (None, "")):
                                mapped[target_code_idx] = row[source_code_idx]
                        mobis_ctype_target_idx = next((i for i, key in enumerate(header_keys) if key.replace(" ", "").lower() == "c/type"), None)
                        mobis_ctype_source_idx = next((i for i, key in enumerate(source_headers) if key.replace(" ", "").lower() == "클레임타입"), None)
                        if mobis_ctype_target_idx is not None and mobis_ctype_source_idx is not None and mobis_ctype_source_idx < len(row):
                            mapped[mobis_ctype_target_idx] = row[mobis_ctype_source_idx]
                        mobis_cause_name_idx = next((i for i, key in enumerate(header_keys) if "원인품명" in key), None)
                        if mobis_cause_name_idx is not None:
                            mobis_cause_name = clean(mapped[mobis_cause_name_idx])
                            mobis_is_converter = any(word in mobis_cause_name.upper() for word in ("CONVERTER", "CATALYTIC"))
                            mobis_major_idx = next((i for i, key in enumerate(header_keys) if "품명대구분" in key), None)
                            if mobis_major_idx is not None:
                                mapped[mobis_major_idx] = "컨버터" if mobis_is_converter else "머플러"
                            mobis_spec_idx = next((i for i, key in enumerate(header_keys) if "컨버터" in key and "사양" in key), None)
                            if mobis_spec_idx is not None and not mobis_is_converter:
                                mapped[mobis_spec_idx] = None
                            mobis_vehicle2_idx = next((i for i, key in enumerate(header_keys) if "차종2" in key.replace(" ", "")), None)
                            mobis_vehicle_name_idx = next((i for i, key in enumerate(header_keys) if key.replace(" ", "") == "차종명"), None)
                            if mobis_vehicle2_idx is not None:
                                mapped[mobis_vehicle2_idx] = "카파(WIA)" if mobis_is_converter else (mapped[mobis_vehicle_name_idx] if mobis_vehicle_name_idx is not None else None)
                        work_code_target_idx = next((i for i, key in enumerate(header_keys) if "주작업코드" in key), None)
                        work_code_source_idx = next((i for i, key in enumerate(source_headers) if key.replace(" ", "") == "작업코드"), None)
                        if work_code_target_idx is not None and work_code_source_idx is not None and work_code_source_idx < len(row):
                            mapped[work_code_target_idx] = row[work_code_source_idx]
                        mobis_oem_notice_idx = next((i for i, key in enumerate(source_headers) if key.replace(" ", "") == "oem통보서번호"), None)
                        if mobis_oem_notice_idx is None:
                            mobis_oem_notice_idx = next((i for i, key in enumerate(source_headers) if "통보서번호" in key or ("oem" in key and "통보서" in key)), None)
                        if mobis_oem_notice_idx is not None and mobis_oem_notice_idx < len(row):
                            mobis_oem_notice = row[mobis_oem_notice_idx]
                            if notice_idx is not None:
                                mapped[notice_idx] = mobis_oem_notice
                            if oem_notice_idx is not None:
                                mapped[oem_notice_idx] = mobis_oem_notice
                            notice_text = clean(mobis_oem_notice)
                            notice_month_match = re.match(r"(20\d{2})(\d{2})", notice_text)
                            if ro_month_idx is not None and notice_month_match:
                                mapped[ro_month_idx] = int(f"{notice_month_match.group(1)}{notice_month_match.group(2)}")
                        mobis_seq_source_idx = next((i for i, key in enumerate(source_headers) if key in ("순위", "no", "no.")), None)
                        if mobis_seq_source_idx is not None and seq_idx is not None and mobis_seq_source_idx < len(row):
                            mapped[seq_idx] = row[mobis_seq_source_idx]
                        # 모비스도 위아와 동일하게 기아 기준 날짜 형식으로 변환합니다.
                        for target_word, source_words in (("생산일자", ("생산일", "생신일")), ("수리일자", ("수리일",)), ("판매일자", ("판매일",))):
                            target_idx = next((i for i, key in enumerate(header_keys) if target_word in key), None)
                            source_idx = next((i for i, key in enumerate(source_headers) if any(word in key for word in source_words)), None)
                            if target_idx is not None and source_idx is not None and source_idx < len(row) and row[source_idx] not in (None, ""):
                                date_text = clean(row[source_idx]).replace("-", "/").replace(".", "/").strip()
                                date_match = re.search(r"(20\d{2})/(\d{1,2})/(\d{1,2})", date_text)
                                compact_match = re.fullmatch(r"(20\d{2})(\d{2})(\d{2})", date_text)
                                if date_match:
                                    mapped[target_idx] = f"{date_match.group(1)}/{int(date_match.group(2)):02d}/{int(date_match.group(3)):02d}"
                                elif compact_match:
                                    mapped[target_idx] = f"{compact_match.group(1)}/{compact_match.group(2)}/{compact_match.group(3)}"
                    # 위아 캠페인Issue는 최종 저장 직전에 다시 확인해 항상 공란으로 유지합니다.
                    if is_wia or (name != kia_name and any("캠페인" in key for key in source_headers) and any("vin" in key for key in source_headers)):
                        for campaign_issue_idx in campaign_issue_target_indices:
                            mapped[campaign_issue_idx] = None
                    if is_wia or is_mobis:
                        # 기아 시트와 동일한 통보서월·년도·분기 파생값을 생성합니다.
                        notice_month_idx = next((i for i, key in enumerate(header_keys) if "통보서월" in key), None)
                        year_idx = next((i for i, key in enumerate(header_keys) if key in ("년도", "연도")), None)
                        quarter_idx = next((i for i, key in enumerate(header_keys) if "분기" in key), None)
                        # 위아·모비스는 각 원본의 월이 아니라 기아 기준 통보서월을 공통 사용합니다.
                        month_value = mapped[ro_month_idx] if is_mobis and ro_month_idx is not None else kia_notice_month
                        if month_value in (None, ""):
                            month_value = mapped[ro_month_idx] if ro_month_idx is not None else None
                        month_match = re.search(r"(20\d{2})[^0-9]?(\d{1,2})", clean(month_value)) if month_value not in (None, "") else None
                        if month_match:
                            year_value = int(month_match.group(1))
                            month_number = int(month_match.group(2))
                            if notice_month_idx is not None:
                                mapped[notice_month_idx] = int(f"{year_value}{month_number:02d}")
                            if year_idx is not None:
                                mapped[year_idx] = kia_year_value if kia_year_value not in (None, "") else f"{year_value}년"
                            if quarter_idx is not None:
                                mapped[quarter_idx] = kia_quarter_value if kia_quarter_value not in (None, "") else f"{(month_number - 1) // 3 + 1}분기"
                    if is_mobis:
                        exact_vehicle_source_idx = next((i for i, key in enumerate(source_headers) if key.replace(" ", "").lower() == "차종명"), None)
                        exact_vehicle_target_idx = next((i for i, key in enumerate(header_keys) if key.replace(" ", "").lower() == "차종명"), None)
                        # 모비스OEM의 AS열(Excel 45열)을 차종명 기준값으로 사용합니다.
                        if exact_vehicle_target_idx is not None and len(row) > 44:
                            mapped[exact_vehicle_target_idx] = row[44]
                        elif exact_vehicle_source_idx is not None and exact_vehicle_target_idx is not None and exact_vehicle_source_idx < len(row):
                            mapped[exact_vehicle_target_idx] = row[exact_vehicle_source_idx]
                    # SEQ와 RO년월은 기아 기준과 같은 숫자형 표기로 통일합니다.
                    if seq_idx is not None and mapped[seq_idx] not in (None, ""):
                        seq_text = clean(mapped[seq_idx]).replace(",", "")
                        if re.fullmatch(r"\d+", seq_text):
                            mapped[seq_idx] = int(seq_text)
                    if ro_month_idx is not None and mapped[ro_month_idx] not in (None, ""):
                        ro_text = clean(mapped[ro_month_idx]).replace("/", "-")
                        ro_match = re.search(r"(20\d{2})[- ]?(\d{1,2})", ro_text)
                        if ro_match:
                            mapped[ro_month_idx] = f"{ro_match.group(1)}/{int(ro_match.group(2)):02d}"
                    ws.append(mapped)
                    count += 1
                    total += 1
                sheet_counts.append(f"{name}: {count:,}건")
            for cell in ws[1]:
                cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
                cell.fill = openpyxl.styles.PatternFill("solid", fgColor="1F4E78")
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                letter = col[0].column_letter
                ws.column_dimensions[letter].width = min(max(max(len(clean(c.value)) for c in col[: min(len(col), 100)]) + 2, 10), 32)
            # 사용자가 저장 위치를 선택하지 않아도 통합 결과를 Excel에서 바로 볼 수 있도록
            # 임시 파일로 열고, 동시에 대시보드에도 표시합니다.
            # 고정 파일명은 이전에 열린 Excel 파일과 충돌해 Permission denied가 발생할 수 있으므로
            # 매 실행마다 잠금되지 않은 고유 임시 파일을 만든다.
            preview_fd, preview_path = tempfile.mkstemp(prefix="고객사별_통합_DATA_", suffix=".xlsx")
            os.close(preview_fd)
            output.save(preview_path)
            os.startfile(preview_path)
            # 고객사별 자동 통합은 클레임 분석 화면과 분리합니다.
            # 통합 결과는 Excel에서만 열고, 클레임 분석 그래프에는 반영하지 않습니다.
            messagebox.showinfo("고객사별 DATA 자동 완료", f"기준 시트: {kia_name}\n통합 건수: {total:,}건\n\n{chr(10).join(sheet_counts)}\n\n통합 결과를 Excel에서 열었습니다.\n클레임 분석 그래프에는 반영하지 않습니다.")
        except Exception as exc:
            messagebox.showerror("고객사별 DATA 자동 오류", str(exc))

    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("모든 Excel 파일", "*.xlsx *.xlsm *.xltx *.xltm *.xls *.xlt *.xlsb"), ("모든 파일", "*.*")])
        if path:
            try:
                self.load(path)
                self.render()
                self.status.config(text=self.status.cget("text") + " · 업로드 데이터 반영 완료 · 분석 새로고침으로 재계산 가능")
            except Exception as e:
                messagebox.showerror("파일 읽기 오류", f"엑셀을 읽지 못했습니다.\n{e}")

    def download_uploaded_data(self):
        """Save the complete uploaded source DATA, independent of active filters."""
        if not getattr(self, "all_rows", None) or not getattr(self, "headers", None):
            messagebox.showwarning("DATA 다운로드", "먼저 DATA를 업로드해 주세요.")
            return
        try:
            wb = Workbook(); ws = wb.active; ws.title = "업로드 DATA"
            ws.append(list(self.headers))
            for row in self.all_rows:
                ws.append(list(row))
            ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
            preview_fd, preview_path = tempfile.mkstemp(prefix="업로드_DATA_", suffix=".xlsx")
            os.close(preview_fd)
            wb.save(preview_path)
            os.startfile(preview_path)
        except Exception as exc:
            messagebox.showerror("DATA 다운로드 오류", str(exc))

    def refresh_analysis(self):
        """마지막 업로드 파일을 다시 읽고 모든 그래프를 재계산한다."""
        # 업로드 직후 원본 파일을 다시 읽으면 Excel 날짜/빈 셀 해석 차이로
        # 77건이 76건으로 되돌아갈 수 있으므로, 현재 보관 중인 DATA를 유지한다.
        if not self.all_rows:
            self.render()
            return
        try:
            self.rows = list(self.all_rows)
            self._fill_tree(self.headers, self.rows[:1000])
            self._populate_filters()
            self.render()
            self.status.config(text=self.status.cget("text") + " · 분석 새로고침 완료")
        except Exception as e:
            messagebox.showerror("분석 새로고침 오류", f"업로드 파일을 다시 읽지 못했습니다.\n{e}")

    def load(self, path):
        suffix = os.path.splitext(path)[1].lower()
        if suffix in (".xls", ".xlt"):
            # 구형 Excel 97-2003 형식은 openpyxl이 읽지 못하므로 xlrd를 통해 읽는다.
            sheets = pd.read_excel(path, sheet_name=None, header=None, engine="xlrd")
            frame = max(sheets.values(), key=lambda df: df.shape[0] * df.shape[1])
            values = frame.where(pd.notna(frame), None).values.tolist()
            sheet_name = next(name for name, df in sheets.items() if df is frame)
            headers, rows = values[0], values[1:]
        elif suffix == ".xlsb":
            try:
                sheets = pd.read_excel(path, sheet_name=None, header=None, engine="pyxlsb")
            except ImportError as exc:
                raise RuntimeError(".xlsb 파일을 읽으려면 pyxlsb 패키지가 필요합니다.\\n설치: pip install pyxlsb") from exc
            frame = max(sheets.values(), key=lambda df: df.shape[0] * df.shape[1])
            values = frame.where(pd.notna(frame), None).values.tolist()
            sheet_name = next(name for name, df in sheets.items() if df is frame)
            headers, rows = values[0], values[1:]
        else:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            # 활성 시트가 빈 시트인 업로드 파일도 처리하도록 데이터가 가장 많은 시트 선택
            ws = max(wb.worksheets, key=lambda sh: (sh.max_row or 0) * (sh.max_column or 0))
            it = ws.iter_rows(values_only=True)
            headers = list(next(it))
            rows = [list(r) for r in it if any(v is not None for v in r)]
            sheet_name = ws.title
        self.headers = headers
        self.rows = [list(r) for r in rows if any(v is not None for v in r)]
        self.all_rows = list(self.rows)
        self.engine.set_data(self.headers, self.all_rows)
        self.source = path
        self._save_data()
        self.file_label.config(text=os.path.basename(path))
        self.status.config(text=f"{len(self.rows):,}건 로드됨 · 열 수 {len(headers)} · 시트: {sheet_name} · 형식: {suffix}")
        self._update_kpis()
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
        self.apply_filters()

    def _update_kpis(self):
        """Update KPI cards from the complete upload and current filtered rows."""
        if not getattr(self, "kpi_labels", None):
            return
        total_claims = len(getattr(self, "all_rows", []))
        filtered_rows = getattr(self, "rows", [])
        occurrence_count = sum(1 for row in filtered_rows if len(row) > 2 and clean(row[2]))
        # 총 생산건수는 업로드 행 수가 아니라 현재 선택 조건에 해당하는 검수 조립수 합계다.
        selected_company = getattr(self, "company_var", tk.StringVar(value="전체")).get()
        assembly_by_month = self._inspection_assembly_by_month(selected_company)
        production_count = int(round(sum(assembly_by_month.values())))
        production_month_count = sum(1 for row in filtered_rows if len(row) > 32 and clean(row[32]))
        ppm = production_month_count / production_count * 1_000_000 if production_count else 0
        self.kpi_labels[0].config(text=f"{total_claims:,}건")
        self.kpi_labels[1].config(text=f"{occurrence_count:,}건")
        self.kpi_labels[2].config(text=f"{production_count:,}건")
        self.kpi_labels[3].config(text=f"{ppm:,.0f}")

    def apply_filters(self):
        self._update_part_options()
        company = self.company_var.get(); model = self.model_var.get(); market = self.market_var.get(); parts = getattr(self, "selected_parts", set()); names = getattr(self, "selected_names", set())
        available_parts = {clean(r[11]) for r in self.all_rows
                           if len(r) > 11 and clean(r[11])
                           and (company == "전체" or (len(r) > 0 and clean(r[0]) == company))
                           and (model == "전체" or (len(r) > 45 and clean(r[45]) == model))
                           and (market == "전체" or (len(r) > 1 and clean(r[1]) == market))
                           and (not names or (len(r) > 41 and clean(r[41]) in names))}
        self.rows = self.engine.filter_rows(company, model, market, parts, names)
        self._update_kpis()
        name_text = "전체" if not names else ", ".join(sorted(names)[:3]) + (" 외" if len(names) > 3 else "")
        all_parts_selected = bool(available_parts) and parts == available_parts
        part_text = "전체" if not parts or all_parts_selected else ", ".join(sorted(parts)[:3]) + (" 외" if len(parts) > 3 else "")
        if all_parts_selected:
            self.part_var.set("전체")
        self.status.config(text=f"{len(self.rows):,}건 분석 중 · 고객사: {company} · 차종: {model} · 품명: {name_text} · 품번: {part_text} · 구분: {market}")
        self._fill_tree([], []) if False else None
        self._schedule_render()

    def _fill_tree(self, headers, rows):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = [f"c{i}" for i in range(len(headers))]
        for i, h in enumerate(headers):
            name = clean(h) or f"열{i+1}"
            self.tree.heading(f"c{i}", text=name)
            # 헤더와 실제 값 중 긴 쪽에 맞춰 표시 폭을 확보하되, 긴 텍스트가
            # 화면 전체를 독점하지 않도록 제한합니다. 전체 열은 가로 스크롤로 확인합니다.
            sample_lengths = [len(clean(row[i])) for row in rows[:100] if len(row) > i and clean(row[i])]
            content_width = max([len(name)] + sample_lengths + [8])
            column_width = min(240, max(82, content_width * 9 + 18))
            self.tree.column(f"c{i}", width=column_width, minwidth=60, anchor="center", stretch=False)
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
        self.fixed_table.place_configure(x=15, y=self.top_canvas.winfo_y() + labels_y, width=105, height=121)
        self.tk.call("raise", str(self.fixed_table))
        fixed_fills = ["#e8f1fb", "#ffffff", "#ffffff", "#ffffff", "#ffffff"]
        for i, name in enumerate(("월", "생산월", "발생월", "조립수", "PPM")):
            yy = labels_y + i*24
            self.fixed_table.create_rectangle(0, yy-labels_y, 105, yy-labels_y+24, fill=fixed_fills[i], outline="#c8d3df")
            self.fixed_table.create_text(52, yy-labels_y+12, text=name, anchor="center", font=(KOREAN_FONT, 9), fill="#20354b")
        # 고정 라벨 표도 본문과 같은 24px 행 간격을 사용하고, 마지막 PPM 하단선을 별도로 보장한다.
        for line_y in (0, 24, 48, 72, 96, 120):
            self.fixed_table.create_line(0, line_y, 104, line_y, fill="#c8d3df", width=1)
        self.fixed_table.create_line(0, 0, 0, 120, fill="#c8d3df", width=1)
        self.fixed_table.create_line(104, 0, 104, 120, fill="#c8d3df", width=1)
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

    def _schedule_render(self):
        """Coalesce resize events so moving/resizing the window does one redraw."""
        if self._render_job is not None:
            try:
                self.after_cancel(self._render_job)
            except tk.TclError:
                pass
        self._render_job = self.after(80, self._render_scheduled)

    def _render_scheduled(self):
        self._render_job = None
        self.render()

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
        self._rounded_panel(x, y, x+w, y+h, radius=12, fill="#ffffff")
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
        if title.startswith("2)") or title.startswith("5)"):
            detail = self.canvas.create_text(x+w-12, y+14, text="자세히 보기 〉", anchor="ne", font=(KOREAN_FONT, 9, "bold"), fill="#1769d4", tags=("chart_detail",))
            self.canvas.tag_bind(detail, "<Button-1>", lambda _event, chart_title=title: self._open_chart_detail(chart_title))
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

    def _open_chart_detail(self, title):
        if title.startswith("2)"):
            self.show_issue_analysis()
        elif title.startswith("5)"):
            self.show_country_analysis()

    def _on_chart_detail_click(self, event):
        # Text tag bindings handle the actual button; this handler keeps the canvas focus behavior stable.
        return None


    def _gradient_bar(self, x1, y1, x2, y2, top_color, bottom_color):
        # 단색 막대: 그라데이션 효과 취소
        self.canvas.create_rectangle(x1, y1, x2, y2, fill=top_color, outline="")

    def _rounded_panel(self, x1, y1, x2, y2, radius=12, fill="#ffffff"):
        """Draw a filled rounded panel without a border on a Tk canvas."""
        r = min(radius, (x2 - x1) / 2, (y2 - y1) / 2)
        self.canvas.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline="")
        self.canvas.create_rectangle(x1, y1 + r, x2, y2 - r, fill=fill, outline="")
        for cx, cy, start in ((x1 + r, y1 + r, 90), (x2 - r, y1 + r, 0), (x2 - r, y2 - r, 270), (x1 + r, y2 - r, 180)):
            self.canvas.create_arc(cx - r, cy - r, cx + r, cy + r, start=start, extent=90, fill=fill, outline="")

    def _monthly_combo(self, x, y, w, h):
        # 발생월: 통보서 앞자리 6개(index 2)
        occur = Counter(month_key(clean(r[2])[:6]) for r in self.rows if len(r)>2 and clean(r[2])[:6])
        # 생산월: 업로드 DATA의 Excel AG열(0-based index 32)을 사용한다.
        # 21년 이전 생산 DATA도 원본에는 유지하되 월별 표에서는 숨긴다.
        assembly_col = 32
        # 생산월 누락/범위 밖 행은 해당 행의 발생월에 배정해 모든 클레임이 생산월에 포함되도록 함
        prod = Counter()
        for r in self.rows:
            if len(r) <= 2: continue
            occurrence_month = month_key(clean(r[2])[:6])
            assembly_text = clean(r[assembly_col]) if len(r) > assembly_col else ""
            assembly_month = month_key(assembly_text) if re.search(r"20\d{2}", assembly_text) else ""
            # AG열에 생산월이 있으면 해당 생산월 그대로 집계한다.
            # 생산월이 비어 있을 때만 발생월을 보완값으로 사용한다.
            prod[assembly_month or occurrence_month] += 1
        def sort_month(v):
            try:
                yy, mm = v.split('.'); return int(yy), int(mm)
            except Exception: return (999, 999)
        # 월별 축은 통보서에서 추출한 발생월만 사용해 생산차량의 과거 연도가 섞이지 않게 함
        # 조립수만 있거나 클레임만 있는 월도 표/그래프에서 유지한다.
        visible_months = set(occur) | set(prod)
        labels = sorted(
            (label for label in visible_months if sort_month(label) >= (21, 1)),
            key=sort_month,
        )
        # 공식 검수 DATA를 읽은 뒤 조립수 전용 월까지 월 축에 추가한다.
        # 생산월 데이터가 없는 경우에도 비교가 가능하도록 더미 생산수 생성
        occ_vals = [occur.get(k, 0) for k in labels]
        prod_vals = [prod.get(k, 0) for k in labels]
        if not any(prod_vals): prod_vals = [v * 80 for v in occ_vals]
        # 더미 생산수 환경의 표시 PPM을 200~500 수준으로 보정
        # 모든 콤보박스가 전체일 때 검수폴더 병합 DATA의 수량누계를 사용한다.
        wia_kappa_rule = (
            self.company_var.get() == "WIA"
            and "카파" in self.model_var.get()
            and self.name_var.get() == "전체"
        )
        wia_ka4_frt_rule = (
            self.company_var.get() == "WIA"
            and "KA4" in self.model_var.get().upper()
            and "MUFFLER ASSY-FR" in self.name_var.get().upper()
        )
        assembly_filters_supported = (
            (self.model_var.get() == "전체" or wia_kappa_rule or wia_ka4_frt_rule)
            and (self.name_var.get() == "전체" or wia_ka4_frt_rule)
            and self.part_var.get() == "전체"
        )
        selected_company = self.company_var.get()
        official_assembly = self._inspection_assembly_by_month(selected_company) if assembly_filters_supported else {}
        visible_months |= set(official_assembly)
        labels = sorted((label for label in visible_months if sort_month(label) >= (21, 1)), key=sort_month)
        if not labels:
            return
        occ_vals = [occur.get(k, 0) for k in labels]
        prod_vals = [prod.get(k, 0) for k in labels]
        # 공식 검수 DATA가 없으면 임의의 더미값을 만들지 않고 0으로 표시한다.
        assembly_vals = [int(round(official_assembly.get(label, 0))) for label in labels]
        # PPM은 생산월 건수 / 검수 조립수 기준으로 대시보드·보고서·KPI를 통일한다.
        rates = [p / a * 1_000_000 if a else 0 for p, a in zip(prod_vals, assembly_vals)]
        bottom=y+h-38; chart_h=h-65; n=len(labels)
        label_w = 105
        cell_w = max(52, (w-65)/max(1,n))
        group = cell_w; left = x + label_w; bw = max(3, group*.32)
        maxv=max(max(occ_vals or [1]), max(prod_vals or [1])); maxr=max(max(rates or [0]), 1)
        for i,(lab,o,p,rate) in enumerate(zip(labels,occ_vals,prod_vals,rates)):
            bx=left+i*group+group*.12; oh=chart_h*o/maxv; ph=chart_h*p/maxv
            # 생산월은 왼쪽(중간 청회색), 발생월은 오른쪽(파랑)으로 배치
            production_color = "#94a3b8"
            occurrence_color = "#2563eb"
            rate_color = "#ef4444"
            self._gradient_bar(bx, bottom-ph, bx+bw, bottom, production_color, production_color)
            self._gradient_bar(bx+bw+2, bottom-oh, bx+bw*2+2, bottom, occurrence_color, occurrence_color)
            if n <= 30:
                self.canvas.create_text(bx+bw*1.5+2, bottom-oh-2, text=f"{o:,}", anchor="s", font=(KOREAN_FONT, 7))
            self.canvas.create_text(bx+bw+2, bottom+4, text=str(lab), anchor="n", angle=0, font=(KOREAN_FONT, 8))
            # 발생율 꺾은선: 보조축을 차트 우측에 대응
            px=bx+bw+1; py=bottom-chart_h*rate/maxr
            if i: self.canvas.create_line(prevx, prevy, px, py, fill=rate_color, width=2)
            self.canvas.create_oval(px-3,py-3,px+3,py+3,fill=rate_color,outline=rate_color)
            prevx,prevy=px,py
        # 그래프 하단 월별 DATA 표
        table_y = bottom + 58
        row_h = 24
        rows = [("월", labels), ("생산월", prod_vals), ("발생월", occ_vals), ("조립수", assembly_vals), ("PPM", [round(v) for v in rates])]
        # 월 헤더만 옅은 파랑으로 두고, 생산월/발생월/조립수/PPM 데이터 행은 흰색으로 통일한다.
        row_fills = ["#e8f1fb", "#ffffff", "#ffffff", "#ffffff", "#ffffff"]
        for ri, (name, vals) in enumerate(rows):
            yy = table_y + ri*row_h
            self.canvas.create_rectangle(x, yy, x+label_w, yy+row_h, fill=row_fills[ri], outline="#b8c7d6")
            self.canvas.create_text(x+label_w/2, yy+row_h/2, text=name, anchor="center", font=(KOREAN_FONT, 9), fill="#20354b")
            for ci, val in enumerate(vals):
                xx = x+label_w+ci*cell_w
                self.canvas.create_rectangle(xx, yy, xx+cell_w, yy+row_h, fill=row_fills[ri], outline="#c8d3df")
                self.canvas.create_text(xx+cell_w/2, yy+row_h/2, text=f"{val:,}" if isinstance(val,(int,float)) else str(val), anchor="center", font=(KOREAN_FONT, 8))
            total_x = x + label_w + len(labels) * cell_w
            # 화면에서 숨긴 21년 이전 월도 전체 발생월 합계에는 포함한다.
            total_occurrence = sum(occur.values())
            total_production = sum(prod.values())
            total_assembly = int(round(sum(official_assembly.values()))) if official_assembly else sum(assembly_vals)
            # 표의 행 순서: 생산월 → 발생월 → PPM
            total_values = ["합계", total_production, total_occurrence, total_assembly, round(total_production / total_assembly * 1_000_000) if total_assembly else 0]
            tv = total_values[ri]
            total_w = max(cell_w, 112)
            self.canvas.create_rectangle(total_x, yy, total_x+total_w, yy+row_h, fill="#dceaf2" if ri == 0 else "#ffffff", outline="#c8d3df")
            self.canvas.create_text(total_x+total_w/2, yy+row_h/2, text=f"{int(tv):,}" if ri == 3 and isinstance(tv, (int, float)) else (f"{tv:,}" if isinstance(tv,(int,float)) else str(tv)), anchor="center", font=(KOREAN_FONT, 8, "bold"))
        # 실제 마지막 합계 셀 끝까지 스크롤 가능하도록 작업영역 확장
        self.top_canvas.configure(scrollregion=(0, 0, total_x + cell_w + 30, self.top_canvas.winfo_reqheight() or 430))
        # 표 전체 외곽 테두리: 마지막 합계 열까지 연결
        table_right = total_x + max(cell_w, 112)
        table_bottom = table_y + len(rows) * row_h
        # 그래프 테두리도 최종 월/합계 열의 실제 끝까지 연장
        self.canvas.create_line(x, y, table_right, y, fill="#555555", width=1)
        self.canvas.create_line(x, y+h, table_right, y+h, fill="#555555", width=1)
        self.canvas.create_line(table_right, y, table_right, y+h, fill="#555555", width=1)
        self.canvas.create_line(x, table_y, table_right, table_y, fill="#c8d3df", width=1)
        self.canvas.create_line(x, table_bottom, table_right, table_bottom, fill="#c8d3df", width=1)
        self.canvas.create_line(x, table_y, x, table_bottom, fill="#c8d3df", width=1)
        self.canvas.create_line(table_right, table_y, table_right, table_bottom, fill="#c8d3df", width=1)
        self.canvas.create_rectangle(x+w-250,y+10,x+w-230,y+25,fill="#2563eb",outline="")
        self.canvas.create_text(x+w-222,y+17,text="발생월",anchor="w",font=(KOREAN_FONT,10))
        self.canvas.create_rectangle(x+w-145,y+10,x+w-125,y+25,fill="#94a3b8",outline="")
        self.canvas.create_text(x+w-117,y+17,text="생산월",anchor="w",font=(KOREAN_FONT,10))
        self.canvas.create_line(x+w-60,y+17,x+w-35,y+17,fill="#ef4444",width=3)
        self.canvas.create_text(x+w-28,y+17,text="발생율",anchor="w",font=(KOREAN_FONT,10))

    def _inspection_assembly_by_month(self, selected_company="전체"):
        """Read official assembly counts from inspection merge output."""
        cache_key = f"{selected_company or '전체'}|{self.model_var.get()}|{self.name_var.get()}|{self.market_var.get()}"
        if not isinstance(self._inspection_assembly_cache, dict):
            self._inspection_assembly_cache = {}
        if cache_key in self._inspection_assembly_cache:
            return self._inspection_assembly_cache[cache_key]
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "검수종합_현황.xlsx")
        if not os.path.exists(path):
            return {}
        # 검수 Excel은 한 번만 읽고 이후 콤보박스 변경에서는 메모리만 필터링한다.
        if not hasattr(self, "_inspection_records"):
            self._inspection_records = None
        if self._inspection_records is None:
            try:
                book = openpyxl.load_workbook(path, read_only=True, data_only=True)
                sheet = book[book.sheetnames[0]]; iterator = sheet.iter_rows(values_only=True); headers = next(iterator, ())
                normalized = {re.sub(r"\s+", "", clean(value)): index for index, value in enumerate(headers)}
                month_col = normalized.get("해당년월", 0); quantity_col = normalized.get("수량누계", 8)
                customer_col = normalized.get("거래처명", 1); vehicle_col = normalized.get("차종", 12)
                name_col = normalized.get("품명", 4); de_col = normalized.get("D/E", 5)
                records = []
                for row in iterator:
                    if len(row) <= max(month_col, quantity_col): continue
                    month = month_key(row[month_col]);
                    if not month: continue
                    try: quantity = float(str(row[quantity_col]).replace(",", "").strip())
                    except (TypeError, ValueError): quantity = 0
                    records.append((month, clean(row[customer_col]) if len(row) > customer_col else "", clean(row[vehicle_col]) if len(row) > vehicle_col else "", clean(row[name_col]) if len(row) > name_col else "", clean(row[de_col]) if len(row) > de_col else "", quantity))
                book.close(); self._inspection_records = records
            except Exception:
                self._inspection_records = []
        totals = Counter()
        for month, customer, vehicle, part_name, de_value, quantity in self._inspection_records:
            if selected_company == "WIA" or "현대위아" in selected_company:
                if "현대위아" not in customer: continue
            elif selected_company == "HMB":
                if "현대자동차(주)울산" not in customer or vehicle not in ("카파", "카파주물"): continue
            elif "모비스" in selected_company or "MOBIS" in selected_company.upper():
                if not any(token in customer.upper() for token in ("모비스", "MOBIS")): continue
            elif selected_company == "기아" or "기아" in selected_company:
                if "기아" not in customer: continue
            elif selected_company == "HMC" or "현대자동차(주)울산" in selected_company:
                if "현대자동차(주)울산" not in customer or "주물" not in vehicle: continue
            elif selected_company == "현대" or "현대자동차" in selected_company:
                if "현대자동차" not in customer: continue
            if self.market_var.get() != "전체" and de_value != self.market_var.get(): continue
            if (selected_company == "WIA" or "현대위아" in selected_company) and "카파" in self.model_var.get():
                if not any(keyword in part_name.upper() for keyword in ("CONVERTER", "CATALYTIC", "MANIFOLD MODULE", "카파")): continue
            if (selected_company == "WIA" or "현대위아" in selected_company) and "KA4" in self.model_var.get().upper() and "MUFFLER ASSY-FR" in self.name_var.get().upper():
                if "FRT" not in part_name.upper() or "MUFFLER" not in part_name.upper(): continue
            totals[month] += quantity
        result = {str(month): float(value) for month, value in totals.items() if str(month)}
        self._inspection_assembly_cache[cache_key] = result
        return result
        try:
            # 집계에 필요한 열만 스트리밍으로 읽어 pandas 전체 형 변환을 피한다.
            book = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheet = book[book.sheetnames[0]]
            row_iter = sheet.iter_rows(values_only=True)
            headers = next(row_iter, ())
            normalized = {re.sub(r"\s+", "", clean(value)): index for index, value in enumerate(headers)}
            # 병합 파일의 표준 열 위치를 fallback으로 사용해 헤더 표기 변형도 처리한다.
            month_col = normalized.get("해당년월", 0)
            quantity_col = normalized.get("수량누계", 8)
            if month_col is None or quantity_col is None:
                book.close()
                return {}
            customer_col = normalized.get("거래처명", 1)
            vehicle_col = normalized.get("차종", 12)
            name_col = normalized.get("품명", 4)
            de_col = normalized.get("D/E", 5)
            totals = Counter()
            for row in row_iter:
                customer = clean(row[customer_col]) if customer_col is not None and len(row) > customer_col else ""
                vehicle = clean(row[vehicle_col]) if vehicle_col is not None and len(row) > vehicle_col else ""
                if selected_company == "WIA" or "현대위아" in selected_company:
                    if "현대위아" not in customer: continue
                elif selected_company == "HMB":
                    if "현대자동차(주)울산" not in customer:
                        continue
                    vehicle_value = clean(row[vehicle_col]) if vehicle_col is not None and len(row) > vehicle_col else ""
                    if vehicle_value not in ("카파", "카파주물"):
                        continue
                elif "모비스" in selected_company or "MOBIS" in selected_company.upper():
                    if not any(token in customer.upper() for token in ("모비스", "MOBIS")):
                        continue
                elif selected_company == "기아" or "기아" in selected_company:
                    if "기아" not in customer: continue
                elif selected_company == "HMC" or "현대자동차(주)울산" in selected_company:
                    if "현대자동차(주)울산" not in customer or "주물" not in vehicle: continue
                elif selected_company == "현대" or "현대자동차" in selected_company:
                    if "현대자동차" not in customer: continue
                if self.market_var.get() != "전체":
                    de_value = clean(row[de_col]) if de_col is not None and len(row) > de_col else ""
                    if de_value != self.market_var.get():
                        continue
                # WIA 카파 조건은 품명에 CONVERTER/CATALYTIC/MANIFOLD MODULE/카파가 포함된
                # 검수 DATA만 조립수 산정에 사용한다.
                if (selected_company == "WIA" or "현대위아" in selected_company) and "카파" in self.model_var.get():
                    part_name = clean(row[name_col]) if name_col is not None and len(row) > name_col else ""
                    upper_name = part_name.upper()
                    if not any(keyword in upper_name for keyword in ("CONVERTER", "CATALYTIC", "MANIFOLD MODULE", "카파")):
                        continue
                if (selected_company == "WIA" or "현대위아" in selected_company) and "KA4" in self.model_var.get().upper() and "MUFFLER ASSY-FR" in self.name_var.get().upper():
                    part_name = clean(row[name_col]) if name_col is not None and len(row) > name_col else ""
                    if "FRT" not in part_name.upper() or "MUFFLER" not in part_name.upper():
                        continue
                if len(row) <= max(month_col, quantity_col): continue
                month = month_key(row[month_col])
                if not month: continue
                try:
                    quantity = float(str(row[quantity_col]).replace(",", "").strip())
                except (TypeError, ValueError):
                    quantity = 0
                totals[month] += quantity
            book.close()
            result = {str(month): float(value) for month, value in totals.items() if str(month)}
            self._inspection_assembly_cache[cache_key] = result
            return result
        except Exception:
            return {}

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
        # 저장된 업로드 DATA가 있으면 실행 때마다 기본 원본으로 덮어쓰지 않는다.
        try:
            if not app.all_rows:
                app.load(default); app.render()
        except Exception: pass
    app.mainloop()
