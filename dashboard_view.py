import os
import tkinter as tk
from tkinter import ttk


class DashboardView:
    """Reference-image-inspired Tkinter shell; analysis remains on ClaimDashboard."""

    def __init__(self, app):
        self.app = app

    def build(self):
        a = self.app
        a.image_refs = []
        # These small assets are consumed by the existing chart renderer.
        for attr, filename, scale in (
            ("issue_icon", "현상명.png", 14),
            ("usage_icon", "사용기간_달력.png", 14),
            ("mileage_icon", "주행거리.png", 14),
            ("country_icon", "발생국가.png", 16),
            ("analysis_icon", "free-icon-growth-3281306.png", 22),
        ):
            raw = a._load_image(filename)
            icon = raw.subsample(scale, scale) if raw else None
            setattr(a, attr, icon)
            if icon:
                a.image_refs.append(icon)
        a.option_add("*Font", ("Malgun Gothic", 10))
        style = ttk.Style(a)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Modern.TFrame", background="#EEF6FF")
        style.configure("Modern.TLabel", background="#FFFFFF", foreground="#102A4C")
        style.configure("Modern.TButton", background="#1677E8", foreground="white", padding=(14, 8), font=("Malgun Gothic", 10, "bold"))
        style.configure("Modern.TCombobox", fieldbackground="#F8FBFF", background="#F8FBFF", padding=5)
        a.configure(background="#EEF6FF")
        body = tk.PanedWindow(a, orient="horizontal", sashwidth=0, bd=0, relief="flat", bg="#EEF6FF")
        body.pack(fill="both", expand=True)
        side = tk.Frame(body, width=220, bg="#082B52")
        side.pack_propagate(False); body.add(side, minsize=220, width=220, stretch="never")
        self._sidebar(side)
        content = tk.Frame(body, bg="#EEF6FF"); body.add(content, minsize=900, stretch="always")
        self._content(content)

    def _sidebar(self, side):
        a = self.app
        graph_raw = a._load_image("free-icon-graph-2848907.png")
        database_raw = a._load_image("free-icon-database-16778425.png")
        graph_icon = graph_raw.subsample(max(1, graph_raw.width() // 22), max(1, graph_raw.height() // 22)) if graph_raw else None
        database_icon = database_raw.subsample(max(1, database_raw.width() // 22), max(1, database_raw.height() // 22)) if database_raw else None
        if graph_icon: a.image_refs.append(graph_icon)
        if database_icon: a.image_refs.append(database_icon)
        logo = a._load_image("company_logo.png")
        if logo:
            small = logo.subsample(max(1, logo.width() // 150), max(1, logo.height() // 48))
            tk.Label(side, image=small, bg="#082B52").pack(pady=(24, 10)); a.image_refs.append(small)
        tk.Label(side, text="클레임 자동 분석", fg="white", bg="#082B52", font=("Malgun Gothic", 18, "bold")).pack(anchor="w", padx=20)
        tk.Label(side, text="Claim Analytics", fg="#A9C8E9", bg="#082B52", font=("Segoe UI", 10)).pack(anchor="w", padx=20, pady=(2, 24))
        st = ttk.Style(a); st.configure("ModernSidebar.Treeview", background="#082B52", fieldbackground="#082B52", foreground="white", rowheight=36, borderwidth=0, font=("Malgun Gothic", 11)); st.map("ModernSidebar.Treeview", background=[("selected", "#1677E8")])
        menu = ttk.Treeview(side, show="tree", selectmode="browse", style="ModernSidebar.Treeview", height=2); menu.pack(fill="x", padx=10)
        menu.tag_configure("section_header", background="#1769D4", foreground="white")
        root = menu.insert("", "end", text="클레임 분석", image=graph_icon, iid="dashboard", open=True, tags=("section_header",))
        menu.insert(root, "end", text="   보고서 출력", iid="report")
        menu.selection_set(root); menu.bind("<<TreeviewSelect>>", a._on_sidebar_select); a.sidebar_menu = menu
        for title, iid, callback, icon in (("고객사별 DATA 저장", "customer_root", a._on_customer_menu_select, database_icon), ("검수폴더 병합", "merge_root", a._on_merge_menu_select, database_icon)):
            t = ttk.Treeview(side, show="tree", selectmode="browse", style="ModernSidebar.Treeview", height=2); t.pack(fill="x", padx=10, pady=(18, 0)); t.tag_configure("section_header", background="#1769D4", foreground="white"); r = t.insert("", "end", text=title, image=icon, iid=iid, open=True, tags=("section_header",)); child = "customer_upload" if iid == "customer_root" else "merge_app"; label = "   DATA 업로드" if iid == "customer_root" else "   검수폴더 병합앱"; t.insert(r, "end", text=label, iid=child); t.bind("<<TreeviewSelect>>", callback)
            if iid == "customer_root": a.customer_menu = t
            else: a.merge_menu = t
        tk.Label(side, text="품질로 더 나은 내일을 만듭니다.\n\nBetter Quality\nA Brighter Tomorrow", fg="#B8D1E8", bg="#082B52", justify="left", anchor="w", font=("Malgun Gothic", 9)).pack(side="bottom", fill="x", padx=20, pady=22)

    def _content(self, content):
        a = self.app
        top = tk.Frame(content, bg="#FFFFFF", padx=16, pady=14, highlightbackground="#D7E6F5", highlightthickness=1); top.pack(fill="x", padx=18, pady=(18, 8))
        tk.Label(top, text="데이터 원본", bg="white", fg="#102A4C", font=("Malgun Gothic", 9, "bold")).pack(side="left")
        a.file_label = ttk.Label(top, text="파일을 선택하세요", style="Modern.TLabel", padding=(10, 0)); a.file_label.pack(side="left", padx=(8, 20))
        self._combo(top, "고객사", "company_var", "company_combo", 12, a.company_changed)
        self._combo(top, "차종", "model_var", "model_combo", 18, a.model_changed)
        self._combo(top, "품명", "name_var", "name_combo", 24, a.open_name_selector)
        self._combo(top, "구분", "market_var", "market_combo", 10, a.apply_filters)
        self._combo(top, "품번", "part_var", "part_combo", 16, a.open_part_selector)
        self._combo(top, "현상분석 구분", "issue_mode_var", "issue_mode_combo", 11, a.render,
                    values=("코드기준", "분석기준"))
        ttk.Button(top, text="분석 시작", command=a.refresh_analysis, style="Modern.TButton").pack(side="right", padx=(12, 0))
        a.status = ttk.Label(content, text="", background="#EEF6FF", foreground="#58718E", padding=(20, 4)); a.status.pack(fill="x")
        a.kpi_frame = tk.Frame(content, bg="#EEF6FF"); a.kpi_frame.pack(fill="x", padx=12, pady=6); a.kpi_labels=[]
        kpi_specs = (
            ("▤", "총 클레임", "#1677E8", "free-icon-data-10139543.png"),
            ("▥", "총 발생건수", "#1677E8", "free-icon-bar-chart-8696653.png"),
            ("⚙", "총 생산건수", "#1677E8", "free-icon-gear-8680172.png"),
            ("◈", "발생률(PPM)", "#F97316", "free-icon-percent-3097292.png"),
        )
        for fallback_icon, title, color, filename in kpi_specs:
            card_shell=tk.Frame(a.kpi_frame,bg="#D7E2EF",bd=0);card_shell.pack(side="left",fill="x",expand=True,padx=6,pady=(0,3))
            card=tk.Canvas(card_shell,bg="#D7E2EF",highlightthickness=0,height=82,bd=0);card.pack(fill="both",expand=True)
            def rounded(canvas, width, height):
                radius=12; canvas.delete("panel")
                canvas.create_rectangle(2+radius, 3, width-2-radius, height-2, fill="white", outline="white", tags="panel")
                canvas.create_rectangle(2, 3+radius, width-2, height-2-radius, fill="white", outline="white", tags="panel")
                for cx,cy,start in ((2+radius,3+radius,90),(width-2-radius,3+radius,0),(width-2-radius,height-2-radius,270),(2+radius,height-2-radius,180)):
                    canvas.create_arc(cx-radius,cy-radius,cx+radius,cy+radius,start=start,extent=90,fill="white",outline="white",tags="panel")
            card.update_idletasks(); rounded(card, max(card.winfo_reqwidth(), 300), 82)
            inner=tk.Frame(card,bg="white",bd=0); card.create_window((14,10),window=inner,anchor="nw")
            raw_icon = a._load_image(filename)
            card_icon = raw_icon.subsample(max(1, raw_icon.width() // 48), max(1, raw_icon.height() // 48)) if raw_icon else None
            if card_icon:
                a.image_refs.append(card_icon)
                tk.Label(inner, image=card_icon, bg="white").pack(side="left",padx=(0,12))
            else:
                tk.Label(inner,text=fallback_icon,bg=color,fg="white",font=("Malgun Gothic",22,"bold"),width=2).pack(side="left",padx=(0,12))
            box=tk.Frame(inner,bg="white");box.pack(side="left");tk.Label(box,text=title,bg="white",fg="#58718E",font=("Malgun Gothic",9)).pack(anchor="w");v=tk.Label(box,text="-",bg="white",fg="#102A4C",font=("Malgun Gothic",21,"bold"));v.pack(anchor="w");a.kpi_labels.append(v)
        a.notebook=ttk.Notebook(content);a.notebook.pack(fill="both",expand=True,padx=18,pady=(4,18));a.dashboard_tab=ttk.Frame(a.notebook);a.detail_tab=ttk.Frame(a.notebook);a.notebook.add(a.dashboard_tab,text="분석 대시보드");a.notebook.add(a.detail_tab,text="원본 데이터")
        a.dashboard_tab.rowconfigure(1,weight=1);a.dashboard_tab.columnconfigure(0,weight=1);a.top_frame=tk.Frame(a.dashboard_tab,bg="white",highlightbackground="#D7E6F5",highlightthickness=1);a.top_frame.grid(row=0,column=0,sticky="ew");a.top_title_frame=tk.Frame(a.top_frame,bg="white");a.top_title_frame.pack(anchor="w",padx=18,pady=(12,4));title_icon=a._load_image("free-icon-data-analytics-8909435.png");title_icon_small=title_icon.subsample(max(1,title_icon.width()//34),max(1,title_icon.height()//34)) if title_icon else None
        if title_icon_small: a.image_refs.append(title_icon_small); tk.Label(a.top_title_frame,image=title_icon_small,bg="white").pack(side="left",padx=(0,8))
        a.top_title=tk.Label(a.top_title_frame,text="1. 월별 클레임 발생현황",font=("Malgun Gothic",18,"bold"),bg="white",fg="#102A4C");a.top_title.pack(side="left");a.month_title_frame=tk.Frame(a.top_frame,bg="white");a.month_title_frame.pack(anchor="w");a.top_canvas=tk.Canvas(a.top_frame,bg="white",height=430,highlightthickness=0);a.top_canvas.pack(fill="x",expand=True);a.top_scroll=ttk.Scrollbar(a.top_frame,orient="horizontal",command=a.top_canvas.xview);a.top_scroll.pack(fill="x");a.top_canvas.configure(xscrollcommand=a.top_scroll.set);a.fixed_table=tk.Canvas(a.top_frame,bg="white",highlightthickness=0,width=105,height=125);a.bottom_canvas=tk.Canvas(a.dashboard_tab,bg="#EEF6FF",highlightthickness=0);a.bottom_canvas.grid(row=1,column=0,sticky="nsew");a.bottom_canvas.bind("<Configure>",lambda e:a._schedule_render());a.canvas=a.bottom_canvas;a.tree=ttk.Treeview(a.detail_tab,show="headings");a.tree.pack(fill="both",expand=True)

    def _combo(self, parent, label, var_name, combo_name, width, command, values=None):
        a=self.app; box=tk.Frame(parent,bg="white");box.pack(side="left",padx=5);tk.Label(box,text=label,bg="white",fg="#58718E",font=("Malgun Gothic",8,"bold")).pack(anchor="w");default = values[0] if values else "전체";var=tk.StringVar(value=default);setattr(a,var_name,var);combo=ttk.Combobox(box,textvariable=var,state="readonly",width=width,style="Modern.TCombobox");
        if values: combo["values"] = values
        combo.pack();setattr(a,combo_name,combo);combo.bind("<<ComboboxSelected>>",lambda e:command())
        if combo_name == "name_combo": combo.bind("<Button-1>", a._name_combo_clicked)
        if combo_name == "part_combo": combo.bind("<Button-1>", a._part_combo_clicked)
