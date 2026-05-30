"""工资单助手主界面"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os

import excel_handler

# ── 字体 ───────────────────────────────────────────────────────
_F = "Microsoft YaHei"
FONT_TITLE  = (_F, 18, "bold")
FONT_STEP   = (_F, 11, "bold")
FONT_BODY   = (_F, 10)
FONT_SMALL  = (_F, 9)
FONT_BTN    = (_F, 11, "bold")
FONT_BTN_SM = (_F, 9)

# ── 颜色 ───────────────────────────────────────────────────────
C_BG        = "#f1f5f9"
C_WHITE     = "#ffffff"
C_STEP_HDR  = "#1e40af"
C_PRIMARY   = "#2563eb"
C_TEAL      = "#0891b2"
C_GREEN     = "#16a34a"
C_AMBER     = "#d97706"
C_RED       = "#dc2626"
C_BORDER    = "#cbd5e1"
C_ROW_ODD   = "#ffffff"
C_ROW_EVEN  = "#f0f7ff"
C_TBL_HDR   = "#dbeafe"

TEMPLATES_FOLDER = "templates"


class WageHelperApp:

    def __init__(self, root: tk.Tk, app_dir: str = None):
        self.root = root
        self.root.title("工资单助手")
        self.root.geometry("900x740")
        self.root.configure(bg=C_BG)
        self.root.resizable(True, True)
        self.root.minsize(720, 560)

        self.app_dir       = app_dir or os.path.dirname(os.path.abspath(__file__))
        self.templates_dir = os.path.join(self.app_dir, TEMPLATES_FOLDER)
        self.temp_bank_path = os.path.join(self.app_dir, "~temp_bank_preview.xlsx")

        self.bank_path  = tk.StringVar()
        self.wage_path  = tk.StringVar()
        self.workers    = []
        self.status_msg = tk.StringVar(value="就绪")

        self._setup_styles()
        self._build_ui()
        self._auto_detect_template()

    # ── 样式 ────────────────────────────────────────────────────

    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("Treeview",
                    font=FONT_BODY, rowheight=27,
                    background=C_ROW_ODD, fieldbackground=C_ROW_ODD,
                    borderwidth=0)
        s.configure("Treeview.Heading",
                    font=FONT_STEP, background=C_TBL_HDR, relief="flat", padding=(8, 5))
        s.map("Treeview",         background=[("selected", "#bfdbfe")])
        s.map("Treeview.Heading", background=[("active",   "#bfdbfe")])

    # ── 主界面 ──────────────────────────────────────────────────

    def _build_ui(self):
        # 顶部标题栏
        title_bar = tk.Frame(self.root, bg=C_STEP_HDR, padx=16, pady=12)
        title_bar.pack(fill=tk.X)
        tk.Label(title_bar, text="工资单助手",
                 font=FONT_TITLE, bg=C_STEP_HDR, fg="white").pack(side=tk.LEFT)
        tk.Label(title_bar, text="  帮助您快速填写银行工资表单",
                 font=(_F, 11), bg=C_STEP_HDR, fg="#93c5fd").pack(side=tk.LEFT, pady=4)

        # 主内容区
        main = tk.Frame(self.root, bg=C_BG, padx=12, pady=8)
        main.pack(fill=tk.BOTH, expand=True)

        self._build_step1(main)
        self._build_step2(main)
        self._build_step4(main)   # 先 pack 到底部，步骤3才能正确填充中间
        self._build_step3(main)   # 最后 pack，填充剩余空间

        # 底部状态栏
        sb = tk.Frame(self.root, bg="#e2e8f0", height=28)
        sb.pack(fill=tk.X, side=tk.BOTTOM)
        sb.pack_propagate(False)
        tk.Label(sb, textvariable=self.status_msg,
                 font=FONT_SMALL, bg="#e2e8f0", fg="#475569", padx=10
                 ).pack(side=tk.LEFT, pady=5)

    # ── 步骤框架工厂 ─────────────────────────────────────────────

    def _step_frame(self, parent, num: int, title: str, expand=False, bottom=False) -> tk.Frame:
        wrapper = tk.Frame(parent, bg=C_BG)
        if bottom:
            wrapper.pack(fill=tk.X, side=tk.BOTTOM, pady=4)
        else:
            wrapper.pack(fill=tk.BOTH if expand else tk.X, expand=expand, pady=4)

        hdr = tk.Frame(wrapper, bg=C_STEP_HDR, padx=14)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=f"步骤 {num}  —  {title}",
                 font=FONT_STEP, bg=C_STEP_HDR, fg="white", pady=8).pack(side=tk.LEFT)

        body = tk.Frame(wrapper, bg=C_WHITE, bd=1, relief="solid",
                        highlightbackground=C_BORDER)
        body.pack(fill=tk.BOTH, expand=expand)
        return body

    # ── 步骤1：选择文件 ──────────────────────────────────────────

    def _build_step1(self, parent):
        body = self._step_frame(parent, 1, "选择文件")
        pad  = tk.Frame(body, bg=C_WHITE, padx=16, pady=12)
        pad.pack(fill=tk.X)

        # 银行表格模板（自动加载，无需选择）
        tk.Label(pad, text="银行表格模板", font=FONT_STEP,
                 bg=C_WHITE, fg="#1e3a5f").pack(anchor=tk.W)

        self.tmpl_status_var = tk.StringVar(value="")
        self.tmpl_status_lbl = tk.Label(pad, textvariable=self.tmpl_status_var,
                                         font=FONT_BODY, bg=C_WHITE, fg=C_AMBER)
        self.tmpl_status_lbl.pack(anchor=tk.W, pady=(3, 0))

        self._btn(pad, "📂 打开模板文件夹", self._open_templates_dir,
                  color=C_TEAL, small=True).pack(anchor=tk.W, pady=(5, 0))

        # 分割线
        tk.Frame(pad, bg=C_BORDER, height=1).pack(fill=tk.X, pady=10)

        # 工资信息表（每次选择）
        tk.Label(pad, text="工资信息表", font=FONT_STEP,
                 bg=C_WHITE, fg="#1e3a5f").pack(anchor=tk.W)

        file_row = tk.Frame(pad, bg=C_WHITE)
        file_row.pack(fill=tk.X, pady=(5, 0))
        tk.Entry(file_row, textvariable=self.wage_path, font=FONT_SMALL,
                 state="readonly", readonlybackground="#f8fafc",
                 relief="solid", bd=1
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._btn(file_row, "选择文件", self._browse_wage,
                  small=True).pack(side=tk.LEFT, padx=(6, 0))

        # 示例数据（测试用）
        tk.Label(pad, text="（测试用）", font=FONT_SMALL,
                 bg=C_WHITE, fg="#9ca3af").pack(anchor=tk.W, pady=(10, 0))
        self._btn(pad, "载入示例数据", self._load_sample,
                  color="#94a3b8", small=True).pack(anchor=tk.W, pady=(3, 0))

    # ── 步骤2：自动填写 ──────────────────────────────────────────

    def _build_step2(self, parent):
        body = self._step_frame(parent, 2, "填写银行表格")
        pad  = tk.Frame(body, bg=C_WHITE, padx=16, pady=12)
        pad.pack(fill=tk.X)

        tk.Label(pad, text="读取工资表，自动填写银行表格，然后在下方核对。",
                 font=FONT_SMALL, bg=C_WHITE, fg="#6b7280").pack(anchor=tk.W, pady=(0, 8))

        row = tk.Frame(pad, bg=C_WHITE)
        row.pack(fill=tk.X)
        self.fill_btn = self._btn(row, "填写银行表格", self._do_fill)
        self.fill_btn.pack(side=tk.LEFT)
        self.fill_lbl = tk.Label(row, text="", font=FONT_BODY, bg=C_WHITE, fg=C_GREEN)
        self.fill_lbl.pack(side=tk.LEFT, padx=14)

    # ── 步骤3：核对信息 ──────────────────────────────────────────

    def _build_step3(self, parent):
        body = self._step_frame(parent, 3, "核对银行表格", expand=True)
        pad  = tk.Frame(body, bg=C_WHITE, padx=16, pady=8)
        pad.pack(fill=tk.BOTH, expand=True)

        top = tk.Frame(pad, bg=C_WHITE)
        top.pack(fill=tk.X, pady=(0, 6))
        tk.Label(top, text="以下是填写好的银行表格内容，双击单元格可修改",
                 font=FONT_SMALL, bg=C_WHITE, fg="#64748b").pack(side=tk.LEFT)
        self._btn(top, "用 Excel 打开查看", self._open_preview_in_excel,
                  color=C_TEAL, small=True).pack(side=tk.RIGHT)
        self.summary_var = tk.StringVar(value="")
        tk.Label(top, textvariable=self.summary_var,
                 font=FONT_SMALL, bg="#eff6ff", fg="#1e40af",
                 padx=8, pady=3, relief="solid").pack(side=tk.RIGHT)

        cols = ("no", "name", "id", "bank_account", "bank_name", "amount")
        col_cfg = (
            ("no",           "序号",     55),
            ("name",         "姓名",    120),
            ("id",           "身份证号", 190),
            ("bank_account", "银行账号", 210),
            ("bank_name",    "开户行",   155),
            ("amount",       "工资（元）", 120),
        )

        tbl = tk.Frame(pad, bg=C_WHITE)
        tbl.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tbl, columns=cols, show="headings", height=13)
        for cid, label, width in col_cfg:
            self.tree.heading(cid, text=label, command=lambda c=cid: self._sort_col(c))
            self.tree.column(cid, width=width, minwidth=40, stretch=True)

        ys = ttk.Scrollbar(tbl, orient="vertical",  command=self.tree.yview)
        xs = ttk.Scrollbar(tbl, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        tbl.grid_rowconfigure(0, weight=1)
        tbl.grid_columnconfigure(0, weight=1)

        self.tree.tag_configure("odd",  background=C_ROW_ODD)
        self.tree.tag_configure("even", background=C_ROW_EVEN)
        self.tree.bind("<Double-1>", self._on_dbl_click)
        self._edit_widget = None

    # ── 步骤4：保存文件 ──────────────────────────────────────────

    def _build_step4(self, parent):
        body = self._step_frame(parent, 4, "保存文件", bottom=True)
        pad  = tk.Frame(body, bg=C_WHITE, padx=16, pady=12)
        pad.pack(fill=tk.X)

        row = tk.Frame(pad, bg=C_WHITE)
        row.pack(fill=tk.X)
        self._btn(row, "保存文件", self._do_save, color=C_GREEN).pack(side=tk.LEFT)
        self.save_lbl = tk.Label(row, text="", font=FONT_BODY, bg=C_WHITE, fg=C_GREEN)
        self.save_lbl.pack(side=tk.LEFT, padx=14)

    # ── 按钮工厂 ─────────────────────────────────────────────────

    def _btn(self, parent, text: str, cmd, color=None, small=False) -> tk.Button:
        color = color or C_PRIMARY
        font  = FONT_BTN_SM if small else FONT_BTN
        px, py = (10, 5) if small else (18, 9)
        return tk.Button(parent, text=text, command=cmd, font=font,
                         bg=color, fg="white", relief="flat",
                         padx=px, pady=py, cursor="hand2",
                         activeforeground="white",
                         activebackground=_darken(color))

    # ── 模板自动检测 ─────────────────────────────────────────────

    def _auto_detect_template(self):
        os.makedirs(self.templates_dir, exist_ok=True)
        try:
            files = sorted(
                f for f in os.listdir(self.templates_dir)
                if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~")
            )
        except OSError:
            files = []

        if len(files) == 1:
            self.bank_path.set(os.path.join(self.templates_dir, files[0]))
            self.tmpl_status_var.set(f"✓ 已加载：{files[0]}")
            self.tmpl_status_lbl.config(fg=C_GREEN)
            self._status(f"银行模板已自动加载：{files[0]}")
        elif len(files) > 1:
            self.bank_path.set(os.path.join(self.templates_dir, files[0]))
            self.tmpl_status_var.set(f"✓ 已加载：{files[0]}（文件夹中共有 {len(files)} 个模板）")
            self.tmpl_status_lbl.config(fg=C_AMBER)
        else:
            self.tmpl_status_var.set("⚠ 未找到模板，请将银行表格格式文件放入模板文件夹")
            self.tmpl_status_lbl.config(fg=C_AMBER)
            self._status("请将银行表格模板放入 templates 文件夹")

    def _open_templates_dir(self):
        os.makedirs(self.templates_dir, exist_ok=True)
        os.startfile(self.templates_dir)

    # ── 选择工资表 ───────────────────────────────────────────────

    def _browse_wage(self):
        p = filedialog.askopenfilename(
            title="选择工资信息表",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        if p:
            self.wage_path.set(p)
            self._status(f"工资表已选择：{os.path.basename(p)}")

    # ── 示例数据 ─────────────────────────────────────────────────

    def _load_sample(self):
        sample_path = os.path.join(self.app_dir, "示例工资表.xlsx")
        if os.path.exists(sample_path):
            try:
                self.workers = excel_handler.read_wage_form(sample_path)
                self.wage_path.set(sample_path)
                self._refresh_table()
                n = len(self.workers)
                self.fill_lbl.config(text=f"✓ 示例数据已载入（{n} 条）", fg=C_GREEN)
                self._status(f"已从示例文件载入 {n} 条数据")
            except Exception as exc:
                messagebox.showerror("错误", f"读取示例文件失败：\n{exc}")
        else:
            messagebox.showwarning("提示",
                f"未找到示例文件：\n{sample_path}\n\n请确认 示例工资表.xlsx 在程序同一文件夹中。")

    # ── 填写银行表格 ─────────────────────────────────────────────

    def _do_fill(self):
        wage = self.wage_path.get()
        bank = self.bank_path.get()

        # 示例数据：wage_path 已指向示例文件，正常走下面流程
        if not wage:
            messagebox.showwarning("提示", "请先选择工资信息表！")
            return
        if not bank:
            messagebox.showwarning("提示", "未找到银行表格模板，请检查 templates 文件夹。")
            return

        try:
            self.fill_btn.config(state=tk.DISABLED, text="填写中…")
            self.root.update()

            # 1. 读取工资表
            workers_raw = excel_handler.read_wage_form(wage)

            # 2. 填写进银行模板，存为临时预览文件
            excel_handler.fill_bank_form(bank, self.temp_bank_path, workers_raw)

            # 3. 读回已填写的银行表格，这才是要核对的内容
            self.workers = excel_handler.read_bank_form(self.temp_bank_path)
            self._refresh_table()

            n = len(self.workers)
            self.fill_lbl.config(text=f"✓ 已填写 {n} 条记录，请在下方核对", fg=C_GREEN)
            self._status(f"银行表格已填写完成，共 {n} 条记录，请核对后保存")

        except Exception as exc:
            messagebox.showerror("错误", f"填写失败：\n\n{exc}")
            self.fill_lbl.config(text="✗ 填写失败", fg=C_RED)
            self._status("填写失败")
        finally:
            self.fill_btn.config(state=tk.NORMAL, text="填写银行表格")

    def _open_preview_in_excel(self):
        if os.path.exists(self.temp_bank_path):
            os.startfile(self.temp_bank_path)
        else:
            messagebox.showwarning("提示", "请先点击 [填写银行表格] 按钮。")

    # ── 保存文件 ─────────────────────────────────────────────────

    def _do_save(self):
        if not self.workers:
            messagebox.showwarning("提示", "没有数据，请先读取工资表。")
            return
        bank = self.bank_path.get()
        if not bank:
            messagebox.showwarning("提示", "未找到银行表格模板，请检查 templates 文件夹。")
            return

        self._sync_table()

        default = f"银行工资表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out = filedialog.asksaveasfilename(
            title="保存文件",
            initialfile=default,
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])
        if not out:
            return

        try:
            excel_handler.fill_bank_form(bank, out, self.workers)
            self.save_lbl.config(text=f"✓ 已保存：{os.path.basename(out)}", fg=C_GREEN)
            self._status(f"文件已保存：{out}")
            if messagebox.askyesno("保存成功", f"文件已成功保存！\n\n{out}\n\n是否打开文件所在文件夹？"):
                os.startfile(os.path.dirname(out))
        except Exception as exc:
            messagebox.showerror("错误", f"保存失败：\n\n{exc}")
            self.save_lbl.config(text="✗ 保存失败", fg=C_RED)

    # ── 表格 ─────────────────────────────────────────────────────

    def _refresh_table(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        total = 0.0
        for i, w in enumerate(self.workers):
            amt = _to_float(w.get("amount", 0))
            total += amt
            self.tree.insert("", tk.END, iid=str(i),
                             tags=(("even" if i % 2 == 0 else "odd"),),
                             values=(i + 1,
                                     w.get("name", ""),
                                     w.get("id", ""),
                                     w.get("bank_account", ""),
                                     w.get("bank_name", ""),
                                     f"{amt:,.2f}"))
        n = len(self.workers)
        self.summary_var.set(f"  共 {n} 人  |  合计工资：¥{total:,.2f}  ")

    def _sync_table(self):
        for i, iid in enumerate(self.tree.get_children()):
            if i >= len(self.workers):
                break
            v = self.tree.item(iid)["values"]
            self.workers[i].update({
                "name":         str(v[1]),
                "id":           str(v[2]),
                "bank_account": str(v[3]),
                "bank_name":    str(v[4]),
                "amount":       _to_float(str(v[5])),
            })

    def _recompute_summary(self):
        total = sum(_to_float(str(self.tree.item(iid)["values"][5]))
                    for iid in self.tree.get_children())
        n = len(self.tree.get_children())
        self.summary_var.set(f"  共 {n} 人  |  合计工资：¥{total:,.2f}  ")

    # ── 单元格编辑 ───────────────────────────────────────────────

    def _on_dbl_click(self, event):
        if self._edit_widget:
            self._edit_widget.destroy()
            self._edit_widget = None

        row = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row or not col:
            return

        col_idx = int(col[1:]) - 1
        if col_idx == 0:
            return

        bbox = self.tree.bbox(row, col)
        if not bbox:
            return
        x, y, w, h = bbox

        vals = list(self.tree.item(row)["values"])
        var  = tk.StringVar(value=str(vals[col_idx]))

        entry = tk.Entry(self.tree, textvariable=var, font=FONT_BODY,
                         relief="solid", bd=1, bg="#fefce8")
        entry.place(x=x, y=y, width=w, height=h)
        entry.select_range(0, tk.END)
        entry.focus()
        self._edit_widget = entry

        def commit(evt=None):
            vals[col_idx] = var.get()
            self.tree.item(row, values=vals)
            if col_idx == 5:
                self._recompute_summary()
            entry.destroy()
            self._edit_widget = None

        def cancel(evt=None):
            entry.destroy()
            self._edit_widget = None

        entry.bind("<Return>",   commit)
        entry.bind("<Tab>",      commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>",   cancel)

    # ── 列排序 ───────────────────────────────────────────────────

    def _sort_col(self, col):
        idx_map = {"no": 0, "name": 1, "id": 2,
                   "bank_account": 3, "bank_name": 4, "amount": 5}
        idx     = idx_map.get(col, 0)
        attr    = f"_sort_{col}_rev"
        reverse = getattr(self, attr, False)
        setattr(self, attr, not reverse)

        items = [(self.tree.item(iid)["values"], iid)
                 for iid in self.tree.get_children()]

        def key(x):
            v = x[0][idx]
            return _to_float(str(v)) if idx in (0, 5) else str(v)

        items.sort(key=key, reverse=reverse)
        for i, (_, iid) in enumerate(items):
            self.tree.move(iid, "", i)
            self.tree.item(iid, tags=(("even" if i % 2 == 0 else "odd"),))

    # ── 状态栏 ───────────────────────────────────────────────────

    def _status(self, msg: str):
        self.status_msg.set(msg)
        self.root.update_idletasks()


# ── 辅助函数 ──────────────────────────────────────────────────────

def _darken(hex_color: str) -> str:
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"#{max(0,r-35):02x}{max(0,g-35):02x}{max(0,b-35):02x}"

def _to_float(val) -> float:
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0
