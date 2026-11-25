import csv
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from urllib.parse import urljoin

import customtkinter as ctk
import openpyxl
from openpyxl.styles import Font, PatternFill
import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

# إعداد المظهر الافتراضي (Light Mode)
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/130.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,de;q=0.8,ar;q=0.7",
}


def normalize_url(raw: str) -> str:
    url = raw.strip()
    if not url:
        return ""
    if not re.match(r"^https?://", url, flags=re.IGNORECASE):
        url = "https://" + url
    return url


def extract_meta_or_canonical_redirect(html_text: str, base_url: str) -> str:
    """يبحث عن تحويل Meta Refresh أو Canonical داخل HTML في حال عدم وجود تحويل HTTP."""
    if not html_text:
        return ""
    try:
        soup = BeautifulSoup(html_text, "html.parser")
        meta = soup.find("meta", attrs={"http-equiv": re.compile(r"^refresh$", re.I)})
        if meta and meta.get("content"):
            match = re.search(r"url\s*=\s*['\"]?([^'\";]+)", str(meta["content"]), re.I)
            if match:
                return urljoin(base_url, match.group(1).strip())

        canonical = soup.find(
            "link",
            rel=lambda v: v and "canonical" in [x.lower() for x in (v if isinstance(v, list) else [v])],
        )
        if canonical and canonical.get("href"):
            return urljoin(base_url, str(canonical["href"]).strip())
    except Exception:
        pass
    return ""


def extract_title_from_html(html_text: str) -> str:
    if not html_text:
        return ""
    try:
        soup = BeautifulSoup(html_text, "html.parser")
        if soup.title and soup.title.string:
            return re.sub(r"\s+", " ", soup.title.string).strip()
    except Exception:
        pass
    return ""


class PlaywrightChecker:
    """يدير متصفح Edge أو Chrome المثبت على النظام عبر Playwright بدون تضمين Chromium."""

    def __init__(self, timeout_sec: float = 20.0, headless: bool = True, browser_channel: str = "auto"):
        self.timeout_ms = int(timeout_sec * 1000)
        self.headless = headless
        self.browser_channel = browser_channel
        self.active_channel = ""
        self._playwright = None
        self._browser = None
        self._context = None
        self._lock = threading.Lock()

    def _launch_system_browser(self):
        if self.browser_channel == "msedge":
            channels = ["msedge"]
        elif self.browser_channel == "chrome":
            channels = ["chrome"]
        else:
            channels = ["msedge", "chrome"]

        last_err = None
        for ch in channels:
            try:
                browser = self._playwright.chromium.launch(channel=ch, headless=self.headless)
                self.active_channel = ch
                return browser
            except Exception as exc:
                last_err = exc

        raise RuntimeError(
            f"تعذر تشغيل متصفح النظام ({', '.join(channels)}). تأكد من تثبيت Microsoft Edge أو Google Chrome.\n{last_err}"
        )

    def start(self):
        with self._lock:
            if self._browser is None:
                self._playwright = sync_playwright().start()
                self._browser = self._launch_system_browser()
                self._context = self._browser.new_context(
                    user_agent=DEFAULT_HEADERS["User-Agent"],
                    ignore_https_errors=True,
                )

    def check_url(self, url: str, wait_extra_sec: float = 2.0) -> dict:
        start_t = time.perf_counter()
        chain = [url]
        page = None
        try:
            with self._lock:
                if self._browser is None:
                    self.start()
                page = self._context.new_page()

            page.set_default_timeout(self.timeout_ms)

            def on_response(resp):
                try:
                    if 300 <= resp.status < 400 and resp.headers.get("location"):
                        target = urljoin(resp.url, resp.headers["location"])
                        if not chain or chain[-1] != target:
                            chain.append(target)
                except Exception:
                    pass

            page.on("response", on_response)

            status_code = ""
            resp = page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms)
            if resp is not None:
                status_code = resp.status

            try:
                page.wait_for_load_state("networkidle", timeout=min(self.timeout_ms, 10000))
            except Exception:
                pass

            if wait_extra_sec > 0:
                time.sleep(wait_extra_sec)

            final_url = page.url or url
            if chain[-1] != final_url:
                chain.append(final_url)

            title = ""
            size_bytes = 0
            try:
                title = re.sub(r"\s+", " ", page.title() or "").strip()
                content = page.content() or ""
                size_bytes = len(content.encode("utf-8", errors="ignore"))
            except Exception:
                pass

            elapsed_ms = int((time.perf_counter() - start_t) * 1000)
            redirect_count = max(0, len(chain) - 1)
            engine_label = f"PW-{self.active_channel or 'browser'}"

            return {
                "original_url": url,
                "final_url": final_url,
                "redirect_count": redirect_count,
                "redirect_chain": " -> ".join(chain),
                "status_code": str(status_code),
                "title": title,
                "size_bytes": size_bytes,
                "response_ms": elapsed_ms,
                "mode": engine_label,
                "error": "",
                "status_bucket": "success" if (isinstance(status_code, int) and 200 <= status_code < 400) else "error",
            }
        except Exception as exc:
            elapsed_ms = int((time.perf_counter() - start_t) * 1000)
            return {
                "original_url": url,
                "final_url": page.url if page else url,
                "redirect_count": max(0, len(chain) - 1),
                "redirect_chain": " -> ".join(chain),
                "status_code": "ERR",
                "title": "",
                "size_bytes": 0,
                "response_ms": elapsed_ms,
                "mode": f"PW-{self.active_channel or 'browser'}",
                "error": str(exc).splitlines()[0][:180],
                "status_bucket": "error",
            }
        finally:
            if page is not None:
                try:
                    page.close()
                except Exception:
                    pass

    def close(self):
        with self._lock:
            for obj in (self._context, self._browser, self._playwright):
                if obj is not None:
                    try:
                        if hasattr(obj, "close"):
                            obj.close()
                        elif hasattr(obj, "stop"):
                            obj.stop()
                    except Exception:
                        pass
            self._context = None
            self._browser = None
            self._playwright = None


class URLCheckerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("URL Checker Pro — By ENG Ahmed Bahgat")
        self.geometry("1280x830")
        self.minsize(1050, 680)

        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.is_running = False

        self.results: list[dict] = []
        self.pw_checker: PlaywrightChecker | None = None

        self._build_ui()
        self._apply_treeview_style()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.grid(row=0, column=0, padx=14, pady=(12, 6), sticky="ew")
        header_frame.grid_columnconfigure(1, weight=1)

        title_lbl = ctk.CTkLabel(
            header_frame,
            text="URL Checker Pro",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        title_lbl.grid(row=0, column=0, padx=16, pady=10, sticky="w")

        subtitle_lbl = ctk.CTkLabel(
            header_frame,
            text="تتبع التحويلات (HTTP & System Edge/Chrome) • فحص متوازي • تصدير XLSX / CSV / TXT",
            font=ctk.CTkFont(size=12),
            text_color="#4B5563",
        )
        subtitle_lbl.grid(row=0, column=1, padx=8, pady=10, sticky="w")

        self.stats_lbl = ctk.CTkLabel(
            header_frame,
            text="الإجمالي: 0 | ناجح: 0 | أخطاء/متوقف: 0",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1D4ED8",
        )
        self.stats_lbl.grid(row=0, column=2, padx=16, pady=10, sticky="e")

        top_grid = ctk.CTkFrame(self, fg_color="transparent")
        top_grid.grid(row=1, column=0, padx=14, pady=6, sticky="ew")
        top_grid.grid_columnconfigure(0, weight=2)
        top_grid.grid_columnconfigure(1, weight=1)

        input_card = ctk.CTkFrame(top_grid, corner_radius=10)
        input_card.grid(row=0, column=0, padx=(0, 6), sticky="nsew")
        input_card.grid_columnconfigure(0, weight=1)
        input_card.grid_rowconfigure(1, weight=1)

        in_bar = ctk.CTkFrame(input_card, fg_color="transparent")
        in_bar.grid(row=0, column=0, padx=12, pady=(10, 4), sticky="ew")

        ctk.CTkLabel(in_bar, text="قائمة الروابط (رابط في كل سطر):", font=ctk.CTkFont(weight="bold")).pack(side="left")

        ctk.CTkButton(
            in_bar, text="تحميل ملف (.txt / .csv / .xlsx)", width=180, height=28, command=self.load_urls_file
        ).pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            in_bar,
            text="حذف المكرر",
            width=105,
            height=28,
            fg_color="#475569",
            hover_color="#334155",
            command=self.deduplicate_input,
        ).pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            in_bar,
            text="مسح",
            width=70,
            height=28,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            command=self.clear_input,
        ).pack(side="right")

        self.url_textbox = ctk.CTkTextbox(input_card, height=138, font=ctk.CTkFont(family="Consolas", size=12))
        self.url_textbox.grid(row=1, column=0, padx=12, pady=(0, 10), sticky="nsew")
        self.url_textbox.insert(
            "1.0",
            "https://eshop.wuerth.de/is-bin/INTERSHOP.enfinity/WFS/1401-B1-Site/en_US/-/EUR/ViewAfterSearch-ExecuteAfterSearch?SearchResultType=&EffectiveSearchTerm=&VisibleSearchTerm=04601012\n",
        )

        ctrl_card = ctk.CTkFrame(top_grid, corner_radius=10)
        ctrl_card.grid(row=0, column=1, padx=(6, 0), sticky="nsew")
        ctrl_card.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(ctrl_card, text="إعدادات الفحص والتحكم", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, columnspan=3, padx=12, pady=(10, 6), sticky="w"
        )

        ctk.CTkLabel(ctrl_card, text="وضع الفحص:").grid(row=1, column=0, padx=10, pady=3, sticky="w")
        self.mode_var = ctk.StringVar(value="Smart (HTTP + System Browser)")
        self.mode_menu = ctk.CTkOptionMenu(
            ctrl_card,
            variable=self.mode_var,
            values=[
                "Smart (HTTP + System Browser)",
                "Fast HTTP Only",
                "System Browser Only",
            ],
        )
        self.mode_menu.grid(row=1, column=1, columnspan=2, padx=10, pady=3, sticky="ew")

        ctk.CTkLabel(ctrl_card, text="المتصفح:").grid(row=2, column=0, padx=10, pady=3, sticky="w")
        self.channel_var = ctk.StringVar(value="Auto (Edge -> Chrome)")
        self.channel_menu = ctk.CTkOptionMenu(
            ctrl_card,
            variable=self.channel_var,
            values=[
                "Auto (Edge -> Chrome)",
                "Microsoft Edge (msedge)",
                "Google Chrome (chrome)",
            ],
        )
        self.channel_menu.grid(row=2, column=1, columnspan=2, padx=10, pady=3, sticky="ew")

        ctk.CTkLabel(ctrl_card, text="Timeout (ث):").grid(row=3, column=0, padx=10, pady=3, sticky="w")
        self.timeout_entry = ctk.CTkEntry(ctrl_card, width=70)
        self.timeout_entry.insert(0, "15")
        self.timeout_entry.grid(row=3, column=1, padx=4, pady=3, sticky="w")

        ctk.CTkLabel(ctrl_card, text="Threads:").grid(row=3, column=1, padx=(50, 0), pady=3, sticky="e")
        self.threads_entry = ctk.CTkEntry(ctrl_card, width=55)
        self.threads_entry.insert(0, "8")
        self.threads_entry.grid(row=3, column=2, padx=10, pady=3, sticky="e")

        ctk.CTkLabel(ctrl_card, text="إعادة محاولة:").grid(row=4, column=0, padx=10, pady=3, sticky="w")
        self.retries_entry = ctk.CTkEntry(ctrl_card, width=70)
        self.retries_entry.insert(0, "1")
        self.retries_entry.grid(row=4, column=1, padx=4, pady=3, sticky="w")

        self.headless_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(ctrl_card, text="Headless", variable=self.headless_var).grid(
            row=4, column=2, padx=10, pady=3, sticky="e"
        )

        btn_row1 = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        btn_row1.grid(row=5, column=0, columnspan=3, padx=10, pady=(6, 3), sticky="ew")
        btn_row1.grid_columnconfigure((0, 1, 2), weight=1)

        self.start_btn = ctk.CTkButton(
            btn_row1, text="▶ بدء الفحص", fg_color="#16A34A", hover_color="#15803D", command=self.start_check
        )
        self.start_btn.grid(row=0, column=0, padx=3, sticky="ew")

        self.pause_btn = ctk.CTkButton(
            btn_row1,
            text="⏸ إيقاف مؤقت",
            fg_color="#D97706",
            hover_color="#B45309",
            state="disabled",
            command=self.toggle_pause,
        )
        self.pause_btn.grid(row=0, column=1, padx=3, sticky="ew")

        self.stop_btn = ctk.CTkButton(
            btn_row1,
            text="⏹ إيقاف",
            fg_color="#DC2626",
            hover_color="#B91C1C",
            state="disabled",
            command=self.stop_check,
        )
        self.stop_btn.grid(row=0, column=2, padx=3, sticky="ew")

        btn_row2 = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        btn_row2.grid(row=6, column=0, columnspan=3, padx=10, pady=(2, 10), sticky="ew")
        btn_row2.grid_columnconfigure((0, 1), weight=1)

        self.retry_failed_btn = ctk.CTkButton(
            btn_row2,
            text="↻ إعادة فحص الفاشلة",
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            command=self.retry_failed,
        )
        self.retry_failed_btn.grid(row=0, column=0, padx=3, sticky="ew")

        self.clear_res_btn = ctk.CTkButton(
            btn_row2,
            text="تصفير النتائج",
            fg_color="#64748B",
            hover_color="#475569",
            command=self.clear_results,
        )
        self.clear_res_btn.grid(row=0, column=1, padx=3, sticky="ew")

        table_card = ctk.CTkFrame(self, corner_radius=10)
        table_card.grid(row=2, column=0, padx=14, pady=(6, 12), sticky="nsew")
        table_card.grid_columnconfigure(0, weight=1)
        table_card.grid_rowconfigure(1, weight=1)

        toolbar = ctk.CTkFrame(table_card, fg_color="transparent")
        toolbar.grid(row=0, column=0, padx=12, pady=(10, 6), sticky="ew")

        ctk.CTkLabel(toolbar, text="فلترة:").pack(side="left", padx=(0, 4))
        self.filter_var = ctk.StringVar(value="الكل")
        self.filter_menu = ctk.CTkOptionMenu(
            toolbar,
            variable=self.filter_var,
            values=["الكل", "ناجح (2xx/3xx)", "مُحوّل (Redirects)", "أخطاء / فاشل"],
            width=160,
            command=lambda _: self.refresh_table(),
        )
        self.filter_menu.pack(side="left", padx=(0, 8))

        self.search_entry = ctk.CTkEntry(toolbar, placeholder_text="بحث في الرابط أو العنوان...", width=240)
        self.search_entry.pack(side="left")
        self.search_entry.bind("<KeyRelease>", lambda _: self.refresh_table())

        ctk.CTkButton(
            toolbar,
            text="تصدير XLSX",
            width=110,
            fg_color="#0F766E",
            hover_color="#115E59",
            command=lambda: self.export_data("xlsx"),
        ).pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            toolbar,
            text="تصدير CSV",
            width=105,
            fg_color="#0369A1",
            hover_color="#075985",
            command=lambda: self.export_data("csv"),
        ).pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            toolbar,
            text="تصدير TXT (Tab)",
            width=120,
            fg_color="#374151",
            hover_color="#1F2937",
            command=lambda: self.export_data("txt"),
        ).pack(side="right")

        tree_frame = ctk.CTkFrame(table_card, fg_color="transparent")
        tree_frame.grid(row=1, column=0, padx=12, pady=(0, 6), sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        columns = (
            "idx",
            "original_url",
            "final_url",
            "status_code",
            "redirects",
            "title",
            "size",
            "ms",
            "mode",
            "error",
        )
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")

        headings = {
            "idx": "#",
            "original_url": "Original URL",
            "final_url": "Final Redirect URL",
            "status_code": "Status",
            "redirects": "Jumps",
            "title": "Page Title",
            "size": "Size (B)",
            "ms": "Time (ms)",
            "mode": "Engine",
            "error": "Error / Notes",
        }
        widths = {
            "idx": 45,
            "original_url": 240,
            "final_url": 290,
            "status_code": 75,
            "redirects": 60,
            "title": 200,
            "size": 85,
            "ms": 80,
            "mode": 95,
            "error": 180,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(
                col,
                width=widths[col],
                anchor="w" if col not in ("idx", "status_code", "redirects", "size", "ms") else "center",
            )

        self.tree.tag_configure("ok", background="#ECFDF5", foreground="#065F46")
        self.tree.tag_configure("redir", background="#EFF6FF", foreground="#1E40AF")
        self.tree.tag_configure("err", background="#FEF2F2", foreground="#991B1B")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        bottom_bar = ctk.CTkFrame(table_card, fg_color="transparent")
        bottom_bar.grid(row=2, column=0, padx=12, pady=(2, 10), sticky="ew")
        bottom_bar.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(bottom_bar)
        self.progress_bar.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.progress_bar.set(0)

        self.progress_lbl = ctk.CTkLabel(bottom_bar, text="0% (0 / 0)", width=130, anchor="e")
        self.progress_lbl.grid(row=0, column=1, sticky="e")

        footer_lbl = ctk.CTkLabel(
            self,
            text="Developed by ENG Ahmed Bahgat",
            font=ctk.CTkFont(size=11),
            text_color="#6B7280",
        )
        footer_lbl.grid(row=3, column=0, pady=(0, 6))

    def _apply_treeview_style(self):
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#E5E7EB", foreground="#111827")

    def _get_selected_channel(self) -> str:
        val = self.channel_var.get()
        if "msedge" in val:
            return "msedge"
        if "chrome" in val:
            return "chrome"
        return "auto"

    def load_urls_file(self):
        path = filedialog.askopenfilename(
            title="اختر ملف الروابط",
            filetypes=[
                ("Supported Files", "*.txt *.csv *.xlsx"),
                ("Text Files", "*.txt"),
                ("Excel Files", "*.xlsx"),
                ("CSV Files", "*.csv"),
            ],
        )
        if not path:
            return
        urls: list[str] = []
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".xlsx":
                wb = openpyxl.load_workbook(path, read_only=True)
                ws = wb.active
                for row in ws.iter_rows(values_only=True):
                    for cell in row:
                        if cell and str(cell).strip().startswith(("http://", "https://", "www.")):
                            urls.append(str(cell).strip())
                            break
                wb.close()
            elif ext == ".csv":
                with open(path, "r", encoding="utf-8", errors="ignore", newline="") as f:
                    reader = csv.reader(f)
                    for row in reader:
                        for cell in row:
                            if cell and str(cell).strip().startswith(("http://", "https://", "www.")):
                                urls.append(str(cell).strip())
                                break
            else:
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    for line in f:
                        if line.strip():
                            urls.append(line.strip())

            if urls:
                self.url_textbox.delete("1.0", "end")
                self.url_textbox.insert("1.0", "\n".join(urls) + "\n")
                messagebox.showinfo("تم التحميل", f"تم تحميل {len(urls)} رابط بنجاح.")
            else:
                messagebox.showwarning("تنبيه", "لم يتم العثور على روابط صالحة في الملف.")
        except Exception as exc:
            messagebox.showerror("خطأ", f"تعذر قراءة الملف:\n{exc}")

    def deduplicate_input(self):
        raw_lines = self.url_textbox.get("1.0", "end").splitlines()
        seen = set()
        unique = []
        for line in raw_lines:
            u = normalize_url(line)
            if u and u not in seen:
                seen.add(u)
                unique.append(u)
        self.url_textbox.delete("1.0", "end")
        if unique:
            self.url_textbox.insert("1.0", "\n".join(unique) + "\n")
        messagebox.showinfo("تمت التصفية", f"الروابط الفريدة: {len(unique)}")

    def clear_input(self):
        self.url_textbox.delete("1.0", "end")

    def clear_results(self):
        if self.is_running:
            return
        self.results.clear()
        self.refresh_table()
        self.progress_bar.set(0)
        self.progress_lbl.configure(text="0% (0 / 0)")
        self._update_stats()

    def _check_single_http(self, url: str, timeout_sec: float) -> dict:
        start_t = time.perf_counter()
        chain = [url]
        try:
            with requests.Session() as sess:
                sess.headers.update(DEFAULT_HEADERS)
                resp = sess.get(url, timeout=timeout_sec, allow_redirects=True)
                for h in resp.history:
                    if chain[-1] != h.url:
                        chain.append(h.url)
                    loc = h.headers.get("Location")
                    if loc:
                        target = urljoin(h.url, loc)
                        if chain[-1] != target:
                            chain.append(target)

                final_url = resp.url
                if chain[-1] != final_url:
                    chain.append(final_url)

                html_text = resp.text if "text/html" in resp.headers.get("Content-Type", "").lower() else ""
                title = extract_title_from_html(html_text)
                meta_target = extract_meta_or_canonical_redirect(html_text, final_url)

                elapsed_ms = int((time.perf_counter() - start_t) * 1000)
                return {
                    "original_url": url,
                    "final_url": final_url,
                    "redirect_count": max(0, len(chain) - 1),
                    "redirect_chain": " -> ".join(chain),
                    "status_code": str(resp.status_code),
                    "title": title,
                    "size_bytes": len(resp.content),
                    "response_ms": elapsed_ms,
                    "mode": "HTTP",
                    "error": "",
                    "status_bucket": "success" if 200 <= resp.status_code < 400 else "error",
                    "_needs_js": bool(meta_target and meta_target != final_url)
                    or ("eshop.wuerth" in final_url.lower() and not title),
                }
        except Exception as exc:
            elapsed_ms = int((time.perf_counter() - start_t) * 1000)
            return {
                "original_url": url,
                "final_url": chain[-1],
                "redirect_count": max(0, len(chain) - 1),
                "redirect_chain": " -> ".join(chain),
                "status_code": "ERR",
                "title": "",
                "size_bytes": 0,
                "response_ms": elapsed_ms,
                "mode": "HTTP",
                "error": str(exc).splitlines()[0][:180],
                "status_bucket": "error",
                "_needs_js": True,
            }

    def _check_with_retries(self, url: str, mode: str, timeout_sec: float, retries: int) -> dict:
        last_res = None
        for attempt in range(retries + 1):
            if self.stop_event.is_set():
                break
            self.pause_event.wait()

            if mode == "System Browser Only":
                res = self.pw_checker.check_url(url) if self.pw_checker else self._check_single_http(url, timeout_sec)
            elif mode == "Fast HTTP Only":
                res = self._check_single_http(url, timeout_sec)
            else:
                res = self._check_single_http(url, timeout_sec)
                if (
                    (res.get("_needs_js") or res["status_bucket"] == "error")
                    and self.pw_checker
                    and not self.stop_event.is_set()
                ):
                    pw_res = self.pw_checker.check_url(url)
                    if pw_res["status_bucket"] == "success" or pw_res["final_url"] != url:
                        res = pw_res

            last_res = res
            if res["status_bucket"] == "success":
                break
            if attempt < retries:
                time.sleep(0.5)

        if last_res is None:
            last_res = {
                "original_url": url,
                "final_url": url,
                "redirect_count": 0,
                "redirect_chain": url,
                "status_code": "STOPPED",
                "title": "",
                "size_bytes": 0,
                "response_ms": 0,
                "mode": mode,
                "error": "Stopped by user",
                "status_bucket": "error",
            }
        last_res.pop("_needs_js", None)
        return last_res

    def start_check(self, custom_urls: list[str] | None = None):
        if self.is_running:
            return

        if custom_urls is not None:
            urls = custom_urls
        else:
            raw_lines = self.url_textbox.get("1.0", "end").splitlines()
            urls = [normalize_url(x) for x in raw_lines if normalize_url(x)]

        if not urls:
            messagebox.showwarning("تنبيه", "الرجاء إدخال رابط واحد على الأقل.")
            return

        try:
            timeout_sec = max(1.0, float(self.timeout_entry.get().strip()))
            threads = max(1, min(50, int(self.threads_entry.get().strip())))
            retries = max(0, min(5, int(self.retries_entry.get().strip())))
        except ValueError:
            messagebox.showerror("خطأ في الإعدادات", "تأكد من إدخال أرقام صحيحة لـ Timeout و Threads و Retries.")
            return

        mode = self.mode_var.get()
        use_pw = mode in ("Smart (HTTP + System Browser)", "System Browser Only")
        channel = self._get_selected_channel()

        self.stop_event.clear()
        self.pause_event.set()
        self.is_running = True

        if custom_urls is None:
            self.results.clear()
            self.refresh_table()

        self.start_btn.configure(state="disabled")
        self.pause_btn.configure(state="normal", text="⏸ إيقاف مؤقت")
        self.stop_btn.configure(state="normal")
        self.retry_failed_btn.configure(state="disabled")

        threading.Thread(
            target=self._worker_runner,
            args=(urls, mode, timeout_sec, threads, retries, use_pw, channel),
            daemon=True,
        ).start()

    def _worker_runner(
        self, urls: list[str], mode: str, timeout_sec: float, threads: int, retries: int, use_pw: bool, channel: str
    ):
        total = len(urls)
        done = 0
        try:
            if use_pw:
                self.pw_checker = PlaywrightChecker(
                    timeout_sec=max(timeout_sec, 15.0),
                    headless=self.headless_var.get(),
                    browser_channel=channel,
                )
                try:
                    self.pw_checker.start()
                except Exception as exc:
                    self.after(
                        0,
                        lambda: messagebox.showwarning(
                            "System Browser", f"تعذر تشغيل متصفح النظام، سيتم الاعتماد على HTTP:\n{exc}"
                        ),
                    )
                    self.pw_checker = None

            effective_threads = 1 if (mode == "System Browser Only" and threads > 4) else threads
            with ThreadPoolExecutor(max_workers=effective_threads) as pool:
                futures = {pool.submit(self._check_with_retries, u, mode, timeout_sec, retries): u for u in urls}
                for fut in as_completed(futures):
                    if self.stop_event.is_set():
                        pool.shutdown(wait=False, cancel_futures=True)
                        break
                    self.pause_event.wait()
                    res = fut.result()
                    done += 1
                    self.after(0, lambda r=res, d=done, t=total: self._on_item_done(r, d, t))
        finally:
            if self.pw_checker is not None:
                self.pw_checker.close()
                self.pw_checker = None
            self.after(0, self._on_finished)

    def _on_item_done(self, item: dict, done: int, total: int):
        replaced = False
        for i, existing in enumerate(self.results):
            if existing["original_url"] == item["original_url"]:
                self.results[i] = item
                replaced = True
                break
        if not replaced:
            self.results.append(item)

        pct = done / max(1, total)
        self.progress_bar.set(pct)
        self.progress_lbl.configure(text=f"{pct*100:.1f}% ({done} / {total})")
        self._update_stats()
        self.refresh_table()

    def _on_finished(self):
        self.is_running = False
        self.start_btn.configure(state="normal")
        self.pause_btn.configure(state="disabled", text="⏸ إيقاف مؤقت")
        self.stop_btn.configure(state="disabled")
        self.retry_failed_btn.configure(state="normal")
        self._update_stats()

    def toggle_pause(self):
        if not self.is_running:
            return
        if self.pause_event.is_set():
            self.pause_event.clear()
            self.pause_btn.configure(text="▶ استئناف")
        else:
            self.pause_event.set()
            self.pause_btn.configure(text="⏸ إيقاف مؤقت")

    def stop_check(self):
        if not self.is_running:
            return
        self.stop_event.set()
        self.pause_event.set()

    def retry_failed(self):
        if self.is_running:
            return
        failed_urls = [r["original_url"] for r in self.results if r["status_bucket"] != "success"]
        if not failed_urls:
            messagebox.showinfo("لا توجد أخطاء", "لا توجد روابط فاشلة لإعادة فحصها.")
            return
        self.start_check(custom_urls=failed_urls)

    def _update_stats(self):
        total = len(self.results)
        ok_cnt = sum(1 for r in self.results if r["status_bucket"] == "success")
        err_cnt = total - ok_cnt
        self.stats_lbl.configure(text=f"الإجمالي: {total} | ناجح: {ok_cnt} | أخطاء/متوقف: {err_cnt}")

    def refresh_table(self):
        flt = self.filter_var.get()
        q = self.search_entry.get().strip().lower()

        for row_id in self.tree.get_children():
            self.tree.delete(row_id)

        idx = 0
        for r in self.results:
            if flt == "ناجح (2xx/3xx)" and r["status_bucket"] != "success":
                continue
            if flt == "مُحوّل (Redirects)" and r["redirect_count"] <= 0:
                continue
            if flt == "أخطاء / فاشل" and r["status_bucket"] == "success":
                continue

            haystack = f"{r['original_url']} {r['final_url']} {r['title']} {r['status_code']} {r['error']}".lower()
            if q and q not in haystack:
                continue

            idx += 1
            tag = "err" if r["status_bucket"] != "success" else ("redir" if r["redirect_count"] > 0 else "ok")
            self.tree.insert(
                "",
                "end",
                values=(
                    idx,
                    r["original_url"],
                    r["final_url"],
                    r["status_code"],
                    r["redirect_count"],
                    r["title"],
                    r["size_bytes"],
                    r["response_ms"],
                    r["mode"],
                    r["error"],
                ),
                tags=(tag,),
            )

    def export_data(self, fmt: str):
        if not self.results:
            messagebox.showwarning("تنبيه", "لا توجد نتائج لتصديرها.")
            return

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"URL_Check_Report_{stamp}.{fmt}"
        path = filedialog.asksaveasfilename(
            title="حفظ التقرير",
            initialfile=default_name,
            defaultextension=f".{fmt}",
            filetypes=[(f"{fmt.upper()} Files", f"*.{fmt}")],
        )
        if not path:
            return

        headers = [
            "Original URL",
            "Final URL",
            "Status Code",
            "Redirect Count",
            "Redirect Chain",
            "Page Title",
            "Size (Bytes)",
            "Response Time (ms)",
            "Engine",
            "Error",
        ]

        def row_vals(r: dict):
            return [
                r["original_url"],
                r["final_url"],
                r["status_code"],
                r["redirect_count"],
                r["redirect_chain"],
                r["title"],
                r["size_bytes"],
                r["response_ms"],
                r["mode"],
                r["error"],
            ]

        try:
            if fmt == "txt":
                with open(path, "w", encoding="utf-8") as f:
                    f.write("\t".join(headers) + "\n")
                    for r in self.results:
                        f.write("\t".join(str(x).replace("\t", " ") for x in row_vals(r)) + "\n")
            elif fmt == "csv":
                with open(path, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(headers)
                    for r in self.results:
                        writer.writerow(row_vals(r))
            elif fmt == "xlsx":
                wb = openpyxl.Workbook()
                ws_all = wb.active
                ws_all.title = "All Results"
                ws_err = wb.create_sheet("Errors Only")

                header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)

                for ws in (ws_all, ws_err):
                    ws.append(headers)
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font

                for r in self.results:
                    vals = row_vals(r)
                    ws_all.append(vals)
                    if r["status_bucket"] != "success":
                        ws_err.append(vals)

                wb.save(path)

            messagebox.showinfo("تم التصدير", f"تم حفظ الملف بنجاح:\n{path}")
        except Exception as exc:
            messagebox.showerror("خطأ في التصدير", f"تعذر حفظ الملف:\n{exc}")

    def _on_close(self):
        self.stop_event.set()
        self.pause_event.set()
        if self.pw_checker is not None:
            self.pw_checker.close()
        self.destroy()


if __name__ == "__main__":
    app = URLCheckerApp()
    app.mainloop()
