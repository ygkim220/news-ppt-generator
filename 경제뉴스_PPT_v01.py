import tkinter as tk
from tkinter import ttk, messagebox
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import re
from concurrent.futures import ThreadPoolExecutor
from email.utils import parsedate_to_datetime
import urllib.parse
import os
import threading
import webbrowser
import json
import glob
from datetime import date

# ==========================================
# Google News RSS 검색 설정
# ==========================================
GOOGLE_NEWS_RSS = "https://news.google.com/rss/search"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}


class TradingMonitorWindow:
    """외국인 / 기관 순매수·순매도 현황 모니터 (네이버 금융 스크래핑)"""

    _COLS = ("순위", "종목명", "현재가", "등락(원)", "거래량")

    def __init__(self, parent):
        self.win = tk.Toplevel(parent)
        self.win.title("외국인/기관 실시간 거래현황")
        self.win.geometry("1000x760")
        self.win.configure(bg="#F5F6F7")
        self.win.resizable(True, True)

        self.market_var = tk.StringVar(value="KOSPI")
        self.auto_var = tk.BooleanVar(value=False)
        self._job = None

        self._build_ui()
        self.refresh()

    def _build_ui(self):
        hdr = tk.Frame(self.win, bg="#1A252F", height=46)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="📊  외국인 / 기관  실시간 거래현황",
                 bg="#1A252F", fg="white",
                 font=("맑은 고딕", 12, "bold")).pack(side="left", padx=14, pady=10)
        tk.Label(hdr, text="※ 네이버 금융 기준  (장중 주기적 갱신)",
                 bg="#1A252F", fg="#85C1E9",
                 font=("맑은 고딕", 8)).pack(side="right", padx=14)

        ctrl = tk.Frame(self.win, bg="#ECF0F1", pady=5)
        ctrl.pack(fill="x")

        tk.Label(ctrl, text="시장:", bg="#ECF0F1",
                 font=("맑은 고딕", 9)).pack(side="left", padx=(12, 3))
        mkt_cb = ttk.Combobox(ctrl, textvariable=self.market_var,
                               values=["KOSPI", "KOSDAQ"], width=8, state="readonly")
        mkt_cb.pack(side="left")
        mkt_cb.bind("<<ComboboxSelected>>", lambda _: self.refresh())

        tk.Button(ctrl, text="🔄 새로고침", command=self.refresh,
                  bg="#3498DB", fg="white", relief="flat",
                  font=("맑은 고딕", 9, "bold"), padx=10, pady=3
                  ).pack(side="left", padx=10)

        tk.Checkbutton(ctrl, text="자동갱신 (30초)", variable=self.auto_var,
                       command=self._toggle_auto,
                       bg="#ECF0F1", font=("맑은 고딕", 9)).pack(side="left")

        self.lbl_status = tk.Label(ctrl, text="", bg="#ECF0F1",
                                    fg="#E74C3C", font=("맑은 고딕", 8))
        self.lbl_status.pack(side="right", padx=4)
        self.lbl_time = tk.Label(ctrl, text="", bg="#ECF0F1",
                                  fg="#7F8C8D", font=("맑은 고딕", 8))
        self.lbl_time.pack(side="right", padx=8)

        body = tk.Frame(self.win, bg="#F5F6F7")
        body.pack(fill="both", expand=True, padx=10, pady=8)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        # ── 순매수 행 ──
        for_buy_lf = tk.LabelFrame(body, text="  🔼  외국인 순매수 상위",
                                    font=("맑은 고딕", 10, "bold"), bg="white")
        for_buy_lf.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
        self.for_buy_tree = self._make_tree(for_buy_lf)

        inst_buy_lf = tk.LabelFrame(body, text="  🔼  기관 순매수 상위",
                                     font=("맑은 고딕", 10, "bold"), bg="white")
        inst_buy_lf.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))
        self.inst_buy_tree = self._make_tree(inst_buy_lf)

        # ── 순매도 행 ──
        for_sell_lf = tk.LabelFrame(body, text="  🔽  외국인 순매도 상위",
                                     font=("맑은 고딕", 10, "bold"), bg="white")
        for_sell_lf.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(4, 0))
        self.for_sell_tree = self._make_tree(for_sell_lf)

        inst_sell_lf = tk.LabelFrame(body, text="  🔽  기관 순매도 상위",
                                      font=("맑은 고딕", 10, "bold"), bg="white")
        inst_sell_lf.grid(row=1, column=1, sticky="nsew", padx=(4, 0), pady=(4, 0))
        self.inst_sell_tree = self._make_tree(inst_sell_lf)

    def _make_tree(self, parent):
        tree = ttk.Treeview(parent, columns=self._COLS, show="headings", height=8)
        for col, w in zip(self._COLS, [35, 145, 85, 80, 80]):
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor="center", minwidth=40)
        sb = tk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True, padx=4, pady=4)
        tree.tag_configure("up", foreground="#E74C3C")
        tree.tag_configure("dn", foreground="#3498DB")
        return tree

    def refresh(self):
        sosok = "0" if self.market_var.get() == "KOSPI" else "1"
        self.lbl_status.config(text="로딩 중...")
        threading.Thread(target=self._fetch_all, args=(sosok,), daemon=True).start()

    def _fetch_all(self, sosok):
        url = "https://finance.naver.com/sise/"
        if sosok == "1":
            url += "?sosok=1"
        try:
            base_hdrs = {**HEADERS, 'Accept-Language': 'ko-KR,ko;q=0.9'}
            r = requests.Session().get(url, headers=base_hdrs, timeout=10)
            r.raise_for_status()
            r.encoding = 'euc-kr'
            soup = BeautifulSoup(r.text, "html.parser")

            raw_fb = self._parse_table(soup, "frgn_deal_tab_0")
            raw_ib = self._parse_table(soup, "organ_deal_tab_0")
            raw_fs = self._parse_table(soup, "frgn_deal_tab_1")
            raw_is = self._parse_table(soup, "organ_deal_tab_1")

            seen, codes = set(), []
            for raw in [raw_fb, raw_ib, raw_fs, raw_is]:
                for row in raw:
                    c = row[4]
                    if c and c not in seen:
                        seen.add(c); codes.append(c)
            vol_map = self._get_volumes(codes)

            f_buy  = self._to_display(raw_fb, vol_map)
            i_buy  = self._to_display(raw_ib, vol_map)
            f_sell = self._to_display(raw_fs, vol_map)
            i_sell = self._to_display(raw_is, vol_map)
        except Exception as e:
            err = [("", f"오류: {str(e)[:40]}", "", "", "", "dn")]
            f_buy = i_buy = f_sell = i_sell = err

        self.win.after(0, lambda: self._populate(f_buy, i_buy, f_sell, i_sell))

    def _parse_table(self, soup, table_id):
        """(순위, 종목명, 현재가, 등락, 코드, tag) 반환"""
        table = soup.find("table", id=table_id)
        if not table:
            return []
        rows = []
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 3:
                continue
            name_a = tds[1].find("a") if len(tds) > 1 else None
            if not name_a:
                continue
            name = name_a.get_text(strip=True)
            price = tds[2].get_text(strip=True) if len(tds) > 2 else ""
            chg_raw = tds[3].get_text(strip=True) if len(tds) > 3 else ""
            chg = chg_raw.replace("상승", "+").replace("하락", "-").replace("보합", "±")
            # 순위: src="...ico_n02.gif" → "2"
            rank_img = tds[0].find("img") if tds else None
            rank = str(len(rows) + 1)
            if rank_img:
                m = re.search(r'ico_n(\d+)\.gif', rank_img.get("src", ""))
                if m:
                    rank = str(int(m.group(1)))
            # 종목코드
            code_m = re.search(r'code=(\d+)', name_a.get("href", ""))
            code = code_m.group(1) if code_m else ""
            tag = "dn" if chg_raw.startswith("하락") else "up"
            rows.append((rank, name, price, chg, code, tag))
        return rows

    def _get_volumes(self, codes):
        """종목별 polling API 병렬 호출로 거래량 수집"""
        if not codes:
            return {}

        def _fetch(code):
            try:
                url = (f"https://polling.finance.naver.com/api/realtime"
                       f"?query=SERVICE_ITEM:{code}")
                r = requests.get(url,
                                 headers={**HEADERS,
                                          "Referer": "https://finance.naver.com/"},
                                 timeout=4)
                items = (r.json().get("result", {})
                                  .get("areas", [{}])[0]
                                  .get("datas", []))
                return code, items[0].get("aq", "") if items else ""
            except Exception:
                return code, ""

        with ThreadPoolExecutor(max_workers=min(len(codes), 12)) as ex:
            return dict(ex.map(_fetch, codes))

    @staticmethod
    def _fmt_vol(v):
        if not v:
            return "-"
        try:
            n = int(v)
            if n >= 100_000_000:
                return f"{n / 100_000_000:.1f}억주"
            if n >= 10_000:
                return f"{n / 10_000:.1f}만주"
            return f"{n:,}주"
        except Exception:
            return str(v)

    def _to_display(self, raw_rows, vol_map):
        """(순위, 종목명, 현재가, 등락, 거래량, tag) → 화면 행 튜플"""
        if not raw_rows:
            return [("", "데이터 없음", "", "", "", "dn")]
        result = []
        for rank, name, price, chg, code, tag in raw_rows:
            vol = self._fmt_vol(vol_map.get(code, ""))
            result.append((rank, name, price, chg, vol, tag))
        return result

    def _populate(self, f_buy, i_buy, f_sell, i_sell):
        for tree, data in (
            (self.for_buy_tree,  f_buy),
            (self.inst_buy_tree, i_buy),
            (self.for_sell_tree, f_sell),
            (self.inst_sell_tree, i_sell),
        ):
            tree.delete(*tree.get_children())
            for row in data:
                *vals, tag = row
                tree.insert("", "end", values=vals, tags=(tag,))
        self.lbl_time.config(text=f"갱신: {datetime.now().strftime('%H:%M:%S')}")
        self.lbl_status.config(text="")

    def _toggle_auto(self):
        if self.auto_var.get():
            self._schedule()
        elif self._job:
            self.win.after_cancel(self._job)
            self._job = None

    def _schedule(self):
        if self.auto_var.get():
            self.refresh()
            self._job = self.win.after(30_000, self._schedule)


class NewsToPPTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Today's News")
        self.root.geometry("780x720")
        self.root.configure(bg="#F5F6F7")

        self.articles = []
        self.session_start_tokens = None
        self.setup_ui()
        self.update_token_monitor()

    # ------------------------------------------
    # UI 구성
    # ------------------------------------------
    def setup_ui(self):
        header = tk.Frame(self.root, bg="#2C3E50", height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="📰 Today's News",
                 bg="#2C3E50", fg="white",
                 font=("맑은 고딕", 14, "bold")).pack(expand=True)

        monitor = tk.Frame(self.root, bg="#1B2631", height=32)
        monitor.pack(fill="x")
        monitor.pack_propagate(False)
        tk.Label(monitor, text="🪙 Claude 토큰 (오늘):",
                 bg="#1B2631", fg="#F39C12",
                 font=("맑은 고딕", 9, "bold")).pack(side="left", padx=10)
        self.token_label = tk.Label(monitor,
                 text="로딩 중...",
                 bg="#1B2631", fg="#ECF0F1",
                 font=("Consolas", 9))
        self.token_label.pack(side="left")
        self.token_session_label = tk.Label(monitor,
                 text="",
                 bg="#1B2631", fg="#85C1E9",
                 font=("Consolas", 9))
        self.token_session_label.pack(side="right", padx=10)

        search_frame = tk.LabelFrame(self.root, text=" 검색 조건 ",
                                     font=("맑은 고딕", 10, "bold"),
                                     bg="white", padx=12, pady=10)
        search_frame.pack(fill="x", padx=15, pady=10)

        kw_frame = tk.Frame(search_frame, bg="white")
        kw_frame.pack(fill="x", pady=4)
        tk.Label(kw_frame, text="키워드:", font=("맑은 고딕", 10),
                 bg="white", width=10, anchor="w").pack(side="left")
        self.keyword_entry = tk.Entry(kw_frame, font=("맑은 고딕", 10), width=45)
        self.keyword_entry.pack(side="left", padx=5)
        self.keyword_entry.insert(0, "환율")
        self.keyword_entry.bind("<Return>", lambda e: self.start_search())

        cnt_frame = tk.Frame(search_frame, bg="white")
        cnt_frame.pack(fill="x", pady=4)
        tk.Label(cnt_frame, text="기사 수:", font=("맑은 고딕", 10),
                 bg="white", width=10, anchor="w").pack(side="left")
        self.count_var = tk.IntVar(value=10)
        ttk.Combobox(cnt_frame, textvariable=self.count_var,
                     values=[5, 10, 15, 20, 30], width=8,
                     state="readonly").pack(side="left", padx=5)

        sort_frame = tk.Frame(search_frame, bg="white")
        sort_frame.pack(fill="x", pady=4)
        tk.Label(sort_frame, text="정렬:", font=("맑은 고딕", 10),
                 bg="white", width=10, anchor="w").pack(side="left")
        self.sort_var = tk.StringVar(value="최신순")
        ttk.Combobox(sort_frame, textvariable=self.sort_var,
                     values=["최신순", "관련도순"], width=10,
                     state="readonly").pack(side="left", padx=5)

        btn_frame = tk.Frame(search_frame, bg="white")
        btn_frame.pack(fill="x", pady=8)
        self.search_btn = tk.Button(btn_frame, text="🔍  뉴스 검색",
                                    command=self.start_search,
                                    bg="#3498DB", fg="white",
                                    font=("맑은 고딕", 10, "bold"),
                                    relief="flat", padx=20, pady=6)
        self.search_btn.pack(side="left", padx=5)

        tk.Button(btn_frame, text="🏦  외국인/기관 거래",
                  command=self.open_trading_monitor,
                  bg="#8E44AD", fg="white",
                  font=("맑은 고딕", 10, "bold"),
                  relief="flat", padx=20, pady=6
                  ).pack(side="left", padx=5)

        self.status_label = tk.Label(self.root,
                                     text="키워드를 입력하고 검색하세요.",
                                     font=("맑은 고딕", 9), bg="#F5F6F7",
                                     fg="#555555")
        self.status_label.pack(pady=3)

        result_frame = tk.LabelFrame(self.root, text=" 검색 결과 ",
                                     font=("맑은 고딕", 10, "bold"),
                                     bg="white", padx=10, pady=5)
        result_frame.pack(fill="both", expand=True, padx=15, pady=10)

        self.result_text = tk.Text(result_frame, font=("맑은 고딕", 9),
                                   bg="#FDFEFE", wrap="word", padx=10, pady=8,
                                   relief="flat")
        scrollbar = tk.Scrollbar(result_frame, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.result_text.pack(side="left", fill="both", expand=True)

        self.result_text.tag_configure("title",
                                       font=("맑은 고딕", 10, "bold"),
                                       foreground="#2980B9")
        self.result_text.tag_configure("source", foreground="#7F8C8D",
                                       font=("맑은 고딕", 8))
        self.result_text.tag_configure("summary", foreground="#34495E")
        self.result_text.tag_configure("idx", foreground="#E67E22",
                                       font=("맑은 고딕", 10, "bold"))

    # ------------------------------------------
    # 검색 처리
    # ------------------------------------------
    def start_search(self):
        keyword = self.keyword_entry.get().strip()
        if not keyword:
            messagebox.showwarning("입력 오류", "검색 키워드를 입력하세요.")
            return

        self.search_btn.config(state="disabled")
        self.status_label.config(text=f"'{keyword}' 검색 중...")
        self.result_text.delete("1.0", tk.END)

        threading.Thread(target=self.do_search,
                         args=(keyword,), daemon=True).start()

    def do_search(self, keyword):
        try:
            count = self.count_var.get()
            self.articles = self.search_google_news(keyword, count)
            self.root.after(0, self.display_results)
        except Exception as e:
            err = str(e)
            self.root.after(0, lambda: self.show_error(err))

    def search_google_news(self, keyword, count):
        query = urllib.parse.quote(keyword)
        url = f"{GOOGLE_NEWS_RSS}?q={query}&hl=ko&gl=KR&ceid=KR:ko"
        r = requests.get(url, headers=HEADERS, timeout=10)
        r.raise_for_status()

        soup = BeautifulSoup(r.text, 'xml')
        items = soup.select('item')

        articles = []
        for item in items[:count]:
            title_tag = item.select_one('title')
            link_tag = item.select_one('link')
            pubdate_tag = item.select_one('pubDate')
            source_tag = item.select_one('source')
            desc_tag = item.select_one('description')

            title_raw = title_tag.get_text(strip=True) if title_tag else ''
            title, source_from_title = self._split_title_source(title_raw)
            source = (source_tag.get_text(strip=True)
                      if source_tag else source_from_title)

            pub_str = pubdate_tag.get_text(strip=True) if pubdate_tag else ''
            dt, date_display = self._parse_pubdate(pub_str)

            articles.append({
                'title': title,
                'link': link_tag.get_text(strip=True) if link_tag else '',
                'summary': self._clean_description(
                    desc_tag.get_text() if desc_tag else '', title),
                'source': source or '미상',
                'date': date_display,
                '_dt': dt,
            })

        if self.sort_var.get() == "최신순":
            articles.sort(key=lambda a: a['_dt'].timestamp() if a['_dt'] else 0,
                          reverse=True)
        return articles

    def _split_title_source(self, title_raw):
        if ' - ' in title_raw:
            parts = title_raw.rsplit(' - ', 1)
            return parts[0].strip(), parts[1].strip()
        return title_raw, ''

    def _parse_pubdate(self, pub_str):
        if not pub_str:
            return None, ''
        try:
            dt = parsedate_to_datetime(pub_str)
            return dt, dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            return None, pub_str

    def _clean_description(self, desc_html, title):
        if not desc_html:
            return ''
        inner = BeautifulSoup(desc_html, 'html.parser')
        texts = [t.strip() for t in inner.stripped_strings
                 if t.strip() and t.strip() != title]
        seen = set()
        unique = []
        for t in texts:
            if t in seen:
                continue
            seen.add(t)
            unique.append(t)
        return ' | '.join(unique[:5])

    def display_results(self):
        if not self.articles:
            self.status_label.config(text="검색 결과가 없습니다.")
            self.result_text.insert(tk.END,
                "\n   검색 결과가 없습니다.\n   다른 키워드로 시도해 보세요.")
            self.search_btn.config(state="normal")
            return

        self.status_label.config(
            text=f"✅ {len(self.articles)}개 기사 검색 완료.")

        for i, art in enumerate(self.articles, 1):
            link_tag = f"link_{i}"
            self.result_text.tag_configure(link_tag,
                                           foreground="#2980B9",
                                           underline=True,
                                           font=("맑은 고딕", 10, "bold"))
            url = art.get('link', '')
            if url:
                self.result_text.tag_bind(link_tag, "<Button-1>",
                    lambda e, u=url: webbrowser.open(u))
                self.result_text.tag_bind(link_tag, "<Enter>",
                    lambda e: self.result_text.config(cursor="hand2"))
                self.result_text.tag_bind(link_tag, "<Leave>",
                    lambda e: self.result_text.config(cursor=""))

            self.result_text.insert(tk.END, f"\n[{i:02d}] ", "idx")
            self.result_text.insert(tk.END, f"{art['title']}\n", link_tag)
            meta = f"      📰 {art['source']}"
            if art['date']:
                meta += f"   📅 {art['date']}"
            self.result_text.insert(tk.END, meta + "\n", "source")
            if art['summary']:
                self.result_text.insert(tk.END,
                    f"      {art['summary']}\n", "summary")

        self.search_btn.config(state="normal")

    def open_trading_monitor(self):
        TradingMonitorWindow(self.root)

    def show_error(self, msg):
        self.status_label.config(text="❌ 오류 발생")
        messagebox.showerror("오류", f"검색 실패:\n{msg}")
        self.search_btn.config(state="normal")

    # ------------------------------------------
    # Claude 토큰 사용량 모니터
    # ------------------------------------------
    def update_token_monitor(self):
        try:
            today_total, total_total = self._aggregate_claude_tokens()

            if self.session_start_tokens is None:
                self.session_start_tokens = total_total
            session_used = total_total - self.session_start_tokens

            self.token_label.config(
                text=f"입력 {self._fmt(today_total['input'])}   "
                     f"출력 {self._fmt(today_total['output'])}   "
                     f"캐시읽기 {self._fmt(today_total['cache_read'])}   "
                     f"캐시생성 {self._fmt(today_total['cache_creation'])}   "
                     f"총 {self._fmt(today_total['total'])}")
            self.token_session_label.config(
                text=f"세션 +{self._fmt(session_used)}")
        except Exception as e:
            self.token_label.config(text=f"(읽기 실패: {str(e)[:40]})")

        self.root.after(5000, self.update_token_monitor)

    def _aggregate_claude_tokens(self):
        today_str = date.today().isoformat()
        today = {'input': 0, 'output': 0,
                 'cache_read': 0, 'cache_creation': 0, 'total': 0}
        all_total = 0

        log_dir = os.path.expanduser("~/.claude/projects")
        pattern = os.path.join(log_dir, "*", "*.jsonl")

        for path in glob.glob(pattern):
            try:
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    for line in f:
                        line = line.strip()
                        if not line or '"usage"' not in line:
                            continue
                        try:
                            obj = json.loads(line)
                        except json.JSONDecodeError:
                            continue

                        usage = (obj.get('message', {}) or {}).get('usage')
                        if not isinstance(usage, dict):
                            continue

                        i = int(usage.get('input_tokens', 0) or 0)
                        o = int(usage.get('output_tokens', 0) or 0)
                        cr = int(usage.get('cache_read_input_tokens', 0) or 0)
                        cc = int(usage.get('cache_creation_input_tokens', 0) or 0)
                        sub = i + o + cr + cc
                        all_total += sub

                        ts = obj.get('timestamp', '')
                        if ts.startswith(today_str):
                            today['input'] += i
                            today['output'] += o
                            today['cache_read'] += cr
                            today['cache_creation'] += cc
                            today['total'] += sub
            except (IOError, OSError):
                continue

        return today, all_total

    @staticmethod
    def _fmt(n):
        if n >= 1_000_000:
            return f"{n / 1_000_000:.2f}M"
        if n >= 1_000:
            return f"{n / 1_000:.1f}K"
        return str(n)


if __name__ == "__main__":
    root = tk.Tk()
    app = NewsToPPTApp(root)
    root.mainloop()
