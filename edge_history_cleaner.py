#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Edge 站点历史清理器

按“名单”清理 Microsoft Edge 浏览记录中的指定站点，其它站点的记录原样保留。

名单规则（每行一条，空行和 # 开头的行忽略）：
    x.com                  匹配 x.com 及其所有子域（含子域开关开启时）
    www.x.com              只针对该主机（含子域开关开启时连同其子域）
    https://x.com/home     取网址中的主机名，规则同上
    *.ads.example.com      强制包含子域

清理内容：
    History                  网址、访问、搜索词、访问分段、来源字段、下载记录
    Top Sites                新标签页/地址栏的常用站点
    Shortcuts                地址栏输入历史提示
    Network Action Predictor 地址栏输入预测、资源预取记录
    Favicons                 这些站点的图标映射与图标本体

用法：
    双击“启动清理器.bat”，或  python edge_history_cleaner.py
    命令行：python edge_history_cleaner.py --site x.com --site t.co --dry-run
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
import queue
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext

    TK_OK = True
except Exception:  # pragma: no cover - 无 tkinter 时退回命令行
    TK_OK = False

APP_NAME = "Edge 站点历史清理器"
APP_VERSION = "1.0"

CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0

DB_FILES = ("History", "Top Sites", "Shortcuts", "Network Action Predictor", "Favicons")
SIDE_FILES = ("History-journal", "Preferences")


# --------------------------------------------------------------------------- #
# 基础工具
# --------------------------------------------------------------------------- #
def user_data_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "Microsoft" / "Edge" / "User Data"


def default_backup_dir() -> Path:
    docs = Path.home() / "Documents"
    return (docs if docs.exists() else Path.home()) / "Edge历史清理备份"


def human_size(num: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            return f"{num:.0f} {unit}" if unit == "B" else f"{num:.1f} {unit}"
        num /= 1024.0
    return f"{num:.1f} GB"


def enable_dpi_awareness() -> None:
    """让高分屏上的文字不发虚（Windows 10+ 按系统 DPI 渲染）。"""
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        return
    except Exception:
        pass
    try:
        import ctypes

        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _run(cmd, timeout: int = 40):
    try:
        return subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            errors="replace",
            timeout=timeout,
            creationflags=CREATE_NO_WINDOW if os.name == "nt" else 0,
        )
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# 站点规则
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class SiteRule:
    host: str
    subdomains: bool


@dataclass(frozen=True)
class TitleRule:
    kind: str                       # "contains" | "regex"
    raw: str                        # 用户写的原始行
    value: str                      # 用于匹配的字符串 / 正则表达式
    pattern: re.Pattern | None = None


class SiteMatcher:
    """把名单编译成匹配器：域名规则 + 网页标题规则。

    标题名单语法：
        R<正则>    网页标题匹配该正则（不区分大小写）
        =<文字>    强制按“标题包含该文字”处理（用于文字本身以 R 开头的情况）
        其它       网页标题包含该行文字
    """

    def __init__(self, text: str = "", include_subdomains: bool = True,
                 title_text: str = "", ignore_case: bool = True):
        self.include_subdomains = include_subdomains
        self.ignore_case = ignore_case
        self.rules: list[SiteRule] = []
        self.title_rules: list[TitleRule] = []
        self.problems: list[str] = []

        for lineno, raw in enumerate(text.splitlines(), 1):
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            rule = self._parse(line, include_subdomains, lineno)
            if rule:
                self.rules.append(rule)

        for lineno, raw in enumerate(title_text.splitlines(), 1):
            line = raw.rstrip()
            if not line.strip() or line.lstrip().startswith("#"):
                continue
            rule = self._parse_title(line.strip(), lineno)
            if rule:
                self.title_rules.append(rule)

    # -- 解析 ------------------------------------------------------------- #
    def _parse(self, entry: str, include_subdomains: bool, lineno: int) -> SiteRule | None:
        text = entry.strip().lower()
        force_sub = False
        if text.startswith("*."):
            force_sub = True
            text = text[2:]
        if not text:
            return None
        if "://" in text or "/" in text:
            host = urlsplit(text if "://" in text else "http://" + text).hostname or ""
        else:
            host = text.split(":")[0]
        host = host.strip(".").lower()
        if not host or any(ch in host for ch in " *?[]{}"):
            self.problems.append(f"第 {lineno} 行无法识别：{entry}")
            return None
        if not re.fullmatch(r"[0-9a-z._-]+", host):
            self.problems.append(f"第 {lineno} 行看起来不是域名：{entry}")
            return None
        return SiteRule(host, include_subdomains or force_sub)

    def _parse_title(self, entry: str, lineno: int) -> TitleRule | None:
        if entry == "R":
            self.problems.append(f"标题第 {lineno} 行只有 R，没有正则内容")
            return None
        if entry.startswith("R"):
            expr = entry[1:]
            try:
                pattern = re.compile(expr, re.IGNORECASE if self.ignore_case else 0)
            except re.error as exc:
                self.problems.append(f"标题第 {lineno} 行正则无效（{exc}）：{entry}")
                return None
            return TitleRule("regex", entry, expr, pattern)
        value = entry[1:] if entry.startswith("=") else entry
        if not value:
            self.problems.append(f"标题第 {lineno} 行内容为空")
            return None
        return TitleRule("contains", entry, value, None)

    # -- 匹配 ------------------------------------------------------------- #
    def has_rules(self) -> bool:
        return bool(self.rules or self.title_rules)

    def host_matches(self, host: str | None) -> bool:
        if not host:
            return False
        host = host.strip(".").lower()
        for rule in self.rules:
            if host == rule.host:
                return True
            if rule.subdomains and host.endswith("." + rule.host):
                return True
        return False

    def title_matches(self, title: str | None) -> bool:
        if not isinstance(title, str) or not title:
            return False
        for rule in self.title_rules:
            if rule.kind == "contains":
                if self.ignore_case:
                    if rule.value.lower() in title.lower():
                        return True
                elif rule.value in title:
                    return True
            elif rule.pattern is not None and rule.pattern.search(title):
                return True
        return False

    def url_matches(self, url: str | None) -> bool:
        if not isinstance(url, str) or not url:
            return False
        try:
            host = urlsplit(url).hostname
        except Exception:
            host = None
        if self.host_matches(host):
            return True
        # Edge 的图标服务：https://edge.microsoft.com/favicon/v1?url=<encode>
        if "favicon" in url and "url=" in url:
            try:
                inner = parse_qs(urlsplit(url).query).get("url", [""])[0]
                return self.host_matches(urlsplit(inner).hostname)
            except Exception:
                return False
        return False

    def page_matches(self, url: str | None, title: str | None) -> bool:
        return self.url_matches(url) or self.title_matches(title)

    def host_under(self, host: str | None, roots) -> bool:
        if not host or not roots:
            return False
        host = host.strip(".").lower()
        for root in roots:
            if host == root:
                return True
            if self.include_subdomains and host.endswith("." + root):
                return True
        return False

    # -- 派生 ------------------------------------------------------------- #
    def clone(self) -> "SiteMatcher":
        other = SiteMatcher(include_subdomains=self.include_subdomains,
                            ignore_case=self.ignore_case)
        other.rules = list(self.rules)
        other.title_rules = list(self.title_rules)
        other.problems = list(self.problems)
        return other

    def add_hosts(self, hosts) -> None:
        existing = {(r.host, r.subdomains) for r in self.rules}
        for host in hosts:
            host = (host or "").strip(".").lower()
            if not host:
                continue
            key = (host, self.include_subdomains)
            if key not in existing:
                self.rules.append(SiteRule(host, self.include_subdomains))
                existing.add(key)

    def without_titles(self) -> "SiteMatcher":
        other = self.clone()
        other.title_rules = []
        return other

    def describe(self) -> str:
        parts = []
        if self.rules:
            parts.append(self.describe_hosts())
        if self.title_rules:
            parts.append(f"标题：{self.describe_titles()}")
        return "  |  ".join(parts) if parts else "（空名单）"

    def describe_hosts(self) -> str:
        return "、".join(f"{r.host}{'（含子域）' if r.subdomains else ''}" for r in self.rules) or "（无）"

    def describe_titles(self) -> str:
        shown = "；".join(r.raw for r in self.title_rules[:6])
        more = "…" if len(self.title_rules) > 6 else ""
        return f"{shown}{more}" if shown else "（无）"


# --------------------------------------------------------------------------- #
# 配置文件与进程
# --------------------------------------------------------------------------- #
@dataclass
class Profile:
    name: str
    path: Path
    label: str = ""

    @property
    def display(self) -> str:
        return f"{self.name} — {self.label}" if self.label else self.name


def _profile_label(folder: Path) -> str:
    prefs = folder / "Preferences"
    if not prefs.exists():
        return ""
    try:
        data = json.loads(prefs.read_text(encoding="utf-8", errors="ignore"))
        return str(data.get("profile", {}).get("name") or "")
    except Exception:
        return ""


def discover_profiles(root: Path | None = None) -> list[Profile]:
    root = Path(root) if root else user_data_dir()
    out: list[Profile] = []
    if not root.exists():
        return out
    try:
        children = sorted(root.iterdir())
    except OSError:
        return out
    for folder in children:
        if not folder.is_dir():
            continue
        if folder.name != "Default" and not re.fullmatch(r"Profile \d+", folder.name):
            continue
        if not (folder / "History").exists() and not (folder / "Preferences").exists():
            continue
        out.append(Profile(folder.name, folder, _profile_label(folder)))
    return out


def edge_exe_path() -> Path | None:
    candidates = [
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
        / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
        / "Microsoft" / "Edge" / "Application" / "msedge.exe",
    ]
    for c in candidates:
        if c.exists():
            return c
    where = shutil.which("msedge")
    return Path(where) if where else None


def edge_pids() -> list[int]:
    result = _run(["tasklist", "/FI", "IMAGENAME eq msedge.exe", "/FO", "CSV", "/NH"], timeout=20)
    if not result or not result.stdout:
        return []
    pids: list[int] = []
    for line in result.stdout.splitlines():
        line = line.strip()
        if not line.startswith('"'):
            continue
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) > 1 and parts[0].lower() == "msedge.exe":
            try:
                pids.append(int(parts[1]))
            except ValueError:
                pass
    return pids


def close_edge(log, wait_graceful: float = 25.0) -> bool:
    if not edge_pids():
        return True
    log("· 正在关闭 Edge（会先请窗口正常退出）…")
    _run(["taskkill", "/IM", "msedge.exe"], timeout=20)
    deadline = time.time() + wait_graceful
    while time.time() < deadline:
        if not edge_pids():
            log("✔ Edge 已关闭")
            return True
        time.sleep(0.7)
    log("⚠ Edge 仍有后台进程，强制结束（未保存的页面内容可能丢失）")
    _run(["taskkill", "/F", "/IM", "msedge.exe"], timeout=20)
    deadline = time.time() + 12
    while time.time() < deadline:
        if not edge_pids():
            log("✔ Edge 已强制关闭")
            return True
        time.sleep(0.7)
    log("✗ 无法关闭 Edge，请手动退出后重试", "error")
    return False


def start_edge(log) -> None:
    exe = edge_exe_path()
    if not exe:
        log("⚠ 未找到 msedge.exe，请手动打开 Edge", "warn")
        return
    try:
        subprocess.Popen([str(exe)], creationflags=CREATE_NO_WINDOW if os.name == "nt" else 0)
        log("· 已重新启动 Edge（可用 Ctrl+Shift+T 恢复标签页）")
    except Exception as exc:
        log(f"⚠ 启动 Edge 失败：{exc}", "warn")


# --------------------------------------------------------------------------- #
# 数据库工具
# --------------------------------------------------------------------------- #
def open_db(path: Path) -> sqlite3.Connection:
    con = sqlite3.connect(str(path), timeout=30, isolation_level=None)
    con.text_factory = lambda b: b.decode("utf-8", "replace")
    try:
        con.execute("PRAGMA secure_delete=ON")
    except sqlite3.DatabaseError:
        pass
    return con


def table_exists(con: sqlite3.Connection, name: str) -> bool:
    try:
        row = con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)).fetchone()
        return bool(row)
    except sqlite3.DatabaseError:
        return False


def table_columns(con: sqlite3.Connection, name: str) -> set[str]:
    try:
        return {r[1] for r in con.execute(f"PRAGMA table_info({name})")}
    except sqlite3.DatabaseError:
        return set()


def _chunks(seq, size: int = 400):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def _ph(n: int) -> str:
    return ",".join("?" * n)


def _subset(con, table: str, column: str, ids, dry: bool, log) -> int:
    """统计或删除 table.column ∈ ids 的行，返回影响行数。"""
    if not ids:
        return 0
    if not table_exists(con, table) or column not in table_columns(con, table):
        log(f"  · 跳过 {table}（此 Edge 版本没有该表/字段）")
        return 0
    total = 0
    try:
        for chunk in _chunks(list(ids)):
            if dry:
                total += con.execute(
                    f"SELECT COUNT(*) FROM {table} WHERE {column} IN ({_ph(len(chunk))})", chunk
                ).fetchone()[0]
            else:
                cur = con.execute(
                    f"DELETE FROM {table} WHERE {column} IN ({_ph(len(chunk))})", chunk
                )
                total += max(cur.rowcount or 0, 0)
    except sqlite3.DatabaseError as exc:
        log(f"  · {table} 处理失败：{exc}", "warn")
    return total


# --------------------------------------------------------------------------- #
# History 清理
# --------------------------------------------------------------------------- #
VISIT_TABLES = (
    ("edge_visits", "visit_id"),
    ("content_annotations", "visit_id"),
    ("context_annotations", "visit_id"),
    ("visit_source", "id"),
)
URL_TABLES = (
    ("keyword_search_terms", "url_id"),
    ("edge_keyword_search_terms_source", "url_id"),
    ("edge_urls", "url_id"),
    ("segments", "url_id"),
)
DOWNLOAD_COLS = (
    "current_path", "target_path", "referrer", "site_url",
    "embedder_download_data", "tab_url", "tab_referrer_url",
)


def _purge_visited_links(con, matcher: SiteMatcher, dry: bool, log) -> int:
    if not table_exists(con, "visited_links"):
        return 0
    rows = con.execute("SELECT id, top_level_url, frame_url FROM visited_links").fetchall()
    hits = [r[0] for r in rows if matcher.url_matches(r[1]) or matcher.url_matches(r[2])]
    return _subset(con, "visited_links", "id", hits, dry, log)


def _purge_downloads(con, matcher: SiteMatcher, dry: bool, log) -> dict:
    out = {"downloads": 0, "downloads_slices": 0, "downloads_url_chains": 0}
    if not table_exists(con, "downloads"):
        return out
    cols = [c for c in DOWNLOAD_COLS if c in table_columns(con, "downloads")]
    if not cols:
        return out
    rows = con.execute("SELECT id," + ",".join(cols) + " FROM downloads").fetchall()
    hits = [r[0] for r in rows if any(matcher.url_matches(v) for v in r[1:] if isinstance(v, str))]
    out["downloads"] = _subset(con, "downloads", "id", hits, dry, log)
    out["downloads_slices"] = _subset(con, "downloads_slices", "download_id", hits, dry, log)
    out["downloads_url_chains"] = _subset(con, "downloads_url_chains", "id", hits, dry, log)
    if table_exists(con, "edge_downloads"):
        _subset(con, "edge_downloads", "id", hits, dry, log)
    return out


def _blank_referrers(con, matcher: SiteMatcher, dry: bool, log) -> int:
    if "external_referrer_url" not in table_columns(con, "visits"):
        return 0
    try:
        rows = con.execute(
            "SELECT id, external_referrer_url FROM visits "
            "WHERE external_referrer_url IS NOT NULL AND external_referrer_url <> ''"
        ).fetchall()
    except sqlite3.DatabaseError:
        return 0
    hits = [r[0] for r in rows if matcher.url_matches(r[1])]
    if not hits:
        return 0
    if not dry:
        for chunk in _chunks(hits):
            con.execute(f"UPDATE visits SET external_referrer_url=NULL WHERE id IN ({_ph(len(chunk))})", chunk)
    return len(hits)


def host_of(url: str | None) -> str:
    try:
        return (urlsplit(url or "").hostname or "").strip(".").lower()
    except Exception:
        return ""


def select_target_urls(con, matcher: SiteMatcher, log) -> tuple[list[int], dict]:
    """挑出要清理的 urls.id：域名命中，或标题命中后连带整站。"""
    info: dict = {"title_pages": 0, "title_roots": [], "title_samples": []}
    cols = table_columns(con, "urls")
    has_title = "title" in cols
    sql = "SELECT id, url, title FROM urls" if has_title else "SELECT id, url FROM urls"
    rows = con.execute(sql).fetchall()

    if not matcher.title_rules:
        return [r[0] for r in rows if matcher.url_matches(r[1])], info
    if not has_title:
        log("  · 这份 History 没有 title 字段，标题规则被跳过", "warn")
        return [r[0] for r in rows if matcher.url_matches(r[1])], info

    by_host: list[int] = []
    by_title: list[tuple[int, str, str]] = []
    for rid, url, title in rows:
        if matcher.url_matches(url):
            by_host.append(rid)
        elif matcher.title_matches(title):
            by_title.append((rid, url, title))

    info["title_pages"] = len(by_title)
    info["title_samples"] = [t for _, _, t in by_title[:5]]
    roots = sorted({host_of(u) for _, u, _ in by_title if host_of(u)})
    info["title_roots"] = roots
    if not roots:
        return by_host, info

    already = set(by_host)
    extra = [rid for rid, url, _ in rows
             if rid not in already and matcher.host_under(host_of(url), roots)]
    return by_host + extra, info


def purge_history(db_path: Path, matcher: SiteMatcher, log, dry: bool = False,
                  referrers: bool = True, vacuum: bool = True) -> dict:
    st: dict = {}
    con = open_db(db_path)
    try:
        url_ids, info = select_target_urls(con, matcher, log)
        st.update({k: v for k, v in info.items() if k != "title_samples"})
        st["title_hosts"] = len(info["title_roots"])
        st["_expanded_hosts"] = info["title_roots"]
        st["_title_samples"] = info["title_samples"]
        st["urls_matched"] = len(url_ids)
        if not url_ids:
            return st

        # 标题命中的站点，其 visited_links / 下载 / 来源字段 也要跟着清
        site_matcher = matcher.without_titles()
        site_matcher.add_hosts(info["title_roots"])

        visit_ids: list[int] = []
        seg_ids: set[int] = set()
        seg_col = "segment_id" in table_columns(con, "visits")
        for chunk in _chunks(url_ids):
            sql = f"SELECT id{' , segment_id' if seg_col else ''} FROM visits WHERE url IN ({_ph(len(chunk))})"
            for row in con.execute(sql, chunk):
                visit_ids.append(row[0])
                if seg_col and row[1]:
                    seg_ids.add(row[1])
        st["visits_matched"] = len(visit_ids)

        if not dry:
            con.execute("BEGIN IMMEDIATE")
            st["urls_deleted"] = _subset(con, "urls", "id", url_ids, False, log)
        else:
            st["urls_deleted"] = len(url_ids)

        for table, column in VISIT_TABLES:
            st[table] = _subset(con, table, column, visit_ids, dry, log)
        for table, column in URL_TABLES:
            st[table] = _subset(con, table, column, url_ids, dry, log)
        if seg_ids:
            st["segments"] = st.get("segments", 0) + _subset(con, "segments", "id", seg_ids, dry, log)
            st["segment_usage"] = _subset(con, "segment_usage", "segment_id", seg_ids, dry, log)
        st["visited_links"] = _subset(con, "visited_links", "link_url_id", url_ids, dry, log)
        st["visited_links"] += _purge_visited_links(con, site_matcher, dry, log)
        st.update(_purge_downloads(con, site_matcher, dry, log))

        st["visits_deleted"] = _subset(con, "visits", "url", url_ids, dry, log)
        if referrers:
            st["referrers"] = _blank_referrers(con, site_matcher, dry, log)

        if not dry:
            con.execute("COMMIT")
            if vacuum:
                con.execute("VACUUM")
            st["integrity"] = con.execute("PRAGMA integrity_check").fetchone()[0]
        return st
    except Exception:
        if con.in_transaction:
            con.rollback()
        raise
    finally:
        con.close()


# --------------------------------------------------------------------------- #
# 其它存储清理
# --------------------------------------------------------------------------- #
def clean_top_sites(db_path: Path, matcher: SiteMatcher, log, dry: bool) -> int:
    return _purge_url_table(db_path, "top_sites", "url", matcher, dry)


def clean_shortcuts(db_path: Path, matcher: SiteMatcher, log, dry: bool) -> int:
    removed = 0
    con = open_db(db_path)
    try:
        if not table_exists(con, "omni_box_shortcuts"):
            return 0
        cols = table_columns(con, "omni_box_shortcuts")
        url_cols = [c for c in ("url", "fill_into_edit") if c in cols]
        rows = con.execute("SELECT id," + ",".join(url_cols) + " FROM omni_box_shortcuts").fetchall()
        hits = [r[0] for r in rows if any(matcher.url_matches(v) for v in r[1:])]
        if not dry and hits:
            con.execute("BEGIN IMMEDIATE")
            for chunk in _chunks(hits):
                con.execute(f"DELETE FROM omni_box_shortcuts WHERE id IN ({_ph(len(chunk))})", chunk)
            con.execute("COMMIT")
        removed = len(hits)
    finally:
        con.close()
    return removed


def _purge_url_table(db_path: Path, table: str, column: str, matcher: SiteMatcher, dry: bool) -> int:
    con = open_db(db_path)
    removed = 0
    try:
        if not table_exists(con, table):
            return 0
        rows = con.execute(f"SELECT {column} FROM {table}").fetchall()
        hits = [r[0] for r in rows if matcher.url_matches(r[0])]
        if hits and not dry:
            con.execute("BEGIN IMMEDIATE")
            for value in hits:
                con.execute(f"DELETE FROM {table} WHERE {column}=?", (value,))
            con.execute("COMMIT")
        removed = len(hits)
    finally:
        con.close()
    return removed


def clean_predictor(db_path: Path, matcher: SiteMatcher, log, dry: bool) -> tuple[int, int]:
    nap = _purge_url_table(db_path, "network_action_predictor", "url", matcher, dry)
    prefetch = 0
    con = open_db(db_path)
    try:
        for table in ("resource_prefetch_predictor_origin", "resource_prefetch_predictor_host_redirect"):
            if not table_exists(con, table):
                continue
            keys = [r[0] for r in con.execute(f"SELECT key FROM {table}")]
            hits = [k for k in keys if isinstance(k, str) and matcher.host_matches(k)]
            if hits and not dry:
                con.execute("BEGIN IMMEDIATE")
                for key in hits:
                    con.execute(f"DELETE FROM {table} WHERE key=?", (key,))
                con.execute("COMMIT")
            prefetch += len(hits)
    finally:
        con.close()
    return nap, prefetch


def clean_favicons(db_path: Path, matcher: SiteMatcher, log, dry: bool) -> tuple[int, int]:
    con = open_db(db_path)
    mappings = icons = 0
    try:
        if not table_exists(con, "icon_mapping"):
            return 0, 0
        rows = con.execute("SELECT id, page_url, icon_id FROM icon_mapping").fetchall()
        hits = [r for r in rows if matcher.url_matches(r[1])]
        mappings = len(hits)

        icon_ids = {r[2] for r in hits}
        if table_exists(con, "favicons"):
            for fid, furl in con.execute("SELECT id, url FROM favicons"):
                if matcher.url_matches(furl):
                    icon_ids.add(fid)

        matched_per_icon: dict[int, int] = {}
        for r in hits:
            matched_per_icon[r[2]] = matched_per_icon.get(r[2], 0) + 1

        if not dry and hits:
            con.execute("BEGIN IMMEDIATE")
            for r in hits:
                con.execute("DELETE FROM icon_mapping WHERE id=?", (r[0],))
            con.execute("COMMIT")

        for icon_id in sorted(icon_ids):
            total = con.execute("SELECT COUNT(*) FROM icon_mapping WHERE icon_id=?", (icon_id,)).fetchone()[0]
            remaining = total - (0 if dry else 0) - matched_per_icon.get(icon_id, 0) * (1 if dry else 0)
            if not dry:
                remaining = con.execute(
                    "SELECT COUNT(*) FROM icon_mapping WHERE icon_id=?", (icon_id,)
                ).fetchone()[0]
            if remaining > 0:
                continue
            icons += 1
            if dry:
                continue
            con.execute("BEGIN IMMEDIATE")
            if table_exists(con, "favicon_bitmaps"):
                con.execute("DELETE FROM favicon_bitmaps WHERE icon_id=?", (icon_id,))
            con.execute("DELETE FROM favicons WHERE id=?", (icon_id,))
            con.execute("COMMIT")
    finally:
        con.close()
    return mappings, icons


def clean_side_stores(profile: Profile, matcher: SiteMatcher, log, dry: bool) -> dict:
    stats: dict[str, int] = {}
    top = profile.path / "Top Sites"
    if top.exists():
        stats["top_sites"] = clean_top_sites(top, matcher, log, dry)
    shortcuts = profile.path / "Shortcuts"
    if shortcuts.exists():
        stats["shortcuts"] = clean_shortcuts(shortcuts, matcher, log, dry)
    nap_db = profile.path / "Network Action Predictor"
    if nap_db.exists():
        nap, prefetch = clean_predictor(nap_db, matcher, log, dry)
        stats["predictor"] = nap
        stats["prefetch"] = prefetch
    favicons = profile.path / "Favicons"
    if favicons.exists():
        maps, icons = clean_favicons(favicons, matcher, log, dry)
        stats["favicon_maps"] = maps
        stats["favicons"] = icons
    return stats


# --------------------------------------------------------------------------- #
# 备份 / 恢复 / 运行
# --------------------------------------------------------------------------- #
def backup_profile(profile: Profile, dest: Path, log) -> list[str]:
    dest.mkdir(parents=True, exist_ok=True)
    copied = []
    for name in DB_FILES + SIDE_FILES:
        src = profile.path / name
        if src.exists():
            try:
                shutil.copy2(src, dest / name)
                copied.append(name)
            except OSError as exc:
                log(f"⚠ 备份 {name} 失败：{exc}", "warn")
    return copied


def restore_profile(profile: Profile, backup_profile_dir: Path, log) -> int:
    count = 0
    for name in DB_FILES + SIDE_FILES:
        src = backup_profile_dir / name
        target = profile.path / name
        if src.exists():
            shutil.copy2(src, target)
            count += 1
        else:
            for suffix in ("-journal", "-wal", "-shm"):
                stale = profile.path / (name + suffix)
                if stale.exists():
                    stale.unlink()
    return count


def _db_size(profile: Profile) -> int:
    total = 0
    for name in DB_FILES:
        path = profile.path / name
        if path.exists():
            total += path.stat().st_size
    return total


def _drop_stale_journal(profile: Profile, log) -> None:
    for name in DB_FILES:
        journal = profile.path / (name + "-journal")
        if journal.exists():
            try:
                journal.unlink()
                log(f"  · 已删除残留日志 {journal.name}")
            except OSError:
                pass


@dataclass
class Options:
    dry_run: bool = False
    backup: bool = True
    backup_dir: Path = field(default_factory=default_backup_dir)
    extra_stores: bool = True
    referrers: bool = True
    vacuum: bool = True
    auto_close: bool = True
    auto_restart: bool = True
    process_control: bool = True


def inspect_profile(profile: Profile, matcher: SiteMatcher, log) -> dict:
    """在不改动真实数据的前提下估算匹配量（复制一份库来读）。"""
    tmp = Path(tempfile.mkdtemp(prefix="ehc_preview_"))
    stats: dict = {}
    try:
        for name in DB_FILES:
            src = profile.path / name
            if src.exists():
                shutil.copy2(src, tmp / name)
        for suffix in ("-journal",):
            src = profile.path / ("History" + suffix)
            if src.exists():
                shutil.copy2(src, tmp / ("History" + suffix))
        history = tmp / "History"
        if history.exists():
            stats.update(purge_history(history, matcher, log, dry=True,
                                       referrers=options_default_referrers, vacuum=False))
        if options_extra_stores:
            copy = Profile(profile.name, tmp, profile.label)
            side = matcher.without_titles()
            side.add_hosts(stats.get("_expanded_hosts") or [])
            stats.update(clean_side_stores(copy, side, log, dry=True))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return stats


# 预览时使用的两个开关（由界面/命令行在调用前设置）
options_default_referrers = True
options_extra_stores = True


def clean_profile(profile: Profile, matcher: SiteMatcher, opts: Options, log) -> dict:
    stats: dict = {}
    history = profile.path / "History"
    if not history.exists():
        log(f"⚠ {profile.name} 没有 History 数据库，跳过", "warn")
        return stats
    size_before = _db_size(profile)

    stats.update(purge_history(history, matcher, log, dry=False,
                               referrers=opts.referrers, vacuum=opts.vacuum))
    if opts.extra_stores:
        side = matcher.without_titles()
        side.add_hosts(stats.get("_expanded_hosts") or [])
        stats.update(clean_side_stores(profile, side, log, dry=False))

    # 自查：确认目标站点/标题已无残留
    left = 0
    con = open_db(history)
    try:
        has_title = "title" in table_columns(con, "urls")
        sql = "SELECT url, title FROM urls" if has_title else "SELECT url FROM urls"
        for row in con.execute(sql):
            if matcher.page_matches(row[0], row[1] if has_title else None):
                left += 1
    finally:
        con.close()
    stats["leftover_urls"] = left
    stats["size_before"] = size_before
    stats["size_after"] = _db_size(profile)
    if left == 0 and (stats.get("urls_deleted") or 0) > 0:
        _drop_stale_journal(profile, log)
    return stats


def run_clean(profiles: list[Profile], matcher: SiteMatcher, opts: Options, log) -> dict:
    result: dict = {"profiles": {}, "backup_dir": None, "targets": matcher.describe_hosts(),
                    "title_rules": [r.raw for r in matcher.title_rules]}
    if not profiles:
        raise RuntimeError("没有选中任何配置文件")
    if not matcher.has_rules():
        raise RuntimeError("名单是空的，先填写要清理的站点或网页标题")

    if opts.process_control:
        if edge_pids():
            if not opts.auto_close:
                raise RuntimeError("Edge 正在运行，请先退出 Edge（或勾选“自动关闭 Edge”）")
            if not close_edge(log):
                raise RuntimeError("Edge 关闭失败，已中止（未改动任何数据）")

    if opts.backup:
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        run_dir = Path(opts.backup_dir) / stamp
        run_dir.mkdir(parents=True, exist_ok=True)
        for profile in profiles:
            copied = backup_profile(profile, run_dir / profile.name, log)
            log(f"  · 已备份 {profile.name}：{', '.join(copied) if copied else '（无文件）'}")
        manifest = {
            "app": APP_NAME,
            "version": APP_VERSION,
            "created": datetime.now().isoformat(timespec="seconds"),
            "targets": [r.host for r in matcher.rules],
            "title_rules": [r.raw for r in matcher.title_rules],
            "include_subdomains": any(r.subdomains for r in matcher.rules),
            "profiles": [p.name for p in profiles],
        }
        (run_dir / "manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        result["backup_dir"] = str(run_dir)
        log(f"✔ 备份完成：{run_dir}")

    for profile in profiles:
        log(f"\n▶ 处理配置文件 {profile.display}")
        stats = clean_profile(profile, matcher, opts, log)
        result["profiles"][profile.name] = stats
        title_note = ""
        if stats.get("title_pages"):
            title_note = f"（标题命中 {stats['title_pages']} 页 → 整站 {stats.get('title_hosts', 0)} 个）"
        log(f"  ✔ 匹配 URL {stats.get('urls_matched', 0)} 条 / 访问 {stats.get('visits_matched', 0)} 条"
            f"{title_note}，残留 {stats.get('leftover_urls', 0)} 条")

    if opts.process_control and opts.auto_restart:
        start_edge(log)
    return result


# --------------------------------------------------------------------------- #
# 报告
# --------------------------------------------------------------------------- #
def format_report(result: dict) -> str:
    lines = [f"目标站点：{result.get('targets', '')}"]
    if result.get("title_rules"):
        lines.append("标题规则：" + "；".join(result["title_rules"]))
    if result.get("backup_dir"):
        lines.append(f"备份目录：{result['backup_dir']}")
    for name, stats in result.get("profiles", {}).items():
        lines.append("")
        lines.append(f"[{name}]")
        pairs = [
            ("匹配 URL", stats.get("urls_matched")),
            ("删除 URL", stats.get("urls_deleted")),
            ("匹配访问", stats.get("visits_matched")),
            ("删除访问", stats.get("visits_deleted")),
            ("标题命中页", stats.get("title_pages")),
            ("标题整站数", stats.get("title_hosts")),
            ("互动记录", stats.get("edge_visits")),
            ("内容标注", stats.get("content_annotations")),
            ("上下文标注", stats.get("context_annotations")),
            ("搜索词", stats.get("keyword_search_terms")),
            ("内嵌链接", stats.get("visited_links")),
            ("访问分段", stats.get("segments")),
            ("来源字段", stats.get("referrers")),
            ("下载记录", stats.get("downloads")),
            ("常用站点", stats.get("top_sites")),
            ("地址栏提示", stats.get("shortcuts")),
            ("输入预测", stats.get("predictor")),
            ("预取记录", stats.get("prefetch")),
            ("图标映射", stats.get("favicon_maps")),
            ("图标本体", stats.get("favicons")),
            ("残留 URL", stats.get("leftover_urls")),
        ]
        for label, value in pairs:
            if value:
                lines.append(f"  {label:<8} {value}")
        if stats.get("size_before"):
            lines.append(f"  数据库   {human_size(stats['size_before'])} → {human_size(stats['size_after'])}")
        if stats.get("integrity"):
            lines.append(f"  完整性   {stats['integrity']}")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# 图形界面
# --------------------------------------------------------------------------- #
TARGET_SAMPLE = """# 每行一个站点，支持域名或完整网址；# 开头是注释
x.com
twitter.com
t.co
"""

TITLE_SAMPLE = r"""# 标题规则：普通行 = 标题包含该文字；R 开头 = 正则（如 R^广告|抽奖）
# 标题命中会清掉该页面所在的整个域名
"""


class CleanerApp(tk.Tk if TK_OK else object):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.minsize(900, 560)
        self._busy = False
        self._q: queue.Queue = queue.Queue()
        self._config_path = Path(os.environ.get("APPDATA", str(Path.home()))) / "EdgeHistoryCleaner" / "config.json"
        self.cfg = self._load_config()
        self._build_style()
        self._build_ui()
        self._load_settings()
        self._fit_window()
        self.after(120, self._drain)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _fit_window(self) -> None:
        """按内容实际需要的高度铺开窗口并居中，避免小屏上被裁掉。"""
        self.update_idletasks()
        need_w = max(self.winfo_reqwidth(), 900)
        need_h = self.winfo_reqheight()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        width = min(need_w, sw - 60)
        height = min(need_h + 6, sh - 80)
        x = max(0, (sw - width) // 2)
        y = max(0, (sh - height) // 2 - 20)
        self.geometry(f"{width}x{height}+{x}+{y}")

    # -- 外观 ------------------------------------------------------------- #
    def _build_style(self):
        try:
            dpi = self.winfo_fpixels("1i")
            self.tk.call("tk", "scaling", max(dpi / 72.0, 1.0))
        except Exception:
            pass
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:
            pass
        font = ("Microsoft YaHei UI", 10)
        style.configure(".", font=font)
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 17, "bold"))
        style.configure("Sub.TLabel", foreground="#5b6470", font=("Microsoft YaHei UI", 9))
        style.configure("Hint.TLabel", foreground="#7a838f", font=("Microsoft YaHei UI", 9))
        style.configure("Go.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(14, 7))
        style.configure("TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"))

    # -- 布局 ------------------------------------------------------------- #
    def _build_ui(self):
        outer = ttk.Frame(self, padding=(16, 12, 16, 12))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)

        head = ttk.Frame(outer)
        head.grid(row=0, column=0, sticky="ew")
        ttk.Label(head, text="Edge 站点历史清理器", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            head,
            text="按名单删除 Edge 浏览记录里指定站点的痕迹，其它站点不受影响；清理前自动备份。",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        # ① 名单
        box1 = ttk.LabelFrame(outer, text="① 站点名单", padding=(10, 8, 10, 10))
        box1.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        box1.columnconfigure(0, weight=1)

        self.txt_targets = tk.Text(box1, height=7, wrap="none", undo=True,
                                   font=("Consolas", 10), relief="solid", borderwidth=1,
                                   highlightthickness=1, padx=8, pady=6)
        self.txt_targets.grid(row=0, column=0, sticky="nsew")
        self.txt_targets.insert("1.0", TARGET_SAMPLE)
        sb = ttk.Scrollbar(box1, orient="vertical", command=self.txt_targets.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.txt_targets.configure(yscrollcommand=sb.set)

        side = ttk.Frame(box1)
        side.grid(row=0, column=2, sticky="ns", padx=(8, 0))
        ttk.Button(side, text="导入名单…", command=self._import_list).pack(fill="x", pady=2)
        ttk.Button(side, text="导出名单…", command=self._export_list).pack(fill="x", pady=2)
        ttk.Button(side, text="清空", command=lambda: self.txt_targets.delete("1.0", "end")).pack(fill="x", pady=2)
        ttk.Label(side, text="用 # 写注释", style="Hint.TLabel").pack(anchor="w", pady=(6, 0))

        self.var_sub = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            box1, text="域名同时匹配其子域（填 x.com 会连 www.x.com、api.x.com 一起清）",
            variable=self.var_sub,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))

        # ② 网页标题名单
        box_titles = ttk.LabelFrame(outer, text="② 网页标题名单（可选）", padding=(10, 8, 10, 10))
        box_titles.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        box_titles.columnconfigure(0, weight=1)

        self.txt_titles = tk.Text(box_titles, height=4, wrap="none", undo=True,
                                  font=("Consolas", 10), relief="solid", borderwidth=1,
                                  highlightthickness=1, padx=8, pady=6)
        self.txt_titles.grid(row=0, column=0, sticky="nsew")
        self.txt_titles.insert("1.0", TITLE_SAMPLE)
        sb2 = ttk.Scrollbar(box_titles, orient="vertical", command=self.txt_titles.yview)
        sb2.grid(row=0, column=1, sticky="ns")
        self.txt_titles.configure(yscrollcommand=sb2.set)

        side2 = ttk.Frame(box_titles)
        side2.grid(row=0, column=2, sticky="ns", padx=(8, 0))
        ttk.Button(side2, text="导入…", command=self._import_titles).pack(fill="x", pady=2)
        ttk.Button(side2, text="导出…", command=self._export_titles).pack(fill="x", pady=2)
        ttk.Button(side2, text="清空", command=lambda: self.txt_titles.delete("1.0", "end")).pack(fill="x", pady=2)

        ttk.Label(
            box_titles,
            text="普通一行 = 标题包含该文字；R 开头 = 后面是正则（如 R^广告｜R\\d{4}），= 开头 = 强制按包含处理。"
                 "标题命中会清掉该页面所在整个域名。",
            style="Hint.TLabel", wraplength=880, justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.var_title_case = tk.BooleanVar(value=False)
        ttk.Checkbutton(box_titles, text="区分大小写", variable=self.var_title_case).grid(
            row=1, column=2, sticky="ne", pady=(6, 0))

        # ③④ 配置 + 选项
        mid = ttk.Frame(outer)
        mid.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
        mid.columnconfigure(0, weight=1)
        mid.columnconfigure(1, weight=1)
        outer.rowconfigure(3, weight=0)

        box_prof = ttk.LabelFrame(mid, text="③ Edge 配置文件", padding=(10, 8))
        box_prof.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.lst_profiles = tk.Listbox(box_prof, selectmode="extended", height=5, activestyle="none",
                                       font=("Microsoft YaHei UI", 9), relief="solid", borderwidth=1,
                                       highlightthickness=0)
        self.lst_profiles.pack(fill="both", expand=True)
        row = ttk.Frame(box_prof)
        row.pack(fill="x", pady=(6, 0))
        ttk.Button(row, text="重新扫描", command=self._refresh_profiles).pack(side="left")
        ttk.Label(row, text="Ctrl/Shift 可多选", style="Hint.TLabel").pack(side="left", padx=8)

        box_opts = ttk.LabelFrame(mid, text="④ 清理选项", padding=(10, 8))
        box_opts.grid(row=0, column=1, sticky="nsew")
        self.var_stores = tk.BooleanVar(value=True)
        self.var_ref = tk.BooleanVar(value=True)
        self.var_backup = tk.BooleanVar(value=True)
        self.var_vacuum = tk.BooleanVar(value=True)
        self.var_autoclose = tk.BooleanVar(value=True)
        self.var_restart = tk.BooleanVar(value=True)

        ttk.Checkbutton(box_opts, text="同时清理常用站点 / 地址栏提示 / 输入预测 / 图标",
                        variable=self.var_stores).pack(anchor="w")
        ttk.Checkbutton(box_opts, text="清空指向这些站点的“来源”字段",
                        variable=self.var_ref).pack(anchor="w")
        ttk.Checkbutton(box_opts, text="清理前自动备份（强烈建议）",
                        variable=self.var_backup).pack(anchor="w")
        ttk.Checkbutton(box_opts, text="收缩数据库，抹掉已删内容的残留字节",
                        variable=self.var_vacuum).pack(anchor="w")
        ttk.Checkbutton(box_opts, text="自动关闭 Edge", variable=self.var_autoclose).pack(anchor="w", pady=(6, 0))
        ttk.Checkbutton(box_opts, text="清理完成后重新打开 Edge", variable=self.var_restart).pack(anchor="w")

        brow = ttk.Frame(box_opts)
        brow.pack(fill="x", pady=(8, 0))
        ttk.Label(brow, text="备份到").pack(side="left")
        self.var_backup_dir = tk.StringVar(value=str(default_backup_dir()))
        ttk.Entry(brow, textvariable=self.var_backup_dir, width=28).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(brow, text="…", width=3, command=self._choose_backup_dir).pack(side="left")

        # 操作区
        actions = ttk.Frame(outer)
        actions.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        self.btn_preview = ttk.Button(actions, text="预览匹配量", command=self._on_preview)
        self.btn_preview.pack(side="left")
        self.btn_run = ttk.Button(actions, text="开始清理", style="Go.TButton", command=self._on_clean)
        self.btn_run.pack(side="left", padx=8)
        self.btn_restore = ttk.Button(actions, text="从备份恢复…", command=self._on_restore)
        self.btn_restore.pack(side="left")
        ttk.Button(actions, text="打开备份目录",
                   command=lambda: self._open_path(self.var_backup_dir.get())).pack(side="left", padx=8)
        self.lbl_status = ttk.Label(actions, text="就绪", style="Hint.TLabel")
        self.lbl_status.pack(side="right")

        # 日志
        box_log = ttk.LabelFrame(outer, text="运行日志", padding=(10, 8))
        box_log.grid(row=5, column=0, sticky="nsew", pady=(10, 0))
        outer.rowconfigure(5, weight=1, minsize=110)
        self.log_view = scrolledtext.ScrolledText(box_log, height=9, wrap="word", state="disabled",
                                                  font=("Consolas", 9), relief="solid", borderwidth=1,
                                                  background="#fbfcfd")
        self.log_view.pack(fill="both", expand=True)
        self.log_view.tag_configure("warn", foreground="#a15c00")
        self.log_view.tag_configure("error", foreground="#b3261e")
        self.log_view.tag_configure("ok", foreground="#0f6b3f")

    # -- 配置读写 --------------------------------------------------------- #
    def _load_config(self) -> dict:
        try:
            return json.loads(self._config_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_config(self) -> None:
        data = {
            "targets": self.txt_targets.get("1.0", "end-1c"),
            "titles": self.txt_titles.get("1.0", "end-1c"),
            "title_case": bool(self.var_title_case.get()),
            "subdomains": bool(self.var_sub.get()),
            "stores": bool(self.var_stores.get()),
            "referrers": bool(self.var_ref.get()),
            "backup": bool(self.var_backup.get()),
            "vacuum": bool(self.var_vacuum.get()),
            "autoclose": bool(self.var_autoclose.get()),
            "restart": bool(self.var_restart.get()),
            "backup_dir": self.var_backup_dir.get(),
        }
        try:
            self._config_path.parent.mkdir(parents=True, exist_ok=True)
            self._config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _load_settings(self) -> None:
        cfg = self.cfg
        if cfg.get("targets"):
            self.txt_targets.delete("1.0", "end")
            self.txt_targets.insert("1.0", cfg["targets"])
        if cfg.get("titles"):
            self.txt_titles.delete("1.0", "end")
            self.txt_titles.insert("1.0", cfg["titles"])
        for key, var in (
            ("subdomains", self.var_sub), ("stores", self.var_stores), ("referrers", self.var_ref),
            ("backup", self.var_backup), ("vacuum", self.var_vacuum),
            ("autoclose", self.var_autoclose), ("restart", self.var_restart),
            ("title_case", self.var_title_case),
        ):
            if key in cfg:
                var.set(bool(cfg[key]))
        if cfg.get("backup_dir"):
            self.var_backup_dir.set(cfg["backup_dir"])
        self._refresh_profiles()

    def _on_close(self) -> None:
        self._save_config()
        self.destroy()

    # -- 配置文件列表 ----------------------------------------------------- #
    def _refresh_profiles(self, preselect_all: bool = True) -> None:
        self.profiles = discover_profiles()
        self.lst_profiles.delete(0, "end")
        for p in self.profiles:
            self.lst_profiles.insert("end", p.display)
        if self.profiles and preselect_all:
            self.lst_profiles.selection_set(0, "end")
        if not self.profiles:
            self._log("⚠ 未找到 Edge 配置文件，请确认 Edge 已安装并至少运行过一次", "warn")

    def _selected_profiles(self) -> list[Profile]:
        return [self.profiles[i] for i in self.lst_profiles.curselection()]

    # -- 名单导入导出 ----------------------------------------------------- #
    def _import_list(self) -> None:
        self._import_text(self.txt_targets, "选择站点名单文件", "站点名单")

    def _export_list(self) -> None:
        self._export_text(self.txt_targets, "导出站点名单", "站点名单.txt", "站点名单")

    def _import_titles(self) -> None:
        self._import_text(self.txt_titles, "选择标题名单文件", "标题名单")

    def _export_titles(self) -> None:
        self._export_text(self.txt_titles, "导出标题名单", "标题名单.txt", "标题名单")

    def _import_text(self, widget: tk.Text, title: str, label: str) -> None:
        path = filedialog.askopenfilename(title=title,
                                          filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            text = Path(path).read_text(encoding="utf-8-sig", errors="replace")
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"读取失败：{exc}")
            return
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        self._log(f"· 已导入{label}：{path}")

    def _export_text(self, widget: tk.Text, title: str, filename: str, label: str) -> None:
        path = filedialog.asksaveasfilename(title=title, defaultextension=".txt",
                                            initialfile=filename,
                                            filetypes=[("文本文件", "*.txt")])
        if not path:
            return
        try:
            Path(path).write_text(widget.get("1.0", "end-1c"), encoding="utf-8")
            self._log(f"· 已导出{label}：{path}")
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"写入失败：{exc}")

    def _choose_backup_dir(self) -> None:
        path = filedialog.askdirectory(title="选择备份目录", initialdir=self.var_backup_dir.get())
        if path:
            self.var_backup_dir.set(path)

    def _open_path(self, path: str) -> None:
        try:
            Path(path).mkdir(parents=True, exist_ok=True)
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"无法打开：{exc}")

    # -- 日志与线程 ------------------------------------------------------- #
    def _log(self, message: str, level: str = "info") -> None:
        self._q.put(("log", (message, level)))

    def _drain(self) -> None:
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "log":
                    text, level = payload
                    text = text.rstrip("\n")
                    leading = ""
                    while text.startswith("\n"):
                        leading += "\n"
                        text = text[1:]
                    self.log_view.configure(state="normal")
                    if leading:
                        self.log_view.insert("end", leading)
                    stamp = datetime.now().strftime("%H:%M:%S")
                    self.log_view.insert("end", f"[{stamp}] {text}\n", level if level != "info" else "")
                    self.log_view.see("end")
                    self.log_view.configure(state="disabled")
                elif kind == "done":
                    self._set_busy(False)
                elif kind == "status":
                    self.lbl_status.configure(text=payload)
                elif kind == "report":
                    self._show_report(payload)
                elif kind == "save_config":
                    self._save_config()
        except queue.Empty:
            pass
        self.after(120, self._drain)

    def _show_report(self, report: str) -> None:
        box = tk.Toplevel(self)
        box.title("清理完成")
        box.geometry("560x440")
        box.transient(self)
        txt = scrolledtext.ScrolledText(box, wrap="word", font=("Consolas", 9))
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", report)
        txt.configure(state="disabled")
        ttk.Button(box, text="关闭", command=box.destroy).pack(pady=(0, 10))
        box.grab_set()

    def _set_busy(self, busy: bool, status: str | None = None) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        for btn in (self.btn_preview, self.btn_run, self.btn_restore):
            btn.configure(state=state)
        if status is not None:
            self.lbl_status.configure(text=status)

    def _run_async(self, task, label: str) -> None:
        if self._busy:
            return
        self._set_busy(True, label)

        def runner():
            try:
                task()
            except Exception as exc:  # noqa: BLE001
                self._log(f"✗ 失败：{exc}", "error")
                self._q.put(("status", "出错"))
            finally:
                self._q.put(("done", None))

        threading.Thread(target=runner, daemon=True).start()

    # -- 组装参数 --------------------------------------------------------- #
    def _collect(self, dry: bool) -> tuple[list[Profile], SiteMatcher, Options] | None:
        global options_default_referrers, options_extra_stores
        text = self.txt_targets.get("1.0", "end-1c")
        title_text = self.txt_titles.get("1.0", "end-1c")
        matcher = SiteMatcher(
            text,
            include_subdomains=bool(self.var_sub.get()),
            title_text=title_text,
            ignore_case=not bool(self.var_title_case.get()),
        )
        for problem in matcher.problems:
            self._log(f"⚠ {problem}", "warn")
        if not matcher.has_rules():
            messagebox.showwarning(APP_NAME, "名单是空的：至少填写一个域名，或一条网页标题规则")
            return None
        profiles = self._selected_profiles()
        if not profiles:
            messagebox.showwarning(APP_NAME, "请先在 ③ 里选中至少一个 Edge 配置文件")
            return None
        opts = Options(
            dry_run=dry,
            backup=bool(self.var_backup.get()),
            backup_dir=Path(self.var_backup_dir.get() or default_backup_dir()),
            extra_stores=bool(self.var_stores.get()),
            referrers=bool(self.var_ref.get()),
            vacuum=bool(self.var_vacuum.get()),
            auto_close=bool(self.var_autoclose.get()),
            auto_restart=bool(self.var_restart.get()),
            process_control=(not dry),
        )
        options_default_referrers = opts.referrers
        options_extra_stores = opts.extra_stores
        return profiles, matcher, opts

    # -- 预览 ------------------------------------------------------------- #
    def _on_preview(self) -> None:
        picked = self._collect(dry=True)
        if not picked:
            return
        profiles, matcher, _opts = picked
        self._log(f"\n=== 预览：{matcher.describe()}（不修改任何数据）===")

        def task():
            for profile in profiles:
                self._log(f"· 扫描 {profile.display}")
                stats = inspect_profile(profile, matcher, self._log)
                if not stats.get("urls_matched"):
                    self._log("  · 没有匹配到记录")
                    continue
                self._log(
                    f"  ✔ 匹配 {stats['urls_matched']} 条 URL / {stats.get('visits_matched', 0)} 次访问"
                    f"；常用站点 {stats.get('top_sites', 0)}，地址栏提示 {stats.get('shortcuts', 0)}，"
                    f"输入预测 {stats.get('predictor', 0)}，图标 {stats.get('favicon_maps', 0)}"
                )
                if stats.get("title_pages"):
                    samples = [str(s).replace("\n", " ")[:38] for s in (stats.get("_title_samples") or [])]
                    roots = stats.get("_expanded_hosts") or []
                    self._log(f"  · 标题命中 {stats['title_pages']} 个页面，连带整站："
                              + "、".join(roots[:6]) + ("…" if len(roots) > 6 else ""))
                    if samples:
                        self._log("  · 命中标题示例：" + " ｜ ".join(samples))
            self._q.put(("status", "预览完成"))

        self._run_async(task, "预览中…")

    # -- 清理 ------------------------------------------------------------- #
    def _on_clean(self) -> None:
        picked = self._collect(dry=False)
        if not picked:
            return
        profiles, matcher, opts = picked
        names = "、".join(p.display for p in profiles)
        extra = "备份到\n" + str(opts.backup_dir) if opts.backup else "不备份（不建议）"
        title_note = f"\n标题规则 {len(matcher.title_rules)} 条（命中后连带清整站）" if matcher.title_rules else ""
        if not messagebox.askyesno(
            APP_NAME,
            f"将要清理：\n{matcher.describe()}{title_note}\n\n"
            f"目标配置文件：{names}\n\n"
            f"{extra}\n\n"
            "过程中 Edge 会被关闭，完成后自动重开。是否继续？",
        ):
            return

        self._log(f"\n=== 开始清理：{matcher.describe()} ===")

        def task():
            result = run_clean(profiles, matcher, opts, self._log)
            report = format_report(result)
            self._log("\n=== 结果 ===\n" + report)
            self._q.put(("report", report))
            self._q.put(("save_config", None))
            self._q.put(("status", "清理完成"))

        self._run_async(task, "清理中…")

    # -- 恢复 ------------------------------------------------------------- #
    def _on_restore(self) -> None:
        root = Path(self.var_backup_dir.get() or default_backup_dir())
        if not root.exists():
            messagebox.showinfo(APP_NAME, f"备份目录不存在：\n{root}")
            return
        backups = sorted([d for d in root.iterdir() if d.is_dir() and (d / "manifest.json").exists()],
                         reverse=True)
        if not backups:
            messagebox.showinfo(APP_NAME, "这个目录里没有找到可恢复的备份")
            return

        dlg = tk.Toplevel(self)
        dlg.title("从备份恢复")
        dlg.geometry("620x400")
        dlg.transient(self)
        ttk.Label(dlg, text="选择一次备份（最新的在最上面）：", padding=(12, 10, 12, 4)).pack(anchor="w")
        lst = tk.Listbox(dlg, font=("Microsoft YaHei UI", 9), relief="solid", borderwidth=1,
                         highlightthickness=0)
        lst.pack(fill="both", expand=True, padx=12)
        for folder in backups:
            try:
                manifest = json.loads((folder / "manifest.json").read_text(encoding="utf-8"))
                desc = "、".join(manifest.get("targets", [])) or "（未记录目标）"
                when = manifest.get("created", folder.name)
            except Exception:
                desc, when = "（清单损坏）", folder.name
            lst.insert("end", f"{folder.name}   {when.replace('T', ' ')}   {desc}")

        def do_restore():
            sel = lst.curselection()
            if not sel:
                messagebox.showinfo(APP_NAME, "请先选中一条备份", parent=dlg)
                return
            folder = backups[sel[0]]
            if edge_pids() and not messagebox.askyesno(
                APP_NAME, "恢复需要关闭 Edge，现在关闭吗？", parent=dlg
            ):
                return
            dlg.destroy()
            self._log(f"\n=== 从备份恢复：{folder} ===")
            profiles = list(self.profiles)
            restart_after = bool(self.var_restart.get())

            def task():
                if edge_pids() and not close_edge(self._log):
                    self._log("✗ Edge 未关闭，已中止恢复", "error")
                    return
                restored = 0
                for profile in profiles:
                    src = folder / profile.name
                    if src.exists():
                        count = restore_profile(profile, src, self._log)
                        restored += count
                        self._log(f"  ✔ {profile.name} 已恢复 {count} 个数据库文件")
                if not restored:
                    self._log("⚠ 这次备份里没有当前配置文件的副本", "warn")
                else:
                    self._q.put(("status", "恢复完成"))
                if restart_after:
                    start_edge(self._log)

            self._run_async(task, "恢复中…")

        bar = ttk.Frame(dlg)
        bar.pack(fill="x", padx=12, pady=10)
        ttk.Button(bar, text="恢复选中备份", command=do_restore).pack(side="left")
        ttk.Button(bar, text="取消", command=dlg.destroy).pack(side="left", padx=8)
        dlg.grab_set()


# --------------------------------------------------------------------------- #
# 命令行
# --------------------------------------------------------------------------- #
def cli_log(message: str, level: str = "info") -> None:
    prefix = {"warn": "⚠ ", "error": "✗ "}.get(level, "")
    print(f"{prefix}{message}", flush=True)


def main(argv: list[str] | None = None) -> int:
    if os.name == "nt":
        try:  # 让 cmd / PowerShell 里的中文正常显示
            import ctypes

            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    parser = argparse.ArgumentParser(
        prog="edge_history_cleaner",
        description="按名单清理 Edge 浏览历史中的指定站点（默认打开图形界面）",
    )
    parser.add_argument("--site", action="append", metavar="DOMAIN", help="目标站点，可重复")
    parser.add_argument("--list-file", metavar="FILE", help="从文本文件读取站点名单")
    parser.add_argument("--title", action="append", metavar="TEXT", help="网页标题规则，可重复；R开头=正则")
    parser.add_argument("--title-file", metavar="FILE", help="从文本文件读取标题名单")
    parser.add_argument("--title-case", action="store_true", help="标题匹配区分大小写")
    parser.add_argument("--profile", action="append", metavar="NAME", help="只处理指定配置文件，可重复")
    parser.add_argument("--profile-dir", metavar="DIR", help="直接指定配置文件目录（调试用）")
    parser.add_argument("--dry-run", action="store_true", help="只统计不修改")
    parser.add_argument("--no-backup", action="store_true", help="不备份")
    parser.add_argument("--backup-dir", metavar="DIR", help="备份目录")
    parser.add_argument("--no-restart", action="store_true", help="清理后不重开 Edge")
    parser.add_argument("--no-stores", action="store_true", help="只清 History，不动其它库")
    parser.add_argument("--keep-referrers", action="store_true", help="不清空来源字段")
    parser.add_argument("--skip-process-control", action="store_true", help="不检测/关闭 Edge（调试用）")
    parser.add_argument("--restore", metavar="RUN_DIR", help="从某次备份目录恢复")
    parser.add_argument("--report", metavar="FILE", help="把结果写入文件")
    parser.add_argument("--version", action="version", version=f"{APP_NAME} {APP_VERSION}")
    args = parser.parse_args(argv)

    if os.name == "nt":
        try:  # 让 cmd / PowerShell 里的中文正常显示
            import ctypes

            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    wanted_gui = not (args.site or args.list_file or args.title or args.title_file
                      or args.profile_dir or args.dry_run or args.restore)
    if wanted_gui and TK_OK:
        enable_dpi_awareness()
        app = CleanerApp()
        app.mainloop()
        return 0
    if wanted_gui:
        cli_log("当前 Python 没有 tkinter，改用命令行模式", "warn")

    text = "\n".join(args.site or [])
    if args.list_file:
        text += "\n" + Path(args.list_file).read_text(encoding="utf-8-sig", errors="replace")
    title_text = "\n".join(args.title or [])
    if args.title_file:
        title_text += "\n" + Path(args.title_file).read_text(encoding="utf-8-sig", errors="replace")
    matcher = SiteMatcher(text, title_text=title_text, ignore_case=not args.title_case)
    for problem in matcher.problems:
        cli_log(problem, "warn")
    if not matcher.has_rules() and not args.restore:
        if matcher.problems:
            cli_log("名单里没有可用的规则（见上面的提示）", "error")
        else:
            cli_log("名单为空，用 --site x.com、--list-file 或 --title 指定", "error")
        return 2

    if args.profile_dir:
        folder = Path(args.profile_dir)
        profiles = [Profile(folder.name, folder, _profile_label(folder))]
    else:
        profiles = discover_profiles()
        if args.profile:
            keep = {n.lower() for n in args.profile}
            profiles = [p for p in profiles if p.name.lower() in keep or p.label.lower() in keep]
    if not profiles:
        cli_log("没有找到匹配的 Edge 配置文件", "error")
        return 2

    if args.restore:
        run_dir = Path(args.restore)
        if not run_dir.exists():
            cli_log(f"备份目录不存在：{run_dir}", "error")
            return 2
        if not args.skip_process_control and edge_pids():
            if not close_edge(cli_log):
                return 2
        restored = 0
        for profile in profiles:
            src = run_dir / profile.name
            if not src.exists():
                cli_log(f"  · 备份里没有 {profile.name}，跳过", "warn")
                continue
            count = restore_profile(profile, src, cli_log)
            restored += count
            cli_log(f"✔ {profile.name} 已恢复 {count} 个数据库文件")
        if not restored:
            cli_log("这次备份里没有可恢复的文件", "error")
            return 1
        if not args.no_restart and not args.skip_process_control:
            start_edge(cli_log)
        return 0

    global options_default_referrers, options_extra_stores
    options_default_referrers = not args.keep_referrers
    options_extra_stores = not args.no_stores

    opts = Options(
        dry_run=args.dry_run,
        backup=not args.no_backup,
        backup_dir=Path(args.backup_dir) if args.backup_dir else default_backup_dir(),
        extra_stores=not args.no_stores,
        referrers=not args.keep_referrers,
        vacuum=True,
        auto_close=not args.skip_process_control,
        auto_restart=not args.no_restart,
        process_control=not args.skip_process_control,
    )

    cli_log(f"目标：{matcher.describe()}")
    cli_log(f"配置文件：{', '.join(p.name for p in profiles)}")

    if args.dry_run:
        total = 0
        for profile in profiles:
            stats = inspect_profile(profile, matcher, cli_log)
            total += stats.get("urls_matched", 0)
            cli_log(f"  {profile.name}: URL {stats.get('urls_matched', 0)} / "
                    f"访问 {stats.get('visits_matched', 0)} / "
                    f"标题命中 {stats.get('title_pages', 0)} / "
                    f"整站 {stats.get('title_hosts', 0)} / "
                    f"常用站点 {stats.get('top_sites', 0)} / "
                    f"地址栏 {stats.get('shortcuts', 0)} / "
                    f"预测 {stats.get('predictor', 0)} / "
                    f"图标 {stats.get('favicon_maps', 0)}")
            roots = stats.get("_expanded_hosts") or []
            if roots:
                cli_log("    标题命中的站点：" + "、".join(roots[:10]) + ("…" if len(roots) > 10 else ""))
        cli_log(f"合计匹配 {total} 条 URL（未做任何修改）")
        return 0

    result = run_clean(profiles, matcher, opts, cli_log)
    report = format_report(result)
    print("\n" + report, flush=True)
    if args.report:
        Path(args.report).write_text(report, encoding="utf-8")
    leftovers = sum(s.get("leftover_urls", 0) for s in result["profiles"].values())
    return 0 if leftovers == 0 else 1


if __name__ == "__main__":
    sys.exit(main())