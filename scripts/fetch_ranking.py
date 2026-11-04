#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用途：
    抓取 gameres 游戏中心（16p.com）新游榜全量数据，并导出为 date.xlsx。
    接口为公开榜单接口，默认参数：国内游戏、近 30 天、每页 50 条。

数据来源：
    https://www.gameres.com/api/public/v1/gamecenter/new_game_list

运行方式：
    python fetch_ranking.py
    python fetch_ranking.py --out ../date.xlsx
    python fetch_ranking.py --max-pages 1 --out ./test.xlsx

依赖：
    requests、openpyxl
"""

import argparse
import datetime as dt
import json
import os
import sys
import time
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

API_URL = "https://www.gameres.com/api/public/v1/gamecenter/new_game_list"
PAGE_SIZE = 50
DATE_RANGE = 30
TYPE_RANGE = 2

DEFAULT_HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "zh-CN,zh;q=0.9,en;q=0.8",
    "origin": "https://www.16p.com",
    "referer": "https://www.16p.com/",
    "sec-ch-ua": '"Not:A-Brand";v="8", "Chromium";v="156", "Microsoft Edge";v="156"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "cross-site",
    "user-agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36 Edg/156.0.0.0"
    ),
}

CONFIG_DEFAULTS = {
    "cookie": "",
    "token": "",
    "token_header": "Authorization",
    "require_auth": False,
    "timeout": 25,
    "sleep_seconds": 1.5,
    "retries": 4,
}

HEADERS = [
    "排名", "游戏名称", "游戏ID", "详情页链接", "评分（5分制）", "评价数",
    "游戏标签", "地区", "榜单类型", "榜单时间范围", "收录/发布日期",
    "开发商", "发行商", "权重（一级）", "权重（二级）", "图标链接",
    "主图链接", "抓取时间",
]

SITE_URL = "https://www.16p.com"


def load_config(config_path: Path) -> dict:
    """读取同目录 config.json，并允许环境变量覆盖 Cookie / Token。

    该榜单接口是公开接口，Cookie 与 Token 可以留空；
    若后续接口需要登录，将 require_auth 设为 true，缺失凭据时会给出明确报错。
    """
    cfg = dict(CONFIG_DEFAULTS)

    if config_path.exists():
        try:
            data = json.loads(config_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise SystemExit(
                f"配置文件解析失败：{config_path}\n"
                f"请检查 JSON 格式是否正确。原始错误：{exc}"
            )
        if not isinstance(data, dict):
            raise SystemExit(f"配置文件必须是 JSON 对象：{config_path}")
        cfg.update({k: v for k, v in data.items() if v is not None and v != ""})

    # 环境变量优先级高于配置文件，便于在 CI 或本地临时注入凭据
    if os.getenv("GAMERES_COOKIE"):
        cfg["cookie"] = os.environ["GAMERES_COOKIE"]
    if os.getenv("GAMERES_TOKEN"):
        cfg["token"] = os.environ["GAMERES_TOKEN"]

    if cfg.get("require_auth") and not (cfg.get("cookie") or cfg.get("token")):
        raise SystemExit(
            "配置要求登录凭据，但未提供 Cookie 或 Token。\n"
            "请在 config.json 中填写 cookie/token，或设置环境变量 "
            "GAMERES_COOKIE / GAMERES_TOKEN。"
        )

    if not (cfg.get("cookie") or cfg.get("token")):
        print("提示：未提供 Cookie/Token，将以匿名方式访问公开接口。")

    return cfg


def build_headers(cfg: dict) -> dict:
    """在基础浏览器请求头上，按需追加 Cookie / Token。"""
    headers = dict(DEFAULT_HEADERS)
    if cfg.get("cookie"):
        headers["Cookie"] = cfg["cookie"]
    if cfg.get("token"):
        token = cfg["token"]
        header_name = cfg.get("token_header") or "Authorization"
        if header_name.lower() == "authorization" and not token.lower().startswith("bearer "):
            token = f"Bearer {token}"
        headers[header_name] = token
    return headers


def unique_join(items) -> str:
    """去重后按中文分号拼接，用于开发商 / 发行商字段。"""
    seen, out = set(), []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            out.append(item)
    return "；".join(out)


def companies(row: dict, role: int) -> str:
    """按 company_role_id 取开发商（1）或发行商（2）。"""
    names = [
        c.get("name")
        for c in (row.get("companys") or [])
        if c.get("company_role_id") == role and c.get("name")
    ]
    return unique_join(names)


def fetch_page(session: requests.Session, page: int, cfg: dict) -> list:
    """抓取单页；失败重试次数由配置控制，重试耗尽后抛出异常。"""
    params = {
        "p": str(page),
        "ps": str(PAGE_SIZE),
        "date_range": str(DATE_RANGE),
        "type_range": str(TYPE_RANGE),
    }
    retries = int(cfg.get("retries", 4))
    last_error = None

    for attempt in range(1, retries + 1):
        try:
            resp = session.get(API_URL, params=params, timeout=cfg.get("timeout", 25))
            resp.raise_for_status()
            payload = resp.json()
            if payload.get("code") != 200:
                raise RuntimeError(f"接口返回 code={payload.get('code')}，message={payload.get('message')}")
            return payload.get("data") or []
        except (requests.RequestException, ValueError, RuntimeError) as exc:
            last_error = exc
            print(f"  第 {page} 页第 {attempt}/{retries} 次请求失败：{exc}")
            if attempt < retries:
                time.sleep(float(cfg.get("sleep_seconds", 1.5)))

    raise RuntimeError(f"第 {page} 页抓取失败，已重试 {retries} 次：{last_error}")


def export_excel(rows: list, out_path: Path, fetched_at: str) -> None:
    """按 18 列中文字段导出 Excel，并做基础排版。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "榜单数据"
    ws.append(HEADERS)

    for row in rows:
        tags = "、".join(row.get("gameplay") or [])
        ws.append([
            row.get("num"),
            row.get("gamename"),
            row.get("gameid"),
            f"{SITE_URL}/{row.get('gameid')}.html",
            row.get("review_rate"),
            row.get("review_num"),
            tags,
            row.get("area"),
            "国内游戏",
            "近30天",
            row.get("format_time"),
            companies(row, 1) or "页面未提供",
            companies(row, 2) or "页面未提供",
            row.get("first_weight_num"),
            row.get("second_weight_num"),
            row.get("iconurl"),
            row.get("mainimg"),
            fetched_at,
        ])

    head_fill = PatternFill("solid", fgColor="DDEBF7")
    thin = Side(style="thin", color="BFBFBF")
    for cell in ws[1]:
        cell.font = Font(bold=True, size=11)
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    widths = [6, 26, 10, 38, 7, 10, 34, 8, 12, 12, 15, 28, 28, 9, 9, 32, 32, 20]
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=cell.column in (7, 12, 13, 16, 17))
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:R{ws.max_row}"
    ws.row_dimensions[1].height = 22
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> None:
    script_dir = Path(__file__).resolve().parent
    project_dir = script_dir.parent

    parser = argparse.ArgumentParser(description="抓取 gameres 新游榜并导出 date.xlsx")
    parser.add_argument("--out", type=Path, default=project_dir / "date.xlsx",
                        help="输出 Excel 路径，默认项目根目录 date.xlsx")
    parser.add_argument("--config", type=Path, default=script_dir / "config.json",
                        help="配置文件路径，默认脚本目录 config.json")
    parser.add_argument("--max-pages", type=int, default=0,
                        help="最多抓取页数，0 表示抓全（调试或抽样时使用）")
    args = parser.parse_args()

    cfg = load_config(args.config)
    session = requests.Session()
    session.headers.update(build_headers(cfg))

    all_rows = []
    page = 1
    try:
        while True:
            rows = fetch_page(session, page, cfg)
            all_rows.extend(rows)
            rank_range = ""
            if rows:
                rank_range = f"（排名 {rows[0].get('num')}–{rows[-1].get('num')}）"
            print(f"第 {page} 页：{len(rows)} 条{rank_range}")

            if not rows or len(rows) < PAGE_SIZE:
                break
            if args.max_pages and page >= args.max_pages:
                print(f"已达到 --max-pages={args.max_pages}，提前结束。")
                break
            page += 1
            time.sleep(float(cfg.get("sleep_seconds", 1.5)))
    except RuntimeError as exc:
        raise SystemExit(f"\n抓取中止：{exc}")

    if not all_rows:
        raise SystemExit("未抓到任何数据，请检查网络、接口状态或配置。")

    all_rows.sort(key=lambda item: int(item.get("num") or 0))
    fetched_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    export_excel(all_rows, args.out, fetched_at)

    ranks = [int(item.get("num")) for item in all_rows]
    game_ids = [item.get("gameid") for item in all_rows]
    print("=" * 50)
    print(f"总条数：{len(all_rows)}")
    print(f"页数：{page}")
    print(f"排名范围：{min(ranks)}–{max(ranks)}，连续：{max(ranks) - min(ranks) + 1 == len(set(ranks))}")
    print(f"游戏 ID：{len(set(game_ids))} 个，无重复：{len(set(game_ids)) == len(game_ids)}")
    print(f"输出文件：{args.out}")


if __name__ == "__main__":
    main()
