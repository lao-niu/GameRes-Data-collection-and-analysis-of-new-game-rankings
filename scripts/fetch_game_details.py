#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用途：
    读取 date.xlsx 的榜单排名，取前 30 名，逐个抓取游戏详情接口，
    为每个游戏生成一份分区块的「游戏名.txt」。

数据来源：
    列表：date.xlsx（由 fetch_ranking.py 生成）
    详情：https://www.gameres.com/api/public/v1/gamecenter/game?gameid=

运行方式：
    python fetch_game_details.py
    python fetch_game_details.py --ranking ../date.xlsx --out-dir ../top30
    python fetch_game_details.py --limit 1 --out-dir ./tmp

依赖：
    requests、openpyxl
"""

import argparse
import datetime as dt
import html
import json
import os
import re
import time
from pathlib import Path

import requests
from openpyxl import load_workbook

API_DETAIL_URL = "https://www.gameres.com/api/public/v1/gamecenter/game"
SITE_URL = "https://www.16p.com"

SOURCE_MAP = {2: "App Store", 3: "TapTap", 5: "Steam", 6: "3839"}
PLATFORM_NAME = {
    "itunes": "iOS（App Store）",
    "taptap": "TapTap",
    "hykb": "3839（好游快爆）",
    "steam": "PC（Steam）",
}

DEFAULT_HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "zh-CN,zh;q=0.9,en;q=0.8",
    "origin": "https://www.16p.com",
    "referer": "https://www.16p.com/",
    "user-agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36 Edg/156.0.0.0"
    ),
    "X-Requested-With": "XMLHttpRequest",
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


def load_config(config_path: Path) -> dict:
    """读取同目录 config.json，并允许环境变量覆盖 Cookie / Token。"""
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


def read_top_games(ranking_path: Path, limit: int) -> list:
    """读取 Excel 表头，按排名升序取前 limit 名。"""
    if not ranking_path.exists():
        raise SystemExit(
            f"找不到榜单文件：{ranking_path}\n"
            "请先运行 fetch_ranking.py 生成 date.xlsx，或用 --ranking 指定文件。"
        )

    wb = load_workbook(ranking_path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        raise SystemExit(f"榜单文件为空：{ranking_path}")

    required = ["排名", "游戏名称", "游戏ID"]
    missing = [name for name in required if name not in header]
    if missing:
        raise SystemExit(f"榜单文件缺少必要列：{'、'.join(missing)}")

    index = {name: header.index(name) for name in required}
    games = []
    for row in rows:
        rank = row[index["排名"]]
        name = row[index["游戏名称"]]
        game_id = row[index["游戏ID"]]
        if rank is None or game_id is None:
            continue
        games.append({"rank": int(rank), "name": str(name or ""), "gameid": int(game_id)})

    games.sort(key=lambda item: item["rank"])
    return games[:limit]


def fetch_detail(session: requests.Session, gameid: int, cfg: dict) -> dict:
    """抓取单个游戏详情；失败重试，耗尽后抛出异常。"""
    retries = int(cfg.get("retries", 4))
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            resp = session.get(
                API_DETAIL_URL,
                params={"gameid": str(gameid)},
                timeout=cfg.get("timeout", 25),
            )
            resp.raise_for_status()
            payload = resp.json()
            if payload.get("code") != 200:
                raise RuntimeError(f"接口返回 code={payload.get('code')}")
            game = (payload.get("data") or {}).get("game") or {}
            if not game.get("gameid"):
                raise RuntimeError("接口未返回 game 数据")
            return game
        except (requests.RequestException, ValueError, RuntimeError) as exc:
            last_error = exc
            print(f"  游戏 {gameid} 第 {attempt}/{retries} 次请求失败：{exc}")
            if attempt < retries:
                time.sleep(float(cfg.get("sleep_seconds", 1.5)))

    raise RuntimeError(f"游戏 {gameid} 详情抓取失败，已重试 {retries} 次：{last_error}")


def unique_join(items) -> str:
    """去重后按中文分号拼接。"""
    seen, out = set(), []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            out.append(item)
    return "；".join(out)


def clean(text) -> str:
    """清理接口中的 HTML 片段。"""
    if text is None:
        return ""
    value = str(text)
    value = re.sub(r"<br\s*/?>", "\n", value, flags=re.I)
    value = re.sub(r"<[^>]+>", "", value)
    value = html.unescape(value).replace("\u002F", "/")
    return value.strip()


def fmt(value) -> str:
    """统一空值口径：缺失一律写「页面未提供」。"""
    if value is None:
        return "页面未提供"
    if isinstance(value, str) and not value.strip():
        return "页面未提供"
    if isinstance(value, (list, tuple)):
        items = [str(item).strip() for item in value if str(item).strip()]
        return "、".join(items) if items else "页面未提供"
    if isinstance(value, dict) and not value:
        return "页面未提供"
    return str(value)


def fmt_bytes(value) -> str:
    """把字节数转成可读包体大小；非正数视为缺失。"""
    try:
        number = int(value)
        if number <= 0:
            return "页面未提供"
        for unit in ["B", "KB", "MB", "GB"]:
            if number < 1024:
                return f"{number} B" if unit == "B" else f"{number:.2f} {unit}"
            number /= 1024
        return f"{number:.2f} TB"
    except (TypeError, ValueError):
        return "页面未提供"


def sanitize_filename(name: str) -> str:
    """把 Windows 非法文件名字符替换为全角字符，并限制长度。"""
    trans = {
        "\\": "＼", "/": "／", ":": "：", "*": "＊", "?": "？",
        '"': "”", "<": "＜", ">": "＞", "|": "｜",
    }
    value = "".join(trans.get(ch, ch) for ch in str(name))
    value = re.sub(r"[\x00-\x1f]", "", value).strip().rstrip(". ")
    return value[:100] or "未命名游戏"


def plat_url(game: dict, source_id: int) -> str:
    """从 plat_list 中取指定平台的商店链接。"""
    for item in game.get("plat_list") or []:
        if item.get("source_id") == source_id and item.get("url"):
            return item["url"]
    return "页面未提供"


def tests_sorted(game: dict) -> list:
    """开测记录按日期倒序。"""
    tests = game.get("test") or []
    return sorted(tests, key=lambda item: str(item.get("testdate") or ""), reverse=True)


def build_platform_blocks(game: dict) -> list:
    """按平台生成 iOS / TapTap / 3839 / Steam 信息区块。"""
    plat = game.get("plat") or {}
    blocks = []

    item = plat.get("itunes")
    if item:
        rate = item.get("averageUserRating")
        lines = [
            "■ iOS（App Store）",
            f'  App Store ID：{fmt(item.get("trackId"))}',
            f'  应用名称：{fmt(item.get("trackCensoredName"))}',
            f'  Bundle ID：{fmt(item.get("bundleId"))}',
            f'  开发者/卖家：{fmt(item.get("sellerName"))}',
            f'  评分：{fmt(rate)}{" / 5" if rate is not None else ""}',
            f'  评价数：{fmt(item.get("userRatingCount"))}',
            f'  版本：{fmt(item.get("version"))}',
            f'  包体大小：{fmt_bytes(item.get("fileSizeBytes"))}',
            f'  上线日期：{fmt(item.get("releaseDate"))}',
            f'  当前版本发布日期：{fmt(item.get("currentVersionReleaseDate"))}',
            f'  最低系统版本：{fmt(item.get("minimumOsVersion"))}',
            f'  分类：{fmt(item.get("primaryGenreName"))}',
            f'  平台标签：{fmt(item.get("genres"))}',
            f'  商店链接：{plat_url(game, 2)}',
            f'  更新时间：{fmt(item.get("updatetime"))}',
        ]
        blocks.append("\n".join(lines))

    item = plat.get("taptap")
    if item:
        rate = item.get("ratingvalue")
        lines = [
            "■ 安卓 / iOS（TapTap）",
            f'  TapTap ID：{fmt(item.get("taptap_id"))}',
            f'  平台评分：{fmt(rate)}{" / 10" if rate is not None else ""}',
            f'  评价数：{fmt(item.get("review_num"))}',
            f'  下载量：{fmt(item.get("download_times"))}',
            f'  包体大小：{fmt(item.get("filesize"))}',
            f'  版本：{fmt(item.get("softwareversion"))}',
            f'  上线日期：{fmt(item.get("datepublished"))}',
            f'  安卓下载链接：{fmt(item.get("androidurl"))}',
            f'  iOS 下载链接：{fmt(item.get("iosurl"))}',
            f'  平台标签：{fmt(item.get("gameplay"))}',
            f'  平台链接：{plat_url(game, 3)}',
            f'  更新时间：{fmt(item.get("updatetime"))}',
        ]
        blocks.append("\n".join(lines))

    item = plat.get("hykb")
    if item:
        rate = item.get("app_rate")
        lines = [
            "■ 安卓（3839 / 好游快爆）",
            f'  App ID：{fmt(item.get("app_id"))}',
            f'  发行商：{fmt(item.get("app_publisher"))}',
            f'  平台评分：{fmt(rate)}{" / 10" if rate is not None else ""}',
            f'  评价数：{fmt(item.get("app_review_number"))}',
            f'  上线日期：{fmt(item.get("app_release_date"))}',
            f'  游戏类型：{fmt(item.get("app_genre"))}',
            f'  平台链接：{plat_url(game, 6)}',
            f'  更新时间：{fmt(item.get("update_time"))}',
        ]
        blocks.append("\n".join(lines))

    item = plat.get("steam")
    if item:
        lines = [
            "■ PC（Steam）",
            f'  Steam AppID：{fmt(item.get("steam_id"))}',
            f'  开发商：{fmt(item.get("app_developer"))}',
            f'  发行商：{fmt(item.get("app_publisher"))}',
            f'  发售日期：{fmt(item.get("app_release_date"))}',
            f'  游戏类型：{fmt(item.get("app_genre"))}',
        ]
        summary = item.get("app_all_review_summary") or ""
        rate = item.get("app_all_favorable_rate") or ""
        number = item.get("app_all_review_number")
        if summary or rate or number is not None:
            parts = []
            if summary:
                parts.append(str(summary))
            if rate:
                parts.append(f"好评率 {rate}")
            if number is not None:
                parts.append(f"共 {number} 条评价")
            lines.append(f"  总体评价：{'，'.join(parts)}")
        else:
            lines.append("  总体评价：页面未提供")

        recent_summary = item.get("app_recent_review_summary") or ""
        recent_rate = item.get("app_recent_favorable_rate") or ""
        recent_number = item.get("app_recent_review_number")
        if recent_summary or recent_rate or recent_number is not None:
            parts = []
            if recent_summary:
                parts.append(str(recent_summary))
            if recent_rate:
                parts.append(f"好评率 {recent_rate}")
            if recent_number is not None:
                parts.append(f"共 {recent_number} 条评价")
            lines.append(f"  近期评价：{'，'.join(parts)}")
        else:
            lines.append("  近期评价：页面未提供")

        sys_req = clean(item.get("app_sys_req") or "")
        match = re.search(r"存储空间:\s*需要\s*([^\n<]+)", sys_req)
        lines.append(f'  存储需求：{match.group(1).strip() if match else "页面未提供"}')
        lines.append(f'  商店链接：{plat_url(game, 5)}')
        lines.append(f'  数据更新时间：{fmt(item.get("update_time"))}')
        blocks.append("\n".join(lines))

    return blocks


def build_txt(game: dict, rank: int, fetch_time: str) -> str:
    """生成单个游戏的分区块文本内容。"""
    lines = [
        "=" * 60,
        f'游戏详情：{game.get("gamename")}',
        "=" * 60,
        "",
        "【基本信息】",
        f"榜单排名：{rank}",
        f'游戏名称：{game.get("gamename")}',
        f'游戏ID：{game.get("gameid")}',
        f'来源详情页：{SITE_URL}/{game.get("gameid")}.html',
        f"抓取时间：{fetch_time}",
        f'地区：{fmt(game.get("area"))}',
        f'详情评分（10分制）：{fmt(game.get("review_rate"))}',
        f'评价数量：{fmt(game.get("review_num"))}',
        f'游戏标签：{fmt("、".join(game.get("gameplay") or []))}',
    ]

    stage = ""
    for info in game.get("need_info") or []:
        if info.get("field") == "gametry" and info.get("value"):
            stage = info["value"]
            break
    lines.append(f"产品阶段：{stage or '页面未提供'}")
    lines.append("")
    lines.append("【开测数据】")

    tests = tests_sorted(game)
    public_tests = [item for item in tests if "公测" in str(item.get("testtype") or "")]
    if public_tests:
        lines.append(f'公测时间：{public_tests[0].get("testdate")}（{public_tests[0].get("testtype")}）')
    else:
        lines.append("公测时间：页面未提供（接口开测记录中无“公测”类型）")

    if tests:
        latest = tests[0]
        test_type = str(latest.get("testtype") or "")
        if str(latest.get("testdate") or "") > dt.date.today().strftime("%Y-%m-%d"):
            status = f'待上线/待测试（最新记录：{latest.get("testdate")} {test_type}）'
        elif "上线" in test_type:
            status = f'已上线（最新记录：{latest.get("testdate")} {test_type}）'
        else:
            status = f'已开测（最新记录：{latest.get("testdate")} {test_type}）'
        lines.append(f"当前测试状态：{status}")
        lines.append(f'最近开测时间：{fmt(latest.get("testdate"))}')
        lines.append(f"测试类型：{test_type or '页面未提供'}")
        lines.append(f"开测记录（共 {len(tests)} 条，按日期倒序）：")
        for index, item in enumerate(tests, 1):
            domestic = "是" if item.get("is_cn") == 1 else "否"
            lines.append(
                f"  {index}. {fmt(item.get('testdate'))}  "
                f"{fmt(item.get('testtype'))}  国内={domestic}"
            )
    else:
        lines.extend([
            "当前测试状态：页面未提供",
            "最近开测时间：页面未提供",
            "测试类型：页面未提供",
            "开测记录：页面未提供",
        ])

    lines.append("")
    lines.append("【厂商与发行】")
    developers = [
        f'{company["name"]}（{SOURCE_MAP.get(company.get("data_source"), "未知来源")}）'
        for company in (game.get("companys") or [])
        if company.get("company_role_id") == 1 and company.get("name")
    ]
    publishers = [
        f'{company["name"]}（{SOURCE_MAP.get(company.get("data_source"), "未知来源")}）'
        for company in (game.get("companys") or [])
        if company.get("company_role_id") == 2 and company.get("name")
    ]
    lines.append(f"开发商（研发）：{unique_join(developers) or '页面未提供'}")
    lines.append(f"发行商：{unique_join(publishers) or '页面未提供'}")
    lines.append("说明：厂商角色按接口 company_role_id 区分（1=开发/研发，2=发行）。")
    lines.append("")
    lines.append("【平台信息】")

    blocks = build_platform_blocks(game)
    if blocks:
        platform_names = [
            PLATFORM_NAME.get(key, key)
            for key in (game.get("plat") or {})
        ]
        lines.append(f"平台汇总：{'、'.join(platform_names)}")
        lines.append("")
        lines.append(("\n" + "-" * 50 + "\n").join(blocks))
    else:
        lines.append("平台信息：页面未提供")

    lines.append("")
    lines.append("=" * 60)
    lines.append("数据来源：玩匠（16p.com）游戏详情接口；抓取时间见上方。")
    return "\n".join(lines)


def main() -> None:
    script_dir = Path(__file__).resolve().parent
    project_dir = script_dir.parent

    parser = argparse.ArgumentParser(description="抓取前 N 名游戏详情并生成 txt")
    parser.add_argument("--ranking", type=Path, default=project_dir / "date.xlsx",
                        help="榜单 Excel 路径，默认项目根目录 date.xlsx")
    parser.add_argument("--out-dir", type=Path, default=project_dir,
                        help="txt 输出目录，默认项目根目录")
    parser.add_argument("--config", type=Path, default=script_dir / "config.json",
                        help="配置文件路径，默认脚本目录 config.json")
    parser.add_argument("--limit", type=int, default=30,
                        help="抓取前多少名，默认 30")
    args = parser.parse_args()

    cfg = load_config(args.config)
    games = read_top_games(args.ranking, args.limit)
    if not games:
        raise SystemExit("榜单中没有可抓取的游戏。")

    session = requests.Session()
    session.headers.update(build_headers(cfg))
    args.out_dir.mkdir(parents=True, exist_ok=True)

    used_names = set()
    success = 0
    for index, item in enumerate(games, 1):
        try:
            game = fetch_detail(session, item["gameid"], cfg)
        except RuntimeError as exc:
            print(f"[{index}/{len(games)}] 跳过：{exc}")
            continue

        base = sanitize_filename(game.get("gamename") or item["name"])
        filename = base
        suffix = 2
        while filename.lower() in used_names:
            filename = f"{base}_{suffix}"
            suffix += 1
        used_names.add(filename.lower())

        fetched_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        content = build_txt(game, item["rank"], fetched_at)
        out_path = args.out_dir / f"{filename}.txt"
        out_path.write_text(content, encoding="utf-8")
        success += 1
        print(f"[{index}/{len(games)}] {item['rank']}. {game.get('gamename') or item['name']} -> {out_path.name}")

        if index < len(games):
            time.sleep(float(cfg.get("sleep_seconds", 1.5)))

    print("=" * 50)
    print(f"完成：{success}/{len(games)}，输出目录：{args.out_dir}")
    if success < len(games):
        raise SystemExit(f"有 {len(games) - success} 个游戏抓取失败，请查看上方日志。")


if __name__ == "__main__":
    main()
