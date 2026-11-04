#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用途：
    读取 date.xlsx，复算报告里的核心指标：
    1) 品类占比（按首个可识别品类词）
    2) 开发商 CR5 / CR10
    3) 评价数 Top10 及其占全表比例
    4) 开测月份分布（从详情 txt 的开测记录中解析）

数据来源：
    date.xlsx（由 fetch_ranking.py 生成）
    详情 txt 目录（由 fetch_game_details.py 生成，默认 ../top30）

运行方式：
    python analysis_ranking.py
    python analysis_ranking.py --ranking ../date.xlsx --details-dir ../top30

依赖：
    pandas、openpyxl
"""

import argparse
import re
from collections import Counter
from pathlib import Path

import pandas as pd

# 品类词表：按游戏标签中出现顺序，取第一个命中的词作为主品类
GENRE_WORDS = [
    "角色扮演", "策略", "模拟经营", "休闲", "动作", "冒险", "卡牌", "射击",
    "体育", "竞速", "塔防", "二次元", "消除", "音乐", "解谜", "放置", "SLG",
    "Roguelike", "MMORPG", "换装", "养成", "经营", "模拟", "RPG", "MOBA",
    "沙盒", "开放世界", "恐怖", "科幻", "武侠", "仙侠", "三国", "传奇",
    "回合制", "即时战略", "跑酷", "宠物", "恋爱", "国风", "暗黑", "篮球",
    "足球", "钓鱼", "烹饪", "装修", "农场", "建造", "生存", "战棋", "自走棋", "格斗",
]

# 平台、活动、商业化标签不参与品类归类
SKIP_TAGS = {
    "编辑推荐", "PC游戏", "多平台", "独立游戏", "买断制", "免费", "付费",
    "单机", "网游", "手游", "联运", "出海", "TapTap制造", "聚光灯GameJam", "Steam",
}


def category_of(tags: str) -> str:
    """取第一个可识别的品类词；没有则归为「其他」。"""
    parts = [part.strip() for part in str(tags or "").split("、") if part.strip()]
    for part in parts:
        if part in SKIP_TAGS:
            continue
        if part in GENRE_WORDS:
            return part
    return "其他"


def parse_review_count(value) -> int:
    """把「万 / 亿」文本评价数换算为整数。"""
    text = str(value or "").strip()
    if not text or text.lower() == "nan":
        return 0
    if "万" in text:
        return int(float(text.replace("万", "")) * 10000)
    if "亿" in text:
        return int(float(text.replace("亿", "")) * 100000000)
    try:
        return int(float(text.replace(",", "")))
    except ValueError:
        return 0


def percent(part: int, total: int) -> str:
    return f"{part / total * 100:.1f}%" if total else "0.0%"


def print_genre(df: pd.DataFrame) -> None:
    counts = Counter(category_of(tag) for tag in df["游戏标签"])
    total = len(df)
    top_four = counts.most_common(4)
    print("一、品类占比（按首个可识别品类词，单主品类口径）")
    for name, count in top_four:
        print(f"  {name}：{count} 条，{percent(count, total)}")
    top_four_sum = sum(count for _, count in top_four)
    other = total - top_four_sum
    # 「其他」拆成两部分，避免与报告口径混淆：
    # 已归类其他 = 命中了品类词但没进前四；未归类 = 标签里一个品类词都没有
    top_four_names = {name for name, _ in top_four}
    classified_other = sum(
        count for name, count in counts.items()
        if name not in top_four_names and name != "其他"
    )
    unclassified = counts.get("其他", 0)
    print(f"  前四合计：{top_four_sum} 条，{percent(top_four_sum, total)}")
    print(f"  其他（已归类，未进前四）：{classified_other} 条，{percent(classified_other, total)}")
    print(f"  其他（未归类，标签无品类词）：{unclassified} 条，{percent(unclassified, total)}")
    print(f"  其他合计：{other} 条，{percent(other, total)}")
    print(f"  总计：{total} 条\n")


def print_developers(df: pd.DataFrame) -> None:
    counts = Counter()
    for value in df["开发商"]:
        text = str(value or "").strip()
        if not text or text == "页面未提供":
            continue
        for name in text.split("；"):
            name = name.strip()
            if name and name != "页面未提供":
                counts[name] += 1

    total = len(df)
    ranked = counts.most_common()
    cr5 = sum(count for _, count in ranked[:5])
    cr10 = sum(count for _, count in ranked[:10])
    print("二、开发商集中度（按榜单记录数计，分母为全部 537 条）")
    if ranked:
        print(f"  Top1：{ranked[0][0]} {ranked[0][1]} 条，{percent(ranked[0][1], total)}")
    print(f"  CR5：{cr5} 条，{percent(cr5, total)}")
    print(f"  CR10：{cr10} 条，{percent(cr10, total)}")
    print(f"  有开发商记录的条数：{sum(1 for v in df['开发商'] if str(v or '').strip() not in ('', '页面未提供'))}\n")


def print_reviews(df: pd.DataFrame) -> None:
    values = [
        (parse_review_count(value), str(name))
        for name, value in zip(df["游戏名称"], df["评价数"])
    ]
    values.sort(key=lambda item: item[0], reverse=True)
    total = sum(value for value, _ in values)
    top10 = values[:10]
    top10_total = sum(value for value, _ in top10)

    print("三、评价数 Top10 及其占全表比例")
    print(f"  全表评价总数：{total:,}")
    print(f"  Top10 合计：{top10_total:,}，占 {percent(top10_total, total)}")
    for index, (value, name) in enumerate(top10, 1):
        print(f"  {index:>2}. {name}：{value:,}")
    print()


def print_test_months(details_dir: Path) -> None:
    print("四、开测月份分布（来自详情 txt 的开测记录）")
    if not details_dir.exists():
        print(f"  未找到详情目录：{details_dir}，已跳过。请先运行 fetch_game_details.py。\n")
        return

    pattern = re.compile(r"^\s*\d+\.\s+(20\d{2}-\d{2}-\d{2})", re.MULTILINE)
    dates = []
    for path in sorted(details_dir.glob("*.txt")):
        dates.extend(pattern.findall(path.read_text(encoding="utf-8")))

    if not dates:
        print(f"  目录 {details_dir} 中没有解析到开测记录，已跳过。\n")
        return

    months = Counter(date[:7] for date in dates)
    months_2026 = sum(count for month, count in months.items() if month.startswith("2026"))
    older = sum(count for month, count in months.items() if not month.startswith("2026"))

    for month in sorted(month for month in months if month.startswith("2026")):
        print(f"  {month}：{months[month]} 条")
    print(f"  2026 年合计：{months_2026} 条")
    print(f"  2025 及以前：{older} 条")
    print(f"  总计：{len(dates)} 条")

    if months_2026:
        october = months.get("2026-10", 0)
        print(f"  2026-10：{october} 条，占 2026 年的 {percent(october, months_2026)}")
    october_15 = sum(1 for date in dates if date == "2026-10-15")
    print(f"  2026-10-15 单日：{october_15} 条\n")


def main() -> None:
    script_dir = Path(__file__).resolve().parent
    project_dir = script_dir.parent

    parser = argparse.ArgumentParser(description="复算新游榜单核心指标")
    parser.add_argument("--ranking", type=Path, default=project_dir / "date.xlsx",
                        help="榜单 Excel 路径，默认项目根目录 date.xlsx")
    parser.add_argument("--details-dir", type=Path, default=project_dir / "top30",
                        help="详情 txt 目录，默认项目根目录 top30")
    args = parser.parse_args()

    if not args.ranking.exists():
        raise SystemExit(
            f"找不到榜单文件：{args.ranking}\n"
            "请先运行 fetch_ranking.py 生成 date.xlsx。"
        )

    df = pd.read_excel(args.ranking, dtype=str)
    required = ["排名", "游戏名称", "游戏标签", "开发商", "评价数"]
    missing = [name for name in required if name not in df.columns]
    if missing:
        raise SystemExit(f"榜单文件缺少必要列：{'、'.join(missing)}")

    df = df.dropna(subset=["排名"]).copy()
    df["排名"] = df["排名"].astype(int)
    df = df.sort_values("排名")

    print(f"读取：{args.ranking}，共 {len(df)} 条\n")
    print_genre(df)
    print_developers(df)
    print_reviews(df)
    print_test_months(args.details_dir)


if __name__ == "__main__":
    main()
