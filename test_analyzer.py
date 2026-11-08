#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""评论分析回归测试。

不依赖 pytest，直接运行：
    python test_analyzer.py
"""

import sys

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from analyze import CommentAgent, _topic_evidence

PASSED = 0
FAILED = 0


def check(name: str, condition: bool, detail: object = "") -> None:
    global PASSED, FAILED
    if condition:
        PASSED += 1
        print(f"[PASS] {name}")
    else:
        FAILED += 1
        print(f"[FAIL] {name} -> {detail}")


def analyze(content: str, rating: int) -> dict:
    agent = CommentAgent(quiet=True)
    return agent._analyze_comment({
        "id": "T",
        "product": "测试商品",
        "user": "测试用户",
        "rating": rating,
        "content": content,
        "time": "",
    })


def main() -> int:
    positive = analyze("质量很好，面料舒服，物流也快，推荐购买。", 5)
    check("正面评论判为正面", positive["sentiment"] == "正面", positive["sentiment"])

    negative = analyze("严重踩雷，穿了两次就起球，做工太差，非常失望。", 1)
    check("负面评论判为负面", negative["sentiment"] == "负面", negative["sentiment"])

    neutral = analyze("东西一般，价格便宜，日常穿可以。", 3)
    check("中评判为中性", neutral["sentiment"] == "中性", neutral["sentiment"])

    negation = analyze("面料透气不闷汗，夏天穿很舒服。", 5)
    check("否定词“不闷汗”不拉低情感", negation["sentiment_score"] > 0.2, negation["sentiment_score"])

    topics = _topic_evidence("收到货发现袖口有线头，洗了一次有点掉色。")
    check(
        "“收到货”不再误判为物流问题",
        "物流包装" not in topics and "质量做工" in topics,
        list(topics),
    )

    negated_quality = _topic_evidence("客服回复很快，走线整齐，没有多余线头，整体很满意。")
    check(
        "“没有多余线头”不判为质量问题",
        (negated_quality.get("质量做工") or {}).get("polarity") != "负面",
        negated_quality.get("质量做工"),
    )

    agent = CommentAgent(quiet=True)
    read_result = agent.read_comments()
    check("读取示例表格", read_result.get("total") == 5, read_result.get("error") or read_result)

    analysis_result = agent.analyze_comments()
    check("报告包含 5 条评论", (agent.report or {}).get("meta", {}).get("total_comments") == 5)
    check("风险等级为高", analysis_result.get("risk_level") == "高", analysis_result.get("risk_level"))
    check(
        "报告包含 meta / summary / comments",
        set(agent.report or {}) == {"meta", "summary", "comments"},
        list(agent.report or {}),
    )

    print(f"\n{PASSED} passed, {FAILED} failed")
    return 1 if FAILED else 0


if __name__ == "__main__":
    raise SystemExit(main())
