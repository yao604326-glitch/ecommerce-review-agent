#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""商品评论语义分析智能体（Tool Calling Agent）。

工作流程：
    1. read_comments     读取 Excel 评论表
    2. analyze_comments  语义分析（情感、主题、风险、运营建议）
    3. save_report       保存 JSON（结构化）与 Markdown（可读）报告

两种运行模式：
    - LLM 模式：配置 OPENAI_API_KEY 后，由大模型通过 function calling 编排工具；
    - 离线模式：未配置 Key（或加 --offline）时，使用内置规则引擎按同样流程执行。

用法：
    python analyze.py
    python analyze.py --offline
    python analyze.py -f comments.xlsx -o reports
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = BASE_DIR / "comments.xlsx"
DEFAULT_OUTPUT_DIR = BASE_DIR / "reports"
MAX_AGENT_STEPS = 8

try:
    from openpyxl import load_workbook

    OPENPYXL_AVAILABLE = True
except ImportError:  # pragma: no cover
    OPENPYXL_AVAILABLE = False


# ---------------------------------------------------------------------------
# 数据字典：电商评论领域的情感词与主题词
# ---------------------------------------------------------------------------

POSITIVE_LEXICON: dict[str, float] = {
    "超出预期": 1.2, "非常满意": 1.2, "很满意": 1.0, "满意": 0.8,
    "很喜欢": 1.0, "喜欢": 0.7, "推荐": 0.8, "回购": 0.9,
    "惊喜": 1.0, "好评": 0.9, "很好": 0.9, "不错": 0.7,
    "挺好": 0.7, "还行": 0.2, "可以": 0.3, "好用": 0.8,
    "值得": 0.7, "舒服": 0.8, "舒适": 0.8, "透气": 0.8,
    "亲肤": 0.8, "凉快": 0.7, "好看": 0.7, "漂亮": 0.7,
    "细致": 0.7, "用心": 0.7, "标准": 0.6, "合身": 0.6,
    "合适": 0.6, "划算": 0.7, "便宜": 0.6, "实惠": 0.7,
    "快": 0.5, "迅速": 0.6, "及时": 0.6, "耐心": 0.6,
    "专业": 0.6, "热情": 0.6, "干净": 0.5, "完好": 0.6,
    "厚实": 0.6, "柔软": 0.7, "顺滑": 0.5,
}

# 负面词表使用正数权重，合并时统一取负号，避免手写遗漏
NEGATIVE_LEXICON: dict[str, float] = {
    "严重踩雷": 1.5, "踩雷": 1.2, "非常失望": 1.3, "失望": 1.0,
    "太差": 1.3, "很差": 1.1, "差评": 1.0, "垃圾": 1.3,
    "欺骗": 1.3, "投诉": 1.0, "维权": 1.0, "起球": 1.0,
    "变形": 0.9, "掉色": 0.9, "褪色": 0.9, "缩水": 0.9,
    "线头": 0.7, "开线": 1.0, "破损": 1.0, "瑕疵": 0.8,
    "色差": 0.9, "差别挺大": 0.9, "货不对板": 1.2, "图文不符": 1.0,
    "偏小": 0.6, "偏大": 0.6, "不合身": 0.8, "紧绷": 0.7,
    "太大": 0.4, "太小": 0.5, "太慢": 1.0, "有点慢": 0.6,
    "很慢": 0.9, "等了": 0.3, "迟迟": 0.7, "物流慢": 0.9,
    "推脱": 1.1, "态度差": 1.1, "不回复": 1.0, "敷衍": 0.9,
    "踢皮球": 1.0, "贵": 0.5, "不值": 0.9, "一般": 0.4,
    "普通": 0.3, "勉强": 0.5, "闷": 0.6, "扎": 0.5,
    "勒": 0.6, "粗糙": 0.8, "薄": 0.3, "不足": 0.4,
    "遗憾": 0.4, "问题": 0.5,
    # 否定式负面（避免“不满意”被拆成“满意”）
    "不好": 0.8, "不舒服": 0.8, "不满意": 1.0, "不推荐": 0.9,
    "不划算": 0.8, "不值当": 0.8,
}

SENTIMENT_LEXICON: dict[str, float] = {
    **POSITIVE_LEXICON,
    **{word: -weight for word, weight in NEGATIVE_LEXICON.items()},
}

TOPIC_KEYWORDS: dict[str, list[str]] = {
    "质量做工": ["质量", "做工", "面料", "材质", "起球", "变形", "掉色", "线头",
                 "走线", "手感", "厚薄", "开线", "缩水", "褪色", "破损"],
    "尺码版型": ["尺码", "偏大", "偏小", "大小", "版型", "合身", "长度", "胸围",
                 "肩宽", "修身", "宽松", "买大", "买小"],
    "物流包装": ["物流", "快递", "发货", "包装", "配送", "签收", "速度",
                 "等待", "五天", "第二天"],
    "客服售后": ["客服", "售后", "退货", "换货", "退款", "回复", "态度", "推脱",
                 "运费", "补发", "维权"],
    "价格性价比": ["价格", "性价比", "便宜", "贵", "划算", "优惠", "促销", "券"],
    "外观颜色": ["颜色", "色差", "款式", "好看", "颜值", "图片", "实物", "版型好看"],
    "穿着体验": ["透气", "闷汗", "舒适", "舒服", "亲肤", "扎", "勒", "夏天",
                 "日常", "百搭", "凉快"],
}

ACTION_MAP: dict[str, str] = {
    "质量做工": "针对起球、线头、变形等质量问题，联系供应商做批次质检，并在详情页补充面料成分与洗涤说明。",
    "尺码版型": "在详情页增加真人试穿与尺码表，明确标注偏大/偏小情况，降低因尺码产生的退货。",
    "物流包装": "与仓配核对发货时效，优化包装防护，必要时更换快递合作方并同步时效预期。",
    "客服售后": "复盘退换货话术与运费规则，给客服授权小额补偿，避免售后问题升级为差评。",
    "价格性价比": "评估促销力度与赠品策略，在详情页强化性价比对比，减少“不值”类负面评价。",
    "外观颜色": "统一模特图与实物色差，补充自然光实拍和色卡说明，降低“图文不符”的预期落差。",
    "穿着体验": "补充场景化穿着说明（透气性、适用季节、搭配建议），把体验卖点前置表达。",
}

COLUMN_ALIASES: dict[str, list[str]] = {
    "id": ["评论id", "评价id", "id", "编号", "序号"],
    "product": ["商品名称", "商品", "产品名称", "product"],
    "user": ["买家昵称", "昵称", "用户", "买家", "user"],
    "rating": ["评分", "星级", "分数", "rating", "star"],
    "content": ["评论内容", "评价内容", "内容", "评论", "评价", "content", "comment"],
    "time": ["评论时间", "评价时间", "时间", "日期", "date", "time"],
}

CLAUSE_SPLIT_RE = re.compile(r"[。！？!?；;\n，,]+")
NEGATION_MARKERS = ("不是", "没有", "不太", "并不", "不", "没", "未", "无", "别")


# ---------------------------------------------------------------------------
# 文本与情感工具函数
# ---------------------------------------------------------------------------

def _squash(value: float) -> float:
    """把无界情感得分压缩到 (-1, 1)。"""
    return value / (abs(value) + 2.0)


def _truncate(text: str, limit: int) -> str:
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"-?\d+(?:\.\d+)?", str(value))
    return float(match.group()) if match else None


def _is_negated(text: str, index: int) -> bool:
    prefix = text[max(0, index - 4): index]
    return any(marker in prefix for marker in NEGATION_MARKERS)


def _match_terms(text: str, lexicon: dict[str, float]) -> list[dict[str, Any]]:
    """最长匹配情感词，并处理“不/没/无”等否定前缀，避免重复计分。"""
    matches: list[dict[str, Any]] = []
    occupied = [False] * len(text)
    for phrase in sorted(lexicon, key=len, reverse=True):
        start = 0
        while True:
            index = text.find(phrase, start)
            if index < 0:
                break
            end = index + len(phrase)
            if any(occupied[index:end]):
                start = index + 1
                continue
            weight = lexicon[phrase]
            negated = _is_negated(text, index)
            matches.append({
                "phrase": phrase,
                "index": index,
                "weight": -weight if negated else weight,
                "negated": negated,
            })
            for pos in range(index, end):
                occupied[pos] = True
            start = end
    matches.sort(key=lambda item: item["index"])
    return matches


def _split_clauses(text: str) -> list[str]:
    return [part.strip() for part in CLAUSE_SPLIT_RE.split(text) if part.strip()]


def _detect_topics(text: str) -> list[str]:
    topics = []
    for topic, keywords in TOPIC_KEYWORDS.items():
        if any(keyword in text for keyword in keywords):
            topics.append(topic)
    return topics


def _topic_evidence(text: str) -> dict[str, dict[str, Any]]:
    """按分句把情感归因到主题，得到“质量做工：负面”这类细粒度结论。"""
    buckets: dict[str, dict[str, Any]] = {}
    for clause in _split_clauses(text):
        topics = _detect_topics(clause)
        if not topics:
            continue
        clause_matches = _match_terms(clause, SENTIMENT_LEXICON)
        clause_score = sum(item["weight"] for item in clause_matches)
        for topic in topics:
            bucket = buckets.setdefault(topic, {"net": 0.0, "evidence": []})
            bucket["net"] += clause_score
            if clause_matches and clause not in bucket["evidence"]:
                bucket["evidence"].append(clause)

    aspects = []
    for topic, bucket in buckets.items():
        net = bucket["net"]
        polarity = "正面" if net > 0.2 else "负面" if net < -0.2 else "中性"
        aspects.append({
            "topic": topic,
            "polarity": polarity,
            "score": round(net, 2),
            "evidence": bucket["evidence"][:2],
        })
    aspects.sort(key=lambda item: (item["polarity"] != "负面", -abs(item["score"])))
    return {item["topic"]: item for item in aspects}


def _comment_summary(sentiment: str, aspects: list[dict[str, Any]]) -> str:
    positive = [item["topic"] for item in aspects if item["polarity"] == "正面"]
    negative = [item["topic"] for item in aspects if item["polarity"] == "负面"]
    if positive and negative:
        return f"整体{sentiment}：认可{'、'.join(positive)}，但对{'、'.join(negative)}有负面反馈。"
    if negative:
        return f"整体{sentiment}：主要不满集中在{'、'.join(negative)}。"
    if positive:
        return f"整体{sentiment}：对{'、'.join(positive)}评价较好。"
    return f"整体{sentiment}，未提及明确主题。"


# ---------------------------------------------------------------------------
# 智能体
# ---------------------------------------------------------------------------

class CommentAgent:
    """一个最小可用的 Tool Calling Agent。

    工具共享同一个会话状态：read_comments 把数据放进内存，analyze_comments
    基于内存数据生成报告，save_report 负责落盘。这样大模型只需要传递少量参数，
    不必把整张评论表在上下文里来回搬运。
    """

    def __init__(
        self,
        input_path: str | Path = DEFAULT_INPUT,
        output_dir: str | Path = DEFAULT_OUTPUT_DIR,
        model: str | None = None,
        force_offline: bool = False,
        quiet: bool = False,
    ) -> None:
        self.input_path = Path(input_path)
        self.output_dir = Path(output_dir)
        self.model = model
        self.force_offline = force_offline
        self.quiet = quiet

        self.comments: list[dict[str, Any]] = []
        self.report: dict[str, Any] | None = None
        self.saved_files: dict[str, str] = {}
        self.client: Any = None
        self.mode = "offline"
        self.offline_reason = ""
        self.trace: list[dict[str, Any]] = []

        self.handlers = {
            "read_comments": self.read_comments,
            "analyze_comments": self.analyze_comments,
            "save_report": self.save_report,
        }

    # ------------------------------ 输出 ------------------------------

    def _info(self, message: str) -> None:
        if not self.quiet:
            print(message, flush=True)

    # ------------------------------ 工具 ------------------------------

    def read_comments(self, input_path: str | None = None, **_: Any) -> dict[str, Any]:
        """读取 Excel 评论表，返回标准化后的评论列表。"""
        if not OPENPYXL_AVAILABLE:
            return {"ok": False, "error": "缺少 openpyxl，请执行：pip install -r requirements.txt"}

        path = Path(input_path) if input_path else self.input_path
        path = path.expanduser()
        if not path.is_absolute():
            path = (Path.cwd() / path).resolve()
        if not path.exists():
            return {"ok": False, "error": f"文件不存在：{path}"}
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            return {"ok": False, "error": f"仅支持 .xlsx/.xlsm 文件，收到：{path.suffix}"}

        try:
            workbook = load_workbook(path, read_only=True, data_only=True)
        except Exception as exc:
            return {"ok": False, "error": f"无法打开 Excel：{exc}"}

        try:
            sheet = workbook.active
            rows = [list(row) for row in sheet.iter_rows(values_only=True)]
        finally:
            workbook.close()

        if not rows:
            return {"ok": False, "error": "表格为空，未读取到任何内容"}

        header_index, mapping = -1, {}
        for index, row in enumerate(rows[:5]):
            candidate = self._map_headers(row)
            if "content" in candidate and len(candidate) >= 3:
                header_index, mapping = index, candidate
                break
        if header_index < 0:
            return {
                "ok": False,
                "error": "未识别到表头，请确认包含“评论ID/评分/评论内容”等列",
                "headers": [str(cell) for cell in rows[0] if cell is not None],
            }

        comments: list[dict[str, Any]] = []
        for offset, row in enumerate(rows[header_index + 1:], start=1):
            content = self._cell(row, mapping.get("content"))
            if not str(content).strip():
                continue
            rating = _to_float(self._cell(row, mapping.get("rating")))
            comments.append({
                "id": str(self._cell(row, mapping.get("id")) or f"C{offset:03d}").strip(),
                "product": str(self._cell(row, mapping.get("product")) or "").strip(),
                "user": str(self._cell(row, mapping.get("user")) or "").strip(),
                "rating": int(rating) if rating is not None and rating.is_integer() else rating,
                "content": str(content).strip(),
                "time": self._format_time(self._cell(row, mapping.get("time"))),
            })

        if not comments:
            return {"ok": False, "error": "表头识别成功，但没有读到有效评论内容"}

        self.comments = comments
        self.input_path = path
        return {
            "ok": True,
            "file": str(path),
            "total": len(comments),
            "columns": sorted(mapping.keys()),
            "preview": [f'{item["id"]} | {item["rating"]}星 | {_truncate(item["content"], 28)}'
                        for item in comments[:3]],
        }

    def analyze_comments(self, **_: Any) -> dict[str, Any]:
        """对已读取的评论做语义分析，生成结构化报告并存入会话状态。"""
        if not self.comments:
            return {"ok": False, "error": "还没有评论数据，请先调用 read_comments"}

        analyzed = [self._analyze_comment(item) for item in self.comments]
        summary = self._build_summary(analyzed)
        self.report = {
            "meta": {
                "report_title": "商品评论语义分析报告",
                "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "source_file": str(self.input_path),
                "total_comments": len(analyzed),
                "agent_mode": "LLM Tool Calling Agent" if self.mode == "llm" else "Offline Rule Agent",
                "model": self.model if self.mode == "llm" else "内置规则引擎",
            },
            "summary": summary,
            "comments": analyzed,
        }

        return {
            "ok": True,
            "total_comments": summary["total_comments"],
            "average_rating": summary["average_rating"],
            "sentiment_distribution": summary["sentiment_distribution"],
            "risk_level": summary["risk_level"],
            "top_issues": [
                {"topic": item["topic"], "negative_mentions": item["negative_mentions"]}
                for item in summary["top_issues"]
            ],
            "action_items": summary["action_items"],
            "next_step": "报告已生成，请调用 save_report 保存 JSON 和 Markdown 文件",
        }

    def save_report(self, output_dir: str | None = None, title: str | None = None, **_: Any) -> dict[str, Any]:
        """把会话中的报告保存为 JSON（结构化）和 Markdown（可读）。"""
        if not self.report:
            return {"ok": False, "error": "没有可保存的报告，请先调用 analyze_comments"}

        directory = Path(output_dir) if output_dir else self.output_dir
        directory = directory.expanduser()
        if not directory.is_absolute():
            directory = (Path.cwd() / directory).resolve()
        directory.mkdir(parents=True, exist_ok=True)

        if title:
            self.report["meta"]["report_title"] = title

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"comment_analysis_{stamp}"
        json_path = directory / f"{base_name}.json"
        markdown_path = directory / f"{base_name}.md"
        suffix = 1
        while json_path.exists() or markdown_path.exists():
            suffix += 1
            json_path = directory / f"{base_name}_{suffix}.json"
            markdown_path = directory / f"{base_name}_{suffix}.md"

        json_path.write_text(
            json.dumps(self.report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        markdown_path.write_text(self._render_markdown(self.report), encoding="utf-8-sig")

        self.saved_files = {"json": str(json_path), "markdown": str(markdown_path)}
        return {
            "ok": True,
            "files": self.saved_files,
            "bytes": {
                "json": json_path.stat().st_size,
                "markdown": markdown_path.stat().st_size,
            },
            "message": "报告已保存",
        }

    # ------------------------------ 工具分发 ------------------------------

    def _dispatch(self, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        started = time.perf_counter()
        handler = self.handlers.get(name)
        if handler is None:
            result = {"ok": False, "error": f"未知工具：{name}"}
        else:
            try:
                result = handler(**arguments)
            except TypeError as exc:
                result = {"ok": False, "error": f"参数不匹配：{exc}"}
            except Exception as exc:  # 工具异常不能打断智能体循环
                result = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}

        elapsed = round((time.perf_counter() - started) * 1000, 1)
        self.trace.append({
            "tool": name,
            "arguments": arguments,
            "ok": bool(result.get("ok")),
            "elapsed_ms": elapsed,
        })
        status = "ok" if result.get("ok") else "error"
        args_preview = _truncate(json.dumps(arguments, ensure_ascii=False), 90)
        self._info(f"  [tool] {name}({args_preview}) -> {status} ({elapsed} ms)")
        if not result.get("ok"):
            self._info(f"         原因：{_truncate(result.get('error', ''), 140)}")
        return result

    # ------------------------------ 两种执行模式 ------------------------------

    def run(self) -> dict[str, Any]:
        self._load_dotenv()
        if not self.force_offline:
            self.client = self._build_client()

        if self.client is not None:
            self.mode = "llm"
            try:
                return self._run_llm()
            except Exception as exc:
                self._info(f"\n[agent] 大模型调用失败：{type(exc).__name__}: {exc}")
                self._info("[agent] 自动切换为离线规则模式，继续完成任务。\n")
                self.mode = "offline"
                self.client = None
                self.report = None
                self.saved_files = {}
                self.trace.clear()
        else:
            self._info(f"[agent] {self.offline_reason or '未配置大模型，使用离线规则模式'}")

        return self._run_offline()

    def _run_offline(self) -> dict[str, Any]:
        if not OPENPYXL_AVAILABLE:
            raise RuntimeError("缺少 openpyxl，请先执行：pip install -r requirements.txt")

        self._info("[agent] 开始执行：read_comments -> analyze_comments -> save_report")
        read_result = self._dispatch("read_comments", {"input_path": str(self.input_path)})
        if not read_result.get("ok"):
            raise RuntimeError(read_result.get("error", "读取失败"))

        analysis_result = self._dispatch("analyze_comments", {})
        if not analysis_result.get("ok"):
            raise RuntimeError(analysis_result.get("error", "分析失败"))

        save_result = self._dispatch("save_report", {"output_dir": str(self.output_dir)})
        if not save_result.get("ok"):
            raise RuntimeError(save_result.get("error", "保存失败"))

        return {
            "ok": True,
            "mode": "offline",
            "final_message": self._compose_summary_text(),
            "files": self.saved_files,
        }

    def _run_llm(self) -> dict[str, Any]:
        self._info(f"[agent] 大模型模式：{self.model}")
        system_prompt = (
            "你是一个电商评论分析智能体。你的职责是通过调用工具完成分析，而不是凭空回答。\n"
            "工具说明：\n"
            "1. read_comments：读取 Excel 评论表；\n"
            "2. analyze_comments：对已读取的评论做语义分析（情感、主题、风险、建议），"
            "必须在 read_comments 之后调用；\n"
            "3. save_report：保存 JSON 与 Markdown 报告，必须在 analyze_comments 之后调用。\n"
            "规则：\n"
            "- 严格按 read_comments -> analyze_comments -> save_report 的顺序调用；\n"
            "- 不要伪造、补全或修改评论数据；\n"
            "- 工具报错时，根据错误信息修正参数后重试；\n"
            "- 报告保存成功后，用中文输出 5 行以内的运营摘要（风险等级、主要问题、优先动作），不要输出 JSON。"
        )
        user_prompt = (
            f"请分析评论文件：{self.input_path}\n"
            f"报告输出目录：{self.output_dir}\n"
            "请调用工具完成读取、分析和保存，然后给出运营摘要。"
        )
        messages: list[dict[str, Any]] = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]

        final_text = ""
        for step in range(1, MAX_AGENT_STEPS + 1):
            response = self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                tools=TOOLS,
                tool_choice="auto",
                temperature=0.2,
            )
            message = response.choices[0].message

            if message.tool_calls:
                messages.append({
                    "role": "assistant",
                    "content": message.content or "",
                    "tool_calls": [
                        {
                            "id": call.id,
                            "type": "function",
                            "function": {
                                "name": call.function.name,
                                "arguments": call.function.arguments or "{}",
                            },
                        }
                        for call in message.tool_calls
                    ],
                })
                for call in message.tool_calls:
                    try:
                        arguments = json.loads(call.function.arguments or "{}")
                        if not isinstance(arguments, dict):
                            raise ValueError("arguments 必须是 JSON 对象")
                    except Exception as exc:
                        result = {"ok": False, "error": f"工具参数不是合法 JSON：{exc}"}
                        self._info(f"  [tool] {call.function.name} -> error (bad json)")
                    else:
                        result = self._dispatch(call.function.name, arguments)
                    messages.append({
                        "role": "tool",
                        "tool_call_id": call.id,
                        "content": json.dumps(result, ensure_ascii=False),
                    })
                continue

            final_text = (message.content or "").strip()
            break

        # 模型可能提前结束或漏调工具，这里做兜底，保证报告一定落盘。
        if not self.report:
            self._info("[agent] 模型未完成分析，自动补齐分析步骤。")
            self._dispatch("analyze_comments", {})
        if not self.saved_files:
            self._info("[agent] 模型未保存报告，自动补齐保存步骤。")
            self._dispatch("save_report", {"output_dir": str(self.output_dir)})

        if not final_text:
            final_text = self._compose_summary_text()

        return {
            "ok": True,
            "mode": "llm",
            "final_message": final_text,
            "files": self.saved_files,
        }

    # ------------------------------ LLM 配置 ------------------------------

    def _load_dotenv(self) -> None:
        try:
            from dotenv import load_dotenv
        except ImportError:
            return
        load_dotenv(BASE_DIR / ".env")

    def _build_client(self) -> Any:
        try:
            from openai import OpenAI
        except ImportError:
            self.offline_reason = "未安装 openai，使用离线规则模式（pip install -r requirements.txt 可启用大模型）"
            return None

        api_key = os.getenv("OPENAI_API_KEY", "").strip()
        base_url = os.getenv("OPENAI_BASE_URL", "").strip() or None
        default_model = "gpt-4o-mini"

        if not api_key:
            deepseek_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
            if deepseek_key:
                api_key = deepseek_key
                base_url = base_url or "https://api.deepseek.com"
                default_model = "deepseek-chat"

        if not api_key:
            self.offline_reason = (
                "未检测到 OPENAI_API_KEY，使用离线规则模式"
                "（配置环境变量后即可启用大模型工具调用）"
            )
            return None

        self.model = self.model or os.getenv("OPENAI_MODEL") or default_model
        return OpenAI(api_key=api_key, base_url=base_url, timeout=60.0, max_retries=2)

    # ------------------------------ 分析实现 ------------------------------

    def _analyze_comment(self, comment: dict[str, Any]) -> dict[str, Any]:
        content = str(comment.get("content") or "")
        matches = _match_terms(content, SENTIMENT_LEXICON)
        raw_score = sum(item["weight"] for item in matches)
        rating = _to_float(comment.get("rating"))
        rating_prior = ((rating - 3.0) / 2.0) if rating is not None else 0.0
        score = round(0.65 * _squash(raw_score) + 0.35 * rating_prior, 3)
        sentiment = "正面" if score >= 0.2 else "负面" if score <= -0.2 else "中性"

        aspect_map = _topic_evidence(content)
        aspects = list(aspect_map.values())

        keywords: list[str] = []
        for item in sorted(matches, key=lambda entry: -abs(entry["weight"])):
            if item["phrase"] not in keywords:
                keywords.append(item["phrase"])
        keywords = keywords[:6]

        severe_words = ("踩雷", "投诉", "维权", "欺骗", "垃圾", "严重", "太差", "退货", "退款", "推脱")
        high_risk = (rating is not None and rating <= 2) or any(word in content for word in severe_words)
        risk = "高" if high_risk else "中" if sentiment == "负面" else "低"

        return {
            **comment,
            "sentiment": sentiment,
            "sentiment_score": score,
            "risk": risk,
            "topics": [item["topic"] for item in aspects],
            "aspects": aspects,
            "keywords": keywords,
            "summary": _comment_summary(sentiment, aspects),
        }

    def _build_summary(self, comments: list[dict[str, Any]]) -> dict[str, Any]:
        total = len(comments)
        ratings = [item["rating"] for item in comments if isinstance(item["rating"], (int, float))]
        average_rating = round(sum(ratings) / len(ratings), 2) if ratings else 0.0

        distribution = {"正面": 0, "中性": 0, "负面": 0}
        for item in comments:
            distribution[item["sentiment"]] += 1

        topic_acc: dict[str, dict[str, int]] = {}
        for item in comments:
            for aspect in item["aspects"]:
                acc = topic_acc.setdefault(aspect["topic"], {"mentions": 0, "正面": 0, "中性": 0, "负面": 0})
                acc["mentions"] += 1
                acc[aspect["polarity"]] += 1

        topic_stats = []
        for topic, acc in topic_acc.items():
            topic_stats.append({
                "topic": topic,
                "mentions": acc["mentions"],
                "positive_mentions": acc["正面"],
                "neutral_mentions": acc["中性"],
                "negative_mentions": acc["负面"],
                "negative_ratio": round(acc["负面"] / acc["mentions"], 2) if acc["mentions"] else 0.0,
            })
        topic_stats.sort(key=lambda item: (-item["negative_mentions"], -item["mentions"]))

        top_issues = []
        for stat in topic_stats:
            if stat["negative_mentions"] <= 0:
                continue
            samples = []
            ranked = sorted(
                comments,
                key=lambda item: (item["sentiment"] != "负面", item["risk"] != "高"),
            )
            for item in ranked:
                if any(a["topic"] == stat["topic"] and a["polarity"] == "负面" for a in item["aspects"]):
                    samples.append(_truncate(item["content"], 62))
                if len(samples) >= 2:
                    break
            top_issues.append({**stat, "samples": samples})

        highlights = [
            {"comment_id": item["id"], "text": _truncate(item["content"], 62)}
            for item in comments if item["sentiment"] == "正面"
        ][:3]

        negative_ratio = distribution["负面"] / total if total else 0.0
        severe = any(item["risk"] == "高" for item in comments)
        if negative_ratio >= 0.4 or (severe and negative_ratio >= 0.2):
            risk_level = "高"
        elif negative_ratio >= 0.2 or severe or any(item["sentiment"] == "负面" for item in comments):
            risk_level = "中"
        else:
            risk_level = "低"

        action_items: list[str] = []
        if risk_level == "高":
            action_items.append(
                "优先在 24 小时内回访评分 ≤2 或强烈负面的买家，妥善处理退换货，避免差评扩散。"
            )
        for stat in topic_stats:
            if stat["negative_mentions"] > 0 and stat["topic"] in ACTION_MAP:
                action_items.append(ACTION_MAP[stat["topic"]])
        if not action_items:
            action_items.append("整体口碑稳定，建议继续保持，并在详情页沉淀优质买家评价。")
        action_items = list(dict.fromkeys(action_items))[:6]

        keyword_counter: Counter[str] = Counter()
        for item in comments:
            keyword_counter.update(item["keywords"])

        positive = distribution["正面"]
        overview = (
            f"本次共分析 {total} 条评论，平均评分 {average_rating} 分；"
            f"正面 {positive} 条、中性 {distribution['中性']} 条、负面 {distribution['负面']} 条，"
            f"整体风险等级为「{risk_level}」。"
        )
        if top_issues:
            overview += f"负面反馈最集中的主题是：{'、'.join(item['topic'] for item in top_issues[:3])}。"

        return {
            "total_comments": total,
            "average_rating": average_rating,
            "sentiment_distribution": distribution,
            "sentiment_ratio": {
                key: round(value / total, 2) if total else 0.0
                for key, value in distribution.items()
            },
            "negative_ratio": round(negative_ratio, 2),
            "risk_level": risk_level,
            "overview": overview,
            "topic_stats": topic_stats,
            "top_issues": top_issues,
            "highlights": highlights,
            "action_items": action_items,
            "keywords": [word for word, _ in keyword_counter.most_common(10)],
        }

    @staticmethod
    def _map_headers(row: list[Any]) -> dict[str, int]:
        cells = [
            str(value or "").strip().lower().replace(" ", "").replace("_", "")
            for value in row
        ]
        mapping: dict[str, int] = {}

        # 先做精确匹配，避免“评论”误匹配到“评论ID”
        for field, aliases in COLUMN_ALIASES.items():
            for index, text in enumerate(cells):
                if text and text in aliases:
                    mapping[field] = index
                    break

        # 再做包含匹配，兼容“评论内容（必填）”这类表头
        for field, aliases in COLUMN_ALIASES.items():
            if field in mapping:
                continue
            for index, text in enumerate(cells):
                if index in mapping.values():
                    continue
                if any(len(alias) >= 2 and alias in text for alias in aliases):
                    mapping[field] = index
                    break
        return mapping

    @staticmethod
    def _cell(row: list[Any], index: int | None) -> Any:
        if index is None or index >= len(row):
            return None
        return row[index]

    @staticmethod
    def _format_time(value: Any) -> str:
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d %H:%M")
        return str(value or "").strip()

    def _compose_summary_text(self) -> str:
        if not self.report:
            return "分析未完成。"
        summary = self.report["summary"]
        distribution = summary["sentiment_distribution"]
        lines = [
            f"分析完成：共 {summary['total_comments']} 条评论，平均评分 {summary['average_rating']}；"
            f"正面 {distribution['正面']} / 中性 {distribution['中性']} / 负面 {distribution['负面']}，"
            f"风险等级「{summary['risk_level']}」。",
        ]
        if summary["top_issues"]:
            issues = "；".join(
                f"{item['topic']}（负面 {item['negative_mentions']} 次）"
                for item in summary["top_issues"][:3]
            )
            lines.append(f"主要问题：{issues}。")
        if summary["action_items"]:
            lines.append(f"优先动作：{summary['action_items'][0]}")
        if self.saved_files:
            lines.append(f"结构化报告：{self.saved_files.get('json', '')}")
            lines.append(f"可读报告：{self.saved_files.get('markdown', '')}")
        return "\n".join(lines)

    @staticmethod
    def _render_markdown(report: dict[str, Any]) -> str:
        meta = report["meta"]
        summary = report["summary"]
        dist = summary["sentiment_distribution"]

        def esc(value: Any) -> str:
            return str(value).replace("|", "\\|").replace("\n", " ")

        lines = [
            f"# {meta['report_title']}",
            "",
            f"- 生成时间：{meta['generated_at']}",
            f"- 数据来源：`{meta['source_file']}`",
            f"- 样本数量：{meta['total_comments']} 条",
            f"- 分析模式：{meta['agent_mode']}",
            f"- 分析引擎：{meta['model']}",
            "",
            "## 一、核心结论",
            "",
            summary["overview"],
            "",
            f"**风险等级：{summary['risk_level']}**",
            "",
            "## 二、情感分布",
            "",
            "| 情感 | 数量 | 占比 |",
            "| --- | ---: | ---: |",
        ]
        for label in ("正面", "中性", "负面"):
            lines.append(
                f"| {label} | {dist[label]} | {summary['sentiment_ratio'][label] * 100:.0f}% |"
            )

        lines += [
            "",
            "## 三、主题洞察",
            "",
            "| 主题 | 提及次数 | 正面 | 中性 | 负面 | 负面占比 |",
            "| --- | ---: | ---: | ---: | ---: | ---: |",
        ]
        for item in summary["topic_stats"]:
            lines.append(
                f"| {item['topic']} | {item['mentions']} | {item['positive_mentions']} | "
                f"{item['neutral_mentions']} | {item['negative_mentions']} | "
                f"{item['negative_ratio'] * 100:.0f}% |"
            )

        lines += ["", "### 主要问题", ""]
        if summary["top_issues"]:
            for item in summary["top_issues"]:
                lines.append(
                    f"- **{item['topic']}**：负面提及 {item['negative_mentions']} 次，"
                    f"负面占比 {item['negative_ratio'] * 100:.0f}%。"
                )
                for sample in item["samples"]:
                    lines.append(f"  > {sample}")
        else:
            lines.append("- 暂未发现集中性的负面主题。")

        lines += ["", "### 用户认可点", ""]
        if summary["highlights"]:
            for item in summary["highlights"]:
                lines.append(f"- `{item['comment_id']}`：{item['text']}")
        else:
            lines.append("- 暂无明显正面原声。")

        lines += [
            "",
            "## 四、评论明细",
            "",
            "| 评论ID | 评分 | 情感 | 风险 | 涉及主题 | 语义摘要 |",
            "| --- | ---: | --- | --- | --- | --- |",
        ]
        for item in report["comments"]:
            lines.append(
                f"| {esc(item['id'])} | {esc(item['rating'] if item['rating'] is not None else '-')} | "
                f"{item['sentiment']} | {item['risk']} | "
                f"{esc('、'.join(item['topics']) or '-')} | {esc(item['summary'])} |"
            )

        lines += ["", "## 五、运营建议", ""]
        for index, action in enumerate(summary["action_items"], start=1):
            lines.append(f"{index}. {action}")

        lines += [
            "",
            "---",
            "",
            "> 本报告由评论分析智能体自动生成；语义结论基于评论原文，建议结合业务实际复核后执行。",
            "",
        ]
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# 工具声明（供大模型 function calling 使用）
# ---------------------------------------------------------------------------

TOOLS: list[dict[str, Any]] = [
    {
        "type": "function",
        "function": {
            "name": "read_comments",
            "description": "读取 Excel 评论表，识别表头并返回标准化评论数据。默认读取 comments.xlsx。",
            "parameters": {
                "type": "object",
                "properties": {
                    "input_path": {
                        "type": "string",
                        "description": "Excel 文件路径，支持 .xlsx 或 .xlsm；不传则使用默认路径",
                    }
                },
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "analyze_comments",
            "description": (
                "对已读取到内存中的评论做语义分析，输出情感倾向、主题归因、风险等级和运营建议。"
                "调用前必须先执行 read_comments。"
            ),
            "parameters": {
                "type": "object",
                "properties": {},
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "save_report",
            "description": (
                "把当前分析结果保存为 JSON 结构化报告和 Markdown 可读报告。"
                "调用前必须先执行 analyze_comments。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "output_dir": {
                        "type": "string",
                        "description": "报告输出目录，默认 reports/",
                    },
                    "title": {
                        "type": "string",
                        "description": "报告标题，可选",
                    },
                },
                "required": [],
            },
        },
    },
]


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def _init_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # pragma: no cover - 仅老版本控制台会触发
            pass


def main() -> int:
    _init_console()
    parser = argparse.ArgumentParser(description="电商评论语义分析智能体（Tool Calling Agent）")
    parser.add_argument("-f", "--file", default=str(DEFAULT_INPUT), help="评论 Excel 路径")
    parser.add_argument("-o", "--output", default=str(DEFAULT_OUTPUT_DIR), help="报告输出目录")
    parser.add_argument("--model", default=None, help="大模型名称，默认读取 OPENAI_MODEL")
    parser.add_argument("--offline", action="store_true", help="强制使用离线规则模式")
    parser.add_argument("--quiet", action="store_true", help="不打印工具调用过程")
    args = parser.parse_args()

    agent = CommentAgent(
        input_path=args.file,
        output_dir=args.output,
        model=args.model,
        force_offline=args.offline,
        quiet=args.quiet,
    )

    try:
        result = agent.run()
    except Exception as exc:
        print(f"\n[agent] 执行失败：{exc}", file=sys.stderr)
        if not args.quiet:
            print("[agent] 请检查输入文件格式与依赖安装情况。", file=sys.stderr)
        return 1

    print("\n" + "=" * 68)
    print(result["final_message"])
    print("=" * 68)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
