"""LLM attribution: prompt building, tolerant JSON parsing, and a minimal
HTTP call that reuses kb-qa-eval's api_client wire-format helpers
(``build_request_payload`` / ``extract_response_text``) without its
evaluation-specific scorecard parser.
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Any

_PROJECT_DIR = Path(__file__).resolve().parent.parent.parent / "kb-qa-eval"
if str(_PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(_PROJECT_DIR))

_CRITERIA_PATH = Path(__file__).resolve().parent.parent / "references" / "attribution_criteria.md"

_PROMPTS_DIR = Path(__file__).resolve().parent.parent / "prompts"
if str(_PROMPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_PROMPTS_DIR))

from attribution_prompt import ATTRIBUTION_PROMPT  # noqa: E402


def load_criteria() -> str:
    try:
        return _CRITERIA_PATH.read_text(encoding="utf-8")
    except OSError:
        return "A=检索召回问题；B=大模型定位问题；C=知识库无答案。"


def build_attribution_prompt(
    *,
    query: str,
    filename: str,
    annotation: str,
    retrieval: str,
    coverage_text: str,
    answer: str,
    wa: float,
    dim_reasons: str,
    original_text: str = "",
) -> str:
    return ATTRIBUTION_PROMPT.format(
        criteria=load_criteria(),
        query=query,
        filename=filename or "(unknown)",
        annotation=annotation or "(无标注)",
        retrieval=(retrieval or "(无召回)")[:4000],
        coverage_text=coverage_text,
        answer=answer or "(空回答)",
        wa=f"{wa:.2f}",
        dim_reasons=dim_reasons or "(无)",
        original_text=original_text or "(未提供)",
    )


def parse_attribution(content: str) -> dict[str, Any]:
    """Tolerant parse: strip code fences, then extract the first JSON object."""
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", content.strip(), flags=re.MULTILINE)
    text = text.strip()
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end > start:
        text = text[start:end + 1]
    try:
        parsed = json.loads(text)
        if isinstance(parsed, dict):
            return parsed
    except (json.JSONDecodeError, ValueError):
        pass
    return {}


def call_attribution_model(
    prompt: str,
    *,
    model: str,
    base_url: str,
    api_key: str,
    protocol: str = "chat_completions",
    max_tokens: int = 4096,
    retries: int = 2,
    require_verification: bool = False,
) -> dict[str, Any]:
    """Call *model* once (with format retry) and parse the attribution JSON.

    With ``require_verification`` (original-text material was provided), a
    parsed result whose 原文核验 field is empty triggers one targeted retry
    nudging the model to fill the three-state verification verdict.
    """
    import requests

    from scripts.api_client import (
        RESPONSE_MODE_TEXT,
        build_request_payload,
        extract_response_text,
    )

    if protocol == "chat_completions":
        endpoint = "chat/completions"
    elif protocol == "responses":
        endpoint = "responses"
    else:
        endpoint = "messages"
    url = f"{base_url.rstrip('/')}/{endpoint}"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    if protocol == "anthropic_messages":
        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json",
        }

    import time

    last_error: Exception | None = None
    last_parsed: dict[str, Any] | None = None
    rate_limit_retries = 0
    max_rate_limit_retries = 3
    attempt = 0
    while True:
        if attempt == 0:
            retry_prompt = prompt
        elif require_verification and last_parsed is not None and not last_parsed.get("原文核验") and attempt < retries - 1:
            retry_prompt = (
                f"{prompt}\n\n注意：本次已提供原文核验材料，【原文核验】字段不得为空字符串，"
                "请依据原文给出 ✅ 已核实 / ⚠️ 部分核实 / ⚠️ 标注问题 三态结论并指明原文位置。"
                "请只输出 JSON，不要包含任何其他文字。"
            )
        else:
            retry_prompt = f"{prompt}\n\n请只输出 JSON，不要包含任何其他文字。"
        payload = build_request_payload(
            model=model,
            prompt=retry_prompt,
            protocol=protocol,
            response_mode=RESPONSE_MODE_TEXT,
            language="zh",
            max_tokens=max_tokens,
        )
        rate_limited = False
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=240)
            response.raise_for_status()
            content = extract_response_text(response.json(), protocol)
            parsed = parse_attribution(content)
            if parsed.get("归因结论") or parsed.get("归因分析"):
                if require_verification and not parsed.get("原文核验") and attempt < retries - 1:
                    last_parsed = parsed
                    last_error = ValueError("原文核验 empty despite provided material")
                else:
                    return parsed
            else:
                last_parsed = None
                last_error = ValueError("attribution JSON missing expected keys")
        except (requests.exceptions.RequestException, ValueError, json.JSONDecodeError) as error:
            last_error = error
            rate_limited = (
                isinstance(error, requests.exceptions.HTTPError)
                and error.response is not None
                and error.response.status_code == 429
            )
        if rate_limited and rate_limit_retries < max_rate_limit_retries:
            # Rate-limit retries resend the same prompt and do not consume
            # the format/verification retry budget.
            rate_limit_retries += 1
            time.sleep(60)
            continue
        attempt += 1
        if attempt >= retries:
            raise RuntimeError(
                f"Attribution API call failed after {retries} retries "
                f"({rate_limit_retries} rate-limit retries): {last_error}"
            )
        time.sleep(2 ** attempt)


def offline_attribution(coverage_text: str, rate: float, wa: float) -> tuple[str, str]:
    """Template attribution used when no model is configured (--no-llm)."""
    if "无金标标注" in coverage_text:
        verdict = "无法判定（无金标标注）"
        analysis = (
            f"该条目缺少金标标注（{coverage_text}），覆盖率无法客观计算，"
            f"自动归因不可靠（WA {wa:.2f}），建议人工复核。"
        )
    elif "未召回" in coverage_text:
        verdict = "A 召回问题（纯）"
        analysis = (
            f"金标段落未召回（{coverage_text}），模型只能基于召回内容作答；"
            f"回答与金标存在偏差属召回缺失的下游结果（WA {wa:.2f}）。"
        )
    elif "部分命中" in coverage_text:
        verdict = "A主+B次"
        analysis = (
            f"金标段落仅部分命中（{coverage_text}），召回上下文不完整；"
            f"模型在信息缺失下作答（WA {wa:.2f}），属召回缺陷前提下的次责。"
        )
    else:
        verdict = "B 大模型定位问题"
        analysis = (
            f"金标信息已在召回上下文中（{coverage_text}），但回答未正确使用，"
            f"属大模型定位问题（WA {wa:.2f}）。"
        )
    return verdict, analysis
