# -*- coding: utf-8 -*-
"""
AI 提取（MiniMax M3）—— 把 PO 文字交给大模型，返回结构化字段 JSON。

配置（任选）：
  - API key：同目录下 .ai_key 文件，或环境变量 MINIMAX_API_KEY
  - 接口/模型：默认 https://api.minimaxi.com/v1 + MiniMax-M3，
              可用环境变量 MINIMAX_BASE_URL / MINIMAX_MODEL 覆盖

对外接口：
  ai_available()        -> bool     是否配了 key（决定要不要走 AI）
  ai_extract_po(text)   -> dict     从 PO 文字抽取字段；失败/未配置返回 {}

返回字段名与表单 / core.FIELD_MAP 完全一致，可直接覆盖到草稿上。
结果仍是“草稿”，务必人工核对（尤其单价/数量/金额）。
"""
import os
import re
import json
import urllib.request
import urllib.error
from pathlib import Path

BASE = Path(__file__).resolve().parent
_KEYFILE = BASE / ".ai_key"

AI_BASE_URL = os.environ.get("MINIMAX_BASE_URL", "https://api.minimaxi.com/v1").rstrip("/")
AI_MODEL = os.environ.get("MINIMAX_MODEL", "MiniMax-M3")
AI_TIMEOUT = int(os.environ.get("MINIMAX_TIMEOUT", "60"))

# 只接受这些字段（与表单一致），其余一律丢弃，防止模型塞垃圾键
ALLOWED_KEYS = {
    "ReferenceNo", "买方名称及地址", "通知方", "品名及HS编码",
    "数量KGS", "单价USD", "付款条款", "贸易术语",
    "装货港", "卸货港", "包装方式",
}

_SYS = "你是严谨的外贸单据信息抽取助手。只输出一个 JSON 对象，不要解释、不要 markdown 代码块、不要多余文字。"

_PROMPT = """从下面这份采购订单(PO)原文里抽取字段，返回严格的 JSON。

规则：
- 只抽取 PO 里**确实出现**的信息；找不到的字段一律填空字符串 ""，绝不要编造。
- 付款条款、贸易术语、品名等照抄 PO 原句，不要改写。
- 数量与单价要换算成统一单位，且只返回**纯数字**（不带单位、不带千分位逗号）：
    · "数量KGS"：换算成公斤。PO 用吨/MT/metric ton 时 ×1000；本来就是 KG/KGS 则原样。
    · "单价USD"：换算成“美元/公斤”。PO 若按每吨计价则 ÷1000；按每公斤则原样。
- 注意标签的各种写法：付款条款可能写成 "Terms of Payment" 或 "Payment Terms"，
  贸易术语在 "Terms of Delivery" 里（如 CIF - Kolkata Port）。

需要的字段（键名固定用下面这些，不要增减键）：
- "ReferenceNo": PO号 / 采购订单号 / 合同号（取真正的订单编号，含数字那个）
- "买方名称及地址": 下单方(买方/收货人)的公司名 + 地址
- "通知方": 通知方 Notify Party（没有就留空）
- "品名及HS编码": 货物名称；若有 HS/HSN 编码，另起一行写 "HS CODE: 编码"
- "数量KGS": 数量(公斤，纯数字)
- "单价USD": 单价(美元/公斤，纯数字)
- "付款条款": 付款条款原句（如 Payable within 45 days due net）
- "贸易术语": 贸易术语+港口（如 CIF Kolkata Port）
- "装货港": 装货港（能判断才填）
- "卸货港": 卸货港（能判断才填）
- "包装方式": 包装方式（如 drum / IBC / ISO tank / bulk 等）

PO 原文：
<<<
{text}
>>>

只输出 JSON。"""


def _api_key() -> str:
    k = os.environ.get("MINIMAX_API_KEY", "").strip()
    if k:
        return k
    if _KEYFILE.exists():
        try:
            return _KEYFILE.read_text(encoding="utf-8").strip()
        except Exception:
            return ""
    return ""


def ai_available() -> bool:
    return bool(_api_key())


def _strip_think(s: str) -> str:
    """剥掉 MiniMax M3 推理模型的 <think>...</think> 段。"""
    return re.sub(r"<think>.*?</think>", "", s, flags=re.S).strip()


def _parse_json(content: str) -> dict:
    """从模型回复里抠出 JSON 对象。"""
    content = _strip_think(content)
    # 去掉可能的 ```json ... ``` 包裹
    content = re.sub(r"^```(?:json)?|```$", "", content.strip(), flags=re.M).strip()
    try:
        return json.loads(content)
    except Exception:
        pass
    # 兜底：取第一个 { 到最后一个 } 之间
    i, j = content.find("{"), content.rfind("}")
    if i != -1 and j != -1 and j > i:
        try:
            return json.loads(content[i:j + 1])
        except Exception:
            return {}
    return {}


def _chat(text: str) -> str:
    """调用 MiniMax，返回模型文本内容。任何失败抛异常给调用方兜底。"""
    key = _api_key()
    if not key:
        raise RuntimeError("未配置 API key")
    body = json.dumps({
        "model": AI_MODEL,
        "messages": [
            {"role": "system", "content": _SYS},
            {"role": "user", "content": _PROMPT.format(text=text[:24000])},
        ],
        "temperature": 0,
    }).encode("utf-8")
    req = urllib.request.Request(
        f"{AI_BASE_URL}/chat/completions",
        data=body,
        headers={
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=AI_TIMEOUT) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return (data.get("choices") or [{}])[0].get("message", {}).get("content", "") or ""


def ai_extract_po(text: str) -> dict:
    """用 AI 从 PO 文字抽取字段。未配置 key / 调用失败 / 解析失败 → 返回 {}。"""
    if not text or not text.strip() or not ai_available():
        return {}
    try:
        raw = _chat(text)
    except Exception:
        return {}
    obj = _parse_json(raw)
    if not isinstance(obj, dict):
        return {}
    out = {}
    for k, v in obj.items():
        if k not in ALLOWED_KEYS:
            continue
        if v in (None, ""):
            continue
        out[k] = v
    return out
