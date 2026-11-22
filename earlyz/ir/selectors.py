"""早睡 (EarlyZ) 3.2 — 稳定 NodeSelector 与身份恢复。

对应架构文档 §4.3/§4.4：
- build_node_id：组合 part_name + node_type + semantic_role + 规范化文本 hash
  + parent identity，生成稳定 ID；
- NodeResolver：按优先级恢复身份（精确 ID → 文本/内容 hash → role+occurrence
  → 语义 fallback），多候选必须报 AMBIGUOUS，禁止静默选中；
- legacy_index_to_selector：兼容层，把旧 para_index 转换为 fallback 信息，
  绝不把裸索引当作稳定定位。
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from typing import Optional

from earlyz.ir.models import DocumentIR, DocumentNode, NodeSelector

_WS_RE = re.compile(r"\s+")


class AmbiguousSelectorError(Exception):
    """选择器命中多个候选，必须停止并报告（§4.4）。

    selector 可以是 NodeSelector 对象或 dict（_locate_paragraphs 的
    XML 级定位用 dict 形态，语义等价）。
    """

    def __init__(self, selector, candidates):
        self.selector = selector
        self.candidates = candidates
        if hasattr(selector, "to_dict"):
            _detail = selector.to_dict()
        else:
            _detail = selector
        super().__init__(
            f"AMBIGUOUS_SELECTOR: {len(candidates)} candidates for {_detail}"
        )


def normalize_text(text: str) -> str:
    """规范化文本：去除首尾空白、压缩内部空白、全角空格转半角。"""
    t = text.replace("\u3000", " ").replace("\xa0", " ")
    t = _WS_RE.sub("", t)  # 排版身份不关心词间空白
    return t.strip()


def _stable_hash(s: str, length: int = 16) -> str:
    return hashlib.sha1(s.encode("utf-8")).hexdigest()[:length]


def build_node_id(
    part_name: str,
    node_type: str,
    semantic_role: Optional[str],
    normalized_text_hash: str,
    parent_id: Optional[str],
    extra_fingerprint: Optional[str] = None,
) -> str:
    """按 §4.3 生成稳定 node_id（不使用数组索引）。"""
    role = semantic_role or "unclassified"
    parts = [part_name, node_type, role, normalized_text_hash]
    if parent_id:
        parts.append(parent_id)
    if extra_fingerprint:
        parts.append(extra_fingerprint)
    return f"{node_type}:{role}:{_stable_hash(':'.join(parts))}"


class NodeResolver:
    """在 DocumentIR 中按优先级恢复节点身份。"""

    def __init__(self, ir: DocumentIR):
        self.ir = ir
        seen: set[int] = set()
        self._nodes: list[DocumentNode] = []
        for coll in (ir.blocks, ir.paragraphs, ir.tables, ir.figures,
                     ir.fields, ir.captions, ir.sections):
            for n in coll:
                if id(n) not in seen:  # dataclass == 会把同值对象误判为同一节点
                    seen.add(id(n))
                    self._nodes.append(n)

    # ── 主入口 ──────────────────────────────────────────────────

    def resolve(self, selector: NodeSelector) -> Optional[DocumentNode]:
        # 1) 精确 node_id
        if selector.node_id:
            exact = self._by_node_id(selector.node_id)
            if exact is not None:
                return exact

        # 2) node_type + text_hash（唯一性校验）
        if selector.text_hash:
            by_hash = self._by_text_hash(selector.text_hash, selector.node_type)
            if by_hash:
                if len(by_hash) > 1 and not selector.occurrence:
                    raise AmbiguousSelectorError(selector, by_hash)
                return by_hash[selector.occurrence if selector.occurrence < len(by_hash) else 0]

        # 3) semantic_role + occurrence
        if selector.semantic_role:
            by_role = self._by_role(selector.semantic_role, selector.node_type)
            if by_role:
                if len(by_role) > 1 and not selector.occurrence:
                    raise AmbiguousSelectorError(selector, by_role)
                idx = selector.occurrence
                return by_role[idx] if idx < len(by_role) else None

        # 4) 语义 fallback（normalized_text）
        if selector.fallback and selector.fallback.get("normalized_text"):
            by_text = self._by_normalized_text(selector.fallback["normalized_text"],
                                               selector.node_type)
            if by_text:
                if len(by_text) > 1 and not selector.occurrence:
                    raise AmbiguousSelectorError(selector, by_text)
                return by_text[selector.occurrence if selector.occurrence < len(by_text) else 0]

        # 5) legacy 索引仅作诊断提示，不直接返回（§4.4 兼容层）
        return None

    # ── 内部检索 ────────────────────────────────────────────────

    def _by_node_id(self, node_id: str) -> Optional[DocumentNode]:
        for n in self._nodes:
            if n.node_id == node_id:
                return n
        return None

    def _by_text_hash(self, text_hash: str, node_type: Optional[str]) -> list[DocumentNode]:
        out = []
        for n in self._nodes:
            if n.text_hash == text_hash and (node_type is None or n.node_type == node_type):
                out.append(n)
        return out

    def _by_role(self, role: str, node_type: Optional[str]) -> list[DocumentNode]:
        out = []
        for n in self._nodes:
            if n.semantic_role == role and (node_type is None or n.node_type == node_type):
                out.append(n)
        return out

    def _by_normalized_text(self, text: str, node_type: Optional[str]) -> list[DocumentNode]:
        norm = normalize_text(text)
        out = []
        for n in self._nodes:
            ntext = normalize_text(getattr(n, "text", "") or "")
            if ntext and ntext == norm and (node_type is None or n.node_type == node_type):
                out.append(n)
        return out


def legacy_index_to_selector(
    part_name: str, node_type: str, index: int
) -> NodeSelector:
    """兼容层：旧 (para|tbl|section)_index → NodeSelector。

    索引只进 fallback 供诊断与人工确认，不作为稳定定位依据（§13.2）。
    """
    return NodeSelector(
        node_type=node_type,
        fallback={"part_name": part_name, "legacy_index": index},
    )
