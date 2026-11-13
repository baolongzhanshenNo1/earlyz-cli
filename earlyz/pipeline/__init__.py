"""
EarlyZ Pipeline — Agent 闭环排版流水线

流程:
    SpecReader(规范模板) → DocDiscover(目标文档) → Reconcile → Writer → Fixer 诊断
    ↓
    Fixer 有残留 → 用户反馈 → 更新 fix_plan → 重跑 Writer → Fixer
    （闭环直到 0 残留）

用法:
    from earlyz.pipeline.agent import EarlyZAgent

    agent = EarlyZAgent()
    result = agent.run(
        spec_path=Path('规范.docx'),
        target_path=Path('论文.docx'),
        output_path=Path('论文_修正版.docx'),
    )
    # result: {"plan": [...], "issues": [...], "fixed_count": N, "remaining": M}
"""

from .agent import EarlyZAgent

__all__ = ["EarlyZAgent"]
