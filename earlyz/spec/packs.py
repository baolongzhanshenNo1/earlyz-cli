"""Template Pack 资产化（§19）：内置规范包元数据发现与版本戳 + 用户模板（C3）。

设计（源自 DeepSeek Harness agent-presets 审计）：
- 发现 = 每次读盘不缓存（运行中新增/删除立即可见）；
- 生效内容（``<id>_pack.json``）与显示元数据（可选侧车 ``<id>_pack.meta.yml``）分离，
  元数据损坏不影响模板可用（display-failure-not-fatal）；
- 信任由根目录决定（ADR-017）：system 根 = 内置不可删；user 根 = 用户自建可删；
- 坏包不隐藏：broken 包留在列表里带原因；
- 版本戳 = 文件戳世代（mtimeMs + size + sha1，ADR-018）。
"""

import hashlib
import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path

_PACKS_DIR = Path(__file__).resolve().parent / "packs"
# 内置规范包 id 白名单（单一真相源：server.py X-Spec-Pack 校验从这导入，禁重复定义）
_BUILTIN_PACK_IDS: frozenset[str] = frozenset({"thesis", "contract", "gov", "bid", "tech"})

# 用户模板根（上传 spec 资产化落点，C3）：<DATA_DIR>/packs_user/*.docx
try:
    from api.host.artifacts import DATA_DIR  # 测试身份（pytest：sys.path=仓库根）
except ImportError:  # 运行身份（agent_loop/server：sys.path=api/）
    from host.artifacts import DATA_DIR
_USER_PACKS = Path(DATA_DIR) / "packs_user"


@dataclass
class PackMeta:
    """一个模板包的显示元数据（id/name/description/信任/规则数/损坏原因）。"""

    id: str
    name: str
    description: str
    domain: str
    rule_count: int
    trust: str              # 'system'（内置根）| 'user'（用户根）——由根目录决定，非自报
    order: int = 0
    broken: str | None = None
    file: str | None = None  # user pack 的真实文件路径（system pack 为 None）


def _read_meta(pack_file: Path) -> dict:
    """可选侧车 <stem>.meta.yml/json（如 thesis_pack.meta.yml）：显示元数据，损坏不影响模板可用。"""
    stem = pack_file.with_suffix("")  # 去掉 .json → thesis_pack
    for suffix in (".meta.yml", ".meta.json"):
        m = stem.with_name(stem.name + suffix)
        if m.exists():
            try:
                if m.suffix == ".yml":
                    import yaml
                    data = yaml.safe_load(m.read_text(encoding="utf-8")) or {}
                else:
                    data = json.loads(m.read_text(encoding="utf-8"))
                return data if isinstance(data, dict) else {}
            except Exception:  # noqa: BLE001 — 元数据损坏回退内置字段
                return {}
    return {}


def _scan_pack_file(p: Path) -> PackMeta:
    """读一个 pack JSON，返回元数据；解析失败 → broken 但保留在列表（坏包不隐藏）。"""
    pid = p.stem.rsplit("_pack", 1)[0]
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:  # noqa: BLE001
        return PackMeta(id=pid, name=p.stem, description=f"模板文件损坏：{e}",
                        domain="", rule_count=0, trust="system", broken=str(e))
    meta = _read_meta(p)
    return PackMeta(
        id=pid,
        name=meta.get("name") or data.get("name") or pid,
        description=meta.get("description") or data.get("description") or "",
        domain=data.get("domain", ""),
        order=meta.get("order", 0),
        rule_count=len(data.get("rules") or []),
        trust="system",
    )


def list_packs() -> list[PackMeta]:
    """内置规范包清单（system 根，每次读盘）。"""
    return sorted((_scan_pack_file(p) for p in _PACKS_DIR.glob("*_pack.json")),
                  key=lambda m: m.order)


def list_packs_dict() -> list[dict]:
    """/api/templates 端点的 JSON 形态（system packs）。"""
    return [asdict(m) for m in list_packs()]


def list_user_packs() -> list[PackMeta]:
    """用户模板清单（user 根，C3）：上传 spec 的资产化落点。

    磁盘文件名 = <sha1前12>.docx（纯 ASCII——pack id 会进 HTTP header，
    含中文会炸 fetch「non ISO-8859-1」）；中文原名存 sidecar meta。
    兼容旧格式 <sha>__<原文件名>.docx：id 归一化为 sha（仍能找到文件）。
    """
    _USER_PACKS.mkdir(parents=True, exist_ok=True)
    out = []
    for p in sorted(_USER_PACKS.glob("*.docx")):
        stem = p.name.rsplit(".docx", 1)[0]
        sha_part = stem.split("__", 1)[0]
        # sidecar 元数据优先（新格式存中文原名）
        display = ""
        sidecar = p.with_suffix(".meta.json")
        if sidecar.exists():
            try:
                display = (json.loads(sidecar.read_text("utf-8")) or {}).get("name", "")
            except Exception:  # noqa: BLE001 — sidecar 损坏回退文件名
                display = ""
        if not display:
            display = p.name.split("__", 1)[1] if "__" in p.name else sha_part
        out.append(PackMeta(
            id=f"user:{sha_part}",
            name=display.rsplit(".docx", 1)[0] if display.endswith(".docx") else display,
            description="用户上传的模板",
            domain="",
            rule_count=0,
            trust="user",
            file=str(p),
        ))
    return out


def pack_stamp(pack_id: str) -> dict:
    """内置规范包版本戳（文件戳世代 + sha1，ADR-018）。未知 id → FileNotFoundError。"""
    p = _PACKS_DIR / f"{pack_id}_pack.json"
    if not p.exists():
        raise FileNotFoundError(f"未知内置规范包: {pack_id}")
    return {
        "mtime_ms": int(p.stat().st_mtime * 1000),
        "size": p.stat().st_size,
        "sha1": hashlib.sha1(p.read_bytes()).hexdigest(),
    }


def find_user_pack_file(pack_id: str) -> Path | None:
    """user:<sha12> → packs_user 内真实文件路径（防路径注入：仅十六进制精确匹配）。

    兼容旧 id user:<sha>__<原名>.docx：截取 sha 段查找。
    """
    if not pack_id.startswith("user:"):
        return None
    sha_part = pack_id[len("user:"):].split("__", 1)[0]
    if not re.fullmatch(r"[0-9a-f]{6,64}", sha_part):
        return None
    p_new = _USER_PACKS / f"{sha_part}.docx"
    if p_new.exists():
        return p_new
    # 旧格式 <sha>__<原文件名>.docx
    hits = [q for q in _USER_PACKS.glob(f"{sha_part}__*.docx")]
    return hits[0] if len(hits) == 1 else None


def user_pack_name(pack_id: str) -> str:
    """user:<sha12> → 侧车里的中文显示名（无侧车回退 sha12；找不到包返回 ""）。

    2026-10-12：会话模板记忆要给 LLM 一份可读名字，原来只有路径。
    """
    f = find_user_pack_file(pack_id)
    if f is None:
        return ""
    meta = _read_meta(f)
    return str(meta.get("name") or f.stem)


def save_user_pack(docx_bytes: bytes, original_name: str) -> dict | None:
    """上传 spec 资产化（C3）：按 sha1 去重，重复上传同文件不重复占盘。

    磁盘文件名 = <sha1前12>.docx（纯 ASCII：pack id 进 HTTP header 不能含
    非 ISO-8859-1 字符，2026-08-15 fetch 崩溃根因）；中文原名存 sidecar。
    返回 pack 元数据 dict（含 id），失败返回 None（不阻断排版主流程）。
    """
    try:
        _USER_PACKS.mkdir(parents=True, exist_ok=True)
        sha = hashlib.sha1(docx_bytes).hexdigest()
        dest = _USER_PACKS / f"{sha[:12]}.docx"
        created = not dest.exists()
        if created:
            dest.write_bytes(docx_bytes)
        # sidecar：中文显示名（id/文件名不含中文）
        dest.with_suffix(".meta.json").write_text(
            json.dumps({"name": Path(original_name).stem, "sha1": sha}, ensure_ascii=False),
            encoding="utf-8",
        )
        return {
            "id": f"user:{sha[:12]}",
            "name": Path(original_name).stem,
            "trust": "user",
            "file": str(dest),
            "created": created,
        }
    except Exception:  # noqa: BLE001 — 资产化失败不阻断上传
        return None


def delete_user_pack(pack_id: str) -> bool:
    """删除用户模板（user 根可删；system 根拒绝——ADR-017）。

    文件与 sidecar 元数据一起清（交换/替换模板时不留残迹；sidecar 留着会让
    `<sha>.meta.json` 在 packs_user 目录里越积越多）。
    """
    p = find_user_pack_file(pack_id)
    if not p:
        return False
    try:
        p.unlink()
        sidecar = p.with_suffix(".meta.json")
        if sidecar.exists():
            sidecar.unlink()
        return True
    except OSError:
        return False
