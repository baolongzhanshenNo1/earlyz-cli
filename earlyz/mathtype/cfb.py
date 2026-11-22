"""OLE 复合文档（CFB，[MS-CFB] v3，512 字节扇区）最小读写：只处理 MathType 公式对象这种
「根存储 + 若干小流」的形态（单层、无子存储）。stdlib only。

读：按 FAT / miniFAT 链取出全部流；写：所有流 < 4096 字节走 mini stream，目录树形状与
MathType 自己写出的对象一致（按 [MS-CFB] 名称比较规则排的二叉树）。
"""
from __future__ import annotations

import struct

MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")
ENDOFCHAIN, FREESECT, FATSECT = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF
SSZ, MSZ, MINI_CUTOFF = 512, 64, 4096


class CfbError(ValueError):
    pass


def read_streams(data: bytes) -> tuple[dict[str, bytes], bytes]:
    """→ ({流名: 内容}, 根存储 CLSID 16 字节)。"""
    if data[:8] != MAGIC:
        raise CfbError("not a compound file")
    ssz = 1 << struct.unpack_from("<H", data, 30)[0]
    mssz = 1 << struct.unpack_from("<H", data, 32)[0]
    n_fat, dir_start = struct.unpack_from("<II", data, 44)
    cutoff, mfat_start = struct.unpack_from("<II", data, 56)[0], struct.unpack_from("<I", data, 60)[0]
    difat_start, n_difat = struct.unpack_from("<II", data, 68)
    difat = list(struct.unpack_from("<109I", data, 76))
    s = difat_start
    while n_difat and s not in (ENDOFCHAIN, FREESECT):
        off = (s + 1) * ssz
        ids = struct.unpack_from(f"<{ssz // 4}I", data, off)
        difat += ids[:-1]
        s, n_difat = ids[-1], n_difat - 1
    difat = difat[:n_fat]

    def sec(i: int) -> bytes:
        off = (i + 1) * ssz
        if off + ssz > len(data) + ssz:      # 末扇区允许不满
            raise CfbError("sector out of range")
        return data[off:off + ssz]

    fat: list[int] = []
    for i in difat:
        fat += struct.unpack(f"<{ssz // 4}I", sec(i).ljust(ssz, b"\0"))

    def chain(start: int, table: list[int]) -> list[int]:
        out, i = [], start
        while i not in (ENDOFCHAIN, FREESECT):
            if i >= len(table) or len(out) > len(table):
                raise CfbError("broken sector chain")
            out.append(i)
            i = table[i]
        return out

    dirb = b"".join(sec(i) for i in chain(dir_start, fat))
    ents = []
    for k in range(len(dirb) // 128):
        e = dirb[k * 128:(k + 1) * 128]
        nlen = struct.unpack_from("<H", e, 64)[0]
        ents.append((e[:max(nlen - 2, 0)].decode("utf-16-le", "replace"), e[66], e[80:96],
                     *struct.unpack_from("<II", e, 116)))
    root = next((x for x in ents if x[1] == 5), None)
    if root is None:
        raise CfbError("no root entry")
    mini = b"".join(sec(i) for i in chain(root[3], fat))
    mfat: list[int] = []
    for i in chain(mfat_start, fat):
        mfat += struct.unpack(f"<{ssz // 4}I", sec(i).ljust(ssz, b"\0"))
    out: dict[str, bytes] = {}
    for name, typ, _clsid, start, size in ents:
        if typ != 2:
            continue
        if size < cutoff:
            raw = b"".join(mini[j * mssz:(j + 1) * mssz] for j in chain(start, mfat))
        else:
            raw = b"".join(sec(j) for j in chain(start, fat))
        out[name] = raw[:size]
    return out, root[2]


def _key(name: str) -> tuple:
    return (len(name), name.upper())       # [MS-CFB] 2.6.4：先比长度，再比大写 UTF-16


def _tree(names: list[str]):
    """有序名 → 平衡二叉树。返回 ({name: (left, right, color)}, 根子节点下标)。

    取中点建树 → 所有叶子深度只差 ≤1；最深一层不满时把这一层涂红（红节点无子、各路径黑高相等），
    满足 [MS-CFB] 2.6.4 对红黑树的要求；完全树则全黑。color：0 = red，1 = black。
    """
    order = sorted(names, key=_key)
    links: dict[str, list] = {n: [NOSTREAM, NOSTREAM, 1] for n in names}
    idx = {n: names.index(n) + 1 for n in names}       # 目录项下标（0 = 根存储）
    depth: dict[str, int] = {}

    def build(lo: int, hi: int, d: int):
        if lo > hi:
            return NOSTREAM
        mid = (lo + hi) // 2
        n = order[mid]
        depth[n] = d
        links[n][0] = build(lo, mid - 1, d + 1)
        links[n][1] = build(mid + 1, hi, d + 1)
        return idx[n]

    root_child = build(0, len(order) - 1, 0)
    if order:
        maxd = max(depth.values())
        if len(order) != (1 << (maxd + 1)) - 1:
            for n, d in depth.items():
                if d == maxd:
                    links[n][2] = 0
    return {n: tuple(v) for n, v in links.items()}, root_child


def write_streams(streams: list[tuple[str, bytes]], clsid: bytes) -> bytes:
    """单层存储（全部流 < 4096 字节）→ CFB 字节串。流顺序 = 目录项顺序。"""
    if any(len(b) >= MINI_CUTOFF for _, b in streams):
        raise CfbError("stream too large for mini stream layout")
    names = [n for n, _ in streams]
    links, root_child = _tree(names)
    # mini stream：每个流按 64 字节对齐顺排
    mini, mfat, starts = b"", [], {}
    for n, b in streams:
        k = (len(b) + MSZ - 1) // MSZ or 0
        starts[n] = len(mini) // MSZ if k else ENDOFCHAIN
        first = len(mini) // MSZ
        for j in range(k):
            mfat.append(first + j + 1 if j < k - 1 else ENDOFCHAIN)
        mini += b.ljust(k * MSZ, b"\0")
    n_mini_sec = (len(mini) + SSZ - 1) // SSZ
    n_mfat_sec = max(1, (len(mfat) * 4 + SSZ - 1) // SSZ)
    n_dir_sec = (len(streams) + 1 + 3) // 4
    # 扇区布局：0 FAT | 1.. 目录 | miniFAT | mini stream 容器
    dir0, mf0 = 1, 1 + n_dir_sec
    ms0 = mf0 + n_mfat_sec
    total = ms0 + n_mini_sec
    if total > SSZ // 4:
        raise CfbError("object too large for single FAT sector")
    fat = [FREESECT] * (SSZ // 4)
    fat[0] = FATSECT

    def link(first: int, k: int):
        for j in range(k):
            fat[first + j] = first + j + 1 if j < k - 1 else ENDOFCHAIN

    link(dir0, n_dir_sec)
    link(mf0, n_mfat_sec)
    link(ms0, n_mini_sec)

    def dirent(name: str, typ: int, color: int, left: int, right: int, child: int,
               cls: bytes, start: int, size: int) -> bytes:
        nb = (name + "\0").encode("utf-16-le") if name else b""
        return (nb.ljust(64, b"\0") + struct.pack("<HBB", len(nb), typ, color)
                + struct.pack("<III", left, right, child) + cls.ljust(16, b"\0")
                + b"\0" * 4 + b"\0" * 16 + struct.pack("<IQ", start, size))   # v3：size 高 4 字节为 0

    ents = [dirent("Root Entry", 5, 1, NOSTREAM, NOSTREAM, root_child, clsid,
                   ms0 if n_mini_sec else ENDOFCHAIN, len(mini))]
    for n, b in streams:
        left, right, color = links[n]
        ents.append(dirent(n, 2, color, left, right, NOSTREAM, b"", starts[n], len(b)))
    while len(ents) % 4:
        ents.append(dirent("", 0, 0, NOSTREAM, NOSTREAM, NOSTREAM, b"", 0, 0))
    hdr = (MAGIC + b"\0" * 16 + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6
           + struct.pack("<IIIIIIIII", 0, 1, dir0, 0, MINI_CUTOFF, mf0, n_mfat_sec,
                         ENDOFCHAIN, 0))
    hdr += struct.pack("<109I", 0, *([FREESECT] * 108))
    mfat_bytes = struct.pack(f"<{len(mfat)}I", *mfat).ljust(n_mfat_sec * SSZ, b"\xff")
    body = (struct.pack(f"<{SSZ // 4}I", *fat) + b"".join(ents) + mfat_bytes
            + mini.ljust(n_mini_sec * SSZ, b"\0"))
    return hdr.ljust(SSZ, b"\0") + body
