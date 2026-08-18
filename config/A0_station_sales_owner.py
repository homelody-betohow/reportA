"""
按站点固定销售负责人（非 AMZ / 多站点平台）。

用于 M2：platform_shop.ops_owner 未命中时（尤其是 LM 的 -ls / -xj 站点）按本表赋值。
新增或调整站点归属时只改本文件。
"""

from __future__ import annotations

# 站点代码 → 销售负责人
STATION_SALES_OWNER: dict[str, str] = {
    "LM-TOTO": "曹乐思",
    "LM-ES-BTH": "陈晓佳",
    "LM-FR-BTH": "陈晓佳",
    "LM-IT-BTH": "陈晓佳",
    "LM-PL-BTH": "陈晓佳",
    "LM-PT-BTH": "陈晓佳",
    "TEMU-AIH": "刘思兰TEMU",
    "TEMU-BV": "陈培TEMU",
    "TEMU-AL": "陈培TEMU",
    "TEMU-HM": "王园芳TEMU",
    "TEMU-KR-A": "李炜玲TEMU",
    "TEMU-KR-B": "李炜玲TEMU",
    "TEMU-KR-C": "李炜玲TEMU",
    "TEMU-HJ-A": "陈培TEMU",
    "TEMU-HJ-B": "陈培TEMU",
    "TEMU-HJ-C": "陈培TEMU",
    "TEMU-NF-A": "刘思兰TEMU",
    "TEMU-NF-B": "刘思兰TEMU",
    "TEMU-NF-C": "刘思兰TEMU",
    "LM-FR-BC-ls": "曹乐思",
    "LM-FR-BC-xj": "陈晓佳",
    "LM-ES-BC-ls": "曹乐思",
    "LM-ES-BC-xj": "陈晓佳",
    "LM-PT-BC-ls": "曹乐思",
    "LM-PT-BC-xj": "陈晓佳",
    "LM-IT-BC-ls": "曹乐思",
    "LM-IT-BC-xj": "陈晓佳",
    "TEMU-BZ": "李炜玲TEMU",
    "TEMU-AQ": "陈培TEMU",
    "LM-FR-RP-ls": "曹乐思",
    "LM-FR-RP-xj": "陈晓佳",
    "LM-ES-RP-ls": "曹乐思",
    "LM-ES-RP-xj": "陈晓佳",
    "LM-PT-RP-ls": "曹乐思",
    "LM-PT-RP-xj": "陈晓佳",
    "LM-IT-RP-ls": "曹乐思",
    "LM-IT-RP-xj": "陈晓佳",
}

# 需要按站点设负责人的站点集合（便于 isin 判断）
STATION_OWNER_KEYS = frozenset(STATION_SALES_OWNER)
