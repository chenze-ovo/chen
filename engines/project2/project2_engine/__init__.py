"""项目2加工。规则只来自调用方传入的配置和店铺表，不引用 order_convert。"""

from .pipeline import build_oversea_copy_rows, process_orders, write_oversea_copy_from_formal

__all__ = ["process_orders", "build_oversea_copy_rows", "write_oversea_copy_from_formal"]
