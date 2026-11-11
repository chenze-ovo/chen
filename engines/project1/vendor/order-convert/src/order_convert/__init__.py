"""订单 Excel 自动转换核心库。"""
from .pipeline import convert_order_file, ConvertResult

__all__ = ["convert_order_file", "ConvertResult"]
__version__ = "1.0.0"
