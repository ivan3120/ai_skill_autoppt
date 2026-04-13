# -*- coding: utf-8 -*-
"""
Services - 服务层
"""

# 延迟导入，避免循环依赖
def __getattr__(name):
    if name == "ExcelParser":
        from scripts.services.excel_parser import ExcelParser
        return ExcelParser
    elif name == "parse_excel":
        from scripts.services.excel_parser import parse_excel
        return parse_excel
    raise AttributeError(f"module has no attribute '{name}'")

__all__ = ["ExcelParser", "parse_excel"]
