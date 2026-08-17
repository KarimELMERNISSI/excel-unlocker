"""
Module de déverrouillage de classeurs Excel (.xlsx, .xlsm).
"""

from .unlock import remove_excel_protection, inspect_excel_file

__all__ = ["remove_excel_protection", "inspect_excel_file"]
