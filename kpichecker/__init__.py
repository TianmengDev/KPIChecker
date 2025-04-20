#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
KPI检查器包，用于检查文档中是否包含KPI自评语句
"""

__version__ = '0.1.0'
__author__ = 'KPI Checker Team'
__description__ = '检查Word文档中的KPI自评语句是否存在'

from .config import Config
from .docx_scanner import DocxScanner
from .docx_parser import DocxParser
from .kpi_checker import KPIChecker
from .report_generator import ReportGenerator
from .file_fixer import FileFixer

__all__ = [
    'Config',
    'DocxScanner',
    'DocxParser',
    'KPIChecker',
    'ReportGenerator',
    'FileFixer',
]
