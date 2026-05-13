"""V2 论文采集器 — 搜作者 -> 全量采集 -> Excel

用法: python collector.py 赵一冉 --max 5
采集: HTML阅读(标题/摘要/关键词/基金/作者简介) + 知网详情页(标题/作者机构/摘要/关键词/基金/专辑/专题/分类号) + XHR元数据
上限: 200篇
依赖: 旧规范文档CDP客户端(E:/名称规范系统/旧规范文档/author_agent)
"""
import sys, os, json, time, re, urllib.request
import pandas as pd

sys.path.insert(0, r"E:\名称规范系统\旧规范文档")
from author_agent.cdp_client import new_tab, close_tab, eval_js, click_at, page_text

PROXY = "http://localhost:3456"
OUTPUT_DIR = r"E:\名称规范系统\新规范文档系统\collected"
MAX_PAPERS = 200