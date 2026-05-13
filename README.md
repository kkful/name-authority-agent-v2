# 名称规范记录智能体 V2

论文全量采集 -> Excel -> 结构化提取

## 阶段1：论文采集器 (collector.py)
- 知网同步XHR搜索
- HTML阅读全文 + 详情页元数据 + 元数据兜底
- 全量输出Excel

## 使用
```bash
python collector.py 赵一冉 --max 5
```
