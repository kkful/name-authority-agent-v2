# 名称规范记录智能体 V2

论文全量采集 -> Excel -> 结构化提取

## v2.1 更新 (2026-05-13)

- **翻页功能**: 从XHR响应中提取briefRequest参数,自动翻页采集多页结果
- 支持 `--max` 参数控制HTML打开篇数(默认20)
- 单次采集上限200篇元数据

## 翻页原理

知网搜索结果每页20条。翻页不是简单的page=N:
1. 首次搜索用 `boolSearch=true`
2. 从返回HTML中提取 `<input id="briefRequest" value="...">` (含翻页状态)
3. 翻页用 `boolSearch=false` + briefRequest值作为QueryJson
4. 循环直到达到目标数量

## 环境要求

- Python 3.9+ / Node.js 22+ / Chrome远程调试
- pandas, openpyxl

## 快速开始

```bash
pip install pandas openpyxl
# Chrome: chrome://inspect/#remote-debugging 勾选Allow
node C:/Users/Administrator/.claude/skills/web-access/scripts/check-deps.mjs
python collector.py 张伟 --max 50
```

首次运行需在Chrome知网页面手动搜索一次激活,之后全自动。
