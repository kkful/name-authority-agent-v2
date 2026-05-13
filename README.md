# 名称规范记录智能体 V2

论文全量采集 -> Excel -> 结构化提取

## 环境要求

- Python 3.9+
- Node.js 22+ (CDP Proxy)
- Chrome 浏览器 (远程调试模式)
- pandas, openpyxl

## 快速开始

### 1. 安装依赖

```bash
pip install pandas openpyxl
```

### 2. 启动 Chrome 远程调试

Chrome 地址栏输入 `chrome://inspect/#remote-debugging`，勾选 "Allow remote debugging"，重启 Chrome。

### 3. 启动 CDP Proxy

```bash
node C:/Users/Administrator/.claude/skills/web-access/scripts/check-deps.mjs
```

### 4. 运行采集器

```bash
python collector.py 赵一冉 --max 5
```

首次运行会在 Chrome 中打开知网页面并弹出绿色提示，手动搜索一次激活标签后，重新运行即可自动采集。

### 5. 输出

Excel 文件保存在 `./collected/` 目录，文件名格式：`{姓名}_论文采集_{时间}.xlsx`

## 参数说明

```bash
python collector.py <姓名> [--max N]
```

- `姓名`: 要搜索的作者名（必填）
- `--max N`: 最多打开的HTML篇数（可选，默认20）

## 采集内容

| 来源 | 提取字段 |
|------|---------|
| HTML阅读 | 标题、摘要、关键词、基金、作者简介、收稿日期 |
| 知网详情页 | 标题、作者机构、摘要、关键词、基金、专辑、专题、分类号 |
| XHR元数据 | 标题、摘要片段（HTML和详情页都不可用时兜底） |

## 上限

单次最多采集 200 篇论文元数据，HTML 打开数量由 `--max` 参数控制。

## 依赖关系

复用旧规范文档的 CDP 客户端：`E:\名称规范系统\旧规范文档\author_agent\cdp_client.py`

如需在其他机器部署，修改 `collector.py` 顶部的 `sys.path.insert` 路径。
