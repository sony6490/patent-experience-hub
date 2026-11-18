# PatentKB — 本地专利项目经验检索台

输入新项目描述，从本地知识库中召回相似历史项目、推荐检索关键词与 IPC 分类号，
并生成检索式。答案带可点击引用，可直达原始文件或 Obsidian 知识页。

## 特点

- **纯本地索引**：文件抽取、分词、SQLite FTS5 全文检索全部在本机完成，建库与增量同步零 API 费用。
- **可选 LLM 增强**：配置 DeepSeek API Key 后生成结构化检索方案；未配置时自动降级为本地模式。
- **增量同步**：新增 / 修改 / 删除 / 复活自动识别；删除采用软删除，超过保留期后物理清理。
- **引用直达**：每条证据可打开原件、在 Obsidian 中打开对应知识页、或在资源管理器中定位。

## 快速开始

```bash
pip install -r requirements.txt
cp config.example.json config.json
# 编辑 config.json，把 root 改为你的知识库根目录
cp config.local.example.json config.local.json   # 可选，填 API Key
python build_index.py                            # 首次建索引（支持断点续跑）
python server.py                                 # 打开 http://127.0.0.1:8766
python sync.py loop                              # 可选：常驻增量同步
```

Windows 用户可直接双击 `start.bat`（启动服务并打开浏览器）和 `start_sync.bat`（常驻同步）。

## 目录结构

| 文件 | 说明 |
|---|---|
| `kb.py` | 零依赖文档抽取与元数据派生工具（被 kbcommon.py 复用） |
| `kbcommon.py` | 文件枚举、抽取、分词、SQLite、项目画像 |
| `build_index.py` | 全量建索引（断点续跑） |
| `server.py` | 本地 HTTP 服务与检索 / LLM 逻辑 |
| `sync.py` | 增量同步守护进程 |
| `repair.py` | 建库后元数据收尾（可选） |
| `rebuild_queries.py` | 单独重抽历史检索式（可选） |
| `index.html` | 前端界面 |
| `config.example.json` | 配置模板（复制为 config.json 后修改） |
| `config.local.example.json` | LLM 配置模板（API Key 只存本地） |

## 索引口径

- 扫描范围、排除目录、可索引扩展名均在 `config.json` 中配置。
- PDF 默认只索引文件名含报告类关键词的文件；专利号命名的文献原文通常不入项目知识库。
- 抽取库：python-docx（docx）、openpyxl（xlsx）、PyMuPDF（pdf）、内置 OOXML 解析（pptx）、
  LibreOffice（doc/xls，需本机安装）。
- 项目画像在文件没有项目编号时按客户目录兜底分组；IPC 只从报告类文档聚合以减少噪音。

## 隐私

- `config.local.json`（含 API Key）与 `data/`（索引库与全文缓存）均已加入 `.gitignore`，不会入库。
- 示例配置中不含任何真实客户名称、项目编号或内部路径。

## 成本

建库与同步 ¥0；启用 LLM 后每次问答约几分钱（按所用模型计费）。
