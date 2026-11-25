# 专利项目经验检索台（PatentKB）

一个**跑在自己电脑上**的专利项目经验检索窗口：输入新项目的技术描述，它从你过去的交付物里
找出相似的历史项目、推荐检索关键词与 IPC 分类号，并起草一份检索方案。答案里每一条结论都带
可点击的引用，直接跳到当初那份 Word / Excel / PDF 原文，或者 Obsidian 里的知识页。

解决的问题很具体：三年前的检索式怎么写的、这个技术方向以前做过哪家的 FTO、当时用过哪些关键词
和分类号——这些东西都躺在共享盘里，但没人搜得到。

> 说明：本仓库是脱敏后的通用版本，不含任何客户资料、索引数据库或 API Key。
> 项目背景与踩坑记录见公众号文章（链接在文末）。

## 它做什么

- **相似历史项目**：按项目编号或客户归组成"项目画像"，用 BM25 分数 + 关键词重叠度 + IPC 重合度排序
- **关键词推荐**：从相似项目的正文里抽技术词，中英分列，附出现频次和来源
- **IPC 推荐**：只在跟查询相关的项目里聚合分类号，按命中项目和文档数打分
- **检索方案**：接 DeepSeek（或任意 OpenAI 兼容接口）把上面的结果整理成可执行的检索式；
  没配 Key 时自动降级为纯本地模式，只给检索建议不给 LLM 结论
- **引用可核验**：证据编号 `[1][2]` 对应真实文件，点击打开原件 / 在资源管理器定位 / 打开 Obsidian 页面
- **增量同步**：新增、修改、删除、恢复自动识别，不需要重建索引

## 技术选择

| 环节 | 选择 | 为什么 |
|---|---|---|
| 抽取 | python-docx / openpyxl / PyMuPDF / 内置 OOXML 解析 / LibreOffice | 旧脚本用正则切 docx 会把表格属性当正文吞掉 |
| 索引 | SQLite FTS5 + jieba 分词 | 零服务依赖，单机可跑，装完就能用 |
| 排序 | BM25 + 项目画像重叠 | 项目级召回比纯文档召回更贴合"找做过什么"这个问题 |
| 向量 | **没用** | 全量 embedding 要花钱要维护，关键词 + 画像已经够用；语义近义（如"避障"↔"碰撞避免"）召回确实有限 |
| 服务 | Python 标准库 `http.server` + ThreadingHTTPServer | 少一个框架少一类环境问题 |
| LLM | OpenAI 兼容接口 | 换模型只改配置，不改代码 |

## 安装与运行

### 1. 依赖

```bash
python -m pip install -r requirements.txt   # jieba, openpyxl, python-docx, PyMuPDF, openai
```

`.doc` / `.xls` 需要本机装 LibreOffice（Windows/macOS/Linux 任一发行版均可），程序按常见路径自动找。
没有也不影响其他格式。

### 2. 配置

```bash
copy config.example.json config.json
copy config.local.example.json config.local.json
```

改 `config.json`：

- `root`：知识库根目录
- `scan_dirs`：要索引的子目录（其余目录不进库）
- `exclude_dirs`：要排除的目录名，例如 `财务资料`、`个人档案`、`node_modules`
- `pdf_whitelist_keywords`：只有文件名匹配这些关键词的 PDF 才抽正文——专利号命名的现有技术原文
  成千上万份，混进项目库只会污染召回
- `server_port`：默认 `8766`，被占用时改这里

`config.local.json` 填 LLM 参数（`api_key`、`base_url`、`model`）。**它已在 .gitignore 里，不会进版本库。**

### 3. 想先看效果：生成演示库

不想拿自己的资料试，可以先生成一份**完全虚构**的演示知识库（4 个假客户、10 个假项目、
70 个文件，含报告 docx / 检索历史 xlsx / 检索式 txt / 沟通纪要 html，以及故意放在
`公司资料/标准报价与结算` 里用来验证「报价书示例分类号不污染 IPC 推荐」的诱饵文件）：

```bash
python tools\make_demo_vault.py            # 默认生成到仓库同级 demo-vault\
```

然后把 `config.json` 的 `root` 指到这个目录，`scan_dirs` 填
`["项目资料", "归档文件", "公司资料"]`，`exclude_dirs` 加上 `个人档案`、`财务资料`，
`code_pattern` 填 `PB-?\d{4}`，正常建索引启动即可。

### 4. 建索引

```bash
python build_index.py                # 全量，支持断点续跑
python build_index.py --workers 8    # 多进程，doc/xls 转换会占用较多 CPU
```

### 5. 启动

Windows 直接双击 **`专利检索台-一键启动.bat`**：它会检查 Python、拉起服务和同步循环、
等接口就绪后打开浏览器；已经在跑就不会重复启动。**`专利检索台-停止服务.bat`** 只结束这两个进程。
`start.bat` / `start_sync.bat` 是单进程的旧入口。

或者手动：

```bash
python server.py
python sync.py loop          # 常驻增量同步（默认 5 分钟一轮）
```

打开 http://127.0.0.1:8766

开机自启：`install_autostart.bat`（右键以管理员身份运行）。

## 使用

- **生成检索方案**：输入技术描述 → 快速 / 深度两种模式 → 得到历史项目、关键词、IPC、检索式、检索计划
- **纯检索**：只搜文档不花钱，适合平时翻旧报告
- **左侧筛选**：客户 / 技术域 / 项目类型 / 年份 / 项目编号；点「筛选」直接过滤文档列表
- **⚙️ 知识库设置**：在网页里勾选要纳入的文件夹，保存后立刻增量同步，不用改 JSON 也不用重启服务。
  一次要移除超过 50 份文档时会先提示，确认后点「强制同步」
- **同步知识库**按钮：手动触发一次增量扫描（`python sync.py once`，需要跳过批量保护时加 `--force`）

## 数据与隐私

- 删除文件采用**软删除**，默认保留 30 天后才物理清理（`tombstone_days`）
- 全文缓存在 `data/text/`，索引在 `data/kb.db`，两者都不进版本库
- 只有本地监听 `127.0.0.1`，不对外网开放
- 发给 LLM 的只有查询词、项目画像摘要和少量正文片段，整篇文档不外传

## 已知限制

- 不做语义向量检索，近义表述召回有限
- 扫描件 PDF 抽不出正文，只能按文件名检索
- 项目归组依赖文件名/目录里的项目编号，缺失时按客户目录兜底，大客户会形成较大的组
- 历史检索式抽取靠"标题含检索式/关键词"+ 大写布尔算符识别，覆盖不全，只当参考样例

## 仓库结构

```
kbcommon.py              配置、抽取、分词、建库、项目画像
build_index.py           全量建索引（多进程 + 断点续跑）
sync.py                  增量同步（once / loop）
server.py                检索 + LLM 方案 + HTTP 接口（标准库实现）
repair.py                建库后元数据收尾
rebuild_queries.py       单独重抽历史检索式
static/index.html        前端界面
kb.py                    被复用的文档抽取工具
tools/make_demo_vault.py 生成虚构演示库（用于试跑与截图）
config.example.json      配置模板
专利检索台-一键启动.bat    检查依赖 → 拉起服务与同步 → 打开浏览器
专利检索台-停止服务.bat    按 data/*.pid 精确结束，不误杀其他 Python
```

## 成本

建库和同步 **¥0**。开启 LLM 后一次问答约几千 tokens，DeepSeek 口径几分钱以内。

---

公众号文章：*（发布后补充链接）*
