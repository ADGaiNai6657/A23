# A23 文档理解与填表

第十七届服创大赛 A23 六周挑战档训练项目。支持 TXT、Markdown、DOCX 和 XLSX 的解析与事实抽取，以及 01 城市经济、03 历史疫情两套 Excel 模板的填写、评估和逐格溯源。[开发规格](docs/开发规格.md)

## 本地运行

需要 Python 3.14、Node.js `^22.18.0 || >=24.12.0` 和 npm。前端用 Vite 构建，由 FastAPI 在同一地址提供页面和接口。

### macOS 和 Linux

首次运行时克隆项目。已经有本地项目和 `.venv` 时，在项目根目录从安装 Python 依赖开始。

```sh
git clone https://github.com/ADGaiNai6657/A23.git
cd A23
python3.14 -m venv .venv
.venv/bin/python -m pip install -r backend/requirements.txt
npm --prefix frontend ci
npm --prefix frontend run build
(cd backend && ../.venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8000)
```

### Windows PowerShell

以下命令从项目根目录执行。

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r backend\requirements.txt
npm.cmd --prefix frontend ci
npm.cmd --prefix frontend run build
Set-Location backend
..\.venv\Scripts\python.exe -m uvicorn main:app --host 127.0.0.1 --port 8000
```

启动后打开 [应用首页](http://127.0.0.1:8000/)、[接口文档](http://127.0.0.1:8000/docs) 或 [健康检查](http://127.0.0.1:8000/health)。健康检查应返回 `{"ok":true}`。前端改动后重新构建并重启服务。若首页提示前端尚未构建，先检查 `frontend/dist/index.html`。macOS Apple Silicon 和 Debian Linux ARM64 已实测，Windows 命令尚待实机验证。

### 配置 DeepSeek API 密钥

在 DeepSeek 开放平台创建密钥，在项目的 `backend/.env` 中写入以下一行。macOS 可在项目根目录用 `open -a TextEdit backend/.env` 打开已有文件，或在 Finder 按 `⌘⇧.` 显示隐藏文件。保存后重启服务。[DeepSeek API 密钥管理](https://platform.deepseek.com/api_keys)

```dotenv
DEEPSEEK_API_KEY=你的实际密钥
```

`backend/.env` 已由 Git 忽略。「抽取事实并入库」会把选中的文档文本分批发给 DeepSeek，使用 `deepseek-flash`。该模型替代了第三周任务书中的 `deepseek-chat`，验收时记录实际使用的模型。解析、仅导入、03 填写和合成演示无需密钥。[DeepSeek 模型与计费](https://api-docs.deepseek.com/quick_start/pricing)

## 用合成样例上手

先停止同端口上已运行的服务。在根目录生成样例，再以演示目录启动。演示数据在 `backend/data/demo/`，与默认数据目录中的日常使用数据分开。

```sh
.venv/bin/python backend/demo.py
(cd backend && A23_DOC_DIR=data/demo/doc A23_DATA_DIR=data/demo/data \
  ../.venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8000)
```

Windows PowerShell 在项目根目录打开新终端，执行以下命令。

```powershell
.\.venv\Scripts\python.exe backend\demo.py
Set-Location backend
$env:A23_DOC_DIR = "$PWD\data\demo\doc"
$env:A23_DATA_DIR = "$PWD\data\demo\data"
..\.venv\Scripts\python.exe -m uvicorn main:app --host 127.0.0.1 --port 8000
```

打开首页，在「填写与验收」选中模板类型并录入表中的文件名，然后点「开始填写」。合成样例已经预填事实库和自检报告，可以直接下载结果或用 Excel 行号和列号追溯来源。

| 模板 | 库里的文档名 | 模板文件名 | 可核对的结果 |
| --- | --- | --- | --- |
| 01 城市经济 | `合成城市报告.docx` | `模板_合成城市.xlsx` | 1 行、写入 5 格。第 2 行第 2 列可查 GDP 原文 |
| 03 历史疫情逐日统计 | `合成逐日统计.xlsx` | `模板_合成逐日.xlsx` | 3 行、写入 19 格、2 格留空。第 3 行第 3 列是有效零值，第 3 行第 4 列可查拒填原因 |

## 用自己的材料

1. 把模板放到项目根目录的 `doc/` 下，保留原始模板和参考表。页面填写的是以「模板」开头或结尾的 XLSX 文件名，同名模板在材料目录中只能有一份。指定了 `A23_DOC_DIR` 时，将模板放到指定目录。项目自带测试集的用法见下节。
2. 在「文档入库」选择源文件。「查看解析块数」支持四种格式，适合先检查结构。要做 01，选城市报告 DOCX，填写可选的逗号分隔指标清单，然后点「抽取事实并入库」。要做 03，选逐日统计 XLSX，点「仅导入文档块」。
3. 在「填写与验收」选择 01 或 03，核对库里的文档名和模板名，点「开始填写」。完成后可以下载 Excel，以及按行列号查值、原文和拒填理由。与该版本匹配的评分会标注为「参考表评分」或「独立原文核对」。同名文件内容变更后，请先重新导入或抽取。
4. 如需比较多份文档的事实，在 `/docs` 中调用 `GET /facts`。用主体、指标、期次和口径定位候选。明确来源优先级时指定 `preferred_doc`，可同时查看其余候选。

直接填写目前对应 01 和 03 的模板结构。TXT 和 Markdown 的块编号使用原文行号，DOCX 和 XLSX 使用文档内块序号。XLSX 的块可以全量入库，交给模型抽取的范围是前 200 个非空数据行。03 填表直接使用指定日期区间的结构化数据。

## 接入 `dataset/测试集.zip`

在项目根目录解压并执行原文核对。首次执行或需要重建事实库时加 `--extract`，这会用已配置的 DeepSeek 密钥抽取 01。之后可省略该参数，复用事实库重跑。脚本将 16 份测试语料逐份解析并在本机入库，对 01 和 03 各填两次，逐格核对结果、原文坐标和模板原件的校验和。

```sh
.venv/bin/python -m zipfile -e dataset/测试集.zip dataset/测试集
.venv/bin/python backend/check_dataset.py --extract
.venv/bin/python backend/check_dataset.py
```

实际结果、逐格来源、抽取记录和独立核对报告保存在 `backend/data/actual/`。运行以下命令，在 [实际数据工作台](http://127.0.0.1:8001/) 选 01 或 03，使用页面预置的文档名和模板名直接实填。也可选中对应原文，再点「一键演示」查看前五个文档块、填写和下载。演示会复用已有事实库。

```sh
(cd backend && A23_DOC_DIR=../dataset/测试集/包含模板文件 A23_DATA_DIR=data/actual \
  ../.venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8001)
```

| 场景 | 原文核对口径 | 实测结果 |
| --- | --- | --- |
| 01 城市经济 | 按报告的正文段落顺序核对城市名和四项数值 | 113 个原文块、100 行，500 格与原文一致 |
| 03 历史疫情 | 按日期和国家排序，从源 XLSX 逐字段对照，空值保留 | 2,542 行 × 7 列，17,464 格与原文一致，330 格为空，误填与漏填均为 0 |

01 使用 `deepseek-flash` 在两个独立数据库各抽取一次，分别耗时 106.12 秒和 120.40 秒，两次都保留 400 条带原文证据的事实。两套模板在同库复验与重建库中的逐格记录一致。真实清单抽取另覆盖药品流通 TXT（保留 28 条）、文旅 Markdown（用原文指标名复跑后保留 3 条）和逐日统计 XLSX（前 200 行保留 400 条）。

03 的 Word 材料记录了单日省级情况，已解析作为范围核对材料。填表采用源 XLSX 中的国家或地区逐日数据。压缩包还含空气质量 Word 模板，填表验收范围为六周任务的 01 和 03。压缩包未附参考结果工作簿，以上数字来自独立原文核对，正式评分等参考表补齐。

## 部署到 Linux 服务器

下面是使用 systemd 的单机部署。示例使用服务账号 `a23`，项目路径 `/opt/a23`，材料和生成数据分别放在 `/srv/a23/doc` 和 `/var/lib/a23`。先安装上述 Python、Node.js、npm 和 Git，并准备可写的部署目录。

```sh
sudo useradd --system --user-group --home-dir /opt/a23 --shell /usr/sbin/nologin a23
sudo install -d -m 0750 -o a23 -g a23 /opt/a23 /srv/a23/doc /var/lib/a23
sudo -u a23 -H git clone https://github.com/ADGaiNai6657/A23.git /opt/a23
sudo -u a23 -H sh -ec '
  cd /opt/a23
  python3.14 -m venv .venv
  .venv/bin/python -m pip install -r backend/requirements.txt
  npm --prefix frontend ci
  npm --prefix frontend run build
'
```

如果账号已经存在，从创建目录开始。首次部署将模板放进 `/srv/a23/doc`，并确保服务账号可读。需要使用模型时，在服务器的 `/opt/a23/backend/.env` 中配置密钥，文件属主设为 `a23`，权限设为 `600`。

将以下内容保存为 `/etc/systemd/system/a23.service`。

```ini
[Unit]
Description=A23 document processing
After=network.target

[Service]
User=a23
Group=a23
UMask=0077
WorkingDirectory=/opt/a23/backend
Environment=A23_DOC_DIR=/srv/a23/doc
Environment=A23_DATA_DIR=/var/lib/a23
ExecStart=/opt/a23/.venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8000
Restart=on-failure

[Install]
WantedBy=multi-user.target
```

```sh
sudo systemctl daemon-reload
sudo systemctl enable --now a23
curl -fsS http://127.0.0.1:8000/health
sudo journalctl -u a23 -n 50
```

应用由单个 Uvicorn 进程提供服务，数据库、报告与生成表保存在 `A23_DATA_DIR`。暂停服务时可备份该目录和材料目录。应用自身没有登录，其他机器访问时通过带访问控制和 HTTPS 的反向代理转发到 `127.0.0.1:8000`，并从同一个域名的根路径提供页面与接口。更新时拉取代码、重跑依赖安装和前端构建，再执行 `sudo systemctl restart a23` 和健康检查。

## 验收

在项目根目录运行自动化与合成验收。先构建前端，以便覆盖静态页面路由。

```sh
.venv/bin/python -m pytest -q backend
.venv/bin/python backend/acceptance.py
```

正式材料到位后，按任务书把它们放入 `doc/` 并保留目录层级。在页面将所需文档入库后，用下列脚本填写并评估。根据实际文件名替换三个位置参数。页面刷新后可看到与这次结果匹配的评分。

```sh
.venv/bin/python backend/check_fill.py --scenario 01 '库内文档名.docx' '模板路径.xlsx' '参考路径.xlsx'
.venv/bin/python backend/check_fill.py --scenario 03 '库内文档名.xlsx' '模板路径.xlsx' '参考路径.xlsx'
```

若服务使用自定义数据目录，传入同一个 `--data-dir`，例如 `--data-dir /var/lib/a23`。正式验收使用 `--basis reference` 和独立的正式判据，判据结构见下节。命令核对同一目录的数据库、输入摘要、工作簿和本次报告。请在复验脚本生成报告后运行验收命令。

已验证原始样本的块数、行数、分母和逐格原文一致性。相对正式参考表的准确率及正式同库与重建库验收等参考表补齐后记录。[第四周任务](六周任务/第四周任务.md) · [第六周任务](六周任务/第六周任务.md)

## 六周交接记录

### 一分钟说明

在新目录克隆后，按本页「本地运行」安装和构建，再按「用合成样例上手」完成两套填表和评分。收到测试集压缩包后，按「接入测试集」运行，页面可显示独立原文核对结果。填表先输出带版本的工作簿和逐格记录，与版本匹配的评分才在页面显示。

### 环境与仓库

实测平台为 macOS Apple Silicon。`python -V` 输出 `Python 3.14.6`，`git --version` 输出 `git version 2.54.0`，`node -v` 输出 `v26.3.0`，`npm -v` 输出 `11.16.0`。项目按用户指定交付到公开 GitHub，以同一账号和项目名查询 Gitee 返回 404。[GitHub 项目](https://github.com/ADGaiNai6657/A23)

### 解析交付

同一份样本各请求两次 `POST /parse`，完整响应一致。耗时含本机请求、解析和响应序列化。XLSX 计时未调用模型。

| 指定样本 | 块数 | 两次耗时（秒） | 定位 |
| --- | ---: | --- | --- |
| 药品流通 TXT | 70 | 0.003、0.001 | 原文行号 |
| 文旅 Markdown | 68 | 0.001、0.001 | 原文行号 |
| 南京统计公报 DOCX | 1,217 | 0.036、0.036 | 271 段和 946 格，按段落号或表号、行列号定位 |
| 逐日统计 XLSX | 15,088 | 0.378、0.392 | 工作表和原表行号 |

不支持的 CSV、无法解码的 TXT、损坏的 DOCX 均返回 400，分别提示格式、内容或文件损坏。[完整 TXT](backend/data/parse_result.json) · [完整 DOCX](backend/data/parse_result_docx.json) · [XLSX 前 200 块](backend/data/parse_result_xlsx.json)

### 第三周真实抽取

模型为 `deepseek-flash`，请求参数为 temperature 0、批大小 30。耗时记录并发批次的实际经过时间。XLSX 入库 15,088 块，前 200 个非空数据行参与模型抽取。

| 样本 | 模式 | 入库块数 | 抽出 | 保留 | 丢弃 | 秒 |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| TXT | 通用 | 70 | 184 | 181 | 3 | 72.46 |
| TXT | 清单 | 70 | 18 | 18 | 0 | 32.31 |
| Markdown | 通用 | 68 | 214 | 214 | 0 | 117.35 |
| Markdown | 清单 | 68 | 3 | 3 | 0 | 9.86 |
| DOCX | 通用 | 1,217 | 2,151 | 1,761 | 390 | 405.11 |
| DOCX | 清单 | 1,217 | 12 | 12 | 0 | 23.01 |
| XLSX | 通用 | 15,088 | 2,430 | 2,390 | 40 | 276.79 |
| XLSX | 清单 | 15,088 | 400 | 400 | 0 | 24.61 |

清单、丢弃原因和各次耗时见抽取记录。首次 DOCX 通用抽取遇到坏字段而中断，修复后保留可独立核对的事实并记下拒收原因。XLSX 两种模式合计保留 2,790 条，导出前 200 条，其他样本完整导出。用 `check_extract.py --export-only` 可从已有数据库生成导出。同后缀有多个样本时，导出名带源文件名。[八次实测](backend/data/extraction_results.json) · [引导档事实](backend/data/facts_export.json) · [导出清单](backend/data/facts_export_manifest.json)

事实保存原指标名，归并键用于查询。期次取值同句的年份、同句最大年份或上文标题，记录期次来源块。同键异值保留全部候选，缺少明确来源优先级时拒绝采用，指定来源后保存采用标记和决策理由。控制用例还验证 2024 和 2025 各自回到原文第 1 行。[多源证据](backend/data/fusion_evidence.json)

### 数字口径与复验

01 的 113 块、100 行及城市顺序是固定锚点。单元格准确率是正确格数除以 500，行完整率是五格都写满的行数除以 100。03 固定 2,542 行，覆盖 17,464 个应填格和 330 个应空格。对格率和漏填率的分母是 17,464，误填率的分母是 330。03 的日期按年月日比较，Excel 数值序列按各工作簿的纪元转换。

同库复验逐项精确比较。重建库复验中，01 的四项模型相关数字可用实测区间，其余判据精确比较。当前两个独立抽取的 01 实测点均为写入 500 格、空 0 格、单元格准确率 100%、行完整率 100%，因此本版原文口径使用对应的单点区间。[原文判据及观测](backend/data/source_expectations.json)

```sh
.venv/bin/python backend/check_dataset.py --data-dir backend/data/actual
.venv/bin/python backend/acceptance.py --basis source --data-dir backend/data/actual \
  --expectations backend/data/source_expectations.json --mode same-db

.venv/bin/python backend/check_dataset.py --data-dir backend/data/actual/rebuilt-new --extract
.venv/bin/python backend/acceptance.py --basis source --data-dir backend/data/actual/rebuilt-new \
  --expectations backend/data/source_expectations.json --mode rebuilt
```

隔离目录实测输出如下。删除报告的随机目录已缩写。

```text
改数 | exit=1 | 01.rows 预期 100，实际 99，检查 source_report.json
改键 | exit=1 | 01.rows 预期 100，实际 None，检查 source_report.json
删除报告 | exit=1 | 验收失败：01 缺少本次报告 <隔离目录>/source_report.json，预期文件存在，实际缺失
恢复 | exit=0 | 验收通过，共 19 条数值判据
```

| 交叉验收项目 | 方式 | 结果 |
| --- | --- | --- |
| 新目录接手 | 独立克隆本地基线并应用本轮补丁，重新安装依赖 | 页面和健康检查可用，36 条回归测试、13 条合成判据通过 |
| 同库复验 | 在隔离目录复制已有库，两次重填并独立核对 | 01 写入 500 格，03 写入 17,464 格并保留 330 个空格，19 条判据通过 |
| 重建库复验 | 独立目录从空库真实重抽，耗时 156.87 秒，再填写两套模板 | 01 保留 400 条事实，两套原文核对均通过，共 19 条判据 |
| 拒填追溯 | 按国家名顺序核对 5 个各有至少两格拒填的国家 | 10 格均可回查原文坐标与拒填理由 |

### 局限与排查

- 正式参考表尚缺。页面的「独立原文核对」用于回归，正式评分走 `check_fill.py`。
- 当前填表范围是 01 和 03 的 XLSX 模板。空气质量 DOCX 模板可解析，填表功能待扩展。
- 在线 XLSX 抽取处理前 200 个非空数据行。入库块数和按日期填表的覆盖范围另行核对。
- 在线抽取有条数波动。药品流通 TXT 的同一清单先后保留 28 条和 18 条，逐条证据校验均已执行。
- 当前桌面工作台已验证 macOS 浏览器，Windows 仍待实机复验。
- 首页不可用先访问 `/health`，模板错误查 `A23_DOC_DIR` 与文件名，报告缺失查 `A23_DATA_DIR` 下的评分、逐格记录及版本。输入更新后，按场景重新导入或抽取，再填写并生成自检报告。
