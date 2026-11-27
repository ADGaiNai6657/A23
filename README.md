# A23 文档理解与填表

第十七届服创大赛 A23 六周挑战档训练项目。支持 TXT、Markdown、DOCX 和 XLSX 的解析与事实抽取，以及 01 城市经济、03 历史疫情两套 Excel 模板的填写、评估和逐格溯源。[开发规格](docs/开发规格.md)

## 项目目录

| 路径 | 内容 |
| --- | --- |
| `backend/` | FastAPI、文档解析、抽取、SQLite、填写、评分和验收脚本 |
| `backend/data/` | 默认运行数据目录、数值判据及 `deliverables/` 验收产物 |
| `frontend/` | Vue 源码、npm 锁文件及本地构建目录 |
| `dataset/` | 公开测试集压缩包 |
| `六周任务/`、`docs/` | 原始要求、实现规格、当前验收结果和交接记录 |

## 本地运行

需要 Python 3.14、Node.js `^22.18.0 || >=24.12.0` 和 npm。前端用 Vite 构建，由 FastAPI 在同一地址提供页面和接口。

### macOS 和 Linux

首次运行时克隆项目。已经有本地项目和 `.venv` 时，在项目根目录从安装 Python 依赖开始。

```sh
git clone https://github.com/ADGaiNai6657/A23.git
cd A23
python3.14 -m venv .venv
.venv/bin/python -m pip install -r backend/requirements.txt
(cd frontend && npm ci && npm run build)
(cd backend && ../.venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8000)
```

### Windows PowerShell

以下命令从项目根目录执行。

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r backend\requirements.txt
Push-Location frontend
npm.cmd ci
npm.cmd run build
Pop-Location
Set-Location backend
..\.venv\Scripts\python.exe -m uvicorn main:app --host 127.0.0.1 --port 8000
```

启动后打开 [应用首页](http://127.0.0.1:8000/)、[接口文档](http://127.0.0.1:8000/docs) 或 [健康检查](http://127.0.0.1:8000/health)。健康检查应返回 `{"ok":true}`。前端改动后重新构建并刷新。若首页提示前端尚未构建，先检查 `frontend/dist/index.html`。本轮复验在 macOS Apple Silicon；历史部署记录含 Debian Linux ARM64，Windows 命令仍待实机验证。

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
(cd backend && A23_DOC_DIR=data/demo/doc A23_DATA_DIR=data/demo/data A23_REFERENCE_DIR=data/demo/doc A23_SYNTHETIC_REFERENCE=1 \
  ../.venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8000)
```

Windows PowerShell 在项目根目录打开新终端，执行以下命令。

```powershell
.\.venv\Scripts\python.exe backend\demo.py
Set-Location backend
$env:A23_DOC_DIR = "$PWD\data\demo\doc"
$env:A23_DATA_DIR = "$PWD\data\demo\data"
$env:A23_REFERENCE_DIR = "$PWD\data\demo\doc"
$env:A23_SYNTHETIC_REFERENCE = "1"
..\.venv\Scripts\python.exe -m uvicorn main:app --host 127.0.0.1 --port 8000
```

打开首页，在「填写与验收」选中模板类型并录入表中的文件名，然后点「开始填写」。合成样例已经预填事实库和自检报告。配置的合成参考表会给本次填写重新评分。可以上传同名原文并点「一键演示（已有库）」，然后下载 Excel 或按行列号追溯。

| 模板 | 库里的文档名 | 模板文件名 | 可核对的结果 |
| --- | --- | --- | --- |
| 01 城市经济 | `合成城市报告.docx` | `模板_合成城市.xlsx` | 1 行、写入 5 格。第 2 行第 2 列可查 GDP 原文 |
| 03 历史疫情逐日统计 | `合成逐日统计.xlsx` | `模板_合成逐日.xlsx` | 3 行、写入 19 格、2 格留空。第 3 行第 3 列是有效零值，第 3 行第 4 列可查拒填原因 |

## 用自己的材料

1. 把模板放到项目根目录的 `doc/` 下，保留原始模板和参考表。页面填写的是以「模板」开头或结尾的 XLSX 文件名，同名模板在材料目录中只能有一份。指定了 `A23_DOC_DIR` 时，将模板放到指定目录。项目自带测试集的用法见下节。
2. 在「文档入库」选择源文件。「查看解析块数」支持四种格式，适合先检查结构。要做 01，选城市报告 DOCX，填写可选的逗号分隔指标清单，然后点「抽取事实并入库」。要做 03，选逐日统计 XLSX，点「仅导入文档块」。
3. 在「填写与验收」选择 01 或 03，核对库里的文档名和模板名，点「开始填写」。完成后可以下载 Excel，以及按行列号查值、原文和拒填理由。与该版本匹配的评分会标注为「参考表评分」或「独立原文核对」。同名文件内容变更后，请先重新导入或抽取。
4. 如需比较多份文档的事实，在 `/docs` 中调用 `GET /facts`。用主体、指标、期次和口径定位候选。有明确来源时指定 `preferred_doc`。用相同字段请求 `POST /facts/decisions` 可保存候选、采用项与理由，凭返回的 ID 查询。

「一键演示（已有库）」走解析、填写、读报告，要求库内已保存同一版本原文及所需事实。「从上传开始」会先为 01 调用 DeepSeek 抽取入库，或为 03 导入 XLSX 再填写。01 会在发出请求前提示密钥、耗时和费用。直接填写目前对应 01 和 03 的模板结构。TXT 和 Markdown 的块编号使用原文行号，DOCX 和 XLSX 使用文档内块序号。XLSX 的块可以全量入库，交给模型抽取的范围是前 200 个非空数据行。03 填表直接使用指定日期区间的结构化数据。

## 自然语言操作

在工作台「自然语言操作」选择不超过 12 MB 的 DOCX 或 XLSX，输入中文指令后点「识别并预览」。预览会把文档内容及指令发送到已配置的 DeepSeek，页面先询问是否继续。核对操作列表中的目标位置、原内容和预期结果，再点「确认并执行」。成功后逐项显示执行结果，编辑或排版任务可下载文件副本。提取任务直接显示原文和第几段或工作表单元格位置。

首版支持指定正文段的文字替换、整段加粗及左/中/右对齐，指定工作表单元格改值和已有范围内的加粗及对齐，以及 DOCX 段落或 XLSX 单元格的原文提取。一条指令可组合多项操作。多个工作表时须在指令中写明工作表名称。段落从 1 开始计数，Excel 目标须位于已有工作表范围内，单次格式范围最多 1000 格。遇目标不明确、越界、操作不受支持或模型/文件错误时会显示中文原因。全部操作通过校验后才生成副本。确认计划绑定上传文件名与 SHA-256，换文件后要重新预览。计划和下载副本在单进程内暂存 30 分钟，重启服务后需要重新执行。

接口使用同一个上传文件。先用 `POST /operations/preview` 发送 multipart 的 `file`、`instruction`，得到 `plan_id` 和逐项预览。再用 `POST /operations` 发送原文件及 `plan_id`，根据返回的 `result_id` 调用 `GET /operations/files/{result_id}`。提取任务不生成副本。[接口说明](docs/开发规格.md#功能与验收)

段落或 Excel 范围的格式操作可只指定加粗、只指定对齐，或同时指定两者；未指定的属性保留。Excel 写入必须显式提供值，`null` 表示清空，数值 `0` 和布尔值可正常写入与提取。Word 文字替换保留文字片段旁的内嵌图片；若替换范围跨越图片、对象或特殊文字节点，预览会拒绝并指出段落。

## 空气质量 Word 模板首版

工作台「空气质量 Word 模板填写」分别上传公开监测 XLSX 和空白 DOCX 模板。`POST /templates/inspect` 校验模板段落、三张表的表头、目标格的合并结构与源表监测时间，显示每市源行数和需要扩充的模板行数；确认后 `POST /templates/fill` 用同一对文件及 `plan_id` 生成副本。页面提供下载和逐格来源（源工作表、行列、原值）；多工作表的列序分别识别，目标格原有段落样式与对齐保留。本版固定处理 2025-11-25 09:00 的德州、潍坊、临沂三市。源表分别有 25、29、25 条站点，前两表各扩一行以保留全部记录；源值「—」原样保留并标作缺失。该链路从结构化 XLSX 直接读取，不调用 DeepSeek。确认计划与两文件 SHA-256 绑定；模板字段、时间或目标格结构不符会在预览拒绝。计划最多保留 100 份、下载结果最多 20 份，均在访问或新增时清理超过 30 分钟的记录。

## 上传 XLSX 目标模板首版

工作台「上传 XLSX 目标模板」分别选择逐日统计源 XLSX 和目标 XLSX，点「预览工作表与表头」，核对目标工作表及每一列对应的 03 字段，再点「确认映射并填写」。页面显示已填格、真空格及逐格来源，可下载填写副本。目标表可调整六个业务字段的列序；若没有日期列，系统在最前面插入日期。预览按工作表列出源表头，若同一源工作表的必要字段重复，会指出冲突列并拒绝生成计划。确认时重新校验两份文件的 SHA-256 和完整映射。

此首版只处理 03 的日期范围、国家/地区和六个既有字段，按日期、国家/地区排序，从结构化 XLSX 直接读取，不调用 DeepSeek。目标模板需有一张可选择的 6 列或 7 列表头；不支持任意业务字段、公式计算、跨来源事实裁决或五套正式模板。结果逐格报告包含目标行列、值或拒填原因、源工作表、源行列和原值。接口为 `POST /templates/xlsx/inspect` 和 `POST /templates/xlsx/fill`，两次均上传 `source` 与 `template`；填写另传预览返回的 `plan_id`、目标 `sheet` 和以目标列字母为键、字段键为值的 JSON `mapping`。下载沿用 `GET /operations/files/{result_id}`。计划与下载副本在单进程内暂存，过期后需要重新预览。

预览会列出目标工作表的实际写入范围，只允许覆盖该范围内的空白单元格。范围内已有公式、内容、批注或合并格时，预览指出位置并拒绝计划；范围外的行列和其他工作表保留原位。六列表头需要插入日期列，因此若表头之外已有内容，也会拒绝以避免移动原有单元格。逐格页面将源表真空、来源冲突和数值精度拒填分开显示；冲突可查看每条候选的值、工作表、坐标和原值。数值采用保守口径：仅写入绝对值小于 10¹⁵、至多 15 位有效数字且转换后十进制值不变的数；其余保持目标真空并以 `precision_loss` 记录原值。副本落盘后还会回读全部目标格，若与报告不一致则不发布结果。

## 接入 `dataset/测试集.zip`

从仓库根目录运行下列命令。样本压缩包使用旧式中文文件名编码，由 `unpack_dataset.py` 解码。示例建两个隔离数据库，先复用各自历史抽取事实，再独立读取公开原文核对结果。已有数据库可省略 `--archived-facts` 做同库复验。空气质量自动化测试会从已跟踪的 `dataset/测试集.zip` 解压到 pytest 临时目录，不依赖仓库内默认的 `dataset/测试集/` 解压目录。

```sh
work_a23=$(mktemp -d)
.venv/bin/python backend/unpack_dataset.py --output "$work_a23/corpus"
.venv/bin/python backend/check_dataset.py --root "$work_a23/corpus" --data-dir "$work_a23/same" \
  --archived-facts backend/data/deliverables/source/city-original-historical.json
.venv/bin/python backend/check_dataset.py --root "$work_a23/corpus" --data-dir "$work_a23/same"
.venv/bin/python backend/check_dataset.py --root "$work_a23/corpus" --data-dir "$work_a23/rebuilt" \
  --archived-facts backend/data/deliverables/source/city-rebuilt-historical.json
.venv/bin/python backend/verify.py --same-db-dir "$work_a23/same" \
  --rebuilt-data-dir "$work_a23/rebuilt" --rebuilt-kind history
```

执行真实空库重抽时，另选新的数据目录，在 `check_dataset.py` 中用 `--extract` 代替 `--archived-facts`，综合验收使用默认的 `--rebuilt-kind online`。抽取会把 01 城市报告分批发送到 DeepSeek 并计费。两种重建都会生成 `acceptance_run.json`，用于核对初始空库、最终库和本次填写。在线模式读取数据库里的原文，检查全部块和 Prompt 摘要，并离线重放原始响应比对实际入库事实。随仓库交付的在线 ZIP 可用公开原文离线核验批次和留档事实。校验原库还需对应运行目录里的 `a23.db`。旧 `backend/data/actual/` 内的原始档案供查阅，当前复验按上面的隔离路径生成。设置 `A23_SOURCE_CHECK_ROOT` 后，页面会对两套公开模板和字节匹配的原文生成本次独立核对报告。

```sh
A23_DOC_DIR="$work_a23/corpus/包含模板文件" A23_DATA_DIR="$work_a23/same" \
  A23_SOURCE_CHECK_ROOT="$work_a23/corpus" \
  .venv/bin/python -m uvicorn main:app --app-dir backend --host 127.0.0.1 --port 8001
```

01 原文核对为 113 块、100 行、500 格；03 为 2,542 行，应填 17,464 格，应空 330 格，实测误填和漏填均为 0。已有参考结果时，可以运行 `check_fill.py` 或设置 `A23_REFERENCE_DIR`。公开压缩包当前包含原文和模板，正式参考结果仍需另行提供。两次历史抽取及本轮结果可独立查阅。[出库产物与核对记录](docs/六周验收记录.md)

### 解析与抽取实测

四份公开样本各调用 `/parse` 两次，完整响应一致。TXT 与 Markdown 都接受 UTF-8 和 UTF-8 BOM，按行切块并保留原行号。DOCX 使用段落或表格坐标，XLSX 使用工作表名和原行号。四格式都返回 `file` 及统一的 `blocks`，只读解析不会入库。

| 样本 | 块数 | 两次解析耗时（秒） |
| --- | ---: | --- |
| 药品流通 TXT | 70 | 0.020、0.001 |
| 文旅 Markdown | 68 | 0.001、0.001 |
| 南京公报 DOCX | 1,217（271 段 + 946 格） | 0.035、0.036 |
| 逐日统计 XLSX | 15,088 | 0.383、0.381 |

GBK TXT、CSV 和损坏 DOCX 实测均返回 400。以下四格式抽取数据来自同一份历史实验记录。模型为 `deepseek-flash`，temperature 0，批大小 30。XLSX 的模型调用范围为前 200 个非空数据行，以控制调用时长。后续 TXT 两轮的输入摘要和原始响应单独保留。[四格式历史记录](backend/data/extraction_results.json) · [TXT 两轮对照](docs/六周验收记录.md#txt-两轮对照)

| 样本 | 模式 | 入库块数 | 候选 | 保留 | 丢弃 | 耗时（秒） |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| TXT | 通用 | 70 | 184 | 181 | 3 | 72.46 |
| TXT | 清单 | 70 | 18 | 18 | 0 | 32.31 |
| Markdown | 通用 | 68 | 214 | 214 | 0 | 117.35 |
| Markdown | 清单 | 68 | 3 | 3 | 0 | 9.86 |
| DOCX | 通用 | 1,217 | 2,151 | 1,761 | 390 | 405.11 |
| DOCX | 清单 | 1,217 | 12 | 12 | 0 | 23.01 |
| XLSX | 通用 | 15,088 | 2,430 | 2,390 | 40 | 276.79 |
| XLSX | 清单 | 15,088 | 400 | 400 | 0 | 24.61 |

指标保留原文说法，同义词按显式清单归并。GDP 总量、GDP 增速、病例数和每百万人病例数各自有独立含义。期次优先看能直接对应到数值的分句或紧邻标题。仍无明确期次时，取同句最大年份，再继承上文最近的年份标题，保留出处块号。两份真实公开公报的同键床位数分别为 1037 和 1029.8 万张，默认保留两条，指定来源后可选用一条。实验 CLI 在 `extraction-runs/` 保存逐批响应、判废和入库记录，网页抽取返回汇总。[多源复核](backend/data/deliverables/source/fusion-real.json) · [解析与抽取函数索引](docs/开发规格.md#解析与抽取函数)

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
  (cd frontend && npm ci && npm run build)
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

核心接口与版本字段见开发规格。`/reports` 可以在工作台用「读取报告」查看当前结果。点击已有库演示时需要先上传原文。试过的三个错误提示为「后端没有响应，请启动服务后重试。」「库里没有这份文档」「没有这份文档的填写记录」。[接口与评分约定](docs/开发规格.md#功能与验收)

在项目根目录运行离线综合验收，包含自动化测试、合成评分、交付工件核验和公开在线归档重放。

```sh
(cd frontend && npm run build)
.venv/bin/python backend/verify.py
```

| 验收入口 | 需要的材料 | 核对内容 |
| --- | --- | --- |
| 默认综合验收 | 已构建前端和随仓库交付的工件 | 合成流程、公开工作簿与报告、在线归档原文和批次重放 |
| 综合验收带 `--same-db-dir` 和 `--rebuilt-data-dir` | 两个完整运行目录，包括原数据库和对应凭证 | 公开原文判据或独立参考判据，重建留档与数据库事实的一致性 |
| `--basis reference --expectations` | 上述两个目录、同版参考工作簿评分和独立判据配置 | 按参考评分，绑定观察点与实际验收的原文、模板和参考版本 |

正式材料到位后，按任务书把它们放入 `doc/` 并保留目录层级。在页面将所需文档入库后，用下列脚本填写并评估。根据实际文件名替换三个位置参数。页面刷新后可看到与这次结果匹配的评分。

```sh
.venv/bin/python backend/check_fill.py --scenario 01 '库内文档名.docx' '模板路径.xlsx' '参考路径.xlsx'
.venv/bin/python backend/check_fill.py --scenario 03 '库内文档名.xlsx' '模板路径.xlsx' '参考路径.xlsx'
```

若服务使用自定义数据目录，传入同一个 `--data-dir`，例如 `--data-dir /var/lib/a23`。正式验收使用 `--basis reference` 和独立的正式判据，判据结构见下节。命令核对同一目录的数据库、输入摘要、工作簿和本次报告。请在复验脚本生成报告后运行验收命令。

正式参考表默认检查固定行数及表头顺序。01 必须有 100 个城市的五列有效值，03 核对应填与应空分母。合成 API 评分需显式设置 `A23_SYNTHETIC_REFERENCE=1`，合成 CLI 评分需加 `--synthetic`。已有本次验收产物时，可用 `check_fill.py --score-existing` 加评分，保留填写批次。正式参考表成绩在材料到位后记录。[第四周任务](六周任务/第四周任务.md) · [第六周任务](六周任务/第六周任务.md)

收到正式参考表后，分别给同库和重建目录中的 01、03 工作簿生成参考评分。命令使用 `check_fill.py <库内文档名> <模板路径> <参考表路径> --scenario 01 --data-dir <本次输出目录> --score-existing`，03 改用 `--scenario 03`。两种模式使用同一对参考表。

```sh
.venv/bin/python backend/verify.py --basis reference --expectations "$work_a23/reference-rules.json" \
  --same-db-dir "$work_a23/same" --rebuilt-data-dir "$work_a23/rebuilt"
```

参考判据 JSON 的 `basis` 填 `reference-workbook`，`reference_sha256` 分别列出 `01` 和 `03` 的参考工作簿摘要。`same-db` 为两场景逐位判据，`rebuilt.01` 为写入格数、空着格数和两个率的区间。`observations` 至少列两条独立记录，每条包括相对于配置文件的 `data_dir` 及 `mode`，其中至少一条是 `rebuilt`。观察点必须与正在验收的 01 原文和模板字节一致，模板样式改变也需要该版本的观察点。重建默认为在线；验收历史事实重建时，观察点和命令分别使用 `rebuilt_kind: history` 与 `--rebuilt-kind history`。正式判据的数值来自相应评分报告。

## 六周交接

### 一分钟版

工作台把原文解析、事实入库、Excel 填写、报告和逐格追溯连在一起。最快的路径是在新环境生成合成样例，构建 Vue 页面，用 FastAPI 提供页面和 API。默认合成验收全程离线。公开数据的原文核对可以复用历史事实；新的 01 抽取需要本机配置 DeepSeek。

### 八行数字口径表

每项的机器判定值见权威 JSON，评分报告按执行时所选依据读取 `eval_report*.json` 或 `source_report*.json`。[判据配置](backend/data/source_expectations.json)

| 数字 | 数据来源 | 同机同库 | 空库重抽 | 与谁比 | 为什么 |
| --- | --- | --- | --- | --- | --- |
| 原文块数 | `/parse` 的 `blocks` 长度 | 逐位 | 逐位 | 判据 `block_count` | 依赖原文字节和解析器 |
| 行数 | `/fill` 的 `rows` | 逐位 | 逐位 | 判据 `rows` | 01 正文筛选和 03 日期、国家排序决定行 |
| 写入格数 | `/fill` 的 `cells_written` | 逐位 | 01 区间，03 逐位 | 判据 `cells_written` | 01 依赖可用事实 |
| 空着的格数 | `/fill` 的 `cells_empty` | 逐位 | 01 区间，03 逐位 | 判据 `cells_empty` | 直接取真实空格，03 另核对固定应空分母 |
| 单元格准确率 | `/reports` 的 `cell_accuracy` | 逐位 | 01 区间 | 判据 `cell_accuracy` | 01 正确格数除以总格数 |
| 行完整率 | `/reports` 的 `row_completeness` | 逐位 | 01 区间 | 判据 `row_completeness` | 01 五格均写入的行数除以总行数 |
| 耗时 | `X-A23-Fill-Seconds` 和 `fill_report*.json` | 记量级 | 记量级 | 验收记录中的本机耗时 | 写表和存储耗时随机器负载改变 |
| 抽取条数 | `/extract` 的 `count`、`kept` | 仅记录 | 仅记录 | 每次抽取日志和差异清单 | 模型返回会变化，在线调用单列 |

03 对格率和漏填率以应填格为分母，误填率以应空格为分母。`wrong_values` 统计应填但填错的格，`missed` 统计应填却留空的格。评分差异保留坐标和实际值。生成完整六字段表的命令如下。

```sh
.venv/bin/python backend/acceptance.py --print-criteria --basis source
```

### 从零到有数

需要 Git、Python 3.14 和 Node.js `^22.18.0 || >=24.12.0`。以下为 macOS/Linux，除克隆和进入目录外都在仓库根目录执行。Windows 用 `py -3.14` 建环境、`.venv\Scripts\python.exe` 运行脚本，进入 `frontend` 后调用 `npm.cmd`，启动前用 PowerShell 的 `$env:变量名` 设置环境变量。

1. 克隆并进入项目。`git clone https://github.com/ADGaiNai6657/A23.git && cd A23`
2. 建独立环境。`python3.14 -m venv .venv`
3. 安装锁定依赖。`.venv/bin/python -m pip install -r backend/requirements.txt`
4. 确定演示数据目录。`export A23_DATA_DIR="$(mktemp -d)"`。仅真实抽取时在 `backend/.env` 配置 `DEEPSEEK_API_KEY`。上述合成路径无需密钥。
5. 安装前端并构建。`(cd frontend && npm ci && npm run build)`
6. 准备原文、模板和参考表。`.venv/bin/python backend/demo.py --inputs-only`
7. 从命令行建库并填表。`.venv/bin/python backend/demo.py`。它对 01 使用固定模型替身并抽取入库，对 03 直接导入。若要演示 03 空库上传，可跳过此步。
8. 启动。`A23_DOC_DIR="$A23_DATA_DIR/demo/doc" A23_REFERENCE_DIR="$A23_DATA_DIR/demo/doc" A23_SYNTHETIC_REFERENCE=1 A23_DATA_DIR="$A23_DATA_DIR/demo/data" .venv/bin/python -m uvicorn main:app --app-dir backend --port 8000`
9. 打开工作台。`http://127.0.0.1:8000`。在页面选择 01，上传 `$A23_DATA_DIR/demo/doc/合成城市报告.docx`，将模板名填为 `模板_合成城市.xlsx`，点「一键演示（已有库）」。期望 1 行、5 格、两个率均为 100%，填写耗时单列。选择 03 后上传同目录的 `合成逐日统计.xlsx`，模板名填 `模板_合成逐日.xlsx`，点「从上传开始」，期望 3 行、19 格、2 格应空。
10. 在第二个终端跑验收。`.venv/bin/python backend/verify.py`。先检查前端构建，再验证首页及实际静态资源、交付文件、合成双模式和临时破坏性测试。公开数据两种模式的命令见「接入测试集」。

### 已知局限

- 正式参考结果仍缺，训练测试集实测使用独立原文核对。03 的填写与原文核对均采用同一逐日工作簿，核对可发现转换与排序错误。
- 训练填表规则固定为 01/03；公开空气质量 Word 模板使用上面的独立上传链路。
- XLSX 全量解析，模型抽取只看前 200 个非空数据行。03 填写按日期范围直接处理结构化数据。
- 事实三关核对块号、原话和数字。模型仍可能抄错主体、统计口径或单位，须按来源追溯。
- 真正空库在线重抽存在波动。药品流通 TXT 本轮留档 26/24 条，修复清单同义词后重放为 26/25 条。
- `dist/`、工作数据库与密钥不入仓。换环境时从交接指定的目录准备材料和数据。
- 本轮跨平台验证范围为 macOS Apple Silicon；Windows 尚无实机证据。表 C 记录的是单人独立 Clone，第三方互验仍待执行。[验收记录与 A/B/C 表](docs/六周验收记录.md)

### 哪些数不能逐位

同库使用 `same-db` 的逐位值。空库重建使用 `rebuilt`，只允许 01 四项模型相关指标使用区间。两个历史实验各保留 400 条城市事实，归档摘要和对应评分支撑当前单点区间。区间内的差异保留在报告，超出区间报错。03 的行数、分母和指标仍逐位。耗时和抽取条数仅供记录。

### 出问题先看三处

1. 检查 `frontend/dist/index.html` 和 `/health`。前端依赖安装和构建均在 `frontend` 执行。
2. 真实抽取失败时检查 `backend/.env`、网络报错和 `extraction-runs/`。代理导致 PyPI TLS 中断时，可在环境允许直连的情况下对安装命令临时取消 `HTTP_PROXY` 和 `HTTPS_PROXY`。
3. 查 `A23_DATA_DIR` 的文档版本、事实和评分报告。模板目录看 `A23_DOC_DIR`，评分依据看 `A23_REFERENCE_DIR` 或 `A23_SOURCE_CHECK_ROOT`。评分与本次 `run_id`、工作簿校验和不符时，页面提示本次暂无评分。

三格追溯、两套场景的拒填与缺失例子、样本耗时和终端留证见验收记录。详情按 Excel 行列号请求，01 和 03 共用一个面板，可查各自结果的全部数据格。03 的 17,794 格按需查询，避免一次加载整份逐格报告。[六周验收记录](docs/六周验收记录.md) · [核心功能说明](docs/核心功能说明.md)

| 结果格 | 判定与值 | 原文证据 |
| --- | --- | --- |
| 公开 01 B2 | 对，上海 GDP 为 56,708.71 亿元 | 城市报告第 4 块／第 4 段，「上海以 56,708.71 亿元的 GDP 总量稳居全国第一」 |
| 合成 01 B2 | 值不符，填入 100，临时对照表指定 999 | 合成城市报告第 1 块／第 1 段，「GDP总量100亿元」 |
| 公开 03 F3 | 应空，Algeria 当天的每日检测数缺值 | 疫情源 XLSX 第 498 块，`Data!499`，按行和「每日检测数」表头可核查原格为空 |

公开 01 的两个率原文核对均为 100%。同库两次填写耗时约 0.01 秒，03 约 0.6 秒。完整测试覆盖改坏评分和恢复，交叉验收本轮采用单人替代路径。[终端记录](backend/data/deliverables/acceptance-failures.txt) · [表 A/B/C](docs/六周验收记录.md#交叉验收)

后续整改的在线空库重抽、TXT 复测及审计条目对照由本地复审记录保存；运行和公开样本复验命令以上文为准。

01 总量与部分周期、跨模板观察点和独立副本的回归见 `backend/test_followup_audit.py` 和 `backend/verify.py`。
