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
3. 在「填写与验收」选择 01 或 03，核对库里的文档名和模板名，点「开始填写」。完成后可以下载 Excel，以及按行列号查值、原文和拒填理由。页面上的准确率来自最近一次用参考表运行的独立自检。
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

正式材料到位后，按任务书把它们放入 `doc/` 并保留目录层级。在页面将所需文档入库后，用下列脚本填写并评估。根据实际文件名替换三个位置参数。页面刷新后可看到最近一次独立自检的结果。

```sh
.venv/bin/python backend/check_fill.py --scenario 01 '库内文档名.docx' '模板路径.xlsx' '参考路径.xlsx'
.venv/bin/python backend/check_fill.py --scenario 03 '库内文档名.xlsx' '模板路径.xlsx' '参考路径.xlsx'
```

若服务使用自定义数据目录，传入同一个 `--data-dir`，例如 `--data-dir /var/lib/a23`。正式验收判据 JSON 的顶层为 `same-db` 和 `rebuilt`，各含 `01`、`03` 的字段与期望值。同库复验用精确值，重建后受模型影响的数值用实测区间 `[下界, 上界]`。执行 `.venv/bin/python backend/acceptance.py --expectations 判据.json --mode same-db` 或将模式改为 `rebuilt`，失败时脚本打印预期值与实测值并返回退出码 1。

已验证原始样本的块数、行数、分母和逐格原文一致性。相对正式参考表的准确率及正式同库与重建库验收等参考表补齐后记录。[第四周任务](六周任务/第四周任务.md) · [第六周任务](六周任务/第六周任务.md)
