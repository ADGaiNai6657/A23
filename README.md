# A23 文档理解与填表

第十七届服创大赛 A23 六周挑战档训练项目。支持四种输入、事实抽取与多源复核、01 城市经济和 03 历史疫情模板填写，以及浏览器逐格追溯。[开发规格](docs/开发规格.md)

## 安装和运行

在项目根目录执行。项目内已有 `.venv` 时，从第二条命令继续。Node.js 需满足 `^22.18.0 || >=24.12.0`。

```sh
python3.14 -m venv .venv
.venv/bin/python -m pip install -r backend/requirements.txt
npm --prefix frontend ci
npm --prefix frontend run build
(cd backend && ../.venv/bin/python -m uvicorn main:app)
```

打开 <http://127.0.0.1:8000> 使用页面，或打开 <http://127.0.0.1:8000/docs> 调试接口。前端改动后重新构建。后端启动时没有前端构建产物，仍可访问 API。

调用真实 DeepSeek 前，在本机创建 `backend/.env` 并设置 `DEEPSEEK_API_KEY`。抽取接口会将分批后的源文档文本发送给 DeepSeek。合成演示、解析、仅导入、03 逐日表填写和回归测试无需密钥。

## 用合成样例演示

```sh
.venv/bin/python backend/demo.py
.venv/bin/python backend/acceptance.py
(cd backend && ../.venv/bin/python -m pytest -q)
```

演示在 `backend/data/demo/` 内生成输入、SQLite、模板副本和两套报告。01 事实来自可复现的合成返回，03 使用合成 XLSX 数据源。页面也可读取这套结果，从项目根目录执行以下命令，再使用表中的文件名。

```sh
cd backend
A23_DOC_DIR=data/demo/doc A23_DATA_DIR=data/demo/data ../.venv/bin/python -m uvicorn main:app
```

| 模板 | 库里的文档名 | 模板文件名 | 合成验收数据 |
| --- | --- | --- | --- |
| 01 城市经济 | `合成城市报告.docx` | `模板_合成城市.xlsx` | 1 行、写入 5 格，单元格准确率与行完整率均为 100% |
| 03 历史疫情 | `合成逐日统计.xlsx` | `模板_合成逐日.xlsx` | 3 行、19 格正确、2 格真空、误填与漏填均为 0 |

点「开始填写」后下载结果，或输入 Excel 行号和列号查询来源。01 的第 2 行第 2 列显示 GDP 的源段落。03 的第 3 行第 3 列展示有效零值，第 3 行第 4 列展示缺值和拒填理由。

## 样本到位后的验收

1. 将原始材料放入与 `backend/` 同级的 `doc/`，保留任务书中的层级和文件名。
2. 在页面或 `/docs` 上传并解析四种格式。01 对城市报告执行「抽取事实并入库」。03 对逐日 XLSX 执行「仅导入文档块」。可通过 `GET /facts` 按主体、指标、期次、口径查询候选和采用结果。明确来源文档时可以传 `preferred_doc`，其他候选仍可查询。
3. 按任务书填写模板、下载结果并核对证据。01 需要人工对照选定的事实及单位，03 需要核对日期区间、国家顺序、真空和所有拒填分类。
4. 用下列脚本分别跑 01 和 03，传入库内文档名、原模板路径和参考路径。刷新页面查看报告。页面的率取自最近一次独立自检。

```sh
.venv/bin/python backend/check_fill.py --scenario 01 '库内文档名.docx' '模板路径.xlsx' '参考路径.xlsx'
.venv/bin/python backend/check_fill.py --scenario 03 '库内文档名.xlsx' '模板路径.xlsx' '参考路径.xlsx'
```

`backend/check_extract.py` 支持对指定文件跑通用和清单模式并导出事实。XLSX 全量解析入库，DeepSeek 只处理前 200 条非空数据行。03 的模板结果直接读取目标日期区间的结构化数据。这个范围选择控制模型调用量，范围外的 XLSX 行仍可用于确定性填表。

正式验收用一份 JSON 记录两个场景的判据，顶层键为 `same-db` 和 `rebuilt`，各含 `01`、`03` 对应的字段和值。区间写成 `[下界, 上界]`。同库复验使用精确值，重建后受模型影响的数值区间由实测确定。执行 `.venv/bin/python backend/acceptance.py --expectations 判据.json --mode same-db` 或将模式改为 `rebuilt`。失败时退出码为 1，并打印失败项、预期值与实测值。

原始样本和参考表尚未到位，当前验收数字来自合成材料。开发规格中的正式块数、行数、分母及跨环境复现需在原始材料上逐项记录。[第四周任务](六周任务/第四周任务.md) · [第六周任务](六周任务/第六周任务.md)
