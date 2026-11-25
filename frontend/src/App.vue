<script setup>
import { computed, onMounted, ref } from 'vue'
import { api } from './api'

const city = {
  doc: '输入_2025中国城市经济百强全景报告.docx',
  template: '模板_2025中国城市经济百强.xlsx',
  output: 'filled_2025中国城市经济百强.xlsx',
  columns: 5,
}
const daily = {
  doc: '输入_COVID-19全球数据集_节选.xlsx',
  template: '',
  output: 'filled_covid19_byday.xlsx',
  columns: 7,
}
const kind = ref('city')
const source = ref(city.doc)
const template = ref(city.template)
const layout = computed(() => kind.value === 'city' ? city : daily)
const file = ref(null)
const metrics = ref('')
const online = ref(false)
const pending = ref('')
const error = ref('')
const notice = ref('')
const result = ref(null)
const report = ref(null)
const elapsed = ref(null)
const row = ref(2)
const column = ref(1)
const detail = ref(null)
const activeDoc = ref('')

async function checkHealth() {
  try {
    online.value = (await api.health()).ok
  } catch {
    online.value = false
  }
}

onMounted(checkHealth)

function selectLayout() {
  source.value = layout.value.doc
  template.value = layout.value.template
  activeDoc.value = ''
  result.value = report.value = detail.value = null
  notice.value = error.value = ''
}

async function run(label, action) {
  error.value = notice.value = ''
  pending.value = label
  try {
    await checkHealth()
    if (!online.value) throw new Error('后端没有响应，请先启动服务')
    await action()
  } catch (failure) {
    error.value = failure.message
  } finally {
    pending.value = ''
  }
}

function useFile(operation) {
  if (!file.value) {
    error.value = '请先选择一份文档'
    return
  }
  run({ parse: '解析中', ingest: '导入中', extract: '抽取中' }[operation], async () => {
    const response = await api[operation](file.value, metrics.value)
    if (operation === 'extract') {
      notice.value = `抽到 ${response.count} 条，保留 ${response.kept} 条有证据的事实。`
    } else {
      const count = operation === 'parse' ? response.blocks.length : response.blocks
      notice.value = `${operation === 'parse' ? '解析' : '导入'}完成，共 ${count} 个块。`
    }
    if (operation !== 'parse') source.value = response.doc
  })
}

async function fill() {
  if (!source.value.trim() || !template.value.trim()) {
    error.value = '请输入文档名和模板文件名'
    return
  }
  result.value = report.value = detail.value = null
  await run('填写中', async () => {
    activeDoc.value = source.value.trim()
    const started = performance.now()
    result.value = await api.fill(activeDoc.value, template.value.trim())
    elapsed.value = ((performance.now() - started) / 1000).toFixed(2)
    try {
      report.value = await api.reports(activeDoc.value, layout.value.output)
    } catch (failure) {
      if (failure.status !== 404) throw failure
      notice.value = '表格已填写，暂无这份文档的自检报告。'
    }
  })
}

function inspectCell() {
  detail.value = null
  return run('查询中', async () => {
    detail.value = await api.cell(activeDoc.value, layout.value.output, row.value, column.value)
  })
}

function percent(value) {
  return value == null ? '—' : `${(value * 100).toFixed(2)}%`
}
</script>

<template>
  <main>
    <header>
      <h1>文档理解与填表</h1>
      <p :class="online ? 'online' : 'offline'">{{ online ? '后端在线' : '后端未连接' }}</p>
    </header>
    <p v-if="pending" role="status">{{ pending }}…</p>
    <p v-if="notice" role="status">{{ notice }}</p>
    <p v-if="error" class="offline" role="alert">{{ error }}</p>

    <section>
      <h2>文档入库</h2>
      <fieldset :disabled="Boolean(pending)">
        <label>选择文档 <input type="file" accept=".txt,.md,.docx,.xlsx" @change="file = $event.target.files[0]"></label>
        <label>指定抽取指标 <input v-model="metrics" placeholder="可选，逗号分隔"></label>
        <div class="actions">
          <button @click="useFile('parse')">查看解析块数</button>
          <button @click="useFile('ingest')">仅导入文档块</button>
          <button class="primary" @click="useFile('extract')">抽取事实并入库</button>
        </div>
      </fieldset>
    </section>

    <section>
      <h2>填写与验收</h2>
      <form @submit.prevent="fill">
        <fieldset :disabled="Boolean(pending)">
          <label>模板类型
            <select v-model="kind" @change="selectLayout">
              <option value="city">01 城市经济</option>
              <option value="daily">03 历史疫情逐日统计</option>
            </select>
          </label>
          <label>库里的文档名 <input v-model="source" required></label>
          <label>模板文件名 <input v-model="template" required></label>
          <button class="primary">开始填写</button>
        </fieldset>
      </form>
      <div v-if="result">
        <dl>
          <div><dt>数据行</dt><dd>{{ result.rows }}</dd></div>
          <div><dt>写入</dt><dd>{{ result.cells_written }}</dd></div>
          <div><dt>空格</dt><dd>{{ result.cells_empty }}</dd></div>
          <div><dt>填写耗时</dt><dd>{{ elapsed }} 秒</dd></div>
        </dl>
        <div v-if="report">
          <p>以下数值来自最近一次自检。</p>
          <p v-if="kind === 'city'">单元格准确率 {{ percent(report.cell_accuracy) }} · 行完整率 {{ percent(report.row_completeness) }}</p>
          <p v-else>对格率 {{ percent(report.correct_rate) }} · 误填率 {{ percent(report.false_fill_rate) }} · 漏填率 {{ percent(report.missed_rate) }}</p>
        </div>
        <a :href="api.download(layout.output)" download>下载填写结果</a>

        <form class="inspect" @submit.prevent="inspectCell">
          <fieldset :disabled="Boolean(pending)">
            <legend>逐格追溯</legend>
            <label>Excel 行号 <input v-model.number="row" type="number" min="2" :max="result.rows + 1" required></label>
            <label>Excel 列号 <input v-model.number="column" type="number" min="1" :max="layout.columns" required></label>
            <button>查看详情</button>
          </fieldset>
        </form>
        <div v-if="detail" class="detail">
          <p>{{ { filled: '已填写', refused: '已拒填', missing: '没有记录' }[detail.decision] }} · {{ detail.reason }}</p>
          <p v-if="detail.value != null">填写值 {{ detail.value }}</p>
          <p v-if="detail.evaluation">最近自检判定 {{ detail.evaluation }}</p>
          <p v-if="detail.source">{{ detail.source.doc }} · 第 {{ detail.source.block_no }} 块</p>
          <blockquote v-if="detail.text">{{ detail.text }}</blockquote>
        </div>
      </div>
    </section>
  </main>
</template>

<style>
:root { font: 16px/1.55 system-ui, sans-serif; color: #173129; background: #f0f4ef; }
* { box-sizing: border-box; }
body { margin: 0; }
main { max-width: 850px; margin: auto; padding: 2rem 1rem 4rem; }
header { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: baseline; gap: .5rem; }
h1 { font-size: clamp(1.8rem, 5vw, 2.6rem); margin: 0 0 1rem; }
h2 { margin-top: 0; font-size: 1.25rem; }
section { background: white; border: 1px solid #cdd9d0; border-radius: 12px; margin-bottom: 1rem; padding: 1.25rem; }
fieldset { border: 0; padding: 0; margin: 0; min-width: 0; }
label { display: block; margin-bottom: .7rem; }
input, select, button { font: inherit; border: 1px solid #9caf9f; border-radius: 6px; padding: .45rem .7rem; background: #fff; color: inherit; }
label input, label select { display: block; width: 100%; margin-top: .25rem; }
.actions { display: flex; flex-wrap: wrap; gap: .5rem; }
button, a { cursor: pointer; }
.primary { background: #174c37; color: white; }
.online, a { color: #176047; }
.offline { color: #9d261a; }
dl { display: grid; grid-template-columns: repeat(auto-fit, minmax(96px, 1fr)); gap: .5rem; }
dl div, .detail { background: #edf4ee; padding: .6rem .8rem; border-radius: 6px; }
dd { margin: 0; font-size: 1.35rem; }
.inspect { margin-top: 1rem; }
.inspect fieldset { display: flex; flex-wrap: wrap; gap: .5rem; align-items: start; }
.inspect input { max-width: 120px; }
.detail { margin-top: .75rem; overflow-wrap: anywhere; }
blockquote { border-left: 3px solid #8aab98; margin-left: 0; padding-left: .8rem; white-space: pre-wrap; }
</style>
