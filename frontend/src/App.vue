<script setup>
import { computed, nextTick, onMounted, ref } from 'vue'
import { api } from './api'

const layouts = {
  city: {
    label: '01 城市经济',
    doc: '2025年中国城市经济百强全景报告.docx',
    template: '2025年中国城市经济百强全景报告-模板.xlsx',
    output: 'filled_2025中国城市经济百强.xlsx',
    columns: 5,
    extension: /\.docx$/i,
    action: 'extract',
    guide: '上传城市报告 DOCX，抽取事实后填写 Excel。',
    rates: [['单元格准确率', 'cell_accuracy'], ['行完整率', 'row_completeness']],
  },
  daily: {
    label: '03 历史疫情逐日统计',
    doc: 'COVID-19全球数据集（节选）.xlsx',
    template: 'COVID-19 模板.xlsx',
    output: 'filled_covid19_byday.xlsx',
    columns: 7,
    extension: /\.xlsx$/i,
    action: 'ingest',
    guide: '导入逐日统计 XLSX，使用表中数据直接填写。',
    rates: [['对格率', 'correct_rate'], ['误填率', 'false_fill_rate'], ['漏填率', 'missed_rate']],
  },
}
const fileActions = [
  { id: 'parse', title: '查看解析块数', pending: '正在解析文档' },
  { id: 'ingest', title: '仅导入文档块', pending: '正在导入文档' },
  { id: 'extract', title: '抽取事实并入库', pending: '正在抽取事实' },
]
const decisions = { filled: '已填写', refused: '已拒填', missing: '没有记录' }
const refusalLabels = {
  missing_fact: '缺少抽取事实', incompatible_value: '数值或单位不符',
  source_blank: '原文缺值', conflict: '来源冲突',
}
const connections = { checking: '正在连接', online: '后端在线', offline: '后端未连接' }

const kind = ref('city')
const layout = computed(() => layouts[kind.value])
const source = ref(layout.value.doc)
const template = ref(layout.value.template)
const file = ref(null)
const fileInput = ref(null)
const dragging = ref(false)
const preview = ref([])
const metrics = ref('')
const connection = ref('checking')
const pending = ref(null)
const feedback = ref({})
const result = ref(null)
const report = ref(null)
const resultHeading = ref(null)
const row = ref(2)
const column = ref(1)
const detail = ref(null)

async function checkHealth() {
  try {
    connection.value = (await api.health()).ok ? 'online' : 'offline'
  } catch {
    connection.value = 'offline'
  }
}

onMounted(checkHealth)

function selectLayout() {
  source.value = layout.value.doc
  template.value = layout.value.template
  result.value = report.value = detail.value = null
  row.value = 2
  column.value = 1
  feedback.value.fill = feedback.value.detail = null
}

function selectFile(selected) {
  feedback.value.input = null
  if (!selected) return
  preview.value = []
  if (!/\.(txt|md|docx|xlsx)$/i.test(selected.name)) {
    file.value = null
    feedback.value.input = { error: true, text: '请选择 TXT、Markdown、DOCX 或 XLSX 文件。' }
    return
  }
  file.value = selected
}

function onPick(event) {
  selectFile(event.target.files[0])
  event.target.value = ''
}

function onDrop(event) {
  dragging.value = false
  if (pending.value) return
  if (event.dataTransfer.files.length !== 1) {
    feedback.value.input = { error: true, text: '请一次拖入一份文档。' }
    return
  }
  selectFile(event.dataTransfer.files[0])
}

function clearFile() {
  file.value = null
  preview.value = []
  feedback.value.input = null
  fileInput.value?.focus()
}

async function run(area, label, action) {
  if (pending.value) return
  feedback.value[area] = null
  pending.value = { area, label }
  try {
    await checkHealth()
    if (connection.value !== 'online') throw new Error('后端没有响应，请启动服务后重试。')
    await action()
  } catch (failure) {
    feedback.value[area] = { error: true, text: failure.message }
  } finally {
    pending.value = null
  }
}

function useFile(action) {
  if (!file.value) {
    feedback.value.input = { error: true, text: '请先选择一份文档。' }
    fileInput.value?.focus()
    return
  }
  return run('input', action.pending, async () => {
    const response = await api[action.id](file.value, metrics.value)
    let text
    if (action.id === 'extract') {
      text = extractionSummary(response)
    } else {
      const count = action.id === 'parse' ? response.blocks.length : response.blocks
      if (action.id === 'parse') preview.value = response.blocks.slice(0, 5)
      text = `${action.id === 'parse' ? '解析' : '导入'}完成，共 ${count} 个块。`
    }
    if (action.id !== 'parse') source.value = response.doc
    feedback.value.input = { text }
  })
}

function extractionSummary(response) {
  const reasons = Object.entries(response.rejected ?? {}).map(([reason, count]) => `${reason} ${count} 条`)
  return `抽到 ${response.count} 条，保留 ${response.kept} 条有证据的事实。${reasons.join('、')}`
}

function showResult(doc, summary, elapsed) {
  result.value = { doc, ...summary, elapsed }
  detail.value = null
  feedback.value.detail = null
  row.value = 2
  column.value = 1
}

function readScore() {
  const doc = source.value.trim()
  if (!doc) {
    feedback.value.fill = { error: true, text: '请输入库里的文档名。' }
    return
  }
  return run('fill', '正在读取自检报告', async () => {
    const score = await api.reports(doc, layout.value.output)
    showResult(doc, score, score.fill_seconds?.toFixed(2))
    report.value = score
    feedback.value.fill = { text: '已载入本次填写的自检报告。' }
  })
}

async function fill(mode = 'fill') {
  const demo = mode !== 'fill'
  if (demo && !file.value) {
    feedback.value.input = { error: true, text: '请先选择要演示的源文档。' }
    fileInput.value?.focus()
    return
  }
  if ((!demo && !source.value.trim()) || !template.value.trim()) {
    feedback.value.fill = { error: true, text: '请输入文档名和模板文件名。' }
    return
  }
  const full = mode === 'full'
  if (full && !layout.value.extension.test(file.value.name)) {
    feedback.value.fill = { error: true, text: '上传的文档格式与所选模板类型不匹配。' }
    return
  }
  if (full && kind.value === 'city' && !window.confirm(
    '将文档分批发送到 DeepSeek，需要配置 API Key，可能等待数分钟并产生 API 费用。继续吗？',
  )) return
  await run('fill', demo ? '正在解析文档' : '正在填写表格', async () => {
    result.value = report.value = detail.value = null
    feedback.value.detail = null
    let sourceSha256
    if (demo) {
      const parsed = await api.parse(file.value)
      if (!parsed.source_sha256) throw new Error('后端未返回原文版本，请重启服务后重试。')
      preview.value = parsed.blocks.slice(0, 5)
      source.value = parsed.file
      sourceSha256 = parsed.source_sha256
      feedback.value.input = { text: `解析完成，共 ${parsed.blocks.length} 个块。` }
      if (full) {
        const extract = layout.value.action === 'extract'
        pending.value.label = extract ? '正在调用 DeepSeek 抽取并入库' : '正在导入文档块'
        const saved = await api[layout.value.action](file.value, metrics.value)
        feedback.value.input = {
          text: extract ? extractionSummary(saved) : `已入库，共 ${saved.blocks} 个块。`,
        }
      }
      pending.value.label = '正在填写表格'
    }
    const doc = source.value.trim()
    const started = performance.now()
    const summary = await api.fill(doc, template.value.trim(), sourceSha256)
    showResult(doc, summary, (summary.fill_seconds ?? (performance.now() - started) / 1000).toFixed(2))
    pending.value.label = '正在读取自检报告'
    try {
      report.value = await api.reports(doc, layout.value.output, summary.result_id)
    } catch (failure) {
      if (failure.status !== 404) throw failure
    }
    feedback.value.fill = { text: `填写完成，共 ${summary.rows} 行。` }
  })
  if (result.value && !feedback.value.fill?.error) {
    await nextTick()
    resultHeading.value?.focus()
  }
}

function inspectCell() {
  detail.value = null
  return run('detail', '正在查询单元格', async () => {
    detail.value = await api.cell(
      result.value.doc, layout.value.output, row.value, column.value, result.value.result_id,
    )
  })
}

function percent(value) {
  return value == null ? '—' : `${(value * 100).toFixed(2)}%`
}
</script>

<template>
  <a class="skip-link" href="#workspace">跳到工作区</a>

  <header class="topbar">
    <div class="topbar-inner">
      <div class="brand">
        <span class="brand-mark" aria-hidden="true">A23</span>
        <span>文档工作台</span>
      </div>
      <div class="toolbar">
        <span class="connection" :data-state="connection" role="status">
          <span class="connection-dot" aria-hidden="true"></span>
          {{ connections[connection] }}
        </span>
        <a class="docs-link" :href="api.docs" target="_blank" rel="noopener" aria-label="接口文档（在新标签页打开）">
          接口文档 <span aria-hidden="true">↗</span>
        </a>
      </div>
    </div>
  </header>

  <main id="workspace" class="workspace" tabindex="-1">
    <header class="page-title">
      <h1>文档理解与填表</h1>
      <p>从源文件提取事实，填写 Excel，再逐格核对出处。</p>
    </header>

    <div class="workflow">
      <section class="panel panel-input" aria-labelledby="input-heading">
        <header class="panel-heading">
          <span class="panel-icon" aria-hidden="true"><span class="icon icon-document"></span></span>
          <div>
            <h2 id="input-heading">文档入库</h2>
            <p>解析可预览文档结构，入库后可用于填表。</p>
          </div>
        </header>

        <fieldset class="input-controls" :disabled="Boolean(pending)">
          <div
            class="upload"
            :class="{ 'is-dragging': dragging, 'has-file': file }"
            @dragover.prevent="dragging = !pending"
            @dragleave="!$event.currentTarget.contains($event.relatedTarget) && (dragging = false)"
            @drop.prevent="onDrop"
          >
            <label class="upload-target" for="source-file">
              <span class="upload-icon" aria-hidden="true"><span class="icon icon-upload">↑</span></span>
              <span class="upload-copy">
                <span class="upload-title">{{ file ? file.name : '选择文档' }}</span>
                <span class="upload-hint" id="file-hint">{{ file ? '点击或拖入新文件可替换' : '拖放或点击选择 · TXT、Markdown、DOCX、XLSX' }}</span>
              </span>
              <input
                id="source-file"
                ref="fileInput"
                class="file-input"
                type="file"
                accept=".txt,.md,.docx,.xlsx"
                aria-label="选择文档"
                aria-describedby="file-hint"
                @change="onPick"
              >
            </label>
            <button v-if="file" class="remove-file" type="button" aria-label="移除已选文档" @click="clearFile">×</button>
          </div>

          <label class="field" for="metrics">
            <span>指定抽取指标 <span class="optional">可选</span></span>
            <input id="metrics" v-model="metrics" type="text" placeholder="例如 GDP, 常住人口">
          </label>

          <div class="file-actions">
            <button
              v-for="action in fileActions"
              :key="action.id"
              class="button"
              :class="action.id === layout.action ? 'button-primary' : 'button-secondary'"
              type="button"
              @click="useFile(action)"
            >{{ action.title }}</button>
          </div>
        </fieldset>

        <p v-if="feedback.input?.error" class="feedback feedback-error" role="alert">{{ feedback.input.text }}</p>
        <p v-else class="feedback" role="status">{{ pending?.area === 'input' ? `${pending.label}…` : feedback.input?.text }}</p>
        <details v-if="preview.length" class="block-preview" open>
          <summary>前 {{ preview.length }} 个文档块</summary>
          <ol>
            <li v-for="(block, index) in preview" :key="index" :value="block.locator.line ?? index + 1">
              {{ block.text }}
            </li>
          </ol>
        </details>
      </section>

      <section class="panel panel-fill" aria-labelledby="fill-heading">
        <header class="panel-heading">
          <span class="panel-icon" aria-hidden="true"><span class="icon icon-grid"></span></span>
          <div>
            <h2 id="fill-heading">填写与验收</h2>
            <p>{{ layout.guide }}</p>
          </div>
        </header>

        <form @submit.prevent="fill()">
          <fieldset class="fill-controls" :disabled="Boolean(pending)">
            <fieldset class="scenario">
              <legend>模板类型</legend>
              <div class="scenario-options">
                <label v-for="(option, key) in layouts" :key="key" class="scenario-option">
                  <input v-model="kind" type="radio" name="scenario" :value="key" @change="selectLayout">
                  <span>{{ option.label }}</span>
                </label>
              </div>
            </fieldset>
            <div class="fill-fields">
              <label class="field" for="doc-name">
                <span>库里的文档名</span>
                <input id="doc-name" v-model="source" type="text" required>
              </label>
              <label class="field" for="template-name">
                <span>模板文件名</span>
                <input id="template-name" v-model="template" type="text" placeholder="模板_文件名.xlsx" required>
              </label>
            </div>
            <div class="fill-actions">
              <button class="button button-secondary" type="button" @click="fill('prepared')">一键演示（已有库）</button>
              <button class="button button-secondary" type="button" @click="fill('full')">从上传开始</button>
              <button class="button button-secondary" type="button" @click="readScore">读取报告</button>
              <button class="button button-primary" type="submit">
                开始填写 <span aria-hidden="true">→</span>
              </button>
            </div>
          </fieldset>
        </form>
        <p class="feedback">已有库演示要求上传的原文已{{ kind === 'city' ? '抽取事实' : '导入' }}。从上传开始会{{ kind === 'city' ? '调用 DeepSeek 并产生 API 用量' : '导入 XLSX，无需调用模型' }}。</p>

        <p v-if="feedback.fill?.error" class="feedback feedback-error" role="alert">{{ feedback.fill.text }}</p>
        <p v-else class="feedback" role="status">{{ pending?.area === 'fill' ? `${pending.label}…` : feedback.fill?.text }}</p>
      </section>
    </div>

    <section class="panel results" aria-labelledby="result-heading">
      <header class="panel-heading result-header">
        <span class="panel-icon" aria-hidden="true"><span class="icon icon-result"></span></span>
        <div>
          <h2 id="result-heading" ref="resultHeading" tabindex="-1">填写结果</h2>
          <p v-if="result" class="result-document">{{ result.doc }}</p>
          <p v-else>填写完成后可下载 Excel，并检查每个单元格的来源。</p>
        </div>
        <a v-if="result" class="button button-secondary download" :href="api.download(layout.output, result.result_id)" download>
          <span aria-hidden="true">↓</span> 下载填写结果
        </a>
      </header>

      <template v-if="result">
        <dl class="stats">
          <div><dt>数据行</dt><dd>{{ result.rows }}</dd></div>
          <div><dt>写入格数</dt><dd>{{ result.cells_written }}</dd></div>
          <div><dt>空格</dt><dd>{{ result.cells_empty }}</dd></div>
          <div><dt>填写耗时</dt><dd>{{ result.elapsed ?? '—' }} <span v-if="result.elapsed != null">秒</span></dd></div>
        </dl>
        <div v-if="report" class="report">
          <p>{{ report.basis === 'independent-source-check' ? '独立原文核对' : '参考表评分' }}</p>
          <dl class="rates">
            <div v-for="[label, key] in layout.rates" :key="key">
              <dt>{{ label }}</dt>
              <dd>{{ percent(report[key]) }}</dd>
            </div>
          </dl>
        </div>
        <p v-else class="report-empty">{{ result.score_failed ? '本次自检生成失败，请检查参考表和核对材料。' : '本次填写暂无匹配的自检报告。' }}</p>
        <p v-if="report && Object.keys(report.refusal_counts ?? {}).length" class="report-empty">
          拒填分类：{{ Object.entries(report.refusal_counts).map(([key, count]) => `${refusalLabels[key]} ${count} 格`).join('、') }}
        </p>

        <div class="trace">
          <form @submit.prevent="inspectCell">
            <fieldset :disabled="Boolean(pending)">
              <legend>逐格追溯</legend>
              <div class="trace-fields">
                <label class="field" for="row">
                  <span>Excel 行号</span>
                  <input id="row" v-model.number="row" type="number" min="2" :max="result.rows + 1" required>
                </label>
                <label class="field" for="column">
                  <span>Excel 列号</span>
                  <input id="column" v-model.number="column" type="number" min="1" :max="layout.columns" required>
                </label>
                <button class="button button-secondary" type="submit">查看详情</button>
              </div>
            </fieldset>
          </form>
          <p v-if="feedback.detail?.error" class="feedback feedback-error" role="alert">{{ feedback.detail.text }}</p>
          <p v-else class="feedback" role="status">{{ pending?.area === 'detail' ? `${pending.label}…` : '' }}</p>
          <div aria-live="polite">
            <div v-if="detail" class="detail">
              <div class="detail-heading">
                <strong class="decision" :data-decision="detail.decision">{{ decisions[detail.decision] }}</strong>
                <span v-if="detail.value != null" class="detail-value">填写值 {{ detail.value }}</span>
              </div>
              <p>{{ detail.reason }}</p>
              <p v-if="detail.evaluation">本次自检判定 {{ detail.evaluation }}</p>
              <p v-if="detail.evidence_status === 'stale'">原文已变更，请重新填写后核查来源。</p>
              <p v-if="detail.source">{{ detail.source.doc }} · 第 {{ detail.source.block_no }} 块</p>
              <p v-if="detail.source?.locator">
                <template v-if="detail.source.locator.sheet">
                  {{ detail.source.locator.sheet }} · 第 {{ detail.source.locator.row }} 行
                </template>
                <template v-else-if="detail.source.locator.paragraph">第 {{ detail.source.locator.paragraph }} 段</template>
              </p>
              <blockquote v-if="detail.text">{{ detail.text }}</blockquote>
            </div>
          </div>
        </div>
      </template>
    </section>
  </main>
</template>
