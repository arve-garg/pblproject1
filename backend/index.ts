import cors from 'cors'
import express from 'express'
import multer from 'multer'
import { parse } from 'csv-parse/sync'
import * as XLSX from 'xlsx'
import os from 'node:os'
import path from 'node:path'
import pdf from 'pdf-parse'
import fs from 'node:fs/promises'

const app = express()
const upload = multer({ dest: path.join(os.tmpdir(), 'lumen-uploads'), limits: { fileSize: 25 * 1024 * 1024, files: 10, fields: 20 } })
app.use(cors())
app.use(express.json({ limit: '2mb' }))

type Row = Record<string, unknown>
type Dataset = { fileName: string; rows: Row[]; columns: string[] }
const datasets = new Map<string, Dataset>()
const demoDatasetId = 'demo'
datasets.set(demoDatasetId, { fileName: 'northstar_sales_q3.csv', rows: [{ Revenue: 42000, Orders: 120, Segment: 'Enterprise' }, { Revenue: 55000, Orders: 140, Segment: 'Mid-market' }, { Revenue: 48000, Orders: 98, Segment: 'SMB' }], columns: ['Revenue', 'Orders', 'Segment'] })

function normalizeRows(rows: Row[]) {
  return rows.filter(Boolean).map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [key.trim(), value])))
}

async function readDataset(file: Express.Multer.File): Promise<Dataset> {
  const extension = file.originalname.toLowerCase().split('.').pop()
  const buffer = await fs.readFile(file.path)
  if (extension === 'csv') {
    const rows = parse(buffer.toString('utf8'), { columns: true, skip_empty_lines: true, bom: true, relax_column_count: true }) as Row[]
    const clean = normalizeRows(rows)
    return { fileName: file.originalname, rows: clean, columns: Object.keys(clean[0] ?? {}) }
  }
  if (extension === 'xlsx' || extension === 'xls') {
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true })
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    const rows = XLSX.utils.sheet_to_json<Row>(sheet, { defval: null })
    const clean = normalizeRows(rows)
    return { fileName: file.originalname, rows: clean, columns: Object.keys(clean[0] ?? {}) }
  }
  if (extension === 'pdf') {
    const extracted = await pdf(buffer)
    const lines = extracted.text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean)
    if (!lines.length) throw new Error(`No selectable text found in ${file.originalname}. Scanned/image-only PDFs need OCR before analysis.`)
    const rows = lines.map((text, index) => ({ line_number: index + 1, text }))
    return { fileName: file.originalname, rows, columns: ['line_number', 'text'] }
  }
  throw new Error('Unsupported format. Use CSV, XLSX, XLS, or PDF.')
}

function numeric(value: unknown) { if (value === null || value === undefined || String(value).trim() === '') return Number.NaN; return typeof value === 'number' ? value : Number(String(value).replace(/[$,%\s,]/g, '')) }
function profile(dataset: Dataset) {
  const columns = dataset.columns.map((name) => {
    const values = dataset.rows.map((row) => row[name]).filter((value) => value !== null && value !== undefined && value !== '')
    const numbers = values.map(numeric).filter(Number.isFinite)
    const isNumeric = numbers.length > values.length * 0.7 && numbers.length > 0
    const unique = [...new Set(values.map(String))]
    return { name, type: isNumeric ? 'numeric' : 'categorical', missing: dataset.rows.length - values.length, unique: unique.length, min: isNumeric ? Math.min(...numbers) : undefined, max: isNumeric ? Math.max(...numbers) : undefined, average: isNumeric ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : undefined, topValues: isNumeric ? undefined : unique.slice(0, 5).map((value) => ({ value, count: values.filter((item) => String(item) === value).length })) }
  })
  const numericColumns = columns.filter((column) => column.type === 'numeric')
  const firstNumeric = numericColumns[0]
  const series = firstNumeric ? dataset.rows.slice(0, 30).map((row, index) => ({ label: index + 1, value: numeric(row[firstNumeric.name]) })).filter((point) => Number.isFinite(point.value)) : []
  return { fileName: dataset.fileName, rows: dataset.rows.length, columns: dataset.columns.length, schema: columns, series, preview: dataset.rows.slice(0, 20) }
}

app.post('/api/analyze', upload.single('file'), async (request, response) => {
  try {
    if (!request.file) return response.status(400).json({ error: 'Attach a dataset in the file field.' })
    const dataset = await readDataset(request.file)
    const id = `${Date.now()}-${Math.random().toString(36).slice(2)}`
    datasets.set(id, dataset)
    return response.json({ id, analysis: profile(dataset) })
  } catch (error) { return response.status(400).json({ error: error instanceof Error ? error.message : 'Could not analyze file.' }) }
  finally { if (request.file?.path) await fs.rm(request.file.path, { force: true }).catch(() => undefined) }
})

app.post('/api/analyze-batch', upload.array('files', 10), async (request, response) => {
  const files = (request.files ?? []) as Express.Multer.File[]
  if (!files.length) return response.status(400).json({ error: 'Choose one or more CSV, Excel, or text-based PDF files.' })
  const results = await Promise.all(files.map(async (file) => {
    try {
      const dataset = await readDataset(file)
      const id = `${Date.now()}-${Math.random().toString(36).slice(2)}`
      datasets.set(id, dataset)
      return { id, analysis: profile(dataset) }
    } catch (error) {
      return { fileName: file.originalname, error: error instanceof Error ? error.message : 'Could not analyze file.' }
    } finally { await fs.rm(file.path, { force: true }).catch(() => undefined) }
  }))
  return response.json({ results, accepted: results.filter((item) => 'analysis' in item).length, failed: results.filter((item) => 'error' in item).length })
})

function getDataset(datasetId?: string) { return datasets.get(datasetId || demoDatasetId) }

app.post('/api/ask', (request, response) => {
  const { datasetId, question } = request.body as { datasetId?: string; question?: string }
  const dataset = getDataset(datasetId)
  if (!dataset || !question?.trim()) return response.status(400).json({ error: 'Select a dataset and enter a question.' })
  const summary = profile(dataset)
  const q = question.trim().toLowerCase()
  const steps: { agent: string; action: string; result: string }[] = []
  steps.push({ agent: 'Planner', action: 'Classify request', result: q.includes('missing') || q.includes('clean') || q.includes('quality') ? 'Data-quality inspection' : q.includes('column') || q.includes('field') || q.includes('schema') ? 'Schema inspection' : q.includes('preview') || q.includes('sample') || q.includes('row') ? 'Row preview' : q.includes('highest') || q.includes('lowest') || q.includes('average') || q.includes('mean') || q.includes('numeric') || q.includes('trend') || q.includes('signal') ? 'Numeric analysis' : 'Dataset overview' })
  steps.push({ agent: 'Data Profiler', action: 'Inspect active dataset', result: dataset.fileName + ': ' + summary.rows + ' rows, ' + summary.columns + ' columns' })
  let answer = ''
  if (q.includes('missing') || q.includes('clean') || q.includes('quality')) {
    const columns = summary.schema.filter((column) => column.missing > 0).sort((a, b) => b.missing - a.missing)
    const total = summary.schema.reduce((sum, column) => sum + column.missing, 0)
    answer = total ? 'I found ' + total + ' missing cells. ' + columns.map((column) => column.name + ': ' + column.missing + ' missing').join('; ') + '. Consider reviewing these fields before selecting a cleaning action.' : 'No missing cells were detected in the active dataset. You can still check for duplicate rows, whitespace, and inconsistent categories in the cleaning pipeline.'
    steps.push({ agent: 'Quality Agent', action: 'Count missing cells', result: total + ' missing cells across ' + columns.length + ' affected fields' })
  } else if (q.includes('column') || q.includes('field') || q.includes('schema')) {
    answer = 'The active dataset has ' + summary.columns + ' fields: ' + summary.schema.map((column) => column.name + ' (' + column.type + ', ' + column.unique + ' unique, ' + column.missing + ' missing)').join('; ') + '.'
    steps.push({ agent: 'Schema Agent', action: 'Summarize columns', result: summary.schema.length + ' field profiles reviewed' })
  } else if (q.includes('preview') || q.includes('sample') || q.includes('row')) {
    answer = 'Here is the available sample from ' + dataset.fileName + ': ' + JSON.stringify(dataset.rows.slice(0, 5))
    steps.push({ agent: 'Sample Agent', action: 'Read first rows', result: 'Returned ' + Math.min(5, dataset.rows.length) + ' sample rows' })
  } else if (q.includes('highest') || q.includes('lowest') || q.includes('average') || q.includes('mean') || q.includes('numeric') || q.includes('trend') || q.includes('signal')) {
    const numericColumns = summary.schema.filter((column) => column.type === 'numeric')
    const stats = numericColumns.map((column) => ({ name: column.name, min: column.min, max: column.max, average: column.average }))
    const target = q.includes('highest') ? stats.reduce((best, current) => !best || (current.max ?? -Infinity) > (best.max ?? -Infinity) ? current : best, undefined as typeof stats[number] | undefined) : q.includes('lowest') ? stats.reduce((best, current) => !best || (current.min ?? Infinity) < (best.min ?? Infinity) ? current : best, undefined as typeof stats[number] | undefined) : undefined
    answer = stats.length ? 'Numeric summary for ' + dataset.fileName + ': ' + stats.map((item) => item.name + ' — mean ' + Number(item.average).toFixed(2) + ', min ' + item.min + ', max ' + item.max).join('; ') + (target ? '. ' + (q.includes('highest') ? 'Largest maximum: ' : 'Smallest minimum: ') + target.name + '.' : '') : 'I could not identify numeric columns in this dataset. Check the detected schema or upload a structured CSV/Excel file.'
    steps.push({ agent: 'Statistics Agent', action: 'Calculate numeric summaries', result: stats.length + ' numeric fields summarized from the uploaded rows' })
  } else {
    answer = 'Dataset overview for ' + dataset.fileName + ': ' + summary.rows + ' rows and ' + summary.columns + ' columns. Fields: ' + summary.schema.map((column) => column.name + ' (' + column.type + ')').join(', ') + '. Try asking about missing values, column types, numeric averages/ranges, or a sample of rows.'
    steps.push({ agent: 'Insight Agent', action: 'Build overview', result: 'Summarized schema and dataset size' })
  }
  steps.push({ agent: 'Evidence Check', action: 'Ground response in data', result: 'Answer derived from the currently selected in-memory dataset; no external facts used' })
  return response.json({ answer, steps, evidence: { fileName: dataset.fileName, rows: summary.rows, columns: summary.columns, schema: summary.schema } })
})
app.post('/api/clean', (request, response) => {
  const { datasetId, options = {} } = request.body as { datasetId?: string; options?: { trimStrings?: boolean; removeEmptyRows?: boolean; dropMissingRows?: boolean; deduplicateRows?: boolean; fillNumericMissing?: boolean; normalizeHeaders?: boolean } }
  const dataset = datasets.get(datasetId || demoDatasetId)
  if (!dataset) return response.status(400).json({ error: 'Upload a dataset before cleaning it.' })
  const before = dataset.rows.length
  let rows = dataset.rows.map((row) => ({ ...row }))
  const steps: { name: string; detail: string; status: string; affected?: number }[] = []
  const record = (name: string, detail: string, affected = 0) => steps.push({ name, detail, affected, status: 'complete' })
  const missingCount = rows.reduce((sum, row) => sum + dataset.columns.filter((column) => row[column] === null || row[column] === undefined || String(row[column]).trim() === '').length, 0)
  record('Profile missing values', `Found ${missingCount} empty cells across ${dataset.columns.length} fields; source values are preserved unless an option below is selected.`)
  if (options.trimStrings) {
    let affected = 0
    rows = rows.map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => {
      if (typeof value === 'string' && value !== value.trim()) affected++
      return [key, typeof value === 'string' ? value.trim() : value]
    })))
    record('Trim text values', `Trimmed leading/trailing whitespace in text cells.`, affected)
  } else steps.push({ name: 'Trim text values', detail: 'Skipped by user.', status: 'skipped' })
  if (options.normalizeHeaders) {
    const renamed = new Map(dataset.columns.map((column) => [column, column.trim().replace(/\s+/g, '_')]))
    rows = rows.map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [renamed.get(key) ?? key, value])))
    record('Normalize column names', 'Trimmed headers and replaced internal spaces with underscores.', dataset.columns.filter((column) => renamed.get(column) !== column).length)
  } else steps.push({ name: 'Normalize column names', detail: 'Skipped by user.', status: 'skipped' })
  if (options.removeEmptyRows) {
    const count = rows.length
    rows = rows.filter((row) => Object.values(row).some((value) => value !== null && value !== undefined && String(value).trim() !== ''))
    record('Remove empty rows', `Removed ${count - rows.length} fully empty rows.`, count - rows.length)
  } else steps.push({ name: 'Remove empty rows', detail: 'Skipped by user.', status: 'skipped' })
  if (options.dropMissingRows) {
    const count = rows.length
    const activeColumns = Object.keys(rows[0] ?? {}).length ? Object.keys(rows[0]) : dataset.columns
    rows = rows.filter((row) => activeColumns.every((column) => row[column] !== null && row[column] !== undefined && String(row[column]).trim() !== ''))
    record('Drop rows with missing values', `Removed ${count - rows.length} rows containing at least one missing value.`, count - rows.length)
  } else steps.push({ name: 'Drop rows with missing values', detail: 'Skipped by user.', status: 'skipped' })
  if (options.fillNumericMissing) {
    let filled = 0
    for (const column of (Object.keys(rows[0] ?? {}).length ? Object.keys(rows[0]) : dataset.columns)) {
      const values = rows.map((row) => numeric(row[column])).filter((value) => Number.isFinite(value) && rowIsPresent(value))
      const nonEmpty = rows.map((row) => row[column]).filter((value) => value !== null && value !== undefined && String(value).trim() !== '')
      const numericColumn = values.length > 0 && values.length / Math.max(nonEmpty.length, 1) >= 0.7
      if (!numericColumn) continue
      const mean = values.reduce((sum, value) => sum + value, 0) / values.length
      rows = rows.map((row) => {
        const value = row[column]
        if (value === null || value === undefined || String(value).trim() === '') { filled++; return { ...row, [column]: Number(mean.toFixed(6)) } }
        return row
      })
    }
    record('Fill numeric missing values', `Filled blank cells in numeric columns with each column's mean.`, filled)
  } else steps.push({ name: 'Fill numeric missing values', detail: 'Skipped by user.', status: 'skipped' })
  if (options.deduplicateRows) {
    const count = rows.length
    const seen = new Set<string>()
    rows = rows.filter((row) => { const key = JSON.stringify(row); if (seen.has(key)) return false; seen.add(key); return true })
    record('Remove duplicate rows', `Removed ${count - rows.length} exact duplicate rows.`, count - rows.length)
  } else steps.push({ name: 'Remove duplicate rows', detail: 'Skipped by user.', status: 'skipped' })
  const columns = Object.keys(rows[0] ?? {}).length ? Object.keys(rows[0]) : dataset.columns
  const escapeCsv = (value: unknown) => { const text = value === null || value === undefined ? '' : String(value); return /[",\r\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text }
  const csv = [columns.map(escapeCsv).join(','), ...rows.map((row) => columns.map((column) => escapeCsv(row[column])).join(','))].join('\n')
  return response.json({ steps, rowsBefore: before, rowsAfter: rows.length, missingCellsBefore: missingCount, downloadName: 'cleaned-' + dataset.fileName.replace(/\.[^.]+$/, '') + '.csv', csv, audit: { options, sourceFile: dataset.fileName, rowsBefore: before, rowsAfter: rows.length, unchangedSource: true } })
})
function rowIsPresent(value: number) { return Number.isFinite(value) }
app.post('/api/train', (request, response) => {
  const { datasetId, target, algorithm, testSize } = request.body as { datasetId?: string; target?: string; algorithm?: string; testSize?: number }
  const dataset = getDataset(datasetId)
  if (!dataset) return response.status(404).json({ error: 'Dataset not found. Select or upload a dataset before evaluating.' })
  const targetColumn = target || dataset.columns.find((column) => profile(dataset).schema.find((item) => item.name === column)?.type === 'numeric')
  if (!targetColumn || !dataset.columns.includes(targetColumn)) return response.status(400).json({ error: `Target field "${targetColumn ?? ''}" was not found in this dataset.` })
  const values = dataset.rows.map((row) => numeric(row[targetColumn])).filter(Number.isFinite)
  if (values.length < 3) return response.status(400).json({ error: 'At least 3 numeric target values are required to evaluate a holdout baseline.' })
  const fraction = Math.min(0.5, Math.max(0.1, (testSize ?? 20) / 100))
  const testCount = Math.max(1, Math.min(values.length - 1, Math.ceil(values.length * fraction)))
  const trainValues = values.slice(0, values.length - testCount)
  const testValues = values.slice(values.length - testCount)
  const prediction = trainValues.reduce((sum, value) => sum + value, 0) / trainValues.length
  const mae = testValues.reduce((sum, value) => sum + Math.abs(value - prediction), 0) / testValues.length
  const actualMean = testValues.reduce((sum, value) => sum + value, 0) / testValues.length
  const totalSquares = testValues.reduce((sum, value) => sum + (value - actualMean) ** 2, 0)
  const residualSquares = testValues.reduce((sum, value) => sum + (value - prediction) ** 2, 0)
  const r2 = totalSquares === 0 ? null : 1 - residualSquares / totalSquares
  return response.json({ model: 'Mean baseline (not ML training)', target: targetColumn, testSize: testSize ?? 20, metrics: { r2: r2 === null ? null : Number(r2.toFixed(4)), mae: Number(mae.toFixed(4)), samples: values.length, trainSamples: trainValues.length, testSamples: testValues.length }, steps: [`Selected numeric target "${targetColumn}"`, `Used the first ${trainValues.length} values for baseline estimation`, `Evaluated the training mean on the final ${testValues.length} values`, 'Calculated MAE and R² from the holdout values'], note: 'This is a deterministic mean baseline, not a learned machine-learning model. The metric values are calculated from the selected dataset.' })
})
app.post('/api/classify', (request, response) => {
  const { datasetId, target } = request.body as { datasetId?: string; target?: string }
  const dataset = getDataset(datasetId)
  if (!dataset) return response.status(404).json({ error: 'Dataset not found. Upload a labelled dataset first.' })
  const targetColumn = String(target ?? '').trim()
  if (!targetColumn || !dataset.columns.includes(targetColumn)) return response.status(400).json({ error: 'Choose an existing target column containing class labels.' })

  const numericColumns = profile(dataset).schema.filter((column) => column.type === 'numeric' && column.name !== targetColumn).map((column) => column.name)
  if (!numericColumns.length) return response.status(400).json({ error: 'This classifier needs at least one numeric feature column besides the target. Encode categorical predictors or select a dataset with numeric features.' })

  const samples = dataset.rows.map((row) => ({ row, label: String(row[targetColumn] ?? '').trim() }))
    .filter((item) => item.label !== '' && item.label.toLowerCase() !== 'null' && item.label.toLowerCase() !== 'undefined' && item.label.toLowerCase() !== 'nan')
  const classes = [...new Set(samples.map((item) => item.label))].sort()
  if (classes.length < 2) return response.status(400).json({ error: 'The target must contain at least two different classes.' })
  const groups = new Map<string, typeof samples>(classes.map((label) => [label, samples.filter((item) => item.label === label)] as [string, typeof samples]))
  if (classes.some((label) => (groups.get(label)?.length ?? 0) < 5)) return response.status(400).json({ error: 'For a meaningful evaluation, each target class needs at least 5 labelled rows. Add more labelled examples and try again.' })
  if (samples.length < 30) return response.status(400).json({ error: 'Use at least 30 labelled rows for a more useful holdout evaluation.' })

  // Deterministic, stratified holdout. All preprocessing and model fitting use training rows only.
  let seed = 42
  const random = () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296 }
  const train: typeof samples = []
  const test: typeof samples = []
  for (const label of classes) {
    const group = [...(groups.get(label) ?? [])]
    for (let i = group.length - 1; i > 0; i--) { const j = Math.floor(random() * (i + 1)); [group[i], group[j]] = [group[j], group[i]] }
    const testCount = Math.max(1, Math.min(group.length - 1, Math.round(group.length * 0.2)))
    test.push(...group.slice(0, testCount))
    train.push(...group.slice(testCount))
  }
  // Impute missing numeric values with training-set medians; never inspect test labels when fitting.
  const medians: Record<string, number> = {}
  for (const column of numericColumns) {
    const values = train.map((item) => numeric(item.row[column])).filter(Number.isFinite).sort((a, b) => a - b)
    medians[column] = values.length ? values[Math.floor(values.length / 2)] : 0
  }
  const vector = (row: Row) => numericColumns.map((column) => {
    const value = numeric(row[column])
    return Number.isFinite(value) ? value : medians[column]
  })
  const trainVectors = train.map((item) => ({ x: vector(item.row), y: item.label }))
  const testVectors = test.map((item) => ({ x: vector(item.row), y: item.label }))
  type Tree = { leaf: string } | { feature: number; threshold: number; left: Tree; right: Tree }
  const gini = (labels: string[]) => {
    if (!labels.length) return 0
    const counts = new Map<string, number>()
    for (const label of labels) counts.set(label, (counts.get(label) ?? 0) + 1)
    return 1 - [...counts.values()].reduce((sum, count) => sum + (count / labels.length) ** 2, 0)
  }
  const majority = (items: { x: number[]; y: string }[]) => {
    const counts = new Map<string, number>()
    for (const item of items) counts.set(item.y, (counts.get(item.y) ?? 0) + 1)
    return [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0][0]
  }
  const maxFeatures = Math.max(1, Math.floor(Math.sqrt(numericColumns.length)))
  const buildTree = (items: { x: number[]; y: string }[], depth: number): Tree => {
    const labels = items.map((item) => item.y)
    if (new Set(labels).size <= 1 || depth >= 14 || items.length < 6) return { leaf: majority(items) }
    const parentGini = gini(labels)
    const featureIndexes = numericColumns.map((_, index) => index)
    for (let i = featureIndexes.length - 1; i > 0; i--) {
      const j = Math.floor(random() * (i + 1)); [featureIndexes[i], featureIndexes[j]] = [featureIndexes[j], featureIndexes[i]]
    }
    let best: { feature: number; threshold: number; gain: number } | undefined
    for (const feature of featureIndexes.slice(0, maxFeatures)) {
      const unique = [...new Set(items.map((item) => item.x[feature]))].sort((a, b) => a - b)
      if (unique.length < 2) continue
      const candidates: number[] = []
      const maxCandidates = 24
      const step = Math.max(1, Math.ceil((unique.length - 1) / maxCandidates))
      for (let i = step - 1; i < unique.length - 1; i += step) candidates.push((unique[i] + unique[i + 1]) / 2)
      if (candidates[candidates.length - 1] !== (unique[unique.length - 2] + unique[unique.length - 1]) / 2) candidates.push((unique[unique.length - 2] + unique[unique.length - 1]) / 2)
      for (const threshold of candidates) {
        const left = items.filter((item) => item.x[feature] <= threshold)
        const right = items.filter((item) => item.x[feature] > threshold)
        if (left.length < 2 || right.length < 2) continue
        const gain = parentGini - (left.length / items.length) * gini(left.map((item) => item.y)) - (right.length / items.length) * gini(right.map((item) => item.y))
        if (!best || gain > best.gain) best = { feature, threshold, gain }
      }
    }
    if (!best || best.gain < 1e-8) return { leaf: majority(items) }
    const left = items.filter((item) => item.x[best!.feature] <= best!.threshold)
    const right = items.filter((item) => item.x[best!.feature] > best!.threshold)
    return { feature: best.feature, threshold: best.threshold, left: buildTree(left, depth + 1), right: buildTree(right, depth + 1) }
  }
  const predictTree = (tree: Tree, x: number[]): string => {
    let node = tree
    while (!('leaf' in node)) node = x[node.feature] <= node.threshold ? node.left : node.right
    return node.leaf
  }
  // Random forest: bootstrap samples, random feature subsets at each split, majority vote.
  const treeCount = 120
  const forest: Tree[] = []
  for (let treeIndex = 0; treeIndex < treeCount; treeIndex++) {
    const bootstrap = Array.from({ length: trainVectors.length }, () => trainVectors[Math.floor(random() * trainVectors.length)])
    forest.push(buildTree(bootstrap, 0))
  }
  const predict = (x: number[]) => {
    const votes = new Map<string, number>()
    for (const tree of forest) {
      const label = predictTree(tree, x)
      votes.set(label, (votes.get(label) ?? 0) + 1)
    }
    return [...votes.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0][0]
  }
  const predictions = testVectors.map((item) => predict(item.x))
  const matrix = classes.map((actual) => classes.map((predicted) => testVectors.filter((item, index) => item.y === actual && predictions[index] === predicted).length))
  const correct = testVectors.filter((item, index) => predictions[index] === item.y).length
  const perClass = classes.map((label, index) => {
    const tp = matrix[index][index]
    const fp = matrix.reduce((sum, row, rowIndex) => sum + (rowIndex === index ? 0 : row[index]), 0)
    const fn = matrix[index].reduce((sum, count, colIndex) => sum + (colIndex === index ? 0 : count), 0)
    const precision = tp + fp ? tp / (tp + fp) : 0
    const recall = tp + fn ? tp / (tp + fn) : 0
    return { class: label, precision: Number(precision.toFixed(4)), recall: Number(recall.toFixed(4)), f1: Number((precision + recall ? 2 * precision * recall / (precision + recall) : 0).toFixed(4)), support: matrix[index].reduce((sum, count) => sum + count, 0) }
  })
  const accuracy = correct / testVectors.length
  return response.json({
    model: 'Random Forest classifier (120 trees)',
    target: targetColumn,
    features: numericColumns,
    split: 'Stratified 80/20 holdout (seed 42)',
    metrics: { accuracy: Number(accuracy.toFixed(4)), correct, testSamples: testVectors.length, trainSamples: trainVectors.length, macroPrecision: Number((perClass.reduce((sum, item) => sum + item.precision, 0) / classes.length).toFixed(4)), macroRecall: Number((perClass.reduce((sum, item) => sum + item.recall, 0) / classes.length).toFixed(4)), macroF1: Number((perClass.reduce((sum, item) => sum + item.f1, 0) / classes.length).toFixed(4)) },
    classes, confusionMatrix: matrix, perClass,
    note: 'Random Forest reduces the single-model bias of the previous Gaussian Naive Bayes baseline. Scores are from held-out rows for this dataset and split; high precision is not guaranteed and depends on data quality, predictive features, class labels, and representativeness.'
  })
})

app.post('/api/predict', (request, response) => {
  const { datasetId, target, inputs } = request.body as { datasetId?: string; target?: string; inputs?: Record<string, unknown> }
  const dataset = getDataset(datasetId)
  if (!dataset) return response.status(400).json({ error: 'Upload a dataset before predicting.' })
  const targetColumn = target || dataset.columns[0]
  const values = dataset.rows.map((row) => numeric(row[targetColumn])).filter(Number.isFinite)
  const prediction = values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0
  return response.json({ prediction: Number(prediction.toFixed(2)), target: targetColumn, inputs: inputs ?? {}, explanation: `Prediction uses the mean of ${values.length} observed ${targetColumn} values in this transparent baseline. Replace this endpoint with your trained model for learned predictions.` })
})

app.patch('/api/datasets/:id', (request, response) => {
  const dataset = datasets.get(request.params.id)
  if (!dataset) return response.status(404).json({ error: 'Dataset not found.' })
  const name = String((request.body as { name?: string }).name ?? '').trim()
  if (!name) return response.status(400).json({ error: 'Dataset name cannot be empty.' })
  dataset.fileName = name
  return response.json({ fileName: name, analysis: profile(dataset) })
})

app.get('/api/health', (_request, response) => response.json({ status: 'ok', service: 'lumen-api' }))
app.use((error: { code?: string; message?: string }, _request: express.Request, response: express.Response, next: express.NextFunction) => {
  if (error.code === 'LIMIT_FILE_SIZE') return response.status(413).json({ error: 'A file exceeds the 25 MB upload limit.' })
  if (error.code === 'LIMIT_FILE_COUNT') return response.status(413).json({ error: 'Upload up to 10 files at a time.' })
  if (error) return response.status(400).json({ error: error.message ?? 'Request failed.' })
  return next()
})
const port = Number(process.env.PORT ?? 8787)
app.listen(port, () => console.log(`Lumen API listening on http://localhost:${port}`))
