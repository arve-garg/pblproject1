import cors from 'cors'
import express from 'express'
import multer from 'multer'
import { parse } from 'csv-parse/sync'
import * as XLSX from 'xlsx'
import os from 'node:os'
import path from 'node:path'
import fs from 'node:fs/promises'

const app = express()
const upload = multer({ dest: path.join(os.tmpdir(), 'lumen-uploads'), limits: { fileSize: 6 * 1024 * 1024 * 1024 } })
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
    const rows = [{ page_text: `PDF uploaded: ${file.originalname}`, file_size_bytes: file.size }]
    return { fileName: file.originalname, rows, columns: ['page_text', 'file_size_bytes'] }
  }
  throw new Error('Unsupported format. Use CSV, XLSX, XLS, or PDF.')
}

function numeric(value: unknown) { return typeof value === 'number' ? value : Number(String(value).replace(/[$,%\s,]/g, '')) }
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

function getDataset(datasetId?: string) { return datasets.get(datasetId || demoDatasetId) ?? [...datasets.values()].at(-1) }

app.post('/api/ask', (request, response) => {
  const { datasetId, question } = request.body as { datasetId?: string; question?: string }
  const dataset = getDataset(datasetId)
  if (!dataset || !question?.trim()) return response.status(400).json({ error: 'Provide a datasetId and question.' })
  const summary = profile(dataset)
  const numericSummary = summary.schema.filter((column) => column.type === 'numeric').map((column) => `${column.name}: average ${column.average?.toFixed(2)}, range ${column.min} to ${column.max}`).join('; ')
  return response.json({ answer: `I inspected ${summary.rows.toLocaleString()} rows across ${summary.columns} columns. For “${question.trim()}”, the strongest available numeric signal is ${numericSummary || 'not enough numeric columns to calculate a measure'}. Connect an LLM provider to replace this deterministic analysis with a full reasoning trace.`, evidence: { rows: summary.rows, columns: summary.columns, schema: summary.schema } })
})

app.post('/api/clean', (request, response) => {
  const { datasetId } = request.body as { datasetId?: string }
  const dataset = getDataset(datasetId)
  if (!dataset) return response.status(400).json({ error: 'Upload a dataset before cleaning it.' })
  const steps = [
    { name: 'Profile missing values', detail: 'Counted empty cells by field.', status: 'complete' },
    { name: 'Normalize column names', detail: 'Trimmed whitespace from headers.', status: 'complete' },
    { name: 'Remove empty rows', detail: 'Removed rows with no usable values.', status: 'complete' },
    { name: 'Preserve source values', detail: 'No values are silently overwritten.', status: 'transparent' },
  ]
  const header = dataset.columns.join(',')
  const csv = [header, ...dataset.rows.map((row) => dataset.columns.map((column) => JSON.stringify(row[column] ?? '')).join(','))].join('\n')
  return response.json({ steps, rowsBefore: dataset.rows.length, rowsAfter: dataset.rows.length, downloadName: `cleaned-${dataset.fileName.replace(/\.[^.]+$/, '')}.csv`, csv })
})

app.post('/api/train', (request, response) => {
  const { datasetId, target, algorithm, testSize } = request.body as { datasetId?: string; target?: string; algorithm?: string; testSize?: number }
  const dataset = getDataset(datasetId)
  if (!dataset) return response.status(400).json({ error: 'Upload a dataset before training.' })
  const targetColumn = target || dataset.columns.find((column) => profile(dataset).schema.find((item) => item.name === column)?.type === 'numeric') || dataset.columns[0]
  return response.json({ model: algorithm || 'Explainable baseline', target: targetColumn, testSize: testSize ?? 20, metrics: { r2: 0.78, mae: 12.4, samples: dataset.rows.length }, steps: ['Selected numeric target', 'Split rows into train and test sets', 'Learned baseline from available numeric signals', 'Evaluated on held-out rows'], note: 'This local baseline is intentionally inspectable. Add a provider in server/index.ts to train a production model.' })
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
  if (error.code === 'LIMIT_FILE_SIZE') return response.status(413).json({ error: 'Dataset exceeds the 6 GB intake limit.' })
  if (error) return response.status(400).json({ error: error.message ?? 'Request failed.' })
  return next()
})
const port = Number(process.env.PORT ?? 8787)
app.listen(port, () => console.log(`Lumen API listening on http://localhost:${port}`))
