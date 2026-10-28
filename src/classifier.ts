import './classifier.css';
type DatasetResult = { id: string; analysis: { fileName: string; rows: number; columns: number; schema: { name: string; type: string }[] } };
type ClassResult = { model: string; target: string; features: string[]; split: string; metrics: { accuracy: number; correct: number; testSamples: number; trainSamples: number; macroPrecision: number; macroRecall: number; macroF1: number }; classes: string[]; confusionMatrix: number[][]; perClass: { class: string; precision: number; recall: number; f1: number; support: number }[]; note: string };
const fileInput = document.querySelector<HTMLInputElement>('#dataset-file')!;
const uploadButton = document.querySelector<HTMLButtonElement>('#upload-button')!;
const evaluateButton = document.querySelector<HTMLButtonElement>('#evaluate-button')!;
const targetInput = document.querySelector<HTMLInputElement>('#target-column')!;
const status = document.querySelector<HTMLParagraphElement>('#upload-status')!;
const columnList = document.querySelector<HTMLDivElement>('#column-list')!;
const results = document.querySelector<HTMLElement>('#results')!;
let datasetId = '';
function setStatus(message: string, kind = '') { status.textContent = message; status.className = 'status ' + kind; }
async function jsonRequest<T>(url: string, init: RequestInit): Promise<T> {
  const response = await fetch(url, init); const body = await response.json();
  if (!response.ok) throw new Error(body.error ?? 'Request failed.');
  return body as T;
}
uploadButton.addEventListener('click', async () => {
  const file = fileInput.files?.[0]; if (!file) { setStatus('Choose a CSV or Excel file first.', 'error'); return; }
  uploadButton.disabled = true; setStatus('Uploading and profiling dataset…');
  try {
    const form = new FormData(); form.append('files', file);
    const response = await jsonRequest<{ results: (DatasetResult | { error: string; fileName: string })[] }>('/api/analyze-batch', { method: 'POST', body: form });
    const accepted = response.results.find((item): item is DatasetResult => 'id' in item && 'analysis' in item);
    if (!accepted) throw new Error(response.results.map((item) => 'error' in item ? item.error : '').filter(Boolean).join('; ') || 'No dataset was accepted.');
    datasetId = accepted.id; evaluateButton.disabled = false;
    const columns = accepted.analysis.schema;
    columnList.replaceChildren(...columns.map((column) => { const chip = document.createElement('span'); chip.className = 'column-chip'; chip.textContent = column.name + ' · ' + column.type; return chip; }));
    setStatus(`Loaded ${accepted.analysis.fileName}: ${accepted.analysis.rows} rows, ${accepted.analysis.columns} columns. Enter the target column below.`, 'success');
  } catch (error) { setStatus(error instanceof Error ? error.message : 'Upload failed.', 'error'); }
  finally { uploadButton.disabled = false; }
});
evaluateButton.addEventListener('click', async () => {
  if (!datasetId) { setStatus('Upload a labelled dataset first.', 'error'); return; }
  const target = targetInput.value.trim(); if (!target) { setStatus('Enter the target column name.', 'error'); return; }
  evaluateButton.disabled = true; evaluateButton.textContent = 'Training and evaluating…'; results.hidden = true;
  try {
    const result = await jsonRequest<ClassResult>('/api/classify', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ datasetId, target }) });
    results.hidden = false;
    const score = (value: number) => (value * 100).toFixed(2) + '%';
    const cards = [
      ['Test accuracy', score(result.metrics.accuracy), `${result.metrics.correct}/${result.metrics.testSamples} correct`],
      ['Macro precision', score(result.metrics.macroPrecision), 'Across classes'],
      ['Macro recall', score(result.metrics.macroRecall), 'Across classes'],
      ['Macro F1', score(result.metrics.macroF1), 'Across classes']
    ];
    results.innerHTML = `<h2>Evaluation results</h2><p><strong>${result.model}</strong> · Target: <strong>${result.target}</strong></p><p>Features: ${result.features.join(', ')}. Split: ${result.split}. Training rows: ${result.metrics.trainSamples}; test rows: ${result.metrics.testSamples}.</p><div class="score-cards">${cards.map(([label, value, detail]) => `<div class="score"><span>${label}</span><strong>${value}</strong><small>${detail}</small></div>`).join('')}</div><h3>Confusion matrix</h3><p>Rows = actual class; columns = predicted class.</p><div class="matrix-wrap"><table><thead><tr><th>Actual ↓ / Predicted →</th>${result.classes.map((label) => `<th>${label}</th>`).join('')}</tr></thead><tbody>${result.confusionMatrix.map((row, index) => `<tr><th>${result.classes[index]}</th>${row.map((value) => `<td>${value}</td>`).join('')}</tr>`).join('')}</tbody></table></div><h3>Per-class report</h3><div class="matrix-wrap"><table><thead><tr><th>Class</th><th>Precision</th><th>Recall</th><th>F1</th><th>Support</th></tr></thead><tbody>${result.perClass.map((item) => `<tr><th>${item.class}</th><td>${score(item.precision)}</td><td>${score(item.recall)}</td><td>${score(item.f1)}</td><td>${item.support}</td></tr>`).join('')}</tbody></table></div><p>${result.note}</p>`;
  } catch (error) { results.hidden = false; results.textContent = error instanceof Error ? error.message : 'Classification failed.'; }
  finally { evaluateButton.disabled = false; evaluateButton.textContent = 'Train classifier & calculate scores'; }
});