import { useState } from 'react'
import * as XLSX from 'xlsx'
import { LensWorkspace } from './components/LensWorkspace'
import { SpreadsheetWorkspace } from './components/SpreadsheetWorkspace'
import {
  applyBdMatches,
  buildBdMap,
  buildClipboardText,
  buildPlanillaClipboardHtml,
  buildPlanillaClipboardText,
  buildPlanillaRowsFromWorkbooks,
  buildPlanillaWorkbook,
  buildWorkbook,
  createEmptyRow,
  createSampleRows,
  findSheetOrThrow,
  parseOcrText,
  updateRowWithMatch,
} from './lib/records'

const views = [
  {
    id: 'lens',
    label: 'Captura',
    title: 'Flujo rapido desde Google Lens',
    description: 'Pegar, corregir y exportar con una paleta verde de oficina.',
  },
  {
    id: 'sheet',
    label: 'Planilla',
    title: 'Vista tipo hoja de calculo',
    description: 'Una superficie de trabajo mas densa para revisar filas como tabla.',
  },
]

function App() {
  const [currentView, setCurrentView] = useState('lens')
  const [ocrText, setOcrText] = useState('')
  const [rows, setRows] = useState(createSampleRows)
  const [copyFeedback, setCopyFeedback] = useState('')
  const [activeRowId, setActiveRowId] = useState('')
  const [bdMap, setBdMap] = useState(null)
  const [bdFileName, setBdFileName] = useState('')
  const [planillaBdFile, setPlanillaBdFile] = useState(null)
  const [planillaBdFileName, setPlanillaBdFileName] = useState('')
  const [planillaReportFile, setPlanillaReportFile] = useState(null)
  const [planillaReportFileName, setPlanillaReportFileName] = useState('')
  const [planillaRows, setPlanillaRows] = useState([])
  const [planillaStatus, setPlanillaStatus] = useState({ message: 'Esperando archivos.', type: 'idle' })

  function processText(text) {
    const parsedRows = parseOcrText(text)
    return bdMap ? applyBdMatches(parsedRows, bdMap) : parsedRows
  }

  function handleTextChange(value) {
    setOcrText(value)
    setRows(processText(value))
    setCopyFeedback('')
  }

  function handleClearText() {
    setOcrText('')
    setRows(createSampleRows())
    setCopyFeedback('')
    setActiveRowId('')
  }

  async function handleBdFileChange(event) {
    const file = event.target.files?.[0]

    if (!file) {
      return
    }

    try {
      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array', raw: false })
      const nextBdMap = buildBdMap(workbook)

      setBdMap(nextBdMap)
      setBdFileName(file.name)
      setRows((currentRows) => applyBdMatches(currentRows, nextBdMap))
      setCopyFeedback(`Base cargada: ${file.name}`)
    } catch (error) {
      setBdMap(null)
      setBdFileName('')
      setCopyFeedback(error instanceof Error ? error.message : 'No se pudo leer la hoja bd del Excel base.')
    }
  }

  function handlePlanillaBdFileChange(event) {
    const file = event.target.files?.[0]

    if (!file) {
      return
    }

    setPlanillaBdFile(file)
    setPlanillaBdFileName(file.name)
    setPlanillaStatus({ message: `BD cargada: ${file.name}`, type: 'idle' })
  }

  function handlePlanillaReportFileChange(event) {
    const file = event.target.files?.[0]

    if (!file) {
      return
    }

    setPlanillaReportFile(file)
    setPlanillaReportFileName(file.name)
    setPlanillaStatus({ message: `REPORTE cargado: ${file.name}`, type: 'idle' })
  }

  async function handlePlanillaProcess() {
    if (!planillaBdFile || !planillaReportFile) {
      setPlanillaStatus({ message: 'Debes seleccionar bd.xlsx y REPORTE.xlsx.', type: 'error' })
      return
    }

    try {
      setPlanillaStatus({ message: 'Procesando archivos...', type: 'idle' })
      const [bdBuffer, reportBuffer] = await Promise.all([
        planillaBdFile.arrayBuffer(),
        planillaReportFile.arrayBuffer(),
      ])
      const bdWorkbook = XLSX.read(bdBuffer, { type: 'array', raw: false })
      const reportWorkbook = XLSX.read(reportBuffer, { type: 'array', raw: false })

      findSheetOrThrow(bdWorkbook, 'bd')
      findSheetOrThrow(reportWorkbook, 'base')

      const nextRows = buildPlanillaRowsFromWorkbooks(bdWorkbook, reportWorkbook)
      setPlanillaRows(nextRows)
      setPlanillaStatus({
        message: nextRows.length
          ? 'Proceso completado. Usa Copiar para Excel y pega desde A2.'
          : 'No se encontraron filas utiles en la hoja bd.',
        type: nextRows.length ? 'success' : 'error',
      })
    } catch (error) {
      setPlanillaRows([])
      setPlanillaStatus({
        message: error instanceof Error ? error.message : 'Ocurrio un error al procesar.',
        type: 'error',
      })
    }
  }

  async function handlePasteText() {
    try {
      const pastedText = await navigator.clipboard.readText()
      setOcrText(pastedText)
      setRows(processText(pastedText))
      setCopyFeedback('')
    } catch {
      setCopyFeedback('No se pudo leer el portapapeles.')
    }
  }

  function handleRowChange(id, field, value) {
    setRows((currentRows) =>
      currentRows.map((row) => (row.id === id ? updateRowWithMatch(row, field, value, bdMap) : row)),
    )
  }

  function handleAddRow() {
    setRows((currentRows) => [...currentRows, createEmptyRow()])
  }

  function handleDeleteRow(id) {
    setRows((currentRows) => {
      const nextRows = currentRows.filter((row) => row.id !== id)
      return nextRows.length ? nextRows : createSampleRows()
    })
  }

  function handleExport() {
    const workbook = buildWorkbook(rows)
    XLSX.writeFile(workbook, 'registros.xlsx')
  }

  async function handleCopyForExcel() {
    try {
      await navigator.clipboard.writeText(buildClipboardText(rows))
      setCopyFeedback('Copiado para pegar directo en Excel.')
    } catch {
      setCopyFeedback('No se pudo copiar. Usa descargar Excel.')
    }
  }

  async function handlePlanillaCopy() {
    if (!planillaRows.length) {
      setPlanillaStatus({ message: 'No hay datos para copiar.', type: 'error' })
      return
    }

    try {
      const text = buildPlanillaClipboardText(planillaRows)
      const html = buildPlanillaClipboardHtml(planillaRows)

      if (window.ClipboardItem && navigator.clipboard.write) {
        await navigator.clipboard.write([
          new ClipboardItem({
            'text/plain': new Blob([text], { type: 'text/plain' }),
            'text/html': new Blob([html], { type: 'text/html' }),
          }),
        ])
      } else {
        await navigator.clipboard.writeText(text)
      }

      setPlanillaStatus({ message: 'Copiado. Pega en Excel desde la celda A2.', type: 'success' })
    } catch {
      setPlanillaStatus({ message: 'No se pudo copiar. Revisa permisos del navegador.', type: 'error' })
    }
  }

  function handlePlanillaExport() {
    if (!planillaRows.length) {
      setPlanillaStatus({ message: 'No hay datos para descargar.', type: 'error' })
      return
    }

    XLSX.writeFile(buildPlanillaWorkbook(planillaRows), 'planilla-procesada.xlsx')
    setPlanillaStatus({ message: 'Excel generado correctamente.', type: 'success' })
  }

  const filledRows = rows.filter((row) => row.number || row.name)
  const rawLineCount = ocrText.split('\n').map((line) => line.trim()).filter(Boolean).length
  const currentViewMeta = views.find((view) => view.id === currentView)

  const sharedProps = {
    ocrText,
    rows,
    copyFeedback,
    activeRowId,
    bdFileName,
    rawLineCount,
    filledRows,
    onTextChange: handleTextChange,
    onPasteText: handlePasteText,
    onClearText: handleClearText,
    onBdFileChange: handleBdFileChange,
    onRowChange: handleRowChange,
    onRowFocus: setActiveRowId,
    onRowBlur: (id) => setActiveRowId((current) => (current === id ? '' : current)),
    onAddRow: handleAddRow,
    onDeleteRow: handleDeleteRow,
    onCopy: handleCopyForExcel,
    onExport: handleExport,
  }

  const planillaProps = {
    planillaBdFileName,
    planillaReportFileName,
    planillaRows,
    planillaStatus,
    onPlanillaBdFileChange: handlePlanillaBdFileChange,
    onPlanillaReportFileChange: handlePlanillaReportFileChange,
    onPlanillaProcess: handlePlanillaProcess,
    onPlanillaCopy: handlePlanillaCopy,
    onPlanillaExport: handlePlanillaExport,
  }

  return (
    <main className="min-h-screen bg-[linear-gradient(180deg,#edf5ef_0%,#eef4ef_42%,#e3ece5_100%)] text-slate-950">
      <header className="border-b border-emerald-900/10 bg-[#0d1b13] text-white shadow-[0_16px_40px_-24px_rgba(6,78,59,0.55)]">
        <div className="mx-auto flex max-w-[1600px] flex-col gap-5 px-4 py-4 sm:px-6 lg:px-8">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-[0.28em] text-emerald-300/75">Workspace</p>
              <h1 className="mt-2 text-3xl font-semibold tracking-tight text-white sm:text-4xl">Cruce de registros con paleta office.</h1>
            </div>
            <div className="max-w-xl text-sm text-emerald-50/75">
              {currentViewMeta?.description}
            </div>
          </div>

          <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
            <nav className="flex flex-wrap gap-2 rounded-2xl border border-white/10 bg-white/5 p-1.5 backdrop-blur">
              {views.map((view) => {
                const active = view.id === currentView
                return (
                  <button
                    key={view.id}
                    type="button"
                    onClick={() => setCurrentView(view.id)}
                    className={`rounded-xl px-4 py-2.5 text-sm font-semibold transition ${
                      active ? 'bg-emerald-500 text-emerald-950 shadow-sm' : 'text-emerald-50/80 hover:bg-white/10'
                    }`}
                  >
                    {view.label}
                  </button>
                )
              })}
            </nav>

            <div className="grid gap-2 sm:grid-cols-4">
              <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3">
                <p className="text-xs uppercase tracking-[0.18em] text-emerald-200/80">Vista</p>
                <strong className="mt-1 block text-sm text-white">{currentViewMeta?.title}</strong>
              </div>
              <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3">
                <p className="text-xs uppercase tracking-[0.18em] text-emerald-200/80">Base</p>
                <strong className="mt-1 block truncate text-sm text-white">{bdFileName || 'Sin archivo'}</strong>
              </div>
              <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3">
                <p className="text-xs uppercase tracking-[0.18em] text-emerald-200/80">Planilla BD</p>
                <strong className="mt-1 block truncate text-sm text-white">{planillaBdFileName || 'Sin archivo'}</strong>
              </div>
              <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3">
                <p className="text-xs uppercase tracking-[0.18em] text-emerald-200/80">Planilla REPORTE</p>
                <strong className="mt-1 block truncate text-sm text-white">{planillaReportFileName || 'Sin archivo'}</strong>
              </div>
              <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3">
                <p className="text-xs uppercase tracking-[0.18em] text-emerald-200/80">Filas Lens</p>
                <strong className="mt-1 block text-sm text-white">{filledRows.length}</strong>
              </div>
            </div>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-[1600px] px-4 py-6 sm:px-6 lg:px-8">
        {currentView === 'sheet' ? <SpreadsheetWorkspace {...planillaProps} /> : <LensWorkspace {...sharedProps} />}
      </div>
    </main>
  )
}

export default App
