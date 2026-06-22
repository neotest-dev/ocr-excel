import { useState } from 'react'
import * as XLSX from 'xlsx'

const sampleRows = [
  { id: crypto.randomUUID(), number: '', name: '', status: 'manual' },
]

function PlusIcon() {
  return (
    <svg viewBox="0 0 20 20" fill="none" aria-hidden="true" className="h-4 w-4">
      <path d="M10 4v12M4 10h12" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    </svg>
  )
}

function CopyIcon() {
  return (
    <svg viewBox="0 0 20 20" fill="none" aria-hidden="true" className="h-4 w-4">
      <rect x="7" y="4" width="9" height="11" rx="2" stroke="currentColor" strokeWidth="1.5" />
      <path d="M5 12H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h6a2 2 0 0 1 2 2v1" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
    </svg>
  )
}

function DownloadIcon() {
  return (
    <svg viewBox="0 0 20 20" fill="none" aria-hidden="true" className="h-4 w-4">
      <path d="M10 3v9m0 0 3-3m-3 3-3-3M4 15.5h12" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  )
}

function TrashIcon() {
  return (
    <svg viewBox="0 0 20 20" fill="none" aria-hidden="true" className="h-4 w-4">
      <path d="M4.5 6h11M8 3.5h4m-6.5 2.5.6 9a2 2 0 0 0 2 1.9h3.8a2 2 0 0 0 2-1.9l.6-9" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M8.5 8.5v5M11.5 8.5v5" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
    </svg>
  )
}

function PencilIcon() {
  return (
    <svg viewBox="0 0 20 20" fill="none" aria-hidden="true" className="h-4 w-4">
      <path d="M3 14.8V17h2.2l8.6-8.6-2.2-2.2L3 14.8ZM12.6 5.4l2.2 2.2 1.1-1.1a1.55 1.55 0 0 0 0-2.2l-.1-.1a1.55 1.55 0 0 0-2.2 0l-1 1.2Z" fill="currentColor" />
    </svg>
  )
}

function normalizeText(text) {
  return text
    .replace(/[|]/g, ' ')
    .replace(/[,:;]/g, ' ')
    .replace(/[.]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

function extractDigits(value) {
  return value.replace(/\D/g, '')
}

function extractNumberAndName(line) {
  const tokens = line.split(/\s+/).filter(Boolean)
  const numberTokens = []
  let splitIndex = 0

  while (splitIndex < tokens.length && /^[\d()._-]+$/.test(tokens[splitIndex])) {
    numberTokens.push(tokens[splitIndex])
    splitIndex += 1
  }

  const number = extractDigits(numberTokens.join(''))
  const name = tokens.slice(splitIndex).join(' ').trim()

  if (number.length >= 3 && name) {
    return { number, name }
  }

  return null
}

function createRow(number = '', name = '') {
  return {
    id: crypto.randomUUID(),
    number,
    name,
    status: 'manual',
  }
}

function normalizeHeader(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase()
}

function normalizeDni(value) {
  return extractDigits(String(value ?? ''))
}

function buildBdMap(workbook) {
  const worksheet = workbook.Sheets.bd

  if (!worksheet) {
    throw new Error('missing_bd_sheet')
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' })
  const dniMap = new Map()

  rows.forEach((row) => {
    const entries = Object.entries(row)
    const dniValue = entries.find(([key]) => normalizeHeader(key) === 'DNI')?.[1] ?? ''
    const nameValue = entries.find(([key]) => normalizeHeader(key) === 'APELLIDOS Y NOMBRES')?.[1] ?? ''
    const dni = normalizeDni(dniValue)
    const name = normalizeText(String(nameValue ?? ''))

    if (dni && name && !dniMap.has(dni)) {
      dniMap.set(dni, name)
    }
  })

  return dniMap
}

function applyBdMatches(rows, bdMap) {
  return rows.map((row) => {
    const dni = normalizeDni(row.number)

    if (!dni) {
      return { ...row, status: row.name ? 'manual' : 'empty' }
    }

    const matchedName = bdMap.get(dni)

    if (matchedName) {
      return {
        ...row,
        number: dni,
        name: matchedName,
        status: 'matched',
      }
    }

    return {
      ...row,
      number: dni,
      status: row.name ? 'missing' : 'pending',
    }
  })
}

function parseLine(line) {
  const cleaned = normalizeText(line)

  if (!cleaned) {
    return null
  }

  const inline = extractNumberAndName(cleaned)

  if (inline) {
    return createRow(inline.number, inline.name)
  }

  const digits = extractDigits(cleaned)
  const name = cleaned.replace(/[\d()._-]/g, ' ').replace(/\s+/g, ' ').trim()

  if (!digits && !name) {
    return null
  }

  return createRow(digits, name)
}

function isNumberOnlyLine(line) {
  return /^[\d()._\-\s]+$/.test(line) && extractDigits(line).length >= 3
}

function looksLikeNameLine(line) {
  return /\p{L}{2,}/u.test(line) && extractDigits(line).length === 0
}

function parseOcrText(text) {
  const lines = text
    .split('\n')
    .map((line) => normalizeText(line))
    .filter(Boolean)

  const inlineRows = []
  const numbers = []
  const names = []

  lines.forEach((line) => {
    const parsed = parseLine(line)

    if (parsed?.number && parsed?.name) {
      inlineRows.push(parsed)
      return
    }

    if (isNumberOnlyLine(line)) {
      numbers.push(extractDigits(line))
      return
    }

    if (looksLikeNameLine(line)) {
      names.push(line)
    }
  })

  if (inlineRows.length) {
    return inlineRows
  }

  const pairCount = Math.max(numbers.length, names.length)
  const rows = Array.from({ length: pairCount }, (_, index) =>
    createRow(numbers[index] ?? '', names[index] ?? ''),
  ).filter((row) => row.number || row.name)

  return rows.length ? rows : sampleRows
}

function buildWorkbook(rows) {
  const sheetRows = rows.filter((row) => row.number || row.name)
  const worksheet = XLSX.utils.aoa_to_sheet([
    ['DNI', 'APELLIDOS Y NOMBRES'],
    ...sheetRows.map((row) => [row.number, row.name]),
  ])

  sheetRows.forEach((row, index) => {
    const numberCell = worksheet[`A${index + 2}`]
    const nameCell = worksheet[`B${index + 2}`]

    if (numberCell) {
      numberCell.t = 's'
      numberCell.z = '@'
      numberCell.v = row.number
    }

    if (nameCell) {
      nameCell.t = 's'
      nameCell.v = row.name
    }
  })

  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Registros')
  return workbook
}

function buildClipboardText(rows) {
  const filledRows = rows.filter((row) => row.number || row.name)
  const lines = filledRows.map((row) => `="${row.number}"\t${row.name}`)
  return ['DNI\tAPELLIDOS Y NOMBRES', ...lines].join('\n')
}

function App() {
  const [ocrText, setOcrText] = useState('')
  const [rows, setRows] = useState(sampleRows)
  const [copyFeedback, setCopyFeedback] = useState('')
  const [activeRowId, setActiveRowId] = useState('')
  const [bdMap, setBdMap] = useState(null)
  const [bdFileName, setBdFileName] = useState('')

  function handleTextChange(value) {
    const parsedRows = parseOcrText(value)
    setOcrText(value)
    setRows(bdMap ? applyBdMatches(parsedRows, bdMap) : parsedRows)
    setCopyFeedback('')
  }

  function handleClearText() {
    setOcrText('')
    setRows(sampleRows)
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
    } catch {
      setBdMap(null)
      setBdFileName('')
      setCopyFeedback('No se pudo leer la hoja bd del Excel base.')
    }
  }

  async function handlePasteText() {
    try {
      const pastedText = await navigator.clipboard.readText()
      const parsedRows = parseOcrText(pastedText)
      setOcrText(pastedText)
      setRows(bdMap ? applyBdMatches(parsedRows, bdMap) : parsedRows)
      setCopyFeedback('')
    } catch {
      setCopyFeedback('No se pudo leer el portapapeles.')
    }
  }

  function handleRowChange(id, field, value) {
    setRows((currentRows) =>
      currentRows.map((row) => {
        if (row.id !== id) {
          return row
        }

        const updatedRow = { ...row, [field]: value }

        if (field === 'number') {
          const normalizedDni = normalizeDni(value)
          const matchedName = bdMap?.get(normalizedDni)

          if (matchedName) {
            return {
              ...updatedRow,
              number: normalizedDni,
              name: matchedName,
              status: 'matched',
            }
          }

          return {
            ...updatedRow,
            number: normalizedDni || value,
            status: updatedRow.name ? 'missing' : 'pending',
          }
        }

        return {
          ...updatedRow,
          status: updatedRow.number ? (bdMap?.has(normalizeDni(updatedRow.number)) ? 'matched' : 'manual') : 'manual',
        }
      }),
    )
  }

  function handleAddRow() {
    setRows((currentRows) => [
      ...currentRows,
      { id: crypto.randomUUID(), number: '', name: '' },
    ])
  }

  function handleDeleteRow(id) {
    setRows((currentRows) => {
      const nextRows = currentRows.filter((row) => row.id !== id)
      return nextRows.length ? nextRows : sampleRows
    })
  }

  function handleExport() {
    const workbook = buildWorkbook(rows)
    XLSX.writeFile(workbook, 'registros.xlsx')
  }

  async function handleCopyForExcel() {
    const clipboardText = buildClipboardText(rows)

    try {
      await navigator.clipboard.writeText(clipboardText)
      setCopyFeedback('Copiado para pegar directo en Excel.')
    } catch {
      setCopyFeedback('No se pudo copiar. Usa descargar Excel.')
    }
  }

  const filledRows = rows.filter((row) => row.number || row.name)
  const rawLineCount = ocrText.split('\n').map((line) => line.trim()).filter(Boolean).length

  return (
    <main className="min-h-screen px-4 py-6 text-slate-100 sm:px-6 lg:px-10">
      <div className="mx-auto max-w-7xl">
        <div className="mb-6 flex flex-col gap-2">
          <p className="text-xs font-semibold uppercase tracking-[0.28em] text-sky-300/80">
            Google Lens a Excel
          </p>
          <h1 className="text-3xl font-semibold tracking-tight text-white sm:text-4xl">
            Pega, ordena y exporta.
          </h1>
          <p className="max-w-2xl text-sm text-slate-300 sm:text-base">
            Pega el texto de Google Lens a la izquierda. La tabla se arma al lado para
            corregir, copiar o descargar.
          </p>
        </div>

        <section className="grid gap-6 lg:grid-cols-[minmax(0,0.95fr)_minmax(0,1.15fr)]">
          <div className="rounded-3xl border border-white/10 bg-slate-950/70 p-4 shadow-2xl shadow-black/30 backdrop-blur sm:p-5">
            <div className="mb-3 flex items-center justify-between gap-3">
              <div>
                <h2 className="text-lg font-semibold text-white">Texto de Google Lens</h2>
                <p className="text-sm text-slate-400">Pega numeros y nombres como salgan.</p>
              </div>
              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handlePasteText}
                  className="inline-flex items-center gap-2 rounded-full border border-emerald-400/30 bg-emerald-400/10 px-4 py-2 text-sm font-medium text-emerald-200 transition hover:bg-emerald-400/20"
                >
                  <CopyIcon />
                  Pegar
                </button>
                <button
                  type="button"
                  onClick={handleClearText}
                  className="inline-flex items-center gap-2 rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm font-medium text-slate-200 transition hover:bg-white/10"
                >
                  <TrashIcon />
                  Limpiar
                </button>
              </div>
            </div>

            <label className="mb-3 flex cursor-pointer items-center justify-between gap-3 rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-sm text-slate-200 transition hover:bg-white/10">
              <div>
                <div className="font-medium text-white">Base Excel</div>
                <div className="text-xs text-slate-400">Carga `bd.xlsx` con hoja `bd`</div>
              </div>
              <div className="truncate text-right text-xs text-slate-400">{bdFileName || 'Seleccionar archivo'}</div>
              <input type="file" accept=".xlsx,.xls,.xlsm" onChange={handleBdFileChange} className="hidden" />
            </label>

            <textarea
              value={ocrText}
              onChange={(event) => handleTextChange(event.target.value)}
              placeholder="Pega aqui"
              className="min-h-[420px] w-full rounded-2xl border border-white/10 bg-slate-900/80 px-4 py-4 text-sm text-slate-100 outline-none transition placeholder:text-slate-500 focus:border-sky-400/50 focus:ring-4 focus:ring-sky-400/10"
            />

            <div className="mt-3 text-sm text-slate-400">
              Lineas detectadas: <span className="font-semibold text-slate-200">{rawLineCount}</span>
            </div>
          </div>

          <div className="rounded-3xl border border-white/10 bg-slate-950/70 p-4 shadow-2xl shadow-black/30 backdrop-blur sm:p-5">
            <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <h2 className="text-lg font-semibold text-white">Tabla</h2>
                <p className="text-sm text-slate-400">Filas armadas: {filledRows.length}</p>
              </div>

              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handleAddRow}
                  className="inline-flex items-center gap-2 rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm font-medium text-slate-200 transition hover:bg-white/10"
                >
                  <PlusIcon />
                  Agregar fila
                </button>
                <button
                  type="button"
                  onClick={handleCopyForExcel}
                  disabled={!filledRows.length}
                  className="inline-flex items-center gap-2 rounded-full border border-emerald-400/30 bg-emerald-400/10 px-4 py-2 text-sm font-medium text-emerald-200 transition hover:bg-emerald-400/20 disabled:cursor-not-allowed disabled:opacity-40"
                >
                  <CopyIcon />
                  Copiar
                </button>
                <button
                  type="button"
                  onClick={handleExport}
                  disabled={!filledRows.length}
                  className="inline-flex items-center gap-2 rounded-full bg-white px-4 py-2 text-sm font-semibold text-slate-950 transition hover:bg-slate-200 disabled:cursor-not-allowed disabled:opacity-40"
                >
                  <DownloadIcon />
                  Descargar
                </button>
              </div>
            </div>

            <div className="overflow-hidden rounded-2xl border border-white/10">
              <div className="max-h-[520px] overflow-auto">
                <table className="min-w-full divide-y divide-white/10 text-sm">
                  <thead className="bg-white/5 text-slate-300">
                    <tr>
                      <th className="px-4 py-3 text-left font-medium">DNI</th>
                      <th className="px-4 py-3 text-left font-medium">APELLIDOS Y NOMBRES</th>
                      <th className="px-4 py-3 text-left font-medium">ESTADO</th>
                      <th className="px-4 py-3 text-left font-medium">Accion</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-white/10 bg-slate-950/30">
                    {rows.map((row) => (
                      <tr
                        key={row.id}
                        className={`align-top transition ${
                          activeRowId === row.id ? 'bg-sky-400/8 ring-1 ring-inset ring-sky-400/20' : ''
                        }`}
                      >
                        <td className="px-3 py-3">
                          <div className="relative">
                            <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-500">
                              <PencilIcon />
                            </span>
                            <input
                              value={row.number}
                              onChange={(event) => handleRowChange(row.id, 'number', event.target.value)}
                              onFocus={() => setActiveRowId(row.id)}
                              onBlur={() => setActiveRowId((current) => (current === row.id ? '' : current))}
                              placeholder="0018476"
                              className="w-full rounded-xl border border-white/10 bg-slate-900/90 py-2 pl-10 pr-3 text-slate-100 outline-none transition focus:border-sky-400/50 focus:ring-4 focus:ring-sky-400/10"
                            />
                          </div>
                        </td>
                        <td className="px-3 py-3">
                          <div className="relative">
                            <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-500">
                              <PencilIcon />
                            </span>
                            <input
                              value={row.name}
                              onChange={(event) => handleRowChange(row.id, 'name', event.target.value)}
                              onFocus={() => setActiveRowId(row.id)}
                              onBlur={() => setActiveRowId((current) => (current === row.id ? '' : current))}
                              placeholder="Nombre completo"
                              className="w-full rounded-xl border border-white/10 bg-slate-900/90 py-2 pl-10 pr-3 text-slate-100 outline-none transition focus:border-sky-400/50 focus:ring-4 focus:ring-sky-400/10"
                            />
                          </div>
                        </td>
                        <td className="px-3 py-3">
                          <span
                            className={`inline-flex rounded-full px-3 py-1 text-xs font-medium ${
                              row.status === 'matched'
                                ? 'bg-emerald-400/15 text-emerald-200'
                                : row.status === 'missing' || row.status === 'pending'
                                  ? 'bg-amber-400/15 text-amber-200'
                                  : 'bg-slate-400/15 text-slate-300'
                            }`}
                          >
                            {row.status === 'matched'
                              ? 'Encontrado'
                              : row.status === 'missing'
                                ? 'No encontrado'
                                : row.status === 'pending'
                                  ? 'Sin nombre'
                                  : 'Manual'}
                          </span>
                        </td>
                        <td className="px-3 py-3">
                          <button
                            type="button"
                            onClick={() => handleDeleteRow(row.id)}
                            className="inline-flex h-9 w-9 items-center justify-center rounded-full border border-rose-400/20 bg-rose-400/10 text-rose-200 transition hover:bg-rose-400/20"
                            aria-label="Quitar fila"
                            title="Quitar fila"
                          >
                            <TrashIcon />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <p className="mt-3 text-sm text-slate-400">
              {copyFeedback || 'Copiar genera texto tabulado listo para pegar directo en Excel.'}
            </p>
          </div>
        </section>
      </div>
    </main>
  )
}

export default App
