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
    <main className="min-h-screen bg-[#1f1f1f] text-[#111827]">
      <div className="border-b border-black/40 bg-[#111111] text-white">
        <div className="flex items-center gap-6 px-4 py-3 text-sm">
          <span className="font-medium text-[#d4d4d4]">Archivo</span>
          <span className="font-medium text-[#d4d4d4]">Inicio</span>
          <span className="font-medium text-[#d4d4d4]">Insertar</span>
          <span className="font-medium text-[#d4d4d4]">Datos</span>
          <span className="font-medium text-[#d4d4d4]">Revisar</span>
          <div className="ml-auto rounded bg-[#2a2a2a] px-3 py-1 text-xs text-[#cbd5e1]">
            Hoja local
          </div>
        </div>
        <div className="border-t border-white/5 bg-[#181818] px-4 py-2 text-xs text-[#a3a3a3]">
          Pegar texto, validar con BD y exportar resultados
        </div>
      </div>

      <div className="border-b border-black/20 bg-[#2b2b2b] px-4 py-2 text-white">
        <div className="grid grid-cols-[96px_minmax(0,1fr)] items-center gap-3">
          <div className="rounded border border-white/10 bg-[#1f1f1f] px-2 py-1 text-sm">A1</div>
          <div className="rounded border border-white/10 bg-[#1f1f1f] px-3 py-1.5 text-sm text-[#d1d5db]">
            Editor de registros
          </div>
        </div>
      </div>

      <div className="mx-auto max-w-[1600px] p-4">
        <section className="grid gap-4 xl:grid-cols-[340px_minmax(0,1fr)]">
          <div className="rounded-2xl border border-[#c9cfd8] bg-[#f3f4f6] p-4 shadow-sm">
            <div className="mb-3 flex items-center justify-between gap-3">
              <div>
                <h2 className="text-base font-semibold text-[#111827]">Entrada</h2>
                <p className="text-sm text-[#4b5563]">Pega aqui el texto de Google Lens.</p>
              </div>
              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handlePasteText}
                  className="inline-flex items-center gap-2 rounded-md border border-[#c7d2fe] bg-white px-3 py-2 text-sm font-medium text-[#1f2937] transition hover:bg-[#eef2ff]"
                >
                  <CopyIcon />
                  Pegar
                </button>
                <button
                  type="button"
                  onClick={handleClearText}
                  className="inline-flex items-center gap-2 rounded-md border border-[#fecaca] bg-white px-3 py-2 text-sm font-medium text-[#991b1b] transition hover:bg-[#fef2f2]"
                >
                  <TrashIcon />
                  Limpiar
                </button>
              </div>
            </div>

            <label className="mb-3 flex cursor-pointer items-center justify-between gap-3 rounded-xl border border-[#d1d5db] bg-white px-4 py-3 text-sm text-[#111827] transition hover:bg-[#f9fafb]">
              <div>
                <div className="font-medium text-[#111827]">Base Excel</div>
                <div className="text-xs text-[#6b7280]">Carga `bd.xlsx` con hoja `bd`</div>
              </div>
              <div className="truncate text-right text-xs text-[#6b7280]">{bdFileName || 'Seleccionar archivo'}</div>
              <input type="file" accept=".xlsx,.xls,.xlsm" onChange={handleBdFileChange} className="hidden" />
            </label>

            <textarea
              value={ocrText}
              onChange={(event) => handleTextChange(event.target.value)}
              placeholder="Pega aqui"
              className="min-h-[520px] w-full rounded-xl border border-[#d1d5db] bg-white px-4 py-4 font-mono text-sm text-[#111827] outline-none transition placeholder:text-[#9ca3af] focus:border-[#60a5fa] focus:ring-4 focus:ring-[#bfdbfe]"
            />

            <div className="mt-3 text-sm text-[#4b5563]">
              Lineas detectadas: <span className="font-semibold text-[#111827]">{rawLineCount}</span>
            </div>
          </div>

          <div className="overflow-hidden rounded-2xl border border-[#bfc7d1] bg-white shadow-sm">
            <div className="border-b border-[#d1d5db] bg-[#f3f4f6] px-4 py-3">
              <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <h2 className="text-base font-semibold text-[#111827]">Hoja</h2>
                <p className="text-sm text-[#4b5563]">Filas armadas: {filledRows.length}</p>
              </div>

              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handleAddRow}
                  className="inline-flex items-center gap-2 rounded-md border border-[#d1d5db] bg-white px-3 py-2 text-sm font-medium text-[#111827] transition hover:bg-[#f9fafb]"
                >
                  <PlusIcon />
                  Agregar fila
                </button>
                <button
                  type="button"
                  onClick={handleCopyForExcel}
                  disabled={!filledRows.length}
                  className="inline-flex items-center gap-2 rounded-md border border-[#bbf7d0] bg-[#f0fdf4] px-3 py-2 text-sm font-medium text-[#166534] transition hover:bg-[#dcfce7] disabled:cursor-not-allowed disabled:opacity-40"
                >
                  <CopyIcon />
                  Copiar
                </button>
                <button
                  type="button"
                  onClick={handleExport}
                  disabled={!filledRows.length}
                  className="inline-flex items-center gap-2 rounded-md bg-[#16a34a] px-3 py-2 text-sm font-semibold text-white transition hover:bg-[#15803d] disabled:cursor-not-allowed disabled:opacity-40"
                >
                  <DownloadIcon />
                  Descargar
                </button>
              </div>
            </div>
            </div>

            <div className="overflow-hidden">
              <div className="max-h-[680px] overflow-auto">
                <table className="min-w-full border-collapse text-sm">
                  <thead className="bg-[#111111] text-white">
                    <tr>
                      <th className="border border-[#d1d5db] px-4 py-2 text-left font-medium">DNI</th>
                      <th className="border border-[#d1d5db] px-4 py-2 text-left font-medium">APELLIDOS Y NOMBRES</th>
                      <th className="border border-[#d1d5db] px-4 py-2 text-left font-medium">ESTADO</th>
                      <th className="border border-[#d1d5db] px-4 py-2 text-left font-medium">ACCION</th>
                    </tr>
                  </thead>
                  <tbody className="bg-white text-[#111827]">
                    {rows.map((row) => (
                      <tr
                        key={row.id}
                        className={`align-top transition ${
                          activeRowId === row.id ? 'bg-[#dbeafe]' : 'hover:bg-[#f9fafb]'
                        }`}
                      >
                        <td className="border border-[#e5e7eb] px-2 py-1">
                          <div className="relative">
                            <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-[#9ca3af]">
                              <PencilIcon />
                            </span>
                            <input
                              value={row.number}
                              onChange={(event) => handleRowChange(row.id, 'number', event.target.value)}
                              onFocus={() => setActiveRowId(row.id)}
                              onBlur={() => setActiveRowId((current) => (current === row.id ? '' : current))}
                              placeholder="0018476"
                              className="w-full rounded-none border border-transparent bg-transparent py-2 pl-10 pr-3 text-[#111827] outline-none transition focus:border-[#93c5fd] focus:bg-[#eff6ff]"
                            />
                          </div>
                        </td>
                        <td className="border border-[#e5e7eb] px-2 py-1">
                          <div className="relative">
                            <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-[#9ca3af]">
                              <PencilIcon />
                            </span>
                            <input
                              value={row.name}
                              onChange={(event) => handleRowChange(row.id, 'name', event.target.value)}
                              onFocus={() => setActiveRowId(row.id)}
                              onBlur={() => setActiveRowId((current) => (current === row.id ? '' : current))}
                              placeholder="Nombre completo"
                              className="w-full rounded-none border border-transparent bg-transparent py-2 pl-10 pr-3 text-[#111827] outline-none transition focus:border-[#93c5fd] focus:bg-[#eff6ff]"
                            />
                          </div>
                        </td>
                        <td className="border border-[#e5e7eb] px-2 py-1">
                          <span
                            className={`inline-flex rounded-full px-3 py-1 text-xs font-medium ${
                              row.status === 'matched'
                                ? 'bg-[#dcfce7] text-[#166534]'
                                : row.status === 'missing' || row.status === 'pending'
                                  ? 'bg-[#fef3c7] text-[#92400e]'
                                  : 'bg-[#e5e7eb] text-[#4b5563]'
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
                        <td className="border border-[#e5e7eb] px-2 py-1">
                          <button
                            type="button"
                            onClick={() => handleDeleteRow(row.id)}
                            className="inline-flex h-9 w-9 items-center justify-center rounded-md border border-[#fecaca] bg-[#fff1f2] text-[#b91c1c] transition hover:bg-[#ffe4e6]"
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

            <p className="border-t border-[#d1d5db] bg-[#f9fafb] px-4 py-3 text-sm text-[#4b5563]">
              {copyFeedback || 'Copiar genera texto tabulado listo para pegar directo en Excel.'}
            </p>
          </div>
        </section>
      </div>
    </main>
  )
}

export default App
