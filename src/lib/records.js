import * as XLSX from 'xlsx'

export function createEmptyRow(number = '', name = '') {
  return {
    id: crypto.randomUUID(),
    number,
    name,
    status: 'manual',
  }
}

export function createSampleRows() {
  return [createEmptyRow()]
}

export function normalizeText(text) {
  return String(text ?? '')
    .replace(/[|]/g, ' ')
    .replace(/[,:;]/g, ' ')
    .replace(/[.]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

export function extractDigits(value) {
  return String(value ?? '').replace(/\D/g, '')
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

function parseLine(line) {
  const cleaned = normalizeText(line)

  if (!cleaned) {
    return null
  }

  const inline = extractNumberAndName(cleaned)

  if (inline) {
    return createEmptyRow(inline.number, inline.name)
  }

  const digits = extractDigits(cleaned)
  const name = cleaned.replace(/[\d()._-]/g, ' ').replace(/\s+/g, ' ').trim()

  if (!digits && !name) {
    return null
  }

  return createEmptyRow(digits, name)
}

function isNumberOnlyLine(line) {
  return /^[\d()._\-\s]+$/.test(line) && extractDigits(line).length >= 3
}

function looksLikeNameLine(line) {
  return /\p{L}{2,}/u.test(line) && extractDigits(line).length === 0
}

export function parseOcrText(text) {
  const lines = String(text ?? '')
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
    createEmptyRow(numbers[index] ?? '', names[index] ?? ''),
  ).filter((row) => row.number || row.name)

  return rows.length ? rows : createSampleRows()
}

function normalizeHeader(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase()
}

function findHeaderValue(row, aliases) {
  const entries = Object.entries(row)
  const match = entries.find(([key]) => aliases.includes(normalizeHeader(key)))
  return match?.[1] ?? ''
}

function findSheetByNames(workbook, names) {
  const sheetName = workbook.SheetNames.find((name) => names.includes(name.trim().toLowerCase()))
  return sheetName ? workbook.Sheets[sheetName] : null
}

export function normalizeDni(value) {
  return extractDigits(String(value ?? ''))
}

export function buildBdMap(workbook) {
  const worksheet = findSheetByNames(workbook, ['bd'])

  if (!worksheet) {
    throw new Error('No se pudo encontrar la hoja bd.')
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' })
  const dniMap = new Map()

  rows.forEach((row) => {
    const dniValue = findHeaderValue(row, ['DNI'])
    const nameValue = findHeaderValue(row, ['APELLIDOS Y NOMBRES'])
    const dni = normalizeDni(dniValue)
    const name = normalizeText(nameValue)

    if (dni && name && !dniMap.has(dni)) {
      dniMap.set(dni, name)
    }
  })

  return dniMap
}

export function buildRowsFromReportWorkbook(workbook) {
  const worksheet =
    findSheetByNames(workbook, ['reporte', 'base', 'hoja1']) ?? workbook.Sheets[workbook.SheetNames[0]]

  if (!worksheet) {
    throw new Error('No se pudo encontrar una hoja util en REPORTE.')
  }

  const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false })
  const rows = jsonRows
    .map((row) => {
      const dni = normalizeDni(findHeaderValue(row, ['DNI', 'DOCUMENTO', 'NUMERO DE DOCUMENTO']))
      const name = normalizeText(
        findHeaderValue(row, ['APELLIDOS Y NOMBRES', 'NOMBRES Y APELLIDOS', 'NOMBRE COMPLETO', 'NOMBRES']),
      )

      if (!dni && !name) {
        return null
      }

      return createEmptyRow(dni, name)
    })
    .filter(Boolean)

  return rows.length ? rows : createSampleRows()
}

export function applyBdMatches(rows, bdMap) {
  return rows.map((row) => {
    const dni = normalizeDni(row.number)

    if (!dni) {
      return { ...row, status: row.name ? 'manual' : 'empty' }
    }

    const matchedName = bdMap?.get(dni)

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

export function updateRowWithMatch(row, field, value, bdMap) {
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
}

export function buildWorkbook(rows) {
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

export function buildClipboardText(rows) {
  const filledRows = rows.filter((row) => row.number || row.name)
  const lines = filledRows.map((row) => `="${row.number}"\t${row.name}`)
  return ['DNI\tAPELLIDOS Y NOMBRES', ...lines].join('\n')
}

export function getStatusLabel(status) {
  if (status === 'matched') return 'Encontrado'
  if (status === 'missing') return 'No encontrado'
  if (status === 'pending') return 'Sin nombre'
  return 'Manual'
}

export function normalizeIdentifier(value) {
  const normalized = String(value ?? '').trim()

  if (!normalized) {
    return ''
  }

  if (/^\d+\.0+$/.test(normalized)) {
    return normalized.replace(/\.0+$/, '')
  }

  return normalized.replace(/\s+/g, '')
}

export function formatToday() {
  const today = new Date()
  return `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`
}

export function findSheetOrThrow(workbook, desiredName) {
  const sheet = findSheetByNames(workbook, [desiredName.toLowerCase()])

  if (!sheet) {
    throw new Error(`No se encontro la hoja ${desiredName}.`)
  }

  return sheet
}

export function buildPlanillaRowsFromWorkbooks(bdWorkbook, reporteWorkbook) {
  const bdSheet = findSheetOrThrow(bdWorkbook, 'bd')
  const reporteSheet = findSheetOrThrow(reporteWorkbook, 'base')
  const bdRows = XLSX.utils.sheet_to_json(bdSheet, { defval: '', raw: false })
  const reporteRows = XLSX.utils.sheet_to_json(reporteSheet, { defval: '', raw: false })

  const reporteIndex = new Map()

  reporteRows.forEach((row) => {
    const dni = normalizeIdentifier(findHeaderValue(row, ['DNI']))

    if (!dni || reporteIndex.has(dni)) {
      return
    }

    reporteIndex.set(dni, {
      codigo1: normalizeIdentifier(findHeaderValue(row, ['CODIGO 1'])),
      codigo2: normalizeIdentifier(findHeaderValue(row, ['CODIGO 2'])),
      nombre: normalizeText(findHeaderValue(row, ['NOMBRES Y APELLIDOS', 'APELLIDOS Y NOMBRES'])),
    })
  })

  const today = formatToday()

  return bdRows
    .map((row) => {
      const dni = normalizeIdentifier(findHeaderValue(row, ['DNI']))
      const nombreBd = normalizeText(findHeaderValue(row, ['APELLIDOS Y NOMBRES', 'NOMBRES Y APELLIDOS']))

      if (!dni && !nombreBd) {
        return null
      }

      const match = reporteIndex.get(dni)

      return {
        id: crypto.randomUUID(),
        FECHA: today,
        DNI: dni,
        'CODIGO 1': match?.codigo1 ?? '',
        'CODIGO 2': match?.codigo2 ?? '',
        'NOMBRES Y APELLIDOS': match?.nombre || nombreBd,
        OBSERVACIONES: '',
        'ALGO MAS': '',
        'ALGO 2': '',
        __matched: Boolean(match),
      }
    })
    .filter(Boolean)
}

export function buildPlanillaClipboardText(rows) {
  return rows
    .map((row) => [
      row.FECHA,
      row.DNI,
      row['CODIGO 1'],
      row['CODIGO 2'],
      row['NOMBRES Y APELLIDOS'],
      row.OBSERVACIONES,
      row['ALGO MAS'],
      row['ALGO 2'],
    ].join('\t'))
    .join('\r\n')
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
}

function buildHtmlCell(value, kind) {
  const safeValue = escapeHtml(value)

  if (kind === 'text') {
    return `<td style="mso-number-format:'\\@'; white-space:pre;">${safeValue}</td>`
  }

  if (kind === 'date') {
    return `<td style="mso-number-format:'dd/mm/yyyy'; white-space:pre;">${safeValue}</td>`
  }

  return `<td style="white-space:pre;">${safeValue}</td>`
}

export function buildPlanillaClipboardHtml(rows) {
  const htmlRows = rows
    .map((row) => [
      buildHtmlCell(row.FECHA, 'date'),
      buildHtmlCell(row.DNI, 'text'),
      buildHtmlCell(row['CODIGO 1'], 'text'),
      buildHtmlCell(row['CODIGO 2'], 'text'),
      buildHtmlCell(row['NOMBRES Y APELLIDOS'], 'default'),
      buildHtmlCell(row.OBSERVACIONES, 'default'),
      buildHtmlCell(row['ALGO MAS'], 'default'),
      buildHtmlCell(row['ALGO 2'], 'default'),
    ].join(''))
    .map((cells) => `<tr>${cells}</tr>`)
    .join('')

  return `<html><body><table>${htmlRows}</table></body></html>`
}

export function buildPlanillaWorkbook(rows) {
  const worksheet = XLSX.utils.aoa_to_sheet([
    ['FECHA', 'DNI', 'CODIGO 1', 'CODIGO 2', 'NOMBRES Y APELLIDOS', 'OBSERVACIONES', 'ALGO MAS', 'ALGO 2'],
    ...rows.map((row) => [
      row.FECHA,
      row.DNI,
      row['CODIGO 1'],
      row['CODIGO 2'],
      row['NOMBRES Y APELLIDOS'],
      row.OBSERVACIONES,
      row['ALGO MAS'],
      row['ALGO 2'],
    ]),
  ])

  rows.forEach((row, index) => {
    const dniCell = worksheet[`B${index + 2}`]
    const codigo1Cell = worksheet[`C${index + 2}`]
    const codigo2Cell = worksheet[`D${index + 2}`]

    ;[dniCell, codigo1Cell, codigo2Cell].forEach((cell, cellIndex) => {
      const values = [row.DNI, row['CODIGO 1'], row['CODIGO 2']]
      if (!cell) return
      cell.t = 's'
      cell.z = '@'
      cell.v = values[cellIndex]
    })
  })

  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Hoja1')
  return workbook
}
