import { CopyIcon, DownloadIcon } from './Icons'

const columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']

function SpreadsheetCell({ children, className = '' }) {
  return <div className={`h-7 overflow-hidden whitespace-nowrap border-r border-b border-[#d9d9d9] px-2 py-1 text-[12px] text-ellipsis ${className}`}>{children}</div>
}

export function SpreadsheetWorkspace({
  planillaBdFileName,
  planillaReportFileName,
  planillaRows,
  planillaStatus,
  onPlanillaBdFileChange,
  onPlanillaReportFileChange,
  onPlanillaProcess,
  onPlanillaCopy,
  onPlanillaExport,
}) {
  const totalRows = planillaRows.length
  const totalMatch = planillaRows.filter((row) => row.__matched).length
  const totalMissing = planillaRows.filter((row) => !row.__matched).length

  function buildGridRows() {
    const renderedRows = []
    renderedRows.push(['FECHA', 'DNI', 'CODIGO 1', 'CODIGO 2', 'NOMBRES Y APELLIDOS', 'OBSERVACIONES', 'ALGO MAS', 'ALGO 2'])

    planillaRows.forEach((row) => {
      renderedRows.push([
        row.FECHA,
        row.DNI,
        row['CODIGO 1'],
        row['CODIGO 2'],
        row['NOMBRES Y APELLIDOS'],
        row.OBSERVACIONES,
        row['ALGO MAS'],
        row['ALGO 2'],
      ])
    })

    const minRows = Math.max(22, planillaRows.length + 8)

    while (renderedRows.length < minRows) {
      renderedRows.push([])
    }

    return renderedRows
  }

  const gridRows = buildGridRows()

  return (
    <section className="overflow-hidden rounded-[30px] border border-[#214b34] bg-[#141414] shadow-[0_30px_90px_-50px_rgba(16,124,65,0.6)]">
      <div className="border-b border-[#2a2a2a] bg-[#111111] px-4 py-3 text-white">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div>
            <p className="text-xs uppercase tracking-[0.22em] text-emerald-300/75">Planilla</p>
            <h2 className="mt-1 text-2xl font-semibold tracking-tight">Cruce entre BD y REPORTE</h2>
          </div>
          <div className="rounded-md border border-[#3a3a3a] bg-[#262626] px-3 py-2 text-sm text-[#bdbdbd]">
            Libro local
          </div>
        </div>
      </div>

      <div className="border-b border-[#2a2a2a] bg-[#151515] px-4 py-2 text-sm text-[#e9e9e9]">
        <div className="flex flex-wrap gap-4">
          <span className="border-b-2 border-[#21a366] pb-1 font-semibold">Inicio</span>
          <span>Insertar</span>
          <span>Datos</span>
          <span>Vista</span>
        </div>
      </div>

      <div className="border-b border-[#2a2a2a] bg-[#1b1b1b] px-4 py-4">
        <div className="grid gap-4 xl:grid-cols-[1.2fr_1fr_1fr_1.2fr]">
          <div className="space-y-2 border-r border-[#3a3a3a] pr-4 last:border-r-0">
            <div className="flex flex-wrap gap-2">
              <label className="inline-flex max-w-full cursor-pointer items-center gap-2 rounded-md border border-[#3a3a3a] bg-[#262626] px-3 py-2 text-sm text-[#f3f3f3]">
                <span>BD</span>
                <span className="max-w-40 truncate text-[#bdbdbd]">{planillaBdFileName || 'bd.xlsx'}</span>
                <input type="file" accept=".xlsx,.xls,.xlsm" onChange={onPlanillaBdFileChange} className="hidden" />
              </label>
              <label className="inline-flex max-w-full cursor-pointer items-center gap-2 rounded-md border border-[#3a3a3a] bg-[#262626] px-3 py-2 text-sm text-[#f3f3f3]">
                <span>REPORTE</span>
                <span className="max-w-40 truncate text-[#bdbdbd]">{planillaReportFileName || 'REPORTE.xlsx'}</span>
                <input type="file" accept=".xlsx,.xls,.xlsm" onChange={onPlanillaReportFileChange} className="hidden" />
              </label>
            </div>
            <div className="text-[11px] text-[#b7b7b7]">Archivos</div>
          </div>

          <div className="space-y-2 border-r border-[#3a3a3a] pr-4">
            <div className="flex flex-wrap gap-2">
              <button
                type="button"
                onClick={onPlanillaProcess}
                className="rounded-md border border-[#0e6938] bg-[#107c41] px-3 py-2 text-sm font-semibold text-white"
              >
                Procesar
              </button>
              <button
                type="button"
                onClick={onPlanillaCopy}
                disabled={!planillaRows.length}
                className="inline-flex items-center gap-2 rounded-md border border-[#355f4b] bg-[#20362b] px-3 py-2 text-sm text-[#d8ffe9] disabled:cursor-not-allowed disabled:opacity-40"
              >
                <CopyIcon className="h-4 w-4" />
                Copiar para Excel
              </button>
              <button
                type="button"
                onClick={onPlanillaExport}
                disabled={!planillaRows.length}
                className="inline-flex items-center gap-2 rounded-md border border-[#355f4b] bg-[#20362b] px-3 py-2 text-sm text-[#d8ffe9] disabled:cursor-not-allowed disabled:opacity-40"
              >
                <DownloadIcon className="h-4 w-4" />
                Descargar
              </button>
            </div>
            <div className="text-[11px] text-[#b7b7b7]">Acciones</div>
          </div>

          <div className="space-y-2 border-r border-[#3a3a3a] pr-4">
            <div className="flex flex-wrap gap-2 text-sm text-[#f3f3f3]">
              <div className="rounded-md border border-[#3a3a3a] bg-[#262626] px-3 py-2">BD: {totalRows}</div>
              <div className="rounded-md border border-[#3a3a3a] bg-[#262626] px-3 py-2">Match: {totalMatch}</div>
              <div className="rounded-md border border-[#3a3a3a] bg-[#262626] px-3 py-2">Manual: {totalMissing}</div>
            </div>
            <div className="text-[11px] text-[#b7b7b7]">Resumen</div>
          </div>

          <div className="space-y-2">
            <div
              className={`rounded-md border px-3 py-2 text-sm ${
                planillaStatus.type === 'success'
                  ? 'border-[#2f6c4c] bg-[#163323] text-[#d8ffe9]'
                  : planillaStatus.type === 'error'
                    ? 'border-[#6b2f2f] bg-[#351919] text-[#ffe2e2]'
                    : 'border-[#3a3a3a] bg-[#1f1f1f] text-[#bdbdbd]'
              }`}
            >
              {planillaStatus.message}
            </div>
            <div className="text-[11px] text-[#b7b7b7]">Estado</div>
          </div>
        </div>
      </div>

      <div className="border-b border-[#2a2a2a] bg-[#1a1a1a] px-4 py-2">
        <div className="grid grid-cols-[78px_minmax(0,1fr)] items-center gap-3 text-sm">
          <div className="rounded border border-[#3a3a3a] bg-[#262626] px-2 py-1 text-center text-[#f3f3f3]">A3</div>
          <div className="rounded border border-[#3a3a3a] bg-[#262626] px-3 py-1.5 text-[#bdbdbd]">
            {planillaRows[0]?.FECHA || '22/06/2026'}
          </div>
        </div>
      </div>

      <div className="overflow-auto bg-[#111111]">
        <div className="min-w-[1360px] bg-white text-black">
          <div className="grid grid-cols-[52px_repeat(14,minmax(80px,1fr))]">
            <div className="flex h-7 items-center justify-center border-r border-b border-[#4a4a4a] bg-[#0f0f10] text-xs text-white" />
            {columns.map((column) => (
              <div key={column} className="flex h-7 items-center justify-center border-r border-b border-[#4a4a4a] bg-[#0f0f10] text-xs text-white">
                {column}
              </div>
            ))}
          </div>

          {gridRows.map((row, rowIndex) => (
            <div key={`row-${rowIndex}`} className="grid grid-cols-[52px_repeat(14,minmax(80px,1fr))]">
              <div className="flex h-7 items-center justify-center border-r border-b border-[#4a4a4a] bg-[#0f0f10] text-xs text-white">
                {rowIndex + 1}
              </div>

              {Array.from({ length: 14 }, (_, colIndex) => {
                const value = row[colIndex] ?? ''
                const isHeaderRow = rowIndex === 0 && colIndex < 8
                const isMiss = rowIndex > 0 && planillaRows[rowIndex - 1] && !planillaRows[rowIndex - 1].__matched && colIndex < 5
                const isSelected = rowIndex === 2 && colIndex === 0

                return (
                  <SpreadsheetCell
                    key={`${rowIndex}-${colIndex}`}
                    className={`${isHeaderRow ? 'bg-[#fcfcfc] font-semibold' : ''} ${isMiss ? 'bg-[#fff9db]' : ''} ${colIndex >= 8 ? 'text-[#666]' : ''} ${isSelected ? 'relative outline-2 -outline-offset-2 outline-[#217346]' : ''}`}
                  >
                    {value}
                  </SpreadsheetCell>
                )
              })}
            </div>
          ))}
        </div>
      </div>

      <div className="flex items-center justify-between border-t border-[#2a2a2a] bg-[#101010] px-4 py-2 text-xs text-[#bdbdbd]">
        <div>Listo</div>
        <div className="flex gap-4">
          <span>Hoja1</span>
          <span>100%</span>
        </div>
      </div>
    </section>
  )
}
