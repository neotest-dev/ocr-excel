import { CopyIcon, DownloadIcon, PencilIcon, PlusIcon, TrashIcon } from './Icons'
import { getStatusLabel } from '../lib/records'

function StatusBadge({ status }) {
  const styles = {
    matched: 'bg-emerald-100 text-emerald-700 ring-1 ring-emerald-200',
    missing: 'bg-amber-100 text-amber-700 ring-1 ring-amber-200',
    pending: 'bg-orange-100 text-orange-700 ring-1 ring-orange-200',
    manual: 'bg-slate-100 text-slate-600 ring-1 ring-slate-200',
    empty: 'bg-slate-100 text-slate-500 ring-1 ring-slate-200',
  }

  return (
    <span className={`inline-flex rounded-full px-3 py-1 text-xs font-medium ${styles[status] || styles.manual}`}>
      {getStatusLabel(status)}
    </span>
  )
}

export function LensWorkspace({
  ocrText,
  rows,
  copyFeedback,
  activeRowId,
  bdFileName,
  rawLineCount,
  filledRows,
  onTextChange,
  onPasteText,
  onClearText,
  onBdFileChange,
  onRowChange,
  onRowFocus,
  onRowBlur,
  onAddRow,
  onDeleteRow,
  onCopy,
  onExport,
}) {
  return (
    <section className="grid gap-5 xl:grid-cols-[360px_minmax(0,1fr)]">
      <div className="overflow-hidden rounded-[28px] border border-emerald-900/20 bg-white shadow-[0_20px_60px_-30px_rgba(6,78,59,0.45)]">
        <div className="border-b border-emerald-100 bg-gradient-to-r from-emerald-700 via-emerald-600 to-green-600 px-5 py-5 text-white">
          <p className="text-xs font-semibold uppercase tracking-[0.24em] text-emerald-100/80">Captura</p>
          <h2 className="mt-2 text-2xl font-semibold tracking-tight">Pega desde Google Lens</h2>
          <p className="mt-2 max-w-sm text-sm text-emerald-50/85">
            Esta herramienta es aparte de Planilla. Aqui pegas texto desde Google Lens y lo cruzas solo con la base BD.
          </p>
        </div>

        <div className="space-y-4 p-5">
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              onClick={onPasteText}
              className="inline-flex items-center gap-2 rounded-xl bg-emerald-700 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-emerald-800"
            >
              <CopyIcon />
              Pegar
            </button>
            <button
              type="button"
              onClick={onClearText}
              className="inline-flex items-center gap-2 rounded-xl border border-rose-200 bg-rose-50 px-4 py-2.5 text-sm font-semibold text-rose-700 transition hover:bg-rose-100"
            >
              <TrashIcon />
              Limpiar
            </button>
          </div>

          <label className="flex cursor-pointer items-center justify-between gap-3 rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 transition hover:bg-white">
            <div>
              <p className="text-sm font-semibold text-slate-900">Base Excel</p>
              <p className="text-xs text-slate-500">Archivo `bd.xlsx` con hoja `bd`</p>
            </div>
            <div className="truncate rounded-lg bg-white px-3 py-2 text-xs text-slate-500 shadow-sm ring-1 ring-slate-200">
              {bdFileName || 'Seleccionar archivo'}
            </div>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={onBdFileChange} className="hidden" />
          </label>

          <textarea
            value={ocrText}
            onChange={(event) => onTextChange(event.target.value)}
            placeholder="Pega aqui"
            className="min-h-[520px] w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-4 font-mono text-sm text-slate-900 outline-none transition placeholder:text-slate-400 focus:border-emerald-400 focus:bg-white focus:ring-4 focus:ring-emerald-100"
          />

          <div className="grid grid-cols-2 gap-3">
            <div className="rounded-2xl bg-emerald-50 px-4 py-3 ring-1 ring-emerald-100">
              <p className="text-xs uppercase tracking-[0.16em] text-emerald-700">Lineas</p>
              <strong className="mt-1 block text-2xl font-semibold text-emerald-950">{rawLineCount}</strong>
            </div>
            <div className="rounded-2xl bg-slate-100 px-4 py-3 ring-1 ring-slate-200">
              <p className="text-xs uppercase tracking-[0.16em] text-slate-600">Filas utiles</p>
              <strong className="mt-1 block text-2xl font-semibold text-slate-950">{filledRows.length}</strong>
            </div>
          </div>
        </div>
      </div>

      <div className="overflow-hidden rounded-[28px] border border-emerald-900/10 bg-white shadow-[0_24px_80px_-40px_rgba(6,78,59,0.35)]">
        <div className="flex flex-col gap-4 border-b border-slate-200 bg-slate-50 px-5 py-4 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <p className="text-xs font-semibold uppercase tracking-[0.24em] text-slate-500">Resultado</p>
            <h2 className="mt-1 text-2xl font-semibold tracking-tight text-slate-950">Tabla revisable</h2>
            <p className="mt-1 text-sm text-slate-500">El estado te ayuda a ver qué filas sí salieron de la base.</p>
          </div>

          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              onClick={onAddRow}
              className="inline-flex items-center gap-2 rounded-xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-semibold text-slate-800 transition hover:bg-slate-100"
            >
              <PlusIcon />
              Agregar fila
            </button>
            <button
              type="button"
              onClick={onCopy}
              disabled={!filledRows.length}
              className="inline-flex items-center gap-2 rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-2.5 text-sm font-semibold text-emerald-700 transition hover:bg-emerald-100 disabled:cursor-not-allowed disabled:opacity-40"
            >
              <CopyIcon />
              Copiar
            </button>
            <button
              type="button"
              onClick={onExport}
              disabled={!filledRows.length}
              className="inline-flex items-center gap-2 rounded-xl bg-emerald-700 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-emerald-800 disabled:cursor-not-allowed disabled:opacity-40"
            >
              <DownloadIcon />
              Descargar
            </button>
          </div>
        </div>

        <div className="max-h-[760px] overflow-auto">
          <table className="min-w-full border-collapse text-sm">
            <thead className="sticky top-0 z-10 bg-emerald-950 text-white">
              <tr>
                <th className="border border-emerald-900 px-4 py-3 text-left font-medium">DNI</th>
                <th className="border border-emerald-900 px-4 py-3 text-left font-medium">APELLIDOS Y NOMBRES</th>
                <th className="border border-emerald-900 px-4 py-3 text-left font-medium">ESTADO</th>
                <th className="border border-emerald-900 px-4 py-3 text-left font-medium">ACCION</th>
              </tr>
            </thead>
            <tbody className="bg-white text-slate-900">
              {rows.map((row) => (
                <tr
                  key={row.id}
                  className={`transition ${activeRowId === row.id ? 'bg-emerald-50' : 'hover:bg-slate-50'}`}
                >
                  <td className="border border-slate-200 px-2 py-1.5">
                    <div className="relative">
                      <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                        <PencilIcon />
                      </span>
                      <input
                        value={row.number}
                        onChange={(event) => onRowChange(row.id, 'number', event.target.value)}
                        onFocus={() => onRowFocus(row.id)}
                        onBlur={() => onRowBlur(row.id)}
                        className="w-full rounded-xl border border-transparent bg-transparent py-2 pl-10 pr-3 outline-none transition focus:border-emerald-300 focus:bg-white focus:ring-4 focus:ring-emerald-100"
                      />
                    </div>
                  </td>
                  <td className="border border-slate-200 px-2 py-1.5">
                    <div className="relative">
                      <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                        <PencilIcon />
                      </span>
                      <input
                        value={row.name}
                        onChange={(event) => onRowChange(row.id, 'name', event.target.value)}
                        onFocus={() => onRowFocus(row.id)}
                        onBlur={() => onRowBlur(row.id)}
                        className="w-full rounded-xl border border-transparent bg-transparent py-2 pl-10 pr-3 outline-none transition focus:border-emerald-300 focus:bg-white focus:ring-4 focus:ring-emerald-100"
                      />
                    </div>
                  </td>
                  <td className="border border-slate-200 px-3 py-1.5">
                    <StatusBadge status={row.status} />
                  </td>
                  <td className="border border-slate-200 px-3 py-1.5">
                    <button
                      type="button"
                      onClick={() => onDeleteRow(row.id)}
                      className="inline-flex h-9 w-9 items-center justify-center rounded-xl border border-rose-200 bg-rose-50 text-rose-700 transition hover:bg-rose-100"
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

        <div className="border-t border-slate-200 bg-slate-50 px-5 py-3 text-sm text-slate-600">
          {copyFeedback || 'Copiar genera texto tabulado listo para pegar directo en Excel.'}
        </div>
      </div>
    </section>
  )
}
