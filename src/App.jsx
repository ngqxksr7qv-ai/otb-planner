import { useState, useMemo, useCallback, useRef, Fragment } from 'react'
import * as XLSX from 'xlsx'
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts'

const MONTHS = ['MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN', 'FEB']

const PRESETS = {
  small: { warning: 1000, critical: 2500, label: 'Small store' },
  mid: { warning: 2000, critical: 5000, label: 'Mid-volume' },
  high: { warning: 5000, critical: 10000, label: 'High-volume' },
}

const DEFAULT_SETTINGS = {
  warningThreshold: 2000,
  criticalThreshold: 5000,
  flagNegativeOTB: true,
  flagZeroOnOrder: true,
}

function formatCurrency(val) {
  if (val == null || isNaN(val)) return '\u2014'
  const neg = val < 0
  const formatted = Math.abs(val).toLocaleString('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0, maximumFractionDigits: 0 })
  return neg ? `-${formatted}` : formatted
}

function parseFile(data) {
  const workbook = XLSX.read(data, { type: 'array' })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })

  if (rows.length < 3) throw new Error('File has insufficient data rows.')

  // Find header row by matching known column labels
  let headerRowIdx = -1
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const row = rows[i].map(c => String(c).trim().toUpperCase())
    if (row.includes('CATEGORY') && row.includes('CLASS') && row.includes('DATA')) {
      headerRowIdx = i
      break
    }
  }
  if (headerRowIdx === -1) throw new Error('Could not find header row with expected columns.')

  const headers = rows[headerRowIdx].map(c => String(c).trim())

  // Build column index map by label (handles out-of-order columns)
  const colIdx = {}
  const upperHeaders = headers.map(h => h.toUpperCase())
  colIdx.category = upperHeaders.indexOf('CATEGORY')
  colIdx.classCode = upperHeaders.indexOf('CLASS')
  colIdx.classDesc = upperHeaders.indexOf('CLASS DESC')
  colIdx.store = upperHeaders.indexOf('STORE')
  colIdx.data = upperHeaders.indexOf('DATA')

  // Map month columns
  colIdx.months = {}
  for (const m of MONTHS) {
    const idx = upperHeaders.indexOf(m)
    if (idx !== -1) colIdx.months[m] = idx
  }

  const dataRows = rows.slice(headerRowIdx + 1)
  const classes = []
  const stores = new Set()

  // Group rows into class records (4 rows per class)
  let i = 0
  while (i < dataRows.length) {
    const row = dataRows[i]
    if (!row || row.length === 0 || !String(row[colIdx.data] || '').trim()) {
      i++
      continue
    }

    const category = String(row[colIdx.category] || '').trim()
    const classCode = row[colIdx.classCode]
    const classDesc = String(row[colIdx.classDesc] || '').trim()
    const store = String(row[colIdx.store] || '').trim()

    if (store) stores.add(store)

    // Collect up to 4 rows that share the same category+class+store grouping
    const group = {}
    let j = i
    while (j < dataRows.length && j < i + 4) {
      const r = dataRows[j]
      const dataLabel = String(r[colIdx.data] || '').trim()
      if (!dataLabel) { j++; continue }

      // Check if this row belongs to the same class group
      const rCat = String(r[colIdx.category] || '').trim()
      const rClass = r[colIdx.classCode]
      const rStore = String(r[colIdx.store] || '').trim()
      if (j > i && (rCat !== category || String(rClass) !== String(classCode) || rStore !== store)) break

      const monthValues = {}
      for (const m of MONTHS) {
        if (colIdx.months[m] != null) {
          const v = r[colIdx.months[m]]
          monthValues[m] = typeof v === 'number' ? v : parseFloat(String(v).replace(/[,$]/g, '')) || 0
        } else {
          monthValues[m] = 0
        }
      }

      if (dataLabel === 'Sales Forecast') group.salesForecast = monthValues
      else if (dataLabel === 'On Order Retail') group.onOrderRetail = monthValues
      else if (dataLabel === 'On Order Cost') group.onOrderCost = monthValues
      else if (dataLabel === 'OTB Cost Original') group.otbCostOriginal = monthValues

      j++
    }
    i = j

    const salesForecast = group.salesForecast || null
    const onOrderRetail = group.onOrderRetail || null
    const onOrderCost = group.onOrderCost || null
    const otbCostOriginal = group.otbCostOriginal || null

    const sumValues = (obj) => obj ? MONTHS.reduce((s, m) => s + (obj[m] || 0), 0) : 0
    const annualForecast = sumValues(salesForecast)
    const annualOTB = sumValues(otbCostOriginal)
    const annualOnOrderCost = sumValues(onOrderCost)
    const otbUtilizationPct = annualOTB !== 0 ? (annualOnOrderCost / annualOTB) * 100 : 0

    const monthlyVariance = {}
    for (const m of MONTHS) {
      const otb = otbCostOriginal ? (otbCostOriginal[m] || 0) : 0
      const forecast = salesForecast ? (salesForecast[m] || 0) : 0
      monthlyVariance[m] = otb - forecast
    }

    const isAllZero = annualForecast === 0 && annualOTB === 0

    classes.push({
      id: `${category}-${classCode}-${store}`,
      category,
      classCode: typeof classCode === 'number' ? classCode : parseInt(classCode) || classCode,
      classDesc,
      store,
      salesForecast,
      onOrderRetail,
      onOrderCost,
      otbCostOriginal,
      annualForecast,
      annualOTB,
      annualOnOrderCost,
      otbUtilizationPct,
      monthlyVariance,
      isAllZero,
      alertMonths: [],
      alertLevel: 'none',
    })
  }

  return { classes, stores: Array.from(stores).sort() }
}

function computeAlerts(classes, settings) {
  return classes.map(cls => {
    if (cls.isAllZero) return { ...cls, alertLevel: 'none', alertMonths: [] }

    const alertMonths = []
    let maxLevel = 'none'

    for (const m of MONTHS) {
      const absVariance = Math.abs(cls.monthlyVariance[m])
      if (absVariance >= settings.criticalThreshold) {
        alertMonths.push(m)
        maxLevel = 'critical'
      } else if (absVariance >= settings.warningThreshold) {
        alertMonths.push(m)
        if (maxLevel !== 'critical') maxLevel = 'warning'
      }
    }

    // Flag negative OTB months
    if (settings.flagNegativeOTB && cls.otbCostOriginal) {
      for (const m of MONTHS) {
        if (cls.otbCostOriginal[m] < 0 && !alertMonths.includes(m)) {
          alertMonths.push(m)
          if (maxLevel === 'none') maxLevel = 'warning'
        }
      }
    }

    // Flag zero on order with large OTB
    if (settings.flagZeroOnOrder && cls.annualOnOrderCost === 0 && cls.annualOTB > 10000) {
      if (maxLevel === 'none') maxLevel = 'warning'
    }

    // Negative annual OTB is always critical
    if (cls.annualOTB < 0) {
      maxLevel = 'critical'
    }

    return { ...cls, alertLevel: maxLevel, alertMonths }
  })
}

const ALERT_ORDER = { critical: 0, warning: 1, none: 2 }

function sortClasses(classes, sortConfig) {
  const sorted = [...classes]
  sorted.sort((a, b) => {
    const comparators = [sortConfig, sortConfig.secondary].filter(Boolean)
    for (const cfg of comparators) {
      let cmp = 0
      switch (cfg.column) {
        case 'alert':
          cmp = ALERT_ORDER[a.alertLevel] - ALERT_ORDER[b.alertLevel]
          break
        case 'category':
          cmp = a.category.localeCompare(b.category)
          break
        case 'classCode':
          cmp = (a.classCode || 0) - (b.classCode || 0)
          break
        case 'annualForecast':
          cmp = b.annualForecast - a.annualForecast
          break
        case 'annualOTB':
          cmp = b.annualOTB - a.annualOTB
          break
        case 'otbUtilizationPct':
          cmp = a.otbUtilizationPct - b.otbUtilizationPct
          break
        default:
          cmp = 0
      }
      if (cfg.direction === 'desc') cmp = -cmp
      if (cmp !== 0) return cmp
    }
    return 0
  })
  return sorted
}

// ── Components ──

function FileUpload({ onFileLoaded }) {
  const [dragOver, setDragOver] = useState(false)
  const [error, setError] = useState(null)
  const inputRef = useRef(null)

  const processFile = useCallback((file) => {
    setError(null)
    if (!file.name.endsWith('.xlsx')) {
      setError('Please upload an .xlsx file in the Kitchen Collage OTB format.')
      return
    }
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const result = parseFile(new Uint8Array(e.target.result))
        onFileLoaded(result)
      } catch (err) {
        setError(`Error parsing file: ${err.message}`)
      }
    }
    reader.readAsArrayBuffer(file)
  }, [onFileLoaded])

  const handleDrop = (e) => {
    e.preventDefault()
    setDragOver(false)
    if (e.dataTransfer.files.length > 0) processFile(e.dataTransfer.files[0])
  }

  return (
    <div className="flex items-center gap-3">
      <div
        className={`border-2 border-dashed rounded-lg px-4 py-2 cursor-pointer transition-colors text-sm ${dragOver ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-gray-400'}`}
        onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
        onDragLeave={() => setDragOver(false)}
        onDrop={handleDrop}
        onClick={() => inputRef.current?.click()}
      >
        Drop .xlsx or click to upload
        <input ref={inputRef} type="file" accept=".xlsx" className="hidden" onChange={(e) => e.target.files[0] && processFile(e.target.files[0])} />
      </div>
      {error && <span className="text-red-600 text-sm">{error}</span>}
    </div>
  )
}

function SettingsPanel({ settings, onSettingsChange }) {
  const [local, setLocal] = useState(settings)
  const [isCustom, setIsCustom] = useState(false)

  const applyPreset = (key) => {
    const p = PRESETS[key]
    const next = { ...local, warningThreshold: p.warning, criticalThreshold: p.critical }
    setLocal(next)
    setIsCustom(false)
  }

  const setCustom = () => {
    setLocal({ ...local, warningThreshold: '', criticalThreshold: '' })
    setIsCustom(true)
  }

  return (
    <div className="bg-gray-50 border-b border-gray-200 px-6 py-4">
      <div className="flex flex-wrap items-center gap-3 mb-3">
        <span className="text-sm font-medium text-gray-700">Presets:</span>
        {Object.entries(PRESETS).map(([key, p]) => (
          <button key={key} onClick={() => applyPreset(key)} className="px-3 py-1 text-xs rounded-full border border-gray-300 hover:bg-gray-200 transition-colors bg-white">{p.label}</button>
        ))}
        <button onClick={setCustom} className={`px-3 py-1 text-xs rounded-full border transition-colors ${isCustom ? 'bg-blue-100 border-blue-400' : 'border-gray-300 hover:bg-gray-200 bg-white'}`}>Custom</button>
      </div>
      <div className="flex flex-wrap items-end gap-6">
        <label className="flex flex-col text-sm">
          <span className="text-gray-600 mb-1">Warning threshold ($)</span>
          <input type="number" value={local.warningThreshold} onChange={(e) => setLocal({ ...local, warningThreshold: e.target.value === '' ? '' : Number(e.target.value) })}
            className="w-32 px-2 py-1 border border-gray-300 rounded text-sm" />
        </label>
        <label className="flex flex-col text-sm">
          <span className="text-gray-600 mb-1">Critical threshold ($)</span>
          <input type="number" value={local.criticalThreshold} onChange={(e) => setLocal({ ...local, criticalThreshold: e.target.value === '' ? '' : Number(e.target.value) })}
            className="w-32 px-2 py-1 border border-gray-300 rounded text-sm" />
        </label>
        <label className="flex items-center gap-2 text-sm cursor-pointer">
          <input type="checkbox" checked={local.flagNegativeOTB} onChange={(e) => setLocal({ ...local, flagNegativeOTB: e.target.checked })} className="rounded" />
          <span className="text-gray-700">Flag negative OTB months</span>
        </label>
        <label className="flex items-center gap-2 text-sm cursor-pointer">
          <input type="checkbox" checked={local.flagZeroOnOrder} onChange={(e) => setLocal({ ...local, flagZeroOnOrder: e.target.checked })} className="rounded" />
          <span className="text-gray-700">Flag zero On Order / large OTB</span>
        </label>
        <button onClick={() => onSettingsChange(local)}
          className="px-4 py-1.5 bg-blue-600 text-white text-sm rounded hover:bg-blue-700 transition-colors">
          Apply
        </button>
      </div>
    </div>
  )
}

function SummaryBar({ data }) {
  const totals = useMemo(() => {
    const totalForecast = data.reduce((s, c) => s + c.annualForecast, 0)
    const totalOTB = data.reduce((s, c) => s + c.annualOTB, 0)
    const totalOnOrder = data.reduce((s, c) => s + c.annualOnOrderCost, 0)
    const criticalCount = data.filter(c => c.alertLevel === 'critical').length
    const warningCount = data.filter(c => c.alertLevel === 'warning').length
    return { totalForecast, totalOTB, totalOnOrder, criticalCount, warningCount }
  }, [data])

  const tiles = [
    { label: 'Total Annual Forecast', value: formatCurrency(totals.totalForecast), color: 'text-gray-900' },
    { label: 'Total Annual OTB', value: formatCurrency(totals.totalOTB), color: totals.totalOTB < 0 ? 'text-red-600' : 'text-gray-900' },
    { label: 'Total On Order Cost', value: formatCurrency(totals.totalOnOrder), color: 'text-gray-900' },
    { label: 'Critical Alerts', value: totals.criticalCount, color: 'text-[#C0392B]', bg: 'bg-red-50' },
    { label: 'Warning Alerts', value: totals.warningCount, color: 'text-[#F39C12]', bg: 'bg-yellow-50' },
  ]

  return (
    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-4 px-6 py-4">
      {tiles.map(t => (
        <div key={t.label} className={`rounded-lg border border-gray-200 p-4 ${t.bg || 'bg-white'}`}>
          <div className="text-xs text-gray-500 uppercase tracking-wide">{t.label}</div>
          <div className={`text-2xl font-semibold mt-1 ${t.color}`}>{t.value}</div>
        </div>
      ))}
    </div>
  )
}

function FilterBar({ categories, filters, onFilterChange }) {
  const [catOpen, setCatOpen] = useState(false)
  const catRef = useRef(null)

  const toggleCategory = (cat) => {
    const next = filters.categories.includes(cat)
      ? filters.categories.filter(c => c !== cat)
      : [...filters.categories, cat]
    onFilterChange({ ...filters, categories: next })
  }

  const selectAllCats = () => onFilterChange({ ...filters, categories: [...categories] })
  const clearAllCats = () => onFilterChange({ ...filters, categories: [] })

  return (
    <div className="flex flex-wrap items-center gap-4 px-6 py-3 bg-white border-b border-gray-200">
      {/* Category multi-select */}
      <div className="relative" ref={catRef}>
        <button onClick={() => setCatOpen(!catOpen)} className="px-3 py-1.5 border border-gray-300 rounded text-sm bg-white hover:bg-gray-50 flex items-center gap-1">
          Categories ({filters.categories.length}/{categories.length})
          <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" /></svg>
        </button>
        {catOpen && (
          <div className="absolute z-50 top-full mt-1 left-0 bg-white border border-gray-200 rounded-lg shadow-lg p-2 max-h-64 overflow-y-auto w-56">
            <div className="flex gap-2 mb-2 pb-2 border-b border-gray-100">
              <button onClick={selectAllCats} className="text-xs text-blue-600 hover:underline">Select all</button>
              <button onClick={clearAllCats} className="text-xs text-blue-600 hover:underline">Clear all</button>
            </div>
            {categories.map(cat => (
              <label key={cat} className="flex items-center gap-2 py-0.5 text-sm cursor-pointer hover:bg-gray-50 px-1 rounded">
                <input type="checkbox" checked={filters.categories.includes(cat)} onChange={() => toggleCategory(cat)} className="rounded" />
                {cat}
              </label>
            ))}
          </div>
        )}
      </div>

      {/* Alert level filter */}
      <select value={filters.alertLevel} onChange={(e) => onFilterChange({ ...filters, alertLevel: e.target.value })}
        className="px-3 py-1.5 border border-gray-300 rounded text-sm bg-white">
        <option value="all">All alerts</option>
        <option value="critical">Critical only</option>
        <option value="warnings">Warnings + Critical</option>
        <option value="none">No alerts</option>
      </select>

      {/* Min forecast filter */}
      <label className="flex items-center gap-2 text-sm">
        <span className="text-gray-600">Min forecast $</span>
        <input type="number" value={filters.minForecast} onChange={(e) => onFilterChange({ ...filters, minForecast: Number(e.target.value) || 0 })}
          className="w-24 px-2 py-1.5 border border-gray-300 rounded text-sm" placeholder="0" />
      </label>

      {/* Search */}
      <input type="text" value={filters.search} onChange={(e) => onFilterChange({ ...filters, search: e.target.value })}
        className="px-3 py-1.5 border border-gray-300 rounded text-sm w-48" placeholder="Search class description..." />
    </div>
  )
}

function UtilizationBar({ pct }) {
  const clampedPct = Math.min(100, Math.max(0, pct))
  let color = '#C0392B'
  if (pct >= 80) color = '#27AE60'
  else if (pct >= 50) color = '#F39C12'

  return (
    <div className="flex items-center gap-2">
      <div className="w-16 h-2.5 bg-gray-200 rounded-full overflow-hidden">
        <div className="h-full rounded-full transition-all" style={{ width: `${clampedPct}%`, backgroundColor: color }} />
      </div>
      <span className="text-xs" style={{ color }}>{pct.toFixed(0)}%</span>
    </div>
  )
}

function MonthCell({ value, variance, settings, otbValue }) {
  let bg = ''
  const absVariance = Math.abs(variance)

  if (absVariance >= settings.criticalThreshold) {
    bg = 'bg-red-200'
  } else if (absVariance >= settings.warningThreshold) {
    bg = 'bg-yellow-100'
  } else if (otbValue < 0) {
    bg = 'bg-red-100'
  }

  return (
    <td className={`px-2 py-1.5 text-right text-xs whitespace-nowrap ${bg} ${value < 0 ? 'text-red-600 font-medium' : 'text-gray-700'}`}>
      {formatCurrency(value)}
    </td>
  )
}

function DetailDrawer({ cls }) {
  const chartData = MONTHS.map(m => ({
    month: m,
    'Sales Forecast': cls.salesForecast?.[m] || 0,
    'OTB Cost Original': cls.otbCostOriginal?.[m] || 0,
    'Surplus/Deficit': cls.monthlyVariance[m],
  }))

  const overCommitted = MONTHS.filter(m => cls.otbCostOriginal?.[m] < 0)
  const worstMonth = overCommitted.length > 0
    ? overCommitted.reduce((worst, m) => (cls.otbCostOriginal[m] < cls.otbCostOriginal[worst] ? m : worst))
    : null

  return (
    <tr>
      <td colSpan={100} className="bg-gray-50 px-6 py-4 border-t border-gray-200">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div>
            <h4 className="text-sm font-semibold text-gray-700 mb-3">Monthly Comparison</h4>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={chartData} margin={{ top: 5, right: 5, bottom: 5, left: 5 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} tickFormatter={v => `$${(v / 1000).toFixed(0)}k`} />
                <Tooltip formatter={(v) => formatCurrency(v)} />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="Sales Forecast" fill="#3B82F6" />
                <Bar dataKey="OTB Cost Original" fill="#F97316" />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div>
            <h4 className="text-sm font-semibold text-gray-700 mb-3">Alert Summary</h4>
            <div className="text-sm text-gray-600 space-y-1">
              {overCommitted.length > 0 ? (
                <p className="text-red-600 font-medium">
                  {overCommitted.length} month{overCommitted.length > 1 ? 's are' : ' is'} over-committed.
                  {worstMonth && ` Largest exposure: ${worstMonth} at ${formatCurrency(cls.otbCostOriginal[worstMonth])}.`}
                </p>
              ) : (
                <p className="text-green-600">No over-committed months.</p>
              )}
              {cls.alertMonths.length > 0 && (
                <p>Alert months: {cls.alertMonths.join(', ')}</p>
              )}
              <p>Annual Forecast: {formatCurrency(cls.annualForecast)}</p>
              <p>Annual OTB: <span className={cls.annualOTB < 0 ? 'text-red-600' : ''}>{formatCurrency(cls.annualOTB)}</span></p>
              <p>Annual On Order: {formatCurrency(cls.annualOnOrderCost)}</p>
              <p>OTB Utilization: {cls.otbUtilizationPct.toFixed(1)}%</p>
              {cls.isAllZero && <p className="text-gray-400 italic">No activity</p>}
            </div>
            <h4 className="text-sm font-semibold text-gray-700 mt-4 mb-2">Surplus / Deficit by Month</h4>
            <ResponsiveContainer width="100%" height={140}>
              <BarChart data={chartData} margin={{ top: 5, right: 5, bottom: 5, left: 5 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} tickFormatter={v => `$${(v / 1000).toFixed(0)}k`} />
                <Tooltip formatter={(v) => formatCurrency(v)} />
                <Bar dataKey="Surplus/Deficit" fill="#6B7280" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </td>
    </tr>
  )
}

function SortHeader({ label, column, sortConfig, onSort }) {
  const isActive = sortConfig.column === column
  const dir = isActive ? sortConfig.direction : null

  return (
    <th
      className="px-2 py-2 text-left text-xs font-medium text-gray-600 uppercase tracking-wider cursor-pointer hover:bg-gray-100 select-none whitespace-nowrap"
      onClick={() => {
        if (isActive) {
          onSort({ ...sortConfig, direction: dir === 'asc' ? 'desc' : 'asc' })
        } else {
          const defaultDir = ['annualForecast', 'annualOTB'].includes(column) ? 'desc' : 'asc'
          onSort({ ...sortConfig, column, direction: defaultDir })
        }
      }}
    >
      {label}
      {isActive && <span className="ml-1">{dir === 'asc' ? '\u25B2' : '\u25BC'}</span>}
    </th>
  )
}

export default function App() {
  const [rawData, setRawData] = useState(null)
  const [stores, setStores] = useState([])
  const [selectedStore, setSelectedStore] = useState('ALL')
  const [settings, setSettings] = useState(DEFAULT_SETTINGS)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [sortConfig, setSortConfig] = useState({ column: 'alert', direction: 'asc' })
  const [filters, setFilters] = useState({ categories: [], alertLevel: 'all', minForecast: 0, search: '' })
  const [showMonthly, setShowMonthly] = useState(true)
  const [selectedRow, setSelectedRow] = useState(null)

  const handleFileLoaded = useCallback(({ classes, stores: fileStores }) => {
    setRawData(classes)
    setStores(fileStores)
    setSelectedStore(fileStores.includes('MAIN') ? 'MAIN' : fileStores[0] || 'ALL')
    const cats = [...new Set(classes.map(c => c.category))].sort()
    setFilters(f => ({ ...f, categories: cats }))
    setSelectedRow(null)
  }, [])

  // Filter by store, then compute alerts
  const storeFiltered = useMemo(() => {
    if (!rawData) return []
    if (selectedStore === 'ALL') {
      // Aggregate across stores: group by category+classCode+classDesc
      const grouped = {}
      for (const cls of rawData) {
        const key = `${cls.category}-${cls.classCode}-${cls.classDesc}`
        if (!grouped[key]) {
          grouped[key] = {
            ...cls,
            id: key,
            store: 'ALL',
            salesForecast: cls.salesForecast ? { ...cls.salesForecast } : null,
            onOrderRetail: cls.onOrderRetail ? { ...cls.onOrderRetail } : null,
            onOrderCost: cls.onOrderCost ? { ...cls.onOrderCost } : null,
            otbCostOriginal: cls.otbCostOriginal ? { ...cls.otbCostOriginal } : null,
          }
        } else {
          const g = grouped[key]
          for (const m of MONTHS) {
            if (cls.salesForecast && g.salesForecast) g.salesForecast[m] = (g.salesForecast[m] || 0) + (cls.salesForecast[m] || 0)
            if (cls.onOrderRetail && g.onOrderRetail) g.onOrderRetail[m] = (g.onOrderRetail[m] || 0) + (cls.onOrderRetail[m] || 0)
            if (cls.onOrderCost && g.onOrderCost) g.onOrderCost[m] = (g.onOrderCost[m] || 0) + (cls.onOrderCost[m] || 0)
            if (cls.otbCostOriginal && g.otbCostOriginal) g.otbCostOriginal[m] = (g.otbCostOriginal[m] || 0) + (cls.otbCostOriginal[m] || 0)
          }
        }
      }
      // Recompute derived fields
      return Object.values(grouped).map(g => {
        const annualForecast = g.salesForecast ? MONTHS.reduce((s, m) => s + (g.salesForecast[m] || 0), 0) : 0
        const annualOTB = g.otbCostOriginal ? MONTHS.reduce((s, m) => s + (g.otbCostOriginal[m] || 0), 0) : 0
        const annualOnOrderCost = g.onOrderCost ? MONTHS.reduce((s, m) => s + (g.onOrderCost[m] || 0), 0) : 0
        const monthlyVariance = {}
        for (const m of MONTHS) monthlyVariance[m] = (g.otbCostOriginal?.[m] || 0) - (g.salesForecast?.[m] || 0)
        return {
          ...g,
          annualForecast,
          annualOTB,
          annualOnOrderCost,
          otbUtilizationPct: annualOTB !== 0 ? (annualOnOrderCost / annualOTB) * 100 : 0,
          monthlyVariance,
          isAllZero: annualForecast === 0 && annualOTB === 0,
        }
      })
    }
    return rawData.filter(c => c.store === selectedStore)
  }, [rawData, selectedStore])

  const alertedData = useMemo(() => computeAlerts(storeFiltered, settings), [storeFiltered, settings])

  const categories = useMemo(() => [...new Set(alertedData.map(c => c.category))].sort(), [alertedData])

  const filteredData = useMemo(() => {
    return alertedData.filter(cls => {
      if (filters.categories.length > 0 && !filters.categories.includes(cls.category)) return false
      if (filters.alertLevel === 'critical' && cls.alertLevel !== 'critical') return false
      if (filters.alertLevel === 'warnings' && cls.alertLevel === 'none') return false
      if (filters.alertLevel === 'none' && cls.alertLevel !== 'none') return false
      if (filters.minForecast > 0 && cls.annualForecast < filters.minForecast) return false
      if (filters.search && !cls.classDesc.toLowerCase().includes(filters.search.toLowerCase())) return false
      return true
    })
  }, [alertedData, filters])

  const sortedData = useMemo(() => sortClasses(filteredData, sortConfig), [filteredData, sortConfig])

  return (
    <div className="min-h-screen bg-gray-50 font-sans text-[13px]">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4 flex-wrap">
        <h1 className="text-lg font-bold text-gray-900 mr-4">OTB Dashboard</h1>
        <FileUpload onFileLoaded={handleFileLoaded} />
        {stores.length > 0 && (
          <select value={selectedStore} onChange={(e) => setSelectedStore(e.target.value)}
            className="px-3 py-1.5 border border-gray-300 rounded text-sm bg-white">
            <option value="ALL">All Stores</option>
            {stores.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        )}
        <div className="ml-auto">
          <button onClick={() => setSettingsOpen(!settingsOpen)}
            className="p-2 rounded hover:bg-gray-100 transition-colors" title="Settings">
            <svg className="w-5 h-5 text-gray-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.066 2.573c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.573 1.066c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.066-2.573c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
            </svg>
          </button>
        </div>
      </header>

      {/* Settings panel */}
      {settingsOpen && (
        <SettingsPanel settings={settings} onSettingsChange={setSettings} />
      )}

      {/* Content */}
      {!rawData ? (
        <div className="flex items-center justify-center h-96 text-gray-400 text-lg">
          Upload an OTB Summary (.xlsx) to get started
        </div>
      ) : (
        <>
          <SummaryBar data={alertedData} />
          <FilterBar categories={categories} filters={filters} onFilterChange={setFilters} />

          {/* Table controls */}
          <div className="px-6 py-2 flex items-center justify-between bg-white border-b border-gray-200">
            <span className="text-sm text-gray-500">{sortedData.length} classes</span>
            <label className="flex items-center gap-2 text-sm cursor-pointer">
              <input type="checkbox" checked={showMonthly} onChange={(e) => setShowMonthly(e.target.checked)} className="rounded" />
              <span className="text-gray-600">Show monthly columns</span>
            </label>
          </div>

          {/* Table */}
          <div className="overflow-x-auto">
            <table className="w-full border-collapse min-w-[900px]">
              <thead className="bg-gray-50 sticky top-0 z-10">
                <tr>
                  <SortHeader label="Alert" column="alert" sortConfig={sortConfig} onSort={setSortConfig} />
                  <SortHeader label="Category" column="category" sortConfig={sortConfig} onSort={setSortConfig} />
                  <SortHeader label="Class #" column="classCode" sortConfig={sortConfig} onSort={setSortConfig} />
                  <th className="px-2 py-2 text-left text-xs font-medium text-gray-600 uppercase tracking-wider">Class Description</th>
                  <SortHeader label="Annual Forecast" column="annualForecast" sortConfig={sortConfig} onSort={setSortConfig} />
                  <SortHeader label="Annual OTB" column="annualOTB" sortConfig={sortConfig} onSort={setSortConfig} />
                  <th className="px-2 py-2 text-left text-xs font-medium text-gray-600 uppercase tracking-wider">On Order Cost</th>
                  <SortHeader label="OTB Util %" column="otbUtilizationPct" sortConfig={sortConfig} onSort={setSortConfig} />
                  {showMonthly && MONTHS.map(m => (
                    <th key={m} className="px-2 py-2 text-right text-xs font-medium text-gray-600 uppercase tracking-wider">{m}</th>
                  ))}
                  <th className="px-2 py-2 text-left text-xs font-medium text-gray-600 uppercase tracking-wider">Alert Months</th>
                </tr>
              </thead>
              <tbody>
                {sortedData.map((cls, idx) => (
                  <Fragment key={cls.id}>
                    <tr
                      className={`border-b border-gray-100 cursor-pointer transition-colors ${idx % 2 === 0 ? 'bg-white' : 'bg-gray-50/50'} hover:bg-blue-50/50 ${selectedRow === cls.id ? 'bg-blue-50' : ''}`}
                      onClick={() => setSelectedRow(selectedRow === cls.id ? null : cls.id)}
                    >
                      <td className="px-2 py-1.5 text-center">
                        {cls.alertLevel === 'critical' ? <span title="Critical">🔴</span> : cls.alertLevel === 'warning' ? <span title="Warning">🟡</span> : <span className="text-gray-300">&mdash;</span>}
                      </td>
                      <td className="px-2 py-1.5 text-gray-800">{cls.category}</td>
                      <td className="px-2 py-1.5 text-gray-800">{cls.classCode}</td>
                      <td className="px-2 py-1.5 text-gray-800 max-w-[180px] truncate" title={cls.classDesc}>{cls.classDesc}{cls.isAllZero && <span className="ml-2 text-xs text-gray-400 italic">No activity</span>}</td>
                      <td className="px-2 py-1.5 text-right font-mono text-gray-800">{formatCurrency(cls.annualForecast)}</td>
                      <td className={`px-2 py-1.5 text-right font-mono ${cls.annualOTB < 0 ? 'text-red-600 font-medium' : 'text-gray-800'}`}>{formatCurrency(cls.annualOTB)}</td>
                      <td className="px-2 py-1.5 text-right font-mono text-gray-800">{formatCurrency(cls.annualOnOrderCost)}</td>
                      <td className="px-2 py-1.5"><UtilizationBar pct={cls.otbUtilizationPct} /></td>
                      {showMonthly && MONTHS.map(m => (
                        <MonthCell
                          key={m}
                          value={cls.otbCostOriginal?.[m]}
                          variance={cls.monthlyVariance[m]}
                          settings={settings}
                          otbValue={cls.otbCostOriginal?.[m] || 0}
                        />
                      ))}
                      <td className="px-2 py-1.5 text-xs text-gray-600 max-w-[120px] truncate" title={cls.alertMonths.join(', ')}>
                        {cls.alertMonths.length > 0 ? cls.alertMonths.join(', ') : '\u2014'}
                      </td>
                    </tr>
                    {selectedRow === cls.id && <DetailDrawer cls={cls} />}
                  </Fragment>
                ))}
              </tbody>
            </table>
          </div>

          {sortedData.length === 0 && (
            <div className="text-center py-12 text-gray-400">No classes match the current filters.</div>
          )}
        </>
      )}
    </div>
  )
}
