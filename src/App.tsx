import { useState, useMemo, useCallback, useRef, useEffect } from 'react'
import './App.css'
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  BarChart,
  Bar,
  ScatterChart,
  Scatter,
  ZAxis,
  Cell,
  AreaChart,
  Area,
  RadarChart,
  Radar,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis,
  PieChart,
  Pie,
} from 'recharts'
import clsx from 'clsx'

// Types
export interface DataRow {
  key: string | number
  [series: string]: string | number
}

interface SeriesMeta {
  id: string
  label: string
  color: string
  visible: boolean
}

// Initial sample data (business financials in $M)
const initialRows: DataRow[] = [
  { key: 2019, Revenue: 120, COGS: 48, "Operating Expenses": 30, "R&D": 12, "Net Profit": 30 },
  { key: 2020, Revenue: 135, COGS: 53, "Operating Expenses": 33, "R&D": 14, "Net Profit": 35 },
  { key: 2021, Revenue: 155, COGS: 60, "Operating Expenses": 36, "R&D": 16, "Net Profit": 43 },
  { key: 2022, Revenue: 178, COGS: 68, "Operating Expenses": 40, "R&D": 18, "Net Profit": 52 },
  { key: 2023, Revenue: 205, COGS: 78, "Operating Expenses": 44, "R&D": 20, "Net Profit": 63 },
  { key: 2024, Revenue: 236, COGS: 90, "Operating Expenses": 49, "R&D": 22, "Net Profit": 75 },
  { key: 2025, Revenue: 264, COGS: 100, "Operating Expenses": 53, "R&D": 24, "Net Profit": 87 },
]

const palette = ['#ef4444', '#f59e0b', '#3b82f6', '#10b981', '#6366f1', '#ec4899', '#14b8a6']

function App() {
  const [rows, setRows] = useState<DataRow[]>(initialRows)
  const [series, setSeries] = useState<SeriesMeta[]>([
    { id: 'Revenue', label: 'Revenue', color: palette[0], visible: true },
    { id: 'COGS', label: 'COGS', color: palette[1], visible: true },
    { id: 'Operating Expenses', label: 'Operating Expenses', color: palette[2], visible: true },
    { id: 'R&D', label: 'R&D', color: palette[3], visible: true },
    { id: 'Net Profit', label: 'Net Profit', color: palette[4], visible: true },
  ])
  const [activeCell, setActiveCell] = useState<{ r: number; c: string } | null>(null)
  const [labelColumn, setLabelColumn] = useState<string>('key')
  const [mappingOpen, setMappingOpen] = useState(true)
  const [editorOpen, setEditorOpen] = useState(false)

  // New chart type state
  type ChartType = 'line' | 'bar' | 'stackedBar' | 'area' | 'radar' | 'pie' | 'donut' | 'histogram' | 'scatter'
  const [chartType, setChartType] = useState<ChartType>('line')
  // Scatter-specific X axis column (defaults to key)
  const [scatterXColumn, setScatterXColumn] = useState<string>('key')
  // Scatter-specific Y axis (defaults to first series)
  const [scatterYColumn, setScatterYColumn] = useState<string>('Revenue')
  // Optional size and color mapping columns
  const [scatterSizeColumn, setScatterSizeColumn] = useState<string>('')
  const [scatterColorColumn, setScatterColorColumn] = useState<string>('')
  // Bar specific color mapping
  const [barColorColumn, setBarColorColumn] = useState<string>('')
  // Pie / Donut specific selected value series
  const [pieSeriesColumn, setPieSeriesColumn] = useState<string>('Revenue')
  // Histogram specific column + bins
  const [histogramColumn, setHistogramColumn] = useState<string>('Revenue')
  const [histogramBins, setHistogramBins] = useState<number>(10)

  // Renaming state
  const [editingSeriesId, setEditingSeriesId] = useState<string | null>(null)
  const [editingValue, setEditingValue] = useState('')

  // Drag state
  const dragSeriesId = useRef<string | null>(null)
  const dragRowKey = useRef<string | number | null>(null)
  const [rowDragOverKey, setRowDragOverKey] = useState<string | number | null>(null)
  const [rowDragPosition, setRowDragPosition] = useState<'above' | 'below' | null>(null)
  // Row header renaming state
  const [editingRowIndex, setEditingRowIndex] = useState<number | null>(null)
  const [editingRowValue, setEditingRowValue] = useState<string>('')

  const toggleSeriesVisible = (id: string) => {
    setSeries(prev => prev.map(s => s.id === id ? { ...s, visible: !s.visible } : s))
  }

  const handleCellChange = useCallback(
    (rowIndex: number, field: string, value: string) => {
      setRows((prev) => {
        const next = [...prev]
        const parsed = value === '' ? '' : Number(value)
        next[rowIndex] = { ...next[rowIndex], [field]: isNaN(parsed as number) ? value : parsed }
        return next
      })
    },
    []
  )

  const removeRow = (index: number) => {
    setRows((prev) => prev.filter((_, i) => i !== index))
  }

  const startRenameSeries = (id: string) => {
    const target = series.find(s => s.id === id)
    if (!target) return
    setEditingSeriesId(id)
    setEditingValue(target.label)
  }

  const commitRenameSeries = () => {
    if (!editingSeriesId) return
    const trimmed = editingValue.trim()
    if (!trimmed) { setEditingSeriesId(null); return }
    // uniqueness (labels only)
    const dup = series.some(s => s.id !== editingSeriesId && s.label.toLowerCase() === trimmed.toLowerCase())
    if (dup) { setEditingSeriesId(null); return }
    setSeries(prev => prev.map(s => s.id === editingSeriesId ? { ...s, label: trimmed } : s))
    setEditingSeriesId(null)
  }

  const cancelRename = () => {
    setEditingSeriesId(null)
  }

  const onSeriesDragStart = (id: string, e: React.DragEvent) => {
    dragSeriesId.current = id
    e.dataTransfer.effectAllowed = 'move'
  }
  const onSeriesDragOver = (e: React.DragEvent) => {
    e.preventDefault()
  }
  const onSeriesDrop = (targetId: string) => {
    const sourceId = dragSeriesId.current
    dragSeriesId.current = null
    if (!sourceId || sourceId === targetId) return
    setSeries(prev => {
      const srcIdx = prev.findIndex(s => s.id === sourceId)
      const tgtIdx = prev.findIndex(s => s.id === targetId)
      if (srcIdx === -1 || tgtIdx === -1) return prev
      const clone = [...prev]
      const [moved] = clone.splice(srcIdx, 1)
      clone.splice(tgtIdx, 0, moved)
      return clone
    })
  }

  const onRowDragStart = (rowKey: string | number, e: React.DragEvent) => {
    // Prevent drag when currently editing a row header
    if (editingRowIndex !== null) { e.preventDefault(); return }
    dragRowKey.current = rowKey
    e.dataTransfer.effectAllowed = 'move'
  }
  const onRowDragOver = (e: React.DragEvent, targetKey: string | number) => {
    e.preventDefault()
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect()
    const midY = rect.top + rect.height / 2
    const pos: 'above' | 'below' = e.clientY < midY ? 'above' : 'below'
    if (rowDragOverKey !== targetKey || rowDragPosition !== pos) {
      setRowDragOverKey(targetKey)
      setRowDragPosition(pos)
    }
  }
  const clearRowDragState = () => {
    setRowDragOverKey(null)
    setRowDragPosition(null)
  }
  const onRowDrop = (targetKey: string | number) => {
    const sourceKey = dragRowKey.current
    dragRowKey.current = null
    if (sourceKey === null) { clearRowDragState(); return }
    setRows(prev => {
      const srcIdx = prev.findIndex(r => r.key === sourceKey)
      const tgtIdx = prev.findIndex(r => r.key === targetKey)
      if (srcIdx === -1 || tgtIdx === -1) return prev
      let insertIndex = tgtIdx + (rowDragPosition === 'below' ? 1 : 0)
      // If moving down and removing earlier element affects index
      if (srcIdx < insertIndex) insertIndex -= 1
      if (insertIndex === srcIdx) return prev
      const clone = [...prev]
      const [moved] = clone.splice(srcIdx, 1)
      clone.splice(insertIndex, 0, moved)
      return clone
    })
    clearRowDragState()
  }
  const onRowDragEnd = () => { clearRowDragState(); dragRowKey.current = null }

  // Row header renaming
  const startEditRowHeader = (rowIndex: number) => {
    setEditingRowIndex(rowIndex)
    setEditingRowValue(String(rows[rowIndex][labelColumn] ?? ''))
  }
  const cancelEditRowHeader = () => {
    setEditingRowIndex(null)
    setEditingRowValue('')
  }
  const commitEditRowHeader = () => {
    if (editingRowIndex === null) return
    const value = editingRowValue.trim()
    if (value === '') { cancelEditRowHeader(); return }
    setRows(prev => {
      const clone = [...prev]
      const row = { ...clone[editingRowIndex] }
      // If label column is 'key' retain numeric type if possible
      if (labelColumn === 'key') {
        const num = Number(value)
        // uniqueness check for key values
        const duplicate = prev.some((r, i) => i !== editingRowIndex && String(r[labelColumn]) === String(value))
        if (duplicate) { return prev } // ignore duplicate commit
        row[labelColumn] = !isNaN(num) && value.match(/^[-+]?\d+(\.\d+)?$/) ? num : value
      } else {
        row[labelColumn] = value
      }
      clone[editingRowIndex] = row
      return clone
    })
    cancelEditRowHeader()
  }

  const visibleSeries = series.filter((s) => s.visible)

  // Determine active X column based on chart type
  const currentXColumn = chartType === 'scatter' ? scatterXColumn : labelColumn
  // Determine if current X column values are numeric (for scatter axis typing)
  const isLabelNumeric = useMemo(
    () => rows.length > 0 && typeof rows[0][currentXColumn] === 'number',
    [rows, currentXColumn]
  )

  // All columns available for label selection (including series columns)
  const labelOptions = useMemo(() => {
    if (!rows.length) return ['key']
    return Object.keys(rows[0])
  }, [rows])

  // Helper render function for charts
  const renderChartContents = () => {
    const commonMargin = { left: 12, right: 12, top: 12, bottom: 12 }
    if (chartType === 'stackedBar') {
      // Color scale (reuse barColorColumn state)
      const barColorScale = (() => {
        if (!barColorColumn) return () => undefined as string | undefined
        const distinct: (string | number | undefined)[] = []
        rows.forEach(r => { const v = r[barColorColumn]; if (!distinct.includes(v)) distinct.push(v) })
        return (v: unknown) => {
          const value = v as string | number | undefined
          const idx = distinct.indexOf(value)
          return palette[idx % palette.length]
        }
      })()
      return (
        <BarChart data={rows} margin={commonMargin}>
          <CartesianGrid strokeDasharray="3 3" stroke="#444" />
          <XAxis dataKey={labelColumn} />
          <YAxis />
          <Tooltip />
          <Legend />
          {visibleSeries.map(s => (
            <Bar key={s.id} dataKey={s.id} name={s.label} fill={s.color} stackId="a">
              {barColorColumn && rows.map((r,i) => (
                <Cell key={i} fill={barColorScale(r[barColorColumn]) || s.color} />
              ))}
            </Bar>
          ))}
        </BarChart>
      )
    }
    if (chartType === 'area') {
      return (
        <AreaChart data={rows} margin={commonMargin}>
          <CartesianGrid strokeDasharray="3 3" stroke="#444" />
          <XAxis dataKey={labelColumn} />
            <YAxis />
            <Tooltip />
            <Legend />
            {visibleSeries.map(s => (
              <Area key={s.id} type="monotone" dataKey={s.id} name={s.label} stroke={s.color} fill={s.color + '55'} isAnimationActive={false} />
            ))}
        </AreaChart>
      )
    }
    if (chartType === 'radar') {
      return (
        <RadarChart data={rows} outerRadius={90} width={480} height={300}>
          <PolarGrid />
          <PolarAngleAxis dataKey={labelColumn} />
          <PolarRadiusAxis />
          <Tooltip />
          <Legend />
          {visibleSeries.map(s => (
            <Radar key={s.id} name={s.label} dataKey={s.id} stroke={s.color} fill={s.color} fillOpacity={0.4} />
          ))}
        </RadarChart>
      )
    }
    if (chartType === 'pie' || chartType === 'donut') {
      const pieData = rows.map(r => ({ label: r[labelColumn] as string | number, value: r[pieSeriesColumn] as number | string }))
      return (
        <PieChart width={480} height={300}>
          <Tooltip />
          <Legend />
          <Pie data={pieData} dataKey="value" nameKey="label" outerRadius={100} innerRadius={chartType === 'donut' ? 60 : 0} label>
            {pieData.map((d,i) => {
              const color = palette[i % palette.length]
              // Use d.label in key to ensure stability and mark d as used
              return <Cell key={String(d.label)+ '-' + i} fill={color} />
            })}
          </Pie>
        </PieChart>
      )
    }
    if (chartType === 'histogram') {
      const values = rows.map(r => r[histogramColumn]).filter(v => typeof v === 'number') as number[]
      if (!values.length) {
        return <div style={{ padding: 16, fontSize: '0.8rem' }}>No numeric data for {histogramColumn}</div>
      }
      const min = Math.min(...values)
      const max = Math.max(...values)
      const binCount = Math.max(1, histogramBins)
      const binSize = (max - min) / binCount || 1
      const bins: { bin: string; count: number }[] = Array.from({ length: binCount }).map(() => ({ bin: '', count: 0 }))
      values.forEach(v => {
        let idx = Math.floor((v - min) / binSize)
        if (idx >= binCount) idx = binCount - 1
        bins[idx].count += 1
      })
      bins.forEach((b,i) => {
        const start = min + i * binSize
        const end = start + binSize
        b.bin = `${Number(start.toFixed(2))}–${Number(end.toFixed(2))}`
      })
      return (
        <BarChart data={bins} margin={commonMargin}>
          <CartesianGrid strokeDasharray="3 3" stroke="#444" />
          <XAxis dataKey="bin" />
          <YAxis />
          <Tooltip />
          <Bar dataKey="count" fill={palette[0]} />
        </BarChart>
      )
    }
    if (chartType === 'bar') {
      // Prepare color scale if coloring by a column
      const barColorScale = (() => {
        if (!barColorColumn) return () => undefined as string | undefined
        const distinct: (string | number | undefined)[] = []
        rows.forEach(r => { const v = r[barColorColumn]; if (!distinct.includes(v)) distinct.push(v) })
        return (v: unknown) => {
          const value = v as string | number | undefined
            const idx = distinct.indexOf(value)
            return palette[idx % palette.length]
        }
      })()
      return (
        <BarChart data={rows} margin={commonMargin}>
          <CartesianGrid strokeDasharray="3 3" stroke="#444" />
          <XAxis dataKey={labelColumn} />
          <YAxis />
          <Tooltip />
          <Legend />
          {visibleSeries.map(s => (
            <Bar key={s.id} dataKey={s.id} name={s.label} fill={s.color}>
              {barColorColumn && rows.map((r,i) => (
                <Cell key={i} fill={barColorScale(r[barColorColumn]) || s.color} />
              ))}
            </Bar>
          ))}
        </BarChart>
      )
    }
    if (chartType === 'scatter') {
      // Build single dataset for selected Y column
      const scatterData = rows.map(r => {
        const x = r[currentXColumn]
        const y = r[scatterYColumn]
        const sizeRaw = scatterSizeColumn ? Number(r[scatterSizeColumn]) : undefined
        const sizeNum = sizeRaw !== undefined && !isNaN(sizeRaw) ? sizeRaw : undefined
        const colorVal = scatterColorColumn ? r[scatterColorColumn] : undefined
        return { x, y, z: sizeNum, colorVal }
      })
      // Color scale for categorical color column
      const colorScale = (() => {
        if (!scatterColorColumn) return () => palette[0]
        const distinct: (string | number | undefined)[] = []
        scatterData.forEach(d => { if (!distinct.includes(d.colorVal)) distinct.push(d.colorVal) })
        return (v: unknown) => {
          const value = v as string | number | undefined
          const idx = distinct.indexOf(value)
          return palette[idx % palette.length]
        }
      })()
      return (
        <ScatterChart margin={commonMargin}>
          <CartesianGrid strokeDasharray="3 3" stroke="#444" />
          <XAxis type={isLabelNumeric ? 'number' : 'category'} dataKey="x" name={currentXColumn === 'key' ? 'Year' : currentXColumn} />
          <YAxis dataKey="y" name={scatterYColumn} />
          {scatterSizeColumn && <ZAxis dataKey="z" range={[60, 400]} name={scatterSizeColumn} />}
          <Tooltip cursor={{ strokeDasharray: '3 3' }} />
          <Legend />
          <Scatter name={scatterYColumn} data={scatterData} fill={palette[0]}>
            {scatterData.map((d, i) => (
              <Cell key={i} fill={scatterColorColumn ? colorScale(d.colorVal) : palette[0]} />
            ))}
          </Scatter>
        </ScatterChart>
      )
    }
    // default line
    return (
      <LineChart data={rows} margin={commonMargin}>
        <CartesianGrid strokeDasharray="3 3" stroke="#444" />
        <XAxis dataKey={labelColumn} />
        <YAxis />
        <Tooltip />
        <Legend />
        {visibleSeries.map(s => (
          <Line key={s.id} type="monotone" dataKey={s.id} name={s.label} stroke={s.color} dot={false} strokeWidth={2} isAnimationActive={false} />
        ))}
      </LineChart>
    )
  }

  // Reusable mapping controls (without collapse wrapper) used inline + modal
  const MappingControls = () => (
    <>
      {chartType !== 'scatter' && (
        <div className="map-group">
          <div className="group-label">Labels</div>
          <LabelColumnSelect
            columns={labelOptions}
            rows={rows}
            value={labelColumn}
            onChange={setLabelColumn}
          />
        </div>
      )}
      {(chartType === 'pie' || chartType === 'donut') && (
        <div className="map-group">
          <div className="group-label">Value series</div>
          <LabelColumnSelect
            columns={series.map(s=>s.id)}
            rows={rows}
            value={pieSeriesColumn}
            onChange={setPieSeriesColumn}
          />
        </div>
      )}
      {chartType === 'histogram' && (
        <div className="map-group">
          <div className="group-label">Column</div>
          <LabelColumnSelect
            columns={series.map(s=>s.id)}
            rows={rows}
            value={histogramColumn}
            onChange={setHistogramColumn}
          />
          <div style={{ marginTop: 6 }}>
            <label style={{ fontSize: '0.7rem', display: 'flex', gap: 4, alignItems: 'center' }}>Bins:
              <input type="number" min={1} max={50} value={histogramBins} style={{ width: 60 }} onChange={e => setHistogramBins(Math.max(1, Math.min(50, Number(e.target.value)||10)))} />
            </label>
          </div>
        </div>
      )}
      {chartType === 'scatter' && (
        <>
          <div className="map-group">
            <div className="group-label">X axis</div>
            <LabelColumnSelect
              columns={labelOptions}
              rows={rows}
              value={scatterXColumn}
              onChange={setScatterXColumn}
            />
          </div>
          <div className="map-group">
            <div className="group-label">Y axis</div>
            <LabelColumnSelect
              columns={series.map(s => s.id)}
              rows={rows}
              value={scatterYColumn}
              onChange={setScatterYColumn}
            />
          </div>
          <div className="map-group">
            <div className="group-label">Size</div>
            <LabelColumnSelect
              columns={labelOptions}
              rows={rows}
              value={scatterSizeColumn}
              onChange={setScatterSizeColumn}
              allowNone
              placeholder="(None)"
            />
          </div>
          <div className="map-group">
            <div className="group-label">Color by</div>
            <LabelColumnSelect
              columns={labelOptions}
              rows={rows}
              value={scatterColorColumn}
              onChange={setScatterColorColumn}
              allowNone
              placeholder="(None)"
            />
          </div>
        </>
      )}
      {chartType !== 'scatter' && (
        <div className="map-group">
          {(chartType === 'line' || chartType === 'bar' || chartType === 'stackedBar' || chartType === 'area' || chartType === 'radar') && (
            <>
              <div className="group-label">{chartType === 'bar' || chartType === 'stackedBar' ? 'Bars' : chartType === 'area' ? 'Areas' : chartType === 'radar' ? 'Radars' : 'Lines'}</div>
              <MultiSelectSeries series={series} onToggle={toggleSeriesVisible} rows={rows} />
            </>
          )}
        </div>
      )}
      {(chartType === 'bar' || chartType === 'stackedBar') && (
        <div className="map-group">
          <div className="group-label">Color by</div>
          <LabelColumnSelect
            columns={labelOptions}
            rows={rows}
            value={barColorColumn}
            onChange={setBarColorColumn}
            allowNone
            placeholder="(None)"
          />
        </div>
      )}
      {chartType === 'scatter' && !series.some(s => s.id === scatterYColumn) && (
        <div style={{ color: 'orange', fontSize: '0.75rem' }}>Selected Y axis series removed.</div>
      )}
    </>
  )

  return (
    <div className="app-root">
      <div className="title-row" style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
        <h2 style={{ margin: 0 }}>Chart</h2>
        <label style={{ fontSize: '0.85rem', display: 'flex', alignItems: 'center', gap: 4 }}>
          Type:
          <select value={chartType} onChange={e => setChartType(e.target.value as ChartType)}>
            <option value="line">Line</option>
            <option value="bar">Bar</option>
            <option value="stackedBar">Stacked Bar</option>
            <option value="area">Area</option>
            <option value="radar">Radar</option>
            <option value="pie">Pie</option>
            <option value="donut">Donut</option>
            <option value="histogram">Histogram</option>
            <option value="scatter">Scatter</option>
          </select>
        </label>
      </div>
      {/* Fullscreen chart section with inline mapping controls */}
      <div className="chart-fullscreen-wrapper">
        <div className="chart-inline-wrapper" style={{ display: 'flex', gap: '1rem', alignItems: 'stretch' }}>
          <div className="chart-fullscreen" style={{ flex: 1, minHeight: 360, position: 'relative' }}>
            {/**
             * Use an explicit pixel height so Recharts can calculate size.
             * Previously height="100%" inside a flex item with auto height caused 0px chart height.
             */}
            <ResponsiveContainer width="100%" height={360}>
              {renderChartContents()}
            </ResponsiveContainer>
            <button className="open-editor-btn" onClick={() => setEditorOpen(true)} style={{ position: 'absolute', top: 8, right: 8, height: 40 }}>Open Editor</button>
          </div>
          <div className="inline-mapping-panel" style={{ width: 280, overflowY: 'auto', border: '1px solid #333', borderRadius: 6, padding: '0.5rem 0.75rem' }}>
            <div className="group-label" style={{ fontWeight: 600, fontSize: '0.8rem', opacity: 0.9, marginBottom: 4 }}>Mapping</div>
            <MappingControls />
          </div>
        </div>
      </div>

      {/* Modal for editor */}
      {editorOpen && (
        <Modal onClose={() => setEditorOpen(false)}>
          <div className="layout">
            <div className="chart-panel">
              <div className="mapping-panel">
                <button className="collapse-btn" onClick={() => setMappingOpen(o => !o)}>
                  <span className="icon">⚙️</span> Chart setup {mappingOpen ? '▾' : '▸'}
                </button>
                {mappingOpen && (
                  <div className="mapping-body">
                    <MappingControls />
                  </div>
                )}
              </div>
              <ResponsiveContainer width="100%" height={300}>
                {renderChartContents()}
              </ResponsiveContainer>
            </div>
            <div className="table-panel">
              <div className="data-table-wrapper">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Year</th>
                      {series.map(s => (
                        <th
                          key={s.id}
                          className={clsx('series-header', { renaming: editingSeriesId === s.id })}
                          draggable
                          onDragStart={(e) => onSeriesDragStart(s.id, e)}
                          onDragOver={onSeriesDragOver}
                          onDrop={() => onSeriesDrop(s.id)}
                          title="Drag to reorder. Double-click or use ✎ to rename"
                        >
                          {editingSeriesId === s.id ? (
                            <input
                              className="th-edit-input"
                              value={editingValue}
                              autoFocus
                              onChange={e => setEditingValue(e.target.value)}
                              onBlur={commitRenameSeries}
                              onKeyDown={e => {
                                if (e.key === 'Enter') { e.preventDefault(); commitRenameSeries() }
                                if (e.key === 'Escape') { e.preventDefault(); cancelRename() }
                              }}
                            />
                          ) : (
                            <div className="series-header-inner" onDoubleClick={() => startRenameSeries(s.id)}>
                              <span className="series-label-text">{s.label}</span>
                              <button
                                className="rename-btn"
                                onClick={(e) => { e.stopPropagation(); startRenameSeries(s.id) }}
                                aria-label="Rename series"
                              >✎</button>
                              <span className="drag-handle" aria-hidden>⋮⋮</span>
                            </div>
                          )}
                        </th>
                      ))}
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((r, ri) => (
                      <tr
                        key={r.key}
                        onDragOver={(e) => onRowDragOver(e, r.key)}
                        onDrop={() => onRowDrop(r.key)}
                        onDragEnd={onRowDragEnd}
                        className={clsx({ 'row-drag-over-above': rowDragOverKey === r.key && rowDragPosition === 'above', 'row-drag-over-below': rowDragOverKey === r.key && rowDragPosition === 'below' })}
                      >
                        <th
                          className={clsx('row-header', { dragging: dragRowKey.current === r.key })}
                          draggable
                          onDragStart={(e) => onRowDragStart(r.key, e)}
                          title={editingRowIndex === ri ? '' : 'Drag to reorder row'}
                          onDoubleClick={() => startEditRowHeader(ri)}
                        >
                          <span className="row-drag-handle" aria-hidden>⋮</span>
                          {editingRowIndex === ri ? (
                            <input
                              className="th-edit-input"
                              value={editingRowValue}
                              autoFocus
                              onChange={e => setEditingRowValue(e.target.value)}
                              onBlur={commitEditRowHeader}
                              onKeyDown={e => {
                                if (e.key === 'Enter') { e.preventDefault(); commitEditRowHeader() }
                                if (e.key === 'Escape') { e.preventDefault(); cancelEditRowHeader() }
                              }}
                              style={{ width: '5rem' }}
                            />
                          ) : (
                            <span className="row-label-text">
                              {String(r['key'])}
                              <button
                                className="rename-btn"
                                onClick={(e) => { e.stopPropagation(); startEditRowHeader(ri) }}
                                aria-label="Rename row header"
                                title="Rename row header"
                                style={{ marginLeft: 4 }}
                              >✎</button>
                            </span>
                          )}
                        </th>
                        {series.map(s => {
                          const val = r[s.id] as number | string | undefined
                          const active = activeCell?.r === ri && activeCell?.c === s.id
                          return (
                            <td key={s.id} className={clsx({ active })} onClick={() => setActiveCell({ r: ri, c: s.id })}>
                              <input value={val ?? ''} onChange={(e) => handleCellChange(ri, s.id, e.target.value)} onFocus={() => setActiveCell({ r: ri, c: s.id })} />
                            </td>
                          )
                        })}
                        <td><button className="danger" onClick={() => removeRow(ri)}>×</button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
          <footer className="footer">Data is in-memory only.</footer>
        </Modal>
      )}
    </div>
  )
}

// Multi-select component
interface MultiSelectProps { series: SeriesMeta[]; onToggle: (id: string) => void; rows: DataRow[] }
function MultiSelectSeries({ series, onToggle, rows }: MultiSelectProps) {
  const [open, setOpen] = useState(false)
  const selected = series.filter(s => s.visible)
  const previewFor = (id: string) => {
    const vals = rows.slice(0, 4).map(r => r[id]).filter(v => v !== undefined)
    if (!vals.length) return ''
    const more = rows.length > 4 ? '…' : ''
    return `${vals.join(', ')}${more}`
  }
  return (
    <div className="multi-select" tabIndex={0} onBlur={() => setOpen(false)}>
      <div className="multi-select-display" onClick={() => setOpen(o => !o)}>
        {selected.length === 0 && <span className="placeholder">Select series</span>}
        {selected.map(s => (
          <span key={s.id} className="chip" onClick={(e) => { e.stopPropagation(); onToggle(s.id) }}>
            <span className="swatch" style={{ background: s.color }} /> {s.label}
          </span>
        ))}
        <span className="caret">{open ? '▴' : '▾'}</span>
      </div>
      {open && (
        <div className="multi-select-menu">
          {series.map(s => (
            <div key={s.id} className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => onToggle(s.id)}>
              <input type="checkbox" readOnly checked={s.visible} />
              <span className="swatch" style={{ background: s.color }} /> {s.label}
              <div className="menu-preview" style={{ fontSize: '0.7rem', opacity: 0.75, marginLeft: 26 }}>{previewFor(s.id)}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

// Label selector with previews
interface LabelSelectProps { columns: string[]; rows: DataRow[]; value: string; onChange: (col: string) => void; placeholder?: string; allowNone?: boolean }
function LabelColumnSelect({ columns, rows, value, onChange, placeholder = 'Select', allowNone = false }: LabelSelectProps) {
  const [open, setOpen] = useState(false)
  const [query, setQuery] = useState('')
  const ref = useRef<HTMLDivElement | null>(null)
  useEffect(() => {
    if (!open) return
    const handler = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false)
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [open])
  const baseCols = columns
  // Removed standalone allCols to satisfy exhaustive-deps; derive inside memo
  const filtered = useMemo(() => {
    const allCols = allowNone ? ['__NONE__', ...baseCols] : baseCols
    if (!query.trim()) return allCols
    const q = query.toLowerCase()
    return allCols.filter(c => c === '__NONE__' ? '(none)'.includes(q) : c.toLowerCase().includes(q))
  }, [allowNone, baseCols, query])
  const previewFor = (col: string) => {
    if (col === '__NONE__') return ''
    const vals = rows.slice(0, 4).map(r => r[col])
    const more = rows.length > 4 ? '…' : ''
    return `${vals.join(', ')}${more}`
  }
  const displayValue = () => {
    if (allowNone && (value === '' || value === '__NONE__')) return placeholder
    if (value === 'key') return 'Year'
    return value
  }
  const commit = (col: string) => {
    if (col === '__NONE__') {
      onChange('')
    } else {
      onChange(col)
    }
    setOpen(false)
    setQuery('')
  }
  const isSelected = (col: string) => {
    if (col === '__NONE__') return allowNone && (value === '' || value === '__NONE__')
    return col === value
  }
  return (
    <div className="single-select" ref={ref}>
      <div
        className="single-select-display"
        role="button"
        tabIndex={0}
        onClick={() => setOpen(o => !o)}
        onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setOpen(o => !o) } }}
      >
        <span>{displayValue()}</span>
        <span className="caret">{open ? '▴' : '▾'}</span>
      </div>
      {open && (
        <div className="single-select-menu" onMouseDown={(e) => e.preventDefault()}>
          <div className="ss-search-wrapper"><input autoFocus placeholder="Search" value={query} onChange={(e) => setQuery(e.target.value)} /></div>
          <div className="ss-options">
            {filtered.map(c => {
              const selected = isSelected(c)
              const label = c === '__NONE__' ? placeholder : (c === 'key' ? 'Year' : c)
              return (
                <div key={c} className={clsx('ss-option', { selected })} onClick={() => commit(c)}>
                  <div className="ss-opt-line1"><span className="col-id">{label}</span>{selected && <span className="check">✓</span>}</div>
                  <div className="ss-opt-preview">{previewFor(c)}</div>
                </div>
              )
            })}
            {filtered.length === 0 && <div className="ss-empty">No matches</div>}
          </div>
        </div>
      )}
    </div>
  )
}

// Modal component
interface ModalProps { children: React.ReactNode; onClose: () => void }
function Modal({ children, onClose }: ModalProps) {
  useEffect(() => {
    const handler = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose])
  return (
    <div className="modal-overlay" role="dialog" aria-modal="true">
      <div className="modal">
        <button className="modal-close" onClick={onClose} aria-label="Close editor">×</button>
        {children}
      </div>
    </div>
  )
}

export default App
