import { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'
import {
  PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer,
} from 'recharts'
import { supabase } from './supabaseClient'

const BASE_CATEGORIES = ['Personal', 'ARES', 'Noble', 'Miscellaneous']
const CUSTOM = 'Custom…'

// Category tones are data-driven, so they stay as inline colours (not utilities).
const CAT_COLORS = {
  Personal: '#2F6F5E',       // pine
  ARES: '#A85433',           // clay
  Noble: '#5A4B8C',          // amethyst
  Miscellaneous: '#96741C',  // antique gold
}
const PALETTE = ['#2C6E7F', '#8A4B6B', '#4F7A3A', '#7A5C3A', '#3F5AA6', '#9A5B2E']

const inr = new Intl.NumberFormat('en-IN', {
  style: 'currency', currency: 'INR', maximumFractionDigits: 2,
})
const today = () => new Date().toISOString().slice(0, 10)

const fmtDay = (d) => {
  const dt = new Date(d)
  if (isNaN(dt)) return d
  return dt.toLocaleDateString('en-IN', { day: '2-digit', month: 'short' })
}

const compactInr = (n) => {
  const a = Math.abs(n)
  if (a >= 1e7) return '₹' + (n / 1e7).toFixed(a % 1e7 ? 1 : 0) + 'Cr'
  if (a >= 1e5) return '₹' + (n / 1e5).toFixed(a % 1e5 ? 1 : 0) + 'L'
  if (a >= 1e3) return '₹' + (n / 1e3).toFixed(a % 1e3 ? 1 : 0) + 'k'
  return inr.format(n)
}

// ---- shared class strings ----
const cardCls = 'bg-surface border border-line rounded-[18px] shadow-card'
const labelCls = 'block mb-1.5 text-[11px] font-semibold tracking-[0.06em] uppercase text-muted'
const inputCls = 'w-full px-3 py-2.5 text-sm text-ink bg-surface border border-line rounded-[10px] placeholder:text-faint focus:outline-none focus:border-accent focus:ring-2 focus:ring-accent/10 transition'
const btnPrimary = 'inline-flex items-center gap-2 px-[18px] py-2.5 text-sm font-semibold rounded-[10px] text-white bg-accent shadow-stat cursor-pointer transition active:translate-y-px hover:bg-accent-dark disabled:opacity-60 disabled:cursor-default disabled:hover:bg-accent focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/20'
const btnGhost = 'inline-flex items-center gap-2 px-[18px] py-2.5 text-sm font-semibold rounded-[10px] text-ink-soft bg-surface border border-line cursor-pointer transition hover:bg-[#F1EFE7] hover:text-ink disabled:opacity-50 disabled:cursor-default focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/20'
const pillBase = 'text-[13px] font-medium px-3 py-1.5 rounded-lg border bg-surface cursor-pointer transition'
const pillSettle = pillBase + ' text-accent border-accent/30 hover:bg-accent/10'
const pillDelete = pillBase + ' text-ink-soft border-line hover:text-danger hover:border-danger/40 hover:bg-danger/[0.08]'

const CHEVRON =
  "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%238B918A' stroke-width='2.2' stroke-linecap='round' stroke-linejoin='round'><path d='M6 9l6 6 6-6'/></svg>"

// ---- module-scope components (OUTSIDE App, so inputs never remount / lose focus) ----

function Icon({ name, size = 16, strokeWidth = 1.8 }) {
  const p = {
    width: size, height: size, viewBox: '0 0 24 24', fill: 'none',
    stroke: 'currentColor', strokeWidth, strokeLinecap: 'round', strokeLinejoin: 'round',
  }
  switch (name) {
    case 'list': return (
      <svg {...p}><path d="M8 6h13M8 12h13M8 18h13" />
        <circle cx="3.5" cy="6" r="1.1" /><circle cx="3.5" cy="12" r="1.1" /><circle cx="3.5" cy="18" r="1.1" /></svg>
    )
    case 'check': return <svg {...p}><path d="M20 6 9 17l-5-5" /></svg>
    case 'download': return <svg {...p}><path d="M12 3v12m0 0 4-4m-4 4-4-4M4 21h16" /></svg>
    case 'plus': return <svg {...p}><path d="M12 5v14M5 12h14" /></svg>
    case 'receipt': return (
      <svg {...p}><path d="M5 3v18l2-1 2 1 2-1 2 1 2-1 2 1V3l-2 1-2-1-2 1-2-1-2 1-2-1Z" />
        <path d="M9 8.5h6M9 12h6" /></svg>
    )
    default: return null
  }
}

function NavButton({ active, onClick, icon, label, count }) {
  return (
    <button onClick={onClick}
      className={'flex-1 lg:flex-none flex items-center gap-2.5 text-left px-[13px] py-[11px] text-sm font-semibold rounded-xl border cursor-pointer transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/20 '
        + (active ? 'bg-ink text-white border-ink shadow-stat' : 'bg-transparent text-ink-soft border-transparent hover:bg-[#EBE9E0]')}>
      <Icon name={icon} />
      <span>{label}</span>
      <span className={'ml-auto text-[12px] font-semibold px-2 py-px rounded-full tabular-nums '
        + (active ? 'bg-white/15 text-white' : 'bg-accent/10 text-accent')}>{count}</span>
    </button>
  )
}

function StatCard({ primary, label, value, color }) {
  return (
    <div className={'border rounded-[14px] shadow-stat px-4 pt-4 pb-[18px] '
      + (primary ? 'bg-[linear-gradient(165deg,#13624F,#0D493A)] border-[#0D493A]' : 'bg-surface border-line')}>
      <div className={'flex items-center gap-[7px] mb-2.5 text-[11px] font-semibold tracking-[0.10em] uppercase '
        + (primary ? 'text-white/70' : 'text-muted')}>
        {color && <span className="w-2 h-2 rounded-full shrink-0" style={{ background: color }} />}
        {label}
      </div>
      <div className={'text-[22px] font-bold tracking-[-0.01em] tabular-nums ' + (primary ? 'text-white' : 'text-ink')}>{value}</div>
    </div>
  )
}

function ChartTooltip({ active, payload, label, kind }) {
  if (!active || !payload || !payload.length) return null
  const item = payload[0]
  return (
    <div className="bg-white border border-line rounded-[10px] shadow-card px-[11px] py-2">
      <div className="text-[11px] font-semibold tracking-[0.10em] uppercase text-muted mb-[3px]">
        {kind === 'date' ? fmtDay(label) : item.name}
      </div>
      <div className="text-sm font-bold text-ink tabular-nums">{inr.format(item.value)}</div>
    </div>
  )
}

function Empty({ title, sub }) {
  return (
    <div className="py-[54px] px-6 text-center">
      <div className="text-[15px] font-semibold text-ink">{title}</div>
      <div className="text-[13px] text-muted mt-[5px]">{sub}</div>
    </div>
  )
}

function ChartEmpty() {
  return <div className="h-full flex items-center justify-center text-faint text-[13px]">No data to chart yet</div>
}

export default function App() {
  const [expenses, setExpenses] = useState([])
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')

  const [view, setView] = useState('current')        // 'current' | 'settled'
  const [catFilter, setCatFilter] = useState('All')
  const [fromDate, setFromDate] = useState('')
  const [toDate, setToDate] = useState('')
  const [search, setSearch] = useState('')

  const [spentOn, setSpentOn] = useState(today())
  const [vendor, setVendor] = useState('')            // stored in DB column `vendor`, shown as "Bill description"
  const [category, setCategory] = useState(BASE_CATEGORIES[0])
  const [customCategory, setCustomCategory] = useState('')
  const [amount, setAmount] = useState('')
  const [notes, setNotes] = useState('')
  const [file, setFile] = useState(null)

  useEffect(() => { loadExpenses() }, [])

  async function loadExpenses() {
    setLoading(true)
    const { data, error } = await supabase
      .from('expenses').select('*').order('spent_on', { ascending: false })
    if (error) setError(error.message)
    else setExpenses(data)
    setLoading(false)
  }

  // Only accept a valid decimal string — blocks letters, e, +, -, and extra dots.
  function onAmountChange(e) {
    const v = e.target.value
    if (v === '' || /^\d*\.?\d*$/.test(v)) setAmount(v)
  }

  async function addExpense() {
    if (!amount || Number(amount) <= 0) { setError('Enter an amount greater than zero.'); return }
    const finalCategory = category === CUSTOM ? customCategory.trim() : category
    if (category === CUSTOM && !finalCategory) { setError('Enter a name for the custom category.'); return }
    setSaving(true); setError('')
    try {
      let receipt_url = null
      if (file) {
        const path = `${Date.now()}-${file.name.replace(/[^a-zA-Z0-9.-]/g, '_')}`
        const { error: upErr } = await supabase.storage.from('receipts').upload(path, file)
        if (upErr) throw upErr
        receipt_url = supabase.storage.from('receipts').getPublicUrl(path).data.publicUrl
      }
      const { error: insErr } = await supabase.from('expenses').insert({
        spent_on: spentOn, vendor: vendor || null, category: finalCategory,
        amount: Number(amount), notes: notes || null, receipt_url, settled: false,
      })
      if (insErr) throw insErr
      setVendor(''); setAmount(''); setNotes(''); setFile(null)
      setSpentOn(today()); setCategory(BASE_CATEGORIES[0]); setCustomCategory('')
      const el = document.getElementById('receipt-input'); if (el) el.value = ''
      await loadExpenses()
    } catch (e) {
      setError(e.message || 'Could not save the expense.')
    } finally { setSaving(false) }
  }

  async function setSettled(id, value) {
    const prev = expenses
    setExpenses(expenses.map(e => e.id === id ? { ...e, settled: value } : e))
    const { error } = await supabase.from('expenses').update({ settled: value }).eq('id', id)
    if (error) { setError(error.message); setExpenses(prev) }
  }

  async function deleteExpense(id) {
    const prev = expenses
    setExpenses(expenses.filter(e => e.id !== id))
    const { error } = await supabase.from('expenses').delete().eq('id', id)
    if (error) { setError(error.message); setExpenses(prev) }
  }

  // ---- derived data ----
  const scoped = expenses.filter(e => (view === 'settled' ? e.settled : !e.settled))
  const currentCount = expenses.filter(e => !e.settled).length
  const settledCount = expenses.filter(e => e.settled).length

  const allCats = []
  BASE_CATEGORIES.forEach(c => allCats.push(c))
  expenses.forEach(e => { if (e.category && !allCats.includes(e.category)) allCats.push(e.category) })
  const colorFor = (cat) => CAT_COLORS[cat] || PALETTE[Math.max(0, allCats.indexOf(cat)) % PALETTE.length]

  const catTotal = (cat) => scoped.filter(e => e.category === cat)
    .reduce((s, e) => s + Number(e.amount || 0), 0)
  const total = scoped.reduce((s, e) => s + Number(e.amount || 0), 0)

  const pieData = allCats.map(c => ({ name: c, value: catTotal(c) })).filter(d => d.value > 0)

  const byDate = {}
  scoped.forEach(e => { byDate[e.spent_on] = (byDate[e.spent_on] || 0) + Number(e.amount || 0) })
  const timeData = Object.entries(byDate).map(([date, value]) => ({ date, value }))
    .sort((a, b) => a.date.localeCompare(b.date))

  const listed = scoped.filter(e => {
    if (catFilter !== 'All' && e.category !== catFilter) return false
    if (fromDate && e.spent_on < fromDate) return false
    if (toDate && e.spent_on > toDate) return false
    if (search) {
      const hay = ((e.notes || '') + ' ' + (e.vendor || '')).toLowerCase()
      if (!hay.includes(search.toLowerCase())) return false
    }
    return true
  })

  function exportToExcel() {
    const rows = listed.map(e => ({
      Date: e.spent_on, 'Bill Description': e.vendor || '', Category: e.category || '',
      'Amount (INR)': e.amount, Notes: e.notes || '',
      Status: e.settled ? 'Settled' : 'Current', Receipt: e.receipt_url || '',
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Expenses')
    XLSX.writeFile(wb, `expenses-${view}-${today()}.xlsx`)
  }

  return (
    <div className="min-h-screen bg-paper text-ink font-sans antialiased">
      <div className="h-[3px] bg-[linear-gradient(90deg,#0F5A47,#2F6F5E_60%,#A85433)]" />

      {/* Header — the serif total is the single "signature" moment */}
      <header className="bg-surface border-b border-line px-4 sm:px-[30px] py-[22px] flex items-center justify-between gap-6">
        <div>
          <h1 className="text-[19px] font-bold tracking-[-0.01em] m-0">Expense Tracker</h1>
          <div className="text-[12.5px] text-muted mt-[3px]">
            {view === 'settled' ? 'Settled records' : 'Current, unsettled spending'}
            {catFilter !== 'All' ? ' · ' + catFilter : ''}
          </div>
        </div>
        <div className="text-right">
          <div className="text-[11px] font-semibold tracking-[0.10em] uppercase text-muted">
            {view === 'settled' ? 'Settled total' : 'Current outstanding'}
          </div>
          <div className="font-serif text-[28px] sm:text-[34px] leading-none tracking-[-0.01em] text-accent tabular-nums mt-[7px]">
            {inr.format(total)}
          </div>
        </div>
      </header>

      <div className="grid grid-cols-1 lg:grid-cols-[220px_1fr] gap-[26px] max-w-[1240px] mx-auto p-4 lg:p-[26px]">
        {/* Sidebar */}
        <aside>
          <div className="flex gap-2 lg:flex-col">
            <NavButton active={view === 'current'} onClick={() => setView('current')} icon="list" label="Current" count={currentCount} />
            <NavButton active={view === 'settled'} onClick={() => setView('settled')} icon="check" label="Settled" count={settledCount} />
          </div>
        </aside>

        {/* Main */}
        <main>
          {/* Category filter tabs */}
          <div className="flex gap-5 flex-wrap items-center border-b border-line mb-[22px]">
            {['All', ...allCats].map(c => (
              <button key={c} onClick={() => setCatFilter(c)}
                className={'bg-transparent border-none cursor-pointer pt-0.5 pb-3 -mb-px text-sm flex items-center gap-[7px] border-b-2 transition '
                  + (catFilter === c ? 'text-ink font-semibold border-accent' : 'text-muted font-medium border-transparent hover:text-ink')}>
                {c !== 'All' && <span className="w-2 h-2 rounded-full" style={{ background: colorFor(c) }} />}
                {c}
              </button>
            ))}
          </div>

          {/* Summary cards */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3.5 mb-[22px] animate-rise">
            <StatCard primary label={view === 'settled' ? 'Settled total' : 'Total spent'} value={inr.format(total)} />
            {allCats.map(c => <StatCard key={c} label={c} color={colorFor(c)} value={inr.format(catTotal(c))} />)}
          </div>

          {/* Charts */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-[18px] mb-[22px]">
            <div className={cardCls + ' p-5'}>
              <div className="text-sm font-semibold tracking-[-0.005em]">Spend by category</div>
              <div className="text-[12px] text-muted mt-0.5">Share of {view} spending</div>
              {pieData.length ? (
                <div className="flex gap-[18px] flex-wrap items-center mt-1.5">
                  <div className="relative w-[190px] h-[200px] shrink-0">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie data={pieData} dataKey="value" nameKey="name"
                          innerRadius={62} outerRadius={92} paddingAngle={2} cornerRadius={4} stroke="none">
                          {pieData.map(d => <Cell key={d.name} fill={colorFor(d.name)} />)}
                        </Pie>
                        <Tooltip content={<ChartTooltip />} />
                      </PieChart>
                    </ResponsiveContainer>
                    <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                      <div className="text-[11px] font-semibold tracking-[0.10em] uppercase text-muted mb-[3px]">Total</div>
                      <div className="font-serif text-[20px] text-ink tabular-nums">{compactInr(total)}</div>
                    </div>
                  </div>
                  <ul className="flex-1 min-w-[150px] list-none m-0 p-0 flex flex-col gap-2.5">
                    {pieData.map(d => (
                      <li key={d.name} className="flex items-center gap-2.5 text-[13px] text-ink-soft">
                        <span className="w-2 h-2 rounded-full shrink-0" style={{ background: colorFor(d.name) }} />
                        <span>{d.name}</span>
                        <span className="ml-auto font-semibold text-ink tabular-nums">{inr.format(d.value)}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              ) : <div className="h-[200px]"><ChartEmpty /></div>}
            </div>

            <div className={cardCls + ' p-5'}>
              <div className="text-sm font-semibold tracking-[-0.005em]">Spend over time</div>
              <div className="text-[12px] text-muted mt-0.5">Daily totals</div>
              <div className="h-[210px] mt-1.5">
                {timeData.length ? (
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={timeData} margin={{ top: 10, right: 6, left: -6, bottom: 0 }}>
                      <defs>
                        <linearGradient id="barFill" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#0F5A47" stopOpacity="0.92" />
                          <stop offset="100%" stopColor="#2F6F5E" stopOpacity="0.5" />
                        </linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="2 5" stroke="#ECEAE1" vertical={false} />
                      <XAxis dataKey="date" tickFormatter={fmtDay} tickLine={false}
                        axisLine={{ stroke: '#E6E3D9' }} tick={{ fontSize: 11, fill: '#8B918A' }} />
                      <YAxis tickFormatter={(v) => compactInr(v).replace('₹', '')} tickLine={false}
                        axisLine={false} width={42} tick={{ fontSize: 11, fill: '#8B918A' }} />
                      <Tooltip cursor={{ fill: 'rgba(15,90,71,.06)' }} content={<ChartTooltip kind="date" />} />
                      <Bar dataKey="value" fill="url(#barFill)" radius={[6, 6, 0, 0]} maxBarSize={46} />
                    </BarChart>
                  </ResponsiveContainer>
                ) : <ChartEmpty />}
              </div>
            </div>
          </div>

          {/* Add expense */}
          <div className={cardCls + ' p-[22px] mb-[22px] animate-rise'}>
            <div className="text-sm font-semibold tracking-[-0.005em] mb-4">Add expense</div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-3.5 items-end">
              <div className="lg:col-span-2">
                <label className={labelCls}>Date</label>
                <input className={inputCls} type="date" value={spentOn} onChange={e => setSpentOn(e.target.value)} />
              </div>
              <div className="lg:col-span-2">
                <label className={labelCls}>Bill description</label>
                <input className={inputCls} placeholder="What was this for?"
                  value={vendor} onChange={e => setVendor(e.target.value)} />
              </div>
              <div className="lg:col-span-2">
                <label className={labelCls}>Category</label>
                <select className={inputCls + ' appearance-none pr-9 cursor-pointer bg-no-repeat'}
                  style={{ backgroundImage: `url("${CHEVRON}")`, backgroundPosition: 'right 12px center' }}
                  value={category} onChange={e => setCategory(e.target.value)}>
                  {BASE_CATEGORIES.map(c => <option key={c}>{c}</option>)}
                  <option value={CUSTOM}>{CUSTOM}</option>
                </select>
                {category === CUSTOM && (
                  <input className={inputCls + ' mt-2'} placeholder="Custom category name"
                    value={customCategory} onChange={e => setCustomCategory(e.target.value)} />
                )}
              </div>
              <div className="lg:col-span-2">
                <label className={labelCls}>Amount</label>
                <div className="relative">
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted text-sm pointer-events-none">₹</span>
                  <input className={inputCls + ' !pl-[26px] tabular-nums'} inputMode="decimal"
                    placeholder="0.00" value={amount} onChange={onAmountChange} />
                </div>
              </div>
              <div className="lg:col-span-2">
                <label className={labelCls}>Notes</label>
                <input className={inputCls} placeholder="Optional" value={notes} onChange={e => setNotes(e.target.value)} />
              </div>
              <div className="lg:col-span-2">
                <label className={labelCls}>Receipt</label>
                <input id="receipt-input" className={inputCls + ' !py-[7px] !text-[13px]'}
                  type="file" accept="image/*" onChange={e => setFile(e.target.files[0] || null)} />
              </div>
            </div>
            <div className="flex items-center gap-3.5 mt-[18px]">
              <button className={btnPrimary} onClick={addExpense} disabled={saving}>
                {saving
                  ? <span className="w-3.5 h-3.5 rounded-full border-2 border-white/40 border-t-white animate-spin inline-block" />
                  : <Icon name="plus" size={16} strokeWidth={2.2} />}
                {saving ? 'Saving…' : 'Add expense'}
              </button>
              {error && <span className="text-[13px] text-danger">{error}</span>}
            </div>
          </div>

          {/* List + toolbar */}
          <div className={cardCls + ' overflow-hidden animate-rise'}>
            <div className="flex gap-3 flex-wrap items-end p-[18px] border-b border-line">
              <div>
                <label className={labelCls}>From</label>
                <input className={inputCls + ' w-[150px]'} type="date" value={fromDate} onChange={e => setFromDate(e.target.value)} />
              </div>
              <div>
                <label className={labelCls}>To</label>
                <input className={inputCls + ' w-[150px]'} type="date" value={toDate} onChange={e => setToDate(e.target.value)} />
              </div>
              <div className="flex-1 min-w-[160px]">
                <label className={labelCls}>Search description or notes</label>
                <input className={inputCls} placeholder="Search…" value={search} onChange={e => setSearch(e.target.value)} />
              </div>
              <button className={btnGhost} onClick={exportToExcel} disabled={!listed.length}>
                <Icon name="download" size={15} strokeWidth={2} /> Export to Excel
              </button>
            </div>

            {loading ? (
              <div className="py-[54px] text-center text-muted">Loading…</div>
            ) : listed.length === 0 ? (
              <Empty
                title={view === 'current' ? 'No current expenses' : 'Nothing settled yet'}
                sub={view === 'current' ? 'Add one above and it will show up here.' : 'Expenses you settle will collect here.'}
              />
            ) : (
              <table className="w-full border-collapse [&_tbody_tr:last-child_td]:border-b-0">
                <thead>
                  <tr>
                    {['Date', 'Bill description', 'Category', 'Amount', 'Receipt', ''].map((h, i) => (
                      <th key={i} className={'text-[11px] font-semibold tracking-[0.06em] uppercase text-muted px-4 py-3 bg-[#FAF9F5] border-b border-line '
                        + (i === 3 ? 'text-right' : 'text-left')}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {listed.map(e => (
                    <tr key={e.id} className="transition hover:bg-[#FAF9F4]">
                      <td className="px-4 py-3.5 text-sm border-b border-line-soft text-muted whitespace-nowrap">{fmtDay(e.spent_on)}</td>
                      <td className="px-4 py-3.5 text-sm border-b border-line-soft text-ink font-medium">{e.vendor || '—'}</td>
                      <td className="px-4 py-3.5 text-sm border-b border-line-soft">
                        <span className="inline-flex items-center gap-1.5 text-[12px] font-semibold px-2.5 py-[3px] rounded-full"
                          style={{ color: colorFor(e.category), background: colorFor(e.category) + '1F' }}>
                          <span className="w-2 h-2 rounded-full" style={{ background: colorFor(e.category) }} />
                          {e.category}
                        </span>
                      </td>
                      <td className="px-4 py-3.5 text-sm border-b border-line-soft text-right font-bold text-ink tabular-nums">{inr.format(e.amount)}</td>
                      <td className="px-4 py-3.5 text-sm border-b border-line-soft">
                        {e.receipt_url
                          ? <a href={e.receipt_url} target="_blank" rel="noreferrer">
                              <img src={e.receipt_url} alt="receipt" className="w-[34px] h-[34px] object-cover rounded-[7px] border border-line block" /></a>
                          : <span className="w-[34px] h-[34px] rounded-[7px] border border-dashed border-line flex items-center justify-center text-faint">
                              <Icon name="receipt" size={15} strokeWidth={1.6} /></span>}
                      </td>
                      <td className="px-4 py-3.5 text-sm border-b border-line-soft text-right whitespace-nowrap">
                        <button className={pillSettle + ' mr-2'} onClick={() => setSettled(e.id, view !== 'settled')}>
                          {view === 'settled' ? 'Reopen' : 'Settle'}
                        </button>
                        <button className={pillDelete} onClick={() => deleteExpense(e.id)}>Delete</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </main>
      </div>
    </div>
  )
}