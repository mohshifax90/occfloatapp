"use client"

import { useMemo, useState } from "react"
import { BarChart3, CalendarDays, Plane, Search, Upload, Users } from "lucide-react"
import { Button } from "@/components/ui/button"
import { DataGrid, type GridColDef } from "@mui/x-data-grid"
import { BarChart } from "@mui/x-charts/BarChart"
import { LineChart, lineClasses, type AnimatedLineProps } from "@mui/x-charts/LineChart"
import Box from "@mui/material/Box"
import Stack from "@mui/material/Stack"
import FormControlLabel from "@mui/material/FormControlLabel"
import Switch from "@mui/material/Switch"
import { ScatterChart } from "@mui/x-charts/ScatterChart"
import { PieChart } from "@mui/x-charts/PieChart"
import Typography from "@mui/material/Typography"
import { SparkLineChart, type SparkLineChartProps } from "@mui/x-charts/SparkLineChart"
import { chartsAxisHighlightClasses } from "@mui/x-charts/ChartsAxisHighlight"
import type { FlightEntry } from "@/components/ops-control-timeline"

type Props = {
  flights: FlightEntry[]
  setFlights: (next: FlightEntry[]) => void
}

type TabKey = "flights" | "monthly" | "service" | "crewpool"
type CrewMetricShape = {
  avgAircraftPerDay: number
  avgBlockPerDay: number
  reqByFdp: number
  reqByFtl: number
  reqCore: number
  totalPerDay: number
  totalReq: number
}

type CrewSection7Baseline = {
  avgAircraftPerDay: number
  avgBlockPerDay: number
  fdpPerAircraftHours: number
  totalFdpDemandPerDay: number
  ftlLimitHoursPer28Days: number
  ftlAllowableAvgPerDay: number
  fdpSustainablePerCrewPerDay: number
  reqByFtl: number
  reqByFdp: number
  reqCore: number
  leave: number
  crewOff: number
  admin: number
  training: number
  nonFlyingPerDay: number
  totalReq: number
  crewRatioMin: number
  crewRatioMax: number
}
type SheetJSWindow = Window & {
  XLSX?: {
    read: (data: ArrayBuffer, opts: { type: "array" }) => {
      SheetNames: string[]
      Sheets: Record<string, unknown>
    }
    utils: {
      sheet_to_json: (ws: unknown, opts?: { header?: number; defval?: unknown }) => unknown[][]
    }
  }
}
const SHEETJS_CDN = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"
let sheetJsPromise: Promise<void> | null = null
const CREW_RATIO_SECTION_7_BASELINE: CrewSection7Baseline = {
  avgAircraftPerDay: 9.25,
  avgBlockPerDay: 50.4,
  fdpPerAircraftHours: 13.5,
  totalFdpDemandPerDay: 135,
  ftlLimitHoursPer28Days: 100,
  ftlAllowableAvgPerDay: 4.5,
  fdpSustainablePerCrewPerDay: 10.5,
  reqByFtl: 11.2,
  reqByFdp: 12.85,
  reqCore: 12,
  leave: 5,
  crewOff: 5,
  admin: 1.3,
  training: 1.3,
  nonFlyingPerDay: 12.6,
  totalReq: 24.6,
  crewRatioMin: 2.4,
  crewRatioMax: 2.5,
}

function loadSheetJs(): Promise<void> {
  if (typeof window === "undefined") return Promise.reject(new Error("SSR"))
  const w = window as SheetJSWindow
  if (w.XLSX) return Promise.resolve()
  if (sheetJsPromise) return sheetJsPromise
  sheetJsPromise = new Promise<void>((resolve, reject) => {
    const existing = document.querySelector(`script[src="${SHEETJS_CDN}"]`) as HTMLScriptElement | null
    if (existing) {
      existing.addEventListener("load", () => resolve())
      existing.addEventListener("error", () => reject(new Error("Failed to load SheetJS")))
      return
    }
    const s = document.createElement("script")
    s.src = SHEETJS_CDN
    s.async = true
    s.onload = () => resolve()
    s.onerror = () => reject(new Error("Failed to load SheetJS"))
    document.head.appendChild(s)
  })
  return sheetJsPromise
}

function canonical(s: unknown): string {
  return String(s || "").trim().toLowerCase().replace(/\s+/g, " ")
}

function excelSerialToYmd(serial: number): string {
  const epoch = new Date(Date.UTC(1899, 11, 30))
  const dt = new Date(epoch.getTime() + Math.floor(serial) * 24 * 60 * 60 * 1000)
  const y = dt.getUTCFullYear()
  const m = String(dt.getUTCMonth() + 1).padStart(2, "0")
  const d = String(dt.getUTCDate()).padStart(2, "0")
  return `${y}-${m}-${d}`
}

function normalizeExcelDate(value: unknown): string {
  if (typeof value === "number") return excelSerialToYmd(value)
  const s = String(value || "").trim()
  if (!s) return ""
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (iso) return s
  const dmy = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/)
  if (dmy) {
    const dd = String(Number(dmy[1])).padStart(2, "0")
    const mm = String(Number(dmy[2])).padStart(2, "0")
    const yyRaw = Number(dmy[3])
    const yyyy = yyRaw < 100 ? 2000 + yyRaw : yyRaw
    return `${yyyy}-${mm}-${dd}`
  }
  const dt = new Date(s)
  if (!Number.isNaN(dt.getTime())) {
    const y = dt.getFullYear()
    const m = String(dt.getMonth() + 1).padStart(2, "0")
    const d = String(dt.getDate()).padStart(2, "0")
    return `${y}-${m}-${d}`
  }
  return ""
}

function normalizeExcelTime(value: unknown): string {
  if (typeof value === "number") {
    const mins = Math.round((value % 1) * 24 * 60)
    const h = String(Math.floor(mins / 60) % 24).padStart(2, "0")
    const m = String(mins % 60).padStart(2, "0")
    return `${h}:${m}`
  }
  const s = String(value || "").trim()
  const m = s.match(/^(\d{1,2}):(\d{2})/)
  if (!m) return ""
  return `${String(Number(m[1])).padStart(2, "0")}:${m[2]}`
}

function toMin(hhmm: string): number | null {
  const m = (hhmm || "").trim().match(/^(\d{1,2}):(\d{2})$/)
  if (!m) return null
  const h = Number(m[1])
  const mm = Number(m[2])
  if (!Number.isInteger(h) || !Number.isInteger(mm) || h < 0 || h > 23 || mm < 0 || mm > 59) return null
  return h * 60 + mm
}

function durationHours(dep: string, arr: string): number {
  const d = toMin(dep)
  const a = toMin(arr)
  if (d == null || a == null) return 0
  const diff = a >= d ? a - d : 24 * 60 - d + a
  return diff / 60
}

function buildServiceWindows(rows: FlightEntry[]) {
  const map = new Map<string, { aircraft: string; date: string; firstDep: number; lastArr: number }>()
  rows.forEach((f) => {
    const aircraft = (f.aircraft || "").trim()
    const date = (f.scheduleDate || "").trim()
    const dep = toMin(f.depTime)
    const arr = toMin(f.arrTime)
    if (!aircraft || !date || dep == null || arr == null) return
    const key = `${aircraft}|${date}`
    const prev = map.get(key)
    if (!prev) {
      map.set(key, { aircraft, date, firstDep: dep, lastArr: arr })
      return
    }
    if (dep < prev.firstDep) prev.firstDep = dep
    if (arr > prev.lastArr) prev.lastArr = arr
  })
  return [...map.values()].map((w) => {
    let mins = w.lastArr - w.firstDep
    if (mins < 0) mins += 24 * 60
    return {
      aircraft: w.aircraft,
      date: w.date,
      serviceHours: mins / 60,
    }
  })
}

function CustomLine(props: AnimatedLineProps) {
  const { d, ownerState, skipAnimation, className, ...other } = props
  return (
    <>
      <path
        d={d}
        stroke={ownerState.gradientId ? `url(#${ownerState.gradientId})` : ownerState.color}
        strokeWidth={ownerState.isHighlighted ? 4 : 2}
        strokeLinejoin="round"
        fill="none"
        filter={ownerState.isHighlighted ? "brightness(120%)" : undefined}
        opacity={ownerState.isFaded ? 0.3 : 1}
        className={className}
      />
      <path d={d} stroke="transparent" strokeWidth={25} fill="none" {...other} />
    </>
  )
}

export default function AircraftReportModule({ flights, setFlights }: Props) {
  const [tab, setTab] = useState<TabKey>("flights")
  const [search, setSearch] = useState("")
  const [fromDate, setFromDate] = useState("")
  const [toDate, setToDate] = useState("")
  const [heatMode, setHeatMode] = useState<"BOTH" | "DEP" | "ARR">("BOTH")
  const [selectedQuarter, setSelectedQuarter] = useState("All")
  const [selectedRegs, setSelectedRegs] = useState<Set<string>>(new Set())
  const [selectedOrigins, setSelectedOrigins] = useState<Set<string>>(new Set())
  const [selectedDestinations, setSelectedDestinations] = useState<Set<string>>(new Set())
  const [showFilterCard, setShowFilterCard] = useState(true)
  const [filtersExpanded, setFiltersExpanded] = useState(true)
  const [openFilterSections, setOpenFilterSections] = useState({
    qtr: true,
    reg: true,
    org: true,
    dst: true,
  })
  const [monthlyWithArea, setMonthlyWithArea] = useState(false)
  const [monthlyChartMonthFilter, setMonthlyChartMonthFilter] = useState<string | null>(null)
  const [kpiIndex, setKpiIndex] = useState<number | null>(null)
  const [monthlyKpiIndex, setMonthlyKpiIndex] = useState<number | null>(null)
  const [serviceKpiIndex, setServiceKpiIndex] = useState<number | null>(null)
  const [flightStatusFilter, setFlightStatusFilter] = useState<"All" | "Departed" | "Landed">("All")
  const [heatCellFilter, setHeatCellFilter] = useState<{ ac: string | null; hour: number | null }>({
    ac: null,
    hour: null,
  })
  const [crewModel, setCrewModel] = useState({
    maxOpHours: 12.5,
    maxFdpHours: 9.5,
    maxFtlHours: 4.5,
    standby: 1,
    admin: 1.5,
    releaseGanOps: 0,
    training: 1,
    crewOff: 4,
    leave: 5,
    captainAvailable: null as number | null,
    foAvailable: null as number | null,
  })

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase()
    return flights.filter((f) => {
      const d = (f.scheduleDate || "").trim()
      if (fromDate && d && d < fromDate) return false
      if (toDate && d && d > toDate) return false
      if (selectedQuarter !== "All") {
        const m = d.match(/^(\d{4})-(\d{2})-\d{2}$/)
        if (m) {
          const y = m[1]
          const month = Number(m[2])
          const q = `Q${Math.floor((month - 1) / 3) + 1}`
          const yq = `${y}-${q}`
          if (yq !== selectedQuarter) return false
        } else {
          return false
        }
      }
      const reg = (f.aircraft || "").trim()
      if (selectedRegs.size > 0 && !selectedRegs.has(reg)) return false
      const org = (f.origin || "").trim()
      if (selectedOrigins.size > 0 && !selectedOrigins.has(org)) return false
      const dst = (f.destination || "").trim()
      if (selectedDestinations.size > 0 && !selectedDestinations.has(dst)) return false
      if (!q) return true
      const blob = [
        f.flightNo,
        f.aircraft,
        f.origin,
        f.destination,
        f.depTime,
        f.arrTime,
        f.pic,
        f.sic,
        f.ca,
        f.remarks,
      ]
        .join(" ")
        .toLowerCase()
      return blob.includes(q)
    })
  }, [flights, search, fromDate, toDate, selectedQuarter, selectedRegs, selectedOrigins, selectedDestinations])
  const filteredAllQuarter = useMemo(() => {
    const q = search.trim().toLowerCase()
    return flights.filter((f) => {
      const d = (f.scheduleDate || "").trim()
      if (fromDate && d && d < fromDate) return false
      if (toDate && d && d > toDate) return false
      const reg = (f.aircraft || "").trim()
      if (selectedRegs.size > 0 && !selectedRegs.has(reg)) return false
      const org = (f.origin || "").trim()
      if (selectedOrigins.size > 0 && !selectedOrigins.has(org)) return false
      const dst = (f.destination || "").trim()
      if (selectedDestinations.size > 0 && !selectedDestinations.has(dst)) return false
      if (!q) return true
      const blob = [
        f.flightNo,
        f.aircraft,
        f.origin,
        f.destination,
        f.depTime,
        f.arrTime,
        f.pic,
        f.sic,
        f.ca,
        f.remarks,
      ]
        .join(" ")
        .toLowerCase()
      return blob.includes(q)
    })
  }, [flights, search, fromDate, toDate, selectedRegs, selectedOrigins, selectedDestinations])
  const currentFilterRangeLabel = useMemo(() => {
    if (selectedQuarter !== "All") return selectedQuarter
    const fmt = (ymd: string) => {
      const m = ymd.match(/^(\d{4})-(\d{2})-(\d{2})$/)
      if (!m) return ymd || "-"
      return `${m[3]}/${m[2]}/${m[1].slice(-2)}`
    }
    const f = fromDate ? fmt(fromDate) : "start"
    const t = toDate ? fmt(toDate) : "end"
    return `${f} to ${t}`
  }, [fromDate, toDate, selectedQuarter])
  const quarterOptions = useMemo(() => {
    const set = new Set<string>()
    flights.forEach((f) => {
      const d = (f.scheduleDate || "").trim()
      const m = d.match(/^(\d{4})-(\d{2})-\d{2}$/)
      if (!m) return
      const y = m[1]
      const month = Number(m[2])
      const q = `Q${Math.floor((month - 1) / 3) + 1}`
      set.add(`${y}-${q}`)
    })
    return ["All", ...Array.from(set).sort()]
  }, [flights])
  const regOptions = useMemo(
    () => Array.from(new Set(flights.map((f) => (f.aircraft || "").trim()).filter(Boolean))).sort(),
    [flights],
  )
  const originOptions = useMemo(
    () => Array.from(new Set(flights.map((f) => (f.origin || "").trim()).filter(Boolean))).sort(),
    [flights],
  )
  const destinationOptions = useMemo(
    () => Array.from(new Set(flights.map((f) => (f.destination || "").trim()).filter(Boolean))).sort(),
    [flights],
  )

  const stats = useMemo(() => {
    let totalHours = 0
    let totalPax = 0
    const ac = new Set<string>()
    filtered.forEach((f) => {
      totalHours += durationHours(f.depTime, f.arrTime)
      ac.add((f.aircraft || "").trim())
      const n = parseInt(String(f.pax || "").replace(/[^0-9-]/g, ""), 10)
      if (!Number.isNaN(n) && n > 0) totalPax += n
    })
    return {
      legs: filtered.length,
      aircraft: [...ac].filter(Boolean).length,
      totalHours,
      totalPax,
    }
  }, [filtered])
  const kpiTimeline = useMemo(() => {
    const byDate = new Map<string, { flights: number; pax: number; hours: number; ac: Set<string> }>()
    filtered.forEach((f) => {
      const date = (f.scheduleDate || "").trim()
      if (!date) return
      const row = byDate.get(date) ?? { flights: 0, pax: 0, hours: 0, ac: new Set<string>() }
      row.flights += 1
      row.hours += durationHours(f.depTime, f.arrTime)
      const n = parseInt(String(f.pax || "").replace(/[^0-9-]/g, ""), 10)
      if (!Number.isNaN(n) && n > 0) row.pax += n
      if ((f.aircraft || "").trim()) row.ac.add((f.aircraft || "").trim())
      byDate.set(date, row)
    })
    const rows = [...byDate.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([date, v]) => ({
        date,
        flights: v.flights,
        pax: v.pax,
        hours: Number(v.hours.toFixed(2)),
        aircraft: v.ac.size,
      }))
    return rows
  }, [filtered])
  const sparkLabels = useMemo(() => kpiTimeline.map((x) => x.date), [kpiTimeline])
  const sparkCommonSx = {
    [`& .${lineClasses.area}`]: { opacity: 0.2 },
    [`& .${lineClasses.line}`]: { strokeWidth: 3 },
    [`& .${chartsAxisHighlightClasses.root}`]: {
      stroke: "rgb(37, 99, 235)",
      strokeDasharray: "none",
      strokeWidth: 2,
    },
  }
  const buildSparkSettings = (data: number[], axisId: string): SparkLineChartProps => ({
    data,
    baseline: "min",
    margin: { bottom: 0, top: 5, left: 4, right: 0 },
    xAxis: { id: axisId, data: sparkLabels },
    yAxis: {
      domainLimit: (_min, maxValue) => ({
        min: -Number(maxValue || 0) / 6,
        max: Number(maxValue || 0),
      }),
    },
    sx: sparkCommonSx,
    slotProps: { lineHighlight: { r: 3 } },
    clipAreaOffset: { top: 2, bottom: 2 },
    axisHighlight: { x: "line" },
  })
  const buildSparkSettingsByLabels = (data: number[], labels: string[], axisId: string): SparkLineChartProps => ({
    data,
    baseline: "min",
    margin: { bottom: 0, top: 5, left: 4, right: 0 },
    xAxis: { id: axisId, data: labels },
    yAxis: {
      domainLimit: (_min, maxValue) => ({
        min: -Number(maxValue || 0) / 6,
        max: Number(maxValue || 0),
      }),
    },
    sx: sparkCommonSx,
    slotProps: { lineHighlight: { r: 3 } },
    clipAreaOffset: { top: 2, bottom: 2 },
    axisHighlight: { x: "line" },
  })

  const monthlyRows = useMemo(() => {
    const map = new Map<string, { monthKey: string; month: string; legs: number; hours: number }>()
    filtered.forEach((f) => {
      const key = (f.scheduleDate || "").slice(0, 7)
      if (!key) return
      const dt = new Date(`${key}-01T00:00:00`)
      const label = Number.isNaN(dt.getTime())
        ? key
        : `${dt.toLocaleString("en-US", { month: "short" })}-${String(dt.getFullYear()).slice(-2)}`
      const item = map.get(key) ?? { monthKey: key, month: label, legs: 0, hours: 0 }
      item.legs += 1
      item.hours += durationHours(f.depTime, f.arrTime)
      map.set(key, item)
    })
    return [...map.values()].sort((a, b) => a.monthKey.localeCompare(b.monthKey))
  }, [filtered])
  const monthlyChartLabels = useMemo(() => monthlyRows.map((m) => m.month), [monthlyRows])
  const monthlyHoursSeries = useMemo(() => monthlyRows.map((m) => Number(m.hours.toFixed(2))), [monthlyRows])
  const monthlyLegsSeries = useMemo(() => monthlyRows.map((m) => m.legs), [monthlyRows])
  const monthlyScatterDataset = useMemo(
    () =>
      monthlyRows.map((m, i) => ({
        id: i,
        x1: i + 1,
        y1: Number(m.hours.toFixed(2)),
        x2: i + 1,
        y2: m.legs,
      })),
    [monthlyRows],
  )
  const monthlyAvgAircraftRows = useMemo(() => {
    const monthDayAircraft = new Map<string, Map<string, Set<string>>>()
    filtered.forEach((f) => {
      const date = (f.scheduleDate || "").trim()
      const ac = (f.aircraft || "").trim()
      if (!date || !ac) return
      const monthKey = date.slice(0, 7)
      if (!monthKey) return
      const dayMap = monthDayAircraft.get(monthKey) ?? new Map<string, Set<string>>()
      const acSet = dayMap.get(date) ?? new Set<string>()
      acSet.add(ac)
      dayMap.set(date, acSet)
      monthDayAircraft.set(monthKey, dayMap)
    })
    return Array.from(monthDayAircraft.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([monthKey, dayMap]) => {
        const dt = new Date(`${monthKey}-01T00:00:00`)
        const month = Number.isNaN(dt.getTime())
          ? monthKey
          : `${dt.toLocaleString("en-US", { month: "short" })}-${String(dt.getFullYear()).slice(-2)}`
        const dailyCounts = Array.from(dayMap.values()).map((set) => set.size)
        const avg = dailyCounts.length
          ? dailyCounts.reduce((a, b) => a + b, 0) / dailyCounts.length
          : 0
        return { id: monthKey, month, avgAircraft: Math.round(avg) }
      })
  }, [filtered])
  const monthlySparkLabels = useMemo(() => monthlyRows.map((m) => m.month), [monthlyRows])
  const monthlyAvgAircraftByMonth = useMemo(
    () => new Map(monthlyAvgAircraftRows.map((r) => [r.id, r.avgAircraft])),
    [monthlyAvgAircraftRows],
  )
  const monthlyTableRows = useMemo(
    () =>
      monthlyRows
        .filter((r) => (monthlyChartMonthFilter ? r.month === monthlyChartMonthFilter : true))
        .map((r) => ({
          ...r,
          avgAircraft: monthlyAvgAircraftByMonth.get(r.monthKey) ?? 0,
        })),
    [monthlyRows, monthlyChartMonthFilter, monthlyAvgAircraftByMonth],
  )
  const nowRef = useMemo(() => new Date(), [])
  const todayYmd = useMemo(
    () => `${nowRef.getFullYear()}-${String(nowRef.getMonth() + 1).padStart(2, "0")}-${String(nowRef.getDate()).padStart(2, "0")}`,
    [nowRef],
  )
  const nowMinToday = useMemo(() => nowRef.getHours() * 60 + nowRef.getMinutes(), [nowRef])
  const getFlightStatus = (f: FlightEntry): "Departed" | "Landed" | "Scheduled" => {
    const d = (f.scheduleDate || "").trim()
    if (!d) return "Scheduled"
    if (d < todayYmd) return "Landed"
    if (d > todayYmd) return "Scheduled"
    const dep = toMin(f.depTime)
    const arr = toMin(f.arrTime)
    if (arr != null && nowMinToday >= arr) return "Landed"
    if (dep != null && nowMinToday >= dep) return "Departed"
    return "Scheduled"
  }
  const flightsByStatus = useMemo(() => {
    if (flightStatusFilter === "All") return filtered
    return filtered.filter((f) => getFlightStatus(f) === flightStatusFilter)
  }, [filtered, flightStatusFilter])
  const monthlyMovementRows = useMemo(
    () =>
      filtered
        .filter((f) => {
          const text = `${f.remarks || ""} ${f.flightNo || ""}`.toLowerCase()
          return !(
            text.includes("cancel") ||
            text.includes("cancelled") ||
            text.includes("canceled") ||
            text.includes("cnl")
          )
        })
        .filter((f) => {
          if (!monthlyChartMonthFilter) return true
          const d = (f.scheduleDate || "").trim()
          const m = d.match(/^(\d{4})-(\d{2})-\d{2}$/)
          if (!m) return false
          const dt = new Date(`${m[1]}-${m[2]}-01T00:00:00`)
          const label = Number.isNaN(dt.getTime())
            ? `${m[1]}-${m[2]}`
            : `${dt.toLocaleString("en-US", { month: "short" })}-${String(dt.getFullYear()).slice(-2)}`
          return label === monthlyChartMonthFilter
        })
        .map((f, i) => {
          const org = (f.origin || "").trim().toUpperCase()
          const dst = (f.destination || "").trim().toUpperCase()
          const isStartFromMle = org === "MLE" && !!dst
          const isEndToMle = dst === "MLE" && !!org
          const movement = isStartFromMle || isEndToMle ? 1 : 0
          const movementType = isStartFromMle
            ? "ORG Start (MLE -> any)"
            : isEndToMle
              ? "DST End (any -> MLE)"
              : "Non-MLE leg"
          return {
            id: `${f.id || i}-mv`,
            scheduleDate: f.scheduleDate,
            flightNo: f.flightNo,
            aircraft: f.aircraft,
            origin: f.origin,
            destination: f.destination,
            route: `${(f.origin || "").trim()}-${(f.destination || "").trim()}`,
            movement,
            movementType,
          }
        }),
    [filtered, monthlyChartMonthFilter],
  )
  const monthlyMovementSummaryRows = useMemo(() => {
    const map = new Map<string, { month: string; aircraft: string; depFromMle: number; arrToMle: number; movements: number }>()
    monthlyMovementRows.forEach((r) => {
      const d = String(r.scheduleDate || "")
      const m = d.match(/^(\d{4})-(\d{2})-\d{2}$/)
      const monthLabel = m
        ? (() => {
            const dt = new Date(`${m[1]}-${m[2]}-01T00:00:00`)
            return Number.isNaN(dt.getTime())
              ? `${m[1]}-${m[2]}`
              : `${dt.toLocaleString("en-US", { month: "short" })}-${String(dt.getFullYear()).slice(-2)}`
          })()
        : "-"
      if (Number(r.movement || 0) !== 1) return
      const aircraft = String(r.aircraft || "-")
      const org = String(r.origin || "").trim().toUpperCase()
      const dst = String(r.destination || "").trim().toUpperCase()
      const key = `${monthLabel}|${aircraft}`
      const prev = map.get(key) ?? { month: monthLabel, aircraft, depFromMle: 0, arrToMle: 0, movements: 0 }
      if (org === "MLE") prev.depFromMle += 1
      if (dst === "MLE") prev.arrToMle += 1
      prev.movements = prev.depFromMle + prev.arrToMle
      map.set(key, prev)
    })
    return Array.from(map.values())
      .sort((a, b) => (a.month === b.month ? a.aircraft.localeCompare(b.aircraft) : a.month.localeCompare(b.month)))
      .map((x, i) => ({ id: `ms-${i}`, ...x }))
  }, [monthlyMovementRows])
  const summaryMetrics = useMemo(() => {
    const mleRows = filtered.filter(
      (f) => (f.origin || "").trim().toUpperCase() === "MLE" || (f.destination || "").trim().toUpperCase() === "MLE",
    )
    const blockHours = filtered.reduce((a, f) => a + durationHours(f.depTime, f.arrTime), 0)
    const flightHours = blockHours
    const pax = mleRows.reduce((a, f) => {
      const n = parseInt(String(f.pax || "").replace(/[^0-9-]/g, ""), 10)
      return a + (!Number.isNaN(n) && n > 0 ? n : 0)
    }, 0)
    const serviceWindows = buildServiceWindows(filtered)
    const serviceTotal = serviceWindows.reduce((a, w) => a + w.serviceHours, 0)
    const serviceAvg = serviceWindows.length ? serviceTotal / serviceWindows.length : 0
    const activeByDay = new Map<string, Set<string>>()
    serviceWindows.forEach((w) => {
      const set = activeByDay.get(w.date) ?? new Set<string>()
      set.add(w.aircraft)
      activeByDay.set(w.date, set)
    })
    const avgAircraftPerDay = activeByDay.size
      ? Array.from(activeByDay.values()).reduce((a, s) => a + s.size, 0) / activeByDay.size
      : 0
    const uniqueAircraft = new Set(filtered.map((f) => (f.aircraft || "").trim()).filter(Boolean)).size
    const avgPerAircraft = uniqueAircraft ? serviceTotal / uniqueAircraft : 0
    return {
      flights: mleRows.length,
      blockHours,
      flightHours,
      passengers: pax,
      serviceTotal,
      serviceAvg,
      avgAircraftPerDay,
      avgPerAircraft,
    }
  }, [filtered])
  const calcCrewMetrics = (rows: FlightEntry[]): CrewMetricShape => {
    const byDay = new Map<string, { ac: Set<string>; block: number }>()
    rows.forEach((f) => {
      const day = (f.scheduleDate || "").trim()
      if (!day) return
      const item = byDay.get(day) ?? { ac: new Set<string>(), block: 0 }
      const ac = (f.aircraft || "").trim()
      if (ac) item.ac.add(ac)
      item.block += durationHours(f.depTime, f.arrTime)
      byDay.set(day, item)
    })
    const days = Array.from(byDay.values())
    const avgAircraftPerDay = days.length ? days.reduce((a, d) => a + d.ac.size, 0) / days.length : 0
    const avgBlockPerDay = days.length ? days.reduce((a, d) => a + d.block, 0) / days.length : 0
    const reqByFdp = (avgAircraftPerDay * crewModel.maxOpHours) / crewModel.maxFdpHours
    const reqByFtl = avgBlockPerDay / crewModel.maxFtlHours
    const reqCore = (reqByFdp + reqByFtl) / 2
    const totalPerDay = reqCore + crewModel.standby + crewModel.admin + crewModel.releaseGanOps + crewModel.training
    const totalReq = totalPerDay + crewModel.crewOff + crewModel.leave
    return {
      avgAircraftPerDay,
      avgBlockPerDay,
      reqByFdp,
      reqByFtl,
      reqCore,
      totalPerDay,
      totalReq,
    }
  }
  const quarterOrder = { Q1: 1, Q2: 2, Q3: 3, Q4: 4 } as const
  const quarterList = useMemo(() => {
    const set = new Set<string>()
    filteredAllQuarter.forEach((f) => {
      const d = (f.scheduleDate || "").trim()
      const m = d.match(/^(\d{4})-(\d{2})-\d{2}$/)
      if (!m) return
      const y = Number(m[1])
      const month = Number(m[2])
      const q = `Q${Math.floor((month - 1) / 3) + 1}`
      set.add(`${y}-${q}`)
    })
    return Array.from(set).sort((a, b) => {
      const ay = Number(a.slice(0, 4))
      const by = Number(b.slice(0, 4))
      if (ay !== by) return ay - by
      const aq = a.slice(5) as keyof typeof quarterOrder
      const bq = b.slice(5) as keyof typeof quarterOrder
      return (quarterOrder[aq] || 9) - (quarterOrder[bq] || 9)
    })
  }, [filteredAllQuarter])
  const currentCrewMetrics = useMemo(() => calcCrewMetrics(filtered), [filtered])
  const currentCrewDerived = useMemo(() => {
    const nonFlyingFromModel = crewModel.leave + crewModel.crewOff + crewModel.admin + crewModel.training + crewModel.releaseGanOps
    const avgAcForRatio = currentCrewMetrics.avgAircraftPerDay
    const ratio = avgAcForRatio > 0 ? currentCrewMetrics.totalReq / avgAcForRatio : 0
    const totalFdpDemandPerDay = currentCrewMetrics.avgAircraftPerDay * CREW_RATIO_SECTION_7_BASELINE.fdpPerAircraftHours
    return {
      nonFlyingFromModel,
      ratio,
      totalFdpDemandPerDay,
    }
  }, [currentCrewMetrics, crewModel])
  const quarterCrewMetrics = useMemo(() => {
    const byQuarter = new Map<string, FlightEntry[]>()
    filteredAllQuarter.forEach((f) => {
      const d = (f.scheduleDate || "").trim()
      const m = d.match(/^(\d{4})-(\d{2})-\d{2}$/)
      if (!m) return
      const y = m[1]
      const month = Number(m[2])
      const q = `Q${Math.floor((month - 1) / 3) + 1}`
      const key = `${y}-${q}`
      const arr = byQuarter.get(key) ?? []
      arr.push(f)
      byQuarter.set(key, arr)
    })
    const out = new Map<string, ReturnType<typeof calcCrewMetrics>>()
    quarterList.forEach((q) => out.set(q, calcCrewMetrics(byQuarter.get(q) ?? [])))
    return out
  }, [filteredAllQuarter, quarterList])

  const serviceRows = useMemo(() => {
    const windows = buildServiceWindows(filtered)
    const map = new Map<string, { aircraft: string; days: number; total: number }>()
    windows.forEach((w) => {
      const item = map.get(w.aircraft) ?? { aircraft: w.aircraft, days: 0, total: 0 }
      item.days += 1
      item.total += w.serviceHours
      map.set(w.aircraft, item)
    })
    return [...map.values()]
      .map((r) => ({
        ...r,
        avg: r.days > 0 ? r.total / r.days : 0,
      }))
      .sort((a, b) => b.avg - a.avg)
  }, [filtered])
  const hourlyCounts = useMemo(() => {
    const bins = Array.from({ length: 24 }, (_, h) => ({ h, count: 0 }))
    filtered.forEach((f) => {
      const d = toMin(f.depTime)
      if (d == null) return
      const h = Math.floor(d / 60)
      bins[h].count += 1
    })
    return bins
  }, [filtered])
  const heatData = useMemo(() => {
    const hours = Array.from({ length: 15 }, (_, i) => i + 5) // 05:00-19:00
    const byAc = new Map<string, number[]>()
    filtered.forEach((f) => {
      const ac = (f.aircraft || "").trim() || "Unknown"
      if (!byAc.has(ac)) byAc.set(ac, Array(24).fill(0))
      const dep = toMin(f.depTime)
      const arr = toMin(f.arrTime)
      if ((heatMode === "DEP" || heatMode === "BOTH") && dep != null) byAc.get(ac)![Math.floor(dep / 60) % 24] += 1
      if ((heatMode === "ARR" || heatMode === "BOTH") && arr != null) byAc.get(ac)![Math.floor(arr / 60) % 24] += 1
    })
    const rows = [...byAc.entries()]
      .map(([ac, bins]) => ({
        ac,
        bins,
        total: bins.reduce((a, b) => a + b, 0),
      }))
      .sort((a, b) => b.total - a.total)
    const values = rows.flatMap((r) => hours.map((h) => r.bins[h])).filter((v) => v > 0)
    const max = Math.max(1, ...values)
    return { hours, rows, max }
  }, [filtered, heatMode])
  const serviceDataRows = useMemo(() => {
    const rows = filtered.filter((f) => {
      if (!heatCellFilter.ac || heatCellFilter.hour == null) return true
      const ac = (f.aircraft || "").trim() || "Unknown"
      if (ac !== heatCellFilter.ac) return false
      const dep = toMin(f.depTime)
      const arr = toMin(f.arrTime)
      const depHr = dep == null ? null : Math.floor(dep / 60) % 24
      const arrHr = arr == null ? null : Math.floor(arr / 60) % 24
      if (heatMode === "DEP") return depHr === heatCellFilter.hour
      if (heatMode === "ARR") return arrHr === heatCellFilter.hour
      return depHr === heatCellFilter.hour || arrHr === heatCellFilter.hour
    })
    return rows.map((f) => ({
      id: f.id,
      date: (() => {
        const s = (f.scheduleDate || "").trim()
        if (!s) return "-"
        const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/)
        if (!m) return s
        return `${m[3]}.${m[2]}.${m[1]}`
      })(),
      ac: (f.aircraft || "").trim() || "-",
      flight: (f.flightNo || "").trim() || "-",
      route: `${(f.origin || "").trim() || "-"}-${(f.destination || "").trim() || "-"}`,
      std: (f.depTime || "").trim() || "-",
      sta: (f.arrTime || "").trim() || "-",
      blockHrs: Number(durationHours(f.depTime, f.arrTime).toFixed(2)),
    }))
  }, [filtered, heatCellFilter, heatMode])
  const serviceMonthly = useMemo(() => {
    const windows = buildServiceWindows(filtered)
    const monthMap = new Map<string, { monthKey: string; label: string; total: number; count: number }>()
    windows.forEach((w) => {
      const d = (w.date || "").trim()
      if (!d) return
      const key = d.slice(0, 7)
      if (!key) return
      const dt = new Date(`${key}-01T00:00:00`)
      const label = Number.isNaN(dt.getTime())
        ? key
        : `${dt.toLocaleString("en-US", { month: "short" })}-${String(dt.getFullYear()).slice(-2)}`
      const item = monthMap.get(key) ?? { monthKey: key, label, total: 0, count: 0 }
      item.total += w.serviceHours
      item.count += 1
      monthMap.set(key, item)
    })
    return [...monthMap.values()]
      .sort((a, b) => a.monthKey.localeCompare(b.monthKey))
      .map((m) => ({ ...m, avg: m.count > 0 ? m.total / m.count : 0 }))
  }, [filtered])
  const serviceSparkRows = useMemo(() => {
    const avgAcByMonthKey = new Map(monthlyAvgAircraftRows.map((r) => [r.id, r.avgAircraft]))
    return serviceMonthly.map((m) => {
      const avgAircraft = avgAcByMonthKey.get(m.monthKey) ?? 0
      const avgPerAircraft = avgAircraft > 0 ? m.total / avgAircraft : 0
      return { label: m.label, total: m.total, avg: m.avg, avgAircraft, avgPerAircraft }
    })
  }, [serviceMonthly, monthlyAvgAircraftRows])
  const serviceSparkLabels = useMemo(() => serviceSparkRows.map((r) => r.label), [serviceSparkRows])

  const flightCols: GridColDef[] = [
    { field: "scheduleDate", headerName: "Date", width: 110 },
    {
      field: "__status",
      headerName: "Status",
      width: 110,
      sortable: false,
      filterable: false,
      renderCell: (params) => {
        const s = String(params.value || "")
        const cls =
          s === "Landed"
            ? "bg-success-subtle text-success-emphasis"
            : s === "Departed"
              ? "bg-info-subtle text-info-emphasis"
              : "bg-secondary-subtle text-secondary-emphasis"
        return <span className={`badge ${cls}`}>{s || "-"}</span>
      },
    },
    { field: "flightNo", headerName: "Flight", width: 100 },
    { field: "aircraft", headerName: "AC", width: 90 },
    { field: "origin", headerName: "ORG", width: 80 },
    { field: "destination", headerName: "DST", width: 80 },
    { field: "depTime", headerName: "STD", width: 80 },
    { field: "arrTime", headerName: "STA", width: 80 },
    { field: "pax", headerName: "Pax", width: 80 },
    { field: "pic", headerName: "PIC", width: 100 },
    { field: "sic", headerName: "SIC", width: 100 },
    { field: "ca", headerName: "CA", width: 100 },
    { field: "remarks", headerName: "Remarks", minWidth: 180, flex: 1 },
  ]

  return (
    <div className="row g-3">
      <div className={showFilterCard ? "col-12 col-xl-3" : "d-none"}>
        <div className="card border h-100">
          <div className="card-body p-3" style={{ maxHeight: filtersExpanded ? "none" : "calc(100vh - 12rem)", overflow: "auto" }}>
            <div className="d-flex align-items-center justify-content-between mb-3">
              <div className="fw-semibold">Filters</div>
              <Button size="sm" variant="outline" onClick={() => setFiltersExpanded((v) => !v)}>
                {filtersExpanded ? "Compact" : "Expand Vertical"}
              </Button>
            </div>
            <button
              type="button"
              className="btn btn-sm btn-light w-100 text-start mb-2"
              onClick={() => setOpenFilterSections((p) => ({ ...p, qtr: !p.qtr }))}
            >
              QTR {openFilterSections.qtr ? "▾" : "▸"}
            </button>
            {openFilterSections.qtr ? (
            <div className="d-flex flex-wrap gap-2 mb-4">
              {quarterOptions.map((q) => (
                <button
                  key={q}
                  type="button"
                  className="btn btn-sm rounded-pill"
                  style={{
                    background: selectedQuarter === q ? "#c7f1ea" : "#e5e7eb",
                    border: `1px solid ${selectedQuarter === q ? "#0f766e" : "#d1d5db"}`,
                    color: selectedQuarter === q ? "#115e59" : "#4b5563",
                    minWidth: 70,
                  }}
                  onClick={() => setSelectedQuarter(q)}
                >
                  {q}
                </button>
              ))}
            </div>
            ) : null}

            <button
              type="button"
              className="btn btn-sm btn-light w-100 text-start mb-2"
              onClick={() => setOpenFilterSections((p) => ({ ...p, reg: !p.reg }))}
            >
              REG {openFilterSections.reg ? "▾" : "▸"}
            </button>
            {openFilterSections.reg ? (
            <div className="d-flex flex-wrap gap-2 mb-4">
              {regOptions.map((reg) => (
                <button
                  key={reg}
                  type="button"
                  className="btn btn-sm rounded-pill"
                  style={{
                    background: selectedRegs.has(reg) ? "#c7f1ea" : "#e5e7eb",
                    border: `1px solid ${selectedRegs.has(reg) ? "#0f766e" : "#d1d5db"}`,
                    color: selectedRegs.has(reg) ? "#115e59" : "#4b5563",
                    minWidth: 82,
                  }}
                  onClick={() =>
                    setSelectedRegs((prev) => {
                      const next = new Set(prev)
                      if (next.has(reg)) next.delete(reg)
                      else next.add(reg)
                      return next
                    })
                  }
                >
                  {reg}
                </button>
              ))}
            </div>
            ) : null}

            <button
              type="button"
              className="btn btn-sm btn-light w-100 text-start mb-2"
              onClick={() => setOpenFilterSections((p) => ({ ...p, org: !p.org }))}
            >
              ORIGIN {openFilterSections.org ? "▾" : "▸"}
            </button>
            {openFilterSections.org ? (
            <div className="d-flex flex-wrap gap-2">
              {originOptions.map((org) => (
                <button
                  key={org}
                  type="button"
                  className="btn btn-sm rounded-pill"
                  style={{
                    background: selectedOrigins.has(org) ? "#c7f1ea" : "#e5e7eb",
                    border: `1px solid ${selectedOrigins.has(org) ? "#0f766e" : "#d1d5db"}`,
                    color: selectedOrigins.has(org) ? "#115e59" : "#4b5563",
                    minWidth: 64,
                  }}
                  onClick={() =>
                    setSelectedOrigins((prev) => {
                      const next = new Set(prev)
                      if (next.has(org)) next.delete(org)
                      else next.add(org)
                      return next
                    })
                  }
                >
                  {org}
                </button>
              ))}
            </div>
            ) : null}
            <button
              type="button"
              className="btn btn-sm btn-light w-100 text-start mt-3 mb-2"
              onClick={() => setOpenFilterSections((p) => ({ ...p, dst: !p.dst }))}
            >
              DEST {openFilterSections.dst ? "▾" : "▸"}
            </button>
            {openFilterSections.dst ? (
            <div className="d-flex flex-wrap gap-2">
              {destinationOptions.map((dst) => (
                <button
                  key={dst}
                  type="button"
                  className="btn btn-sm rounded-pill"
                  style={{
                    background: selectedDestinations.has(dst) ? "#c7f1ea" : "#e5e7eb",
                    border: `1px solid ${selectedDestinations.has(dst) ? "#0f766e" : "#d1d5db"}`,
                    color: selectedDestinations.has(dst) ? "#115e59" : "#4b5563",
                    minWidth: 64,
                  }}
                  onClick={() =>
                    setSelectedDestinations((prev) => {
                      const next = new Set(prev)
                      if (next.has(dst)) next.delete(dst)
                      else next.add(dst)
                      return next
                    })
                  }
                >
                  {dst}
                </button>
              ))}
            </div>
            ) : null}
          </div>
        </div>
      </div>
      <div className={showFilterCard ? "col-12 col-xl-9" : "col-12"}>
    <div className="d-flex flex-column gap-3">
      <div className="d-flex justify-content-end">
        <Button
          size="sm"
          variant="outline"
          onClick={() => setShowFilterCard((v) => !v)}
        >
          {showFilterCard ? "Hide Filters" : "Show Filters"}
        </Button>
      </div>
      <div className="card border">
        <div className="card-body d-flex flex-wrap align-items-end gap-2">
          <div className="me-auto">
            <div className="text-uppercase text-muted small fw-semibold">Aircraft Report</div>
            <div className="small text-muted">Native OCC module with filter + analytics flow.</div>
          </div>
          <div className="input-group" style={{ maxWidth: 320 }}>
            <span className="input-group-text"><Search size={14} /></span>
            <input className="form-control" placeholder="Search flights..." value={search} onChange={(e) => setSearch(e.target.value)} />
          </div>
          <select
            className="form-select"
            style={{ width: 160 }}
            value={flightStatusFilter}
            onChange={(e) => setFlightStatusFilter(e.target.value as "All" | "Departed" | "Landed")}
          >
            <option value="All">Status: All</option>
            <option value="Departed">Departed</option>
            <option value="Landed">Landed</option>
          </select>
          <input type="date" className="form-control" style={{ width: 150 }} value={fromDate} onChange={(e) => setFromDate(e.target.value)} />
          <input type="date" className="form-control" style={{ width: 150 }} value={toDate} onChange={(e) => setToDate(e.target.value)} />
          <label className="btn btn-outline-secondary mb-0 d-inline-flex align-items-center gap-1">
            <Upload size={14} /> Upload Excel
            <input
              type="file"
              accept=".xlsx,.xls"
              className="d-none"
              onChange={async (e) => {
                const file = e.target.files?.[0]
                if (!file) return
                try {
                  await loadSheetJs()
                  const w = window as SheetJSWindow
                  if (!w.XLSX) throw new Error("SheetJS missing")
                  const buf = await file.arrayBuffer()
                  const wb = w.XLSX.read(buf, { type: "array" })
                  const ws = wb.Sheets[wb.SheetNames[0]]
                  const matrix = w.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" })
                  if (!Array.isArray(matrix) || matrix.length === 0) throw new Error("No data")
                  const rows2d = matrix as unknown[][]
                  let headerRowIndex = -1
                  const required = ["travel date", "ac", "org", "dst", "std", "sta"]
                  for (let i = 0; i < Math.min(rows2d.length, 25); i += 1) {
                    const hdr = rows2d[i].map(canonical)
                    const hit = required.filter((k) => hdr.includes(k)).length
                    if (hit >= 4) {
                      headerRowIndex = i
                      break
                    }
                  }
                  if (headerRowIndex < 0) throw new Error("Header not found")
                  const headers = rows2d[headerRowIndex].map((h) => String(h || "").trim())
                  const map = new Map<string, number>()
                  headers.forEach((h, i) => map.set(canonical(h), i))
                  const idx = (name: string) => map.get(canonical(name))
                  const dataRows = rows2d.slice(headerRowIndex + 1).filter((r) =>
                    r.some((v) => String(v || "").trim() !== ""),
                  )
                  const now = new Date().toISOString()
                  const next = dataRows.map((r, i) => {
                    const get = (name: string) => {
                      const p = idx(name)
                      return p == null ? "" : r[p]
                    }
                    return {
                      id: `xls_${Date.now()}_${i}`,
                      createdAt: now,
                      scheduleDate: normalizeExcelDate(get("Travel Date")),
                      aircraft: String(get("AC") || "").trim(),
                      flightNo: String(get("Flight") || "").trim(),
                      origin: String(get("ORG") || "").trim(),
                      destination: String(get("DST") || "").trim(),
                      depTime: normalizeExcelTime(get("STD")),
                      arrTime: normalizeExcelTime(get("STA")),
                      pic: String(get("PIC") || "").trim(),
                      sic: String(get("SIC") || "").trim(),
                      ca: String(get("CA") || "").trim(),
                      pax: String(get("Pax") || "").trim(),
                      remarks: String(get("Remarks") || "").trim(),
                    }
                  }).filter((x) => x.scheduleDate || x.flightNo || x.aircraft)
                  setFlights(next)
                } catch {
                  window.alert("Invalid Excel format. Please upload FlightData Excel with Travel Date, AC, ORG, DST, STD, STA columns.")
                } finally {
                  e.currentTarget.value = ""
                }
              }}
            />
          </label>
        </div>
      </div>

      <div className="d-flex flex-wrap gap-2">
        <Button variant={tab === "flights" ? "default" : "outline"} onClick={() => setTab("flights")}>
          <Plane size={14} className="me-1" /> Flights
        </Button>
        <Button variant={tab === "monthly" ? "default" : "outline"} onClick={() => setTab("monthly")}>
          <CalendarDays size={14} className="me-1" /> Monthly
        </Button>
        <Button variant={tab === "service" ? "default" : "outline"} onClick={() => setTab("service")}>
          <BarChart3 size={14} className="me-1" /> Service Window
        </Button>
        <Button variant={tab === "crewpool" ? "default" : "outline"} onClick={() => setTab("crewpool")}>
          <Users size={14} className="me-1" /> Crew Pool
        </Button>
      </div>

      {tab === "flights" ? (
        <div className="d-flex flex-column gap-3">
          <div className="row g-2">
            {[
              {
                title: "Flights",
                value: stats.legs.toLocaleString(),
                color: "rgb(37, 99, 235)",
                data: kpiTimeline.map((x) => x.flights),
                axisId: "kpi-flights",
              },
              {
                title: "Aircraft",
                value: stats.aircraft.toLocaleString(),
                color: "rgb(15, 118, 110)",
                data: kpiTimeline.map((x) => x.aircraft),
                axisId: "kpi-aircraft",
              },
              {
                title: "Total Hrs",
                value: stats.totalHours.toFixed(1),
                color: "rgb(109, 40, 217)",
                data: kpiTimeline.map((x) => x.hours),
                axisId: "kpi-hours",
              },
              {
                title: "Total Pax",
                value: stats.totalPax.toLocaleString(),
                color: "rgb(194, 65, 12)",
                data: kpiTimeline.map((x) => x.pax),
                axisId: "kpi-pax",
              },
            ].map((card) => (
              <div key={card.title} className="col-12 col-md-6 col-xl-3">
                <div className="card border h-100">
                  <div className="card-body p-3">
                    <Typography sx={{ color: "rgb(117,117,117)", fontWeight: 500, fontSize: "0.85rem", pb: 0.5 }}>
                      {kpiIndex == null ? card.title : (sparkLabels[kpiIndex] || card.title)}
                    </Typography>
                    <div className="d-flex align-items-end justify-content-between gap-2">
                      <Typography sx={{ fontSize: "1.25rem", fontWeight: 600, color: card.color }}>
                        {kpiIndex == null ? card.value : Number(card.data[kpiIndex] || 0).toLocaleString()}
                      </Typography>
                      <Box sx={{ width: 170, height: 42 }}>
                        <SparkLineChart
                          height={42}
                          width={170}
                          area
                          showHighlight
                          color={card.color}
                          onHighlightedAxisChange={(axisItems) => {
                            setKpiIndex(axisItems[0]?.dataIndex ?? null)
                          }}
                          highlightedAxis={
                            kpiIndex == null ? [] : [{ axisId: card.axisId, dataIndex: kpiIndex }]
                          }
                          {...buildSparkSettings(card.data, card.axisId)}
                        />
                      </Box>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
          <div className="row g-3">
            <div className="col-12 col-xl-7">
              <div className="card border h-100">
                <div className="card-body">
                  <div className="small fw-semibold mb-2">Departures by Hour</div>
                  <Box sx={{ width: "100%", height: 220 }}>
                    <LineChart
                      series={[{ data: hourlyCounts.map((b) => b.count), label: "Departures", area: true, color: "#2563eb" }]}
                      xAxis={[{ scaleType: "point", data: hourlyCounts.map((b) => `${String(b.h).padStart(2, "0")}:00`), height: 28 }]}
                      sx={{
                        [`& .${lineClasses.line}`]: {
                          display: "none",
                        },
                      }}
                      margin={{ right: 24, left: 36, top: 12, bottom: 24 }}
                    />
                  </Box>
                </div>
              </div>
            </div>
            <div className="col-12 col-xl-5">
              <div className="card border h-100">
                <div className="card-body">
                  <div className="small fw-semibold mb-2">Top Aircraft by Avg Service Hrs</div>
                  <Box sx={{ width: "100%", height: 220 }}>
                    <LineChart
                      series={[
                        {
                          data: serviceRows.slice(0, 10).map((r) => Number(r.avg.toFixed(2))),
                          label: "Avg Service Hrs",
                          area: true,
                          color: "#0f766e",
                        },
                      ]}
                      xAxis={[{ scaleType: "point", data: serviceRows.slice(0, 10).map((r) => r.aircraft), height: 28 }]}
                      sx={{
                        [`& .${lineClasses.line}`]: {
                          display: "none",
                        },
                      }}
                      margin={{ right: 24, left: 36, top: 12, bottom: 24 }}
                    />
                  </Box>
                </div>
              </div>
            </div>
          </div>
          <div className="card border">
            <div style={{ height: 560 }}>
              <DataGrid
                rows={flightsByStatus.map((f) => ({ ...f, __status: getFlightStatus(f) }))}
                columns={flightCols}
                disableRowSelectionOnClick
              />
            </div>
          </div>
        </div>
      ) : null}

      {tab === "monthly" ? (
        <div className="d-flex flex-column gap-3">
          <div className="row g-2">
            {[
              {
                title: "Monthly Flights",
                value: monthlyLegsSeries.reduce((a, b) => a + b, 0).toLocaleString(),
                color: "rgb(37, 99, 235)",
                data: monthlyLegsSeries,
                axisId: "mkpi-flights",
              },
              {
                title: "Monthly Block Hrs",
                value: `${monthlyHoursSeries.reduce((a, b) => a + b, 0).toFixed(1)}h`,
                color: "rgb(109, 40, 217)",
                data: monthlyHoursSeries,
                axisId: "mkpi-hours",
              },
              {
                title: "Avg AC / Month",
                value: (monthlyAvgAircraftRows.length ? Math.round(monthlyAvgAircraftRows.reduce((a, b) => a + b.avgAircraft, 0) / monthlyAvgAircraftRows.length) : 0).toString(),
                color: "rgb(15, 118, 110)",
                data: monthlyAvgAircraftRows.map((r) => r.avgAircraft),
                axisId: "mkpi-ac",
              },
              {
                title: "Months",
                value: monthlyRows.length.toLocaleString(),
                color: "rgb(194, 65, 12)",
                data: monthlyRows.map((_r, i) => i + 1),
                axisId: "mkpi-months",
              },
            ].map((card) => (
              <div key={card.title} className="col-12 col-md-6 col-xl-3">
                <div className="card border h-100">
                  <div className="card-body p-3">
                    <Typography sx={{ color: "rgb(117,117,117)", fontWeight: 500, fontSize: "0.85rem", pb: 0.5 }}>
                      {monthlyKpiIndex == null ? card.title : (monthlySparkLabels[monthlyKpiIndex] || card.title)}
                    </Typography>
                    <div className="d-flex align-items-end justify-content-between gap-2">
                      <Typography sx={{ fontSize: "1.25rem", fontWeight: 600, color: card.color }}>
                        {monthlyKpiIndex == null ? card.value : Number(card.data[monthlyKpiIndex] || 0).toLocaleString()}
                      </Typography>
                      <Box sx={{ width: 170, height: 42 }}>
                        <SparkLineChart
                          height={42}
                          width={170}
                          area
                          showHighlight
                          color={card.color}
                          onHighlightedAxisChange={(axisItems) => setMonthlyKpiIndex(axisItems[0]?.dataIndex ?? null)}
                          highlightedAxis={monthlyKpiIndex == null ? [] : [{ axisId: card.axisId, dataIndex: monthlyKpiIndex }]}
                          {...buildSparkSettingsByLabels(card.data, monthlySparkLabels, card.axisId)}
                        />
                      </Box>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
          <div className="card border">
            <div className="card-body">
              <Stack direction={{ xs: "column", xl: "row" }} spacing={1} sx={{ width: "100%" }}>
                <Box sx={{ flexGrow: 1 }}>
                  <LineChart
                    series={[
                      { data: monthlyHoursSeries, label: "Block Hrs", area: monthlyWithArea },
                      { data: monthlyLegsSeries, label: "Flights", area: monthlyWithArea },
                    ]}
                    xAxis={[{ scaleType: "point", data: monthlyChartLabels }]}
                    slots={monthlyWithArea ? {} : { line: CustomLine }}
                    height={360}
                  />
                  <div className="d-flex align-items-center gap-2 flex-wrap mt-2">
                    <span className="small text-muted">Chart Filter:</span>
                    <button
                      type="button"
                      className={`btn btn-sm rounded-pill ${monthlyChartMonthFilter === null ? "btn-success" : "btn-outline-secondary"}`}
                      onClick={() => setMonthlyChartMonthFilter(null)}
                    >
                      All
                    </button>
                    {monthlyChartLabels.map((m) => (
                      <button
                        key={`mf-${m}`}
                        type="button"
                        className={`btn btn-sm rounded-pill ${monthlyChartMonthFilter === m ? "btn-success" : "btn-outline-secondary"}`}
                        onClick={() => setMonthlyChartMonthFilter(m)}
                      >
                        {m}
                      </button>
                    ))}
                  </div>
                </Box>
                <Stack direction={{ xs: "row", xl: "column" }} spacing={2} useFlexGap sx={{ justifyContent: "center", flexWrap: "wrap" }}>
                  <FormControlLabel control={<Switch checked={monthlyWithArea} onChange={(e) => setMonthlyWithArea(e.target.checked)} />} label="Fill line area" />
                </Stack>
              </Stack>
            </div>
          </div>
          <div className="card border">
            <div className="card-body">
              <div className="small fw-semibold mb-2">Avg Aircraft Utilized / Month</div>
              {monthlyAvgAircraftRows.length === 0 ? (
                <div className="small text-muted">No monthly aircraft utilization data.</div>
              ) : (
                <>
                  <LineChart
                    height={280}
                    xAxis={[{ scaleType: "point", data: monthlyAvgAircraftRows.map((r) => r.month), label: "Month" }]}
                    yAxis={[{ label: "Avg Aircraft / Day", min: 0 }]}
                    series={[
                      {
                        data: monthlyAvgAircraftRows.map((r) => r.avgAircraft),
                        label: "Avg Aircraft Utilized / Month",
                        color: "#0f766e",
                        showMark: true,
                      },
                    ]}
                  />
                  <div
                    className="mt-0 px-1"
                    style={{
                      display: "grid",
                      gridTemplateColumns: `repeat(${Math.max(monthlyAvgAircraftRows.length, 1)}, minmax(0, 1fr))`,
                      gap: "4px",
                    }}
                  >
                    {monthlyAvgAircraftRows.map((r) => (
                      <span
                        key={`avg-${r.id}`}
                        className="badge border text-center"
                        style={{
                          background:
                            r.avgAircraft >= 9
                              ? "#dcfce7"
                              : r.avgAircraft >= 7
                                ? "#fef3c7"
                                : "#fee2e2",
                          color:
                            r.avgAircraft >= 9
                              ? "#166534"
                              : r.avgAircraft >= 7
                                ? "#92400e"
                                : "#991b1b",
                          borderColor:
                            r.avgAircraft >= 9
                              ? "#86efac"
                              : r.avgAircraft >= 7
                                ? "#fcd34d"
                                : "#fca5a5",
                        }}
                      >
                        {Math.round(r.avgAircraft)}
                      </span>
                    ))}
                  </div>
                </>
              )}
            </div>
          </div>
          <div className="card border">
            <div style={{ height: 360 }}>
              <DataGrid
                rows={monthlyTableRows.map((r) => ({ id: r.month, ...r, hours: Number(r.hours.toFixed(2)) }))}
                columns={[
                  { field: "month", headerName: "Month", width: 140 },
                  { field: "legs", headerName: "Flights", width: 120 },
                  { field: "hours", headerName: "Block Hrs", width: 140 },
                  {
                    field: "avgAircraft",
                    headerName: "Avg Aircraft / Month",
                    width: 180,
                    valueFormatter: (params) => String(Math.round(Number(params || 0))),
                  },
                ]}
                disableRowSelectionOnClick
              />
            </div>
          </div>
          <div className="card border">
            <div className="card-body pb-2">
              <div className="small fw-semibold">Movement Data Sheet (ORI-DEST = 1 movement)</div>
              <div className="small text-muted">
                Count by Aircraft. Movement = ORG MLE{">"}any or any-{">"}MLE DST.
              </div>
            </div>
            <div style={{ height: 320 }}>
              <DataGrid
                rows={monthlyMovementSummaryRows}
                columns={[
                  { field: "month", headerName: "Month", width: 120 },
                  { field: "aircraft", headerName: "Aircraft", width: 120 },
                  { field: "depFromMle", headerName: "MLE -> Any", width: 120 },
                  { field: "arrToMle", headerName: "Any -> MLE", width: 120 },
                  { field: "movements", headerName: "Movements", width: 120 },
                ]}
                disableRowSelectionOnClick
              />
            </div>
          </div>
          <div className="card border">
            <div className="card-body pb-2">
              <div className="small fw-semibold">Movement Data Sheet (Detail)</div>
              <div className="small text-muted">
                {monthlyChartMonthFilter ? `Filtered month: ${monthlyChartMonthFilter}` : "All months"} • Rule: ORG Start MLE-{">"}any / DST End any-{">"}MLE = 1
              </div>
            </div>
            <div style={{ height: 420 }}>
              <DataGrid
                rows={monthlyMovementRows}
                columns={[
                  { field: "scheduleDate", headerName: "Date", width: 110 },
                  { field: "flightNo", headerName: "Flight", width: 100 },
                  { field: "aircraft", headerName: "AC", width: 90 },
                  { field: "origin", headerName: "ORI", width: 90 },
                  { field: "destination", headerName: "DEST", width: 90 },
                  { field: "route", headerName: "ORI-DEST", minWidth: 130, flex: 1 },
                  { field: "movementType", headerName: "Movement Type", minWidth: 210, flex: 1 },
                  { field: "movement", headerName: "Movement", width: 110 },
                ]}
                disableRowSelectionOnClick
              />
            </div>
          </div>
        </div>
      ) : null}

      {tab === "service" ? (
        <div className="d-flex flex-column gap-3">
          <div className="row g-2">
            {[
              {
                title: "Service Total",
                value: `${summaryMetrics.serviceTotal.toFixed(1)}h`,
                color: "rgb(37, 99, 235)",
                data: serviceSparkRows.map((r) => Number(r.total.toFixed(1))),
                axisId: "skpi-total",
              },
              {
                title: "Service Avg",
                value: `${summaryMetrics.serviceAvg.toFixed(2)}h`,
                color: "rgb(109, 40, 217)",
                data: serviceSparkRows.map((r) => Number(r.avg.toFixed(2))),
                axisId: "skpi-avg",
              },
              {
                title: "Avg AC / Day",
                value: summaryMetrics.avgAircraftPerDay.toFixed(2),
                color: "rgb(15, 118, 110)",
                data: serviceSparkRows.map((r) => r.avgAircraft),
                axisId: "skpi-ac",
              },
              {
                title: "Avg / Aircraft",
                value: `${summaryMetrics.avgPerAircraft.toFixed(1)}h`,
                color: "rgb(194, 65, 12)",
                data: serviceSparkRows.map((r) => Number(r.avgPerAircraft.toFixed(1))),
                axisId: "skpi-perac",
              },
            ].map((card) => (
              <div key={card.title} className="col-12 col-md-6 col-xl-3">
                <div className="card border h-100">
                  <div className="card-body p-3">
                    <Typography sx={{ color: "rgb(117,117,117)", fontWeight: 500, fontSize: "0.85rem", pb: 0.5 }}>
                      {serviceKpiIndex == null ? card.title : (serviceSparkLabels[serviceKpiIndex] || card.title)}
                    </Typography>
                    <div className="d-flex align-items-end justify-content-between gap-2">
                      <Typography sx={{ fontSize: "1.25rem", fontWeight: 600, color: card.color }}>
                        {serviceKpiIndex == null ? card.value : Number(card.data[serviceKpiIndex] || 0).toLocaleString()}
                      </Typography>
                      <Box sx={{ width: 170, height: 42 }}>
                        <SparkLineChart
                          height={42}
                          width={170}
                          area
                          showHighlight
                          color={card.color}
                          onHighlightedAxisChange={(axisItems) => setServiceKpiIndex(axisItems[0]?.dataIndex ?? null)}
                          highlightedAxis={serviceKpiIndex == null ? [] : [{ axisId: card.axisId, dataIndex: serviceKpiIndex }]}
                          {...buildSparkSettingsByLabels(card.data, serviceSparkLabels, card.axisId)}
                        />
                      </Box>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
          <div className="card border">
            <div className="card-body">
              <div className="d-flex align-items-center justify-content-between mb-2 gap-2 flex-wrap">
                <div className="small fw-semibold">Service Window Heatmap (Aircraft x Hour)</div>
                <div className="btn-group btn-group-sm">
                  <button type="button" className={`btn ${heatMode === "BOTH" ? "btn-success" : "btn-outline-secondary"}`} onClick={() => setHeatMode("BOTH")}>Both</button>
                  <button type="button" className={`btn ${heatMode === "DEP" ? "btn-success" : "btn-outline-secondary"}`} onClick={() => setHeatMode("DEP")}>Dep</button>
                  <button type="button" className={`btn ${heatMode === "ARR" ? "btn-success" : "btn-outline-secondary"}`} onClick={() => setHeatMode("ARR")}>Arr</button>
                </div>
              </div>
              <div className="table-responsive">
                <table
                  className="table table-sm align-middle mb-0"
                  style={{ fontSize: "0.58rem", lineHeight: 1.05 }}
                >
                  <thead>
                    <tr>
                      <th>AC \\ Hr</th>
                      {heatData.hours.map((h) => (
                        <th key={h} className="text-center">{String(h).padStart(2, "0")}:00</th>
                      ))}
                      <th className="text-end">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {heatData.rows.map((r) => (
                      <tr key={r.ac}>
                        <td className="fw-semibold">{r.ac}</td>
                        {heatData.hours.map((h) => {
                          const v = r.bins[h]
                          const ratio = heatData.max <= 0 ? 0 : v / heatData.max
                          let bg = "#e5e7eb"
                          let fg = "#334155"
                          if (v > 0 && ratio <= 0.35) bg = "#cfe7a6"
                          else if (ratio <= 0.55) bg = "#e7cd63"
                          else if (ratio <= 0.75) {
                            bg = "#e98034"
                            fg = "#ffffff"
                          } else if (ratio > 0.75) {
                            bg = "#b52e25"
                            fg = "#ffffff"
                          }
                          const active =
                            heatCellFilter.ac === r.ac && heatCellFilter.hour === h
                          return (
                            <td
                              key={`${r.ac}-${h}`}
                              className="text-center fw-semibold"
                              style={{
                                background: bg,
                                color: fg,
                                cursor: "pointer",
                                outline: active ? "2px solid #2563eb" : undefined,
                                outlineOffset: active ? "-2px" : undefined,
                              }}
                              onClick={() => {
                                setHeatCellFilter((prev) => {
                                  if (prev.ac === r.ac && prev.hour === h) return { ac: null, hour: null }
                                  return { ac: r.ac, hour: h }
                                })
                              }}
                            >
                              {v || ""}
                            </td>
                          )
                        })}
                        <td className="text-end fw-semibold">{r.total}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="d-flex align-items-center gap-3 mt-2 small text-muted flex-wrap">
                <span>Movements:</span>
                <span className="px-2 py-1 rounded" style={{ background: "#e5e7eb" }}>0</span>
                <span className="px-2 py-1 rounded" style={{ background: "#cfe7a6" }}>low</span>
                <span className="px-2 py-1 rounded" style={{ background: "#e7cd63" }}>mid</span>
                <span className="px-2 py-1 rounded text-white" style={{ background: "#e98034" }}>high</span>
                <span className="px-2 py-1 rounded text-white" style={{ background: "#b52e25" }}>peak</span>
              </div>
            </div>
          </div>
          <div className="card border">
            <div className="card-body">
              <div className="small fw-semibold mb-3">Combined Service Averages</div>
              <div className="row g-3">
                <div className="col-12 col-xl-6">
                  <div className="small text-muted mb-1">Monthly Avg Service Hrs</div>
                  {serviceMonthly.length === 0 ? (
                    <div className="small text-muted">No monthly data.</div>
                  ) : (
                    <LineChart
                      height={240}
                      xAxis={[{ scaleType: "point", data: serviceMonthly.map((m) => m.label), label: "Month" }]}
                      yAxis={[{ label: "Avg Hrs" }]}
                      series={[
                        {
                          data: serviceMonthly.map((m) => Number(m.avg.toFixed(2))),
                          label: "Monthly Avg Service Hrs",
                          color: "#d33d34",
                        },
                      ]}
                      margin={{ left: 54, right: 18, top: 20, bottom: 36 }}
                    />
                  )}
                </div>
                <div className="col-12 col-xl-6">
                  <div className="small text-muted mb-1">Average Service Hours by Aircraft</div>
                  {serviceRows.length === 0 ? (
                    <div className="small text-muted">No aircraft service data.</div>
                  ) : (
                    <LineChart
                      height={240}
                      xAxis={[{ scaleType: "point", data: serviceRows.slice(0, 12).map((r) => r.aircraft), label: "Aircraft" }]}
                      yAxis={[{ label: "Avg Hrs" }]}
                      series={[
                        { data: serviceRows.slice(0, 12).map((r) => Number(r.avg.toFixed(2))), label: "Avg Service Hours per Day", color: "#3b7a73" },
                      ]}
                      margin={{ left: 54, right: 18, top: 20, bottom: 36 }}
                    />
                  )}
                </div>
              </div>
            </div>
          </div>
          <div className="card border">
            <div className="card-body pb-2">
              <div className="d-flex align-items-center justify-content-between gap-2 mb-2">
                <div className="fw-semibold">Service Window Data</div>
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setHeatCellFilter({ ac: null, hour: null })}
                >
                  Clear Cell Filter
                </Button>
              </div>
              <div className="small text-muted mb-2">
                {heatCellFilter.ac && heatCellFilter.hour != null
                  ? `Filtered by cell: ${heatCellFilter.ac} @ ${String(heatCellFilter.hour).padStart(2, "0")}:00 (${heatMode})`
                  : `Showing all filtered records (${serviceDataRows.length.toLocaleString()})`}
              </div>
            </div>
            <div style={{ height: 440 }}>
              <DataGrid
                rows={serviceDataRows}
                columns={[
                  { field: "date", headerName: "Date", width: 120 },
                  { field: "ac", headerName: "AC", width: 100 },
                  { field: "flight", headerName: "Flight", width: 110 },
                  { field: "route", headerName: "Route", minWidth: 160, flex: 1 },
                  { field: "std", headerName: "STD", width: 90 },
                  { field: "sta", headerName: "STA", width: 90 },
                  { field: "blockHrs", headerName: "Block Hrs", width: 110 },
                ]}
                disableRowSelectionOnClick
              />
            </div>
          </div>
        </div>
      ) : null}

      {tab === "crewpool" ? (
        <div className="d-flex flex-column gap-3">
          <div className="card border">
            <div className="card-body">
              <div className="h5 mb-3">Crew Requirement Mapping</div>
              <div className="table-responsive">
                <table className="table table-sm align-middle mb-0">
                  <thead>
                    <tr>
                      <th>#</th>
                      <th>Requirement Item</th>
                      <th className="text-end">Value (Current Filter)</th>
                      {quarterList.map((q) => (
                        <th key={q} className="text-end">{q}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {([
                      { no: "1", label: "No of Aircraft Operational / Day (Avg)", valueFn: (m: ReturnType<typeof calcCrewMetrics>) => `${m.avgAircraftPerDay.toFixed(2)}` },
                      { no: "2", label: "Maximum Hours Available for Operational / Day", valueFn: () => `${crewModel.maxOpHours.toFixed(2)}h` },
                      { no: "3", label: "Maximum FDP allowed / Day", valueFn: () => `${crewModel.maxFdpHours.toFixed(2)}h` },
                      { no: "4", label: "Maximum FTL allowed / Day / Crew", valueFn: () => `${crewModel.maxFtlHours.toFixed(2)}h` },
                      { no: "5", label: "Daily Avg Schedule Block Hours / Day (Date range)", valueFn: (m: ReturnType<typeof calcCrewMetrics>) => `${m.avgBlockPerDay.toFixed(2)}h` },
                      { no: "6", label: "Crew set Required / Day (Aircraft Availability + FDP) = (Avg AC/Day × 13:30) ÷ 10:00", valueFn: (m: ReturnType<typeof calcCrewMetrics>) => `${m.reqByFdp.toFixed(2)}` },
                      { no: "7", label: "Crew set Required / Day (FTL) = Avg Block/Day ÷ 4:30", valueFn: (m: ReturnType<typeof calcCrewMetrics>) => `${m.reqByFtl.toFixed(2)}` },
                      { no: "8", label: "Avg Crew Set Required / Day (FTL & FDP)", valueFn: (m: ReturnType<typeof calcCrewMetrics>) => `${m.reqCore.toFixed(2)}` },
                      { no: "9", label: "Standby Set Required / Day", valueFn: () => `${crewModel.standby.toFixed(2)}` },
                      { no: "10", label: "Admin Duty / Chief Pilot / Training Captain", valueFn: () => `${crewModel.admin.toFixed(2)}`, editableKey: "admin" as const },
                      { no: "10A", label: "Release to Gan Ops", valueFn: () => `${crewModel.releaseGanOps.toFixed(2)}`, editableKey: "releaseGanOps" as const },
                      { no: "11", label: "Ground Training / Flight Training", valueFn: () => `${crewModel.training.toFixed(2)}` },
                      { no: "12", label: "Total Crew Set Required per Day", valueFn: (m: ReturnType<typeof calcCrewMetrics>) => `${m.totalPerDay.toFixed(2)}` },
                      { no: "13", label: "Avg Number of Crew Off per Day", valueFn: () => `${crewModel.crewOff.toFixed(2)}`, editableKey: "crewOff" as const },
                      { no: "14", label: "Avg Number of Crew on Leave / Day", valueFn: () => `${crewModel.leave.toFixed(2)}`, editableKey: "leave" as const },
                      { no: "15", label: "Total Number of Crew Requirement", valueFn: (m: ReturnType<typeof calcCrewMetrics>) => `${m.totalReq.toFixed(2)}` },
                      { no: "", label: "Number of Captain Available for Rostering", valueFn: () => "-" },
                    ] as Array<{ no: string; label: string; valueFn: (m: ReturnType<typeof calcCrewMetrics>) => string; editableKey?: "admin" | "releaseGanOps" | "crewOff" | "leave" }>).map(({ no, label, valueFn, editableKey }) => (
                      <tr key={`${no}-${label}`}>
                        <td>{no}</td>
                        <td>{label}</td>
                        <td className="text-end">
                          {editableKey ? (
                            <input
                              type="number"
                              step="0.1"
                              min="0"
                              value={crewModel[editableKey]}
                              onChange={(e) => {
                                const raw = Number(e.target.value)
                                const next = Number.isFinite(raw) && raw >= 0 ? raw : 0
                                setCrewModel((prev) => ({ ...prev, [editableKey]: next }))
                              }}
                              className="form-control form-control-sm ms-auto"
                              style={{ width: 110, textAlign: "right" }}
                            />
                          ) : (
                            valueFn(currentCrewMetrics)
                          )}
                        </td>
                        {quarterList.map((q) => (
                          <td key={`${no}-${q}`} className="text-end">
                            {valueFn(quarterCrewMetrics.get(q) || calcCrewMetrics([]))}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
          <div className="card border">
            <div className="card-body">
              <div className="h6 mb-2">Crew Ratio Justification Comparison</div>
              <div className="small text-muted mb-3">
                Report baseline from PDF vs current system values (active filters/date range).
              </div>
              <div className="table-responsive">
                <table className="table table-sm align-middle mb-0">
                  <thead>
                    <tr>
                      <th className="py-1">Metric</th>
                      <th className="py-1">Report Detail (Short)</th>
                      <th className="text-end py-1">Standard Value</th>
                      <th className="text-end py-1">System (Current Filter: {currentFilterRangeLabel})</th>
                      <th className="text-end py-1">Variance</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[
                      {
                        label: "Avg aircraft utilized per day",
                        detail:
                          "Daily demand baseline. Report states ~9.25 aircraft/day, with operational trend to 10 active aircraft.",
                        report: CREW_RATIO_SECTION_7_BASELINE.avgAircraftPerDay,
                        system: currentCrewMetrics.avgAircraftPerDay,
                      },
                      {
                        label: "Avg block hours per day",
                        detail:
                          "Daily flying workload baseline from actual utilization data. This drives FTL-based crew need.",
                        report: CREW_RATIO_SECTION_7_BASELINE.avgBlockPerDay,
                        system: currentCrewMetrics.avgBlockPerDay,
                      },
                      {
                        label: "Total FDP demand per day",
                        detail:
                          "FDP demand model: Avg aircraft/day × FDP allocated per aircraft (13.5h). Report uses this as daily FDP exposure.",
                        report: CREW_RATIO_SECTION_7_BASELINE.totalFdpDemandPerDay,
                        system: currentCrewDerived.totalFdpDemandPerDay,
                      },
                      {
                        label: "Crew required (FTL)",
                        detail:
                          "FTL constraint: Crew required = Avg Block Hours/Day ÷ 4.5h allowable avg per crew/day (to protect 28-day FTL limits).",
                        report: CREW_RATIO_SECTION_7_BASELINE.reqByFtl,
                        system: currentCrewMetrics.reqByFtl,
                      },
                      {
                        label: "Crew required (FDP)",
                        detail:
                          "FDP constraint: Crew required = Total FDP Demand/Day ÷ 10.5h sustainable FDP per crew/day. Report treats FDP as limiting factor.",
                        report: CREW_RATIO_SECTION_7_BASELINE.reqByFdp,
                        system: currentCrewMetrics.reqByFdp,
                      },
                      {
                        label: "Average flying crews required/day",
                        detail:
                          "Planning outcome for line flying coverage, balancing FTL and FDP constraints into a daily crew requirement.",
                        report: CREW_RATIO_SECTION_7_BASELINE.reqCore,
                        system: currentCrewMetrics.reqCore,
                      },
                      {
                        label: "Non-flying crews/day",
                        detail:
                          "Non-flying allocation: leave, recurrent off/rest, admin duties, and training/check obligations that remove line availability.",
                        report: CREW_RATIO_SECTION_7_BASELINE.nonFlyingPerDay,
                        system: currentCrewDerived.nonFlyingFromModel,
                        reportBreakdown: `Leave ${CREW_RATIO_SECTION_7_BASELINE.leave.toFixed(2)}, Off ${CREW_RATIO_SECTION_7_BASELINE.crewOff.toFixed(2)}, Admin ${CREW_RATIO_SECTION_7_BASELINE.admin.toFixed(2)}, Training ${CREW_RATIO_SECTION_7_BASELINE.training.toFixed(2)}, Gan ${Number(0).toFixed(2)}`,
                        systemBreakdown: `Leave ${crewModel.leave.toFixed(2)}, Off ${crewModel.crewOff.toFixed(2)}, Admin ${crewModel.admin.toFixed(2)}, Training ${crewModel.training.toFixed(2)}, Gan ${crewModel.releaseGanOps.toFixed(2)}`,
                      },
                      {
                        label: "Total crews required",
                        detail:
                          "Total requirement: flying crews needed + non-flying daily unavailability. This is the core staffing demand baseline.",
                        report: CREW_RATIO_SECTION_7_BASELINE.totalReq,
                        system: currentCrewMetrics.totalReq,
                      },
                      {
                        label: "Crew ratio per aircraft",
                        detail:
                          "Final ratio metric: Total Crews Required ÷ Avg Aircraft Utilized. Report conclusion band is 2.4–2.5 per aircraft.",
                        report:
                          CREW_RATIO_SECTION_7_BASELINE.avgAircraftPerDay > 0
                            ? CREW_RATIO_SECTION_7_BASELINE.totalReq / CREW_RATIO_SECTION_7_BASELINE.avgAircraftPerDay
                            : 0,
                        system: currentCrewDerived.ratio,
                      },
                    ].map((row) => {
                      const delta = row.system - row.report
                      const cls = delta > 0.05 ? "text-success" : delta < -0.05 ? "text-danger" : "text-muted"
                      return (
                        <tr key={row.label}>
                          <td className="py-1">{row.label}</td>
                          <td className="text-muted py-1" style={{ fontSize: "inherit" }}>
                            {row.detail}
                            {"reportBreakdown" in row ? (
                              <div className="mt-1" style={{ fontSize: "inherit" }}>
                                <div><span className="fw-semibold">Standard:</span> {row.reportBreakdown}</div>
                                <div><span className="fw-semibold">System:</span> {row.systemBreakdown}</div>
                              </div>
                            ) : null}
                          </td>
                          <td className="text-end py-1">
                            {row.report.toFixed(2)}
                          </td>
                          <td className="text-end py-1">
                            {row.system.toFixed(2)}
                          </td>
                          <td className={`text-end fw-semibold py-1 ${cls}`}>
                            {delta >= 0 ? "+" : ""}
                            {delta.toFixed(2)}
                          </td>
                        </tr>
                      )
                    })}
                    <tr>
                      <td className="py-1">Target ratio band</td>
                      <td className="text-end py-1">{CREW_RATIO_SECTION_7_BASELINE.crewRatioMin.toFixed(1)} - {CREW_RATIO_SECTION_7_BASELINE.crewRatioMax.toFixed(1)}</td>
                      <td className="text-end py-1">{currentCrewDerived.ratio.toFixed(2)}</td>
                      <td className="text-end text-muted py-1">reference</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </div>
      </div>
    </div>
  )
}
