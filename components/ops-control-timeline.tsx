"use client"

import { useEffect, useMemo, useRef, useState } from "react"
import {
  AlertTriangle,
  CheckCircle2,
  Clipboard,
  ClipboardCheck,
  Clock,
  Coffee,
  Eye,
  EyeOff,
  FileText,
  Plane,
  PlaneLanding,
  PlaneTakeoff,
  Trash2,
  Upload,
  Users,
  UsersRound,
} from "lucide-react"

// ---------------------------------------------------------------------------
// Public API: a single Entry-shaped flight record persisted in DataStore.
// Each row has the union of upload metadata + per-flight fields. Strings only
// (Entry from the parent app is `Record<string, string>`).
// ---------------------------------------------------------------------------
export type FlightEntry = {
  id: string
  createdAt: string
  scheduleDate: string
  aircraft: string
  flightNo: string
  origin: string
  destination: string
  depTime: string
  arrTime: string
  pic: string
  sic: string
  ca: string
  pax: string
  remarks: string
}

// ---------------------------------------------------------------------------
// SheetJS CDN loader. We avoid adding `xlsx` to package.json (it's no longer
// on npm registry) and instead lazily fetch the SheetJS standalone bundle the
// first time the user uploads a file. The script attaches `XLSX` to window.
// ---------------------------------------------------------------------------
const SHEETJS_CDN = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"

type SheetJSWindow = Window & {
  XLSX?: {
    read: (data: ArrayBuffer, opts: { type: "array" }) => {
      SheetNames: string[]
      Sheets: Record<
        string,
        { [cell: string]: { v?: unknown; w?: string; t?: string } } & {
          "!ref"?: string
        }
      >
    }
    utils: {
      sheet_to_json: (
        ws: unknown,
        opts?: {
          header?: number | "A"
          raw?: boolean
          defval?: unknown
        },
      ) => unknown[][]
      decode_range: (ref: string) => {
        s: { c: number; r: number }
        e: { c: number; r: number }
      }
      encode_cell: (a: { r: number; c: number }) => string
    }
  }
}

let sheetJsPromise: Promise<void> | null = null
function loadSheetJs(): Promise<void> {
  if (typeof window === "undefined") return Promise.reject(new Error("SSR"))
  const w = window as SheetJSWindow
  if (w.XLSX) return Promise.resolve()
  if (sheetJsPromise) return sheetJsPromise
  sheetJsPromise = new Promise<void>((resolve, reject) => {
    const existing = document.querySelector(
      `script[src="${SHEETJS_CDN}"]`,
    ) as HTMLScriptElement | null
    if (existing) {
      existing.addEventListener("load", () => resolve())
      existing.addEventListener("error", () =>
        reject(new Error("Failed to load SheetJS")),
      )
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

// ---------------------------------------------------------------------------
// Time helpers — keep everything in "HH:MM" + minute offsets.
// ---------------------------------------------------------------------------
function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n)
}

function formatTimeCell(value: unknown): string {
  if (value == null || value === "") return ""
  if (value instanceof Date) {
    return `${pad2(value.getHours())}:${pad2(value.getMinutes())}`
  }
  if (typeof value === "number") {
    // Excel time fraction (0-1) → hours/minutes
    if (value > 0 && value < 2) {
      const totalMin = Math.round(value * 24 * 60)
      const h = Math.floor(totalMin / 60) % 24
      const m = totalMin % 60
      return `${pad2(h)}:${pad2(m)}`
    }
    // Integer HHMM (e.g., 700 → 07:00, 1140 → 11:40)
    if (Number.isInteger(value) && value >= 0 && value <= 2400) {
      const h = Math.floor(value / 100)
      const m = value % 100
      if (h >= 0 && h < 24 && m >= 0 && m < 60) {
        return `${pad2(h)}:${pad2(m)}`
      }
    }
  }
  const s = String(value).trim()
  const m = s.match(/^(\d{1,2}):(\d{2})/)
  if (m) return `${pad2(Number(m[1]))}:${m[2]}`
  // Integer-shaped string like "700" or "1140"
  const intMatch = s.match(/^(\d{3,4})$/)
  if (intMatch) {
    const n = Number(intMatch[1])
    const h = Math.floor(n / 100)
    const mm = n % 100
    if (h >= 0 && h < 24 && mm >= 0 && mm < 60) {
      return `${pad2(h)}:${pad2(mm)}`
    }
  }
  return s
}

function parseTimeToMin(hhmm: string): number | null {
  if (!hhmm) return null
  const m = hhmm.match(/^(\d{1,2}):(\d{2})/)
  if (!m) return null
  return Number(m[1]) * 60 + Number(m[2])
}

function flightDurationMin(depTime: string, arrTime: string): number {
  const dep = parseTimeToMin(depTime)
  const arr = parseTimeToMin(arrTime)
  if (dep == null || arr == null) return 0
  if (arr >= dep) return arr - dep
  // Overnight arrival
  return 24 * 60 - dep + arr
}

function normalizeDateValueToYmd(value: unknown): string | null {
  if (value == null || value === "") return null
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${pad2(value.getMonth() + 1)}-${pad2(value.getDate())}`
  }
  if (typeof value === "number") {
    // Excel serial date (very rough guard)
    if (value > 20000 && value < 80000) {
      const ms = Math.round((value - 25569) * 86400 * 1000)
      const d = new Date(ms)
      if (!Number.isNaN(d.getTime())) {
        return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`
      }
    }
  }
  const s = String(value).trim()
  const ymd = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/)
  if (ymd) return `${ymd[1]}-${pad2(Number(ymd[2]))}-${pad2(Number(ymd[3]))}`
  const dmy = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/)
  if (dmy) {
    const yyyy = dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3]
    return `${yyyy}-${pad2(Number(dmy[2]))}-${pad2(Number(dmy[1]))}`
  }
  return null
}

function extractScheduleDate(rows: unknown[][]): string | null {
  const limitRows = Math.min(rows.length, 8)
  for (let r = 0; r < limitRows; r++) {
    const row = rows[r] || []
    for (let c = 0; c < row.length; c++) {
      const ymd = normalizeDateValueToYmd(row[c])
      if (ymd) return ymd
    }
  }
  return null
}

function crew4(value: string): string {
  return (value || "").trim().toUpperCase().slice(0, 4)
}

function normalizeFlightNo(value: string): string {
  return (value || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "")
}

function isSimilarFlightNo(a: string, b: string): boolean {
  const na = normalizeFlightNo(a)
  const nb = normalizeFlightNo(b)
  if (!na || !nb) return false
  if (na === nb) return true
  return na.includes(nb) || nb.includes(na)
}

function isMleLeg(f: FlightEntry): boolean {
  const o = (f.origin || "").trim().toUpperCase()
  const d = (f.destination || "").trim().toUpperCase()
  return o === "MLE" || d === "MLE"
}

// ---------------------------------------------------------------------------
// Excel parser. Two layouts are supported:
//
//   AAOC (preferred, one row per leg):
//     row N    → main headers (FLT NUMBER | AIRCRAFT REG | ROUTE | TIMING | ...)
//     row N+1  → sub-headers (FROM | TO | STD | STA | EMB | DEMB | ...)
//     row N+2+ → one row per flight leg
//
//   Legacy SCHEDULE (column-blocks per aircraft):
//     row 3    → aircraft tail #s
//     row 5    → block headers (FLIGHT | ROUTE | TIME | REMARKS | E | D | ...)
//     row 6+   → one row per 10-min slot
//
// We auto-detect by scanning the first 20 rows for "AIRCRAFT REG" or
// "FLT NUMBER".
// ---------------------------------------------------------------------------
function findHeaderColumn(row: unknown[], label: string): number {
  for (let c = 0; c < row.length; c++) {
    if (String(row[c] ?? "").trim().toUpperCase() === label) return c
  }
  return -1
}

function parseAaoc(rows: unknown[][]): {
  flights: Omit<FlightEntry, "id" | "createdAt" | "scheduleDate">[]
  warnings: string[]
} {
  const warnings: string[] = []

  // Locate main header row.
  let mainHeaderRow = -1
  let fltCol = -1
  let regCol = -1
  for (let r = 0; r < Math.min(rows.length, 25); r++) {
    const row = rows[r] || []
    const f = findHeaderColumn(row, "FLT NUMBER")
    const a = findHeaderColumn(row, "AIRCRAFT REG")
    if (f >= 0 && a >= 0) {
      mainHeaderRow = r
      fltCol = f
      regCol = a
      break
    }
  }
  if (mainHeaderRow < 0) {
    warnings.push(
      "Could not find 'FLT NUMBER' / 'AIRCRAFT REG' headers in the first 25 rows.",
    )
    return { flights: [], warnings }
  }

  const sub = (rows[mainHeaderRow + 1] || []) as unknown[]
  const fromCol = findHeaderColumn(sub, "FROM")
  const toCol = findHeaderColumn(sub, "TO")
  const stdCol = findHeaderColumn(sub, "STD")
  const staCol = findHeaderColumn(sub, "STA")
  const embCol = findHeaderColumn(sub, "EMB")
  const dembCol = findHeaderColumn(sub, "DEMB")
  if (fromCol < 0 || toCol < 0 || stdCol < 0 || staCol < 0) {
    warnings.push(
      "Sub-headers FROM / TO / STD / STA not found in the row below main headers.",
    )
    return { flights: [], warnings }
  }

  const main = (rows[mainHeaderRow] || []) as unknown[]
  const remarksCol = findHeaderColumn(main, "REMARKS")
  const counterCol = findHeaderColumn(main, "COUNTER CLOSE TIME")
  const picColMain = findHeaderColumn(main, "PIC")
  const sicColMain = findHeaderColumn(main, "SIC")
  // C/A column has slash in real files; tolerate both forms.
  const caColMain = (() => {
    const a = findHeaderColumn(main, "C/A")
    return a >= 0 ? a : findHeaderColumn(main, "CA")
  })()
  const picColSub = findHeaderColumn(sub, "PIC")
  const sicColSub = findHeaderColumn(sub, "SIC")
  const caColSub = (() => {
    const a = findHeaderColumn(sub, "C/A")
    return a >= 0 ? a : findHeaderColumn(sub, "CA")
  })()

  const flights: Omit<FlightEntry, "id" | "createdAt" | "scheduleDate">[] = []
  for (let r = mainHeaderRow + 2; r < rows.length; r++) {
    const row = rows[r] || []
    const flightNo = String(row[fltCol] ?? "").trim()
    const aircraft = String(row[regCol] ?? "").trim()
    if (!flightNo || !aircraft) continue
    if (flightNo.match(/SHUT|BREAK|HOLD/i)) continue
    const depTime = formatTimeCell(row[stdCol])
    const arrTime = formatTimeCell(row[staCol])
    if (!depTime || !arrTime) continue
    const origin = String(row[fromCol] ?? "").trim()
    const destination = String(row[toCol] ?? "").trim()
    const embVal = embCol >= 0 ? String(row[embCol] ?? "").trim() : ""
    const dembVal = dembCol >= 0 ? String(row[dembCol] ?? "").trim() : ""
    const pax = embVal || dembVal
    let remarks = ""
    if (remarksCol >= 0) remarks = String(row[remarksCol] ?? "").trim()
    if (!remarks && counterCol >= 0) {
      remarks = String(row[counterCol] ?? "").trim()
    }
    const pic =
      picColMain >= 0
        ? String(row[picColMain] ?? "").trim()
        : picColSub >= 0
        ? String(row[picColSub] ?? "").trim()
        : ""
    const sic =
      sicColMain >= 0
        ? String(row[sicColMain] ?? "").trim()
        : sicColSub >= 0
        ? String(row[sicColSub] ?? "").trim()
        : ""
    const ca =
      caColMain >= 0
        ? String(row[caColMain] ?? "").trim()
        : caColSub >= 0
        ? String(row[caColSub] ?? "").trim()
        : ""
    flights.push({
      aircraft,
      flightNo,
      origin,
      destination,
      depTime,
      arrTime,
      pic,
      sic,
      ca,
      pax,
      remarks,
    })
  }

  flights.sort((a, b) => {
    if (a.aircraft !== b.aircraft) return a.aircraft.localeCompare(b.aircraft)
    return (parseTimeToMin(a.depTime) ?? 0) - (parseTimeToMin(b.depTime) ?? 0)
  })

  return { flights, warnings }
}

function detectFormat(rows: unknown[][]): "aaoc" | "legacy" {
  for (let r = 0; r < Math.min(rows.length, 20); r++) {
    const row = rows[r] || []
    for (const cell of row) {
      const v = String(cell ?? "").trim().toUpperCase()
      if (v === "AIRCRAFT REG" || v === "FLT NUMBER") return "aaoc"
    }
  }
  return "legacy"
}

type RawLeg = {
  aircraft: string
  flightNo: string
  route: string
  time: string
  remarks: string
  e: string
  d: string
  rowIdx: number
}

function parseLegacySchedule(rows: unknown[][]): {
  flights: Omit<FlightEntry, "id" | "createdAt" | "scheduleDate">[]
  warnings: string[]
} {
  const warnings: string[] = []
  const headerRow = (rows[4] || []) as unknown[]
  const aircraftRow = (rows[2] || []) as unknown[]

  const flightCols: number[] = []
  for (let c = 0; c < headerRow.length; c++) {
    const val = String(headerRow[c] ?? "")
      .trim()
      .toUpperCase()
    if (val === "FLIGHT") flightCols.push(c)
  }
  if (flightCols.length === 0) {
    warnings.push(
      "Could not find any 'FLIGHT' column header in row 5 — is this the SCHEDULE sheet?",
    )
    return { flights: [], warnings }
  }

  const blocks = flightCols.map((startCol) => ({
    startCol,
    aircraft: String(aircraftRow[startCol] ?? "").trim(),
  }))

  const legs: RawLeg[] = []
  for (let r = 5; r < rows.length; r++) {
    const row = rows[r] || []
    for (const block of blocks) {
      const flightNo = row[block.startCol]
      if (flightNo == null || String(flightNo).trim() === "") continue
      const flightStr = String(flightNo).trim()
      // Skip non-flight markers like "DAY SHUT DOWN"
      if (flightStr.match(/SHUT|BREAK|HOLD/i)) continue
      legs.push({
        aircraft: block.aircraft,
        flightNo: flightStr,
        route: String(row[block.startCol + 1] ?? "").trim(),
        time: formatTimeCell(row[block.startCol + 2]),
        remarks: String(row[block.startCol + 3] ?? "").trim(),
        e: String(row[block.startCol + 4] ?? "").trim(),
        d: String(row[block.startCol + 5] ?? "").trim(),
        rowIdx: r,
      })
    }
  }

  // Pair depart/arrive by aircraft+flightNo (in row order).
  const grouped = new Map<string, RawLeg[]>()
  legs.forEach((leg) => {
    const key = `${leg.aircraft}::${leg.flightNo}`
    const list = grouped.get(key) ?? []
    list.push(leg)
    grouped.set(key, list)
  })

  const flights: Omit<FlightEntry, "id" | "createdAt" | "scheduleDate">[] = []
  grouped.forEach((entries) => {
    entries.sort((a, b) => a.rowIdx - b.rowIdx)
    if (entries.length >= 2) {
      // Pair sequential entries (legs round-trip same flight number repeatedly).
      for (let i = 0; i + 1 < entries.length; i += 2) {
        const dep = entries[i]
        const arr = entries[i + 1]
        flights.push({
          aircraft: dep.aircraft,
          flightNo: dep.flightNo,
          origin: dep.route,
          destination: arr.route,
          depTime: dep.time,
          arrTime: arr.time,
          pic: "",
          sic: "",
          ca: "",
          pax: dep.e || arr.d || "",
          remarks: dep.remarks || arr.remarks || "",
        })
      }
      if (entries.length % 2 === 1) {
        const trailing = entries[entries.length - 1]
        flights.push({
          aircraft: trailing.aircraft,
          flightNo: trailing.flightNo,
          origin: trailing.route,
          destination: "",
          depTime: trailing.time,
          arrTime: trailing.time,
          pic: "",
          sic: "",
          ca: "",
          pax: trailing.e || trailing.d || "",
          remarks: trailing.remarks,
        })
      }
    } else if (entries.length === 1) {
      const only = entries[0]
      flights.push({
        aircraft: only.aircraft,
        flightNo: only.flightNo,
        origin: only.route,
        destination: "",
        depTime: only.time,
        arrTime: only.time,
        pic: "",
        sic: "",
        ca: "",
        pax: only.e || only.d || "",
        remarks: only.remarks,
      })
    }
  })

  // Sort by aircraft tail then depart time.
  flights.sort((a, b) => {
    if (a.aircraft !== b.aircraft) return a.aircraft.localeCompare(b.aircraft)
    return (parseTimeToMin(a.depTime) ?? 0) - (parseTimeToMin(b.depTime) ?? 0)
  })

  return { flights, warnings }
}

function parseSchedule(rows: unknown[][]): {
  flights: Omit<FlightEntry, "id" | "createdAt" | "scheduleDate">[]
  warnings: string[]
} {
  const fmt = detectFormat(rows)
  return fmt === "aaoc" ? parseAaoc(rows) : parseLegacySchedule(rows)
}

async function parseExcelFile(file: File): Promise<{
  flights: Omit<FlightEntry, "id" | "createdAt" | "scheduleDate">[]
  warnings: string[]
  detectedScheduleDate: string | null
  detectedReleaseBy: string | null
  detectedReleaseTime: string | null
  detectedReleaseVersion: string | null
}> {
  await loadSheetJs()
  const w = window as SheetJSWindow
  if (!w.XLSX) throw new Error("SheetJS failed to load.")
  const buf = await file.arrayBuffer()
  const wb = w.XLSX.read(buf, { type: "array" })
  // Prefer a sheet called AAOC or SCHEDULE, otherwise first sheet.
  const upper = wb.SheetNames.map((n) => n.toUpperCase())
  const idx =
    upper.findIndex((n) => n.includes("AAOC")) >= 0
      ? upper.findIndex((n) => n.includes("AAOC"))
      : upper.findIndex((n) => n.includes("SCHEDULE")) >= 0
      ? upper.findIndex((n) => n.includes("SCHEDULE"))
      : 0
  const sheetName = wb.SheetNames[idx]
  if (!sheetName) throw new Error("No sheets found in workbook.")
  const sheet = wb.Sheets[sheetName]
  const rows = w.XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: "",
  }) as unknown[][]
  const parsed = parseSchedule(rows)
  const extractMetaValue = (keys: string[]): string | null => {
    const topRows = Math.min(rows.length, 16)
    for (let r = 0; r < topRows; r++) {
      const row = rows[r] || []
      for (let c = 0; c < row.length; c++) {
        const cell = String(row[c] ?? "").trim()
        if (!cell) continue
        const upper = cell.toUpperCase()
        const matched = keys.some((k) => upper.includes(k))
        if (!matched) continue
        const inline = cell.split(":").slice(1).join(":").trim()
        if (inline) return inline
        const right = String(row[c + 1] ?? "").trim()
        if (right) return right
      }
    }
    return null
  }
  const normalizeTimeMeta = (value: string | null): string | null => {
    if (!value) return null
    const t = formatTimeCell(value)
    if (parseTimeToMin(t) != null) return t
    const m = value.match(/\b(\d{1,2}):(\d{2})\b/)
    if (m) return `${pad2(Number(m[1]))}:${m[2]}`
    return null
  }
  return {
    ...parsed,
    detectedScheduleDate: extractScheduleDate(rows),
    detectedReleaseBy: extractMetaValue(["RELEASE BY", "RELEASED BY", "PREPARED BY"]),
    detectedReleaseTime: normalizeTimeMeta(
      extractMetaValue(["RELEASE TIME", "TIME RELEASED", "ISSUE TIME", "TIME"]),
    ),
    detectedReleaseVersion: extractMetaValue(["RELEASE VERSION", "VERSION", "REVISION", "REV"]),
  }
}

// ---------------------------------------------------------------------------
// Component.
// ---------------------------------------------------------------------------
export default function OpsControlTimeline(props: {
  flights: FlightEntry[]
  setFlights: (next: FlightEntry[]) => void
}) {
  const { flights: allFlights, setFlights } = props
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const [busy, setBusy] = useState(false)
  const [errorMsg, setErrorMsg] = useState<string | null>(null)
  const [warnings, setWarnings] = useState<string[]>([])
  const [lastImported, setLastImported] = useState<{
    fileName: string
    count: number
    at: string
    releaseVersion: string
    releaseBy: string
    releaseTime: string
    isInitial: boolean
  } | null>(null)
  const [isImportModalOpen, setIsImportModalOpen] = useState(false)
  const [scheduleDate, setScheduleDate] = useState<string>(() => {
    const d = new Date()
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`
  })
  // Scope Ops timeline computations to the selected day only.
  // This prevents heavy multi-day datasets from freezing the screen.
  const flights = useMemo(
    () => allFlights.filter((f) => (f.scheduleDate || "").trim() === scheduleDate),
    [allFlights, scheduleDate],
  )

  // Tick interval slider — discrete steps in minutes. Default 15.
  const INTERVAL_OPTIONS = [5, 10, 15, 30, 60]
  const [intervalIdx, setIntervalIdx] = useState(2)
  const tickIntervalMin = INTERVAL_OPTIONS[intervalIdx] ?? 15

  // Show / hide crew codes on flight bars.
  const [showCrew, setShowCrew] = useState(true)
  const [releaseVersion, setReleaseVersion] = useState("Initial")
  const [releaseBy, setReleaseBy] = useState("")
  const [releaseTime, setReleaseTime] = useState(() => {
    const d = new Date()
    return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`
  })
  const [isInitialSchedule, setIsInitialSchedule] = useState(false)

  // Drag-and-drop: when a bar is dropped we open a confirmation dialog
  // showing the target's existing schedule, the new STD/STA after the
  // horizontal drag delta, and any time-overlap conflicts. Cancelling
  // leaves the flight where it was.
  //   • mode "vertical"   — aircraft change. The whole trip group is
  //     reassigned together (each leg keeps its own STD/STA).
  //   • mode "horizontal" — only the time changed. Just the dragged
  //     segment moves; the rest of the trip stays put.
  const [moveDialog, setMoveDialog] = useState<{
    flightId: string
    fromAircraft: string
    toAircraft: string
    newStdMin: number
    newStaMin: number
    mode: "vertical" | "horizontal"
    groupFlightIds: string[]
  } | null>(null)
  const [draggingId, setDraggingId] = useState<string | null>(null)
  // X-offset of the cursor inside the bar at drag start, so dropping
  // computes new bar-left from the cursor without it leaping under the
  // pointer.
  const dragOffsetXRef = useRef(0)
  const timelineScrollRef = useRef<HTMLDivElement | null>(null)
  const lastCenteredKeyRef = useRef<string>("")

  // Live "now" indicator — refresh every 30s while the view is mounted.
  const [nowMin, setNowMin] = useState<number>(() => {
    const d = new Date()
    return d.getHours() * 60 + d.getMinutes()
  })
  useEffect(() => {
    const tick = () => {
      const d = new Date()
      setNowMin(d.getHours() * 60 + d.getMinutes())
    }
    tick()
    const id = window.setInterval(tick, 30_000)
    return () => window.clearInterval(id)
  }, [])

  // ----- handle file -----
  const handleFile = async (file: File) => {
    setBusy(true)
    setErrorMsg(null)
    setWarnings([])
    try {
      const {
        flights: parsed,
        warnings: warns,
        detectedScheduleDate,
        detectedReleaseBy,
        detectedReleaseTime,
        detectedReleaseVersion,
      } = await parseExcelFile(file)
      if (parsed.length === 0) {
        setErrorMsg(
          "No flights were parsed from the file. Check that the schedule sheet has the expected layout (FLIGHT | ROUTE | TIME | ... headers in row 5).",
        )
        return
      }
      if (detectedScheduleDate && detectedScheduleDate !== scheduleDate) {
        setErrorMsg(
          `Schedule date mismatch. Selected date is ${scheduleDate}, but file date is ${detectedScheduleDate}.`,
        )
        return
      }
      const resolvedReleaseBy = (detectedReleaseBy || releaseBy).trim()
      const resolvedReleaseTime = (detectedReleaseTime || releaseTime).trim()
      const resolvedReleaseVersion = (detectedReleaseVersion || releaseVersion || "Initial").trim()
      if (!resolvedReleaseBy || !resolvedReleaseTime) {
        setErrorMsg(
          "Could not find Release By/Release Time in sheet. Please fill them in the import popup and try again.",
        )
        return
      }
      const now = new Date().toISOString()
      const entries: FlightEntry[] = parsed.map((f, i) => ({
        id: `flt_${Date.now().toString(36)}_${i.toString(36)}`,
        createdAt: now,
        scheduleDate,
        ...f,
      }))
      setFlights(entries)
      setWarnings(warns)
      setLastImported({
        fileName: file.name,
        count: entries.length,
        at: new Date().toLocaleTimeString(),
        releaseVersion: resolvedReleaseVersion,
        releaseBy: resolvedReleaseBy,
        releaseTime: resolvedReleaseTime,
        isInitial: isInitialSchedule,
      })
      setReleaseBy(resolvedReleaseBy)
      setReleaseTime(resolvedReleaseTime)
      setReleaseVersion(resolvedReleaseVersion)
      setIsImportModalOpen(false)
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : String(err))
    } finally {
      setBusy(false)
    }
  }

  const onPickFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file) {
      const today = new Date()
      const todayKey = `${today.getFullYear()}-${pad2(today.getMonth() + 1)}-${pad2(today.getDate())}`
      if (scheduleDate > todayKey && !isInitialSchedule) {
        setErrorMsg("This is a future schedule date. Please mark it as Initial Schedule before importing.")
        e.currentTarget.value = ""
        return
      }
      setErrorMsg(null)
      void handleFile(file)
    }
    e.currentTarget.value = ""
  }

  const onDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    const file = e.dataTransfer?.files?.[0]
    if (file) {
      const today = new Date()
      const todayKey = `${today.getFullYear()}-${pad2(today.getMonth() + 1)}-${pad2(today.getDate())}`
      if (scheduleDate > todayKey && !isInitialSchedule) {
        setErrorMsg("This is a future schedule date. Please mark it as Initial Schedule before importing.")
        return
      }
      setErrorMsg(null)
      void handleFile(file)
    }
  }

  // ----- group + frame for timeline -----
  const aircraftList = useMemo(() => {
    const seen = new Set<string>()
    const list: string[] = []
    flights.forEach((f) => {
      if (!seen.has(f.aircraft)) {
        seen.add(f.aircraft)
        list.push(f.aircraft)
      }
    })
    return list.sort()
  }, [flights])

  const flightsByAircraft = useMemo(() => {
    const map = new Map<string, FlightEntry[]>()
    flights.forEach((f) => {
      const list = map.get(f.aircraft) ?? []
      list.push(f)
      map.set(f.aircraft, list)
    })
    map.forEach((list) =>
      list.sort(
        (a, b) =>
          (parseTimeToMin(a.depTime) ?? 0) - (parseTimeToMin(b.depTime) ?? 0),
      ),
    )
    return map
  }, [flights])

  // ----- Trip-group detection -----
  // Returns a Map<flightId, groupId>. Two flights share a groupId when they
  // belong to the same round-trip. Three patterns are supported:
  //   1. Same flight number on the same aircraft, multiple legs ("combined"
  //      flight, e.g. 7000 / 7000 making MLE→X→MLE).
  //   2. Pair of consecutive even/odd numbers on the same aircraft (e.g.
  //      2200 outbound + 2201 return, 2306 + 2307). Even is outbound, odd
  //      = even+1 is the return.
  //   3. Single segment (everything else, including odd-only legs from a
  //      previous-day layover like 2789, 2089, 2389).
  const tripGroups = useMemo(() => {
    const groups = new Map<string, string>()
    const numOf = (s: string): number | null => {
      const digits = (s || "").replace(/[^0-9]/g, "")
      if (!digits) return null
      const n = parseInt(digits, 10)
      return Number.isFinite(n) ? n : null
    }
    flights.forEach((f) => {
      if (groups.has(f.id)) return
      const num = numOf(f.flightNo)

      // Pattern 1: same-flight-no, same aircraft, more than one leg.
      const sameFn = flights.filter(
        (x) =>
          x.aircraft === f.aircraft &&
          isSimilarFlightNo(x.flightNo, f.flightNo),
      )
      if (sameFn.length > 1) {
        const gid = `same:${f.aircraft}:${normalizeFlightNo(f.flightNo)}`
        sameFn.forEach((x) => groups.set(x.id, gid))
        return
      }

      // Pattern 2: even/odd consecutive pair on same aircraft.
      if (num != null) {
        const partnerNum = num % 2 === 0 ? num + 1 : num - 1
        const partner = flights.find((x) => {
          if (x.aircraft !== f.aircraft) return false
          if (x.id === f.id) return false
          const xn = numOf(x.flightNo)
          return xn === partnerNum
        })
        if (partner) {
          const lower = Math.min(num, partnerNum)
          const gid = `pair:${f.aircraft}:${lower}`
          groups.set(f.id, gid)
          groups.set(partner.id, gid)
          return
        }
      }

      // Pattern 4: route-based fallback. If the flight is an MLE outbound
      // (MLE→X) and there's a later leg on the same aircraft returning
      // X→MLE, treat them as one trip even when their flight numbers
      // don't follow patterns 1 or 2. Catches custom number schemes.
      const o = (f.origin || "").trim().toUpperCase()
      const d = (f.destination || "").trim().toUpperCase()
      const fDep = parseTimeToMin(f.depTime)
      if (o === "MLE" && d && d !== "MLE" && fDep != null) {
        const partner = flights.find((x) => {
          if (x.aircraft !== f.aircraft) return false
          if (x.id === f.id) return false
          if (groups.has(x.id)) return false
          const xo = (x.origin || "").trim().toUpperCase()
          const xd = (x.destination || "").trim().toUpperCase()
          if (xo !== d || xd !== "MLE") return false
          const xDep = parseTimeToMin(x.depTime)
          return xDep != null && xDep >= fDep
        })
        if (partner) {
          const gid = `route:${f.aircraft}:${d}:${f.id}`
          groups.set(f.id, gid)
          groups.set(partner.id, gid)
          return
        }
      }

      // Pattern 3: single segment — group of one.
      groups.set(f.id, f.id)
    })
    return groups
  }, [flights])

  // ----- Schedule brief -----
  // Computes operational summary used by the Schedule Brief card.
  //   • mleDeps / mleArrs — counts of legs originating / terminating at MLE
  //   • totalPax — sum of pax across all legs
  //   • mleDepBreaks — fleet-wide gaps between consecutive MLE departures,
  //     a.k.a. windows when no flight is departing Male'.
  //   • routeConflicts — flights heading to the same outstation (not MLE)
  //     whose STD or STA falls within 30 min of another flight's STD or STA.
  const brief = useMemo(() => {
    const isMle = (s: string) => (s || "").trim().toUpperCase() === "MLE"

    const mleDeparturesArr = flights
      .filter((f) => isMle(f.origin))
      .map((f) => ({ f, depMin: parseTimeToMin(f.depTime) }))
      .filter((x): x is { f: FlightEntry; depMin: number } => x.depMin != null)
      .sort((a, b) => a.depMin - b.depMin)

    const mleArrivalsArr = flights
      .filter((f) => isMle(f.destination))
      .map((f) => ({ f, arrMin: parseTimeToMin(f.arrTime) }))
      .filter((x): x is { f: FlightEntry; arrMin: number } => x.arrMin != null)
      .sort((a, b) => a.arrMin - b.arrMin)

    const totalPax = flights.reduce((s, f) => {
      const n = parseInt((f.pax || "").replace(/[^0-9-]/g, ""), 10)
      return s + (Number.isFinite(n) ? n : 0)
    }, 0)

    // Free-window detection: gaps between consecutive MLE events. We compute
    // three views of the event timeline:
    //   • "dep"  — only STD-from-MLE events
    //   • "arr"  — only STA-to-MLE events
    //   • "both" — union of both, sorted by time
    // A free window is any gap ≥ 20 min between consecutive events.
    const MIN_BREAK_MIN = 20
    const computeBreaks = (events: number[]) => {
      const out: { startMin: number; endMin: number; durationMin: number }[] = []
      for (let i = 0; i < events.length - 1; i++) {
        const gap = events[i + 1] - events[i]
        if (gap >= MIN_BREAK_MIN) {
          out.push({
            startMin: events[i],
            endMin: events[i + 1],
            durationMin: gap,
          })
        }
      }
      return out
    }
    const depTimes = mleDeparturesArr.map((x) => x.depMin)
    const arrTimes = mleArrivalsArr.map((x) => x.arrMin)
    const bothTimes = [...depTimes, ...arrTimes].sort((a, b) => a - b)
    const breaksByMode = {
      dep: computeBreaks(depTimes),
      arr: computeBreaks(arrTimes),
      both: computeBreaks(bothTimes),
    }

    // Route conflicts — same non-MLE destination + STD or STA within 30 min.
    const TOL = 30
    const flightsWithTimes = flights.map((f) => ({
      f,
      dep: parseTimeToMin(f.depTime),
      arr: parseTimeToMin(f.arrTime),
    }))
    const routeConflicts: {
      a: FlightEntry
      b: FlightEntry
      reason: string
      destination: string
    }[] = []
    for (let i = 0; i < flightsWithTimes.length; i++) {
      for (let j = i + 1; j < flightsWithTimes.length; j++) {
        const A = flightsWithTimes[i]
        const B = flightsWithTimes[j]
        const destA = (A.f.destination || "").trim().toUpperCase()
        const destB = (B.f.destination || "").trim().toUpperCase()
        if (!destA || !destB) continue
        if (destA !== destB) continue
        if (destA === "MLE") continue
        const stdClose =
          A.dep != null && B.dep != null && Math.abs(A.dep - B.dep) <= TOL
        const staClose =
          A.arr != null && B.arr != null && Math.abs(A.arr - B.arr) <= TOL
        if (stdClose || staClose) {
          routeConflicts.push({
            a: A.f,
            b: B.f,
            destination: destA,
            reason:
              stdClose && staClose
                ? "STD and STA within 30 min"
                : stdClose
                ? "STD within 30 min"
                : "STA within 30 min",
          })
        }
      }
    }

    return {
      mleDeparturesCount: mleDeparturesArr.length,
      mleArrivalsCount: mleArrivalsArr.length,
      totalPax,
      breaksByMode,
      routeConflicts,
    }
  }, [flights])
  const avgServiceByAircraft = useMemo(() => {
    const map = new Map<string, { totalMin: number; legs: number }>()
    flights.forEach((f) => {
      const tail = (f.aircraft || "").trim()
      if (!tail) return
      const durMin = flightDurationMin(f.depTime, f.arrTime)
      const entry = map.get(tail) ?? { totalMin: 0, legs: 0 }
      entry.totalMin += Math.max(durMin, 0)
      entry.legs += 1
      map.set(tail, entry)
    })
    return [...map.entries()]
      .map(([tail, v]) => ({
        tail,
        avgHrs: v.legs > 0 ? v.totalMin / v.legs / 60 : 0,
        legs: v.legs,
      }))
      .sort((a, b) => b.avgHrs - a.avgHrs)
  }, [flights])

  // Local copy-state for the brief's clipboard button.
  const [briefCopied, setBriefCopied] = useState(false)
  // Which MLE event timeline drives the free-window list:
  //   • dep  — only STD-from-MLE events
  //   • arr  — only STA-to-MLE events
  //   • both — union of both
  const [breakMode, setBreakMode] = useState<"dep" | "arr" | "both">("dep")
  const [breakWindowEnabled, setBreakWindowEnabled] = useState(false)
  const [conflictWindowEnabled, setConflictWindowEnabled] = useState(false)

  // Horizontal pixel density. Aircraft are rows; time flows left → right.
  const PIXELS_PER_HOUR = useMemo(() => {
    switch (tickIntervalMin) {
      case 5:
        return 520
      case 10:
        return 360
      case 15:
        return 280
      case 30:
        return 190
      case 60:
      default:
        return 140
    }
  }, [tickIntervalMin])

  const routeConflictBands = useMemo(() => {
    const out: Array<{ startMin: number; endMin: number; label: string }> = []
    brief.routeConflicts.forEach((c) => {
      const aDep = parseTimeToMin(c.a.depTime)
      const aArr = parseTimeToMin(c.a.arrTime)
      const bDep = parseTimeToMin(c.b.depTime)
      const bArr = parseTimeToMin(c.b.arrTime)
      const mins = [aDep, aArr, bDep, bArr].filter((v): v is number => v != null)
      if (!mins.length) return
      const start = Math.max(Math.min(...mins) - 15, 0)
      const end = Math.min(Math.max(...mins) + 15, 24 * 60)
      out.push({ startMin: start, endMin: end, label: c.destination || "Conflict" })
    })
    out.sort((x, y) => x.startMin - y.startMin)
    const merged: Array<{ startMin: number; endMin: number; labels: Set<string> }> = []
    out.forEach((b) => {
      const last = merged[merged.length - 1]
      if (!last || b.startMin > last.endMin) {
        merged.push({ startMin: b.startMin, endMin: b.endMin, labels: new Set([b.label]) })
      } else {
        last.endMin = Math.max(last.endMin, b.endMin)
        last.labels.add(b.label)
      }
    })
    return merged.map((m) => ({ startMin: m.startMin, endMin: m.endMin, label: Array.from(m.labels).join(", ") }))
  }, [brief.routeConflicts])

  // Auto-fit time range to data, snapped to the chosen interval boundaries.
  const range = useMemo(() => {
    let min = Number.POSITIVE_INFINITY
    let max = Number.NEGATIVE_INFINITY
    flights.forEach((f) => {
      const dep = parseTimeToMin(f.depTime)
      const arr = parseTimeToMin(f.arrTime)
      if (dep != null) min = Math.min(min, dep)
      if (arr != null) max = Math.max(max, arr)
    })
    // Timeline always starts at 05:30 AM as requested.
    const fixedStartMin = 5 * 60 + 30
    if (!Number.isFinite(min) || !Number.isFinite(max)) {
      return { startMin: fixedStartMin, endMin: 22 * 60 }
    }
    const step = tickIntervalMin
    return {
      startMin: fixedStartMin,
      endMin: Math.ceil(max / step) * step,
    }
  }, [flights, tickIntervalMin])

  const totalMin = Math.max(tickIntervalMin, range.endMin - range.startMin)
  const trackWidth = (totalMin / 60) * PIXELS_PER_HOUR

  const ticks = useMemo(() => {
    const out: {
      min: number
      x: number
      label: string
      major: boolean
    }[] = []
    for (let t = range.startMin; t <= range.endMin; t += tickIntervalMin) {
      const major = t % 60 === 0
      out.push({
        min: t,
        x: ((t - range.startMin) / 60) * PIXELS_PER_HOUR,
        label: `${pad2(Math.floor(t / 60) % 24)}:${pad2(t % 60)}`,
        major,
      })
    }
    return out
  }, [range, tickIntervalMin, PIXELS_PER_HOUR])

  const nowX = useMemo(() => {
    if (nowMin < range.startMin || nowMin > range.endMin) return null
    return ((nowMin - range.startMin) / 60) * PIXELS_PER_HOUR
  }, [nowMin, range, PIXELS_PER_HOUR])

  const AIRCRAFT_COL_W = 130

  useEffect(() => {
    const scroller = timelineScrollRef.current
    if (!scroller) return
    if (nowX == null) return
    if (flights.length === 0) return

    const centerKey = `${scheduleDate}|${tickIntervalMin}|${range.startMin}|${range.endMin}|${flights.length}`
    if (lastCenteredKeyRef.current === centerKey) return

    const timelineViewportW = Math.max(0, scroller.clientWidth - AIRCRAFT_COL_W)
    const rawTarget = AIRCRAFT_COL_W + nowX - timelineViewportW / 2
    const maxLeft = Math.max(0, scroller.scrollWidth - scroller.clientWidth)
    const targetLeft = Math.max(0, Math.min(rawTarget, maxLeft))
    scroller.scrollLeft = targetLeft
    lastCenteredKeyRef.current = centerKey
  }, [
    AIRCRAFT_COL_W,
    flights.length,
    nowX,
    range.endMin,
    range.startMin,
    scheduleDate,
    tickIntervalMin,
  ])

  // Color a flight bar by tail # so the same aircraft stays one color. Palette
  // is wide enough that even large fleets get distinct tints before recycling.
  const colorForAircraft = (tail: string): { bg: string; fg: string; bd: string } => {
    const palette = [
      { bg: "#dbeafe", fg: "#1d4ed8", bd: "#93c5fd" },
      { bg: "#dcfce7", fg: "#15803d", bd: "#86efac" },
      { bg: "#ede9fe", fg: "#6d28d9", bd: "#c4b5fd" },
      { bg: "#fef3c7", fg: "#92400e", bd: "#fcd34d" },
      { bg: "#fee2e2", fg: "#b91c1c", bd: "#fca5a5" },
      { bg: "#cffafe", fg: "#0e7490", bd: "#67e8f9" },
      { bg: "#fce7f3", fg: "#be185d", bd: "#f9a8d4" },
      { bg: "#fef9c3", fg: "#854d0e", bd: "#fde68a" },
      { bg: "#e0f2fe", fg: "#075985", bd: "#7dd3fc" },
      { bg: "#f3e8ff", fg: "#7e22ce", bd: "#d8b4fe" },
      { bg: "#ffedd5", fg: "#9a3412", bd: "#fdba74" },
      { bg: "#d1fae5", fg: "#047857", bd: "#6ee7b7" },
      { bg: "#e0e7ff", fg: "#3730a3", bd: "#a5b4fc" },
      { bg: "#fae8ff", fg: "#86198f", bd: "#f0abfc" },
    ]
    let h = 0
    for (let i = 0; i < tail.length; i++) h = (h * 31 + tail.charCodeAt(i)) >>> 0
    return palette[h % palette.length]
  }

  return (
    <div className="d-flex flex-column gap-3">
      {/* ----- Upload card ----- */}
      <div className="card border-0 shadow-sm" style={{ borderRadius: 12 }}>
        <div className="card-body">
          <div className="d-flex align-items-center justify-content-between gap-2 flex-wrap">
            <h2 className="h6 fw-semibold mb-0 d-flex align-items-center gap-2">
              <span
                className="d-inline-flex align-items-center justify-content-center rounded"
                style={{
                  width: 28,
                  height: 28,
                  background: "#dbeafe",
                  color: "#1d4ed8",
                }}
              >
                <Upload size={15} />
              </span>
              Import flight schedule
            </h2>
            <button
              type="button"
              className="btn btn-primary btn-sm d-flex align-items-center gap-1"
              onClick={() => setIsImportModalOpen(true)}
            >
              <Upload size={14} />
              Import
            </button>
          </div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            className="d-none"
            onChange={onPickFile}
          />
          {errorMsg ? (
            <div className="alert alert-danger small d-flex align-items-start gap-2 mt-3 mb-0">
              <AlertTriangle size={16} className="mt-1" />
              <div>{errorMsg}</div>
            </div>
          ) : null}
          {warnings.length > 0 ? (
            <div className="alert alert-warning small d-flex align-items-start gap-2 mt-3 mb-0">
              <AlertTriangle size={16} className="mt-1" />
              <div>
                {warnings.map((w, i) => (
                  <div key={i}>{w}</div>
                ))}
              </div>
            </div>
          ) : null}
          {lastImported ? (
              <div className="alert alert-success small d-flex align-items-start gap-2 mt-3 mb-0">
              <CheckCircle2 size={16} className="mt-1" />
              <div>
                Imported {lastImported.count} flight legs from{" "}
                <strong>{lastImported.fileName}</strong> at {lastImported.at}. Release:{" "}
                <strong>{lastImported.releaseVersion}</strong> · {lastImported.releaseTime} ·{" "}
                {lastImported.releaseBy}
                {lastImported.isInitial ? " · Initial schedule" : ""}
              </div>
            </div>
          ) : null}
        </div>
      </div>

      {isImportModalOpen ? (
        <div
          className="modal show d-block"
          tabIndex={-1}
          style={{ background: "rgba(0,0,0,0.5)" }}
          onClick={() => setIsImportModalOpen(false)}
        >
          <div className="modal-dialog modal-dialog-centered" onClick={(e) => e.stopPropagation()}>
            <div className="modal-content border-0 shadow" style={{ borderRadius: 14 }}>
              <div className="modal-header border-0 pb-0">
                <h5 className="modal-title">Import flight schedule</h5>
                <button className="btn btn-sm btn-light" onClick={() => setIsImportModalOpen(false)}>
                  Close
                </button>
              </div>
              <div className="modal-body">
                <div className="d-flex align-items-center gap-2 mb-3">
                  <label className="form-label small fw-semibold mb-0 me-1">Schedule date</label>
                  <input
                    type="date"
                    className="form-control form-control-sm"
                    style={{ width: 170 }}
                    value={scheduleDate}
                    onChange={(e) => setScheduleDate(e.target.value)}
                  />
                </div>
                <div className="row g-2 mb-3">
                  <div className="col-sm-4">
                    <label className="form-label small fw-semibold mb-1">Release version</label>
                    <input
                      type="text"
                      className="form-control form-control-sm"
                      value={releaseVersion}
                      onChange={(e) => setReleaseVersion(e.target.value)}
                      placeholder="Initial / Rev1"
                    />
                  </div>
                  <div className="col-sm-4">
                    <label className="form-label small fw-semibold mb-1">Release time</label>
                    <input
                      type="time"
                      className="form-control form-control-sm"
                      value={releaseTime}
                      onChange={(e) => setReleaseTime(e.target.value)}
                    />
                  </div>
                  <div className="col-sm-4">
                    <label className="form-label small fw-semibold mb-1">Release by</label>
                    <input
                      type="text"
                      className="form-control form-control-sm"
                      value={releaseBy}
                      onChange={(e) => setReleaseBy(e.target.value)}
                      placeholder="Controller name"
                    />
                  </div>
                  <div className="col-12">
                    <label className="d-flex align-items-center gap-2 small fw-semibold mb-0">
                      <input
                        type="checkbox"
                        className="form-check-input mt-0"
                        checked={isInitialSchedule}
                        onChange={(e) => setIsInitialSchedule(e.target.checked)}
                      />
                      Initial schedule
                    </label>
                    <div className="text-muted small mt-1">
                      For future schedule dates (e.g., tomorrow), Initial schedule must be enabled before import.
                    </div>
                  </div>
                </div>
                <div
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={onDrop}
                  className="rounded p-3 d-flex flex-column align-items-center text-center gap-2"
                  style={{
                    border: "2px dashed var(--bs-border-color, #cbd5e1)",
                    background: "var(--bs-body-tertiary-bg, #f8fafc)",
                    cursor: busy ? "wait" : "pointer",
                    minHeight: 96,
                  }}
                  onClick={() => fileInputRef.current?.click()}
                >
                  <Upload size={20} color="#1d4ed8" />
                  <div className="fw-semibold">
                    {busy ? "Parsing…" : "Drop or click to upload .xlsx schedule"}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {/* ----- Schedule brief ----- */}
      {flights.length > 0 ? (() => {
        const fmtMin = (m: number) =>
          `${pad2(Math.floor(m / 60) % 24)}:${pad2(m % 60)}`
        const fmtDur = (m: number) => {
          const h = Math.floor(m / 60)
          const min = m % 60
          if (h && min) return `${h}h ${min}m`
          if (h) return `${h}h`
          return `${min}m`
        }
        const breaks = brief.breaksByMode[breakMode]
        const breakLabel =
          breakMode === "dep"
            ? "MLE departure-free windows"
            : breakMode === "arr"
            ? "MLE arrival-free windows"
            : "MLE event-free windows (dep + arr)"
        const buildText = (): string => {
          const lines: string[] = []
          lines.push(`Schedule brief — ${scheduleDate || "today"}`)
          lines.push("")
          lines.push(`Total legs: ${flights.length}`)
          lines.push(`STD from MLE: ${brief.mleDeparturesCount}`)
          lines.push(`STA to MLE:   ${brief.mleArrivalsCount}`)
          lines.push(`Total pax:    ${brief.totalPax}`)
          lines.push("")
          lines.push(`${breakLabel} (${breaks.length}):`)
          if (breaks.length === 0) {
            lines.push("  (none)")
          } else {
            breaks.forEach((b) =>
              lines.push(
                `  ${fmtMin(b.startMin)} → ${fmtMin(b.endMin)}  (${fmtDur(b.durationMin)})`,
              ),
            )
          }
          lines.push("")
          lines.push(`Route conflicts (${brief.routeConflicts.length}):`)
          if (brief.routeConflicts.length === 0) {
            lines.push("  (none)")
          } else {
            brief.routeConflicts.forEach((c) =>
              lines.push(
                `  ${c.destination}: ${c.a.flightNo} (${c.a.aircraft} ${c.a.depTime}/${c.a.arrTime}) ↔ ${c.b.flightNo} (${c.b.aircraft} ${c.b.depTime}/${c.b.arrTime}) — ${c.reason}`,
              ),
            )
          }
          return lines.join("\n")
        }
        const onCopy = async () => {
          try {
            await navigator.clipboard.writeText(buildText())
            setBriefCopied(true)
            window.setTimeout(() => setBriefCopied(false), 1500)
          } catch {
            // Clipboard may be denied — fall back to selection.
            const ta = document.createElement("textarea")
            ta.value = buildText()
            document.body.appendChild(ta)
            ta.select()
            try {
              document.execCommand("copy")
              setBriefCopied(true)
              window.setTimeout(() => setBriefCopied(false), 1500)
            } finally {
              document.body.removeChild(ta)
            }
          }
        }
        const stat = (
          icon: React.ReactNode,
          label: string,
          value: React.ReactNode,
          tint: { bg: string; fg: string },
        ) => (
          <div
            className="rounded p-3 d-flex align-items-center gap-3"
            style={{
              background: tint.bg,
              border: `1px solid ${tint.bg}`,
              minWidth: 0,
            }}
          >
            <span
              className="d-inline-flex align-items-center justify-content-center rounded"
              style={{
                width: 36,
                height: 36,
                background: "#fff",
                color: tint.fg,
                flex: "0 0 auto",
              }}
            >
              {icon}
            </span>
            <div style={{ minWidth: 0 }}>
              <div
                className="text-uppercase fw-semibold"
                style={{
                  fontSize: 10,
                  letterSpacing: 0.5,
                  color: tint.fg,
                  opacity: 0.85,
                }}
              >
                {label}
              </div>
              <div
                className="fw-bold"
                style={{ fontSize: 22, lineHeight: 1.1, color: tint.fg }}
              >
                {value}
              </div>
            </div>
          </div>
        )
        return (
          <div className="card border-0 shadow-sm" style={{ borderRadius: 12, order: 3 }}>
            <div className="card-body">
              <div className="d-flex align-items-center justify-content-between mb-3 flex-wrap gap-2">
                <h2 className="h6 fw-semibold mb-0 d-flex align-items-center gap-2">
                  <span
                    className="d-inline-flex align-items-center justify-content-center rounded"
                    style={{
                      width: 28,
                      height: 28,
                      background: "#dbeafe",
                      color: "#1d4ed8",
                    }}
                  >
                    <FileText size={15} />
                  </span>
                  Schedule brief
                </h2>
                <button
                  type="button"
                  className="btn btn-sm btn-outline-secondary d-flex align-items-center gap-1"
                  onClick={onCopy}
                  title="Copy brief as text"
                >
                  {briefCopied ? (
                    <>
                      <ClipboardCheck size={14} /> Copied
                    </>
                  ) : (
                    <>
                      <Clipboard size={14} /> Copy brief
                    </>
                  )}
                </button>
              </div>

              <div
                className="d-grid gap-2 mb-3"
                style={{
                  gridTemplateColumns:
                    "repeat(auto-fit, minmax(180px, 1fr))",
                }}
              >
                {stat(
                  <PlaneTakeoff size={18} />,
                  "STD from MLE",
                  brief.mleDeparturesCount,
                  { bg: "#dbeafe", fg: "#1d4ed8" },
                )}
                {stat(
                  <PlaneLanding size={18} />,
                  "STA to MLE",
                  brief.mleArrivalsCount,
                  { bg: "#dcfce7", fg: "#15803d" },
                )}
                {stat(
                  <UsersRound size={18} />,
                  "Total pax",
                  brief.totalPax.toLocaleString(),
                  { bg: "#ede9fe", fg: "#6d28d9" },
                )}
                {stat(
                  <AlertTriangle size={18} />,
                  "Route conflicts",
                  brief.routeConflicts.length,
                  brief.routeConflicts.length > 0
                    ? { bg: "#fee2e2", fg: "#b91c1c" }
                    : { bg: "#f1f5f9", fg: "#475569" },
                )}
              </div>

              <div className="row g-3">
                {/* Break Window */}
                <div className="col-md-6">
                  <div className="d-flex align-items-center justify-content-between mb-2 flex-wrap gap-2">
                    <div className="fw-semibold small d-flex align-items-center gap-2">
                      <Coffee size={14} color="#92400e" />
                      Break Window
                    </div>
                    <div className="d-flex align-items-center gap-2">
                      <div className="btn-group btn-group-sm" role="group" aria-label="Conflict window toggle">
                        <button
                          type="button"
                          className={`btn ${conflictWindowEnabled ? "btn-danger" : "btn-outline-secondary"}`}
                          onClick={() => setConflictWindowEnabled(true)}
                          title="Show conflict bands on Daily Ops Timeline"
                        >
                          Conflict On
                        </button>
                        <button
                          type="button"
                          className={`btn ${!conflictWindowEnabled ? "btn-secondary" : "btn-outline-secondary"}`}
                          onClick={() => setConflictWindowEnabled(false)}
                          title="Hide conflict bands from Daily Ops Timeline"
                        >
                          Conflict Off
                        </button>
                      </div>
                      <div className="btn-group btn-group-sm" role="group" aria-label="Break window toggle">
                        <button
                          type="button"
                          className={`btn ${breakWindowEnabled ? "btn-success" : "btn-outline-secondary"}`}
                          onClick={() => setBreakWindowEnabled(true)}
                          title="Show break window on Daily Ops Timeline"
                        >
                          On
                        </button>
                        <button
                          type="button"
                          className={`btn ${!breakWindowEnabled ? "btn-secondary" : "btn-outline-secondary"}`}
                          onClick={() => setBreakWindowEnabled(false)}
                          title="Hide break window from Daily Ops Timeline"
                        >
                          Off
                        </button>
                      </div>
                      <div
                        className="btn-group btn-group-sm"
                        role="group"
                        aria-label="MLE event window mode"
                      >
                        <button
                          type="button"
                          className={`btn ${
                            breakMode === "dep"
                              ? "btn-primary"
                              : "btn-outline-secondary"
                          } d-flex align-items-center gap-1`}
                          onClick={() => setBreakMode("dep")}
                          title="Gaps between MLE departures (STD from MLE)"
                        >
                          <PlaneTakeoff size={12} /> Dep
                        </button>
                        <button
                          type="button"
                          className={`btn ${
                            breakMode === "arr"
                              ? "btn-primary"
                              : "btn-outline-secondary"
                          } d-flex align-items-center gap-1`}
                          onClick={() => setBreakMode("arr")}
                          title="Gaps between MLE arrivals (STA to MLE)"
                        >
                          <PlaneLanding size={12} /> Arr
                        </button>
                        <button
                          type="button"
                          className={`btn ${
                            breakMode === "both"
                              ? "btn-primary"
                              : "btn-outline-secondary"
                          }`}
                          onClick={() => setBreakMode("both")}
                          title="Gaps with neither departure nor arrival on MLE"
                        >
                          Both
                        </button>
                      </div>
                      <span className="badge bg-warning-subtle text-warning-emphasis">
                        {breaks.length}
                      </span>
                    </div>
                  </div>
                  {breaks.length === 0 ? (
                    <div className="text-muted small fst-italic">
                      {breakMode === "dep"
                        ? "No gaps ≥ 20 min between MLE departures."
                        : breakMode === "arr"
                        ? "No gaps ≥ 20 min between MLE arrivals."
                        : "No gaps ≥ 20 min between MLE departures or arrivals."}
                    </div>
                  ) : (
                    <div
                      className="rounded"
                      style={{
                        border: "1px solid var(--bs-border-color, #e2e8f0)",
                        maxHeight: 220,
                        overflow: "auto",
                      }}
                    >
                      <table
                        className="table table-sm align-middle mb-0"
                        style={{ fontSize: 12 }}
                      >
                        <thead>
                          <tr className="text-uppercase small text-muted">
                            <th>Window</th>
                            <th className="text-end">Duration</th>
                          </tr>
                        </thead>
                        <tbody>
                          {breaks.map((b, i) => (
                            <tr key={i}>
                              <td className="fw-semibold">
                                {fmtMin(b.startMin)} → {fmtMin(b.endMin)}
                              </td>
                              <td className="text-end">
                                <span
                                  className="badge"
                                  style={{
                                    background: "#fef3c7",
                                    color: "#92400e",
                                  }}
                                >
                                  {fmtDur(b.durationMin)}
                                </span>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>

                {/* Average service hrs by aircraft (moved to right side) */}
                <div className="col-md-6">
                  <div className="d-flex align-items-center justify-content-between mb-2 flex-wrap gap-2">
                    <div className="fw-semibold small d-flex align-items-center gap-2">
                      <Clock size={14} color="#1d4ed8" />
                      Average service hrs by aircraft
                    </div>
                    <span className="badge bg-primary-subtle text-primary-emphasis">
                      {avgServiceByAircraft.length}
                    </span>
                  </div>
                  {avgServiceByAircraft.length === 0 ? (
                    <div className="text-muted small fst-italic">
                      No aircraft timing data available.
                    </div>
                  ) : (
                    <div
                      className="rounded"
                      style={{
                        border: "1px solid var(--bs-border-color, #e2e8f0)",
                        maxHeight: 220,
                        overflow: "auto",
                      }}
                    >
                      <table
                        className="table table-sm align-middle mb-0"
                        style={{ fontSize: 12 }}
                      >
                        <thead>
                          <tr className="text-uppercase small text-muted">
                            <th>Aircraft</th>
                            <th className="text-end">Avg hrs</th>
                            <th className="text-end">Legs</th>
                          </tr>
                        </thead>
                        <tbody>
                          {avgServiceByAircraft.map((row) => (
                            <tr key={row.tail}>
                              <td className="fw-semibold">{row.tail}</td>
                              <td className="text-end">{row.avgHrs.toFixed(2)}</td>
                              <td className="text-end">{row.legs}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </div>
              <div className="row g-3 mt-1">
                <div className="col-12">
                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <div className="fw-semibold small d-flex align-items-center gap-2">
                      <AlertTriangle size={14} color="#b91c1c" />
                      Route conflicts (≤30 min, non-MLE)
                    </div>
                    <span
                      className={`badge ${
                        brief.routeConflicts.length > 0
                          ? "bg-danger-subtle text-danger-emphasis"
                          : "bg-success-subtle text-success-emphasis"
                      }`}
                    >
                      {brief.routeConflicts.length}
                    </span>
                  </div>
                  {brief.routeConflicts.length === 0 ? (
                    <div className="text-muted small fst-italic">
                      No flights to the same outstation within 30 min.
                    </div>
                  ) : (
                    <div
                      className="rounded"
                      style={{
                        border: "1px solid var(--bs-border-color, #e2e8f0)",
                        maxHeight: 220,
                        overflow: "auto",
                      }}
                    >
                      <table
                        className="table table-sm align-middle mb-0"
                        style={{ fontSize: 12 }}
                      >
                        <thead>
                          <tr className="text-uppercase small text-muted">
                            <th>Dest</th>
                            <th>Flight A</th>
                            <th>Flight B</th>
                            <th>Reason</th>
                          </tr>
                        </thead>
                        <tbody>
                          {brief.routeConflicts.map((c, i) => (
                            <tr key={i} style={{ background: "#fef2f2" }}>
                              <td className="fw-semibold">{c.destination}</td>
                              <td>
                                <div className="fw-semibold">{c.a.flightNo}</div>
                                <div className="text-muted small">
                                  {c.a.aircraft} · {c.a.depTime}/{c.a.arrTime}
                                </div>
                              </td>
                              <td>
                                <div className="fw-semibold">{c.b.flightNo}</div>
                                <div className="text-muted small">
                                  {c.b.aircraft} · {c.b.depTime}/{c.b.arrTime}
                                </div>
                              </td>
                              <td className="small text-muted">{c.reason}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        )
      })() : null}

      {/* ----- Timeline ----- */}
      <div className="card border-0 shadow-sm" style={{ borderRadius: 12, order: 2 }}>
        <div className="card-body">
          <div className="d-flex align-items-center justify-content-between mb-3 flex-wrap gap-2">
            <h2 className="h6 fw-semibold mb-0 d-flex align-items-center gap-2">
              <span
                className="d-inline-flex align-items-center justify-content-center rounded"
                style={{
                  width: 28,
                  height: 28,
                  background: "#dbeafe",
                  color: "#1d4ed8",
                }}
              >
                <Plane size={15} />
              </span>
              Daily ops timeline
              {flights.length > 0 ? (
                <span className="badge bg-primary-subtle text-primary-emphasis ms-2">
                  {flights.length} legs · {aircraftList.length} aircraft
                </span>
              ) : null}
            </h2>
            <div className="d-flex align-items-center gap-3 flex-wrap">
              <div className="d-flex align-items-center gap-2">
                <label
                  className="form-label small fw-semibold mb-0"
                  htmlFor="ops-tick-interval"
                >
                  Interval
                </label>
                <input
                  id="ops-tick-interval"
                  type="range"
                  className="form-range"
                  min={0}
                  max={INTERVAL_OPTIONS.length - 1}
                  step={1}
                  value={intervalIdx}
                  onChange={(e) => setIntervalIdx(Number(e.target.value))}
                  style={{ width: 130 }}
                />
                <span
                  className="badge"
                  style={{
                    background: "#dbeafe",
                    color: "#1d4ed8",
                    fontWeight: 600,
                    minWidth: 56,
                  }}
                >
                  {tickIntervalMin} min
                </span>
              </div>
              <button
                type="button"
                className={`btn btn-sm d-flex align-items-center gap-1 ${
                  showCrew ? "btn-primary" : "btn-outline-secondary"
                }`}
                onClick={() => setShowCrew((v) => !v)}
                title={showCrew ? "Hide crew codes" : "Show crew codes"}
              >
                {showCrew ? <Eye size={14} /> : <EyeOff size={14} />}
                <Users size={14} />
                {showCrew ? "Crew on" : "Crew off"}
              </button>
              {flights.length > 0 ? (
                <button
                  className="btn btn-sm btn-outline-danger d-flex align-items-center gap-1"
                  onClick={() => {
                    if (
                      window.confirm(
                        "Clear all imported flights? This removes them from the local store.",
                      )
                    ) {
                      setFlights([])
                      setLastImported(null)
                    }
                  }}
                >
                  <Trash2 size={14} />
                  Clear
                </button>
              ) : null}
            </div>
          </div>

          {flights.length === 0 ? (
            <div className="text-muted small text-center py-5">
              No schedule loaded yet. Upload a flight schedule to populate the
              timeline.
            </div>
          ) : (
            <div
              ref={timelineScrollRef}
              style={{
                border: "1px solid var(--bs-border-color, #e2e8f0)",
                borderRadius: 8,
                overflow: "auto",
                maxHeight: 640,
                position: "relative",
              }}
            >
              {(() => {
                const HEADER_H = 40
                const ROW_H = 60
                const totalGridW = AIRCRAFT_COL_W + trackWidth
                return (
                  <div style={{ position: "relative", width: totalGridW }}>
                    {/* ===== Header row: aircraft-corner + time tick header ===== */}
                    <div
                      style={{
                        display: "flex",
                        position: "sticky",
                        top: 0,
                        zIndex: 4,
                        background: "var(--bs-body-tertiary-bg, #f8fafc)",
                        borderBottom:
                          "1px solid var(--bs-border-color, #e2e8f0)",
                      }}
                    >
                      <div
                        style={{
                          width: AIRCRAFT_COL_W,
                          flex: "0 0 auto",
                          height: HEADER_H,
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          fontSize: 11,
                          fontWeight: 600,
                          color: "var(--bs-secondary-color)",
                          letterSpacing: 0.5,
                          textTransform: "uppercase",
                          borderRight:
                            "1px solid var(--bs-border-color, #e2e8f0)",
                          position: "sticky",
                          left: 0,
                          background: "var(--bs-body-tertiary-bg, #f8fafc)",
                          zIndex: 5,
                        }}
                      >
                        Aircraft
                      </div>
                      <div
                        style={{
                          width: trackWidth,
                          flex: "0 0 auto",
                          height: HEADER_H,
                          position: "relative",
                        }}
                      >
                        {ticks.map((t) => (
                          <div
                            key={t.min}
                            style={{
                              position: "absolute",
                              left: t.x,
                              top: 0,
                              height: "100%",
                              padding: "0 6px",
                              fontSize: t.major ? 11 : 10,
                              fontWeight: t.major ? 600 : 500,
                              color: t.major
                                ? "var(--bs-secondary-color)"
                                : "var(--bs-tertiary-color, #94a3b8)",
                              display: "flex",
                              alignItems: "center",
                              borderLeft:
                                t.x === 0
                                  ? "none"
                                  : t.major
                                  ? "1px solid var(--bs-border-color, #cbd5e1)"
                                  : "1px dashed var(--bs-border-color, #e2e8f0)",
                            }}
                          >
                            {t.label}
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* ===== Body: aircraft labels (sticky left) + flight rows ===== */}
                    {aircraftList.map((tail) => {
                      const lanes = flightsByAircraft.get(tail) ?? []
                      const c = colorForAircraft(tail)
                      return (
                        <div
                          key={tail}
                          style={{
                            display: "flex",
                            borderBottom:
                              "1px solid var(--bs-border-color, #e2e8f0)",
                          }}
                        >
                          {/* Aircraft label cell — sticky left */}
                          <div
                            style={{
                              width: AIRCRAFT_COL_W,
                              flex: "0 0 auto",
                              height: ROW_H,
                              display: "flex",
                              alignItems: "center",
                              padding: "0 12px",
                              gap: 8,
                              position: "sticky",
                              left: 0,
                              background:
                                "var(--bs-body-tertiary-bg, #f8fafc)",
                              borderRight:
                                "1px solid var(--bs-border-color, #e2e8f0)",
                              zIndex: 3,
                            }}
                          >
                            <span
                              className="d-inline-flex align-items-center justify-content-center rounded-circle"
                              style={{
                                width: 28,
                                height: 28,
                                background: c.bg,
                                color: c.fg,
                                flex: "0 0 auto",
                              }}
                            >
                              <Plane size={14} />
                            </span>
                            <div
                              className="fw-bold text-truncate"
                              style={{ fontSize: 13 }}
                            >
                              {tail}
                            </div>
                          </div>
                          {/* Aircraft lane — flight bars laid out horizontally */}
                          <div
                            onDragOver={(e) => {
                              if (draggingId) {
                                e.preventDefault()
                                e.dataTransfer.dropEffect = "move"
                              }
                            }}
                            onDrop={(e) => {
                              e.preventDefault()
                              const id =
                                e.dataTransfer.getData("text/plain") || draggingId
                              setDraggingId(null)
                              if (!id) return
                              const moving = flights.find((x) => x.id === id)
                              if (!moving) return

                              // Compute new STD from drop X within the lane.
                              // Subtract the drag-start cursor offset so the
                              // bar's left edge tracks where the user grabbed
                              // it, then snap to the current tick interval.
                              const laneRect = (
                                e.currentTarget as HTMLDivElement
                              ).getBoundingClientRect()
                              const dropLeftPx =
                                e.clientX - laneRect.left - dragOffsetXRef.current
                              const rawMin =
                                (dropLeftPx / PIXELS_PER_HOUR) * 60 + range.startMin
                              const snap = tickIntervalMin
                              const oldDep =
                                parseTimeToMin(moving.depTime) ?? range.startMin
                              const oldArr =
                                parseTimeToMin(moving.arrTime) ?? oldDep
                              const dur = Math.max(0, oldArr - oldDep)
                              const clampedMin = Math.max(0, Math.min(rawMin, 24 * 60 - dur))
                              const newStdMin = Math.round(clampedMin / snap) * snap
                              const newStaMin = newStdMin + dur

                              const sameAircraft = moving.aircraft === tail
                              const sameTime = newStdMin === oldDep
                              if (sameAircraft && sameTime) return

                              // Aircraft change → whole trip group moves
                              // together (each leg keeps its own time).
                              // Same aircraft, time change → only this
                              // segment moves to the new STD/STA.
                              if (!sameAircraft) {
                                const gid = tripGroups.get(id) ?? id
                                const groupFlightIds = flights
                                  .filter(
                                    (x) => (tripGroups.get(x.id) ?? x.id) === gid,
                                  )
                                  .map((x) => x.id)
                                setMoveDialog({
                                  flightId: id,
                                  fromAircraft: moving.aircraft,
                                  toAircraft: tail,
                                  newStdMin: oldDep,
                                  newStaMin: oldArr,
                                  mode: "vertical",
                                  groupFlightIds,
                                })
                              } else {
                                setMoveDialog({
                                  flightId: id,
                                  fromAircraft: moving.aircraft,
                                  toAircraft: tail,
                                  newStdMin,
                                  newStaMin,
                                  mode: "horizontal",
                                  groupFlightIds: [id],
                                })
                              }
                            }}
                            style={{
                              width: trackWidth,
                              flex: "0 0 auto",
                              height: ROW_H,
                              position: "relative",
                              background:
                                draggingId &&
                                flights.find((x) => x.id === draggingId)?.aircraft !==
                                  tail
                                  ? "rgba(29,78,216,0.08)"
                                  : "transparent",
                              transition: "background 120ms",
                            }}
                          >
                            {/* vertical interval grid */}
                            {ticks.map((t) => (
                              <div
                                key={t.min}
                                style={{
                                  position: "absolute",
                                  left: t.x,
                                  top: 0,
                                  bottom: 0,
                                  width: 1,
                                  background:
                                    t.x === 0
                                      ? "transparent"
                                      : "var(--bs-border-color, #e2e8f0)",
                                  opacity: t.major ? 0.7 : 0.35,
                                }}
                              />
                            ))}
                            {/* flight bars (horizontal) */}
                            {(() => {
                              const bars = lanes
                                .map((f) => {
                                  const dep = parseTimeToMin(f.depTime)
                                  if (dep == null) return null
                                  const durMin = flightDurationMin(f.depTime, f.arrTime)
                                  const startMinAbs = dep
                                  const endMinAbs = dep + Math.max(durMin, 1)
                                  const left =
                                    ((startMinAbs - range.startMin) / 60) *
                                    PIXELS_PER_HOUR
                                  const right =
                                    ((endMinAbs - range.startMin) / 60) *
                                    PIXELS_PER_HOUR
                                  const width = Math.max(26, right - left)
                                  return { f, left, width }
                                })
                                .filter((x): x is { f: FlightEntry; left: number; width: number } => Boolean(x))

                              return (
                                <>
                                  {/* Trip-group connector: bars in the same
                                      group (same flight number across legs,
                                      or even/odd pair) on this aircraft are
                                      stitched together with a single thin
                                      line at the bottom. */}
                                  {(() => {
                                    const out: React.ReactNode[] = []
                                    const byGroup = new Map<string, typeof bars>()
                                    bars.forEach((b) => {
                                      const gid = tripGroups.get(b.f.id) ?? b.f.id
                                      const list = byGroup.get(gid) ?? []
                                      list.push(b)
                                      byGroup.set(gid, list)
                                    })
                                    byGroup.forEach((list, gid) => {
                                      if (list.length < 2) return
                                      const sorted = [...list].sort(
                                        (a, b) => a.left - b.left,
                                      )
                                      for (let i = 0; i < sorted.length - 1; i++) {
                                        const A = sorted[i]
                                        const B = sorted[i + 1]
                                        const x1 = A.left + A.width
                                        const x2 = B.left
                                        if (x2 <= x1) continue
                                        // Bold connector — thicker stroke,
                                        // foreground-tint colour, and dot
                                        // caps at each end so the link
                                        // between segments is unmistakable.
                                        out.push(
                                          <div
                                            key={`conn-${gid}-${i}-line`}
                                            style={{
                                              position: "absolute",
                                              left: x1,
                                              top: ROW_H - 9,
                                              width: x2 - x1,
                                              height: 3,
                                              background: c.fg,
                                              borderRadius: 2,
                                              opacity: 1,
                                            }}
                                          />,
                                        )
                                        // Round dot caps anchor the line to
                                        // each bar so the connector reads as
                                        // a deliberate stitch, not a stray
                                        // grid line.
                                        out.push(
                                          <div
                                            key={`conn-${gid}-${i}-capL`}
                                            style={{
                                              position: "absolute",
                                              left: x1 - 3,
                                              top: ROW_H - 12,
                                              width: 7,
                                              height: 7,
                                              background: c.fg,
                                              borderRadius: "50%",
                                              opacity: 1,
                                            }}
                                          />,
                                        )
                                        out.push(
                                          <div
                                            key={`conn-${gid}-${i}-capR`}
                                            style={{
                                              position: "absolute",
                                              left: x2 - 4,
                                              top: ROW_H - 12,
                                              width: 7,
                                              height: 7,
                                              background: c.fg,
                                              borderRadius: "50%",
                                              opacity: 1,
                                            }}
                                          />,
                                        )
                                      }
                                    })
                                    return out
                                  })()}

                                  {bars.map(({ f, left, width }) => {
                                    const mleStyle = isMleLeg(f)
                                      ? { bg: "#f3f4f6", fg: "#374151", bd: "#d1d5db" }
                                      : c
                                    // Whole trip group dims when any of its
                                    // segments is being dragged, so the user
                                    // can see at a glance which legs will
                                    // travel together on a vertical drop.
                                    const dragGid = draggingId
                                      ? tripGroups.get(draggingId) ?? draggingId
                                      : null
                                    const myGid = tripGroups.get(f.id) ?? f.id
                                    const isDragging =
                                      dragGid != null && dragGid === myGid
                                    return (
                                    <div
                                      key={f.id}
                                      draggable
                                      onDragStart={(e) => {
                                        e.dataTransfer.effectAllowed = "move"
                                        e.dataTransfer.setData("text/plain", f.id)
                                        const rect = (
                                          e.currentTarget as HTMLDivElement
                                        ).getBoundingClientRect()
                                        dragOffsetXRef.current = e.clientX - rect.left
                                        setDraggingId(f.id)
                                      }}
                                      onDragEnd={() => setDraggingId(null)}
                                      title={`${f.flightNo} · ${f.origin}${f.destination ? "→" + f.destination : ""} · STD ${f.depTime} / STA ${f.arrTime}${f.pax ? " · " + f.pax + " pax" : ""}`}
                                      style={{
                                        position: "absolute",
                                        left,
                                        top: 6,
                                        height: ROW_H - 12,
                                        width,
                                        background: mleStyle.bg,
                                        color: mleStyle.fg,
                                        border: `1px solid ${mleStyle.bd}`,
                                        borderRadius: 0,
                                        padding: "3px 7px 6px",
                                        fontSize: 11,
                                        fontWeight: 600,
                                        overflow: "hidden",
                                        lineHeight: 1.15,
                                        display: "flex",
                                        flexDirection: "column",
                                        justifyContent: "center",
                                        boxShadow:
                                          "0 1px 2px rgba(15,23,42,0.06)",
                                        cursor: "grab",
                                        opacity: isDragging ? 0.45 : 1,
                                      }}
                                    >
                                      <div
                                        style={{
                                          position: "absolute",
                                          right: -1,
                                          top: 0,
                                          bottom: 0,
                                          width: 1,
                                          background: mleStyle.fg,
                                          opacity: 0.45,
                                        }}
                                        title={`STA ${f.arrTime}`}
                                      />
                                      <div
                                        className="text-truncate"
                                        style={{ fontSize: 12, fontWeight: 700 }}
                                      >
                                        {f.flightNo}
                                      </div>
                                      <div className="text-truncate" style={{ fontSize: 10, opacity: 0.95 }}>
                                        {f.origin}
                                        {f.destination ? `→${f.destination}` : ""}
                                        {f.pax ? ` · ${f.pax} pax` : ""}
                                      </div>
                                      <div className="text-truncate" style={{ fontSize: 10, opacity: 0.9 }}>
                                        STD {f.depTime} · STA {f.arrTime}
                                      </div>
                                      {showCrew && (f.pic || f.sic || f.ca) ? (
                                        <div className="text-truncate" style={{ fontSize: 9, opacity: 0.9 }}>
                                          PIC {crew4(f.pic)} · SIC {crew4(f.sic)} · CA {crew4(f.ca)}
                                        </div>
                                      ) : null}
                                    </div>
                                    )
                                  })}
                                </>
                              )
                            })()}
                          </div>
                        </div>
                      )
                    })}

                    {/* MLE break-duration bands — vertical highlights
                        spanning all aircraft rows behind the flight bars.
                        Reflects whichever mode the brief toggle is set to:
                        Dep / Arr / Both. */}
                    {breakWindowEnabled ? brief.breaksByMode[breakMode].map((b, i) => {
                      if (b.endMin <= range.startMin || b.startMin >= range.endMin) {
                        return null
                      }
                      const clipStart = Math.max(b.startMin, range.startMin)
                      const clipEnd = Math.min(b.endMin, range.endMin)
                      const xLeft =
                        ((clipStart - range.startMin) / 60) * PIXELS_PER_HOUR
                      const xRight =
                        ((clipEnd - range.startMin) / 60) * PIXELS_PER_HOUR
                      const w = xRight - xLeft
                      if (w <= 0) return null
                      const fmtMin = (m: number) =>
                        `${pad2(Math.floor(m / 60) % 24)}:${pad2(m % 60)}`
                      const fmtDur = (m: number) => {
                        const h = Math.floor(m / 60)
                        const min = m % 60
                        if (h && min) return `${h}h ${min}m`
                        if (h) return `${h}h`
                        return `${min}m`
                      }
                      return (
                        <div
                          key={`break-${i}`}
                          style={{
                            position: "absolute",
                            left: AIRCRAFT_COL_W + xLeft,
                            top: HEADER_H,
                            width: w,
                            bottom: 0,
                            background:
                              "repeating-linear-gradient(135deg, rgba(254,243,199,0.45) 0 8px, rgba(254,243,199,0.18) 8px 16px)",
                            borderLeft: "1px dashed #d97706",
                            borderRight: "1px dashed #d97706",
                            zIndex: 1,
                            pointerEvents: "none",
                          }}
                          title={`Break ${fmtMin(b.startMin)} → ${fmtMin(b.endMin)} (${fmtDur(b.durationMin)})`}
                        >
                          <div
                            style={{
                              position: "absolute",
                              top: 4,
                              left: 4,
                              right: 4,
                              display: "flex",
                              justifyContent: "center",
                            }}
                          >
                            <span
                              style={{
                                background: "#fef3c7",
                                color: "#92400e",
                                fontSize: 10,
                                fontWeight: 700,
                                padding: "2px 8px",
                                borderRadius: 999,
                                border: "1px solid #fcd34d",
                                whiteSpace: "nowrap",
                                letterSpacing: 0.3,
                                boxShadow: "0 1px 2px rgba(0,0,0,0.05)",
                              }}
                            >
                              {fmtDur(b.durationMin)} ·{" "}
                              {fmtMin(b.startMin)}–{fmtMin(b.endMin)}
                            </span>
                          </div>
                        </div>
                      )
                    }) : null}

                    {/* Route conflict bands (red) */}
                    {conflictWindowEnabled
                      ? routeConflictBands.map((b, i) => {
                          if (b.endMin <= range.startMin || b.startMin >= range.endMin) return null
                          const clipStart = Math.max(b.startMin, range.startMin)
                          const clipEnd = Math.min(b.endMin, range.endMin)
                          const xLeft = ((clipStart - range.startMin) / 60) * PIXELS_PER_HOUR
                          const xRight = ((clipEnd - range.startMin) / 60) * PIXELS_PER_HOUR
                          const w = xRight - xLeft
                          if (w <= 0) return null
                          const fmtMin = (m: number) => `${pad2(Math.floor(m / 60) % 24)}:${pad2(m % 60)}`
                          return (
                            <div
                              key={`conf-band-${i}`}
                              style={{
                                position: "absolute",
                                left: AIRCRAFT_COL_W + xLeft,
                                top: HEADER_H,
                                width: w,
                                bottom: 0,
                                background:
                                  "repeating-linear-gradient(135deg, rgba(220,38,38,0.22) 0 8px, rgba(220,38,38,0.1) 8px 16px)",
                                borderLeft: "1px solid #dc2626",
                                borderRight: "1px solid #dc2626",
                                zIndex: 2,
                                pointerEvents: "none",
                              }}
                              title={`Conflict ${fmtMin(b.startMin)} → ${fmtMin(b.endMin)} | ${b.label}`}
                            />
                          )
                        })
                      : null}

                    {/* Live "now" indicator — vertical red line spanning all rows */}
                    {nowX != null ? (
                      <div
                        style={{
                          position: "absolute",
                          left: AIRCRAFT_COL_W + nowX,
                          top: HEADER_H,
                          bottom: 0,
                          width: 0,
                          borderLeft: "2px solid #dc2626",
                          zIndex: 6,
                          pointerEvents: "none",
                        }}
                      >
                        <div
                          style={{
                            position: "absolute",
                            top: -16,
                            left: -28,
                            background: "#dc2626",
                            color: "#fff",
                            fontSize: 10,
                            fontWeight: 700,
                            padding: "1px 6px",
                            borderRadius: 4,
                            letterSpacing: 0.5,
                            whiteSpace: "nowrap",
                            boxShadow: "0 1px 2px rgba(0,0,0,0.15)",
                          }}
                        >
                          NOW {pad2(Math.floor(nowMin / 60))}:
                          {pad2(nowMin % 60)}
                        </div>
                      </div>
                    ) : null}
                  </div>
                )
              })()}
            </div>
          )}
        </div>
      </div>

      {/* ----- Move flight conflict dialog ----- */}
      {moveDialog ? (() => {
        const moving = flights.find((x) => x.id === moveDialog.flightId)
        if (!moving) return null
        const movingIds = new Set(moveDialog.groupFlightIds)
        const groupMembers = flights
          .filter((x) => movingIds.has(x.id))
          .sort(
            (a, b) =>
              (parseTimeToMin(a.depTime) ?? 0) - (parseTimeToMin(b.depTime) ?? 0),
          )
        const targetSchedule = (flightsByAircraft.get(moveDialog.toAircraft) ?? [])
          .filter((x) => !movingIds.has(x.id))
        const minToHHMM = (m: number) =>
          `${pad2(Math.floor(m / 60) % 24)}:${pad2(m % 60)}`
        const newStd = minToHHMM(moveDialog.newStdMin)
        const newSta = minToHHMM(moveDialog.newStaMin)
        const aircraftChanged = moveDialog.fromAircraft !== moveDialog.toAircraft
        const isVertical = moveDialog.mode === "vertical"
        const isHorizontal = moveDialog.mode === "horizontal"

        // For each member compute its "after move" STD/STA. Vertical mode
        // preserves each leg's existing time. Horizontal mode only retimes
        // the primary; the rest of the group stays where it was.
        const newTimesById = new Map<string, { dep: number; arr: number }>()
        groupMembers.forEach((m) => {
          if (isHorizontal && m.id === moveDialog.flightId) {
            newTimesById.set(m.id, {
              dep: moveDialog.newStdMin,
              arr: moveDialog.newStaMin,
            })
          } else {
            const dep = parseTimeToMin(m.depTime)
            const arr = parseTimeToMin(m.arrTime) ?? dep
            if (dep != null && arr != null) {
              newTimesById.set(m.id, { dep, arr })
            }
          }
        })

        const conflictIds = new Set<string>()
        targetSchedule.forEach((t) => {
          const tDep = parseTimeToMin(t.depTime)
          const tArr = parseTimeToMin(t.arrTime) ?? tDep
          if (tDep == null || tArr == null) return
          for (const tm of newTimesById.values()) {
            if (tDep < tm.arr && tm.dep < tArr) {
              conflictIds.add(t.id)
              break
            }
          }
        })
        const hasConflicts = conflictIds.size > 0
        const stdChanged = newStd !== moving.depTime

        const close = () => setMoveDialog(null)
        const confirmMove = () => {
          setFlights(
            flights.map((x) => {
              if (!movingIds.has(x.id)) return x
              if (isHorizontal && x.id === moveDialog.flightId) {
                return { ...x, depTime: newStd, arrTime: newSta }
              }
              if (isVertical) {
                return { ...x, aircraft: moveDialog.toAircraft }
              }
              return x
            }),
          )
          setMoveDialog(null)
        }
        return (
          <div
            className="modal show d-block"
            tabIndex={-1}
            style={{ background: "rgba(0,0,0,0.5)" }}
            onClick={close}
          >
            <div
              className="modal-dialog modal-dialog-centered modal-lg"
              onClick={(e) => e.stopPropagation()}
            >
              <div
                className="modal-content border-0 shadow"
                style={{ borderRadius: 14 }}
              >
                <div className="modal-header border-0 pb-0">
                  <h5 className="modal-title d-flex align-items-center gap-2">
                    {hasConflicts ? (
                      <AlertTriangle size={18} color="#dc2626" />
                    ) : (
                      <CheckCircle2 size={18} color="#15803d" />
                    )}
                    {isVertical
                      ? groupMembers.length > 1
                        ? `Move trip (${groupMembers.length} segments)`
                        : `Move flight ${moving.flightNo}`
                      : `Reschedule flight ${moving.flightNo}`}
                  </h5>
                  <button className="btn btn-sm btn-light" onClick={close}>
                    Close
                  </button>
                </div>
                <div className="modal-body">
                  <div
                    className="rounded p-3 mb-3"
                    style={{
                      background: "var(--bs-body-tertiary-bg, #f8fafc)",
                      border: "1px solid var(--bs-border-color, #e2e8f0)",
                    }}
                  >
                    <div className="small text-muted mb-1">
                      {aircraftChanged
                        ? groupMembers.length > 1
                          ? "Reassigning trip"
                          : "Reassigning"
                        : "On"}
                    </div>
                    <div className="d-flex align-items-center gap-2 fw-semibold">
                      <span className="badge bg-secondary-subtle text-secondary-emphasis">
                        {moveDialog.fromAircraft}
                      </span>
                      {aircraftChanged ? (
                        <>
                          <span>→</span>
                          <span
                            className="badge"
                            style={{ background: "#dbeafe", color: "#1d4ed8" }}
                          >
                            {moveDialog.toAircraft}
                          </span>
                        </>
                      ) : null}
                    </div>
                    {isVertical && groupMembers.length > 1 ? (
                      <div className="mt-2">
                        <div className="small text-muted mb-1">
                          Trip segments (each keeps its STD/STA):
                        </div>
                        <ul className="list-unstyled small mb-0 ps-0">
                          {groupMembers.map((m) => (
                            <li key={m.id}>
                              <strong>{m.flightNo}</strong> · {m.origin}
                              {m.destination ? `→${m.destination}` : ""} · STD{" "}
                              {m.depTime} / STA {m.arrTime}
                              {m.pax ? ` · ${m.pax} pax` : ""}
                            </li>
                          ))}
                        </ul>
                      </div>
                    ) : (
                      <>
                        <div className="small mt-2">
                          <strong>{moving.flightNo}</strong> · {moving.origin}
                          {moving.destination ? `→${moving.destination}` : ""}
                          {moving.pax ? ` · ${moving.pax} pax` : ""}
                        </div>
                        <div className="small mt-2 d-flex align-items-center gap-2 flex-wrap">
                          {stdChanged ? (
                            <>
                              <span className="text-muted text-decoration-line-through">
                                STD {moving.depTime} / STA {moving.arrTime}
                              </span>
                              <span>→</span>
                              <span
                                className="fw-semibold"
                                style={{ color: "#1d4ed8" }}
                              >
                                STD {newStd} / STA {newSta}
                              </span>
                            </>
                          ) : (
                            <span>
                              STD {moving.depTime} / STA {moving.arrTime}
                            </span>
                          )}
                        </div>
                      </>
                    )}
                  </div>

                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <div className="fw-semibold small">
                      Other flights on {moveDialog.toAircraft}
                    </div>
                    {hasConflicts ? (
                      <span className="badge bg-danger-subtle text-danger-emphasis">
                        {conflictIds.size} conflict
                        {conflictIds.size === 1 ? "" : "s"}
                      </span>
                    ) : (
                      <span className="badge bg-success-subtle text-success-emphasis">
                        No time conflicts
                      </span>
                    )}
                  </div>

                  {targetSchedule.length === 0 ? (
                    <div className="text-muted small fst-italic">
                      {aircraftChanged
                        ? `No other flights currently on ${moveDialog.toAircraft}.`
                        : `No other flights on ${moveDialog.toAircraft}.`}
                    </div>
                  ) : (
                    <div className="table-responsive">
                      <table
                        className="table table-sm align-middle mb-0"
                        style={{ fontSize: 12 }}
                      >
                        <thead>
                          <tr className="text-uppercase small text-muted">
                            <th style={{ width: 36 }}></th>
                            <th>Flight</th>
                            <th>Route</th>
                            <th>STD</th>
                            <th>STA</th>
                            <th>Crew</th>
                          </tr>
                        </thead>
                        <tbody>
                          {targetSchedule.map((t) => {
                            const conflict = conflictIds.has(t.id)
                            return (
                              <tr
                                key={t.id}
                                style={{
                                  background: conflict
                                    ? "#fef2f2"
                                    : "transparent",
                                }}
                              >
                                <td>
                                  {conflict ? (
                                    <AlertTriangle size={14} color="#dc2626" />
                                  ) : null}
                                </td>
                                <td className="fw-semibold">{t.flightNo}</td>
                                <td>
                                  {t.origin}
                                  {t.destination ? `→${t.destination}` : ""}
                                  {t.pax ? ` · ${t.pax} pax` : ""}
                                </td>
                                <td>{t.depTime || "—"}</td>
                                <td>{t.arrTime || "—"}</td>
                                <td className="text-muted small">
                                  {t.pic || t.sic || t.ca
                                    ? `${crew4(t.pic)} / ${crew4(t.sic)} / ${crew4(t.ca)}`
                                    : "—"}
                                </td>
                              </tr>
                            )
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}

                  {hasConflicts ? (
                    <div className="alert alert-warning small mt-3 mb-0 d-flex align-items-start gap-2">
                      <AlertTriangle size={16} className="mt-1" />
                      <div>
                        {aircraftChanged ? "Moving" : "Rescheduling"} this leg
                        will overlap <strong>{conflictIds.size}</strong>{" "}
                        existing flight
                        {conflictIds.size === 1 ? "" : "s"} on{" "}
                        {moveDialog.toAircraft}. You can still proceed, but
                        the schedule will need a manual fix.
                      </div>
                    </div>
                  ) : null}
                </div>
                <div className="modal-footer border-0">
                  <button className="btn btn-light" onClick={close}>
                    Cancel
                  </button>
                  <button
                    className={`btn ${hasConflicts ? "btn-danger" : "btn-primary"}`}
                    onClick={confirmMove}
                  >
                    {hasConflicts
                      ? aircraftChanged
                        ? "Move anyway"
                        : "Reschedule anyway"
                      : aircraftChanged
                      ? "Confirm move"
                      : "Confirm reschedule"}
                  </button>
                </div>
              </div>
            </div>
          </div>
        )
      })() : null}
    </div>
  )
}
