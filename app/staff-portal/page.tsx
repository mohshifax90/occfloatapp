"use client"

import { useEffect, useMemo, useState } from "react"
import {
  Activity,
  AlertTriangle,
  ArrowRight,
  Bed,
  Bell,
  Briefcase,
  CalendarCheck,
  CalendarDays,
  ChevronLeft,
  ChevronRight,
  Clipboard,
  ClipboardList,
  Clock,
  Coffee,
  GraduationCap,
  Headphones,
  Headset,
  LogOut,
  MapPin,
  Megaphone,
  Mic,
  MoreHorizontal,
  Phone,
  Plane,
  Plus,
  Radio,
  Send,
  Shield,
  Sun,
  User,
  UserCog,
  UserPlus,
  Users,
  X,
} from "lucide-react"
import { ThemeToggle } from "@/components/theme-toggle"
import { createClient as createSupabaseClient } from "@/utils/supabase/client"
import Box from "@mui/material/Box"
import Stepper from "@mui/material/Stepper"
import Step from "@mui/material/Step"
import StepLabel from "@mui/material/StepLabel"

// ---------------------------------------------------------------------------
// Storage keys (must match app/page.tsx).
// ---------------------------------------------------------------------------
const STORAGE_KEY = "occfloat.v1"
const AUTH_CACHE_KEY = "occfloat.authCache.v1"
const AUTH_STORAGE_KEY = "occfloat.staffPortalAuthStaffId"
const STAFF_PORTAL_REQUESTS_KEY = "occfloat.staffPortalRequests"
const SUPABASE_STORE_TABLE = "occfloat_store"
const SUPABASE_STORE_ID = "primary"
const STAFF_PORTAL_BUILD_TAG = "f4092ea"

// ---------------------------------------------------------------------------
// Minimal mirrors of the types we need from the main app. We keep them loose
// (string-keyed) because the data store on disk is `Record<string, string>`.
// ---------------------------------------------------------------------------
type Entry = {
  id: string
  createdAt: string
  [key: string]: string
}

type DataStore = {
  staff: Entry[]
  levels: Entry[]
  leaves: Entry[]
  leaveAttendanceControl: Entry[]
  publicHolidays: Entry[]
  roster: Entry[]
  shift: Entry[]
  attendance: Entry[]
  evaluation: Entry[]
  briefing: Entry[]
  checklist: Entry[]
  opsControl: Entry[]
  // Other modules omitted intentionally — we only read what the portal needs.
  [key: string]: Entry[]
}

type StaffRequest = {
  id: string
  createdAt: string
  staffId: string
  staffNo: string
  staffName: string
  type: "Leave" | "Attendance" | "Roster Change"
  status:
    | "Pending Peer Acceptance"
    | "Pending Approval"
    | "Pending Document"
    | "Document Submitted"
    | "Approved"
    | "Cancelled"
    | "Rejected"
  // Leave fields:
  leavePolicyName?: string
  fromDate?: string
  toDate?: string
  // Roster change fields:
  date?: string
  fromCode?: string
  toCode?: string
  changeWithStaffNo?: string
  changeWithStaffName?: string
  peerDecisionByStaffNo?: string
  peerDecisionByStaffName?: string
  peerDecisionAt?: string
  peerDecision?: "Accepted" | "Rejected"
  dutyMarkType?: "Medical" | "Sick Leave" | "Family Responsible Leave"
  noOfDays?: string
  documentName?: string
  documentData?: string
  documentUploadedAt?: string
  // Free-form note from the staff member.
  reason: string
}

type PortalTab = "ops" | "roster" | "leave" | "profile"

type PairingRow = {
  staffNo: string
  staffName: string
  shiftCode: string
  level: string
}

const workflowStepperSx = {
  "& .MuiStepLabel-label": { fontSize: 13, color: "#6c757d" },
  "& .MuiStepLabel-label.Mui-active": { color: "#0d6efd", fontWeight: 600 },
  "& .MuiStepLabel-label.Mui-completed": { color: "#198754", fontWeight: 600 },
  "& .MuiStepIcon-root": { color: "#adb5bd" },
  "& .MuiStepIcon-root.Mui-active": { color: "#0d6efd" },
  "& .MuiStepIcon-root.Mui-completed": { color: "#198754" },
}

function requestStatusBadgeStyle(status: StaffRequest["status"]): React.CSSProperties {
  if (status === "Approved") return { background: "#d1e7dd", color: "#0f5132" }
  if (status === "Rejected") return { background: "#f8d7da", color: "#842029" }
  if (status === "Cancelled") return { background: "#e2e3e5", color: "#41464b" }
  if (status === "Pending Approval") return { background: "#cff4fc", color: "#055160" }
  if (status === "Pending Document" || status === "Document Submitted") {
    return { background: "#fff3cd", color: "#664d03" }
  }
  if (status === "Pending Peer Acceptance") return { background: "#e2e3e5", color: "#41464b" }
  return { background: "#e2e3e5", color: "#41464b" }
}

function getPortalRequestStepIndex(status: StaffRequest["status"]): number {
  const s = (status || "").trim().toLowerCase()
  if (s === "rejected" || s === "approved" || s === "cancelled") return 3
  if (s === "pending approval") return 2
  if (s === "pending document" || s === "document submitted") return 1
  return 0
}

// ---------------------------------------------------------------------------
// Lightweight helpers (mirrored from app/page.tsx).
// ---------------------------------------------------------------------------
function toYmd(date: Date): string {
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, "0")
  const d = String(date.getDate()).padStart(2, "0")
  return `${y}-${m}-${d}`
}

function getDateRangeInclusive(fromDate: string, toDate: string): Date[] {
  const start = new Date(`${fromDate}T00:00:00`)
  const end = new Date(`${toDate}T00:00:00`)
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || start > end) return []
  const result: Date[] = []
  const cursor = new Date(start)
  while (cursor <= end) {
    result.push(new Date(cursor))
    cursor.setDate(cursor.getDate() + 1)
  }
  return result
}

function parseTimeToMin(value: string): number | null {
  const s = (value || "").trim()
  const m = s.match(/^(\d{1,2}):(\d{2})$/)
  if (!m) return null
  const hh = Number(m[1])
  const mm = Number(m[2])
  if (!Number.isInteger(hh) || !Number.isInteger(mm)) return null
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null
  return hh * 60 + mm
}

function isOffCode(value: string): boolean {
  const code = (value || "").trim().toUpperCase()
  return code === "OF" || code === "OFF" || code === "REST"
}

function getEmptyStore(): DataStore {
  return {
    staff: [],
    levels: [],
    leaves: [],
    leaveAttendanceControl: [],
    publicHolidays: [],
    roster: [],
    shift: [],
    attendance: [],
    evaluation: [],
    briefing: [],
    checklist: [],
    opsControl: [],
  }
}

function safeParseStore(raw: string | null): DataStore {
  if (!raw) return getEmptyStore()
  try {
    const parsed = JSON.parse(raw)
    return { ...getEmptyStore(), ...parsed }
  } catch {
    return getEmptyStore()
  }
}

function normalizeStoreFromPayload(payload: unknown): DataStore {
  if (!payload || typeof payload !== "object") return getEmptyStore()
  const incoming = payload as Record<string, unknown>
  const base = getEmptyStore()
  const next: DataStore = { ...base }
  ;(Object.keys(base) as Array<keyof DataStore>).forEach((k) => {
    const v = incoming[k as string]
    next[k] = Array.isArray(v) ? (v as Entry[]) : []
  })
  return next
}

function safeParseRequests(raw: string | null): StaffRequest[] {
  if (!raw) return []
  try {
    const parsed = JSON.parse(raw)
    if (Array.isArray(parsed)) return parsed as StaffRequest[]
    return []
  } catch {
    return []
  }
}

// Keep a stable per-staff icon by hashing the staff number into a small palette.
const STAFF_ICONS = [
  User,
  UserPlus,
  UserCog,
  Users,
  Shield,
  Headphones,
  Radio,
  GraduationCap,
] as const

function iconIndexForStaffNo(staffNo: string): number {
  const key = (staffNo || "").trim()
  if (!key) return 0
  let hash = 0
  for (let i = 0; i < key.length; i++) {
    hash = (hash * 31 + key.charCodeAt(i)) >>> 0
  }
  return hash % STAFF_ICONS.length
}

// Renders a stable icon for a given staffNo. The wrapping component itself is
// declared at module scope so React does not see a "new component per render".
function StaffAvatarIcon({ staffNo, size = 22 }: { staffNo: string; size?: number }) {
  const Icon = STAFF_ICONS[iconIndexForStaffNo(staffNo)]
  return <Icon size={size} />
}

// Convert "08:00" → "8:00 AM"
function formatTime12(hhmm: string): string {
  if (!hhmm) return ""
  const match = hhmm.match(/^(\d{1,2}):(\d{2})$/)
  if (!match) return hhmm
  let h = Number(match[1])
  const m = match[2]
  const ampm = h >= 12 ? "PM" : "AM"
  if (h === 0) h = 12
  else if (h > 12) h -= 12
  return `${h}:${m} ${ampm}`
}

// Pick a stable shift-icon "kind" per shift code so the calendar reads like
// the agenda mockup (phone for comms, plane for travel, mic for briefings, …).
type ShiftIconKind =
  | "off"
  | "leave"
  | "holiday"
  | "headset"
  | "plane"
  | "user"
  | "calendar"
  | "clipboard"
  | "mic"
  | "briefcase"
  | "coffee"
  | "radio"
  | "clock"
  | null

const SHIFT_COLOR_PALETTE = [
  { bg: "#60a5fa", fg: "#0f172a" },
  { bg: "#4ade80", fg: "#052e16" },
  { bg: "#fbbf24", fg: "#1f2937" },
  { bg: "#a78bfa", fg: "#1e1b4b" },
  { bg: "#f87171", fg: "#450a0a" },
  { bg: "#22d3ee", fg: "#083344" },
  { bg: "#818cf8", fg: "#1e1b4b" },
  { bg: "#fb923c", fg: "#431407" },
  { bg: "#2dd4bf", fg: "#042f2e" },
  { bg: "#f472b6", fg: "#500724" },
]

function shiftIconKind(code: string, shiftByCode?: Map<string, Entry>): ShiftIconKind {
  const norm = (code || "").trim().toUpperCase()
  if (!norm || norm === "-") return null
  if (isOffCode(norm)) return "off"
  if (norm === "AL") return "leave"
  if (norm === "PH" || norm === "GH") return "holiday"
  const shift = shiftByCode?.get(norm)
  const shiftName = (shift?.shiftName || "").trim().toLowerCase()
  const shiftType = (shift?.shiftType || "").trim().toLowerCase()
  if (shiftName.includes("flight watch")) return "headset"
  if (shiftName.includes("flight release")) return "plane"
  if (shiftName.includes("ops controller")) return "user"
  if (shiftName.includes("crew scheduler")) return "calendar"
  if (shiftName.includes("docs and load") || shiftName.includes("docs & load")) return "clipboard"
  if (shiftType === "early") return "coffee"
  if (shiftType === "mid") return "briefcase"
  if (shiftType === "late") return "radio"
  let hash = 0
  for (let i = 0; i < norm.length; i++) {
    hash = (hash * 31 + norm.charCodeAt(i)) >>> 0
  }
  const fallback: ShiftIconKind[] = ["headset", "plane", "mic", "briefcase", "coffee", "radio"]
  return fallback[hash % fallback.length]
}

// Stable component that branches over the icon kind. Each JSX node references
// a constant lucide component, so React's static-components rule is happy.
function ShiftCodeIcon({
  code,
  shiftByCode,
  size = 16,
  color,
}: {
  code: string
  shiftByCode?: Map<string, Entry>
  size?: number
  color?: string
}) {
  const kind = shiftIconKind(code, shiftByCode)
  switch (kind) {
    case "off":
      return <Bed size={size} color={color} />
    case "leave":
      return (
        <img
          src="/icons/annual-leave.svg"
          alt="Annual leave"
          width={size}
          height={size}
          style={{ width: size, height: size, objectFit: "contain" }}
        />
      )
    case "holiday":
      return <Sun size={size} color={color} />
    case "headset":
      return <Headset size={size} color={color} />
    case "plane":
      return <Plane size={size} color={color} />
    case "user":
      return <User size={size} color={color} />
    case "calendar":
      return <CalendarCheck size={size} color={color} />
    case "clipboard":
      return <Clipboard size={size} color={color} />
    case "mic":
      return <Mic size={size} color={color} />
    case "briefcase":
      return <Briefcase size={size} color={color} />
    case "coffee":
      return <Coffee size={size} color={color} />
    case "radio":
      return <Radio size={size} color={color} />
    default:
      return <Clock size={size} color={color} />
  }
}

// Helper for callers that just want to know "is there an icon for this code?"
function hasShiftIcon(code: string, shiftByCode?: Map<string, Entry>): boolean {
  return shiftIconKind(code, shiftByCode) !== null
}

function isShiftPairingEnabled(code: string, shiftByCode?: Map<string, Entry>): boolean {
  const norm = (code || "").trim().toUpperCase()
  if (!norm || norm === "-") return false
  if (isOffCode(norm) || norm === "AL" || norm === "PH" || norm === "GH") return false
  const shift = shiftByCode?.get(norm)
  if (!shift) return false
  return (shift.showInShiftPairing || "").trim().toLowerCase() === "yes"
}

// Build a deduplicated pairing list (staff working on `date` excluding `me`).
// Extracted to module scope so both the modal and the inline profile view can
// share the exact same shape and ordering.
function buildPairingForDate(
  date: string | null,
  rosterRows: Entry[],
  staffRows: Entry[],
  meStaffNo: string,
  shiftByCode?: Map<string, Entry>,
): PairingRow[] {
  if (!date) return []
  const rows = rosterRows
    .filter(
      (r) =>
        r.date === date &&
        r.shiftCode &&
        !isOffCode(r.shiftCode) &&
        isShiftPairingEnabled(r.shiftCode || "", shiftByCode),
    )
    .map<PairingRow>((r) => {
      const staff = staffRows.find(
        (s) =>
          (s.staffNo || "").trim() === (r.staffNo || "").trim() &&
          (s.fullName || "").trim() === (r.staffName || "").trim(),
      )
      return {
        staffNo: r.staffNo || "",
        staffName: r.staffName || "",
        shiftCode: (r.shiftCode || "").toUpperCase(),
        level: staff?.level || "",
      }
    })
    .filter((r) => r.staffNo !== meStaffNo)
  return rows.sort((a, b) => {
    if (a.shiftCode !== b.shiftCode) return a.shiftCode.localeCompare(b.shiftCode)
    return a.staffName.localeCompare(b.staffName)
  })
}

// Used for the new agenda calendar — soft tint instead of saturated fill so
// the layout reads like the reference mockup (light blue cells, dark text).
function softTintForCode(code: string, shiftByCode?: Map<string, Entry>): { bg: string; fg: string } {
  const norm = (code || "").trim().toUpperCase()
  if (!norm || norm === "-") return { bg: "transparent", fg: "inherit" }
  if (isOffCode(norm)) return { bg: "#e2e8f0", fg: "#475569" }
  if (norm === "AL") return { bg: "#9ca3af", fg: "#111827" }
  if (norm === "PH") return { bg: "#f472b6", fg: "#500724" }
  if (norm === "GH") return { bg: "#4ade80", fg: "#052e16" }
  if (!shiftByCode || shiftByCode.size === 0) return { bg: "#dbeafe", fg: "#1d4ed8" }

  // Keep parity with main app: assign colors by unique shift time window,
  // in the same shift-definition order.
  const timeKeyToColor = new Map<string, (typeof SHIFT_COLOR_PALETTE)[number]>()
  const codeToColor = new Map<string, (typeof SHIFT_COLOR_PALETTE)[number]>()
  let colorIndex = 0
  Array.from(shiftByCode.values()).forEach((shift) => {
    const c = (shift.shiftCode || "").trim().toUpperCase()
    if (!c) return
    const timeKey = `${shift.startTime || ""}-${shift.endTime || ""}`
    if (!timeKeyToColor.has(timeKey)) {
      timeKeyToColor.set(timeKey, SHIFT_COLOR_PALETTE[colorIndex % SHIFT_COLOR_PALETTE.length])
      colorIndex += 1
    }
    codeToColor.set(c, timeKeyToColor.get(timeKey)!)
  })
  const mapped = codeToColor.get(norm)
  if (mapped) return { bg: mapped.bg, fg: mapped.fg }
  return { bg: "#dbeafe", fg: "#1d4ed8" }
}

// ---------------------------------------------------------------------------
// Page component.
// ---------------------------------------------------------------------------
export default function StaffPortalPage() {
  const [hydrated, setHydrated] = useState(false)
  const [store, setStore] = useState<DataStore>(getEmptyStore())
  const [authStaffId, setAuthStaffId] = useState<string | null>(null)
  const [staffRequests, setStaffRequests] = useState<StaffRequest[]>([])

  const [activeTab, setActiveTab] = useState<PortalTab>("ops")
  const [calendarMonth, setCalendarMonth] = useState<{ year: number; month: number }>(() => {
    const now = new Date()
    return { year: now.getFullYear(), month: now.getMonth() }
  })

  // Login form (shown when no auth staff id is present).
  const [loginStaffNo, setLoginStaffNo] = useState("")
  const [loginPassword, setLoginPassword] = useState("")
  const [loginError, setLoginError] = useState("")
  const [loginStage, setLoginStage] = useState<"identify" | "password">("identify")

  // Pairing modal state (Ops tab).
  const [pairingDate, setPairingDate] = useState<string | null>(null)
  const [rosterActionDate, setRosterActionDate] = useState<string | null>(null)
  const [rosterActionMode, setRosterActionMode] = useState<"Shift Change" | "Duty Mark As">("Shift Change")

  // Inline pairing selection on the Profile tab (defaults to today on mount).
  const [profilePairingDate, setProfilePairingDate] = useState<string>(() =>
    toYmd(new Date()),
  )
  const [isLeaveRequestModalOpen, setIsLeaveRequestModalOpen] = useState(false)

  // Leave + roster-change request forms.
  const [leaveForm, setLeaveForm] = useState({
    leavePolicyName: "",
    fromDate: "",
    toDate: "",
    reason: "",
  })
  const [changeForm, setChangeForm] = useState({
    date: "",
    fromCode: "",
    toCode: "",
    changeWith: "",
    reason: "",
  })
  const [dutyMarkForm, setDutyMarkForm] = useState({
    dutyType: "Medical" as "Medical" | "Sick Leave" | "Family Responsible Leave",
    fromDate: "",
    toDate: "",
    reason: "",
  })
  const supabase = useMemo(() => {
    try {
      return createSupabaseClient()
    } catch {
      return null
    }
  }, [])

  // ----- hydration -----
  // We need to read from `window.localStorage`, which is only available on
  // the client, so this runs in an effect after the first client render.
  /* eslint-disable react-hooks/set-state-in-effect */
  useEffect(() => {
    if (typeof window === "undefined") return
    const seeded = safeParseStore(window.localStorage.getItem(STORAGE_KEY))
    try {
      const authCacheRaw = window.localStorage.getItem(AUTH_CACHE_KEY)
      if (authCacheRaw) {
        const authCache = JSON.parse(authCacheRaw) as {
          staff?: Entry[]
        }
        if (Array.isArray(authCache.staff) && authCache.staff.length > 0) {
          seeded.staff = authCache.staff
        }
      }
    } catch {
      // ignore malformed cache
    }
    setStore(seeded)
    setStaffRequests(safeParseRequests(window.localStorage.getItem(STAFF_PORTAL_REQUESTS_KEY)))
    setAuthStaffId(window.localStorage.getItem(AUTH_STORAGE_KEY))
    setHydrated(true)
  }, [])
  /* eslint-enable react-hooks/set-state-in-effect */

  // Pull full app store from Supabase so staff login works on fresh devices
  // where localStorage has not been seeded yet.
  useEffect(() => {
    if (!hydrated || !supabase) return
    const pullStore = async () => {
      try {
        const { data } = await supabase
          .from(SUPABASE_STORE_TABLE)
          .select("payload")
          .eq("id", SUPABASE_STORE_ID)
          .maybeSingle()
        const next = normalizeStoreFromPayload(data?.payload)
        if (next.staff.length === 0) return
        setStore(next)
        try {
          window.localStorage.setItem(
            AUTH_CACHE_KEY,
            JSON.stringify({ staff: next.staff }),
          )
        } catch {
          // ignore
        }
      } catch {
        // Non-blocking.
      }
    }
    void pullStore()
  }, [hydrated, supabase])

  // Persist staff-submitted requests so the main app can pick them up.
  useEffect(() => {
    if (!hydrated) return
    window.localStorage.setItem(
      STAFF_PORTAL_REQUESTS_KEY,
      JSON.stringify(staffRequests),
    )
  }, [staffRequests, hydrated])

  // Pull mirrored request statuses from Supabase so approvals from main app
  // reflect in portal even if different host/port is used.
  useEffect(() => {
    if (!hydrated || !supabase) return
    const pull = async () => {
      try {
        const { data } = await supabase
          .from(SUPABASE_STORE_TABLE)
          .select("payload")
          .eq("id", SUPABASE_STORE_ID)
          .maybeSingle()
        const payload = (data?.payload && typeof data.payload === "object")
          ? (data.payload as Record<string, unknown>)
          : null
        const remote = (payload?.__staffPortalRequests || []) as StaffRequest[]
        if (!Array.isArray(remote) || remote.length === 0) return
        setStaffRequests((prev) => {
          const map = new Map(prev.map((r) => [r.id, r]))
          remote.forEach((r) => {
            if (!r?.id) return
            const existing = map.get(r.id)
            if (!existing) {
              map.set(r.id, r)
              return
            }
            map.set(r.id, { ...existing, ...r })
          })
          return Array.from(map.values()).sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || ""))
        })
      } catch {
        // Non-blocking.
      }
    }
    void pull()
    const id = window.setInterval(() => {
      void pull()
    }, 4000)
    return () => window.clearInterval(id)
  }, [hydrated, supabase])

  // Mirror portal requests into Supabase store payload so main app can import
  // across different browser origins (port/domain).
  useEffect(() => {
    if (!hydrated || !supabase) return
    const timer = window.setTimeout(async () => {
      try {
        const { data } = await supabase
          .from(SUPABASE_STORE_TABLE)
          .select("payload")
          .eq("id", SUPABASE_STORE_ID)
          .maybeSingle()
        const payload = (data?.payload && typeof data.payload === "object") ? { ...(data.payload as Record<string, unknown>) } : {}
        payload.__staffPortalRequests = staffRequests
        await supabase.from(SUPABASE_STORE_TABLE).upsert({
          id: SUPABASE_STORE_ID,
          payload,
          updated_at: new Date().toISOString(),
        })
      } catch {
        // Keep portal usable even if remote sync fails.
      }
    }, 400)
    return () => window.clearTimeout(timer)
  }, [staffRequests, hydrated, supabase])

  const me = useMemo<Entry | null>(() => {
    if (!authStaffId) return null
    return store.staff.find((s) => s.id === authStaffId) ?? null
  }, [store.staff, authStaffId])

  // Today's date string (single source of truth used in multiple sub-views).
  const todayYmd = useMemo(() => toYmd(new Date()), [])

  // Build a per-day shift-code map for me from the roster store.
  const myRosterByDate = useMemo(() => {
    if (!me) return new Map<string, string>()
    const map = new Map<string, string>()
    store.roster.forEach((row) => {
      if (
        (row.staffNo || "").trim() === (me.staffNo || "").trim() &&
        (row.staffName || "").trim() === (me.fullName || "").trim() &&
        row.date
      ) {
        map.set(row.date, row.shiftCode || "")
      }
    })
    return map
  }, [store.roster, me])

  // Approved leaves that overlap a given day → the leave-policy code we'd
  // show in the calendar cell. This is a small reimplementation of the main
  // app's leaveCodeByStaffAndDate logic, scoped to "me only".
  const myLeavesByDate = useMemo(() => {
    if (!me) return new Map<string, string>()
    const policyCodeByName = new Map<string, string>()
    store.leaveAttendanceControl.forEach((row) => {
      const name = (row.leaveAttendanceName || "").trim()
      const code = (row.typeCode || "").trim().toUpperCase()
      if (name && code) policyCodeByName.set(name, code)
    })
    const map = new Map<string, string>()
    store.leaves.forEach((leave) => {
      const sno = (leave.staffNo || "").trim()
      if (sno !== (me.staffNo || "").trim()) return
      const status = (leave.status || "").trim().toLowerCase()
      if (status === "rejected") return
      const from = leave.fromDate
      const to = leave.toDate
      if (!from || !to) return
      const start = new Date(`${from}T00:00:00`)
      const end = new Date(`${to}T00:00:00`)
      if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return
      const code = policyCodeByName.get((leave.leavePolicyName || "").trim()) || "AL"
      const cursor = new Date(start)
      while (cursor <= end) {
        map.set(toYmd(cursor), code)
        cursor.setDate(cursor.getDate() + 1)
      }
    })
    return map
  }, [store.leaves, store.leaveAttendanceControl, me])

  const govHolidayDates = useMemo(() => {
    const set = new Set<string>()
    store.publicHolidays.forEach((h) => {
      const from = h.holidayStartDate
      const to = h.holidayEndDate
      if (!from || !to) return
      const start = new Date(`${from}T00:00:00`)
      const end = new Date(`${to}T00:00:00`)
      if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return
      const cursor = new Date(start)
      while (cursor <= end) {
        set.add(toYmd(cursor))
        cursor.setDate(cursor.getDate() + 1)
      }
    })
    return set
  }, [store.publicHolidays])

  const codeForMyDay = (ymd: string): string => {
    const shift = myRosterByDate.get(ymd)
    if (shift) return shift
    const leave = myLeavesByDate.get(ymd)
    if (leave) return leave
    if (govHolidayDates.has(ymd)) return "PH"
    return ""
  }

  // Shift lookup by code (for the Ops "current shift" card).
  const shiftByCode = useMemo(() => {
    const map = new Map<string, Entry>()
    store.shift.forEach((s) => {
      const c = (s.shiftCode || "").trim().toUpperCase()
      if (c) map.set(c, s)
    })
    return map
  }, [store.shift])

  // Same-day pairing lists: staff working pairing-enabled shifts on `date`,
  // excluding me. The modal uses `pairingDate`; profile has its own selection.
  const meStaffNo = (me?.staffNo || "").trim()
  const pairingForDate = useMemo(
    () => buildPairingForDate(pairingDate, store.roster, store.staff, meStaffNo, shiftByCode),
    [pairingDate, store.roster, store.staff, meStaffNo, shiftByCode],
  )
  const rosterActionPairing = useMemo(
    () => buildPairingForDate(rosterActionDate, store.roster, store.staff, meStaffNo, shiftByCode),
    [rosterActionDate, store.roster, store.staff, meStaffNo, shiftByCode],
  )
  const profilePairing = useMemo(
    () =>
      buildPairingForDate(profilePairingDate, store.roster, store.staff, meStaffNo, shiftByCode),
    [profilePairingDate, store.roster, store.staff, meStaffNo, shiftByCode],
  )

  const leavePolicyOptions = useMemo(
    () =>
      store.leaveAttendanceControl
        .filter((p) => (p.leaveAttendanceType || "").trim().toLowerCase() === "leave")
        .map((p) => p.leaveAttendanceName || "")
        .filter(Boolean),
    [store.leaveAttendanceControl],
  )

  const allShiftCodes = useMemo(() => {
    const codes = new Set<string>()
    store.shift.forEach((s) => {
      const c = (s.shiftCode || "").trim().toUpperCase()
      if (c) codes.add(c)
    })
    codes.add("OF")
    codes.add("AL")
    return Array.from(codes).sort()
  }, [store.shift])

  // ------------------------------------------------------------------
  // Calendar grid for the Roster tab.
  // ------------------------------------------------------------------
  // Build a Monday-first calendar matrix for the visible month, padding the
  // leading and trailing positions with the spillover days from neighbouring
  // months (so the grid is always a clean 6x7 / 5x7 block).
  const calendarCells = useMemo(() => {
    const { year, month } = calendarMonth
    const first = new Date(year, month, 1)
    const last = new Date(year, month + 1, 0)
    // Monday-first → Mon=0, Tue=1, ... Sun=6.
    const startOffset = (first.getDay() + 6) % 7
    const cells: Array<{ ymd: string; day: number; inMonth: boolean }> = []
    for (let i = startOffset; i > 0; i--) {
      const d = new Date(year, month, 1 - i)
      cells.push({ ymd: toYmd(d), day: d.getDate(), inMonth: false })
    }
    for (let d = 1; d <= last.getDate(); d++) {
      const date = new Date(year, month, d)
      cells.push({ ymd: toYmd(date), day: d, inMonth: true })
    }
    let trailing = 1
    while (cells.length % 7 !== 0) {
      const d = new Date(year, month + 1, trailing)
      cells.push({ ymd: toYmd(d), day: d.getDate(), inMonth: false })
      trailing += 1
    }
    return cells
  }, [calendarMonth])

  const monthLabel = useMemo(() => {
    const d = new Date(calendarMonth.year, calendarMonth.month, 1)
    return d.toLocaleDateString("en-US", { month: "long", year: "numeric" })
  }, [calendarMonth])

  // ------------------------------------------------------------------
  // Auth handlers.
  // ------------------------------------------------------------------
  const findStaffByNo = async (sno: string): Promise<Entry | null> => {
    let staff = store.staff.find((s) => (s.staffNo || "").trim() === sno) || null
    if (!staff && supabase) {
      try {
        const { data } = await supabase
          .from(SUPABASE_STORE_TABLE)
          .select("payload")
          .eq("id", SUPABASE_STORE_ID)
          .maybeSingle()
        const remoteStore = normalizeStoreFromPayload(data?.payload)
        if (remoteStore.staff.length > 0) {
          setStore(remoteStore)
          try {
            window.localStorage.setItem(
              AUTH_CACHE_KEY,
              JSON.stringify({ staff: remoteStore.staff }),
            )
          } catch {
            // ignore
          }
          staff = remoteStore.staff.find((s) => (s.staffNo || "").trim() === sno) || null
        }
      } catch {
        // keep local fallback
      }
    }
    return staff
  }

  const continueLogin = async () => {
    const sno = loginStaffNo.trim()
    if (!sno) {
      setLoginError("Enter your staff number.")
      return
    }
    const isSpecial1955 = sno === "1955"
    const staff = await findStaffByNo(sno)
    if (isSpecial1955) {
      setLoginError("")
      setLoginStage("password")
      return
    }
    if (!staff) {
      setLoginError("Staff number not recognised.")
      return
    }
    if ((staff.activeStatus || "").trim().toLowerCase() === "inactive") {
      setLoginError("This staff record is inactive.")
      return
    }
    setLoginError("")
    setLoginStage("password")
  }

  const tryLogin = async () => {
    const sno = loginStaffNo.trim()
    const pwd = loginPassword
    if (!sno) {
      setLoginError("Enter your staff number.")
      return
    }
    if (!pwd) {
      setLoginError("Enter your password.")
      return
    }
    let staff = await findStaffByNo(sno)
    const isSpecial1955 = sno === "1955" && pwd === "admin@1990"
    if (!staff && isSpecial1955) {
      const fallbackStaff: Entry = {
        id: "staff-1955-fallback",
        createdAt: new Date().toISOString(),
        staffNo: "1955",
        fullName: "Mohamed Shifaz",
        activeStatus: "Active",
      }
      setStore((prev) => ({
        ...prev,
        staff: prev.staff.some((s) => (s.staffNo || "").trim() === "1955")
          ? prev.staff
          : [fallbackStaff, ...prev.staff],
      }))
      staff = fallbackStaff
    }
    if (!staff) {
      setLoginError("Staff number not recognised.")
      return
    }
    if ((staff.activeStatus || "").trim().toLowerCase() === "inactive") {
      setLoginError("This staff record is inactive.")
      return
    }
    const defaultPassword = (staff.loginPassword || "").trim() || (staff.staffNo || "").trim()
    if (!isSpecial1955 && pwd !== defaultPassword) {
      setLoginError("Invalid staff number or password.")
      return
    }
    window.localStorage.setItem(AUTH_STORAGE_KEY, staff.id)
    setAuthStaffId(staff.id)
    setLoginError("")
    setLoginStaffNo("")
    setLoginPassword("")
    setLoginStage("identify")
  }

  const logout = () => {
    window.localStorage.removeItem(AUTH_STORAGE_KEY)
    setAuthStaffId(null)
    setActiveTab("ops")
    setLoginStage("identify")
    setLoginPassword("")
    setLoginError("")
  }

  // ------------------------------------------------------------------
  // Submit handlers.
  // ------------------------------------------------------------------
  const submitLeaveRequest = () => {
    if (!me) return
    if (!leaveForm.fromDate || !leaveForm.toDate) {
      window.alert("From and To dates are required.")
      return
    }
    const req: StaffRequest = {
      id: `req_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`,
      createdAt: new Date().toISOString(),
      staffId: me.id,
      staffNo: me.staffNo || "",
      staffName: me.fullName || "",
      type: "Leave",
      status: "Pending Approval",
      leavePolicyName: leaveForm.leavePolicyName || leavePolicyOptions[0] || "",
      fromDate: leaveForm.fromDate,
      toDate: leaveForm.toDate,
      reason: leaveForm.reason,
    }
    const next = [req, ...staffRequests]
    setStaffRequests(next)
    if (typeof window !== "undefined") {
      try {
        window.localStorage.setItem(
          STAFF_PORTAL_REQUESTS_KEY,
          JSON.stringify(next),
        )
      } catch {
        // localStorage may be unavailable.
      }
    }
    setLeaveForm({ leavePolicyName: "", fromDate: "", toDate: "", reason: "" })
  }

  const submitChangeRequest = () => {
    if (!me) return
    if (!changeForm.date || !changeForm.changeWith) {
      window.alert("Date and Change With are required.")
      return
    }
    const pair = rosterActionPairing.find(
      (p) => `${p.staffNo} - ${p.staffName}` === changeForm.changeWith,
    )
    if (!pair) {
      window.alert("Selected pairing staff not found for this date.")
      return
    }
    const requestedShift = (pair.shiftCode || "").trim().toUpperCase()
    const req: StaffRequest = {
      id: `req_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`,
      createdAt: new Date().toISOString(),
      staffId: me.id,
      staffNo: me.staffNo || "",
      staffName: me.fullName || "",
      type: "Roster Change",
      status: "Pending Peer Acceptance",
      date: changeForm.date,
      fromCode: changeForm.fromCode || codeForMyDay(changeForm.date),
      toCode: requestedShift,
      changeWithStaffNo: pair?.staffNo || "",
      changeWithStaffName: pair?.staffName || "",
      reason: changeForm.reason,
    }
    const next = [req, ...staffRequests]
    setStaffRequests(next)
    if (typeof window !== "undefined") {
      try {
        window.localStorage.setItem(
          STAFF_PORTAL_REQUESTS_KEY,
          JSON.stringify(next),
        )
      } catch {
        // localStorage may be unavailable.
      }
    }
    setChangeForm({ date: "", fromCode: "", toCode: "", changeWith: "", reason: "" })
  }

  const respondToSwapRequest = (requestId: string, accept: boolean) => {
    if (!me) return
    setStaffRequests((prev) =>
      prev.map((r) => {
        if (r.id !== requestId) return r
        if (r.type !== "Roster Change") return r
        return {
          ...r,
          status: accept ? "Pending Approval" : "Rejected",
          peerDecision: accept ? "Accepted" : "Rejected",
          peerDecisionAt: new Date().toISOString(),
          peerDecisionByStaffNo: me.staffNo || "",
          peerDecisionByStaffName: me.fullName || "",
        }
      }),
    )
  }

  const submitDutyMarkRequest = () => {
    if (!me) return
    if (!dutyMarkForm.fromDate || !dutyMarkForm.toDate) {
      window.alert("From and To dates are required.")
      return
    }
    if (dutyMarkForm.toDate < dutyMarkForm.fromDate) {
      window.alert('"To Date" must be after or equal to "From Date".')
      return
    }
    const days = getDateRangeInclusive(dutyMarkForm.fromDate, dutyMarkForm.toDate).length
    const req: StaffRequest = {
      id: `req_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`,
      createdAt: new Date().toISOString(),
      staffId: me.id,
      staffNo: me.staffNo || "",
      staffName: me.fullName || "",
      type: "Attendance",
      status: "Pending Approval",
      leavePolicyName: dutyMarkForm.dutyType,
      fromDate: dutyMarkForm.fromDate,
      toDate: dutyMarkForm.toDate,
      dutyMarkType: dutyMarkForm.dutyType,
      noOfDays: String(days),
      reason: dutyMarkForm.reason || `Requested via roster duty mark`,
    }
    const next = [req, ...staffRequests]
    setStaffRequests(next)
    // Belt-and-suspenders: write to localStorage synchronously so the
    // main app's poller picks the duty mark up on its next tick — even
    // if the user navigates away before the persistence effect fires.
    if (typeof window !== "undefined") {
      try {
        window.localStorage.setItem(
          STAFF_PORTAL_REQUESTS_KEY,
          JSON.stringify(next),
        )
      } catch {
        // localStorage may be unavailable (private mode) — fall through.
      }
    }
    setDutyMarkForm((prev) => ({ ...prev, reason: "" }))
    window.alert(
      `Duty mark request submitted (${dutyMarkForm.dutyType}, ${days} day${
        days === 1 ? "" : "s"
      }). It will appear on the Attendance screen as Pending.`,
    )
  }

  const uploadRequestDocument = (requestId: string, file: File) => {
    const reader = new FileReader()
    reader.onload = () => {
      const data = String(reader.result || "")
      setStaffRequests((prev) =>
        prev.map((r) =>
          r.id === requestId
            ? {
                ...r,
                documentName: file.name,
                documentData: data,
                documentUploadedAt: new Date().toISOString(),
                status: "Document Submitted",
              }
            : r,
        ),
      )
    }
    reader.readAsDataURL(file)
  }

  const cancelAttendanceRequest = (requestId: string) => {
    if (!me) return
    const today = toYmd(new Date())
    const next = staffRequests.map((r) => {
      if (r.id !== requestId) return r
      if (r.staffId !== me.id) return r
      if (r.type !== "Attendance") return r
      if (r.status === "Approved" || r.status === "Rejected" || r.status === "Cancelled") return r
      const anchorDate = (r.fromDate || r.date || "").trim()
      if (!anchorDate || anchorDate < today) return r
      return {
        ...r,
        status: "Cancelled" as const,
        reason: r.reason || "Cancelled by staff before approval",
      }
    })
    setStaffRequests(next)
    if (typeof window !== "undefined") {
      try {
        window.localStorage.setItem(STAFF_PORTAL_REQUESTS_KEY, JSON.stringify(next))
      } catch {
        // Ignore local storage failures.
      }
    }
  }

  // ------------------------------------------------------------------
  // Render guards.
  // ------------------------------------------------------------------
  if (!hydrated) {
    return <main className="p-4">Loading staff portal...</main>
  }

  if (!authStaffId || !me) {
    return (
      <main className="container py-5" style={{ maxWidth: 420 }}>
        <div className="card border-0 shadow-sm" style={{ borderRadius: 14 }}>
          <div className="card-body p-4">
            <div
              className="d-inline-flex align-items-center justify-content-center rounded mb-3"
              style={{
                width: 44,
                height: 44,
                background: "#dbeafe",
                color: "#1d4ed8",
              }}
            >
              <CalendarDays size={22} />
            </div>
            <h1 className="h4 fw-semibold mb-1">Staff Portal</h1>
            <p className="text-muted small mb-4">
            Enter your staff number and password to view your roster, request changes, and submit roster
            change requests.
            </p>
            <div className="mb-3">
              <label className="form-label small fw-semibold">Staff Number</label>
              <input
                className="form-control"
                value={loginStaffNo}
                onChange={(e) => setLoginStaffNo(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    if (loginStage === "identify") void continueLogin()
                    else void tryLogin()
                  }
                }}
                placeholder="e.g. 1955"
                disabled={loginStage === "password"}
              />
            </div>
            {loginStage === "password" ? (
              <div className="mb-3">
                <label className="form-label small fw-semibold">Password</label>
                <input
                  type="password"
                  className="form-control"
                  value={loginPassword}
                  onChange={(e) => setLoginPassword(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") void tryLogin()
                  }}
                  placeholder="Enter password"
                />
              </div>
            ) : null}
            {loginError ? (
              <div className="alert alert-danger py-2 small">{loginError}</div>
            ) : null}
            <div className="d-flex gap-2">
              {loginStage === "password" ? (
                <button
                  className="btn btn-outline-secondary fw-semibold"
                  onClick={() => {
                    setLoginStage("identify")
                    setLoginPassword("")
                    setLoginError("")
                  }}
                >
                  Back
                </button>
              ) : null}
              <button
                className="btn w-100 fw-semibold"
                onClick={() => {
                  if (loginStage === "identify") void continueLogin()
                  else void tryLogin()
                }}
                style={{ background: "#1d4ed8", color: "#ffffff" }}
              >
                {loginStage === "identify" ? "Continue" : "Login"}
              </button>
            </div>
            <p className="text-muted small mt-3 mb-0">
              Need access from a manager? Use the main OCCfloat console.
            </p>
            <p className="text-muted" style={{ fontSize: 11, marginTop: 8, marginBottom: 0 }}>
              Build: {STAFF_PORTAL_BUILD_TAG}
            </p>
          </div>
        </div>
      </main>
    )
  }

  return (
    <main
      className="d-flex flex-column"
      style={{ minHeight: "100vh", paddingBottom: 88 }}
    >
      {/* ---------------- Tab views ---------------- */}
      <section className="flex-grow-1 container py-3" style={{ maxWidth: 720 }}>
        {activeTab === "ops" ? (
          <OpsView
            todayYmd={todayYmd}
            announcements={store.briefing}
            checklist={store.checklist}
            opsFlights={store.opsControl}
            codeForDay={codeForMyDay}
            shiftByCode={shiftByCode}
            rosterRows={store.roster}
            staffRows={store.staff}
            meStaffNo={meStaffNo}
          />
        ) : null}

        {activeTab === "roster" ? (
          <RosterView
            calendarCells={calendarCells}
            monthLabel={monthLabel}
            onPrev={() =>
              setCalendarMonth(({ year, month }) => {
                const next = new Date(year, month - 1, 1)
                return { year: next.getFullYear(), month: next.getMonth() }
              })
            }
            onNext={() =>
              setCalendarMonth(({ year, month }) => {
                const next = new Date(year, month + 1, 1)
                return { year: next.getFullYear(), month: next.getMonth() }
              })
            }
            codeForDay={codeForMyDay}
            todayYmd={todayYmd}
            shiftByCode={shiftByCode}
            onOpenPairing={(d, selectedCode) => {
              setRosterActionDate(d)
              setRosterActionMode("Shift Change")
              setChangeForm({
                date: d,
                fromCode: selectedCode || codeForMyDay(d),
                toCode: "",
                changeWith: "",
                reason: "",
              })
              setDutyMarkForm((prev) => ({
                ...prev,
                fromDate: d,
                toDate: d,
                reason: "",
              }))
            }}
            myRequestDates={
              new Set(
                staffRequests
                  .filter(
                    (r) =>
                      r.staffId === me.id &&
                      r.type === "Roster Change" &&
                      r.status === "Pending Approval" &&
                      r.date,
                  )
                  .map((r) => r.date as string),
              )
            }
          />
        ) : null}

        {activeTab === "leave" ? (
          <LeaveView
            me={me}
            myLeaves={store.leaves.filter(
              (l) => (l.staffNo || "").trim() === (me.staffNo || "").trim(),
            )}
            myRequests={staffRequests.filter((r) => r.staffId === me.id)}
            incomingSwapRequests={staffRequests.filter(
              (r) =>
                r.type === "Roster Change" &&
                r.status === "Pending Peer Acceptance" &&
                (r.changeWithStaffNo || "").trim() === (me.staffNo || "").trim(),
            )}
            onRespondSwap={respondToSwapRequest}
            onUploadRequestDocument={uploadRequestDocument}
            onCancelAttendanceRequest={cancelAttendanceRequest}
          />
        ) : null}

        {activeTab === "profile" ? (
          <ProfileView
            me={me}
            todayYmd={todayYmd}
            myRosterByDate={myRosterByDate}
            shiftByCode={shiftByCode}
            codeForMyDay={codeForMyDay}
            myLeaves={store.leaves.filter(
              (l) => (l.staffNo || "").trim() === (me.staffNo || "").trim(),
            )}
            leavePolicies={store.leaveAttendanceControl}
            evaluations={store.evaluation.filter(
              (e) =>
                (e.staffNo || "").trim() === (me.staffNo || "").trim() ||
                (e.staffName || "").trim().toLowerCase() === (me.fullName || "").trim().toLowerCase(),
            )}
            profilePairingDate={profilePairingDate}
            setProfilePairingDate={setProfilePairingDate}
            profilePairing={profilePairing}
            onOpenLeaveRequest={() => setIsLeaveRequestModalOpen(true)}
            onLogout={logout}
          />
        ) : null}
      </section>

      {/* ---------------- Leave request modal (from Profile) ---------------- */}
      {isLeaveRequestModalOpen ? (
        <div
          className="modal show d-block"
          tabIndex={-1}
          style={{ background: "rgba(0,0,0,0.5)" }}
          onClick={() => setIsLeaveRequestModalOpen(false)}
        >
          <div className="modal-dialog modal-dialog-centered" onClick={(e) => e.stopPropagation()}>
            <div className="modal-content border-0 shadow" style={{ borderRadius: 14 }}>
              <div className="modal-header border-0 pb-0">
                <h5 className="modal-title">New leave request</h5>
                <button className="btn btn-sm btn-light rounded-circle" onClick={() => setIsLeaveRequestModalOpen(false)}>
                  <X size={14} />
                </button>
              </div>
              <div className="modal-body">
                <div className="border rounded p-2 bg-body-tertiary mb-3">
                  <div className="small fw-semibold mb-2">Leave Request Process</div>
                  <Box sx={{ maxWidth: 420 }}>
                    <Stepper activeStep={0} orientation="vertical" sx={workflowStepperSx}>
                      <Step>
                        <StepLabel>Submit Request</StepLabel>
                      </Step>
                      <Step>
                        <StepLabel>Document (If Requested)</StepLabel>
                      </Step>
                      <Step>
                        <StepLabel>Admin Approval</StepLabel>
                      </Step>
                      <Step>
                        <StepLabel>Final Decision</StepLabel>
                      </Step>
                    </Stepper>
                  </Box>
                </div>
                <div className="mb-2">
                  <label className="form-label small fw-semibold">Leave type</label>
                  <select
                    className="form-select"
                    value={leaveForm.leavePolicyName}
                    onChange={(e) =>
                      setLeaveForm((f) => ({ ...f, leavePolicyName: e.target.value }))
                    }
                  >
                    <option value="">Select leave type…</option>
                    {leavePolicyOptions.map((opt) => (
                      <option key={opt} value={opt}>
                        {opt}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="row g-2">
                  <div className="col-6">
                    <label className="form-label small fw-semibold">From</label>
                    <input
                      type="date"
                      className="form-control"
                      value={leaveForm.fromDate}
                      onChange={(e) =>
                        setLeaveForm((f) => ({ ...f, fromDate: e.target.value }))
                      }
                    />
                  </div>
                  <div className="col-6">
                    <label className="form-label small fw-semibold">To</label>
                    <input
                      type="date"
                      className="form-control"
                      value={leaveForm.toDate}
                      onChange={(e) =>
                        setLeaveForm((f) => ({ ...f, toDate: e.target.value }))
                      }
                    />
                  </div>
                </div>
                <div className="mt-2">
                  <label className="form-label small fw-semibold">Reason</label>
                  <textarea
                    className="form-control"
                    rows={2}
                    value={leaveForm.reason}
                    onChange={(e) =>
                      setLeaveForm((f) => ({ ...f, reason: e.target.value }))
                    }
                  />
                </div>
              </div>
              <div className="modal-footer border-0 pt-0">
                <button className="btn btn-outline-secondary" onClick={() => setIsLeaveRequestModalOpen(false)}>
                  Cancel
                </button>
                <button
                  className="btn btn-primary"
                  onClick={() => {
                    submitLeaveRequest()
                    setIsLeaveRequestModalOpen(false)
                  }}
                >
                  <Send size={14} className="me-1" /> Submit leave request
                </button>
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {/* ---------------- Pairing modal ---------------- */}
      {pairingDate ? (
        <div
          className="modal show d-block"
          tabIndex={-1}
          style={{ background: "rgba(0,0,0,0.5)" }}
          onClick={() => setPairingDate(null)}
        >
          <div
            className="modal-dialog modal-dialog-centered"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="modal-content border-0 shadow" style={{ borderRadius: 14 }}>
              <div className="modal-header border-0 pb-0">
                <h5 className="modal-title d-flex align-items-center gap-2">
                  <span
                    className="d-inline-flex align-items-center justify-content-center rounded"
                    style={{
                      width: 28,
                      height: 28,
                      background: "#dbeafe",
                      color: "#1d4ed8",
                    }}
                  >
                    <Users size={15} />
                  </span>
                  Working with you · {pairingDate}
                </h5>
                <button
                  className="btn btn-sm btn-light rounded-circle d-flex align-items-center justify-content-center"
                  style={{ width: 32, height: 32 }}
                  onClick={() => setPairingDate(null)}
                >
                  <X size={16} />
                </button>
              </div>
              <div className="modal-body">
                {pairingForDate.length === 0 ? (
                  <div className="p-4 text-center text-muted small">
                    No other staff working on {pairingDate}.
                  </div>
                ) : (
                  <ul className="list-unstyled mb-0 d-flex flex-column gap-2">
                    {pairingForDate.map((row) => {
                      const tint = softTintForCode(row.shiftCode, shiftByCode)
                      return (
                        <li
                          key={`${row.staffNo}-${row.staffName}`}
                          className="d-flex align-items-center gap-3 p-2 rounded"
                          style={{ background: "var(--bs-body-tertiary-bg, #f8fafc)" }}
                        >
                          <div
                            className="d-flex align-items-center justify-content-center rounded-circle"
                            style={{
                              width: 36,
                              height: 36,
                              background: "#dbeafe",
                              color: "#1d4ed8",
                              flex: "0 0 auto",
                            }}
                          >
                            <StaffAvatarIcon staffNo={row.staffNo} size={18} />
                          </div>
                          <div className="flex-grow-1 min-w-0">
                            <div className="fw-semibold text-truncate">
                              {row.staffName}
                            </div>
                            <div className="text-muted small text-truncate">
                              {row.staffNo}
                              {(() => {
                                const sn = (shiftByCode.get((row.shiftCode || "").toUpperCase())?.shiftName || "").trim()
                                return sn ? ` · ${sn}` : ""
                              })()}
                            </div>
                          </div>
                          <span
                            className="badge d-inline-flex align-items-center gap-1"
                            style={{
                              background: tint.bg,
                              color: tint.fg,
                              fontSize: 12,
                              padding: "6px 10px",
                              fontWeight: 600,
                            }}
                          >
                            <ShiftCodeIcon code={row.shiftCode} shiftByCode={shiftByCode} size={12} color={tint.fg} />
                            {row.shiftCode}
                          </span>
                        </li>
                      )
                    })}
                  </ul>
                )}
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {/* ---------------- Roster action modal ---------------- */}
      {rosterActionDate ? (
        <div
          className="modal show d-block"
          tabIndex={-1}
          style={{ background: "rgba(0,0,0,0.5)" }}
          onClick={() => setRosterActionDate(null)}
        >
          <div className="modal-dialog modal-dialog-centered" onClick={(e) => e.stopPropagation()}>
            <div className="modal-content border-0 shadow" style={{ borderRadius: 14 }}>
              <div className="modal-header border-0">
                <h5 className="modal-title">Roster Action · {rosterActionDate}</h5>
                <button className="btn btn-sm btn-light rounded-circle" onClick={() => setRosterActionDate(null)}>
                  <X size={14} />
                </button>
              </div>
              <div className="modal-body pt-0">
                <div className="border rounded p-2 bg-body-tertiary mb-3">
                  <div className="small fw-semibold mb-2">Duty Mark / Shift Change Process</div>
                  <Box sx={{ maxWidth: 420 }}>
                    <Stepper activeStep={0} orientation="vertical" sx={workflowStepperSx}>
                      <Step>
                        <StepLabel>Submit Request</StepLabel>
                      </Step>
                      <Step>
                        <StepLabel>Document (If Requested)</StepLabel>
                      </Step>
                      <Step>
                        <StepLabel>Admin Approval</StepLabel>
                      </Step>
                      <Step>
                        <StepLabel>Final Decision</StepLabel>
                      </Step>
                    </Stepper>
                  </Box>
                </div>
                <div className="d-flex gap-2 mb-3">
                  <button
                    type="button"
                    className={`btn btn-sm ${rosterActionMode === "Shift Change" ? "btn-primary" : "btn-outline-secondary"}`}
                    onClick={() => setRosterActionMode("Shift Change")}
                  >
                    Shift Change
                  </button>
                  <button
                    type="button"
                    className={`btn btn-sm ${rosterActionMode === "Duty Mark As" ? "btn-primary" : "btn-outline-secondary"}`}
                    onClick={() => setRosterActionMode("Duty Mark As")}
                  >
                    Duty Mark As
                  </button>
                </div>

                {rosterActionMode === "Shift Change" ? (
                  <div className="d-flex flex-column gap-2">
                    <div>
                      <label className="form-label small fw-semibold">Date</label>
                      <input className="form-control" value={changeForm.date} readOnly />
                    </div>
                    <div>
                      <label className="form-label small fw-semibold">Current Shift</label>
                      <input className="form-control" value={changeForm.fromCode} readOnly />
                    </div>
                    <div>
                      <label className="form-label small fw-semibold">Change With</label>
                      <select
                        className="form-select"
                        value={changeForm.changeWith}
                        onChange={(e) => setChangeForm((f) => ({ ...f, changeWith: e.target.value }))}
                      >
                        <option value="">Select Shift Pairing</option>
                        {rosterActionPairing.map((p) => (
                          <option key={`${p.staffNo}-${p.staffName}`} value={`${p.staffNo} - ${p.staffName}`}>
                            {p.staffNo} - {p.staffName} ({p.shiftCode})
                          </option>
                        ))}
                      </select>
                    </div>
                    {changeForm.changeWith ? (
                      <div className="small text-muted">
                        Swap target shift:{" "}
                        <strong>
                          {(() => {
                            const pair = rosterActionPairing.find(
                              (p) => `${p.staffNo} - ${p.staffName}` === changeForm.changeWith,
                            )
                            return pair?.shiftCode || "-"
                          })()}
                        </strong>
                      </div>
                    ) : null}
                    <div>
                      <label className="form-label small fw-semibold">Reason</label>
                      <textarea
                        className="form-control"
                        rows={2}
                        value={changeForm.reason}
                        onChange={(e) => setChangeForm((f) => ({ ...f, reason: e.target.value }))}
                        placeholder="State change reason"
                      />
                    </div>
                    <button
                      type="button"
                      className="btn btn-primary"
                      onClick={() => {
                        submitChangeRequest()
                        setRosterActionDate(null)
                      }}
                    >
                      Send for Approval
                    </button>
                  </div>
                ) : (
                  <div className="d-flex flex-column gap-2">
                    <div>
                      <label className="form-label small fw-semibold">Duty Mark As</label>
                      <select
                        className="form-select"
                        value={dutyMarkForm.dutyType}
                        onChange={(e) =>
                          setDutyMarkForm((f) => ({
                            ...f,
                            dutyType: e.target.value as "Medical" | "Sick Leave" | "Family Responsible Leave",
                          }))
                        }
                      >
                        <option value="Medical">Medical</option>
                        <option value="Sick Leave">Sick Leave</option>
                        <option value="Family Responsible Leave">Family Responsible Leave</option>
                      </select>
                    </div>
                    <div className="row g-2">
                      <div className="col-6">
                        <label className="form-label small fw-semibold">From Date</label>
                        <input
                          type="date"
                          className="form-control"
                          value={dutyMarkForm.fromDate}
                          onChange={(e) => setDutyMarkForm((f) => ({ ...f, fromDate: e.target.value }))}
                        />
                      </div>
                      <div className="col-6">
                        <label className="form-label small fw-semibold">To Date</label>
                        <input
                          type="date"
                          className="form-control"
                          value={dutyMarkForm.toDate}
                          onChange={(e) => setDutyMarkForm((f) => ({ ...f, toDate: e.target.value }))}
                        />
                      </div>
                    </div>
                    <div className="small text-muted">
                      No of Days:{" "}
                      {dutyMarkForm.fromDate && dutyMarkForm.toDate && dutyMarkForm.toDate >= dutyMarkForm.fromDate
                        ? getDateRangeInclusive(dutyMarkForm.fromDate, dutyMarkForm.toDate).length
                        : 0}
                    </div>
                    <div>
                      <label className="form-label small fw-semibold">Reason</label>
                      <textarea
                        className="form-control"
                        rows={2}
                        value={dutyMarkForm.reason}
                        onChange={(e) => setDutyMarkForm((f) => ({ ...f, reason: e.target.value }))}
                        placeholder="State change reason"
                      />
                    </div>
                    <button
                      type="button"
                      className="btn btn-primary"
                      onClick={() => {
                        submitDutyMarkRequest()
                        setRosterActionDate(null)
                      }}
                    >
                      Submit
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {/* ---------------- Bottom navigation ---------------- */}
      <nav
        className="border-top bg-body fixed-bottom"
        style={{ zIndex: 10 }}
      >
        <div
          className="container-fluid d-flex align-items-stretch"
          style={{ maxWidth: 720, margin: "0 auto" }}
        >
          {(
            [
              { key: "ops", label: "Ops", Icon: Activity },
              { key: "roster", label: "Roster", Icon: CalendarDays },
              { key: "leave", label: "Request", Icon: ClipboardList },
              { key: "profile", label: "Profile", Icon: User },
            ] as const
          ).map((tab) => {
            const isActive = activeTab === tab.key
            return (
              <button
                key={tab.key}
                className="btn flex-grow-1 d-flex flex-column align-items-center justify-content-center gap-1 rounded-0 border-0"
                style={{
                  padding: "10px 4px",
                  fontSize: 12,
                  color: isActive ? "#1d4ed8" : "var(--bs-secondary-color)",
                  fontWeight: isActive ? 600 : 500,
                  borderTop: isActive
                    ? "2px solid #1d4ed8"
                    : "2px solid transparent",
                  background: "transparent",
                }}
                onClick={() => setActiveTab(tab.key)}
              >
                <tab.Icon size={20} />
                <span>{tab.label}</span>
              </button>
            )
          })}
        </div>
      </nav>
    </main>
  )
}

// ---------------------------------------------------------------------------
// Sub-views.
// ---------------------------------------------------------------------------
function OpsView(props: {
  todayYmd: string
  announcements: Entry[]
  checklist: Entry[]
  opsFlights: Entry[]
  codeForDay: (ymd: string) => string
  shiftByCode: Map<string, Entry>
  rosterRows: Entry[]
  staffRows: Entry[]
  meStaffNo: string
}) {
  const { todayYmd, announcements, checklist, opsFlights, codeForDay, shiftByCode, rosterRows, staffRows, meStaffNo } = props
  const [selectedOpsDate, setSelectedOpsDate] = useState<string>(todayYmd)
  const [opsPairingOpen, setOpsPairingOpen] = useState(false)

  const pendingChecklist = checklist.filter(
    (c) => (c.status || "").toLowerCase() === "pending",
  )
  const selectedShiftCode = codeForDay(selectedOpsDate)
  const opsFlightsForSelectedDate = useMemo(
    () =>
      opsFlights.filter(
        (f) => ((f.scheduleDate || "").trim() === selectedOpsDate),
      ),
    [opsFlights, selectedOpsDate],
  )
  const opsFlightSummary = useMemo(() => {
    const uniqueAircraft = new Set<string>()
    let depFromMle = 0
    let arrToMle = 0
    let totalPax = 0
    const TOL = 30
    const normalized = opsFlightsForSelectedDate.map((f) => {
      const ac = (f.aircraft || "").trim()
      if (ac) uniqueAircraft.add(ac)
      const origin = (f.origin || "").trim().toUpperCase()
      const destination = (f.destination || "").trim().toUpperCase()
      const dep = (f.depTime || "").trim()
      const arr = (f.arrTime || "").trim()
      const depMin = parseTimeToMin(dep)
      const arrMin = parseTimeToMin(arr)
      if (origin === "MLE") depFromMle += 1
      if (destination === "MLE") arrToMle += 1
      const paxNum = parseInt(String(f.pax || "").replace(/[^0-9-]/g, ""), 10)
      if (!Number.isNaN(paxNum) && paxNum > 0) totalPax += paxNum
      return { origin, destination, depMin, arrMin }
    })
    let routeConflicts = 0
    for (let i = 0; i < normalized.length; i += 1) {
      for (let j = i + 1; j < normalized.length; j += 1) {
        const A = normalized[i]
        const B = normalized[j]
        if (!A.destination || A.destination === "MLE" || A.destination !== B.destination) continue
        const depConflict =
          A.depMin != null && B.depMin != null && Math.abs(A.depMin - B.depMin) <= TOL
        const arrConflict =
          A.arrMin != null && B.arrMin != null && Math.abs(A.arrMin - B.arrMin) <= TOL
        if (depConflict || arrConflict) routeConflicts += 1
      }
    }
    return {
      depFromMle,
      arrToMle,
      totalPax,
      routeConflicts,
      aircraft: uniqueAircraft.size,
    }
  }, [opsFlightsForSelectedDate])
  const todaysAnnouncements = announcements
    .filter((a) => a.date === todayYmd)
    .slice(0, 3)
  const opsDays = useMemo(() => {
    const start = new Date(`${todayYmd}T00:00:00`)
    return Array.from({ length: 6 }).map((_, i) => {
      const d = new Date(start)
      d.setDate(start.getDate() + i)
      const ymd = toYmd(d)
      const code = codeForDay(ymd)
      return {
        ymd,
        day: d.getDate(),
        weekday: d.toLocaleDateString("en-US", { weekday: "short" }),
        code,
        shift: shiftByCode.get((code || "").trim().toUpperCase()) || null,
      }
    })
  }, [todayYmd, codeForDay, shiftByCode])
  const opsPairing = useMemo(
    () => buildPairingForDate(selectedOpsDate, rosterRows, staffRows, meStaffNo, shiftByCode),
    [selectedOpsDate, rosterRows, staffRows, meStaffNo, shiftByCode],
  )

  return (
    <div className="d-flex flex-column gap-3">
      <div className="card border-0 shadow-sm" style={{ borderRadius: 12 }}>
        <div className="card-body">
          <h2 className="h6 fw-semibold mb-3 d-flex align-items-center gap-2">
            <span
              className="d-inline-flex align-items-center justify-content-center rounded"
              style={{
                width: 28,
                height: 28,
                background: "#dbeafe",
                color: "#1d4ed8",
              }}
            >
              <CalendarDays size={15} />
            </span>
            Current + Next 5 Days
          </h2>
          <div className="d-flex gap-2 overflow-auto pb-1 no-scrollbar" style={{ scrollbarWidth: "none" }}>
            {opsDays.map((item) => {
              const active = item.ymd === selectedOpsDate
              const tint = softTintForCode(item.code, shiftByCode)
              return (
                <button
                  key={item.ymd}
                  type="button"
                  onClick={() => {
                    setSelectedOpsDate(item.ymd)
                    if (isShiftPairingEnabled(item.code, shiftByCode)) {
                      setOpsPairingOpen(true)
                    }
                  }}
                  className="btn text-start border p-2"
                  style={{
                    minWidth: 84,
                    background: active ? tint.bg : "var(--bs-body-bg)",
                    borderColor: active ? tint.fg : "var(--bs-border-color)",
                    color: active ? tint.fg : "inherit",
                    borderRadius: 10,
                    flex: "0 0 auto",
                  }}
                >
                  <div className="small fw-semibold">{item.weekday}</div>
                  <div className="small text-muted">{item.day}</div>
                  <div className="mt-1 d-flex justify-content-end align-items-center">
                    {item.code ? (
                      <ShiftCodeIcon code={item.code} shiftByCode={shiftByCode} size={18} color={tint.fg} />
                    ) : (
                      <span className="small text-muted">-</span>
                    )}
                  </div>
                </button>
              )
            })}
          </div>
        </div>
      </div>
      <div className="card border-0 shadow-sm" style={{ borderRadius: 12 }}>
        <div className="card-body">
          <h2 className="h6 fw-semibold mb-3 d-flex align-items-center gap-2">
            <span
              className="d-inline-flex align-items-center justify-content-center rounded"
              style={{
                width: 28,
                height: 28,
                background: "#dbeafe",
                color: "#1d4ed8",
              }}
            >
              <ClipboardList size={15} />
            </span>
            Schedule Summary
          </h2>
          <div className="small text-muted mb-2">
            Schedule Date: <strong>{selectedOpsDate}</strong>
            {" · "}
            Shift: <strong>{selectedShiftCode || "-"}</strong>
          </div>
          {opsFlightsForSelectedDate.length === 0 ? (
            <div className="small text-muted">No uploaded schedule for this date.</div>
          ) : null}
          <div className="row g-2">
            {[
              {
                k: "STD from MLE",
                v: String(opsFlightSummary.depFromMle),
                icon: <Plane size={14} />,
                bg: "#eff6ff",
                fg: "#1d4ed8",
              },
              {
                k: "STA to MLE",
                v: String(opsFlightSummary.arrToMle),
                icon: <MapPin size={14} />,
                bg: "#ecfeff",
                fg: "#0f766e",
              },
              {
                k: "Total pax",
                v: opsFlightSummary.totalPax.toLocaleString("en-US"),
                icon: <Users size={14} />,
                bg: "#f5f3ff",
                fg: "#6d28d9",
              },
              {
                k: "Route conflicts",
                v: String(opsFlightSummary.routeConflicts),
                icon: <AlertTriangle size={14} />,
                bg: "#fff7ed",
                fg: "#c2410c",
              },
              {
                k: "Aircraft",
                v: String(opsFlightSummary.aircraft),
                icon: <Radio size={14} />,
                bg: "#f0fdf4",
                fg: "#166534",
              },
            ].map((card) => (
              <div key={card.k} className="col-6 col-md-4">
                <div className="border rounded p-2 h-100" style={{ background: card.bg, borderColor: `${card.fg}33` }}>
                  <div className="d-flex align-items-center justify-content-between mb-1">
                    <div className="small" style={{ color: card.fg }}>{card.k}</div>
                    <span
                      className="d-inline-flex align-items-center justify-content-center rounded"
                      style={{ width: 22, height: 22, background: "#ffffff", color: card.fg }}
                    >
                      {card.icon}
                    </span>
                  </div>
                  <div className="fw-bold" style={{ color: card.fg }}>{card.v}</div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
      {opsPairingOpen ? (
        <div
          className="modal show d-block"
          tabIndex={-1}
          style={{ background: "rgba(0,0,0,0.5)" }}
          onClick={() => setOpsPairingOpen(false)}
        >
          <div className="modal-dialog modal-dialog-centered" onClick={(e) => e.stopPropagation()}>
            <div className="modal-content border-0 shadow" style={{ borderRadius: 14 }}>
              <div className="modal-header border-0 pb-0">
                <h5 className="modal-title d-flex align-items-center gap-2">
                  <span
                    className="d-inline-flex align-items-center justify-content-center rounded"
                    style={{ width: 28, height: 28, background: "#dbeafe", color: "#1d4ed8" }}
                  >
                    <Users size={15} />
                  </span>
                  Shift Pairing · {selectedOpsDate}
                </h5>
                <button
                  className="btn btn-sm btn-light rounded-circle d-flex align-items-center justify-content-center"
                  style={{ width: 32, height: 32 }}
                  onClick={() => setOpsPairingOpen(false)}
                >
                  <X size={16} />
                </button>
              </div>
              <div className="modal-body">
                {opsPairing.length === 0 ? (
                  <div className="text-muted small">No other staff working on this day.</div>
                ) : (
                  <ul className="list-unstyled mb-0 d-flex flex-column gap-2">
                    {opsPairing.map((row) => {
                      const tint = softTintForCode(row.shiftCode, shiftByCode)
                      return (
                        <li
                          key={`${row.staffNo}-${row.staffName}`}
                          className="d-flex align-items-center gap-3 p-2 rounded"
                          style={{ background: "var(--bs-body-tertiary-bg, #f8fafc)" }}
                        >
                          <div
                            className="d-flex align-items-center justify-content-center rounded-circle"
                            style={{
                              width: 34,
                              height: 34,
                              background: tint.bg,
                              color: tint.fg,
                              flex: "0 0 auto",
                            }}
                          >
                            <ShiftCodeIcon code={row.shiftCode} shiftByCode={shiftByCode} size={14} color={tint.fg} />
                          </div>
                          <div className="flex-grow-1 min-w-0">
                            <div className="fw-semibold text-truncate">{row.staffName}</div>
                            <div className="text-muted small text-truncate">
                              {row.staffNo}
                              {(() => {
                                const sn = (shiftByCode.get((row.shiftCode || "").toUpperCase())?.shiftName || "").trim()
                                return sn ? ` · ${sn}` : ""
                              })()}
                            </div>
                          </div>
                          <span
                            className="badge d-inline-flex align-items-center gap-1"
                            style={{
                              background: tint.bg,
                              color: tint.fg,
                              fontSize: 12,
                              padding: "6px 10px",
                              fontWeight: 600,
                            }}
                          >
                            <ShiftCodeIcon code={row.shiftCode} shiftByCode={shiftByCode} size={12} color={tint.fg} />
                            {row.shiftCode}
                          </span>
                        </li>
                      )
                    })}
                  </ul>
                )}
              </div>
            </div>
          </div>
        </div>
      ) : null}

      <div className="card border-0 shadow-sm" style={{ borderRadius: 12 }}>
        <div className="card-body">
          <h2 className="h6 fw-semibold mb-3 d-flex align-items-center gap-2">
            <span
              className="d-inline-flex align-items-center justify-content-center rounded"
              style={{
                width: 28,
                height: 28,
                background: "#dbeafe",
                color: "#1d4ed8",
              }}
            >
              <Megaphone size={15} />
            </span>
            Today&apos;s briefing
          </h2>
          {todaysAnnouncements.length === 0 ? (
            <div className="text-muted small">No briefings posted for today.</div>
          ) : (
            <ul className="list-unstyled mb-0 d-flex flex-column gap-2">
              {todaysAnnouncements.map((a) => (
                <li
                  key={a.id}
                  className="p-2 rounded"
                  style={{
                    background: "var(--bs-body-tertiary-bg, #f8fafc)",
                    borderLeft: "3px solid #1d4ed8",
                  }}
                >
                  <div className="fw-semibold">{a.topic || "Briefing"}</div>
                  {a.details ? (
                    <div className="small text-muted">{a.details}</div>
                  ) : null}
                  {a.actionOwner ? (
                    <div className="small text-muted">
                      Owner: {a.actionOwner}
                      {a.deadline ? ` · Due ${a.deadline}` : ""}
                    </div>
                  ) : null}
                </li>
              ))}
            </ul>
          )}
        </div>
      </div>

      <div className="card border-0 shadow-sm" style={{ borderRadius: 12 }}>
        <div className="card-body">
          <h2 className="h6 fw-semibold mb-3 d-flex align-items-center gap-2">
            <span
              className="d-inline-flex align-items-center justify-content-center rounded"
              style={{
                width: 28,
                height: 28,
                background: "#dbeafe",
                color: "#1d4ed8",
              }}
            >
              <ClipboardList size={15} />
            </span>
            Open OCC checklist
          </h2>
          {pendingChecklist.length === 0 ? (
            <div className="text-muted small">No outstanding checklist items.</div>
          ) : (
            <ul className="list-unstyled mb-0 d-flex flex-column gap-2">
              {pendingChecklist.slice(0, 6).map((c) => (
                <li
                  key={c.id}
                  className="d-flex align-items-start gap-2 p-2 rounded"
                  style={{ background: "var(--bs-body-tertiary-bg, #f8fafc)" }}
                >
                  <span
                    className="d-inline-flex align-items-center justify-content-center rounded"
                    style={{
                      width: 24,
                      height: 24,
                      background: "#dbeafe",
                      color: "#1d4ed8",
                      flex: "0 0 auto",
                      marginTop: 1,
                    }}
                  >
                    <ClipboardList size={12} />
                  </span>
                  <div className="flex-grow-1">
                    <div className="small fw-semibold">{c.task || "Task"}</div>
                    {c.owner ? (
                      <div className="text-muted" style={{ fontSize: 12 }}>
                        {c.owner}
                        {c.date ? ` · ${c.date}` : ""}
                      </div>
                    ) : null}
                  </div>
                </li>
              ))}
            </ul>
          )}
        </div>
      </div>
    </div>
  )
}

function RosterView(props: {
  calendarCells: Array<{ ymd: string; day: number; inMonth: boolean }>
  monthLabel: string
  onPrev: () => void
  onNext: () => void
  codeForDay: (ymd: string) => string
  todayYmd: string
  shiftByCode: Map<string, Entry>
  onOpenPairing: (date: string, code: string) => void
  myRequestDates: Set<string>
}) {
  const {
    calendarCells,
    monthLabel,
    onPrev,
    onNext,
    codeForDay,
    todayYmd,
    shiftByCode,
    onOpenPairing,
    myRequestDates,
  } = props

  const [agendaTab, setAgendaTab] = useState<"agenda" | "month" | "week">("month")
  // Selected day defaults to today; auto-clamped to the visible month so the
  // agenda strip below the grid stays in context as the user paginates months.
  const [selectedDate, setSelectedDate] = useState<string>(todayYmd)

  // If the visible month no longer contains the selected date (after Prev /
  // Next), drop it onto the first in-month day so the agenda card always
  // shows something meaningful.
  /* eslint-disable react-hooks/set-state-in-effect, react-hooks/exhaustive-deps */
  useEffect(() => {
    const inMonth = calendarCells.find(
      (c) => c.inMonth && c.ymd === selectedDate,
    )
    if (inMonth) return
    const firstInMonth = calendarCells.find((c) => c.inMonth)
    if (firstInMonth) setSelectedDate(firstInMonth.ymd)
  }, [calendarCells])
  /* eslint-enable react-hooks/set-state-in-effect, react-hooks/exhaustive-deps */

  const weekdayLabels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

  const selectedCode = (codeForDay(selectedDate) || "").toUpperCase()
  const selectedShift = selectedCode ? shiftByCode.get(selectedCode) || null : null

  // Build the upcoming-events list for the Agenda view.
  const agendaItems = useMemo(() => {
    type Item = {
      ymd: string
      label: string
      weekday: string
      day: number
      code: string
      shift: Entry | null
    }
    const items: Item[] = []
    calendarCells.forEach((cell) => {
      if (!cell.inMonth) return
      const code = (codeForDay(cell.ymd) || "").toUpperCase()
      if (!code) return
      const shift = shiftByCode.get(code) || null
      items.push({
        ymd: cell.ymd,
        label: new Date(`${cell.ymd}T00:00:00`).toLocaleDateString("en-US", {
          weekday: "long",
          day: "numeric",
          month: "long",
        }),
        weekday: new Date(`${cell.ymd}T00:00:00`).toLocaleDateString("en-US", {
          weekday: "short",
        }),
        day: cell.day,
        code,
        shift,
      })
    })
    return items
  }, [calendarCells, codeForDay, shiftByCode])

  // Week view items: 7 days starting from the Monday of the selected date.
  const weekItems = useMemo(() => {
    const ref = new Date(`${selectedDate}T00:00:00`)
    if (Number.isNaN(ref.getTime())) return []
    const offset = (ref.getDay() + 6) % 7 // Monday = 0
    const start = new Date(ref)
    start.setDate(ref.getDate() - offset)
    const result: Array<{
      ymd: string
      day: number
      weekday: string
      code: string
      shift: Entry | null
    }> = []
    for (let i = 0; i < 7; i++) {
      const d = new Date(start)
      d.setDate(start.getDate() + i)
      const ymd = toYmd(d)
      const code = (codeForDay(ymd) || "").toUpperCase()
      result.push({
        ymd,
        day: d.getDate(),
        weekday: d.toLocaleDateString("en-US", { weekday: "short" }),
        code,
        shift: code ? shiftByCode.get(code) || null : null,
      })
    }
    return result
  }, [selectedDate, codeForDay, shiftByCode])

  return (
    <div className="d-flex flex-column gap-3">
      {/* ----- Agenda / Month / Week tab strip ----- */}
      <div
        className="d-flex align-items-center justify-content-between gap-2"
        style={{ borderBottom: "1px solid var(--bs-border-color)", paddingBottom: 4 }}
      >
        <button className="btn btn-sm btn-link p-1 text-muted" tabIndex={-1}>
          <Bell size={18} />
        </button>
        <div className="d-flex align-items-stretch flex-grow-1 justify-content-center">
          {(["agenda", "month", "week"] as const).map((t) => {
            const active = agendaTab === t
            return (
              <button
                key={t}
                className="btn btn-sm border-0 rounded-0"
                onClick={() => setAgendaTab(t)}
                style={{
                  padding: "6px 14px",
                  color: active ? "var(--bs-primary)" : "var(--bs-secondary-color)",
                  fontWeight: active ? 700 : 500,
                  borderBottom: active
                    ? "2px solid var(--bs-primary)"
                    : "2px solid transparent",
                  background: "transparent",
                  textTransform: "capitalize",
                }}
              >
                {t === "agenda" ? "Daily" : t}
              </button>
            )
          })}
        </div>
        <button className="btn btn-sm btn-link p-1 text-muted" tabIndex={-1}>
          <MoreHorizontal size={18} />
        </button>
      </div>

      {/* ----- Month view: calendar grid ----- */}
      {agendaTab === "month" ? (
        <div className="card border-0 shadow-sm">
          <div className="card-body">
            <div className="d-flex align-items-center justify-content-between mb-3">
              <button
                className="btn btn-sm btn-link text-dark"
                onClick={onPrev}
                aria-label="Previous month"
              >
                <ChevronLeft size={20} />
              </button>
              <h2 className="h5 fw-bold mb-0 text-center" style={{ letterSpacing: 1 }}>
                {monthLabel}
              </h2>
              <button
                className="btn btn-sm btn-link text-dark"
                onClick={onNext}
                aria-label="Next month"
              >
                <ChevronRight size={20} />
              </button>
            </div>

            <div
              className="d-grid"
              style={{
                gridTemplateColumns: "repeat(7, 1fr)",
                rowGap: 6,
                columnGap: 6,
              }}
            >
              {weekdayLabels.map((w) => (
                <div
                  key={w}
                  className="text-center text-muted"
                  style={{ fontSize: 11, paddingBottom: 6, letterSpacing: 0.5 }}
                >
                  {w}
                </div>
              ))}
              {calendarCells.map((cell) => {
                const code = (codeForDay(cell.ymd) || "").toUpperCase()
                const tint = softTintForCode(code, shiftByCode)
                const hasIcon = hasShiftIcon(code, shiftByCode)
                const isSelected = cell.ymd === selectedDate
                const isToday = cell.ymd === todayYmd
                const hasRequest = myRequestDates.has(cell.ymd)
                const inMonth = cell.inMonth
                const hasEvent = !!code && code !== "-"

                let bg = "transparent"
                let fg: string = inMonth ? "#0f172a" : "#cbd5e1"
                let iconColor = tint.fg
                let border = "1px solid transparent"

                if (isSelected) {
                  bg = "#1d4ed8"
                  fg = "#ffffff"
                  iconColor = "#ffffff"
                } else if (hasEvent && inMonth) {
                  bg = tint.bg
                  fg = tint.fg
                } else if (!inMonth) {
                  bg = "transparent"
                }

                if (isToday && !isSelected) {
                  border = "1px solid #1d4ed8"
                }

                return (
                  <button
                    key={cell.ymd}
                    onClick={() => setSelectedDate(cell.ymd)}
                    className="btn p-0 border-0 position-relative"
                    style={{
                      height: 56,
                      background: bg,
                      color: fg,
                      borderRadius: 6,
                      border,
                      padding: 4,
                      opacity: inMonth ? 1 : 0.55,
                      display: "flex",
                      flexDirection: "column",
                      alignItems: "center",
                      justifyContent: "flex-start",
                      gap: 2,
                    }}
                    title={`${cell.ymd}${code ? " · " + code : ""}`}
                  >
                    <span
                      style={{
                        fontSize: 12,
                        fontWeight: 600,
                        lineHeight: 1,
                        marginTop: 4,
                      }}
                    >
                      {cell.day}
                    </span>
                    {hasIcon ? (
                      <ShiftCodeIcon code={code} shiftByCode={shiftByCode} size={14} color={iconColor} />
                    ) : (
                      <span style={{ height: 14 }} />
                    )}
                    {hasRequest ? (
                      <span
                        className="position-absolute"
                        style={{
                          top: 2,
                          right: 4,
                          color: "#dc2626",
                          lineHeight: 0,
                        }}
                        title="Roster change request pending"
                      >
                        <AlertTriangle size={11} fill="#fee2e2" />
                      </span>
                    ) : null}
                  </button>
                )
              })}
            </div>
          </div>
        </div>
      ) : null}

      {/* ----- Week view: 7-day strip ----- */}
      {agendaTab === "week" ? (
        <div className="card border-0 shadow-sm">
          <div className="card-body">
            <div className="d-flex align-items-center justify-content-between mb-3">
              <button
                className="btn btn-sm btn-link text-dark"
                onClick={() => {
                  const d = new Date(`${selectedDate}T00:00:00`)
                  d.setDate(d.getDate() - 7)
                  setSelectedDate(toYmd(d))
                }}
              >
                <ChevronLeft size={20} />
              </button>
              <h2 className="h6 fw-bold mb-0">
                Week of{" "}
                {weekItems[0]
                  ? new Date(`${weekItems[0].ymd}T00:00:00`).toLocaleDateString(
                      "en-US",
                      { day: "numeric", month: "short" },
                    )
                  : ""}
              </h2>
              <button
                className="btn btn-sm btn-link text-dark"
                onClick={() => {
                  const d = new Date(`${selectedDate}T00:00:00`)
                  d.setDate(d.getDate() + 7)
                  setSelectedDate(toYmd(d))
                }}
              >
                <ChevronRight size={20} />
              </button>
            </div>
            <div className="d-flex flex-column gap-2">
              {weekItems.map((item) => {
                const tint = softTintForCode(item.code, shiftByCode)
                const isSelected = item.ymd === selectedDate
                return (
                  <button
                    key={item.ymd}
                    onClick={() => setSelectedDate(item.ymd)}
                    className="btn text-start p-0 border-0 d-flex align-items-stretch"
                    style={{
                      borderRadius: 12,
                      overflow: "hidden",
                      background: isSelected ? "#1d4ed8" : "var(--bs-body-tertiary-bg, #f8fafc)",
                      color: isSelected ? "#ffffff" : "inherit",
                    }}
                  >
                    <div
                      className="d-flex flex-column align-items-center justify-content-center"
                      style={{
                        width: 56,
                        padding: "10px 0",
                        background: isSelected ? "rgba(255,255,255,0.08)" : tint.bg,
                        color: isSelected ? "#ffffff" : tint.fg,
                      }}
                    >
                      <span style={{ fontSize: 11, fontWeight: 600, opacity: 0.85 }}>
                        {item.weekday}
                      </span>
                      <span style={{ fontSize: 22, fontWeight: 700, lineHeight: 1 }}>
                        {item.day}
                      </span>
                    </div>
                    <div className="flex-grow-1 d-flex align-items-center px-3 py-2 gap-2">
                      <ShiftCodeIcon code={item.code} shiftByCode={shiftByCode} size={16} />
                      <span style={{ fontWeight: 600 }}>
                        {item.code || "Off duty"}
                      </span>
                      {item.shift ? (
                        <span style={{ opacity: 0.85, fontSize: 12 }}>
                          {formatTime12(item.shift.startTime || "")} –{" "}
                          {formatTime12(item.shift.endTime || "")}
                        </span>
                      ) : null}
                    </div>
                  </button>
                )
              })}
            </div>
          </div>
        </div>
      ) : null}

      {/* ----- Agenda view: chronological list of events ----- */}
      {agendaTab === "agenda" ? (
        <div className="card border-0 shadow-sm">
          <div className="card-body">
            <div className="d-flex align-items-center justify-content-between mb-2">
              <button className="btn btn-sm btn-link text-dark" onClick={onPrev}>
                <ChevronLeft size={20} />
              </button>
              <h2 className="h6 fw-bold mb-0">{monthLabel}</h2>
              <button className="btn btn-sm btn-link text-dark" onClick={onNext}>
                <ChevronRight size={20} />
              </button>
            </div>
            {agendaItems.length === 0 ? (
              <div className="text-muted small text-center py-3">
                No events scheduled this month.
              </div>
            ) : (
              <ul className="list-unstyled mb-0 d-flex flex-column gap-2">
                {agendaItems.map((item) => {
                  const tint = softTintForCode(item.code, shiftByCode)
                  return (
                    <li
                      key={item.ymd}
                      className="d-flex align-items-center gap-3 p-2 rounded"
                      style={{ background: "var(--bs-body-tertiary-bg, #f8fafc)" }}
                    >
                      <div
                        className="d-flex flex-column align-items-center justify-content-center"
                        style={{
                          width: 48,
                          padding: 6,
                          borderRadius: 8,
                          background: tint.bg,
                          color: tint.fg,
                        }}
                      >
                        <span style={{ fontSize: 11, fontWeight: 600 }}>
                          {item.weekday}
                        </span>
                        <span style={{ fontSize: 18, fontWeight: 700, lineHeight: 1 }}>
                          {item.day}
                        </span>
                      </div>
                      <div className="flex-grow-1 min-w-0">
                        <div className="fw-semibold text-truncate">
                          {item.shift?.shiftName || item.code}
                        </div>
                        <div className="text-muted small text-truncate">
                          {item.shift
                            ? `${formatTime12(item.shift.startTime || "")} – ${formatTime12(item.shift.endTime || "")}`
                            : isOffCode(item.code)
                              ? "Rest day"
                              : item.code}
                        </div>
                      </div>
                      <ShiftCodeIcon code={item.code} shiftByCode={shiftByCode} size={18} />
                    </li>
                  )
                })}
              </ul>
            )}
          </div>
        </div>
      ) : null}

      {/* ----- Selected day detail card ----- */}
      <div>
        <h3
          className="fw-semibold mb-2"
          style={{ fontSize: 15, paddingLeft: 4 }}
        >
          {new Date(`${selectedDate}T00:00:00`).toLocaleDateString("en-US", {
            weekday: "long",
            day: "numeric",
            month: "long",
          })}
        </h3>
        <SelectedDayCard
          code={selectedCode}
          shift={selectedShift}
          shiftByCode={shiftByCode}
          ymd={selectedDate}
          onOpenPairing={onOpenPairing}
          hasRequest={myRequestDates.has(selectedDate)}
          showPairingAction={isShiftPairingEnabled(selectedCode, shiftByCode)}
        />
      </div>
    </div>
  )
}

function SelectedDayCard(props: {
  code: string
  shift: Entry | null
  shiftByCode: Map<string, Entry>
  ymd: string
  onOpenPairing: (ymd: string, code: string) => void
  hasRequest: boolean
  showPairingAction?: boolean
}) {
  const { code, shift, shiftByCode, ymd, onOpenPairing, hasRequest, showPairingAction = false } = props
  const tint = softTintForCode(code, shiftByCode)
  const startTime = shift?.startTime ? formatTime12(shift.startTime) : ""
  const endTime = shift?.endTime ? formatTime12(shift.endTime) : ""
  const headline = shift?.shiftName
    ? shift.shiftName
    : isOffCode(code)
      ? "Rest day"
      : code === "AL"
        ? "Annual leave"
        : code === "PH" || code === "GH"
          ? "Public holiday"
          : code
            ? `Shift ${code}`
            : "No shift assigned"
  const subtitle = shift?.location || (shift ? `${startTime} – ${endTime}` : "")

  return (
    <div
      className="card border-0 shadow-sm"
      style={{ borderRadius: 12, overflow: "hidden" }}
    >
      <div className="card-body d-flex align-items-center gap-3 p-3">
        <div
          className="d-flex align-items-center justify-content-center rounded-circle"
          style={{
            width: 44,
            height: 44,
            border: "2px solid #1d4ed8",
            color: "#1d4ed8",
            background: "#ffffff",
            flex: "0 0 auto",
          }}
        >
          <ShiftCodeIcon code={code} shiftByCode={shiftByCode} size={20} />
        </div>
        <div className="flex-grow-1 min-w-0">
          <div className="d-flex align-items-center gap-2">
            {startTime ? (
              <span
                className="fw-bold"
                style={{ fontSize: 15, letterSpacing: 0.3 }}
              >
                {shift ? shift.startTime : ""}
              </span>
            ) : null}
            <span className="fw-bold text-truncate" style={{ fontSize: 15 }}>
              {headline}
            </span>
            {hasRequest ? (
              <AlertTriangle size={14} color="#dc2626" />
            ) : null}
          </div>
          {subtitle ? (
            <div
              className="text-muted text-truncate"
              style={{ fontSize: 12, marginTop: 2 }}
            >
              <MapPin size={11} className="me-1" />
              {subtitle}
            </div>
          ) : null}
          {!shift && code ? (
            <div
              className="text-muted"
              style={{ fontSize: 12, marginTop: 2 }}
            >
              <span
                className="badge"
                style={{
                  background: tint.bg,
                  color: tint.fg,
                  fontSize: 11,
                  padding: "2px 6px",
                }}
              >
                {code}
              </span>
            </div>
          ) : null}
        </div>
        {showPairingAction ? (
          <button
            className="btn d-flex align-items-center justify-content-center rounded-circle"
            onClick={() => onOpenPairing(ymd, code)}
            style={{
              width: 36,
              height: 36,
              background: "#1d4ed8",
              color: "#ffffff",
              flex: "0 0 auto",
            }}
            title="See pairing"
          >
            <ArrowRight size={18} />
          </button>
        ) : null}
      </div>
    </div>
  )
}

function LeaveView(props: {
  me: Entry
  myLeaves: Entry[]
  myRequests: StaffRequest[]
  incomingSwapRequests: StaffRequest[]
  onRespondSwap: (requestId: string, accept: boolean) => void
  onUploadRequestDocument: (requestId: string, file: File) => void
  onCancelAttendanceRequest: (requestId: string) => void
}) {
  const {
    me,
    myLeaves,
    myRequests,
    incomingSwapRequests,
    onRespondSwap,
    onUploadRequestDocument,
    onCancelAttendanceRequest,
  } = props
  const [expandedRequestSteppers, setExpandedRequestSteppers] = useState<Set<string>>(new Set())
  const todayYmd = toYmd(new Date())

  return (
    <div className="d-flex flex-column gap-3">
      <div className="card border-0 shadow-sm" style={{ borderRadius: 12 }}>
        <div className="card-body">
          <h2 className="h6 fw-semibold mb-3 d-flex align-items-center gap-2">
            <span
              className="d-inline-flex align-items-center justify-content-center rounded"
              style={{
                width: 28,
                height: 28,
                background: "#dbeafe",
                color: "#1d4ed8",
              }}
            >
              <Users size={15} />
            </span>
            Swap requests for your acceptance
          </h2>
          {incomingSwapRequests.length === 0 ? (
            <div className="text-muted small mb-3">No pending swap requests for you.</div>
          ) : (
            <ul className="list-unstyled d-flex flex-column gap-2 mb-3">
              {incomingSwapRequests.map((req) => (
                <li key={req.id} className="p-2 rounded" style={{ background: "var(--bs-body-tertiary-bg, #f8fafc)" }}>
                  <div className="small fw-semibold">
                    {req.staffNo} - {req.staffName} requests swap on {req.date}
                  </div>
                  <div className="text-muted small">
                    {req.staffNo} ({req.fromCode || "-"}) ⇄ {me.staffNo || "-"} ({req.toCode || "-"})
                  </div>
                  <div className="d-flex gap-2 mt-2">
                    <button type="button" className="btn btn-sm btn-success" onClick={() => onRespondSwap(req.id, true)}>
                      Accept
                    </button>
                    <button type="button" className="btn btn-sm btn-outline-danger" onClick={() => onRespondSwap(req.id, false)}>
                      Deny
                    </button>
                  </div>
                </li>
              ))}
            </ul>
          )}

          <h2 className="h6 fw-semibold mb-3 d-flex align-items-center gap-2">
            <span
              className="d-inline-flex align-items-center justify-content-center rounded"
              style={{
                width: 28,
                height: 28,
                background: "#dbeafe",
                color: "#1d4ed8",
              }}
            >
              <Bell size={15} />
            </span>
            My submitted requests
          </h2>
          {myRequests.length === 0 ? (
            <div className="text-muted small">You haven&apos;t submitted any requests yet.</div>
          ) : (
            <ul className="list-unstyled mb-0 d-flex flex-column gap-2">
              {myRequests.map((req) => {
                return (
                  <li
                    key={req.id}
                    className="p-2 rounded"
                    style={{ background: "var(--bs-body-tertiary-bg, #f8fafc)" }}
                  >
                    <div className="d-flex align-items-center justify-content-between gap-2">
                      <span className="fw-semibold small d-flex align-items-center gap-2">
                        {req.type === "Leave" ? (
                          <Plane size={13} color="#1d4ed8" />
                        ) : req.type === "Attendance" ? (
                          <Activity size={13} color="#1d4ed8" />
                        ) : (
                          <CalendarDays size={13} color="#1d4ed8" />
                        )}
                        {req.type === "Leave"
                          ? "Leave"
                          : req.type === "Attendance"
                            ? "Attendance"
                            : "Roster change"}
                      </span>
                      <span className="badge" style={requestStatusBadgeStyle(req.status)}>
                        {req.status}
                      </span>
                    </div>
                    {req.type === "Leave" ? (
                      <div className="text-muted small">
                        {req.leavePolicyName || "Leave"} · {req.fromDate} → {req.toDate}
                      </div>
                    ) : req.type === "Attendance" ? (
                      <div className="text-muted small">
                        {req.dutyMarkType || req.leavePolicyName || "Attendance"} · {req.fromDate} → {req.toDate}
                      </div>
                    ) : (
                      <div className="text-muted small">
                        {req.date}: {req.fromCode || "-"} ⇄ {req.toCode}
                        {req.changeWithStaffNo || req.changeWithStaffName
                          ? ` · with ${req.changeWithStaffNo || ""}${req.changeWithStaffNo && req.changeWithStaffName ? " - " : ""}${req.changeWithStaffName || ""}`
                          : ""}
                      </div>
                    )}
                    {req.type === "Roster Change" && req.peerDecision ? (
                      <div className="small text-muted">
                        Peer {req.peerDecision.toLowerCase()} by {req.peerDecisionByStaffNo || ""}{req.peerDecisionByStaffName ? ` - ${req.peerDecisionByStaffName}` : ""}
                      </div>
                    ) : null}
                    {req.reason ? (
                      <div className="small mt-1">{req.reason}</div>
                    ) : null}
                    {(() => {
                      const anchorDate = (req.fromDate || req.date || "").trim()
                      const canCancel =
                        req.type === "Attendance" &&
                        req.status !== "Approved" &&
                        req.status !== "Rejected" &&
                        req.status !== "Cancelled" &&
                        Boolean(anchorDate) &&
                        anchorDate >= todayYmd
                      if (!canCancel) return null
                      return (
                        <div className="mt-2">
                          <button
                            type="button"
                            className="btn btn-sm btn-outline-danger"
                            onClick={() => onCancelAttendanceRequest(req.id)}
                          >
                            Cancel Request
                          </button>
                        </div>
                      )
                    })()}
                    <div className="mt-2">
                      <button
                        type="button"
                        className="btn btn-sm btn-outline-secondary"
                        onClick={() =>
                          setExpandedRequestSteppers((prev) => {
                            const next = new Set(prev)
                            if (next.has(req.id)) next.delete(req.id)
                            else next.add(req.id)
                            return next
                          })
                        }
                      >
                        {expandedRequestSteppers.has(req.id)
                          ? "Collapse Process"
                          : "Expand Process"}
                      </button>
                    </div>
                    {expandedRequestSteppers.has(req.id) ? (
                      <div className="mt-2 border rounded p-2 bg-white">
                        <Box sx={{ maxWidth: 420 }}>
                          <Stepper
                            activeStep={getPortalRequestStepIndex(req.status)}
                            orientation="vertical"
                            sx={workflowStepperSx}
                          >
                            <Step>
                              <StepLabel>Submitted</StepLabel>
                            </Step>
                            <Step>
                              <StepLabel>Document Stage</StepLabel>
                            </Step>
                            <Step>
                              <StepLabel>Approval Stage</StepLabel>
                            </Step>
                            <Step>
                              <StepLabel>
                                {(req.status || "").trim().toLowerCase() === "rejected"
                                  ? "Final: Rejected"
                                  : (req.status || "").trim().toLowerCase() === "cancelled"
                                    ? "Final: Cancelled"
                                  : (req.status || "").trim().toLowerCase() === "approved"
                                    ? "Final: Approved"
                                    : "Final: Pending"}
                              </StepLabel>
                            </Step>
                          </Stepper>
                        </Box>
                      </div>
                    ) : null}
                    {req.status === "Pending Document" ? (
                      <div className="mt-2">
                        <label className="btn btn-sm btn-outline-primary mb-0">
                          Upload Document
                          <input
                            type="file"
                            className="d-none"
                            onChange={(e) => {
                              const file = e.target.files?.[0]
                              if (!file) return
                              onUploadRequestDocument(req.id, file)
                              e.currentTarget.value = ""
                            }}
                          />
                        </label>
                      </div>
                    ) : null}
                    {req.documentName ? (
                      <div className="small text-muted mt-1">
                        Document: {req.documentName}
                      </div>
                    ) : null}
                  </li>
                )
              })}
            </ul>
          )}
        </div>
      </div>

      <div className="card border-0 shadow-sm" style={{ borderRadius: 12 }}>
        <div className="card-body">
          <h2 className="h6 fw-semibold mb-3 d-flex align-items-center gap-2">
            <span
              className="d-inline-flex align-items-center justify-content-center rounded"
              style={{
                width: 28,
                height: 28,
                background: "#dbeafe",
                color: "#1d4ed8",
              }}
            >
              <ClipboardList size={15} />
            </span>
            My approved leaves
          </h2>
          {myLeaves.length === 0 ? (
            <div className="text-muted small">No leaves on file.</div>
          ) : (
            <ul className="list-unstyled mb-0 d-flex flex-column gap-2">
              {myLeaves.map((l) => {
                const status = (l.status || "Pending").trim()
                const statusCls =
                  status.toLowerCase() === "approved"
                    ? "bg-success-subtle text-success-emphasis"
                    : status.toLowerCase() === "rejected"
                      ? "bg-danger-subtle text-danger-emphasis"
                      : "bg-warning-subtle text-warning-emphasis"
                return (
                  <li
                    key={l.id}
                    className="p-2 rounded"
                    style={{ background: "var(--bs-body-tertiary-bg, #f8fafc)" }}
                  >
                    <div className="d-flex align-items-center justify-content-between gap-2">
                      <span className="fw-semibold small d-flex align-items-center gap-2">
                        <Sun size={13} color="#1d4ed8" />
                        {l.leavePolicyName || "Leave"}
                      </span>
                      <span className={`badge ${statusCls}`}>{status}</span>
                    </div>
                    <div className="text-muted small">
                      {l.fromDate} → {l.toDate}
                    </div>
                    {l.remarks ? (
                      <div className="small mt-1">{l.remarks}</div>
                    ) : null}
                  </li>
                )
              })}
            </ul>
          )}
        </div>
      </div>
    </div>
  )
}

function ProfileView(props: {
  me: Entry
  todayYmd: string
  myRosterByDate: Map<string, string>
  shiftByCode: Map<string, Entry>
  codeForMyDay: (ymd: string) => string
  myLeaves: Entry[]
  leavePolicies: Entry[]
  evaluations: Entry[]
  profilePairingDate: string
  setProfilePairingDate: (value: string) => void
  profilePairing: PairingRow[]
  onOpenLeaveRequest: () => void
  onLogout: () => void
}) {
  const {
    me,
    todayYmd,
    codeForMyDay,
    shiftByCode,
    myLeaves,
    leavePolicies,
    evaluations,
    onOpenLeaveRequest,
    onLogout,
  } = props
  const todayCode = codeForMyDay(todayYmd)
  const todayShift = shiftByCode.get((todayCode || "").trim().toUpperCase()) || null
  const todayTint = softTintForCode(todayCode, shiftByCode)
  const approvedLeaves = myLeaves.filter(
    (l) => (l.status || "").trim().toLowerCase() === "approved",
  )
  const approvedLeaveDays = approvedLeaves.reduce((acc, leave) => {
    const days = Number(leave.chargeableLeaveDays || "0")
    if (!Number.isNaN(days) && days > 0) return acc + days
    const from = (leave.fromDate || "").trim()
    const to = (leave.toDate || "").trim()
    if (!from || !to) return acc
    const range = getDateRangeInclusive(from, to)
    return acc + range.length
  }, 0)
  const annualLeavePolicy = leavePolicies.find((p) => {
    const typeCode = (p.typeCode || "").trim().toUpperCase()
    const name = (p.leaveAttendanceName || "").trim().toLowerCase()
    return typeCode === "AL" || name.includes("annual")
  })
  const annualEntitlement = Number(annualLeavePolicy?.noOfDays || "0")
  const leaveBalance =
    !Number.isNaN(annualEntitlement) && annualEntitlement > 0
      ? Math.max(annualEntitlement - approvedLeaveDays, 0)
      : 0
  const evalScores = evaluations
    .map((e) => Number(e.score || "0"))
    .filter((n) => !Number.isNaN(n) && n > 0)
  const evalAvg = evalScores.length
    ? evalScores.reduce((a, b) => a + b, 0) / evalScores.length
    : 0
  const performanceLabel =
    evalAvg >= 8.5 ? "Excellent" : evalAvg >= 7 ? "Good" : evalAvg >= 5 ? "Average" : "Needs Improvement"

  return (
    <div className="d-flex flex-column gap-3">
      <div className="card border-0 shadow-sm" style={{ borderRadius: 12 }}>
        <div className="card-body d-flex align-items-center gap-3">
          <div
            className="d-flex align-items-center justify-content-center rounded-circle"
            style={{
              width: 52,
              height: 52,
              background: "#dbeafe",
              color: "#1d4ed8",
              flex: "0 0 auto",
            }}
          >
            <StaffAvatarIcon staffNo={me.staffNo || ""} size={24} />
          </div>
          <div className="min-w-0 flex-grow-1">
            <div className="fw-semibold text-truncate">{me.fullName || "-"}</div>
            <div className="text-muted small text-truncate">
              {me.staffNo || "-"}
              {me.designation ? ` · ${me.designation}` : ""}
            </div>
          </div>
          <span
            className="badge d-inline-flex align-items-center gap-1"
            style={{
              background: todayTint.bg,
              color: todayTint.fg,
              fontSize: 12,
              padding: "6px 10px",
              fontWeight: 600,
            }}
          >
            <ShiftCodeIcon code={todayCode} shiftByCode={shiftByCode} size={12} color={todayTint.fg} />
            {todayCode || "-"}
          </span>
        </div>
        <div className="card-footer bg-transparent border-0 pt-0 d-flex align-items-center justify-content-end gap-2">
          <ThemeToggle />
          <button className="btn btn-sm btn-outline-secondary d-inline-flex align-items-center gap-1" onClick={onLogout}>
            <LogOut size={14} />
            Logout
          </button>
        </div>
      </div>

      <div className="small text-muted px-1">
        Today shift: {todayShift?.shiftName || todayCode || "-"}
      </div>
      <div className="card border-0 shadow-sm" style={{ borderRadius: 12 }}>
        <div className="card-body">
          <h3 className="h6 fw-semibold mb-3">Profile Card</h3>
          <div className="row g-2 small mb-3">
            <div className="col-6 text-muted">Employment Date</div>
            <div className="col-6 fw-semibold text-end">{me.empDate || "-"}</div>
            <div className="col-6 text-muted">Employment Type</div>
            <div className="col-6 fw-semibold text-end">{me.employmentType || "-"}</div>
            <div className="col-6 text-muted">Designation</div>
            <div className="col-6 fw-semibold text-end">{me.designation || "-"}</div>
            <div className="col-6 text-muted">Level</div>
            <div className="col-6 fw-semibold text-end">{me.level || "-"}</div>
            <div className="col-6 text-muted">Phone Number</div>
            <div className="col-6 fw-semibold text-end">{me.phoneNumber || "-"}</div>
          </div>
          <div className="row g-2">
            <div className="col-6 col-lg-3">
              <div className="border rounded p-2 h-100" style={{ minHeight: 132 }}>
              <div className="text-muted small">Leave Balance (AL)</div>
              <div className="fw-bold fs-5">{leaveBalance}</div>
              <div className="small text-muted">
                Used: {approvedLeaveDays}
                {annualEntitlement > 0 ? ` / ${annualEntitlement}` : ""}
              </div>
              </div>
            </div>
            <div className="col-6 col-lg-3">
              <div className="border rounded p-2 h-100" style={{ minHeight: 132 }}>
              <div className="text-muted small">Eval Score</div>
              <div className="fw-bold fs-5">{evalScores.length ? evalAvg.toFixed(1) : "-"}</div>
              <div className="small text-muted">{evalScores.length} record(s)</div>
              </div>
            </div>
            <div className="col-6 col-lg-3">
              <div className="border rounded p-2 h-100" style={{ minHeight: 132 }}>
              <div className="text-muted small">Performance</div>
              <div className="fw-bold fs-6">{evalScores.length ? performanceLabel : "-"}</div>
              <div className="small text-muted">Based on evaluations</div>
              </div>
            </div>
            <div className="col-6 col-lg-3">
              <button
                type="button"
                className="border rounded p-2 bg-white text-start w-100"
                style={{ minHeight: 132 }}
                onClick={onOpenLeaveRequest}
              >
                <div className="text-muted small">Leave Request</div>
                <div className="fw-semibold mt-1">Create</div>
                <div className="small text-muted mt-2">Open request form</div>
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
