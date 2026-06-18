"use client"

import { useEffect, useMemo, useRef, useState } from "react"
import {
  Activity,
  Award,
  CalendarCheck,
  CalendarDays,
  ClipboardList,
  Clock,
  Flag,
  GraduationCap,
  Layers,
  ListChecks,
  Megaphone,
  Mic,
  Moon,
  Headset,
  Plane,
  ShieldAlert,
  Settings,
  Clipboard,
  User,
  Sun,
  Sunrise,
  Sunset,
  Shield,
  TrendingUp,
  UserCog,
  UserPlus,
  Users,
  LogOut,
  X,
} from "lucide-react"
import { Button } from "@/components/ui/button"
import { ThemeToggle } from "@/components/theme-toggle"
import { createClient as createSupabaseClient } from "@/utils/supabase/client"
import { DataGrid, type GridColDef, type GridRenderCellParams } from "@mui/x-data-grid"
import Box from "@mui/material/Box"
import Stepper from "@mui/material/Stepper"
import Step from "@mui/material/Step"
import StepLabel from "@mui/material/StepLabel"
import OpsControlTimeline, { type FlightEntry } from "@/components/ops-control-timeline"
import AircraftReportModule from "@/components/aircraft-report-module"

type ModuleKey =
  | "dashboard"
  | "crewRostering"
  | "crewLeavePlanner"
  | "crewDataBase"
  | "crewTrainingPlan"
  | "roleManagement"
  | "userManagement"
  | "staff"
  | "levels"
  | "levelShiftPriority"
  | "rosterPriorityType"
  | "opsControl"
  | "mxTypeManager"
  | "aircraftReport"
  | "leaveAttendanceControl"
  | "publicHolidays"
  | "leaves"
  | "roster"
  | "rosterApprovals"
  | "shift"
  | "attendance"
  | "evaluation"
  | "checklist"
  | "briefing"

type EntryValue = string
type Entry = {
  id: string
  createdAt: string
  [key: string]: EntryValue
}
type DataStore = Record<ModuleKey, Entry[]>
type RosterComputeLog = {
  id: string
  at: string
  date: string
  level: string
  staff: string
  code: string
  severity: "info" | "warn" | "error"
  message: string
}
type RosterChangeLog = {
  id: string
  at: string
  requestedBy: string
  date: string
  staffNo: string
  staffName: string
  type: "Manual" | "Upload" | "Roster Manual Change"
  action:
    | "Added"
    | "Updated"
    | "Removed"
    | "Change Submitted"
    | "Approved Change"
    | "Rejected Change"
  details: string
  reason: string
}
type RosterChangeRequest = {
  id: string
  createdAt: string
  status: "Pending Approval" | "Approved" | "Rejected"
  changeType: "Mutual Swap" | "Shift Reassign" | "State Override"
  date: string
  staffNo: string
  staffName: string
  level: string
  fromCode: string
  toCode: string
  targetStaffNo: string
  targetStaffName: string
  conflictStaffNo: string
  conflictStaffName: string
  conflictFromCode: string
  conflictToCode: string
  reason: string
}
type StaffPortalRequest = {
  id: string
  createdAt?: string
  status:
    | "Pending Peer Acceptance"
    | "Pending Approval"
    | "Pending Document"
    | "Document Submitted"
    | "Approved"
    | "Cancelled"
    | "Rejected"
  type: "Leave" | "Attendance" | "Roster Change"
  staffNo?: string
  staffName?: string
  date?: string
  fromCode?: string
  toCode?: string
  changeWithStaffNo?: string
  changeWithStaffName?: string
  reason?: string
  dutyMarkType?: string
  leavePolicyName?: string
  fromDate?: string
  toDate?: string
  noOfDays?: string
  documentName?: string
  documentData?: string
  documentUploadedAt?: string
}

type FieldType =
  | "text"
  | "date"
  | "time"
  | "number"
  | "select"
  | "checkbox"
  | "staffLookup"
  | "multiselect"
  | "image"
  | "textarea"
type FieldConfig = {
  key: string
  label: string
  type: FieldType
  required?: boolean
  options?: string[]
}
type ModuleConfig = {
  key: ModuleKey
  title: string
  description: string
  fields: FieldConfig[]
  columns: string[]
}

const STORAGE_KEY = "occfloat.v1"
const AUTH_CACHE_KEY = "occfloat.authCache.v1"
const AUTH_STORAGE_KEY = "occfloat.mainAuthStaffId"
const STAFF_PORTAL_REQUESTS_KEY = "occfloat.staffPortalRequests"
const SUPABASE_STORE_TABLE = "occfloat_store"
const SUPABASE_STORE_ID = "primary"
const SUPABASE_ROSTER_AUDIT_TABLE = "occfloat_roster_audit_logs"
const SUPABASE_URL = process.env.NEXT_PUBLIC_SUPABASE_URL || ""
const SUPABASE_PUBLISHABLE_KEY = process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY || ""
const REMOTE_SAVE_MIN_INTERVAL_MS = 60_000
const REMOTE_SAVE_TIMEOUT_MS = 15_000
const SHEETJS_CDN = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"
const STEP_COLORS = {
  active: "#0d6efd",
  completed: "#198754",
  pending: "#adb5bd",
}

type SheetJSWindow = Window & {
  XLSX?: {
    read: (data: ArrayBuffer, opts: { type: "array"; cellDates?: boolean }) => {
      SheetNames: string[]
      Sheets: Record<string, unknown>
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
    }
  }
}

let sheetJsLoaderPromise: Promise<void> | null = null
function loadSheetJs(): Promise<void> {
  if (typeof window === "undefined") return Promise.reject(new Error("Excel import is available in browser only."))
  const w = window as SheetJSWindow
  if (w.XLSX) return Promise.resolve()
  if (sheetJsLoaderPromise) return sheetJsLoaderPromise
  sheetJsLoaderPromise = import("xlsx")
    .then((xlsxModule) => {
      w.XLSX = xlsxModule as unknown as SheetJSWindow["XLSX"]
    })
    .catch(
      () =>
        new Promise<void>((resolve, reject) => {
          const existing = document.querySelector(`script[src="${SHEETJS_CDN}"]`) as HTMLScriptElement | null
          if (existing) {
            existing.addEventListener("load", () => resolve())
            existing.addEventListener("error", () => reject(new Error("Failed to load Excel parser.")))
            return
          }
          const script = document.createElement("script")
          script.src = SHEETJS_CDN
          script.async = true
          script.onload = () => resolve()
          script.onerror = () => reject(new Error("Failed to load Excel parser."))
          document.head.appendChild(script)
        }),
    )
  return sheetJsLoaderPromise
}

function getAttendanceStepIndexFromForm(form: Record<string, string>): number {
  const status = (form.status || "").trim().toLowerCase()
  const workflow = (form.workflowStatus || "").trim().toLowerCase()
  const approval = (form.approvalStatus || "").trim().toLowerCase()
  const documentStatus = (form.documentStatus || "").trim().toLowerCase()
  if (status === "rejected" || workflow === "rejected") return 3
  if (status === "approved" || workflow === "approved" || approval === "approved") return 3
  if (
    workflow.includes("approval") ||
    approval.includes("pending") ||
    status.includes("pending approval")
  ) {
    return 2
  }
  if (
    workflow.includes("document") ||
    documentStatus.includes("request") ||
    documentStatus.includes("submitted") ||
    status.includes("document")
  ) {
    return 1
  }
  return 0
}

const workflowStepperSx = {
  "& .MuiStepLabel-label": { fontSize: 13, color: "#6c757d" },
  "& .MuiStepLabel-label.Mui-active": { color: "#0d6efd", fontWeight: 600 },
  "& .MuiStepLabel-label.Mui-completed": { color: "#198754", fontWeight: 600 },
  "& .MuiStepIcon-root": { color: STEP_COLORS.pending },
  "& .MuiStepIcon-root.Mui-active": { color: STEP_COLORS.active },
  "& .MuiStepIcon-root.Mui-completed": { color: STEP_COLORS.completed },
}

function attendanceStatusStyle(value: string): React.CSSProperties {
  const v = (value || "").trim().toLowerCase()
  if (v.includes("approved")) return { background: "#d1e7dd", color: "#0f5132" }
  if (v.includes("rejected")) return { background: "#f8d7da", color: "#842029" }
  if (v.includes("document") || v.includes("request")) {
    return { background: "#fff3cd", color: "#664d03" }
  }
  if (v.includes("approval") || v.includes("pending")) {
    return { background: "#cff4fc", color: "#055160" }
  }
  return { background: "#e2e3e5", color: "#41464b" }
}
const ACCESS_CONTROL_FLOWS = [
  "Dashboard > View",
  "Crew Operation > Crew Rostering > View",
  "Crew Operation > Crew Rostering > Edit",
  "Crew Operation > Crew Leave Planner > View",
  "Crew Operation > Crew Leave Planner > Edit",
  "Crew Operation > Crew Data Base > View",
  "Crew Operation > Crew Data Base > Edit",
  "Crew Operation > Crew Training Plan > View",
  "Crew Operation > Crew Training Plan > Edit",

  "Staff > Create Staff > View",
  "Staff > Create Staff > Edit",
  "Staff > Staff Level > View",
  "Staff > Staff Level > Edit",
  "Staff > Leave Type > View",
  "Staff > Leave Type > Edit",
  "Staff > Public Holidays > View",
  "Staff > Public Holidays > Edit",
  "Staff > Staff Leave > View",
  "Staff > Staff Leave > Edit",
  "Staff > Attendance > View",
  "Staff > Attendance > Edit",
  "Staff > Evaluation > View",
  "Staff > Evaluation > Edit",

  "Roster > Roster > View",
  "Roster > Roster > Generate",
  "Roster > Roster > Download",
  "Roster > Roster > Upload",
  "Roster > Roster > Edit",
  "Roster > Roster Approvals > View",
  "Roster > Roster Approvals > Edit",
  "Roster > Shift Setup > View",
  "Roster > Shift Setup > Edit",
  "Roster > Level Shift Priority > View",
  "Roster > Level Shift Priority > Edit",
  "Roster > Roster Priority Type > View",
  "Roster > Roster Priority Type > Edit",

  "Daily Ops > Ops Control > View",
  "Daily Ops > Ops Control > Edit",
  "Daily Ops > MX Type Manager > View",
  "Daily Ops > MX Type Manager > Edit",
  "Daily Ops > Aircraft Report > View",
  "Daily Ops > Aircraft Report > Edit",
  "Daily Ops > Daily Checklist > View",
  "Daily Ops > Daily Checklist > Edit",
  "Daily Ops > Briefing > View",
  "Daily Ops > Briefing > Edit",

  "Admin > Role Management > View",
  "Admin > Role Management > Edit",
  "Admin > User Management > View",
  "Admin > User Management > Edit",
]

const MODULE_VIEW_FLOW: Record<ModuleKey, string> = {
  dashboard: "Dashboard > View",
  crewRostering: "Crew Operation > Crew Rostering > View",
  crewLeavePlanner: "Crew Operation > Crew Leave Planner > View",
  crewDataBase: "Crew Operation > Crew Data Base > View",
  crewTrainingPlan: "Crew Operation > Crew Training Plan > View",
  staff: "Staff > Create Staff > View",
  levels: "Staff > Staff Level > View",
  leaveAttendanceControl: "Staff > Leave Type > View",
  publicHolidays: "Staff > Public Holidays > View",
  leaves: "Staff > Staff Leave > View",
  attendance: "Staff > Attendance > View",
  evaluation: "Staff > Evaluation > View",
  roster: "Roster > Roster > View",
  rosterApprovals: "Roster > Roster Approvals > View",
  shift: "Roster > Shift Setup > View",
  levelShiftPriority: "Roster > Level Shift Priority > View",
  rosterPriorityType: "Roster > Roster Priority Type > View",
  opsControl: "Daily Ops > Ops Control > View",
  mxTypeManager: "Daily Ops > MX Type Manager > View",
  aircraftReport: "Daily Ops > Aircraft Report > View",
  checklist: "Daily Ops > Daily Checklist > View",
  briefing: "Daily Ops > Briefing > View",
  roleManagement: "Admin > Role Management > View",
  userManagement: "Admin > User Management > View",
}

const MODULES: ModuleConfig[] = [
  {
    key: "dashboard",
    title: "Dashboard",
    description: "Overview of OCC operations and quick access to modules.",
    fields: [],
    columns: [],
  },
  {
    key: "crewRostering",
    title: "Crew Rostering",
    description: "Blank module workspace for Crew Operation > Crew Rostering.",
    fields: [],
    columns: [],
  },
  {
    key: "crewLeavePlanner",
    title: "Crew Leave Planner",
    description: "Blank module workspace for Crew Operation > Crew Leave Planner.",
    fields: [],
    columns: [],
  },
  {
    key: "crewDataBase",
    title: "Crew Data Base",
    description: "Manage crew profiles, status, and bulk updates.",
    fields: [
      { key: "crewCode", label: "Crew Code", type: "text", required: true },
      { key: "crewName", label: "Crew Name", type: "text", required: true },
      { key: "crewType", label: "Crew Type", type: "text" },
      { key: "crewCategory", label: "Crew Category", type: "text" },
      { key: "joiningDate", label: "Joining Date", type: "date" },
      { key: "releaseDate", label: "Release Date", type: "date" },
      { key: "emailId", label: "Email ID", type: "text" },
      { key: "mobileNo", label: "Mobile No", type: "text" },
      { key: "employeeNo", label: "Employee No", type: "text" },
      {
        key: "gender",
        label: "Gender",
        type: "select",
        options: ["", "M", "F"],
      },
      {
        key: "activeStatus",
        label: "Active Status",
        type: "select",
        options: ["Active", "Inactive"],
        required: true,
      },
      { key: "inactiveDate", label: "Inactive Date", type: "date" },
      { key: "inactiveReason", label: "Inactive Reason", type: "textarea" },
    ],
    columns: [
      "crewCode",
      "crewName",
      "crewType",
      "crewCategory",
      "joiningDate",
      "releaseDate",
      "emailId",
      "mobileNo",
      "employeeNo",
      "gender",
      "activeStatus",
      "inactiveDate",
      "inactiveReason",
    ],
  },
  {
    key: "crewTrainingPlan",
    title: "Crew Training Plan",
    description: "Blank module workspace for Crew Operation > Crew Training Plan.",
    fields: [],
    columns: [],
  },
  {
    key: "roleManagement",
    title: "Role Management",
    description: "Create roles and map control flow access for each role.",
    fields: [
      { key: "roleName", label: "Role Name", type: "text", required: true },
      {
        key: "controlFlows",
        label: "Control Flows",
        type: "multiselect",
        options: ACCESS_CONTROL_FLOWS,
        required: true,
      },
      { key: "notes", label: "Notes", type: "textarea" },
    ],
    columns: ["roleName", "controlFlows", "notes"],
  },
  {
    key: "userManagement",
    title: "User Management",
    description: "Manage Dispatch Staff Users and General Users.",
    fields: [
      {
        key: "userType",
        label: "User Type",
        type: "select",
        options: ["Dispatch Staff User", "General User"],
        required: true,
      },
      {
        key: "dispatchStaff",
        label: "Dispatch Staff",
        type: "select",
        options: [],
      },
      { key: "fullName", label: "Full Name", type: "text" },
      { key: "userName", label: "User Name", type: "text", required: true },
      { key: "email", label: "Email", type: "text" },
      { key: "password", label: "Password", type: "text" },
      { key: "roleName", label: "Role", type: "select", options: [], required: true },
      {
        key: "status",
        label: "Status",
        type: "select",
        options: ["Active", "Inactive"],
        required: true,
      },
    ],
    columns: ["userType", "dispatchStaff", "fullName", "userName", "email", "roleName", "status"],
  },
  {
    key: "staff",
    title: "Create Staff",
    description: "Create and manage OCC staff profiles.",
    fields: [
      { key: "avatar", label: "Avatar", type: "image" },
      { key: "staffNo", label: "Staff No", type: "text", required: true },
      { key: "fullName", label: "Full Name", type: "text", required: true },
      { key: "phoneNumber", label: "Phone Number", type: "text", required: true },
      { key: "empDate", label: "EMP Date", type: "date", required: true },
      {
        key: "employmentType",
        label: "Employment Type",
        type: "select",
        options: ["Permanent", "Contract", "Temporary", "Intern"],
        required: true,
      },
      {
        key: "addressPresent",
        label: "Address Present",
        type: "textarea",
        required: true,
      },
      {
        key: "addressPermanent",
        label: "Address Permanent",
        type: "textarea",
        required: true,
      },
      { key: "designation", label: "Designation", type: "text", required: true },
      {
        key: "level",
        label: "Level",
        type: "select",
        options: [],
      },
      { key: "newLevel", label: "Create New Level", type: "text" },
      {
        key: "activeStatus",
        label: "Active Status",
        type: "select",
        options: ["Active", "Inactive"],
        required: true,
      },
      { key: "inactiveDate", label: "Inactive Date", type: "date" },
      { key: "inactiveReason", label: "Inactive Reason", type: "textarea" },
    ],
    columns: [
      "avatar",
      "staffNo",
      "fullName",
      "phoneNumber",
      "empDate",
      "employmentType",
      "addressPresent",
      "addressPermanent",
      "designation",
      "level",
      "activeStatus",
      "inactiveDate",
      "inactiveReason",
    ],
  },
  {
    key: "levels",
    title: "Create Level",
    description: "Create and manage staff levels used across modules.",
    fields: [
      { key: "levelName", label: "Level Name", type: "text", required: true },
      {
        key: "levelRank",
        label: "Rank (Bar)",
        type: "select",
        options: ["4", "3", "2", "1", "0"],
        required: true,
      },
      { key: "levelDescription", label: "Level Description", type: "textarea" },
    ],
    columns: ["levelName", "levelRank", "levelDescription"],
  },
  {
    key: "levelShiftPriority",
    title: "Level Shift Priority",
    description: "Map level-wise priority order to roster-priority number groups.",
    fields: [
      {
        key: "levelName",
        label: "Level",
        type: "select",
        options: [],
        required: true,
      },
      {
        key: "priorityOrder",
        label: "Priority Order (Number)",
        type: "number",
        required: true,
      },
      {
        key: "rosterPriorityNumbers",
        label: "Roster Priority Number(s)",
        type: "multiselect",
        options: [],
        required: true,
      },
      { key: "notes", label: "Notes", type: "textarea" },
    ],
    columns: ["levelName", "priorityOrder", "rosterPriorityNumbers", "notes"],
  },
  {
    key: "leaveAttendanceControl",
    title: "Leave Type",
    description: "Define leave and attendance policy rules for employment types.",
    fields: [
      {
        key: "leaveAttendanceName",
        label: "Leave / Attendance Name",
        type: "text",
        required: true,
      },
      {
        key: "leaveAttendanceType",
        label: "Leave / Attendance Type",
        type: "select",
        options: ["Leave", "Attendance"],
        required: true,
      },
      { key: "typeCode", label: "Type Code", type: "text", required: true },
      { key: "noOfDays", label: "No of Days", type: "number", required: true },
      {
        key: "resetAt",
        label: "Reset At",
        type: "select",
        options: [
          "Completion of First Employment Year, then Every Employment Year (Annual Reset)",
        ],
        required: true,
      },
      {
        key: "eligibleEmploymentTypes",
        label: "Eligible Employment Type",
        type: "multiselect",
        options: ["Permanent", "Contract", "Intern"],
        required: true,
      },
      {
        key: "countPublicHoliday",
        label: "Count Public Holiday in Leave Duration",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "nonCountWeekdays",
        label: "Weekdays Not Counted as Leave",
        type: "multiselect",
        options: [
          "Monday",
          "Tuesday",
          "Wednesday",
          "Thursday",
          "Friday",
          "Saturday",
          "Sunday",
          "Gov Public Holidays",
        ],
      },
      {
        key: "documentRequired",
        label: "Document Required",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "needApproval",
        label: "Need Approval",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "allowNoOfDaysWithoutDocument",
        label: "Allow No of Days Without Document",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "consecutiveDaysWithoutDocument",
        label: "Consecutive Days (Without Document)",
        type: "number",
      },
    ],
    columns: [
      "leaveAttendanceName",
      "leaveAttendanceType",
      "typeCode",
      "noOfDays",
      "resetAt",
      "eligibleEmploymentTypes",
      "countPublicHoliday",
      "nonCountWeekdays",
      "documentRequired",
      "needApproval",
      "allowNoOfDaysWithoutDocument",
      "consecutiveDaysWithoutDocument",
    ],
  },
  {
    key: "publicHolidays",
    title: "Government Public Holidays",
    description: "Manage government public holiday date ranges used in leave calculations.",
    fields: [
      { key: "holidayStartDate", label: "Gov Holiday Start Date", type: "date", required: true },
      { key: "holidayEndDate", label: "Gov Holiday End Date", type: "date", required: true },
      { key: "holidayName", label: "Gov Holiday Name", type: "text", required: true },
    ],
    columns: ["holidayStartDate", "holidayEndDate", "holidayName"],
  },
  {
    key: "leaves",
    title: "Staff Leaves",
    description: "Track leave requests and approval status.",
    fields: [
      { key: "staffLookup", label: "Lookup Staff", type: "staffLookup", required: true },
      {
        key: "leavePolicyName",
        label: "Leave Policy",
        type: "select",
        options: [],
        required: true,
      },
      { key: "fromDate", label: "From Date", type: "date", required: true },
      { key: "toDate", label: "To Date", type: "date", required: true },
      {
        key: "status",
        label: "Status",
        type: "select",
        options: ["Pending", "Approved", "Rejected"],
        required: true,
      },
      { key: "remarks", label: "Remarks", type: "textarea" },
    ],
    columns: [
      "staffNo",
      "staffName",
      "leavePolicyName",
      "fromDate",
      "toDate",
      "totalCalendarDays",
      "excludedWeekdayDetails",
      "publicHolidayDays",
      "chargeableLeaveDays",
      "status",
      "remarks",
    ],
  },
  {
    key: "roster",
    title: "Roster",
    description: "Manage daily roster assignment and responsibilities.",
    fields: [
      { key: "date", label: "Date", type: "date", required: true },
      { key: "staffName", label: "Staff Name", type: "text", required: true },
      { key: "role", label: "Role", type: "text", required: true },
      {
        key: "shiftCode",
        label: "Shift",
        type: "select",
        options: ["Morning", "Evening", "Night"],
        required: true,
      },
      { key: "location", label: "Location", type: "text" },
    ],
    columns: ["date", "staffName", "role", "shiftCode", "location"],
  },
  {
    key: "rosterApprovals",
    title: "Roster Approvals",
    description: "Review and action roster change approval requests.",
    fields: [],
    columns: [],
  },
  {
    key: "shift",
    title: "Shift",
    description: "Define shift windows and assign responsible levels.",
    fields: [
      { key: "shiftIcon", label: "Shift Icon", type: "image" },
      { key: "shiftName", label: "Shift Name", type: "text", required: true },
      { key: "shiftCode", label: "Shift Code", type: "text", required: true },
      {
        key: "rosterPriority",
        label: "Roster Priority Order",
        type: "number",
        required: true,
      },
      { key: "startTime", label: "Start Time", type: "time", required: true },
      { key: "endTime", label: "End Time", type: "time", required: true },
      {
        key: "shiftType",
        label: "Shift Type",
        type: "select",
        options: ["Early", "Mid", "Late"],
        required: false,
      },
      {
        key: "assignedLevels",
        label: "Assigned Level",
        type: "multiselect",
        options: [],
        required: true,
      },
      {
        key: "showInShiftPairing",
        label: "Preset for Shift Pairing",
        type: "checkbox",
      },
      { key: "notes", label: "Notes", type: "textarea" },
    ],
    columns: [
      "shiftIcon",
      "shiftName",
      "shiftCode",
      "rosterPriority",
      "startTime",
      "endTime",
      "shiftDuration",
      "shiftType",
      "assignedLevels",
      "showInShiftPairing",
      "notes",
    ],
  },
  {
    key: "rosterPriorityType",
    title: "Roster Priority Type",
    description:
      "Map shift codes into Mandatory/Required/Optional/Admin and set daily/weekday requirement.",
    fields: [
      {
        key: "priorityType",
        label: "Roster Priority Type",
        type: "select",
        options: ["Mandatory", "Required", "Optional", "Admin"],
        required: true,
      },
      {
        key: "shiftCodes",
        label: "Shift Selection (Multiple)",
        type: "multiselect",
        options: [],
        required: true,
      },
      {
        key: "requiredDaily",
        label: "Required Daily",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "requiredWeekdays",
        label: "Week Days Required",
        type: "multiselect",
        options: ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
      },
      { key: "notes", label: "Notes", type: "textarea" },
    ],
    columns: ["priorityType", "shiftCodes", "requiredDaily", "requiredWeekdays", "notes"],
  },
  {
    key: "opsControl",
    title: "Ops Control",
    description: "Daily Ops control workspace.",
    fields: [],
    columns: [],
  },
  {
    key: "aircraftReport",
    title: "Aircraft Report",
    description: "Aircraft utilization reporting dashboard with filters and analytics.",
    fields: [],
    columns: [],
  },
  {
    key: "mxTypeManager",
    title: "MX Type Manager",
    description: "Manage maintenance types for Ops Control MX workflow.",
    fields: [],
    columns: [],
  },
  {
    key: "attendance",
    title: "Attendance",
    description: "Review attendance requests from staff portal and process approvals/documents.",
    fields: [
      { key: "date", label: "Date", type: "date", required: true },
      { key: "fromDate", label: "Requested From", type: "date" },
      { key: "toDate", label: "Requested To", type: "date" },
      { key: "noOfDays", label: "No of Days", type: "number" },
      { key: "attendanceType", label: "Attendance Type", type: "text" },
      { key: "staffNo", label: "Staff No", type: "text" },
      { key: "staffName", label: "Staff Name", type: "text", required: true },
      {
        key: "status",
        label: "Status",
        type: "select",
        options: [
          "Pending",
          "Pending Document Upload",
          "Pending Approval",
          "Approved",
          "Rejected",
          "Present",
          "Late",
          "Absent",
          "On Leave",
        ],
        required: true,
      },
      { key: "workflowStatus", label: "Workflow", type: "text" },
      { key: "approvalStatus", label: "Approval", type: "text" },
      { key: "documentStatus", label: "Document", type: "text" },
      { key: "requestSource", label: "Source", type: "text" },
      { key: "checkIn", label: "Check In", type: "time" },
      { key: "remarks", label: "Remarks", type: "textarea" },
    ],
    columns: [
      "date",
      "staffNo",
      "staffName",
      "attendanceType",
      "fromDate",
      "toDate",
      "noOfDays",
      "status",
      "documentName",
      "requestSource",
      "remarks",
    ],
  },
  {
    key: "evaluation",
    title: "Evaluation",
    description: "Record staff performance and evaluation scores.",
    fields: [
      { key: "date", label: "Date", type: "date", required: true },
      { key: "staffName", label: "Staff Name", type: "text", required: true },
      { key: "category", label: "Category", type: "text", required: true },
      { key: "score", label: "Score (1-10)", type: "number", required: true },
      { key: "feedback", label: "Feedback", type: "textarea" },
    ],
    columns: ["date", "staffName", "category", "score", "feedback"],
  },
  {
    key: "checklist",
    title: "Daily OCC Checklist",
    description: "Manage daily OCC actions and completion status.",
    fields: [
      { key: "date", label: "Date", type: "date", required: true },
      { key: "task", label: "Checklist Task", type: "text", required: true },
      {
        key: "status",
        label: "Status",
        type: "select",
        options: ["Pending", "Completed"],
        required: true,
      },
      { key: "owner", label: "Owner", type: "text", required: true },
      { key: "remarks", label: "Remarks", type: "textarea" },
    ],
    columns: ["date", "task", "status", "owner", "remarks"],
  },
  {
    key: "briefing",
    title: "Briefing",
    description: "Capture operational briefing points and action owners.",
    fields: [
      { key: "date", label: "Date", type: "date", required: true },
      { key: "topic", label: "Topic", type: "text", required: true },
      { key: "details", label: "Details", type: "textarea", required: true },
      { key: "actionOwner", label: "Action Owner", type: "text" },
      { key: "deadline", label: "Deadline", type: "date" },
    ],
    columns: ["date", "topic", "details", "actionOwner", "deadline"],
  },
]

type NavItem = {
  key: ModuleKey
  label: string
  icon: typeof Users
}

type NavGroup = {
  title: string
  icon: typeof Users
  items: NavItem[]
}

const NAV_GROUPS: NavGroup[] = [
  {
    title: "Crew Operation",
    icon: Plane,
    items: [
      { key: "crewRostering", label: "Crew Rostering", icon: CalendarDays },
      { key: "crewLeavePlanner", label: "Crew Leave Planner", icon: ClipboardList },
      { key: "crewDataBase", label: "Crew Data Base", icon: UserPlus },
      { key: "crewTrainingPlan", label: "Crew Training Plan", icon: Award },
    ],
  },
  {
    title: "Staff",
    icon: Users,
    items: [
      { key: "staff", label: "Create Staff", icon: UserPlus },
      { key: "levels", label: "Staff Level", icon: GraduationCap },
      { key: "leaveAttendanceControl", label: "Leave Type", icon: UserCog },
      { key: "publicHolidays", label: "Public Holidays", icon: Flag },
      { key: "leaves", label: "Staff Leave", icon: ClipboardList },
      { key: "attendance", label: "Attendance", icon: Activity },
      { key: "evaluation", label: "Evaluation", icon: Award },
    ],
  },
  {
    title: "Roster",
    icon: CalendarDays,
    items: [
      { key: "roster", label: "Roster", icon: CalendarDays },
      { key: "rosterApprovals", label: "Roster Approvals", icon: ClipboardList },
      { key: "shift", label: "Shift", icon: Clock },
      { key: "levelShiftPriority", label: "Level Shift Priority", icon: Layers },
      { key: "rosterPriorityType", label: "Roster Priority", icon: TrendingUp },
    ],
  },
  {
    title: "Daily Ops",
    icon: Activity,
    items: [
      { key: "opsControl", label: "Ops Control", icon: Settings },
      { key: "mxTypeManager", label: "MX Type Manager", icon: Settings },
      { key: "aircraftReport", label: "Aircraft Report", icon: TrendingUp },
      { key: "checklist", label: "Daily Checklist", icon: ListChecks },
      { key: "briefing", label: "Briefing", icon: Megaphone },
    ],
  },
  {
    title: "Admin",
    icon: Shield,
    items: [
      { key: "roleManagement", label: "Role Management", icon: Shield },
      { key: "userManagement", label: "User Management", icon: UserCog },
    ],
  },
]

function getEmptyStore(): DataStore {
  return {
    dashboard: [],
    crewRostering: [],
    crewLeavePlanner: [],
    crewDataBase: [],
    crewTrainingPlan: [],
    roleManagement: [],
    userManagement: [],
    staff: [],
    levels: [],
    levelShiftPriority: [],
    rosterPriorityType: [],
    opsControl: [],
    mxTypeManager: [],
    aircraftReport: [],
    leaveAttendanceControl: [],
    publicHolidays: [],
    leaves: [],
    roster: [],
    rosterApprovals: [],
    shift: [],
    attendance: [],
    evaluation: [],
    checklist: [],
    briefing: [],
  }
}

function compactStoreForLocal(storage: DataStore): DataStore {
  const MAX_LOCAL_OPS_ROWS = 800
  const compactOps = (storage.opsControl || [])
    .slice(-MAX_LOCAL_OPS_ROWS)
    .map((row) => ({
      ...row,
      remarks: String(row.remarks || "").slice(0, 120),
    }))
  return {
    ...storage,
    opsControl: compactOps,
  }
}

function buildEmptyForm(module: ModuleConfig): Record<string, string> {
  const next: Record<string, string> = {}
  module.fields.forEach((field) => {
    if (field.type === "checkbox") {
      next[field.key] = "No"
      return
    }
    if (field.type === "select" && field.options?.length) {
      next[field.key] = field.options[0]
      return
    }
    if (field.type === "multiselect") {
      next[field.key] = ""
      return
    }
    next[field.key] = ""
  })
  return next
}

function labelize(key: string): string {
  if (key === "chargeableLeaveDays") return "Leave Days"
  if (key === "publicHolidayDays") return "Gov Public Holiday Days"
  return key.replace(/([A-Z])/g, " $1").replace(/^./, (s) => s.toUpperCase())
}

function splitMultiValue(value: string): string[] {
  return Array.from(
    new Set(
      value
        .split(",")
        .map((item) => item.trim())
        .filter(Boolean),
    ),
  )
}

function normalizeText(value: string): string {
  return (value || "").trim().toLowerCase()
}

function buildAttendanceDedupKey(input: {
  staffNo?: string
  attendanceType?: string
  fromDate?: string
  toDate?: string
  requestSource?: string
}): string {
  const staffNo = (input.staffNo || "").trim().toLowerCase()
  const attendanceType = normalizeText(input.attendanceType || "")
  const fromDate = (input.fromDate || "").trim()
  const toDate = (input.toDate || "").trim() || fromDate
  const source = normalizeText(input.requestSource || "staff portal")
  return `${staffNo}::${attendanceType}::${fromDate}::${toDate}::${source}`
}

function isOffCode(value: string): boolean {
  const code = (value || "").trim().toUpperCase()
  return code === "OF" || code === "OFF" || code === "REST"
}

function normalizePriorityTypeName(value: string): string {
  const normalized = normalizeText(value)
  if (!normalized) return ""
  if (normalized.includes("mandatory")) return "mandatory"
  if (normalized.includes("required")) return "required"
  if (normalized.includes("optional")) return "optional"
  if (normalized.includes("admin")) return "admin"
  if (normalized.includes("rest") || normalized.includes("off")) return "rest/off"
  return normalized
}

function levelIncludedInAssigned(assignedLevels: string, level: string): boolean {
  const normalizedLevel = normalizeText(level)
  if (!normalizedLevel) return false
  return splitMultiValue(assignedLevels).some((item) => normalizeText(item) === normalizedLevel)
}

function parseLeadingNumber(value: string): number | null {
  const match = (value || "").trim().match(/^\d+/)
  if (!match) return null
  const n = Number(match[0])
  return Number.isFinite(n) ? n : null
}

function isYesLike(value: unknown): boolean {
  if (value === true) return true
  if (typeof value === "string") {
    const v = value.trim().toLowerCase()
    return v === "yes" || v === "true" || v === "1" || v === "on"
  }
  if (typeof value === "number") return value === 1
  return false
}

function parseCsvLine(line: string): string[] {
  const result: string[] = []
  let current = ""
  let inQuotes = false
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"'
        i += 1
      } else {
        inQuotes = !inQuotes
      }
      continue
    }
    if (ch === "," && !inQuotes) {
      result.push(current.trim())
      current = ""
      continue
    }
    current += ch
  }
  result.push(current.trim())
  return result
}

function toYmd(date: Date): string {
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, "0")
  const d = String(date.getDate()).padStart(2, "0")
  return `${y}-${m}-${d}`
}

function formatYmdToDdMmYy(ymd: string): string {
  const match = ymd.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (!match) return ymd
  const [, year, month, day] = match
  return `${day}/${month}/${year.slice(-2)}`
}

function formatYmdToWeekdayShort(ymd: string): string {
  const date = new Date(`${ymd}T00:00:00`)
  if (Number.isNaN(date.getTime())) return ""
  return date.toLocaleDateString("en-US", { weekday: "short" })
}

function formatDateText(value: string): string {
  if (!value) return value
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return formatYmdToDdMmYy(value)
  }
  const parsed = new Date(value)
  if (!Number.isNaN(parsed.getTime())) {
    const dd = String(parsed.getDate()).padStart(2, "0")
    const mm = String(parsed.getMonth() + 1).padStart(2, "0")
    const yy = String(parsed.getFullYear()).slice(-2)
    const hh = String(parsed.getHours()).padStart(2, "0")
    const min = String(parsed.getMinutes()).padStart(2, "0")
    return `${dd}/${mm}/${yy} ${hh}:${min}`
  }
  return value.replace(/\b(\d{4})-(\d{2})-(\d{2})\b/g, (_m, y, m, d) => {
    return `${d}/${m}/${String(y).slice(-2)}`
  })
}

function isDateLikeColumn(column: string): boolean {
  const key = column.toLowerCase()
  return (
    key.includes("date") ||
    key.includes("time") ||
    key.endsWith("at") ||
    key === "from" ||
    key === "to"
  )
}

function normalizeCrewCode(value: string): string {
  return (value || "").trim().toUpperCase()
}

function addDaysToYmd(ymd: string, n: number): string {
  const d = new Date(`${ymd}T00:00:00`)
  d.setDate(d.getDate() + n)
  return toYmd(d)
}

function diffDaysYmd(fromYmd: string, toYmd: string): number {
  if (!fromYmd || !toYmd) return 0
  const a = new Date(`${fromYmd}T00:00:00`)
  const b = new Date(`${toYmd}T00:00:00`)
  if (Number.isNaN(a.getTime()) || Number.isNaN(b.getTime())) return 0
  return Math.round((b.getTime() - a.getTime()) / (24 * 60 * 60 * 1000))
}

function overlapDaysInclusive(startA: string, endA: string, startB: string, endB: string): number {
  if (!startA || !endA || !startB || !endB) return 0
  const s = startA > startB ? startA : startB
  const e = endA < endB ? endA : endB
  if (e < s) return 0
  return getDateRangeInclusive(s, e).length
}

function parseLocalDateTime(ymd: string, hm: string): Date {
  return new Date(`${ymd}T${hm}:00`)
}

function addHoursToDate(date: Date, hours: number): Date {
  return new Date(date.getTime() + hours * 60 * 60 * 1000)
}

function formatDateToYmdHm(date: Date): { ymd: string; hm: string } {
  return {
    ymd: toYmd(date),
    hm: `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`,
  }
}

type RosterDayTimingSegment = {
  day: string
  type: string
  restFrom: string
  restTo: string
  rosterFrom: string
  rosterTo: string
}

function buildRosterDaySegments(startDateTime: Date, endDateTime: Date, segmentType: string) {
  const segments: RosterDayTimingSegment[] = []
  if (endDateTime <= startDateTime) return segments
  let cursor = new Date(startDateTime)
  while (cursor < endDateTime) {
    const day = toYmd(cursor)
    const nextDayStart = new Date(`${addDaysToYmd(day, 1)}T00:00:00`)
    const segmentEnd = endDateTime < nextDayStart ? endDateTime : nextDayStart
    const from = formatDateToYmdHm(cursor).hm
    const toMeta = formatDateToYmdHm(segmentEnd)
    const to = toMeta.ymd === day ? toMeta.hm : "23:59"
    const isRest = segmentType === "REST"
    segments.push({
      day,
      type: segmentType,
      restFrom: isRest ? from : "",
      restTo: isRest ? to : "",
      rosterFrom: isRest ? "" : from,
      rosterTo: isRest ? "" : to,
    })
    cursor = segmentEnd
  }
  return segments
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

function normalizeDateOnly(value: string): string {
  const raw = (value || "").trim()
  if (!raw) return ""
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw
  const isoDatePrefix = raw.match(/^(\d{4}-\d{2}-\d{2})T/)
  if (isoDatePrefix) return isoDatePrefix[1]
  const parsed = new Date(raw)
  if (!Number.isNaN(parsed.getTime())) return toYmd(parsed)
  return ""
}

function calculateShiftDuration(startTime: string, endTime: string): string {
  const parseMinutes = (value: string) => {
    const [h = "0", m = "0"] = value.split(":")
    const hours = Number(h)
    const mins = Number(m)
    if (Number.isNaN(hours) || Number.isNaN(mins)) return 0
    return hours * 60 + mins
  }

  const start = parseMinutes(startTime)
  const end = parseMinutes(endTime)
  let duration = end - start
  if (duration < 0) duration += 24 * 60

  const hours = Math.floor(duration / 60)
  const minutes = duration % 60
  return `${hours}h ${minutes}m`
}

function parseShiftDurationToHours(duration: string): number {
  const hMatch = duration.match(/(\d+)\s*h/)
  const mMatch = duration.match(/(\d+)\s*m/)
  const h = hMatch ? Number(hMatch[1]) : 0
  const m = mMatch ? Number(mMatch[1]) : 0
  return h + m / 60
}

function calculateShiftDurationHoursFromTimes(startTime: string, endTime: string): number {
  const parseMinutes = (value: string) => {
    const [h = "0", m = "0"] = value.split(":")
    const hours = Number(h)
    const mins = Number(m)
    if (Number.isNaN(hours) || Number.isNaN(mins)) return 0
    return hours * 60 + mins
  }
  const start = parseMinutes(startTime)
  const end = parseMinutes(endTime)
  let duration = end - start
  if (duration < 0) duration += 24 * 60
  return duration / 60
}

const DEFAULT_CREW_DUTY_TYPES = [
  { dutyType: "Flight Duty", rosterBarMap: "ROSTER DAY" },
  { dutyType: "Ground Duty", rosterBarMap: "ROSTER DAY" },
  { dutyType: "Release - OF", rosterBarMap: "REST" },
]

function getCrewDutyRosterBarMap(dutyType: string): "REST" | "ROSTER DAY" {
  const normalized = dutyType.trim().toLowerCase()
  const match = DEFAULT_CREW_DUTY_TYPES.find((row) => row.dutyType.toLowerCase() === normalized)
  if (match?.rosterBarMap === "REST") return "REST"
  if (normalized.includes("of") || normalized.includes("release") || normalized.includes("rest")) return "REST"
  return "ROSTER DAY"
}

function getDefaultDutyCodeColor(code: string): string {
  const palette: Record<string, string> = {
    FD: "#0ea5a5",
    GD: "#2563eb",
    OF: "#334155",
    DP: "#f97316",
  }
  const normalized = code.trim().toUpperCase()
  if (palette[normalized]) return palette[normalized]
  const colors = ["#0f766e", "#1d4ed8", "#7c3aed", "#be123c", "#b45309", "#047857", "#4338ca"]
  const seed = normalized.split("").reduce((acc, ch) => acc + ch.charCodeAt(0), 0)
  return colors[seed % colors.length]
}

function combineDateAndTime(ymd: string, hhmm: string): Date {
  return new Date(`${ymd}T${hhmm || "00:00"}:00`)
}

function getShiftEndDateTime(ymd: string, startTime: string, endTime: string): Date {
  const start = combineDateAndTime(ymd, startTime)
  const end = combineDateAndTime(ymd, endTime)
  if (end <= start) {
    end.setDate(end.getDate() + 1)
  }
  return end
}

function getCurrentMonthRange() {
  const now = new Date()
  const first = new Date(now.getFullYear(), now.getMonth(), 1)
  const last = new Date(now.getFullYear(), now.getMonth() + 1, 0)
  return { fromDate: toYmd(first), toDate: toYmd(last) }
}

function getCurrentAndNextMonthRange() {
  return getCurrentYearRange()
}

function getCurrentYearRange() {
  const now = new Date()
  const first = new Date(now.getFullYear(), 0, 1)
  const last = new Date(now.getFullYear(), 11, 31)
  return { fromDate: toYmd(first), toDate: toYmd(last) }
}

function getCurrentSixMonthRange() {
  const now = new Date()
  const first = new Date(now.getFullYear(), now.getMonth(), 1)
  const last = new Date(now.getFullYear(), now.getMonth() + 6, 0)
  return { fromDate: toYmd(first), toDate: toYmd(last) }
}

function getSixMonthRangeFromDate(ymd?: string) {
  const base = ymd ? new Date(`${ymd}T00:00:00`) : new Date()
  const safeBase = Number.isNaN(base.getTime()) ? new Date() : base
  const first = new Date(safeBase.getFullYear(), safeBase.getMonth(), 1)
  const last = new Date(safeBase.getFullYear(), safeBase.getMonth() + 6, 0)
  return { fromDate: toYmd(first), toDate: toYmd(last) }
}

function getDefaultRosterGenerationRange() {
  const now = new Date()
  const end = new Date(now)
  end.setMonth(end.getMonth() + 1)
  return { fromDate: toYmd(now), toDate: toYmd(end) }
}

function getStaffSelectionToken(staff: Entry): string {
  const id = (staff.id || "").trim()
  if (id) return id
  return `${(staff.staffNo || "").trim()}::${(staff.fullName || "").trim()}`
}

// Solid, saturated cell colors — readable but not washed out. `bg` is the
// cell background, `text` is the foreground (chosen for >= 4.5:1 contrast).
const SHIFT_COLOR_PALETTE = [
  { bg: "#60a5fa", text: "#0f172a" }, // blue
  { bg: "#4ade80", text: "#052e16" }, // green
  { bg: "#fbbf24", text: "#1f2937" }, // amber
  { bg: "#a78bfa", text: "#1e1b4b" }, // purple
  { bg: "#f87171", text: "#450a0a" }, // rose
  { bg: "#22d3ee", text: "#083344" }, // cyan
  { bg: "#818cf8", text: "#1e1b4b" }, // indigo
  { bg: "#fb923c", text: "#431407" }, // orange
  { bg: "#2dd4bf", text: "#042f2e" }, // teal
  { bg: "#f472b6", text: "#500724" }, // pink
]

// Special code colors. OF uses black background + white text (no border)
// matching the same code-cell shape as duty codes.
const ROSTER_OFF_COLOR = { bg: "#334155", text: "#f8fafc" } // lighter slate
const ROSTER_AL_COLOR = { bg: "#9ca3af", text: "#111827" } // light gray
const ROSTER_PH_COLOR = { bg: "#f472b6", text: "#500724" } // lighter pink
const ROSTER_GH_COLOR = { bg: "#4ade80", text: "#052e16" } // lighter green

function getShiftLucideIcon(entry: Entry) {
  const shiftType = normalizeText(entry.shiftType || "")
  const shiftName = normalizeText(entry.shiftName || "")
  const shiftCode = normalizeText(entry.shiftCode || "")
  if (shiftName.includes("flight watch")) return Headset
  if (shiftName.includes("flight release")) return Plane
  if (shiftName.includes("ops controller")) return User
  if (shiftName.includes("crew scheduler")) return CalendarCheck
  if (shiftName.includes("docs and load") || shiftName.includes("docs & load")) return Clipboard
  if (shiftType === "early" || shiftName.includes("early")) return Sunrise
  if (shiftType === "mid" || shiftName.includes("mid")) return Sun
  if (shiftType === "late" || shiftName.includes("late")) return Sunset
  if (
    shiftName.includes("night") ||
    shiftCode.startsWith("n") ||
    shiftCode.startsWith("d")
  ) {
    return Moon
  }
  if (shiftName.includes("admin") || shiftCode.startsWith("o")) return ShieldAlert
  return Mic
}

export default function Page() {
  const supabase = useMemo(() => {
    try {
      return createSupabaseClient()
    } catch {
      return null
    }
  }, [])
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const forceRemoteSyncRef = useRef(false)
  const skipNextRemoteSaveRef = useRef(false)
  const lastSyncedStoreJsonRef = useRef("")
  const lastRemoteSaveStartedAtRef = useRef(0)
  const rosterUploadInputRef = useRef<HTMLInputElement | null>(null)
  const crewRosterExcelInputRef = useRef<HTMLInputElement | null>(null)
  const rosterGridWrapRef = useRef<HTMLDivElement | null>(null)
  const hasLoadedRef = useRef(false)
  const disableRemoteSyncRef = useRef(false)
  const isOnlineRef = useRef(typeof window === "undefined" ? true : window.navigator.onLine)
  const [isOnline, setIsOnline] = useState(() =>
    typeof window === "undefined" ? true : window.navigator.onLine,
  )
  const [store, setStore] = useState<DataStore>(getEmptyStore)
  const [activeModule, setActiveModule] = useState<ModuleKey>("dashboard")
  const [isFormOpen, setIsFormOpen] = useState(false)
  const [staffLookupQuery, setStaffLookupQuery] = useState("")
  const [editingEntryId, setEditingEntryId] = useState<string | null>(null)
  const [crewDbSearch, setCrewDbSearch] = useState("")
  const [crewLeavePolicyForm, setCrewLeavePolicyForm] = useState({
    policyName: "",
    crewType: "Captain",
    workPatternDays: "90",
    leavePatternDays: "30",
    codeThresholdsJson: "{}",
    activeStatus: "Active",
  })
  const [crewLeavePolicyThresholds, setCrewLeavePolicyThresholds] = useState<Record<string, string>>({})
  const [crewLeaveAssignmentForm, setCrewLeaveAssignmentForm] = useState({
    crewCode: "",
    policyId: "",
  })
  const [crewLeaveMarkForm, setCrewLeaveMarkForm] = useState({
    crewCode: "",
    markType: "MC",
    cycleNumber: "",
    fromDate: toYmd(new Date()),
    toDate: toYmd(new Date()),
    reason: "",
  })
  const [editingCrewLeaveMarkId, setEditingCrewLeaveMarkId] = useState<string | null>(null)
  const [crewLeaveGeneratorForm, setCrewLeaveGeneratorForm] = useState({
    crewCode: "",
    years: "2",
    replaceExisting: "Yes",
  })
  const [crewLeavePlannerTab, setCrewLeavePlannerTab] = useState<
    | "rotationTypes"
    | "attendanceCodes"
    | "dutyTypes"
    | "dutyCodes"
    | "dailyAttendance"
    | "leavePatternManager"
    | "generator"
    | "blocks"
    | "timeline"
    | "conflicts"
  >("rotationTypes")
  const [crewLeavePlannerFocusMode, setCrewLeavePlannerFocusMode] = useState(false)
  const [leavePatternManagerSearch, setLeavePatternManagerSearch] = useState("")
  const [leavePatternManagerCrewCode, setLeavePatternManagerCrewCode] = useState("")
  const [leavePatternSegmentDraft, setLeavePatternSegmentDraft] = useState<Record<string, { start: string; end: string }>>({})
  const [leavePatternExpandedByCycle, setLeavePatternExpandedByCycle] = useState<Record<string, boolean>>({})
  const [leaveDetailModal, setLeaveDetailModal] = useState<{
    open: boolean
    blockId: string
    actualLeaveSegmentId?: string
    crewCode: string
    crewName: string
    cycleNumber: string
    leaveStart: string
    leaveEnd: string
    policyName: string
  }>({
    open: false,
    blockId: "",
    actualLeaveSegmentId: "",
    crewCode: "",
    crewName: "",
    cycleNumber: "",
    leaveStart: "",
    leaveEnd: "",
    policyName: "",
  })
  const [leaveDetailDraft, setLeaveDetailDraft] = useState({ start: "", end: "" })
  const [leaveDetailSelectedCycle, setLeaveDetailSelectedCycle] = useState("")
  const [editingLeaveSegmentId, setEditingLeaveSegmentId] = useState("")
  const [editingLeaveSegmentDraft, setEditingLeaveSegmentDraft] = useState({ cycleNumber: "", start: "", end: "" })
  const [isRotationModalOpen, setIsRotationModalOpen] = useState(false)
  const [editingRotationId, setEditingRotationId] = useState<string | null>(null)
  const [crewLeaveBlocksFilter, setCrewLeaveBlocksFilter] = useState({
    crewCode: "",
    fromDate: "",
    toDate: "",
  })
  const [crewLeaveTimelineFilter, setCrewLeaveTimelineFilter] = useState(() => getCurrentYearRange())
  const [crewLeaveTimelineCrewTypeFilter, setCrewLeaveTimelineCrewTypeFilter] = useState("Captain")
  const [showTimelineWpLpSummary, setShowTimelineWpLpSummary] = useState(true)
  const [showTimelineWorkBars, setShowTimelineWorkBars] = useState(true)
  const [showTimelineLeaveBars, setShowTimelineLeaveBars] = useState(true)
  const [isCrewRosterGeneratorOpen, setIsCrewRosterGeneratorOpen] = useState(false)
  const defaultCrewRosterGenerationRange = useMemo(() => getCurrentSixMonthRange(), [])
  const [crewRosterGeneratorForm, setCrewRosterGeneratorForm] = useState({
    fromDate: defaultCrewRosterGenerationRange.fromDate,
    toDate: defaultCrewRosterGenerationRange.toDate,
    crewCode: "",
    useFixedRestWeekday: false,
    fixedRestWeekdays: [] as string[],
    defaultDutyCode: "FD",
    dutyCodeSequence: "FD",
    maxDutyTimePerDay: "13",
    maxRecoveryRestGap: "110",
    restPeriod: "58",
    localNight: "22:00-06:00",
    localDay: "06:00-22:00",
  })
  const [crewRosterGeneratorTab, setCrewRosterGeneratorTab] = useState<"setup" | "rules">("setup")
  const [crewLeaveTimelineDayWidth, setCrewLeaveTimelineDayWidth] = useState(36)
  const crewLeaveTimelineScrollRef = useRef<HTMLDivElement | null>(null)
  const leaveBarDragRef = useRef<{
    blockId: string
    startX: number
    startY: number
    moved: boolean
  } | null>(null)
  const lastLeaveBarDragAtRef = useRef(0)
  const lastRosterBarSelectionAtRef = useRef(0)
  const isCrewLeavePlannerActive = activeModule === "crewLeavePlanner"
  const isCrewTimelineActive = isCrewLeavePlannerActive && crewLeavePlannerTab === "timeline"
  const isCrewBlocksActive = isCrewLeavePlannerActive && crewLeavePlannerTab === "blocks"
  const isCrewConflictsActive = isCrewLeavePlannerActive && crewLeavePlannerTab === "conflicts"
  const [draggingLeaveBlockId, setDraggingLeaveBlockId] = useState("")
  const [lastLeaveBarAdjustment, setLastLeaveBarAdjustment] = useState<{
    crewCode: string
    cycleNumber: string
    leaveDeltaDays: number
    nextWorkDeltaDays: number
  } | null>(null)
  const [pendingLeaveBarAdjustment, setPendingLeaveBarAdjustment] = useState<{
    blockId: string
    crewCode: string
    crewName: string
    cycleNumber: string
    policyName: string
    leaveStart: string
    leaveEnd: string
    deltaDays: number
  } | null>(null)
  const [rosterBarDetailModal, setRosterBarDetailModal] = useState<{
    open: boolean
    crewCode: string
    crewName: string
    rosterType: string
    startDate: string
    startTime: string
    endDate: string
    endTime: string
    durationHours: string
    daySegments: RosterDayTimingSegment[]
    dutyAssignments: Array<{
      id: string
      code: string
      dutyType: string
      startDate: string
      startTime: string
      endDate: string
      endTime: string
    }>
  }>({
    open: false,
    crewCode: "",
    crewName: "",
    rosterType: "",
    startDate: "",
    startTime: "",
    endDate: "",
    endTime: "",
    durationHours: "",
    daySegments: [],
    dutyAssignments: [],
  })
  const [pendingActualLeaveDrop, setPendingActualLeaveDrop] = useState<{
    open: boolean
    sourceBlockId: string
    crewCode: string
    crewName: string
    cycleNumber: string
    startDate: string
    endDate: string
    plannedDays: number
    conflicts: Array<{ type: string; code: string; start: string; end: string; startTime: string; endTime: string }>
    override: boolean
  }>({
    open: false,
    sourceBlockId: "",
    crewCode: "",
    crewName: "",
    cycleNumber: "",
    startDate: "",
    endDate: "",
    plannedDays: 0,
    conflicts: [],
    override: false,
  })
  const [pendingCrewDutyDrop, setPendingCrewDutyDrop] = useState<{
    open: boolean
    crewCode: string
    crewName: string
    code: string
    dutyType: string
    rosterBarMap: "REST" | "ROSTER DAY"
    startDate: string
    startTime: string
    endDate: string
    endTime: string
    conflicts: Array<{ id: string; type: string; code: string; start: string; end: string; startTime: string; endTime: string }>
    mode: "override" | "after"
  }>({
    open: false,
    crewCode: "",
    crewName: "",
    code: "",
    dutyType: "",
    rosterBarMap: "ROSTER DAY",
    startDate: "",
    startTime: "",
    endDate: "",
    endTime: "",
    conflicts: [],
    mode: "override",
  })
  const [crewOpsCodeForm, setCrewOpsCodeForm] = useState({
    code: "MC",
    name: "",
    allowLeaveDeduction: "Yes",
  })
  const [crewDutyCodeForm, setCrewDutyCodeForm] = useState({
    code: "",
    name: "",
    dutyType: "Flight Duty",
    startTime: "00:00",
    endTime: "00:00",
    color: "#0ea5a5",
  })
  const [editingCrewDutyCodeId, setEditingCrewDutyCodeId] = useState<string | null>(null)
  const [crewDutyTypeForm, setCrewDutyTypeForm] = useState({
    dutyType: "",
    rosterBarMap: "ROSTER DAY",
  })
  const [editingCrewDutyTypeId, setEditingCrewDutyTypeId] = useState<string | null>(null)
  const [editingCrewDutyTypeName, setEditingCrewDutyTypeName] = useState("")
  // Conflict-detection config: max number of crew (of a given crewType) allowed
  // to be on leave on the same day. "_default" applies when a specific crewType
  // entry is missing.
  const [crewLeaveConflictConfig, setCrewLeaveConflictConfig] = useState<{
    thresholdsByType: Record<string, number>
    defaultThreshold: number
    fromDate: string
    toDate: string
    crewTypeFilter: string
    severityFilter: "all" | "high" | "medium" | "low"
  }>({
    thresholdsByType: {},
    defaultThreshold: 2,
    fromDate: "",
    toDate: "",
    crewTypeFilter: "",
    severityFilter: "all",
  })
  const [isInactivePopupOpen, setIsInactivePopupOpen] = useState(false)
  const [inactivePopupStaffId, setInactivePopupStaffId] = useState<string | null>(null)
  const [inactivePopupFromEdit, setInactivePopupFromEdit] = useState(false)
  const [inactiveDateDraft, setInactiveDateDraft] = useState("")
  const [inactiveReasonDraft, setInactiveReasonDraft] = useState("")
  const [isRosterGeneratorOpen, setIsRosterGeneratorOpen] = useState(false)
  const [isRosterIOModalOpen, setIsRosterIOModalOpen] = useState(false)
  const [isSidebarOpen, setIsSidebarOpen] = useState(false)
  const [openGroupTitle, setOpenGroupTitle] = useState<string | null>(null)
  const [rosterView, setRosterView] = useState({
    staffSearch: "",
    shiftFilter: "ALL",
    sortBy: "staffNoAsc",
  })
  const [rosterDistributionExcluded, setRosterDistributionExcluded] = useState<Set<string>>(
    () => new Set(),
  )
  const [rosterComputeLogs, setRosterComputeLogs] = useState<RosterComputeLog[]>([])
  const [rosterComputeRangeDraft, setRosterComputeRangeDraft] = useState(() => getCurrentMonthRange())
  const [rosterComputeRange, setRosterComputeRange] = useState(() => getCurrentMonthRange())
  const [rosterChangeLogs, setRosterChangeLogs] = useState<RosterChangeLog[]>([])
  const [rosterChangeRequests, setRosterChangeRequests] = useState<RosterChangeRequest[]>([])
  const [isRosterChangeModalOpen, setIsRosterChangeModalOpen] = useState(false)
  const [rosterChangeCell, setRosterChangeCell] = useState<{
    date: string
    staffNo: string
    staffName: string
    level: string
    code: string
  } | null>(null)
  const [rosterChangeType, setRosterChangeType] = useState<RosterChangeRequest["changeType"]>("Shift Reassign")
  const [rosterChangeToCode, setRosterChangeToCode] = useState("")
  const [rosterChangeTargetStaff, setRosterChangeTargetStaff] = useState("")
  const [rosterChangeReason, setRosterChangeReason] = useState("")
  const [rosterChangeConflictToCode, setRosterChangeConflictToCode] = useState("")
  const [rosterExportRange, setRosterExportRange] = useState(() => getCurrentMonthRange())
  const [rosterUploadAdminMode, setRosterUploadAdminMode] = useState(false)
  const [rosterDirectEditMode, setRosterDirectEditMode] = useState(false)
  const [isAuthReady, setIsAuthReady] = useState(() => typeof window !== "undefined")
  const [authStaffId, setAuthStaffId] = useState<string | null>(() => {
    if (typeof window === "undefined") return null
    return window.localStorage.getItem(AUTH_STORAGE_KEY)
  })
  const [loginStaffNo, setLoginStaffNo] = useState("")
  const [loginPassword, setLoginPassword] = useState("")
  const [showLoginPassword, setShowLoginPassword] = useState(false)
  const [loginError, setLoginError] = useState("")
  const [mustChangePassword, setMustChangePassword] = useState(false)
  const [newPassword, setNewPassword] = useState("")
  const [confirmNewPassword, setConfirmNewPassword] = useState("")
  const [passwordChangeError, setPasswordChangeError] = useState("")
  const monthViewRange = useMemo(() => {
    const currentMonth = getCurrentMonthRange()
    const maxGeneratedDate = store.roster.reduce<string>((max, row) => {
      const date = (row.date || "").trim()
      if (!date) return max
      return date > max ? date : max
    }, "")

    return {
      fromDate: currentMonth.fromDate,
      toDate:
        maxGeneratedDate && maxGeneratedDate > currentMonth.toDate
          ? maxGeneratedDate
          : currentMonth.toDate,
    }
  }, [store.roster])
  const defaultGenerationRange = useMemo(() => getDefaultRosterGenerationRange(), [])
  const [rosterSettings, setRosterSettings] = useState({
    fromDate: defaultGenerationRange.fromDate,
    toDate: defaultGenerationRange.toDate,
    staffId: "ALL",
    offType: "Fixed",
    fixedOffWeekdays: "Friday",
    maxShiftHours: "49",
    rotationalRestHours: "24",
    shiftCode: "",
  })
  const [hydrated, setHydrated] = useState(false)
  const [syncStatus, setSyncStatus] = useState<"loading" | "saving" | "synced" | "local-only" | "error">(
    "loading",
  )

  const activeConfig = useMemo(
    () => MODULES.find((m) => m.key === activeModule) ?? MODULES[0],
    [activeModule],
  )
  const isCrewOperationBlankModule =
    activeConfig.key === "crewRostering" ||
    activeConfig.key === "crewLeavePlanner" ||
    activeConfig.key === "crewTrainingPlan"
  const filteredActiveRows = useMemo(() => {
    if (activeConfig.key !== "crewDataBase") return store[activeConfig.key]
    const q = normalizeText(crewDbSearch)
    if (!q) return store.crewDataBase
    return store.crewDataBase.filter((row) => {
      const crewCode = normalizeText(row.crewCode || "")
      const crewName = normalizeText(row.crewName || "")
      const staffId = normalizeText(row.employeeNo || "")
      return crewCode.includes(q) || crewName.includes(q) || staffId.includes(q)
    })
  }, [activeConfig.key, crewDbSearch, store])
  const crewTypes = useMemo(
    () =>
      Array.from(
        new Set(
          store.crewDataBase
            .map((row) => (row.crewType || "").trim())
            .filter(Boolean),
        ),
      ).sort(),
    [store.crewDataBase],
  )
  const crewLeavePolicies = useMemo(
    () =>
      store.crewLeavePlanner.filter(
        (row) => (row.recordType || "").trim().toLowerCase() === "policy",
      ),
    [store.crewLeavePlanner],
  )
  const crewLeaveAssignments = useMemo(
    () =>
      store.crewLeavePlanner.filter(
        (row) => (row.recordType || "").trim().toLowerCase() === "assignment",
      ),
    [store.crewLeavePlanner],
  )
  const crewLeaveMarks = useMemo(
    () =>
      store.crewLeavePlanner.filter(
        (row) => (row.recordType || "").trim().toLowerCase() === "mark",
      ),
    [store.crewLeavePlanner],
  )
  const crewGeneratedBlocks = useMemo(
    () =>
      store.crewLeavePlanner.filter(
        (row) => (row.recordType || "").trim().toLowerCase() === "generatedblock",
      ),
    [store.crewLeavePlanner],
  )
  const crewGeneratedRosterBars = useMemo(
    () =>
      store.crewLeavePlanner.filter(
        (row) => (row.recordType || "").trim().toLowerCase() === "generatedroster",
      ),
    [store.crewLeavePlanner],
  )
  const crewOpsCodeConfigs = useMemo(
    () =>
      store.crewLeavePlanner.filter(
        (row) => (row.recordType || "").trim().toLowerCase() === "codeconfig",
      ),
    [store.crewLeavePlanner],
  )
  const crewDutyCodeConfigs = useMemo(
    () =>
      store.crewLeavePlanner.filter(
        (row) => (row.recordType || "").trim().toLowerCase() === "dutycodeconfig",
      ),
    [store.crewLeavePlanner],
  )
  const crewDutyTypeConfigs = useMemo(() => {
    const map = new Map<string, Entry>()
    DEFAULT_CREW_DUTY_TYPES.forEach((row) => {
      map.set(row.dutyType.toLowerCase(), {
        id: `default-duty-type-${row.dutyType.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`,
        createdAt: "",
        recordType: "dutyTypeConfig",
        dutyType: row.dutyType,
        rosterBarMap: row.rosterBarMap,
        isDefault: "Yes",
      })
    })
    store.crewLeavePlanner
      .filter((row) => (row.recordType || "").trim().toLowerCase() === "dutytypeconfig")
      .forEach((row) => {
        const dutyType = (row.dutyType || "").trim()
        if (!dutyType) return
        if (row.hidden === "Yes") {
          map.delete(dutyType.toLowerCase())
          return
        }
        map.set(dutyType.toLowerCase(), { ...row, isDefault: row.isDefault || "No" })
      })
    return Array.from(map.values()).sort((a, b) =>
      (a.dutyType || "").localeCompare(b.dutyType || "", undefined, { sensitivity: "base" }),
    )
  }, [store.crewLeavePlanner])
  const crewDutyTypeMap = useMemo(() => {
    const map = new Map<string, string>()
    crewDutyTypeConfigs.forEach((row) => {
      const dutyType = (row.dutyType || "").trim()
      if (!dutyType) return
      map.set(dutyType.toLowerCase(), row.rosterBarMap === "REST" ? "REST" : "ROSTER DAY")
    })
    return map
  }, [crewDutyTypeConfigs])
  const crewDutyCodeColorMap = useMemo(() => {
    const map = new Map<string, string>()
    crewDutyCodeConfigs.forEach((row) => {
      const code = (row.code || "").trim().toUpperCase()
      if (!code) return
      map.set(code, row.color || getDefaultDutyCodeColor(code))
    })
    return map
  }, [crewDutyCodeConfigs])
  const crewOpsCodeAllowMap = useMemo(() => {
    const map = new Map<string, boolean>()
    crewOpsCodeConfigs.forEach((row) => {
      const code = (row.code || "").trim().toUpperCase()
      if (!code) return
      map.set(code, isYesLike(row.allowLeaveDeduction))
    })
    return map
  }, [crewOpsCodeConfigs])
  const crewOpsDeductionCodes = useMemo(
    () =>
      crewOpsCodeConfigs
        .filter((row) => isYesLike(row.allowLeaveDeduction))
        .map((row) => ({
          code: (row.code || "").trim().toUpperCase(),
          name: (row.codeName || "").trim() || (row.code || "").trim().toUpperCase(),
        }))
        .filter((row) => Boolean(row.code)),
    [crewOpsCodeConfigs],
  )
  const crewWithReleaseDate = useMemo(
    () => store.crewDataBase.filter((c) => Boolean((c.releaseDate || "").trim())),
    [store.crewDataBase],
  )
  const filteredCrewGeneratedBlocks = useMemo(() => {
    if (!isCrewBlocksActive) return [] as Entry[]
    return crewGeneratedBlocks.filter((b) => {
      const crewCode = normalizeCrewCode(b.crewCode || "")
      const startDate = (b.startDate || "").trim()
      const endDate = (b.endDate || "").trim()
      if (crewLeaveBlocksFilter.crewCode && crewCode !== crewLeaveBlocksFilter.crewCode) return false
      if (crewLeaveBlocksFilter.fromDate && endDate < crewLeaveBlocksFilter.fromDate) return false
      if (crewLeaveBlocksFilter.toDate && startDate > crewLeaveBlocksFilter.toDate) return false
      return true
    })
  }, [crewGeneratedBlocks, crewLeaveBlocksFilter, isCrewBlocksActive])
  const crewLeaveTimelineView = useMemo(() => {
    if (!isCrewTimelineActive) {
      return {
        dates: [] as string[],
        totalDays: 0,
        crewRows: [] as Array<{ crewCode: string; crewName: string; blocks: Entry[]; rosterBars: Entry[] }>,
      }
    }
    const fromDate = (crewLeaveTimelineFilter.fromDate || "").trim()
    const toDate = (crewLeaveTimelineFilter.toDate || "").trim()
    if (!fromDate || !toDate || toDate < fromDate) {
      return {
        dates: [] as string[],
        totalDays: 0,
        crewRows: [] as Array<{ crewCode: string; crewName: string; blocks: Entry[]; rosterBars: Entry[] }>,
      }
    }

    const dates = getDateRangeInclusive(fromDate, toDate).map((d) => toYmd(d)).slice(0, 370)
    const totalDays = dates.length
    const blocksInWindow = crewGeneratedBlocks.filter((b) => {
      const start = (b.startDate || "").trim()
      const end = (b.endDate || "").trim()
      if (!start || !end) return false
      return overlapDaysInclusive(start, end, fromDate, toDate) > 0
    })

    const rosterInWindow = crewGeneratedRosterBars.filter((r) => {
      const start = (r.startDate || "").trim()
      const end = (r.endDate || "").trim()
      if (!start || !end) return false
      return overlapDaysInclusive(start, end, fromDate, toDate) > 0
    })
    const crewNameByCode = new Map<string, string>()
    store.crewDataBase.forEach((c) => {
      const code = normalizeCrewCode(c.crewCode || "")
      if (!code) return
      crewNameByCode.set(code, (c.crewName || "").trim())
    })
    const actualLeaveRosterBars = store.crewLeavePlanner
      .filter((row) => (row.recordType || "").trim().toLowerCase() === "actualleavesegment")
      .filter((row) => {
        const start = (row.startDate || "").trim()
        const end = (row.endDate || "").trim()
        if (!start || !end) return false
        return overlapDaysInclusive(start, end, fromDate, toDate) > 0
      })
      .map((row) => {
        const crewCode = normalizeCrewCode(row.crewCode || "")
        return {
          ...row,
          crewCode,
          crewName: row.crewName || crewNameByCode.get(crewCode) || "",
          rosterType: "ACTUAL LEAVE",
          startTime: "00:00",
          endTime: "23:59",
        }
      })

    const crewMap = new Map<string, { crewCode: string; crewName: string; blocks: Entry[]; rosterBars: Entry[] }>()
    blocksInWindow.forEach((b) => {
      const crewCode = normalizeCrewCode(b.crewCode || "")
      if (!crewCode) return
      if (!crewMap.has(crewCode)) {
        crewMap.set(crewCode, {
          crewCode,
          crewName: (b.crewName || "").trim(),
          blocks: [],
          rosterBars: [],
        })
      }
      crewMap.get(crewCode)?.blocks.push(b)
    })
    rosterInWindow.forEach((r) => {
      const crewCode = normalizeCrewCode(r.crewCode || "")
      if (!crewCode) return
      if (!crewMap.has(crewCode)) {
        crewMap.set(crewCode, {
          crewCode,
          crewName: (r.crewName || "").trim(),
          blocks: [],
          rosterBars: [],
        })
      }
      crewMap.get(crewCode)?.rosterBars.push(r)
    })
    actualLeaveRosterBars.forEach((r) => {
      const crewCode = normalizeCrewCode(r.crewCode || "")
      if (!crewCode) return
      if (!crewMap.has(crewCode)) {
        crewMap.set(crewCode, {
          crewCode,
          crewName: (r.crewName || "").trim(),
          blocks: [],
          rosterBars: [],
        })
      }
      crewMap.get(crewCode)?.rosterBars.push(r)
    })
    crewMap.forEach((row) => {
      row.blocks.sort((a, b) => (a.startDate || "").localeCompare(b.startDate || ""))
      const seenActualLeaveBars = new Set<string>()
      row.rosterBars = row.rosterBars.filter((bar) => {
        if ((bar.rosterType || "").trim().toUpperCase() !== "ACTUAL LEAVE") return true
        const key = [
          normalizeCrewCode(bar.crewCode || ""),
          (bar.sourceBlockId || "").trim() || (bar.cycleNumber || "").trim(),
          (bar.startDate || "").trim(),
          (bar.endDate || bar.startDate || "").trim(),
        ].join("::")
        if (seenActualLeaveBars.has(key)) return false
        seenActualLeaveBars.add(key)
        return true
      })
      row.rosterBars.sort((a, b) => {
        const aActual = (a.rosterType || "").trim().toUpperCase() === "ACTUAL LEAVE"
        const bActual = (b.rosterType || "").trim().toUpperCase() === "ACTUAL LEAVE"
        if (aActual !== bActual) return aActual ? 1 : -1
        return (a.startDate || "").localeCompare(b.startDate || "")
      })
    })
    const staffNoByCrewCode = new Map<string, string>()
    store.crewDataBase.forEach((c) => {
      const code = normalizeCrewCode(c.crewCode || "")
      if (!code) return
      staffNoByCrewCode.set(code, (c.employeeNo || "").trim())
    })
    let crewRows = Array.from(crewMap.values()).sort((a, b) => {
      const aStaff = staffNoByCrewCode.get(a.crewCode) || ""
      const bStaff = staffNoByCrewCode.get(b.crewCode) || ""
      const aNum = Number(aStaff)
      const bNum = Number(bStaff)
      const bothNumeric = Number.isFinite(aNum) && Number.isFinite(bNum) && aStaff !== "" && bStaff !== ""
      if (bothNumeric && aNum !== bNum) return aNum - bNum
      const byStaffText = aStaff.localeCompare(bStaff, undefined, { numeric: true, sensitivity: "base" })
      if (byStaffText !== 0) return byStaffText
      return a.crewCode.localeCompare(b.crewCode)
    })
    if (crewLeaveTimelineCrewTypeFilter) {
      crewRows = crewRows.filter((row) =>
        row.blocks.some((b) => (b.crewType || "").trim() === crewLeaveTimelineCrewTypeFilter),
      )
    }
    return { dates, totalDays, crewRows }
  }, [
    crewGeneratedBlocks,
    crewGeneratedRosterBars,
    crewLeaveTimelineFilter,
    crewLeaveTimelineCrewTypeFilter,
    store.crewDataBase,
    store.crewLeavePlanner,
    isCrewTimelineActive,
  ])
  useEffect(() => {
    if (!isCrewTimelineActive) return
    setCrewLeaveTimelineFilter(getCurrentYearRange())
  }, [isCrewTimelineActive])
  useEffect(() => {
    if (!isCrewTimelineActive) return
    const el = crewLeaveTimelineScrollRef.current
    if (!el) return
    const today = toYmd(new Date())
    const idx = crewLeaveTimelineView.dates.indexOf(today)
    if (idx < 0) return
    const laneStartX = 220
    const dayWidth = crewLeaveTimelineDayWidth
    const centerToday = () => {
      const targetX =
        laneStartX + idx * dayWidth + dayWidth / 2 - Math.max((el.clientWidth - laneStartX) / 2, 0)
      el.scrollLeft = Math.max(targetX, 0)
    }
    centerToday()
    const ids = [80, 220, 500].map((delay) => window.setTimeout(centerToday, delay))
    return () => ids.forEach((id) => window.clearTimeout(id))
  }, [isCrewTimelineActive, crewLeaveTimelineView.dates, crewLeaveTimelineDayWidth])
  const computedCrewGeneratedBlocks = useMemo(() => {
    const byCrewCycle = new Map<string, { work?: Entry; leave?: Entry }>()
    crewGeneratedBlocks.forEach((b) => {
      const key = `${normalizeCrewCode(b.crewCode || "")}::${b.cycleNumber || ""}`
      const slot = byCrewCycle.get(key) || {}
      if ((b.blockType || "").toLowerCase() === "work") slot.work = b
      if ((b.blockType || "").toLowerCase() === "leave") slot.leave = b
      byCrewCycle.set(key, slot)
    })

    const marksByCrew = new Map<string, Entry[]>()
    crewLeaveMarks.forEach((m) => {
      const key = normalizeCrewCode(m.crewCode || "")
      const arr = marksByCrew.get(key) || []
      arr.push(m)
      marksByCrew.set(key, arr)
    })

    const canDeduct = {
      MC: crewOpsCodeAllowMap.get("MC") === true,
      AB: crewOpsCodeAllowMap.get("AB") === true,
      SD: crewOpsCodeAllowMap.get("SD") === true,
    }

    const out: Array<Entry & { actualDays?: string; deductionDays?: string; delta?: string }> = []
    crewGeneratedBlocks.forEach((b) => {
      if ((b.blockType || "").toLowerCase() !== "leave") {
        out.push({ ...b, actualDays: b.plannedDays || "0", deductionDays: "0", delta: "0" })
        return
      }
      const crewCode = normalizeCrewCode(b.crewCode || "")
      const cycle = b.cycleNumber || ""
      const key = `${crewCode}::${cycle}`
      const pair = byCrewCycle.get(key)
      const marks = (marksByCrew.get(crewCode) || []).filter((m) => {
        const markCycle = (m.cycleNumber || "").trim()
        return markCycle === cycle
      })
      let thresholds: Record<string, string> = {}
      try {
        thresholds = JSON.parse(b.codeThresholdsJson || "{}") as Record<string, string>
      } catch {
        thresholds = {}
      }
      const marksByCode = new Map<string, number>()
      marks.forEach((m) => {
        const code = (m.markType || "").trim().toUpperCase()
        if (!code) return
        const days = Number(m.noOfDays || "0") || 0
        marksByCode.set(code, (marksByCode.get(code) || 0) + days)
      })
      let deduction = 0
      const codesToEvaluate = new Set<string>([
        ...Object.keys(thresholds || {}).map((k) => (k || "").trim().toUpperCase()).filter(Boolean),
        ...Array.from(crewOpsCodeAllowMap.entries())
          .filter(([, allow]) => allow)
          .map(([code]) => (code || "").trim().toUpperCase())
          .filter(Boolean),
      ])
      codesToEvaluate.forEach((code) => {
        const total = marksByCode.get(code) || 0
        const threshold = Number((thresholds || {})[code] || "0") || 0
        const enabled =
          code === "MC"
            ? canDeduct.MC
            : code === "AB"
              ? canDeduct.AB
              : code === "SD"
                ? canDeduct.SD
                : crewOpsCodeAllowMap.get(code) === true
        if (!enabled) return
        deduction += Math.max(total - threshold, 0)
      })
      const planned = Number(b.plannedDays || "0") || 0
      const actual = Math.max(planned - deduction, 0)
      out.push({
        ...b,
        actualDays: String(actual),
        deductionDays: String(deduction),
        delta: String(actual - planned),
      })
    })
    return out
  }, [crewGeneratedBlocks, crewLeaveMarks, crewOpsCodeAllowMap])
  const selectedCrewGeneratedSummary = useMemo(() => {
    const selectedCrewCode = normalizeCrewCode(crewLeaveAssignmentForm.crewCode || "")
    if (!selectedCrewCode) return null
    const rows = crewGeneratedBlocks.filter((b) => normalizeCrewCode(b.crewCode || "") === selectedCrewCode)
    if (!rows.length) {
      return {
        crewCode: selectedCrewCode,
        exists: false,
        totalBlocks: 0,
        firstDate: "",
        lastDate: "",
        cycleCount: 0,
      }
    }
    const starts = rows.map((r) => (r.startDate || "").trim()).filter(Boolean).sort()
    const ends = rows.map((r) => (r.endDate || "").trim()).filter(Boolean).sort()
    const cycleSet = new Set(rows.map((r) => (r.cycleNumber || "").trim()).filter(Boolean))
    return {
      crewCode: selectedCrewCode,
      exists: true,
      totalBlocks: rows.length,
      firstDate: starts[0] || "",
      lastDate: ends[ends.length - 1] || "",
      cycleCount: cycleSet.size,
    }
  }, [crewLeaveAssignmentForm.crewCode, crewGeneratedBlocks])
  const dailyAttendanceCycleOptions = useMemo(() => {
    const crewCode = normalizeCrewCode(crewLeaveMarkForm.crewCode || "")
    if (!crewCode) return [] as string[]
    const set = new Set<string>()
    computedCrewGeneratedBlocks.forEach((b) => {
      if (normalizeCrewCode(b.crewCode || "") !== crewCode) return
      if ((b.blockType || "").trim().toLowerCase() !== "work") return
      const c = (b.cycleNumber || "").trim()
      if (c) set.add(c)
    })
    return Array.from(set).sort((a, b) => (Number(a) || 0) - (Number(b) || 0))
  }, [computedCrewGeneratedBlocks, crewLeaveMarkForm.crewCode])

  const findMatchedWorkCycle = (crewCode: string, fromDate: string, toDate: string): string => {
    const rows = computedCrewGeneratedBlocks
      .filter((b) => normalizeCrewCode(b.crewCode || "") === crewCode)
      .filter((b) => (b.blockType || "").trim().toLowerCase() === "work")
    const latestCreatedAt = rows.reduce((max, row) => {
      const ts = (row.createdAt || "").trim()
      if (!ts) return max
      return !max || ts > max ? ts : max
    }, "")
    const scoreRows = (input: Entry[]) =>
      input
        .map((b) => {
          const ws = (b.startDate || "").trim()
          const we = (b.endDate || "").trim()
          return {
            cycle: (b.cycleNumber || "").trim(),
            overlap: overlapDaysInclusive(fromDate, toDate, ws, we),
            createdAt: (b.createdAt || "").trim(),
          }
        })
        .filter((x) => x.cycle && x.overlap > 0)
        .sort(
          (a, b) =>
            b.overlap - a.overlap ||
            b.createdAt.localeCompare(a.createdAt) ||
            (Number(a.cycle) || 0) - (Number(b.cycle) || 0),
        )
    const latestMatches = scoreRows(rows.filter((b) => !latestCreatedAt || (b.createdAt || "").trim() === latestCreatedAt))
    if (latestMatches.length) return latestMatches[0].cycle
    const anyMatches = scoreRows(rows)
    return anyMatches[0]?.cycle || ""
  }

  useEffect(() => {
    const crewCode = normalizeCrewCode(crewLeaveMarkForm.crewCode || "")
    const fromDate = (crewLeaveMarkForm.fromDate || "").trim()
    const toDate = (crewLeaveMarkForm.toDate || "").trim() || fromDate
    if (!crewCode || !fromDate || !toDate) return
    const autoCycle = findMatchedWorkCycle(crewCode, fromDate, toDate)
    setCrewLeaveMarkForm((prev) => {
      if ((prev.cycleNumber || "").trim() === autoCycle) return prev
      return { ...prev, cycleNumber: autoCycle }
    })
  }, [
    computedCrewGeneratedBlocks,
    crewLeaveMarkForm.crewCode,
    crewLeaveMarkForm.fromDate,
    crewLeaveMarkForm.toDate,
    crewLeaveMarkForm.cycleNumber,
  ])
  const leavePatternManagerCrewMatches = useMemo(() => {
    const q = normalizeText(leavePatternManagerSearch)
    if (!q) return store.crewDataBase
    return store.crewDataBase.filter((c) => {
      const code = normalizeText(c.crewCode || "")
      const empNo = normalizeText(c.employeeNo || "")
      const name = normalizeText(c.crewName || "")
      return code.includes(q) || empNo.includes(q) || name.includes(q)
    })
  }, [leavePatternManagerSearch, store.crewDataBase])
  useEffect(() => {
    if (!leavePatternManagerCrewMatches.length) {
      if (leavePatternManagerCrewCode) setLeavePatternManagerCrewCode("")
      return
    }
    const exists = leavePatternManagerCrewMatches.some(
      (c) => normalizeCrewCode(c.crewCode || "") === leavePatternManagerCrewCode,
    )
    if (exists) return
    setLeavePatternManagerCrewCode(normalizeCrewCode(leavePatternManagerCrewMatches[0].crewCode || ""))
  }, [leavePatternManagerCrewMatches, leavePatternManagerCrewCode])
  const leavePatternManagerRows = useMemo(() => {
    const selectedCode = normalizeCrewCode(leavePatternManagerCrewCode || "")
    if (!selectedCode) return [] as Array<Record<string, string>>
    const actualUsedByCycle = new Map<string, number>()
    store.crewLeavePlanner
      .filter((row) => (row.recordType || "").trim().toLowerCase() === "actualleavesegment")
      .filter((row) => normalizeCrewCode(row.crewCode || "") === selectedCode)
      .forEach((row) => {
        const cycle = (row.cycleNumber || "").trim()
        if (!cycle) return
        const d = Number(row.plannedDays || "0")
        const prev = actualUsedByCycle.get(cycle) || 0
        actualUsedByCycle.set(cycle, prev + (Number.isFinite(d) ? d : 0))
      })
    const crewRows = computedCrewGeneratedBlocks.filter((b) => normalizeCrewCode(b.crewCode || "") === selectedCode)
    const latestCreatedAt = crewRows.reduce((max, row) => {
      const ts = (row.createdAt || "").trim()
      if (!ts) return max
      return !max || ts > max ? ts : max
    }, "")
    const blockRows = crewRows
      .filter((b) => !latestCreatedAt || (b.createdAt || "").trim() === latestCreatedAt)
      .sort((a, b) => {
        const ca = Number(a.cycleNumber || "0") || 0
        const cb = Number(b.cycleNumber || "0") || 0
        if (ca !== cb) return ca - cb
        return (a.blockType || "").localeCompare(b.blockType || "")
      })
    const byCycle = new Map<string, { work?: Entry; leave?: Entry; leaveActual?: string; leaveDeduction?: string }>()
    blockRows.forEach((b) => {
      const k = (b.cycleNumber || "").trim() || "-"
      const slot = byCycle.get(k) || {}
      if ((b.blockType || "").toLowerCase() === "work") slot.work = b
      if ((b.blockType || "").toLowerCase() === "leave") {
        slot.leave = b
        slot.leaveActual = b.actualDays || b.plannedDays || "0"
        slot.leaveDeduction = b.deductionDays || "0"
      }
      byCycle.set(k, slot)
    })
    return Array.from(byCycle.entries())
      .sort((a, b) => (Number(a[0]) || 0) - (Number(b[0]) || 0))
      .map(([cycle, v]) => ({
        ...(() => {
          const planned = Number(v.leave?.plannedDays || "0") || 0
          const deduction = Number(v.leaveDeduction || "0") || 0
          const baseAfterDeduction = Math.max(planned - deduction, 0)
          const usedByActualSegments = actualUsedByCycle.get(cycle) || 0
          const rem = Math.max(baseAfterDeduction - usedByActualSegments, 0)
          return { rem: String(rem) }
        })(),
        id: `rt${cycle}`,
        cycleNumber: cycle,
        cycleLabel: `RT${cycle}`,
        wpStart: v.work?.startDate || "",
        wpEnd: v.work?.endDate || "",
        wpDays: v.work?.plannedDays || "0",
        adjust: v.leaveDeduction || "0",
        lpStart: v.leave?.startDate || "",
        lpEnd: v.leave?.endDate || "",
        lpDays: v.leave?.plannedDays || "0",
      }))
  }, [computedCrewGeneratedBlocks, leavePatternManagerCrewCode, store.crewLeavePlanner])
  const leaveBlockByCrewCycle = useMemo(() => {
    const map = new Map<string, Entry>()
    const latestByCrew = new Map<string, string>()
    computedCrewGeneratedBlocks.forEach((b) => {
      const crewCode = normalizeCrewCode(b.crewCode || "")
      if (!crewCode) return
      const ts = (b.createdAt || "").trim()
      if (!ts) return
      const prev = latestByCrew.get(crewCode) || ""
      if (!prev || ts > prev) latestByCrew.set(crewCode, ts)
    })
    computedCrewGeneratedBlocks.forEach((b) => {
      if ((b.blockType || "").toLowerCase() !== "leave") return
      const crewCode = normalizeCrewCode(b.crewCode || "")
      const latestTs = latestByCrew.get(crewCode) || ""
      if (latestTs && (b.createdAt || "").trim() !== latestTs) return
      const key = `${crewCode}::${(b.cycleNumber || "").trim()}`
      const existing = map.get(key)
      if (!existing || (b.createdAt || "") > (existing.createdAt || "")) map.set(key, b)
    })
    return map
  }, [computedCrewGeneratedBlocks])
  const leavePatternActualSegments = useMemo(
    () =>
      store.crewLeavePlanner.filter(
        (row) => (row.recordType || "").trim().toLowerCase() === "actualleavesegment",
      ),
    [store.crewLeavePlanner],
  )
  const leavePatternSegmentsByCycle = useMemo(() => {
    const selectedCode = normalizeCrewCode(leavePatternManagerCrewCode || "")
    const byCycle = new Map<string, Entry[]>()
    leavePatternActualSegments.forEach((row) => {
      if (normalizeCrewCode(row.crewCode || "") !== selectedCode) return
      const cycle = (row.cycleNumber || "").trim()
      if (!cycle) return
      const arr = byCycle.get(cycle) || []
      arr.push(row)
      byCycle.set(cycle, arr)
    })
    byCycle.forEach((arr, cycle) => {
      arr.sort((a, b) => (a.startDate || "").localeCompare(b.startDate || ""))
      byCycle.set(cycle, arr)
    })
    return byCycle
  }, [leavePatternActualSegments, leavePatternManagerCrewCode])
  const leaveDetailActiveCycleNumber = (leaveDetailSelectedCycle || leaveDetailModal.cycleNumber || "").trim()
  const leaveDetailSegmentsForActiveCycle = useMemo(() => {
    if (!leaveDetailModal.open) return [] as Entry[]
    return (
      store.crewLeavePlanner
        .filter((row) => (row.recordType || "").trim().toLowerCase() === "actualleavesegment")
        .filter((row) => normalizeCrewCode(row.crewCode || "") === normalizeCrewCode(leaveDetailModal.crewCode))
        .filter((row) => (row.cycleNumber || "").trim() === leaveDetailActiveCycleNumber)
        .sort((a, b) => (a.startDate || "").localeCompare(b.startDate || ""))
    )
  }, [store.crewLeavePlanner, leaveDetailModal, leaveDetailActiveCycleNumber])
  const leaveDetailUnassignedActualLeaveSegments = useMemo(() => {
    if (!leaveDetailModal.open) return [] as Entry[]
    const crewCode = normalizeCrewCode(leaveDetailModal.crewCode || "")
    return store.crewLeavePlanner
      .filter((row) => (row.recordType || "").trim().toLowerCase() === "actualleavesegment")
      .filter((row) => normalizeCrewCode(row.crewCode || "") === crewCode)
      .filter((row) => !(row.cycleNumber || "").trim() || !(row.sourceBlockId || "").trim())
      .sort((a, b) => (a.startDate || "").localeCompare(b.startDate || ""))
  }, [store.crewLeavePlanner, leaveDetailModal])
  const leaveDetailSelectedActualLeaveSegment = useMemo(() => {
    const segmentId = (leaveDetailModal.actualLeaveSegmentId || "").trim()
    if (!segmentId) return null
    return (
      store.crewLeavePlanner.find(
        (row) =>
          row.id === segmentId &&
          (row.recordType || "").trim().toLowerCase() === "actualleavesegment",
      ) || null
    )
  }, [store.crewLeavePlanner, leaveDetailModal.actualLeaveSegmentId])
  const leaveDetailGeneratedRows = useMemo(() => {
    if (!leaveDetailModal.open) return [] as Entry[]
    const crewCode = normalizeCrewCode(leaveDetailModal.crewCode || "")
    const selectedPolicy = (leaveDetailModal.policyName || "").trim()
    const selectedBlockId = (leaveDetailModal.blockId || "").trim()

    const selectedBlock = computedCrewGeneratedBlocks.find((b) => (b.id || "").trim() === selectedBlockId)
    const selectedCreatedAt = (selectedBlock?.createdAt || "").trim()

    const base = computedCrewGeneratedBlocks
      .filter((b) => normalizeCrewCode(b.crewCode || "") === crewCode)
      .filter((b) => !selectedPolicy || (b.policyName || "").trim() === selectedPolicy)
      .filter((b) => !selectedCreatedAt || (b.createdAt || "").trim() === selectedCreatedAt)

    const byCycleType = new Map<string, Entry>()
    base.forEach((b) => {
      const key = `${(b.cycleNumber || "").trim()}::${(b.blockType || "").trim().toLowerCase()}`
      const existing = byCycleType.get(key)
      if (!existing) {
        byCycleType.set(key, b)
        return
      }
      if ((b.createdAt || "") > (existing.createdAt || "")) {
        byCycleType.set(key, b)
      }
    })
    return Array.from(byCycleType.values()).sort((a, b) => {
      const ca = Number(a.cycleNumber || "0") || 0
      const cb = Number(b.cycleNumber || "0") || 0
      if (ca !== cb) return ca - cb
      const sa = (a.startDate || "").trim()
      const sb = (b.startDate || "").trim()
      if (sa !== sb) return sa.localeCompare(sb)
      return (a.blockType || "").localeCompare(b.blockType || "")
    })
  }, [computedCrewGeneratedBlocks, leaveDetailModal])
  const leaveDetailPatternRows = useMemo(() => {
    const usedByCycle = new Map<string, number>()
    if (leaveDetailModal.open) {
      store.crewLeavePlanner
        .filter((row) => (row.recordType || "").trim().toLowerCase() === "actualleavesegment")
        .filter((row) => normalizeCrewCode(row.crewCode || "") === normalizeCrewCode(leaveDetailModal.crewCode))
        .forEach((row) => {
          const cycle = (row.cycleNumber || "").trim()
          if (!cycle) return
          const d = Number(row.plannedDays || "0")
          const prev = usedByCycle.get(cycle) || 0
          usedByCycle.set(cycle, prev + (Number.isFinite(d) ? d : 0))
        })
    }
    const byCycle = new Map<string, { work?: Entry; leave?: Entry }>()
    leaveDetailGeneratedRows.forEach((b) => {
      const cycle = (b.cycleNumber || "").trim() || "-"
      const slot = byCycle.get(cycle) || {}
      if ((b.blockType || "").toLowerCase() === "work") slot.work = b
      if ((b.blockType || "").toLowerCase() === "leave") slot.leave = b
      byCycle.set(cycle, slot)
    })
    return Array.from(byCycle.entries())
      .sort((a, b) => (Number(a[0]) || 0) - (Number(b[0]) || 0))
      .map(([cycle, v]) => ({
        ...(() => {
          const planned = Number(v.leave?.plannedDays || "0") || 0
          const deduction = Number(v.leave?.deductionDays || "0") || 0
          const remBase = Math.max(planned - deduction, 0)
          const used = usedByCycle.get(cycle) || 0
          const rem = Math.max(remBase - used, 0)
          return { rem: String(rem) }
        })(),
        id: `ld_rt${cycle}`,
        cycleNumber: cycle,
        cycleLabel: `Cycle ${cycle}`,
        wpStart: v.work?.startDate || "",
        wpEnd: v.work?.endDate || "",
        wpDays: v.work?.plannedDays || "0",
        wpOriginalStart: v.work?.originalStartDate || v.work?.startDate || "",
        wpOriginalEnd: v.work?.originalEndDate || v.work?.endDate || "",
        adjust: v.leave?.deductionDays || "0",
        lpStart: v.leave?.startDate || "",
        lpEnd: v.leave?.endDate || "",
        lpDays: v.leave?.plannedDays || "0",
        lpOriginalStart: v.leave?.originalStartDate || v.leave?.startDate || "",
        lpOriginalEnd: v.leave?.originalEndDate || v.leave?.endDate || "",
      }))
  }, [leaveDetailGeneratedRows, leaveDetailModal, store.crewLeavePlanner])
  const leaveDetailActiveCycleRow = useMemo(
    () => leaveDetailPatternRows.find((r) => (r.cycleNumber || "").trim() === leaveDetailActiveCycleNumber) || null,
    [leaveDetailPatternRows, leaveDetailActiveCycleNumber],
  )
  const leaveDetailCycleUsedDays = useMemo(
    () =>
      leaveDetailSegmentsForActiveCycle.reduce((acc, s) => {
        const d = Number(s.plannedDays || "0")
        return acc + (Number.isFinite(d) ? d : 0)
      }, 0),
    [leaveDetailSegmentsForActiveCycle],
  )
  const leaveDetailCycleRemainingDays = useMemo(() => {
    const remBase = Number(leaveDetailActiveCycleRow?.rem || "0") || 0
    return Math.max(remBase, 0)
  }, [leaveDetailActiveCycleRow])
  const leaveDetailCycleDeductionRows = useMemo(() => {
    if (!leaveDetailModal.open || !leaveDetailActiveCycleNumber) return [] as Array<{
      code: string
      markedDays: number
      threshold: number
      deductedDays: number
    }>
    const crewCode = normalizeCrewCode(leaveDetailModal.crewCode || "")
    if (!crewCode) return []
    const activeLeaveBlock = leaveDetailGeneratedRows.find(
      (b) =>
        (b.blockType || "").trim().toLowerCase() === "leave" &&
        (b.cycleNumber || "").trim() === leaveDetailActiveCycleNumber,
    )
    if (!activeLeaveBlock) return []
    let thresholds: Record<string, string> = {}
    try {
      thresholds = JSON.parse(activeLeaveBlock.codeThresholdsJson || "{}") as Record<string, string>
    } catch {
      thresholds = {}
    }
    const marksByCode = new Map<string, number>()
    crewLeaveMarks
      .filter((m) => normalizeCrewCode(m.crewCode || "") === crewCode)
      .forEach((m) => {
        const markCycle = (m.cycleNumber || "").trim()
        const code = (m.markType || "").trim().toUpperCase()
        if (!code) return
        if (markCycle !== leaveDetailActiveCycleNumber) return
        const days = Number(m.noOfDays || "0") || 0
        if (days <= 0) return
        marksByCode.set(code, (marksByCode.get(code) || 0) + days)
      })
    const codesToEvaluate = new Set<string>([
      ...Object.keys(thresholds || {}).map((k) => (k || "").trim().toUpperCase()).filter(Boolean),
      ...Array.from(crewOpsCodeAllowMap.entries())
        .filter(([, allow]) => allow)
        .map(([code]) => (code || "").trim().toUpperCase())
        .filter(Boolean),
    ])
    return Array.from(codesToEvaluate)
      .map((code) => {
        const markedDays = marksByCode.get(code) || 0
        const threshold = Number((thresholds || {})[code] || "0") || 0
        const enabled = crewOpsCodeAllowMap.get(code) === true
        const deductedDays = enabled ? Math.max(markedDays - threshold, 0) : 0
        return { code, markedDays, threshold, deductedDays }
      })
      .filter((r) => r.markedDays > 0 || r.deductedDays > 0)
      .sort((a, b) => a.code.localeCompare(b.code))
  }, [
    leaveDetailModal,
    leaveDetailActiveCycleNumber,
    leaveDetailGeneratedRows,
    crewLeaveMarks,
    crewOpsCodeAllowMap,
  ])
  const crewTimelineCrewTypes = useMemo(() => {
    const set = new Set<string>()
    crewGeneratedBlocks.forEach((b) => {
      const t = (b.crewType || "").trim()
      if (t) set.add(t)
    })
    return Array.from(set).sort((a, b) => a.localeCompare(b))
  }, [crewGeneratedBlocks])
  useEffect(() => {
    if (!crewTimelineCrewTypes.length) return
    if (crewLeaveTimelineCrewTypeFilter && crewTimelineCrewTypes.includes(crewLeaveTimelineCrewTypeFilter)) return
    if (crewTimelineCrewTypes.includes("Captain")) {
      setCrewLeaveTimelineCrewTypeFilter("Captain")
      return
    }
    setCrewLeaveTimelineCrewTypeFilter(crewTimelineCrewTypes[0] || "")
  }, [crewTimelineCrewTypes, crewLeaveTimelineCrewTypeFilter])
  const crewTypeWpLpDailySummary = useMemo(() => {
    if (!isCrewTimelineActive) {
      return [] as Array<{ date: string; crewType: string; wp: number; lp: number; total: number }>
    }
    const fromDate = (crewLeaveTimelineFilter.fromDate || "").trim()
    const toDate = (crewLeaveTimelineFilter.toDate || "").trim()
    if (!fromDate || !toDate || toDate < fromDate) {
      return [] as Array<{ date: string; crewType: string; wp: number; lp: number; total: number }>
    }
    const days = getDateRangeInclusive(fromDate, toDate).map((d) => toYmd(d))
    const out: Array<{ date: string; crewType: string; wp: number; lp: number; total: number }> = []
    days.forEach((day) => {
      const byType = new Map<string, { wpCrew: Set<string>; lpCrew: Set<string> }>()
      crewGeneratedBlocks.forEach((b) => {
        const start = (b.startDate || "").trim()
        const end = (b.endDate || "").trim()
        if (!start || !end || day < start || day > end) return
        const crewCode = normalizeCrewCode(b.crewCode || "")
        if (!crewCode) return
        const crewType = (b.crewType || "").trim() || "Unassigned"
        if (!byType.has(crewType)) byType.set(crewType, { wpCrew: new Set<string>(), lpCrew: new Set<string>() })
        const slot = byType.get(crewType)!
        const blockType = (b.blockType || "").trim().toLowerCase()
        if (blockType === "work") slot.wpCrew.add(crewCode)
        if (blockType === "leave") slot.lpCrew.add(crewCode)
      })
      Array.from(byType.entries())
        .sort((a, b) => a[0].localeCompare(b[0]))
        .forEach(([crewType, v]) => {
          out.push({
            date: day,
            crewType,
            wp: v.wpCrew.size,
            lp: v.lpCrew.size,
            total: v.wpCrew.size + v.lpCrew.size,
          })
        })
    })
    return out
  }, [crewGeneratedBlocks, crewLeaveTimelineFilter.fromDate, crewLeaveTimelineFilter.toDate, isCrewTimelineActive])
  const crewTypeDailySummaryMap = useMemo(() => {
    const byType = new Map<string, Map<string, { wp: number; lp: number }>>()
    crewTypeWpLpDailySummary.forEach((row) => {
      if (crewLeaveTimelineCrewTypeFilter && row.crewType !== crewLeaveTimelineCrewTypeFilter) return
      const typeMap = byType.get(row.crewType) || new Map<string, { wp: number; lp: number }>()
      typeMap.set(row.date, { wp: row.wp, lp: row.lp })
      byType.set(row.crewType, typeMap)
    })
    return Array.from(byType.entries()).sort((a, b) => a[0].localeCompare(b[0]))
  }, [crewTypeWpLpDailySummary, crewLeaveTimelineCrewTypeFilter])
  const deductionByGeneratedBlockId = useMemo(() => {
    const map = new Map<string, number>()
    computedCrewGeneratedBlocks.forEach((b) => {
      const id = (b.id || "").trim()
      if (!id) return
      const d = Number(b.deductionDays || "0") || 0
      map.set(id, d)
    })
    return map
  }, [computedCrewGeneratedBlocks])
  const crewTimelineMonthBoundaryIndexes = useMemo(() => {
    if (!isCrewTimelineActive) return new Set<number>()
    const set = new Set<number>()
    for (let i = 0; i < crewLeaveTimelineView.dates.length - 1; i += 1) {
      const cur = crewLeaveTimelineView.dates[i]?.slice(0, 7)
      const next = crewLeaveTimelineView.dates[i + 1]?.slice(0, 7)
      if (cur && next && cur !== next) set.add(i)
    }
    return set
  }, [crewLeaveTimelineView.dates, isCrewTimelineActive])

  // Conflict / overlap detection for the Crew Leave Planner.
  // Detects:
  //   1. Self-overlap   — same crew has two LEAVE blocks that overlap.
  //   2. Coverage       — too many crew of the same crewType on LEAVE on the same day
  //                       (threshold configurable per type, falls back to default).
  //   3. OutOfWindow    — leave block falls outside crew's joiningDate / releaseDate.
  const crewLeaveConflicts = useMemo(() => {
    type Conflict = {
      id: string
      type: "Overlap" | "Coverage" | "OutOfWindow"
      severity: "high" | "medium" | "low"
      crewCode: string
      crewName: string
      crewType: string
      startDate: string
      endDate: string
      message: string
    }
    if (!isCrewConflictsActive) return [] as Conflict[]
    const conflicts: Conflict[] = []

    // Only LEAVE blocks matter for conflict checks.
    const leaveBlocks = crewGeneratedBlocks
      .filter((b) => (b.blockType || "").toLowerCase() === "leave")
      .filter((b) => {
        const s = (b.startDate || "").trim()
        const e = (b.endDate || "").trim()
        if (!s || !e) return false
        if (crewLeaveConflictConfig.fromDate && e < crewLeaveConflictConfig.fromDate) return false
        if (crewLeaveConflictConfig.toDate && s > crewLeaveConflictConfig.toDate) return false
        if (
          crewLeaveConflictConfig.crewTypeFilter &&
          (b.crewType || "").trim() !== crewLeaveConflictConfig.crewTypeFilter
        )
          return false
        return true
      })

    // 1. Self-overlap detection: group leave blocks by crewCode and compare pairs.
    const byCrew = new Map<string, Entry[]>()
    for (const b of leaveBlocks) {
      const key = normalizeCrewCode(b.crewCode || "")
      if (!key) continue
      const list = byCrew.get(key) || []
      list.push(b)
      byCrew.set(key, list)
    }
    for (const [crewCode, list] of byCrew.entries()) {
      const sorted = [...list].sort((a, b) =>
        (a.startDate || "").localeCompare(b.startDate || ""),
      )
      for (let i = 1; i < sorted.length; i += 1) {
        const prev = sorted[i - 1]
        const curr = sorted[i]
        if ((curr.startDate || "") <= (prev.endDate || "")) {
          conflicts.push({
            id: `ovl_${prev.id}_${curr.id}`,
            type: "Overlap",
            severity: "high",
            crewCode,
            crewName: prev.crewName || curr.crewName || "",
            crewType: prev.crewType || curr.crewType || "",
            startDate: curr.startDate || "",
            endDate:
              (prev.endDate || "") > (curr.endDate || "")
                ? prev.endDate || ""
                : curr.endDate || "",
            message: `Overlapping leave blocks: ${prev.startDate}–${prev.endDate} and ${curr.startDate}–${curr.endDate}`,
          })
        }
      }
    }

    // 2. Coverage threshold: per crewType, count concurrent leaves per day.
    //    Build per-type, per-day counts and emit one conflict per (type, day) over threshold.
    const perTypePerDay = new Map<string, Map<string, Set<string>>>()
    for (const b of leaveBlocks) {
      const crewType = (b.crewType || "").trim() || "Unknown"
      const start = b.startDate || ""
      const end = b.endDate || ""
      if (!start || !end || start > end) continue
      // Cap any single block at 400 day-stamps to avoid runaway loops on bad data.
      let day = start
      let guard = 0
      const dayMap = perTypePerDay.get(crewType) || new Map<string, Set<string>>()
      perTypePerDay.set(crewType, dayMap)
      while (day <= end && guard < 400) {
        const set = dayMap.get(day) || new Set<string>()
        set.add(normalizeCrewCode(b.crewCode || ""))
        dayMap.set(day, set)
        day = addDaysToYmd(day, 1)
        guard += 1
      }
    }
    const defaultThreshold = Math.max(0, crewLeaveConflictConfig.defaultThreshold)
    for (const [crewType, dayMap] of perTypePerDay.entries()) {
      const threshold =
        crewLeaveConflictConfig.thresholdsByType[crewType] !== undefined
          ? Math.max(0, crewLeaveConflictConfig.thresholdsByType[crewType])
          : defaultThreshold
      // Collapse consecutive over-threshold days into ranges for readability.
      const sortedDays = Array.from(dayMap.keys()).sort()
      let rangeStart: string | null = null
      let rangeEnd: string | null = null
      let rangeCrewSet = new Set<string>()
      let rangeMaxCount = 0
      const flushRange = () => {
        if (!rangeStart || !rangeEnd) return
        const over = rangeMaxCount - threshold
        const severity: Conflict["severity"] =
          over >= 2 ? "high" : over === 1 ? "medium" : "low"
        conflicts.push({
          id: `cov_${crewType}_${rangeStart}_${rangeEnd}`,
          type: "Coverage",
          severity,
          crewCode: Array.from(rangeCrewSet).filter(Boolean).join(", ") || "-",
          crewName: "(multiple)",
          crewType,
          startDate: rangeStart,
          endDate: rangeEnd,
          message: `${rangeMaxCount} ${crewType} on leave (limit ${threshold})${
            rangeStart === rangeEnd ? "" : ` across ${rangeStart}–${rangeEnd}`
          }`,
        })
      }
      for (const day of sortedDays) {
        const crewSet = dayMap.get(day) || new Set<string>()
        const count = crewSet.size
        if (count > threshold) {
          if (rangeStart && rangeEnd && addDaysToYmd(rangeEnd, 1) === day) {
            rangeEnd = day
            for (const c of crewSet) rangeCrewSet.add(c)
            if (count > rangeMaxCount) rangeMaxCount = count
          } else {
            flushRange()
            rangeStart = day
            rangeEnd = day
            rangeCrewSet = new Set<string>(crewSet)
            rangeMaxCount = count
          }
        }
      }
      flushRange()
    }

    // 3. Out-of-window: leave outside crew's joining/release dates.
    const crewByCode = new Map<string, Entry>()
    for (const c of store.crewDataBase) {
      crewByCode.set(normalizeCrewCode(c.crewCode || ""), c)
    }
    for (const b of leaveBlocks) {
      const crew = crewByCode.get(normalizeCrewCode(b.crewCode || ""))
      if (!crew) continue
      const joining = (crew.joiningDate || "").trim()
      const release = (crew.releaseDate || "").trim()
      const start = (b.startDate || "").trim()
      const end = (b.endDate || "").trim()
      if (joining && start && start < joining) {
        conflicts.push({
          id: `oow_pre_${b.id}`,
          type: "OutOfWindow",
          severity: "medium",
          crewCode: normalizeCrewCode(b.crewCode || ""),
          crewName: b.crewName || crew.crewName || "",
          crewType: b.crewType || crew.crewType || "",
          startDate: start,
          endDate: end,
          message: `Leave starts ${start} before joining date ${joining}`,
        })
      }
      if (release && end && end > release) {
        conflicts.push({
          id: `oow_post_${b.id}`,
          type: "OutOfWindow",
          severity: "medium",
          crewCode: normalizeCrewCode(b.crewCode || ""),
          crewName: b.crewName || crew.crewName || "",
          crewType: b.crewType || crew.crewType || "",
          startDate: start,
          endDate: end,
          message: `Leave ends ${end} after release date ${release}`,
        })
      }
    }

    // Apply severity filter last.
    const sevFilter = crewLeaveConflictConfig.severityFilter
    const filtered =
      sevFilter === "all" ? conflicts : conflicts.filter((c) => c.severity === sevFilter)

    // Sort by severity then by start date.
    const sevRank: Record<Conflict["severity"], number> = { high: 0, medium: 1, low: 2 }
    filtered.sort((a, b) => {
      if (sevRank[a.severity] !== sevRank[b.severity]) return sevRank[a.severity] - sevRank[b.severity]
      return (a.startDate || "").localeCompare(b.startDate || "")
    })
    return filtered
  }, [crewGeneratedBlocks, store.crewDataBase, crewLeaveConflictConfig, isCrewConflictsActive])

  const crewLeaveConflictsCountsBySeverity = useMemo(() => {
    const counts = { high: 0, medium: 0, low: 0 }
    for (const c of crewLeaveConflicts) counts[c.severity] += 1
    return counts
  }, [crewLeaveConflicts])
  const rotationQuickPresets = [
    { work: 73, leave: 19 },
    { work: 72, leave: 19 },
    { work: 72, leave: 18 },
    { work: 75, leave: 15 },
    { work: 73, leave: 18 },
    { work: 60, leave: 13 },
    { work: 58, leave: 15 },
    { work: 60, leave: 15 },
    { work: 45, leave: 10 },
    { work: 50, leave: 10 },
    { work: 30, leave: 7 },
    { work: 21, leave: 7 },
  ]
  const rotationYearPreview = useMemo(() => {
    const work = Number(crewLeavePolicyForm.workPatternDays || "0")
    const leave = Number(crewLeavePolicyForm.leavePatternDays || "0")
    const cycle = work + leave
    if (!Number.isFinite(work) || !Number.isFinite(leave) || cycle <= 0) return null
    const cyclesPerYear = 365 / cycle
    const nearest = Math.round(cyclesPerYear)
    const covered = nearest * cycle
    const diff = covered - 365
    return {
      work,
      leave,
      cycle,
      cyclesPerYear,
      nearest,
      covered,
      diff,
      leavePerYear: leave * cyclesPerYear,
      leaveShare: (leave / cycle) * 100,
    }
  }, [crewLeavePolicyForm.workPatternDays, crewLeavePolicyForm.leavePatternDays])

  const [forms, setForms] = useState<Record<ModuleKey, Record<string, string>>>(() => {
    return MODULES.reduce((acc, module) => {
      acc[module.key] = buildEmptyForm(module)
      return acc
    }, {} as Record<ModuleKey, Record<string, string>>)
  })

  const savePayloadJsonToSupabase = async (payloadJson: string) => {
    if (!SUPABASE_URL || !SUPABASE_PUBLISHABLE_KEY) {
      return { error: new Error("Missing Supabase configuration.") }
    }
    const controller = new AbortController()
    const timeoutId = window.setTimeout(() => controller.abort(), REMOTE_SAVE_TIMEOUT_MS)
    try {
      const response = await fetch(`${SUPABASE_URL}/rest/v1/${SUPABASE_STORE_TABLE}?id=eq.${SUPABASE_STORE_ID}`, {
        method: "PATCH",
        headers: {
          apikey: SUPABASE_PUBLISHABLE_KEY,
          Authorization: `Bearer ${SUPABASE_PUBLISHABLE_KEY}`,
          "Content-Type": "application/json",
          Prefer: "return=minimal",
        },
        body: `{"payload":${payloadJson}}`,
        signal: controller.signal,
      })
      if (!response.ok) {
        return { error: new Error(`Supabase save failed: ${response.status}`) }
      }
      return { error: null }
    } catch (error) {
      return { error }
    } finally {
      window.clearTimeout(timeoutId)
    }
  }

  useEffect(() => {
    let mounted = true
    const loadData = async () => {
      if (mounted) setHydrated(true)
      let loaded = false
      let loadedFromLocalFirst = false
      let localFirstSnapshot: DataStore | null = null

      // Prefer local snapshot first so localhost edits (e.g. Crew DB release date)
      // are not overwritten by older remote payload on refresh.
      const rawLocalFirst = window.localStorage.getItem(STORAGE_KEY)
      if (rawLocalFirst) {
        try {
          const parsed = JSON.parse(rawLocalFirst) as DataStore
          localFirstSnapshot = {
            ...getEmptyStore(),
            ...parsed,
          }
          if (mounted) {
            setStore(localFirstSnapshot)
            if (supabase && !disableRemoteSyncRef.current) {
              setSyncStatus("loading")
            } else {
              setSyncStatus("local-only")
            }
          }
          loaded = true
          loadedFromLocalFirst = true
        } catch {
          // Ignore malformed local payload and continue remote/local fallback below.
        }
      }

      if (supabase && !disableRemoteSyncRef.current) {
        const { data, error } = await supabase
          .from(SUPABASE_STORE_TABLE)
          .select("payload")
          .eq("id", SUPABASE_STORE_ID)
          .maybeSingle()

        if (!error && data && typeof data.payload === "object" && data.payload !== null) {
          const parsed = data.payload as DataStore
          const parsedRemote = {
            ...getEmptyStore(),
            ...parsed,
          }
          const scoreStore = (value: DataStore) =>
            (value.staff?.length || 0) +
            (value.roster?.length || 0) +
            (value.opsControl?.length || 0) +
            (value.crewDataBase?.length || 0) +
            (value.crewLeavePlanner?.length || 0)
          const mergeStaffProfileUpdates = (local: DataStore, remote: DataStore): DataStore => {
            const remoteById = new Map((remote.staff || []).map((row) => [row.id, row]))
            let changed = false
            const staff = (local.staff || []).map((localRow) => {
              const remoteRow = remoteById.get(localRow.id)
              if (!remoteRow) return localRow
              const localProfileAt = Date.parse(localRow.profileUpdatedAt || "")
              const remoteProfileAt = Date.parse(remoteRow.profileUpdatedAt || "")
              const remoteIsNewer =
                Number.isFinite(remoteProfileAt) &&
                (!Number.isFinite(localProfileAt) || remoteProfileAt > localProfileAt)
              const hasRemoteProfileFields = Boolean(
                remoteRow.phoneNumber ||
                  remoteRow.addressPresent ||
                  remoteRow.avatar ||
                  remoteRow.signatureData,
              )
              const localMissingProfileFields =
                !localRow.phoneNumber ||
                !localRow.addressPresent ||
                !localRow.avatar ||
                !localRow.signatureData
              if (!remoteIsNewer && !(hasRemoteProfileFields && localMissingProfileFields)) return localRow
              changed = true
              return {
                ...localRow,
                phoneNumber: remoteRow.phoneNumber || localRow.phoneNumber || "",
                addressPresent: remoteRow.addressPresent || localRow.addressPresent || "",
                avatar: remoteRow.avatar || localRow.avatar || "",
                signatureData: remoteRow.signatureData || localRow.signatureData || "",
                profileUpdatedAt: remoteRow.profileUpdatedAt || localRow.profileUpdatedAt || "",
              }
            })
            return changed ? { ...local, staff } : local
          }
          // If local snapshot already loaded in this session, keep it as source
          // of truth and avoid overwriting with potentially older remote payload.
          if (mounted && !loadedFromLocalFirst) {
            skipNextRemoteSaveRef.current = true
            lastSyncedStoreJsonRef.current = JSON.stringify(parsedRemote)
            setStore(parsedRemote)
            try {
              window.localStorage.setItem(STORAGE_KEY, JSON.stringify(parsed))
            } catch {
              try {
                window.localStorage.setItem(STORAGE_KEY, JSON.stringify(compactStoreForLocal(parsed)))
              } catch {
                window.localStorage.setItem(
                  STORAGE_KEY,
                  JSON.stringify({
                    ...parsed,
                    opsControl: [],
                  }),
                )
              }
            }
            setSyncStatus("synced")
          } else if (mounted && loadedFromLocalFirst) {
            // If local cache looks partial/truncated but remote has richer data,
            // prefer remote and refresh local cache.
            const localScore = localFirstSnapshot ? scoreStore(localFirstSnapshot) : 0
            const remoteScore = scoreStore(parsedRemote)
            if (remoteScore > localScore) {
              skipNextRemoteSaveRef.current = true
              lastSyncedStoreJsonRef.current = JSON.stringify(parsedRemote)
              setStore(parsedRemote)
              try {
                window.localStorage.setItem(STORAGE_KEY, JSON.stringify(parsedRemote))
              } catch {
                // non-blocking
              }
              setSyncStatus("synced")
            } else {
              const merged = localFirstSnapshot
                ? mergeStaffProfileUpdates(localFirstSnapshot, parsedRemote)
                : parsedRemote
              skipNextRemoteSaveRef.current = true
              lastSyncedStoreJsonRef.current = JSON.stringify(merged)
              setStore(merged)
              try {
                window.localStorage.setItem(STORAGE_KEY, JSON.stringify(merged))
              } catch {
                // non-blocking
              }
              setSyncStatus("synced")
            }
          }
          loaded = true
        } else if (error) {
          if (mounted) setSyncStatus("error")
        }
      }

      if (!loaded && mounted) {
        const raw = window.localStorage.getItem(STORAGE_KEY)
        if (raw) {
          try {
            const parsed = JSON.parse(raw) as DataStore
            setStore({
              ...getEmptyStore(),
              ...parsed,
            })
          } catch {
            setStore(getEmptyStore())
          }
        } else {
          setStore(getEmptyStore())
        }
        if (!supabase) setSyncStatus("local-only")
      }

      if (mounted) {
        hasLoadedRef.current = true
      }
    }

    loadData()

    return () => {
      mounted = false
      if (saveTimerRef.current) {
        clearTimeout(saveTimerRef.current)
        saveTimerRef.current = null
      }
    }
  }, [supabase])

  useEffect(() => {
    if (!hydrated || !hasLoadedRef.current) return
    const storeJson = JSON.stringify(store)
    try {
      window.localStorage.setItem(STORAGE_KEY, storeJson)
    } catch {
      try {
        window.localStorage.setItem(STORAGE_KEY, JSON.stringify(compactStoreForLocal(store)))
      } catch {
        window.localStorage.setItem(
          STORAGE_KEY,
          JSON.stringify({
            ...store,
            opsControl: [],
          }),
        )
      }
    }
    try {
      window.localStorage.setItem(
        AUTH_CACHE_KEY,
        JSON.stringify({
          staff: store.staff || [],
          userManagement: store.userManagement || [],
        }),
      )
    } catch {
      // non-blocking
    }

    if (!supabase || disableRemoteSyncRef.current || !isOnlineRef.current) {
      if (!supabase || !isOnlineRef.current) setSyncStatus("local-only")
      return
    }
    if (!forceRemoteSyncRef.current && storeJson === lastSyncedStoreJsonRef.current) {
      setSyncStatus("synced")
      return
    }
    if (skipNextRemoteSaveRef.current && !forceRemoteSyncRef.current) {
      skipNextRemoteSaveRef.current = false
      lastSyncedStoreJsonRef.current = storeJson
      setSyncStatus("synced")
      return
    }
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current)

    const now = Date.now()
    const sinceLastRemoteSave = now - lastRemoteSaveStartedAtRef.current
    const intervalDelay =
      forceRemoteSyncRef.current || sinceLastRemoteSave >= REMOTE_SAVE_MIN_INTERVAL_MS
        ? 0
        : REMOTE_SAVE_MIN_INTERVAL_MS - sinceLastRemoteSave
    const delayMs = forceRemoteSyncRef.current ? 0 : Math.max(5000, intervalDelay)
    saveTimerRef.current = setTimeout(async () => {
      setSyncStatus("saving")
      let error: unknown = null
      try {
        lastRemoteSaveStartedAtRef.current = Date.now()
        const result = await savePayloadJsonToSupabase(storeJson)
        error = result.error
      } catch (err) {
        error = err
      }
      saveTimerRef.current = null

      if (error) {
        if (typeof window !== "undefined" && !window.navigator.onLine) {
          setSyncStatus("local-only")
          return
        }
        setSyncStatus("error")
        return
      }
      forceRemoteSyncRef.current = false
      lastSyncedStoreJsonRef.current = storeJson
      setSyncStatus("synced")
    }, delayMs)
  }, [store, hydrated, supabase])

  useEffect(() => {
    if (typeof window === "undefined") return
    const savedAuthStaffId = window.localStorage.getItem(AUTH_STORAGE_KEY)
    setAuthStaffId(savedAuthStaffId || null)
    setIsAuthReady(true)
  }, [])

  useEffect(() => {
    if (typeof window === "undefined") return
    const handleOnline = () => {
      isOnlineRef.current = true
      setIsOnline(true)
      forceRemoteSyncRef.current = true
      setSyncStatus("loading")
      setStore((prev) => ({ ...prev }))
    }
    const handleOffline = () => {
      isOnlineRef.current = false
      setIsOnline(false)
      setSyncStatus("local-only")
    }
    window.addEventListener("online", handleOnline)
    window.addEventListener("offline", handleOffline)
    return () => {
      window.removeEventListener("online", handleOnline)
      window.removeEventListener("offline", handleOffline)
    }
  }, [])

  // First-run bootstrap: seed a default admin user when none exist so the
  // operator can log in. Credentials: username "admin" / password "admin".
  // Runs once after hydration; no-op if any user already exists.
  useEffect(() => {
    if (!hydrated) return
    if (store.userManagement.length > 0) return
    const adminEntry: Entry = {
      id: `usr_admin_${Date.now()}`,
      createdAt: new Date().toISOString(),
      userType: "Dispatch Staff User",
      dispatchStaff: "",
      fullName: "Default Administrator",
      userName: "admin",
      email: "",
      password: "admin",
      roleName: "Administrator",
      status: "Active",
    }
    setStore((prev) =>
      prev.userManagement.length > 0
        ? prev
        : { ...prev, userManagement: [adminEntry] },
    )
    // Only run when userManagement transitions to empty after hydration.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hydrated])

  const dashboardStats = useMemo(() => {
    const pendingLeaves = store.leaves.filter((item) => item.status === "Pending").length
    const presentCount = store.attendance.filter((item) => item.status === "Present").length
    const pendingChecklist = store.checklist.filter((item) => item.status === "Pending").length
    const totalBriefings = store.briefing.length
    return [
      { label: "Pending Leaves", value: pendingLeaves },
      { label: "Present Attendance", value: presentCount },
      { label: "Pending Checklist", value: pendingChecklist },
      { label: "Briefing Items", value: totalBriefings },
    ]
  }, [store])

  const currentForm = forms[activeConfig.key]
  const activeStaff = useMemo(
    () => store.staff.filter((item) => item.activeStatus !== "Inactive"),
    [store.staff],
  )
  const rosterDates = useMemo(
    () => getDateRangeInclusive(monthViewRange.fromDate, monthViewRange.toDate).map((d) => toYmd(d)),
    [monthViewRange.fromDate, monthViewRange.toDate],
  )

  const selectedRosterStaff = useMemo(() => {
    if (rosterSettings.staffId === "ALL") return activeStaff
    return activeStaff.filter((s) => getStaffSelectionToken(s) === rosterSettings.staffId)
  }, [activeStaff, rosterSettings.staffId])

  useEffect(() => {
    if (rosterSettings.staffId === "ALL") return
    const exists = activeStaff.some(
      (staff) => getStaffSelectionToken(staff) === rosterSettings.staffId,
    )
    if (!exists) {
      setRosterSettings((prev) => ({
        ...prev,
        staffId: "ALL",
      }))
    }
  }, [activeStaff, rosterSettings.staffId])

  const shiftsByCode = useMemo(() => {
    const map = new Map<string, Entry>()
    store.shift.forEach((s) => {
      if (s.shiftCode?.trim()) map.set(s.shiftCode.trim(), s)
    })
    return map
  }, [store.shift])

  const rosterShiftColorsByCode = useMemo(() => {
    const timeKeyToColor = new Map<string, (typeof SHIFT_COLOR_PALETTE)[number]>()
    const codeToColor = new Map<string, (typeof SHIFT_COLOR_PALETTE)[number]>()
    let colorIndex = 0

    store.shift.forEach((shift) => {
      const code = (shift.shiftCode || "").trim()
      if (!code) return
      const timeKey = `${shift.startTime || ""}-${shift.endTime || ""}`
      if (!timeKeyToColor.has(timeKey)) {
        timeKeyToColor.set(timeKey, SHIFT_COLOR_PALETTE[colorIndex % SHIFT_COLOR_PALETTE.length])
        colorIndex += 1
      }
      codeToColor.set(code, timeKeyToColor.get(timeKey)!)
    })

    return codeToColor
  }, [store.shift])

  const rosterEligibleShiftCodes = useMemo(() => {
    if (rosterSettings.staffId === "ALL") {
      return Array.from(
        new Set(store.shift.map((s) => s.shiftCode?.trim()).filter((x): x is string => Boolean(x))),
      )
    }
    const staff = activeStaff.find((s) => getStaffSelectionToken(s) === rosterSettings.staffId)
    if (!staff) return []
    return store.shift
      .filter((s) => levelIncludedInAssigned(s.assignedLevels || "", staff.level || ""))
      .map((s) => s.shiftCode?.trim())
      .filter((x): x is string => Boolean(x))
  }, [store.shift, activeStaff, rosterSettings.staffId])
  const staffLookupResults = useMemo(() => {
    if (activeConfig.key !== "leaves") return []
    const q = staffLookupQuery.trim().toLowerCase()
    const source = store.staff.filter((item) => item.activeStatus !== "Inactive")
    if (!q) return source
    return source
      .filter((item) => {
        const no = (item.staffNo ?? "").toLowerCase()
        const name = (item.fullName ?? "").toLowerCase()
        return no.includes(q) || name.includes(q)
      })
  }, [activeConfig.key, staffLookupQuery, store.staff])

  const govHolidayDateKeys = useMemo(() => {
    const set = new Set<string>()
    store.publicHolidays.forEach((holiday) => {
      const range = getDateRangeInclusive(holiday.holidayStartDate, holiday.holidayEndDate)
      range.forEach((day) => set.add(toYmd(day)))
    })
    return set
  }, [store.publicHolidays])

  const leaveCodeByStaffAndDate = useMemo(() => {
    const map = new Map<string, string>()
    const policyMetaByName = new Map<
      string,
      { typeCode: string; nonCountWeekdays: Set<string> }
    >()
    store.leaveAttendanceControl.forEach((policy) => {
      const name = (policy.leaveAttendanceName || "").trim()
      if (!name) return
      policyMetaByName.set(name, {
        typeCode: (policy.typeCode || "").trim() || "LV",
        nonCountWeekdays: new Set(splitMultiValue(policy.nonCountWeekdays || "")),
      })
    })
    const weekdayMap = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

    store.leaves
      .filter((leave) => (leave.status || "") !== "Rejected")
      .forEach((leave) => {
        const staffNo = (leave.staffNo || "").trim()
        if (!staffNo || !leave.fromDate || !leave.toDate) return
        const policyMeta = policyMetaByName.get((leave.leavePolicyName || "").trim()) || {
          typeCode: "LV",
          nonCountWeekdays: new Set<string>(),
        }
        const leaveCode = policyMeta.typeCode
        const days = getDateRangeInclusive(leave.fromDate, leave.toDate)
        days.forEach((day) => {
          const ymd = toYmd(day)
          const weekdayName = weekdayMap[day.getDay()]
          // Key by staff number + date so leave markers remain stable even
          // if display names change later.
          const key = `${staffNo}::${ymd}`
          if (govHolidayDateKeys.has(ymd)) {
            map.set(key, "GH")
            return
          }
          if (policyMeta.nonCountWeekdays.has(weekdayName)) {
            map.set(key, "PH")
            return
          }
          map.set(key, leaveCode)
        })
      })
    return map
  }, [store.leaves, store.leaveAttendanceControl, govHolidayDateKeys])

  const rosterMatrixRows = useMemo(() => {
    const rows = activeStaff.map((staff) => {
      const cells: Record<string, string> = {}
      rosterDates.forEach((date) => {
        const hit = store.roster.find(
          (r) => r.staffName === staff.fullName && r.staffNo === staff.staffNo && r.date === date,
        )
        if (hit) {
          cells[date] = hit.shiftCode || "ASSIGNED"
          return
        }
        const leaveKey = `${(staff.staffNo || "").trim()}::${date}`
        const leaveCode = leaveCodeByStaffAndDate.get(leaveKey)
        if (leaveCode) {
          cells[date] = leaveCode
          return
        }
        cells[date] = govHolidayDateKeys.has(date) ? "PH" : "-"
      })
      return {
        staffId: getStaffSelectionToken(staff),
        staffNo: staff.staffNo || "",
        staffName: staff.fullName || "",
        level: staff.level || "-",
        cells,
      }
    })
    return rows
  }, [activeStaff, rosterDates, store.roster, leaveCodeByStaffAndDate, govHolidayDateKeys])

  const pendingDutyMarkLockKeys = useMemo(() => {
    const locked = new Set<string>()
    const isTerminal = (value: string) => {
      const v = normalizeText(value || "")
      return v === "approved" || v === "rejected" || v === "cancelled"
    }
    store.attendance.forEach((entry) => {
      const source = normalizeText(entry.requestSource || "")
      if (source !== "staffportal") return
      const workflow = String(entry.workflowStatus || entry.status || "")
      if (isTerminal(workflow)) return
      const staffNo = (entry.staffNo || "").trim()
      const fromDate = (entry.fromDate || "").trim()
      const toDate = (entry.toDate || "").trim() || fromDate
      if (!staffNo || !fromDate || !toDate) return
      getDateRangeInclusive(fromDate, toDate).forEach((day) => {
        locked.add(`${staffNo}::${toYmd(day)}`)
      })
    })
    return locked
  }, [store.attendance])

  const rosterVisibleRows = useMemo(() => {
    const q = rosterView.staffSearch.trim().toLowerCase()
    let rows = rosterMatrixRows
    if (q) {
      rows = rows.filter(
        (row) =>
          (row.staffNo || "").toLowerCase().includes(q) || (row.staffName || "").toLowerCase().includes(q),
      )
    }

    if (rosterView.shiftFilter !== "ALL") {
      rows = rows.filter((row) =>
        rosterDates.some((date) => (row.cells[date] || "") === rosterView.shiftFilter),
      )
    }

    const sorted = [...rows]
    if (rosterView.sortBy === "staffNoAsc") {
      sorted.sort((a, b) => (a.staffNo || "").localeCompare(b.staffNo || ""))
    } else if (rosterView.sortBy === "staffNoDesc") {
      sorted.sort((a, b) => (b.staffNo || "").localeCompare(a.staffNo || ""))
    } else if (rosterView.sortBy === "nameAsc") {
      sorted.sort((a, b) => (a.staffName || "").localeCompare(b.staffName || ""))
    } else if (rosterView.sortBy === "nameDesc") {
      sorted.sort((a, b) => (b.staffName || "").localeCompare(a.staffName || ""))
    }
    return sorted
  }, [rosterMatrixRows, rosterView, rosterDates])

  const rosterShiftFilterOptions = useMemo(() => {
    const set = new Set<string>()
    rosterMatrixRows.forEach((row) => {
      rosterDates.forEach((date) => {
        const code = row.cells[date] || ""
        if (code && code !== "-") set.add(code)
      })
    })
    return Array.from(set).sort()
  }, [rosterMatrixRows, rosterDates])

  const rosterComputeLogsInRange = useMemo(() => {
    const from = (rosterComputeRange.fromDate || "").trim()
    const to = (rosterComputeRange.toDate || "").trim()
    if (!from || !to || from > to) return rosterComputeLogs
    return rosterComputeLogs.filter((log) => {
      const d = (log.date || "").trim()
      if (!d || d === "-") return true
      return d >= from && d <= to
    })
  }, [rosterComputeLogs, rosterComputeRange])

  useEffect(() => {
    if (activeModule !== "roster") return
    if (rosterDates.length === 0) return
    const today = toYmd(new Date())
    if (!rosterDates.includes(today)) return
    const host = rosterGridWrapRef.current
    if (!host) return

    const centerToday = () => {
      const scroller = host.querySelector(".MuiDataGrid-virtualScroller") as HTMLElement | null
      const headerCell = host.querySelector(
        `.MuiDataGrid-columnHeader[data-field="${today}"]`,
      ) as HTMLElement | null
      if (!scroller || !headerCell) return
      const targetLeft =
        headerCell.offsetLeft - scroller.clientWidth / 2 + headerCell.clientWidth / 2
      scroller.scrollLeft = Math.max(0, targetLeft)
    }

    const id = window.setTimeout(centerToday, 60)
    return () => window.clearTimeout(id)
  }, [activeModule, rosterDates, rosterVisibleRows.length, rosterDirectEditMode])

  // Per-day per-code count of staff for the distribution table below the
  // roster grid. Both REST and OFF are merged under the displayed "OF" code.
  // Order: shift codes by their `rosterPriority` value (defined on the Shift
  // module), with alphabetical fallback. AL, GH, PH, OF stay at the tail.
  const rosterDistribution = useMemo(() => {
    const codeSet = new Set<string>()
    const counts = new Map<string, Map<string, number>>() // code -> date -> count
    rosterMatrixRows.forEach((row) => {
      rosterDates.forEach((date) => {
        const raw = (row.cells[date] || "").trim()
        if (!raw || raw === "-" || raw === "UNSET") return
        const code = raw === "REST" || raw === "OFF" ? "OF" : raw
        codeSet.add(code)
        if (!counts.has(code)) counts.set(code, new Map())
        const dateMap = counts.get(code)!
        dateMap.set(date, (dateMap.get(date) || 0) + 1)
      })
    })

    // Build code -> priority lookup from the Shift module config
    const priorityByCode = new Map<string, number>()
    store.shift.forEach((s) => {
      const code = (s.shiftCode || "").trim()
      if (!code) return
      const raw = (s.rosterPriority || "").toString().trim()
      const match = raw.match(/^\d+/)
      const n = Number(match ? match[0] : raw)
      priorityByCode.set(code, Number.isFinite(n) ? n : 9999)
    })

    const tail = ["AL", "GH", "PH", "OF"]
    const heads = Array.from(codeSet)
      .filter((c) => !tail.includes(c))
      .sort((a, b) => {
        const pa = priorityByCode.has(a) ? (priorityByCode.get(a) as number) : 9999
        const pb = priorityByCode.has(b) ? (priorityByCode.get(b) as number) : 9999
        if (pa !== pb) return pa - pb
        return a.localeCompare(b)
      })
    const ordered = [...heads, ...tail.filter((c) => codeSet.has(c))]
    return { codes: ordered, counts }
  }, [rosterMatrixRows, rosterDates, store.shift])

  const persistRosterAuditRun = async (runId: string, logs: RosterComputeLog[]) => {
    if (!supabase || disableRemoteSyncRef.current || logs.length === 0) return
    const scopeStaff =
      rosterSettings.staffId === "ALL"
        ? "ALL"
        : selectedRosterStaff
            .map((s) => `${s.staffNo || ""} - ${s.fullName || ""}`.trim())
            .filter(Boolean)
            .join(", ")

    const rows = logs.map((log) => ({
      run_id: runId,
      log_id: log.id,
      logged_at: log.at,
      date_key: log.date,
      level_name: log.level,
      staff_ref: log.staff,
      code: log.code,
      severity: log.severity,
      message: log.message,
      roster_from_date: rosterSettings.fromDate,
      roster_to_date: rosterSettings.toDate,
      staff_scope: scopeStaff || "-",
      staff_count: selectedRosterStaff.length,
    }))

    const { error } = await supabase.from(SUPABASE_ROSTER_AUDIT_TABLE).insert(rows)
    if (error) {
      // Keep UI functional even when audit table is not yet created.
      console.error("Failed to persist roster audit logs:", error.message)
    }
  }

  const generateRoster = () => {
    const runId = `${Date.now()}-${Math.random().toString(16).slice(2)}`
    if (!rosterSettings.fromDate || !rosterSettings.toDate) {
      window.alert("Please set roster date range.")
      return
    }
    if (rosterSettings.fromDate > rosterSettings.toDate) {
      window.alert('"To Date" must be after or equal to "From Date".')
      return
    }
    if (selectedRosterStaff.length === 0) {
      window.alert("No active staff available for generation.")
      return
    }

    const fixedOffSet = new Set(splitMultiValue(rosterSettings.fixedOffWeekdays))
    const weekdayMap = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    // Lookup: shift code → shiftType ("Early" | "Mid" | "Late" | "")
    // Used by Late-after-REST and Early-before-REST sequencing rules.
    const shiftTypeByCode = new Map<string, string>()
    store.shift.forEach((s) => {
      const code = (s.shiftCode || "").trim()
      if (!code) return
      shiftTypeByCode.set(code, (s.shiftType || "").trim())
    })
    const todayYmd = toYmd(new Date())
    if (rosterSettings.fromDate < todayYmd || rosterSettings.toDate < todayYmd) {
      window.alert("Roster cannot be generated for past dates.")
      return
    }
    const dates = getDateRangeInclusive(rosterSettings.fromDate, rosterSettings.toDate)
    const leaveBlocks = store.leaves.filter((l) => (l.status || "") !== "Rejected")
    const maxHours = Number(rosterSettings.maxShiftHours || "49")
    const maxHoursCap = Math.max(Number(maxHours || 0), 1)
    const restHours = Math.max(Number(rosterSettings.rotationalRestHours || "24"), 0)
    const selectedCode = rosterSettings.shiftCode.trim()

    const removeStaffNames = new Set(selectedRosterStaff.map((s) => s.fullName))
    const removeStaffNos = new Set(selectedRosterStaff.map((s) => s.staffNo))

    const kept = store.roster.filter(
      (r) =>
        !(
          removeStaffNames.has(r.staffName || "") &&
          removeStaffNos.has(r.staffNo || "") &&
          r.date >= rosterSettings.fromDate &&
          r.date <= rosterSettings.toDate
        ),
    )

    const generated: Entry[] = []
    const computeLogs: RosterComputeLog[] = []
    const pushComputeLog = ({
      date,
      level = "-",
      staff = "-",
      code = "-",
      severity = "info",
      message,
    }: {
      date: string
      level?: string
      staff?: string
      code?: string
      severity?: "info" | "warn" | "error"
      message: string
    }) => {
      computeLogs.push({
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        at: new Date().toISOString(),
        date,
        level,
        staff,
        code,
        severity,
        message,
      })
    }
    const unfilledPriorityReport: { date: string; level: string; code: string; reason: string }[] = []
    const normalizeLevel = (value: string) => (value || "UNLEVELLED").trim().toLowerCase()
    const levelDayAssignments = new Map<
      string,
      {
        usedCodes: Set<string>
        usedTimeKeys: Set<string>
        restCount: number
      }
    >()
    const dayAssignments = new Map<
      string,
      {
        usedCodes: Set<string>
        usedTimeKeys: Set<string>
        restCount: number
      }
    >()
    const parsePriority = (value: string) => {
      const match = (value || "").trim().match(/^\d+/)
      const n = Number(match ? match[0] : value)
      return Number.isNaN(n) ? 99 : n
    }
    const staffLevelByNo = new Map<string, string>()
    store.staff.forEach((s) => {
      const no = (s.staffNo || "").trim()
      if (no) staffLevelByNo.set(no, normalizeLevel(s.level || ""))
    })
    const levelPriorityRules = new Map<string, Map<number, number[]>>()
    store.levelShiftPriority.forEach((rule) => {
      const levelKey = normalizeLevel(rule.levelName || "")
      if (!levelKey) return
      const priority = Number((rule.priorityOrder || "").trim())
      if (!Number.isFinite(priority)) return
      const rosterPriorities = splitMultiValue(rule.rosterPriorityNumbers || "")
        .map((n) => parseLeadingNumber(n))
        .filter((n): n is number => n !== null)
        .filter((n) => Number.isFinite(n))
      if (rosterPriorities.length === 0) return
      if (!levelPriorityRules.has(levelKey)) {
        levelPriorityRules.set(levelKey, new Map<number, number[]>())
      }
      const byPriority = levelPriorityRules.get(levelKey)!
      if (!byPriority.has(priority)) byPriority.set(priority, [])
      const merged = [...(byPriority.get(priority) || []), ...rosterPriorities]
      byPriority.set(
        priority,
        Array.from(new Set(merged)).sort((a, b) => a - b),
      )
    })
    const priorityTypeRank: Record<string, number> = {
      mandatory: 1,
      required: 2,
      rest: 3,
      off: 3,
      "rest/off": 3,
      "rest off": 3,
      optional: 4,
      admin: 5,
    }
    const shiftTypeRulesByCode = new Map<
      string,
      Array<{
        typeName: string
        rank: number
        requiredDaily: boolean
        weekdays: Set<string>
      }>
    >()
    store.rosterPriorityType.forEach((rule) => {
      const typeName = normalizePriorityTypeName(rule.priorityType || "")
      const rank = priorityTypeRank[typeName] ?? 99
      const requiredDaily = normalizeText(rule.requiredDaily || "") === "yes"
      const weekdays = new Set(splitMultiValue(rule.requiredWeekdays || ""))
      splitMultiValue(rule.shiftCodes || "")
        .map((code) => code.trim())
        .filter(Boolean)
        .forEach((code) => {
          if (!shiftTypeRulesByCode.has(code)) shiftTypeRulesByCode.set(code, [])
          shiftTypeRulesByCode.get(code)!.push({ typeName, rank, requiredDaily, weekdays })
        })
    })
    const codeTypeRankByCode = new Map<string, number>()
    shiftTypeRulesByCode.forEach((rules, code) => {
      const bestRank = rules.reduce((min, rule) => Math.min(min, rule.rank), 99)
      codeTypeRankByCode.set(code, bestRank)
    })

    // Tracks the most recent REST date per staff. Declared here (before the
    // `kept.forEach` pre-pass) so that pre-existing REST rows can populate
    // it. Used by the TP3 REST/OFF planning step to prefer the staff who
    // rested least recently.
    const lastRestDateByStaff = new Map<string, string>()

    kept.forEach((row) => {
      const rowDate = row.date || ""
      if (rowDate < rosterSettings.fromDate || rowDate > rosterSettings.toDate) return
      if (!dayAssignments.has(rowDate)) {
        dayAssignments.set(rowDate, {
          usedCodes: new Set<string>(),
          usedTimeKeys: new Set<string>(),
          restCount: 0,
        })
      }
      const levelKey = staffLevelByNo.get((row.staffNo || "").trim()) || normalizeLevel("")
      const levelDayKey = `${rowDate}::${levelKey}`
      if (!levelDayAssignments.has(levelDayKey)) {
        levelDayAssignments.set(levelDayKey, {
          usedCodes: new Set<string>(),
          usedTimeKeys: new Set<string>(),
          restCount: 0,
        })
      }
      const state = levelDayAssignments.get(levelDayKey)!
      const dayState = dayAssignments.get(rowDate)!
      const code = (row.shiftCode || "").trim()
      if (isOffCode(code)) {
        state.restCount += 1
        dayState.restCount += 1
        const staffToken = `${(row.staffNo || "").trim()}::${(row.staffName || "").trim()}`
        const prev = lastRestDateByStaff.get(staffToken)
        if (!prev || rowDate > prev) lastRestDateByStaff.set(staffToken, rowDate)
        return
      }
      if (!code || isOffCode(code) || code === "UNSET") return
      state.usedCodes.add(code)
      dayState.usedCodes.add(code)
      const meta = shiftsByCode.get(code)
      if (meta) {
        const timeKey = `${meta.startTime || ""}-${meta.endTime || ""}`
        state.usedTimeKeys.add(timeKey)
        dayState.usedTimeKeys.add(timeKey)
      }
    })
    // Rolling-window settings:
    //   maxHoursCap (per cycle, default 35) is enforced over a sliding window
    //   of the last (lookbackDays + 1) calendar days — i.e. the last
    //   lookbackDays days BEFORE the day under consideration plus today.
    //   With lookbackDays = 4 and a 35-hour cap, this enforces the
    //   "5 consecutive days, no more than 35 hours total" rule.
    const lookbackDays = 4

    const subtractDaysYmd = (ymd: string, n: number) => {
      const d = new Date(`${ymd}T00:00:00`)
      d.setDate(d.getDate() - n)
      return toYmd(d)
    }
    const addDaysYmd = (ymd: string, n: number) => {
      const d = new Date(`${ymd}T00:00:00`)
      d.setDate(d.getDate() + n)
      return toYmd(d)
    }

    const shiftHoursForCode = (code: string) => {
      const meta = shiftsByCode.get(code)
      if (!meta) return 0
      const fromTimes = calculateShiftDurationHoursFromTimes(
        meta.startTime || "",
        meta.endTime || "",
      )
      if (fromTimes && fromTimes > 0) return fromTimes
      return parseShiftDurationToHours(meta.shiftDuration || "8h 0m") || 0
    }

    // Synthetic pre-roster history: rows that DIDN'T actually happen but
    // are added to make different staff start the roster at different
    // cycle phases. Without this, with no real history every staff begins
    // at 0h and they all hit the 35h/5-day cap on the same day, causing
    // the entire team to rest in lockstep. By giving staff[i] (i mod
    // cycleDays) fake past shift days, their rolling-window hours-before
    // fromDate is staggered. The fake entries naturally age out of the
    // window after `lookbackDays` real days.
    // Rolling sum of hours worked by `staff` in the lookbackDays calendar days
    // immediately before `ymd`. Sources, in priority order, are:
    //   1) rows already produced in THIS run (`generated`)
    //   2) preserved rows outside the regenerated range (`kept`)
    //   3) synthetic stagger history (only when no real row exists)
    // REST / OFF / UNSET / blank shift codes contribute zero, which is what
    // lets a previous REST naturally drop out of the window.
    const computeRollingHoursBefore = (staff: Entry, ymd: string) => {
      const sno = (staff.staffNo || "").trim()
      const sname = (staff.fullName || "").trim()
      let total = 0
      for (let i = 1; i <= lookbackDays; i++) {
        const d = subtractDaysYmd(ymd, i)
        const match =
          generated.find(
            (r) =>
              (r.staffNo || "").trim() === sno &&
              (r.staffName || "").trim() === sname &&
              r.date === d,
          ) ||
          kept.find(
            (r) =>
              (r.staffNo || "").trim() === sno &&
              (r.staffName || "").trim() === sname &&
              r.date === d,
          )
        if (match) {
          const code = (match.shiftCode || "").trim()
          if (!code || isOffCode(code) || code === "UNSET") continue
          total += shiftHoursForCode(code)
          continue
        }
        // No historical roster row for this date.
      }
      return total
    }
    const computeConsecutiveWorkDaysBefore = (staff: Entry, ymd: string) => {
      const sno = (staff.staffNo || "").trim()
      const sname = (staff.fullName || "").trim()
      let streak = 0
      for (let i = 1; i <= 62; i++) {
        const d = subtractDaysYmd(ymd, i)
        const match =
          generated.find(
            (r) =>
              (r.staffNo || "").trim() === sno &&
              (r.staffName || "").trim() === sname &&
              r.date === d,
          ) ||
          kept.find(
            (r) =>
              (r.staffNo || "").trim() === sno &&
              (r.staffName || "").trim() === sname &&
              r.date === d,
          )
        if (!match) break
        const code = (match.shiftCode || "").trim()
        if (!code || isOffCode(code) || code === "UNSET") break
        streak += 1
      }
      return streak
    }

    const staffState = new Map<
      string,
      {
        allocatedHours: number
        consecutiveWorkDays: number
        forcedRestUntilDate: string | null
        restUntil: Date | null
        lastShiftCode: string
        // Last few shift codes (most recent first) — used to penalize staff
        // who would receive the same code multiple days in a row. Window of
        // RECENT_CODE_WINDOW entries; REST/OFF/leave codes are also tracked.
        recentShiftCodes: string[]
        // shiftType (Early/Mid/Late) of the most recent worked shift. Used
        // to alternate types between OFs (Early today → Late tomorrow).
        lastShiftType: string
        // End-time of the most recent worked shift (REST/OFF excluded). Used
        // to anchor the cooldown rule: next shift's start must be at least
        // `restHours` hours after this timestamp.
        lastShiftEndTime: Date | null
      }
    >()
    // How many recent shift codes to keep per staff for the no-repeat sort.
    const RECENT_CODE_WINDOW = 3
    // Average shift hours: best-effort from configured shifts. Falls back
    // to maxHours/cycleDays if no shift duration is parseable.
    const cycleDays = lookbackDays + 1
    const sampleShiftHours = (() => {
      const durations = store.shift
        .map((s) => shiftHoursForCode((s.shiftCode || "").trim()))
        .filter((h) => h > 0)
      if (durations.length === 0) return Math.max(1, Math.floor(maxHours / cycleDays))
      return durations.reduce((a, b) => a + b, 0) / durations.length
    })()
    const maxConsecutiveWorkDays = Math.max(
      1,
      Math.floor(Math.max(Number(maxHours || 0), 1) / Math.max(sampleShiftHours, 1)),
    )
    // No synthetic pre-roster offset: REST decisions are now based only on
    // real roster history and current generation window.
    selectedRosterStaff.forEach((staff) => {
      // computeRollingHoursBefore now factors in synthetic history when
      // no real row is found for a past day, so different staff naturally
      // start at different cycle phases.
      staffState.set(getStaffSelectionToken(staff), {
        allocatedHours: computeRollingHoursBefore(staff, rosterSettings.fromDate),
        consecutiveWorkDays: computeConsecutiveWorkDaysBefore(staff, rosterSettings.fromDate),
        forcedRestUntilDate: null,
        restUntil: null,
        lastShiftCode: "",
        recentShiftCodes: [],
        lastShiftType: "",
        lastShiftEndTime: null,
      })
    })

    const getLevelDayState = (ymd: string, level: string) => {
      const levelDayKey = `${ymd}::${normalizeLevel(level || "")}`
      if (!levelDayAssignments.has(levelDayKey)) {
        levelDayAssignments.set(levelDayKey, {
          usedCodes: new Set<string>(),
          usedTimeKeys: new Set<string>(),
          restCount: 0,
        })
      }
      return levelDayAssignments.get(levelDayKey)!
    }

    const getDayState = (ymd: string) => {
      if (!dayAssignments.has(ymd)) {
        dayAssignments.set(ymd, {
          usedCodes: new Set<string>(),
          usedTimeKeys: new Set<string>(),
          restCount: 0,
        })
      }
      return dayAssignments.get(ymd)!
    }

    const addRow = (staff: Entry, ymd: string, shiftCode: string, location: string) => {
      generated.push({
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        createdAt: new Date().toISOString(),
        date: ymd,
        staffNo: staff.staffNo,
        staffName: staff.fullName,
        role: staff.designation || "",
        shiftCode,
        location,
      })
    }

    if (rosterSettings.offType === "Fixed") {
      selectedRosterStaff.forEach((staff) => {
        const eligible = store.shift
          .filter((s) => levelIncludedInAssigned(s.assignedLevels || "", staff.level || ""))
          .sort((a, b) => parsePriority(a.rosterPriority || "") - parsePriority(b.rosterPriority || ""))
        dates.forEach((day) => {
          const ymd = toYmd(day)
          if (fixedOffSet.has(weekdayMap[day.getDay()])) {
            addRow(staff, ymd, "OF", "Fixed Off")
            return
          }
          const fallback = selectedCode ? shiftsByCode.get(selectedCode) : eligible[0]
          addRow(staff, ymd, fallback?.shiftCode || selectedCode || "UNSET", "Auto Generated")
        })
      })
    } else {
      const dayStart = (ymd: string) => new Date(`${ymd}T00:00:00`)
      // Cooldown rule: after a worked shift ends, the staff must rest for
      // `restHours` (configured rotational rest hours, e.g. 40h). The next
      // shift may start no earlier than `lastShiftEnd + restHours`. This is
      // computed once per state-change from the staff's most recent shift
      // end time. If the staff has no recorded shift end (start of roster
      // with no synthetic history), no cooldown is set.
      const MALDIVES_OFFSET_MS = 5 * 60 * 60 * 1000
      const cooldownFromShiftEnd = (lastEnd: Date | null): Date | null => {
        if (!lastEnd) return null
        return new Date(lastEnd.getTime() + restHours * 60 * 60 * 1000)
      }
      const buildShiftOptions = (staff: Entry, ymd: string) =>
        store.shift
          .filter((s) => levelIncludedInAssigned(s.assignedLevels || "", staff.level || ""))
          .map((shift) => {
            const code = (shift.shiftCode || "").trim()
            if (!code) return null
            const startTime = shift.startTime || "00:00"
            const endTime = shift.endTime || "00:00"
            return {
              code,
              priority: parsePriority(shift.rosterPriority || ""),
              timeKey: `${startTime}-${endTime}`,
              startDateTime: combineDateAndTime(ymd, startTime),
              endDateTime: getShiftEndDateTime(ymd, startTime, endTime),
              durationHours:
                calculateShiftDurationHoursFromTimes(startTime, endTime) ||
                parseShiftDurationToHours(shift.shiftDuration || "8h 0m"),
            }
          })
          .filter((x): x is NonNullable<typeof x> => Boolean(x))

      dates.forEach((day) => {
        const ymd = toYmd(day)
        const weekdayName = weekdayMap[day.getDay()]
        const dayState = getDayState(ymd)
        const assignedToday = new Set<string>()
        // Tracks codes already reported as unfilled today so the same
        // shift code isn't recorded once per level iteration.
        const reportedUnfilled = new Set<string>()

        // Refresh each staff's allocatedHours from the rolling 5-day window
        // (4 prior days + today). The synthetic stagger history is folded
        // in by computeRollingHoursBefore for past days that have no real
        // row yet, so it transparently desyncs the team in early days and
        // ages out naturally as real days fill the window.
        selectedRosterStaff.forEach((staff) => {
          const token = getStaffSelectionToken(staff)
          const state = staffState.get(token)
          if (!state) return
          state.allocatedHours = computeRollingHoursBefore(staff, ymd)
          state.consecutiveWorkDays = computeConsecutiveWorkDaysBefore(staff, ymd)
          if (state.forcedRestUntilDate && state.forcedRestUntilDate < ymd) {
            state.forcedRestUntilDate = null
          }
        })
        const contexts = selectedRosterStaff.map((staff) => {
          const token = getStaffSelectionToken(staff)
          const state = staffState.get(token)!
          const startOfDay = dayStart(ymd)
          if (state.restUntil && state.restUntil <= startOfDay) {
            state.restUntil = null
          }
          const hasFutureLeave = leaveBlocks.some(
            (l) =>
              l.staffNo === staff.staffNo &&
              l.fromDate <= ymd &&
              l.toDate >= ymd &&
              ymd >= todayYmd,
          )
          const options = buildShiftOptions(staff, ymd)
          const remainingHours = maxHoursCap - state.allocatedHours
          // Two filter layers:
          //  1) restRespectingOptions: shifts that start after active cooldown.
          //     Top-priority order uses this.
          //  2) availableOptions: restRespectingOptions AND would NOT push
          //     the staff past max-hours cap. Lower priorities use this.
          const restRespectingOptions = state.restUntil
            ? options.filter((o) => o.startDateTime >= state.restUntil!)
            : options
          const availableOptions = restRespectingOptions.filter(
            (o) => o.durationHours <= remainingHours,
          )
          const mustRestFromCooldown =
            Boolean(state.restUntil) && restRespectingOptions.length === 0
          const mustRestFromCap = state.allocatedHours >= maxHoursCap
          const mustRestFromConsecutive = state.consecutiveWorkDays >= maxConsecutiveWorkDays
          const mustRestFromForced =
            Boolean(state.forcedRestUntilDate) && ymd <= (state.forcedRestUntilDate || "")
          const mustRest =
            mustRestFromCap || mustRestFromCooldown || mustRestFromConsecutive || mustRestFromForced
          return {
            staff,
            token,
            state,
            options,
            restRespectingOptions,
            availableOptions,
            hasFutureLeave,
            mustRest,
            mustRestFromCap,
            mustRestFromCooldown,
            mustRestFromConsecutive,
            mustRestFromForced,
          }
        })

        // === TP3 REST/OFF planning pass ===
        // Per spec: TP3 (Rest/Off) is a priority type evaluated BEFORE TP4
        // Optional and TP5 Admin. We choose the most-rest-needing candidates
        // (consecutiveWorkDays desc, allocatedHours desc, last-rest-date asc)
        // and mark them as planned-rest tokens. Those tokens are excluded from
        // the Optional/Admin pass and picked up by the final REST pass.
        //
        // Two refinements per user feedback "different staff should have
        // different OFF days":
        //   1. Apply a GLOBAL per-day cap so OFFs don't stack on one date.
        //      Cap = max(1, ceil(availableStaffToday / cycleDays)). With 9
        //      available staff and a 5-day cycle this gives ~2 OFFs/day.
        //   2. Skip any staff who ALREADY rested on the previous day so that
        //      the same person doesn't get back-to-back OFFs while another
        //      staff hasn't rested yet.
        const plannedRestTokens = new Set<string>()
        const levelGroups = new Map<string, (typeof contexts)[number][]>()
        const eligibleContexts = contexts.filter((ctx) => !ctx.hasFutureLeave)
        eligibleContexts.forEach((ctx) => {
          const levelKey = normalizeLevel(ctx.staff.level || "")
          if (!levelGroups.has(levelKey)) levelGroups.set(levelKey, [])
          levelGroups.get(levelKey)!.push(ctx)
        })
        const dailyRestCap = Math.max(
          1,
          Math.ceil(eligibleContexts.length / Math.max(cycleDays, 1)),
        )
        // Compute the previous calendar day's YMD so we can avoid resting the
        // same staff two days in a row when other staff are still candidates.
        const prevYmd = (() => {
          const d = new Date(`${ymd}T00:00:00Z`)
          d.setUTCDate(d.getUTCDate() - 1)
          return d.toISOString().slice(0, 10)
        })()
        // Build one candidate per level (the most-rest-needing), then rank
        // candidates GLOBALLY and take only the top `dailyRestCap`.
        const levelTopCandidates: { ctx: (typeof contexts)[number]; levelKey: string }[] = []
        levelGroups.forEach((group, levelKey) => {
          if (group.length < 2) return
          const candidate = [...group]
            .filter((ctx) => !ctx.mustRestFromForced)
            .filter((ctx) => {
              // Skip staff who rested yesterday so OFFs don't repeat on the
              // same person while others haven't rested. Only enforced when
              // the level has another viable candidate.
              const t = `${(ctx.staff.staffNo || "").trim()}::${(ctx.staff.fullName || "").trim()}`
              const last = lastRestDateByStaff.get(t) || ""
              const otherViable = group.some(
                (g) =>
                  g !== ctx &&
                  !g.mustRestFromForced &&
                  (lastRestDateByStaff.get(
                    `${(g.staff.staffNo || "").trim()}::${(g.staff.fullName || "").trim()}`,
                  ) || "") !== prevYmd,
              )
              return !(last === prevYmd && otherViable)
            })
            .sort((a, b) => {
              if (a.state.consecutiveWorkDays !== b.state.consecutiveWorkDays) {
                return b.state.consecutiveWorkDays - a.state.consecutiveWorkDays
              }
              if (a.state.allocatedHours !== b.state.allocatedHours) {
                return b.state.allocatedHours - a.state.allocatedHours
              }
              const aToken = `${(a.staff.staffNo || "").trim()}::${(a.staff.fullName || "").trim()}`
              const bToken = `${(b.staff.staffNo || "").trim()}::${(b.staff.fullName || "").trim()}`
              const aLastRest = lastRestDateByStaff.get(aToken) || ""
              const bLastRest = lastRestDateByStaff.get(bToken) || ""
              return aLastRest.localeCompare(bLastRest)
            })[0]
          if (!candidate) return
          levelTopCandidates.push({ ctx: candidate, levelKey })
        })
        // Rank the per-level top candidates globally by rest-need so the
        // staff who needs rest most gets the OFF, regardless of level.
        levelTopCandidates
          .sort((a, b) => {
            if (a.ctx.state.consecutiveWorkDays !== b.ctx.state.consecutiveWorkDays) {
              return b.ctx.state.consecutiveWorkDays - a.ctx.state.consecutiveWorkDays
            }
            if (a.ctx.state.allocatedHours !== b.ctx.state.allocatedHours) {
              return b.ctx.state.allocatedHours - a.ctx.state.allocatedHours
            }
            const aT = `${(a.ctx.staff.staffNo || "").trim()}::${(a.ctx.staff.fullName || "").trim()}`
            const bT = `${(b.ctx.staff.staffNo || "").trim()}::${(b.ctx.staff.fullName || "").trim()}`
            return (lastRestDateByStaff.get(aT) || "").localeCompare(
              lastRestDateByStaff.get(bT) || "",
            )
          })
          .slice(0, dailyRestCap)
          .forEach(({ ctx, levelKey }) => {
            plannedRestTokens.add(ctx.token)
            pushComputeLog({
              date: ymd,
              level: ctx.staff.level || levelKey || "(unleveled)",
              staff: `${ctx.staff.staffNo || ""} - ${ctx.staff.fullName || ""}`.trim() || "-",
              code: "REST",
              severity: "info",
              message: `TP3 planned REST/OFF (cap=${dailyRestCap}/day, evaluated before TP4 Optional / TP5 Admin)`,
            })
          })

        const canUseOption = (ctx: (typeof contexts)[number], option: (typeof contexts)[number]["availableOptions"][number]) => {
          // Uniqueness is enforced on shift CODE only (one staff per code per
          // day globally). Time-key collision is intentionally NOT blocked:
          // multiple codes (e.g. A and E for Flight Watch) may share the
          // same hours and both must still be fillable. The per-staff
          // "one shift per day" rule is enforced separately via
          // assignedToday.has(ctx.token).
          const levelState = getLevelDayState(ymd, ctx.staff.level || "")
          if (dayState.usedCodes.has(option.code)) return false
          if (levelState.usedCodes.has(option.code)) return false
          return true
        }

        const assignOption = (
          ctx: (typeof contexts)[number],
          option: (typeof contexts)[number]["availableOptions"][number],
          location = "Auto Generated",
        ) => {
          const levelState = getLevelDayState(ymd, ctx.staff.level || "")
          addRow(ctx.staff, ymd, option.code, location)
          assignedToday.add(ctx.token)
          ctx.state.lastShiftCode = option.code
          ctx.state.lastShiftType = (shiftTypeByCode.get(option.code) || "").toLowerCase()
          ctx.state.recentShiftCodes = [option.code, ...ctx.state.recentShiftCodes].slice(
            0,
            RECENT_CODE_WINDOW,
          )
          // Track the actual end-of-shift moment so all future cooldown
          // calculations key off it (next shift may start no earlier than
          // lastShiftEndTime + restHours).
          ctx.state.lastShiftEndTime = option.endDateTime
          if (ctx.state.restUntil && option.startDateTime >= ctx.state.restUntil) {
            ctx.state.restUntil = null
          }
          ctx.state.allocatedHours += option.durationHours
          // Arm the per-staff cooldown so the next shift can't start within
          // restHours of this shift's end. We do NOT zero allocatedHours here
          // — the rolling window will naturally drop old hours as the loop
          // advances day-by-day.
          if (ctx.state.allocatedHours >= Math.max(Number(maxHours || 0), 1)) {
            ctx.state.restUntil = cooldownFromShiftEnd(ctx.state.lastShiftEndTime)
          }
          dayState.usedCodes.add(option.code)
          dayState.usedTimeKeys.add(option.timeKey)
          levelState.usedCodes.add(option.code)
          levelState.usedTimeKeys.add(option.timeKey)
        }

        // === Per-shift outer loop ===
        // Build a global, priority-sorted queue of shifts. Each shift carries
        // its OWN level-priority chain via `assignedLevels` (the user-defined
        // order in which levels should be considered for that shift). For
        // example, shift A might have assignedLevels = "Assistant, Officer,
        // Supervisor, Executive, Lead" — meaning: try Assistant first, fall
        // through to Officer if Assistant has no candidate, etc.
        type ShiftSpec = {
          code: string
          priority: number
          chain: string[] // level names in declared order
          rawAssignedLevels: string
          typeRank: number
          typeName: string
          requiredToday: boolean
        }
        const shiftQueue: ShiftSpec[] = store.shift
          .map((s) => {
            const code = (s.shiftCode || "").trim()
            const rules = shiftTypeRulesByCode.get(code) || []
            const ranked = rules.sort((a, b) => a.rank - b.rank)[0]
            const requiredToday = rules.some(
              (r) => r.requiredDaily || (!r.requiredDaily && r.weekdays.has(weekdayName)),
            )
            return {
              code,
              priority: parsePriority(s.rosterPriority || ""),
              chain: splitMultiValue(s.assignedLevels || ""),
              rawAssignedLevels: s.assignedLevels || "",
              typeRank: ranked?.rank ?? 99,
              typeName: ranked?.typeName || "untyped",
              requiredToday,
            }
          })
          .filter((s) => s.code && s.chain.length > 0)
          .sort(
            (a, b) =>
              Number(b.requiredToday) - Number(a.requiredToday) ||
              a.typeRank - b.typeRank ||
              a.priority - b.priority ||
              a.code.localeCompare(b.code),
          )

        // Mandatory ↔ Required ↔ Optional split. The user's spec defines
        // priority 1 = mandatory (Flight Watch A,E / Flight Release B,F) and
        // priority 2 = required (D,S). Anything ≥3 is optional. Mandatory
        // can bypass the hour cap (cooldown still respected) and gets a
        // capacity-exhausted unfilled report. Required respects cap.
        const mustPriorityCutoff = 2

        shiftQueue.forEach((shift) => {
          // Globally one staff per shift code per day.
          if (dayState.usedCodes.has(shift.code)) return

          const isMustPriority =
            shift.requiredToday &&
            shift.typeRank <= mustPriorityCutoff &&
            shift.typeRank < (priorityTypeRank.optional ?? 4)
          // Must (priority 1-2) bypasses cap, respects cooldown only.
          // Optional (≥3) respects both cap and cooldown.
          const sourceOptions = (ctx: (typeof contexts)[number]) =>
            isMustPriority ? ctx.options : ctx.availableOptions
          const eligibleForShift = (ctx: (typeof contexts)[number]) => {
            if (ctx.hasFutureLeave) return false
            if (assignedToday.has(ctx.token)) return false
            if (isMustPriority) return !ctx.mustRestFromForced
            return !ctx.mustRest
          }

          // Walk the shift's chain in declared order. The FIRST level with
          // an eligible candidate wins. We do NOT skip to a later level
          // until the earlier one is fully exhausted (no candidates).
          let assigned = false
          for (let chainIdx = 0; chainIdx < shift.chain.length; chainIdx++) {
            const levelRaw = shift.chain[chainIdx]
            const levelKey = normalizeLevel(levelRaw)
            const candidates = contexts
              .filter((ctx) => normalizeLevel(ctx.staff.level || "") === levelKey)
              .filter(eligibleForShift)
              .map((ctx) => ({
                ctx,
                option: sourceOptions(ctx).find((o) => o.code === shift.code),
              }))
              .filter(
                (x): x is { ctx: (typeof contexts)[number]; option: NonNullable<typeof x.option> } =>
                  Boolean(x.option),
              )
              .filter(
                ({ ctx, option }) =>
                  ctx.state.allocatedHours + option.durationHours <= maxHoursCap &&
                  ctx.state.consecutiveWorkDays + 1 <= maxConsecutiveWorkDays,
              )
              .filter(({ ctx, option }) => canUseOption(ctx, option))
              // Sort tiers, applied in order:
              //   1) Hard no-3-in-a-row: deprioritize staff whose last 2
              //      worked codes already match this shift code.
              //   2) shiftType === "Late": prefer staff whose previous day
              //      was REST (Late shift after rest rule).
              //   3) shiftType === "Early": prefer staff whose projected
              //      hours after this shift hit the cap, forcing REST
              //      tomorrow (Early shift before rest rule).
              //   4) Alternate Early/Late between OFs: prefer staff whose
              //      lastShiftType differs from this shift's type.
              //   5) Recent-code repeat count: prefer staff with FEWER
              //      occurrences of this code in their recent window.
              //   6) Distribute load: lowest accumulated hours first.
              .sort((a, b) => {
                const stype = (shiftTypeByCode.get(shift.code) || "").toLowerCase()

                // (1) Tier 1: hard 3-in-a-row block. If a staff has had this
                // code on each of the last 2 worked days, pin them last so a
                // 3rd consecutive same-code is only ever picked when there's
                // no other candidate.
                const sameLast2 = (codes: string[]) =>
                  codes.length >= 2 && codes[0] === shift.code && codes[1] === shift.code
                const aTriple = sameLast2(a.ctx.state.recentShiftCodes) ? 1 : 0
                const bTriple = sameLast2(b.ctx.state.recentShiftCodes) ? 1 : 0
                if (aTriple !== bTriple) return aTriple - bTriple

                if (stype === "late") {
                  const aRested = a.ctx.state.lastShiftCode === "REST" ? 0 : 1
                  const bRested = b.ctx.state.lastShiftCode === "REST" ? 0 : 1
                  if (aRested !== bRested) return aRested - bRested
                } else if (stype === "early") {
                  const aProjectsRest =
                    a.ctx.state.allocatedHours + a.option.durationHours >= maxHoursCap ? 0 : 1
                  const bProjectsRest =
                    b.ctx.state.allocatedHours + b.option.durationHours >= maxHoursCap ? 0 : 1
                  if (aProjectsRest !== bProjectsRest) return aProjectsRest - bProjectsRest
                }

                // (4) Alternate Early/Late between OFs: if today is Early,
                // prefer staff whose last worked type was Late, and vice
                // versa. Skip when stype is "" or the staff has no lastType.
                if (stype === "early" || stype === "late") {
                  const opposite = stype === "early" ? "late" : "early"
                  const aMatches = a.ctx.state.lastShiftType === opposite ? 0 : 1
                  const bMatches = b.ctx.state.lastShiftType === opposite ? 0 : 1
                  if (aMatches !== bMatches) return aMatches - bMatches
                }

                // (5) Recent-code repetition: count this code's occurrences
                // in the last RECENT_CODE_WINDOW worked days. Lower is better.
                const aRepCount = a.ctx.state.recentShiftCodes.filter(
                  (c) => c === shift.code,
                ).length
                const bRepCount = b.ctx.state.recentShiftCodes.filter(
                  (c) => c === shift.code,
                ).length
                if (aRepCount !== bRepCount) return aRepCount - bRepCount

                if (a.ctx.state.consecutiveWorkDays !== b.ctx.state.consecutiveWorkDays) {
                  return a.ctx.state.consecutiveWorkDays - b.ctx.state.consecutiveWorkDays
                }
                return a.ctx.state.allocatedHours - b.ctx.state.allocatedHours
              })

            if (candidates.length > 0) {
              const picked = candidates[0]
              const tag = chainIdx === 0
                ? `Auto RP${shift.priority}/TP${shift.typeRank} ${levelRaw}`
                : `Auto RP${shift.priority}/TP${shift.typeRank} ${shift.chain[0]}→${levelRaw}`
              assignOption(picked.ctx, picked.option, tag)
              pushComputeLog({
                date: ymd,
                level: picked.ctx.staff.level || "(unleveled)",
                staff:
                  `${picked.ctx.staff.staffNo || ""} - ${picked.ctx.staff.fullName || ""}`.trim() || "-",
                code: shift.code,
                severity: "info",
                message: `Assigned ${shift.code} via roster priority chain (${tag}, type=${shift.typeName})`,
              })
              assigned = true
              break
            }
          }

          if (assigned) return

          // === Mandatory rescue pass ===
          // For unfilled MUST priority shifts only, retry the chain with
          // cooldown bypassed but the cap (35h) STILL enforced explicitly.
          // Implements "rest can be moved back and forth, if not exceeding
          // 35hrs": rather than leave a P1/P2 mandatory open while staff
          // sit on a cooldown-induced REST, we let one of them work today
          // if their rolling hours allow it. The cap stays absolute.
          if (isMustPriority) {
            for (let chainIdx = 0; chainIdx < shift.chain.length; chainIdx++) {
              const levelRaw = shift.chain[chainIdx]
              const levelKey = normalizeLevel(levelRaw)
              const rescueCandidates = contexts
                .filter((ctx) => normalizeLevel(ctx.staff.level || "") === levelKey)
                .filter((ctx) => !ctx.hasFutureLeave)
                .filter((ctx) => !assignedToday.has(ctx.token))
                // Use unfiltered options (ignores cooldown gate) but apply
                // an explicit cap check below so totals can't exceed 35h.
                .map((ctx) => ({
                  ctx,
                  option: ctx.options.find((o) => o.code === shift.code),
                }))
                .filter(
                  (x): x is { ctx: (typeof contexts)[number]; option: NonNullable<typeof x.option> } =>
                    Boolean(x.option),
                )
                .filter(
                  ({ ctx, option }) =>
                    ctx.state.allocatedHours + option.durationHours <= maxHoursCap &&
                    ctx.state.consecutiveWorkDays + 1 <= maxConsecutiveWorkDays,
                )
                .filter(({ ctx, option }) => canUseOption(ctx, option))
                .sort((a, b) => {
                  if (a.ctx.state.consecutiveWorkDays !== b.ctx.state.consecutiveWorkDays) {
                    return a.ctx.state.consecutiveWorkDays - b.ctx.state.consecutiveWorkDays
                  }
                  return a.ctx.state.allocatedHours - b.ctx.state.allocatedHours
                })
              if (rescueCandidates.length > 0) {
                const picked = rescueCandidates[0]
                const tag =
                  chainIdx === 0
                    ? `Auto RP${shift.priority}/TP${shift.typeRank} ${levelRaw} (rescue)`
                    : `Auto RP${shift.priority}/TP${shift.typeRank} ${shift.chain[0]}→${levelRaw} (rescue)`
                assignOption(picked.ctx, picked.option, tag)
                // Re-arm the 40h gap from the rescue shift's end so the
                // staff's NEXT shift can't start within restHours of this
                // one, even if assignOption cleared the prior restUntil.
                const refreshed = cooldownFromShiftEnd(picked.ctx.state.lastShiftEndTime)
                if (refreshed) picked.ctx.state.restUntil = refreshed
                pushComputeLog({
                  date: ymd,
                  level: picked.ctx.staff.level || "(unleveled)",
                  staff:
                    `${picked.ctx.staff.staffNo || ""} - ${picked.ctx.staff.fullName || ""}`.trim() || "-",
                  code: shift.code,
                  severity: "warn",
                  message: `Mandatory rescue: cooldown bypass to fill P${shift.priority} ${shift.code} (rolling ${(picked.ctx.state.allocatedHours + picked.option.durationHours).toFixed(1)}h ≤ cap ${maxHoursCap}h)`,
                })
                assigned = true
                break
              }
            }
          }
          if (assigned) return

          // Nothing worked across the entire chain. Record only for must
          // priorities, dedupe per (date,code), and skip when capacity is
          // truly exhausted (all staff blocked).
          if (!isMustPriority) return
          if (reportedUnfilled.has(shift.code)) return
          // Capacity-exhausted check is scoped to staff WITHIN this shift's
          // chain. Staff at other levels can't legally fill this code, so
          // their availability is irrelevant. Example: an "O" shift whose
          // assignedLevels is just Lead has only one possible filler; when
          // that Lead is on leave, the shift is structurally unfillable
          // and we silently skip rather than report a fake gap.
          const chainLevelKeys = new Set(
            shift.chain.map((l) => normalizeLevel(l)),
          )
          const chainContexts = contexts.filter((c) =>
            chainLevelKeys.has(normalizeLevel(c.staff.level || "")),
          )
          // If the chain has zero staff at all (e.g. all chain levels have
          // no selected staff), that's a permanent structural limit.
          // If chain has staff but they're all on leave / cooldown / cap /
          // already assigned, that's also a limit for this date.
          const chainCapacityExhausted =
            chainContexts.length === 0 ||
            chainContexts.every(
              (c) => c.hasFutureLeave || c.mustRest || assignedToday.has(c.token),
            )
          if (chainCapacityExhausted) return
          // Build a diagnostic showing what happened at each chain level so
          // the user can see whether the gap is "no staff at this level" vs
          // "staff existed but were all blocked".
          const chainTrace = shift.chain.map((levelRaw) => {
            const levelKey = normalizeLevel(levelRaw)
            const here = contexts.filter(
              (c) => normalizeLevel(c.staff.level || "") === levelKey,
            )
            if (here.length === 0) return `${levelRaw}=NO-STAFF`
            const eligible = here.filter((c) => eligibleForShift(c))
            if (eligible.length === 0) {
              const blockedReasons = here.map((c) => {
                if (c.hasFutureLeave) return "leave"
                if (assignedToday.has(c.token)) return "assigned"
                if (c.mustRestFromCooldown) return "cooldown"
                if (c.mustRest) return "cap"
                return "?"
              })
              return `${levelRaw}=BLOCKED(${blockedReasons.join(",")})`
            }
            const hasOption = eligible.some((c) =>
              (isMustPriority ? c.options : c.availableOptions).some(
                (o) => o.code === shift.code,
              ),
            )
            return hasOption
              ? `${levelRaw}=NO-CANDIDATE`
              : `${levelRaw}=SHIFT-NOT-ASSIGNED-TO-LEVEL`
          })
          let reason = `chain tried: ${chainTrace.join(" → ")}`
          if (contexts.length === 0) {
            reason = "no active staff selected for this run"
          }
          unfilledPriorityReport.push({
            date: ymd,
            level: shift.chain[0] || "(unleveled)",
            code: shift.code,
            reason,
          })
          pushComputeLog({
            date: ymd,
            level: shift.chain[0] || "(unleveled)",
            code: shift.code,
            severity: "error",
            message: `Unfilled must priority shift: ${reason}`,
          })
          reportedUnfilled.add(shift.code)
        })

        // Hard daily required backfill (before optional/rest):
        // if any required P1/P2 code is still open after chain passes, try
        // cross-level assignment using any unassigned active staff who can
        // take that exact code within hour/day caps.
        const requiredCodesToday = Array.from(
          new Set(
            shiftQueue
              .filter((s) => s.requiredToday && s.typeRank <= mustPriorityCutoff)
              .map((s) => s.code),
          ),
        )
        requiredCodesToday.forEach((code) => {
          if (dayState.usedCodes.has(code)) return
          const rescue = contexts
            .filter((ctx) => !ctx.hasFutureLeave)
            .filter((ctx) => !assignedToday.has(ctx.token))
            .map((ctx) => ({
              ctx,
              option: ctx.options.find((o) => o.code === code),
            }))
            .filter(
              (x): x is { ctx: (typeof contexts)[number]; option: NonNullable<typeof x.option> } =>
                Boolean(x.option),
            )
            .filter(
              ({ ctx, option }) =>
                ctx.state.allocatedHours + option.durationHours <= maxHoursCap &&
                ctx.state.consecutiveWorkDays + 1 <= maxConsecutiveWorkDays,
            )
            .filter(({ ctx, option }) => canUseOption(ctx, option))
            .sort((a, b) => a.ctx.state.allocatedHours - b.ctx.state.allocatedHours)
          if (rescue.length === 0) return
          const picked = rescue[0]
          assignOption(picked.ctx, picked.option, "Required backfill (cross-level)")
          pushComputeLog({
            date: ymd,
            level: picked.ctx.staff.level || "(unleveled)",
            staff: `${picked.ctx.staff.staffNo || ""} - ${picked.ctx.staff.fullName || ""}`.trim() || "-",
            code,
            severity: "warn",
            message: `Required backfill assigned for ${code} before optional/rest`,
          })
        })

        contexts
          .filter(
            (ctx) =>
              !ctx.hasFutureLeave &&
              !ctx.mustRest &&
              !assignedToday.has(ctx.token) &&
              !plannedRestTokens.has(ctx.token),
          )
          .forEach((ctx) => {
            // Per spec: TP4 (Optional) and TP5 (Admin) are LOWER priority than
            // TP3 REST/OFF. This pass only fills staff who are (a) not on
            // leave, (b) not a `mustRest` constraint case, (c) not already
            // booked, and (d) NOT in the planned-rest set from the TP3 pass
            // above. Cap and cooldown remain enforced.
            let options = ctx.availableOptions
              .filter((o) => canUseOption(ctx, o))
              .filter(
                (o) =>
                  ctx.state.allocatedHours + o.durationHours <= maxHoursCap &&
                  ctx.state.consecutiveWorkDays + 1 <= maxConsecutiveWorkDays,
              )
            if (options.length === 0) return
            if (selectedCode) {
              const byCode = options.find((o) => o.code === selectedCode)
              if (byCode) {
                assignOption(ctx, byCode)
                pushComputeLog({
                  date: ymd,
                  level: ctx.staff.level || "(unleveled)",
                  staff: `${ctx.staff.staffNo || ""} - ${ctx.staff.fullName || ""}`.trim() || "-",
                  code: byCode.code,
                  severity: "info",
                  message: `Assigned from manual shift filter (${selectedCode})`,
                })
                return
              }
            }
            // Order options by priority, then preferentially:
            //   - exclude codes already in the recent window (no-repeat)
            //   - prefer opposite shiftType vs lastShiftType (Early↔Late alt)
            options = options.sort((a, b) => a.priority - b.priority || a.startDateTime.getTime() - b.startDateTime.getTime())
            if (options.length > 1) {
              const recent = ctx.state.recentShiftCodes
              const opposite =
                ctx.state.lastShiftType === "early"
                  ? "late"
                  : ctx.state.lastShiftType === "late"
                    ? "early"
                    : ""
              const ranked = [...options].sort((a, b) => {
                // Hard 3-in-a-row block
                const aTriple =
                  recent.length >= 2 && recent[0] === a.code && recent[1] === a.code ? 1 : 0
                const bTriple =
                  recent.length >= 2 && recent[0] === b.code && recent[1] === b.code ? 1 : 0
                if (aTriple !== bTriple) return aTriple - bTriple
                // Alternate Early/Late vs lastShiftType
                if (opposite) {
                  const aType = (shiftTypeByCode.get(a.code) || "").toLowerCase()
                  const bType = (shiftTypeByCode.get(b.code) || "").toLowerCase()
                  const aMatch = aType === opposite ? 0 : 1
                  const bMatch = bType === opposite ? 0 : 1
                  if (aMatch !== bMatch) return aMatch - bMatch
                }
                // Recent code repetition count
                const aRep = recent.filter((c) => c === a.code).length
                const bRep = recent.filter((c) => c === b.code).length
                if (aRep !== bRep) return aRep - bRep
                return 0
              })
              const picked = ranked[0]
              if (picked && picked.code !== options[0].code) {
                assignOption(ctx, picked)
                pushComputeLog({
                  date: ymd,
                  level: ctx.staff.level || "(unleveled)",
                  staff: `${ctx.staff.staffNo || ""} - ${ctx.staff.fullName || ""}`.trim() || "-",
                  code: picked.code,
                  severity: "info",
                  message: `Assigned non-repeat shift after priority phase (recent=${recent.join(",") || "-"}, prevType=${ctx.state.lastShiftType || "-"})`,
                })
                return
              }
            }
            assignOption(ctx, options[0])
            pushComputeLog({
              date: ymd,
              level: ctx.staff.level || "(unleveled)",
              staff: `${ctx.staff.staffNo || ""} - ${ctx.staff.fullName || ""}`.trim() || "-",
              code: options[0].code,
              severity: "info",
              message: `Assigned fallback shift after priority phase`,
            })
          })

        // Diagnostic: explain WHY a staff is falling through to REST. The
        // reason is written into the row.location field so it shows up in
        // the roster grid / exports for inspection.
        const restReasonFor = (ctx: (typeof contexts)[number]) => {
          const cap = maxHoursCap
          if (ctx.hasFutureLeave) return `leave on ${ymd}`
          if (ctx.state.restUntil) {
            // Display in Maldives local time (UTC+5) to match the configured
            // operating timezone — avoids confusion with raw UTC values.
            const mvt = new Date(ctx.state.restUntil.getTime() + MALDIVES_OFFSET_MS)
            const iso = mvt.toISOString().slice(0, 16).replace("T", " ")
            return `cooldown until ${iso} MVT`
          }
          if (ctx.state.allocatedHours >= cap) {
            return `rolling=${ctx.state.allocatedHours.toFixed(1)}h >= cap=${cap}h (last ${lookbackDays}d)`
          }
          if (ctx.state.consecutiveWorkDays >= maxConsecutiveWorkDays) {
            return `consecutive=${ctx.state.consecutiveWorkDays}d >= max=${maxConsecutiveWorkDays}d`
          }
          if (ctx.state.forcedRestUntilDate && ymd <= ctx.state.forcedRestUntilDate) {
            return `forced rest until ${formatYmdToDdMmYy(ctx.state.forcedRestUntilDate)}`
          }
          const levelKey = normalizeLevel(ctx.staff.level || "")
          const levelState = getLevelDayState(ymd, ctx.staff.level || "")
          const usedCodes = Array.from(levelState.usedCodes).join(",") || "none"
          const optionCount = ctx.availableOptions.length
          if (optionCount === 0) {
            // either no shifts at all assigned to this level, or every option
            // would push them past the cap
            const totalOptions = ctx.options.length
            if (totalOptions === 0) {
              return `no shifts assigned to level "${levelKey}"`
            }
            return `all ${totalOptions} shift options would exceed cap (rolling=${ctx.state.allocatedHours.toFixed(1)}h, cap=${cap}h)`
          }
          const blockedBy = ctx.availableOptions
            .filter((o) => levelState.usedCodes.has(o.code) || dayState.usedCodes.has(o.code))
            .map((o) => o.code)
          if (blockedBy.length === optionCount) {
            return `all ${optionCount} eligible codes already taken today (used: ${usedCodes})`
          }
          return `no candidate selected at level "${levelKey}" (options=${optionCount}, used: ${usedCodes})`
        }

        // Per-day diagnostic summary BEFORE the final REST pass so the user
        // can see why TP4/TP5 weren't enough to cover everyone. This lists
        // every still-unassigned staff with their constraint flags and the
        // codes available to them at their level.
        const stillUnassigned = contexts.filter(
          (ctx) => !ctx.hasFutureLeave && !assignedToday.has(ctx.token),
        )
        if (stillUnassigned.length > 0) {
          const totals = {
            cap: stillUnassigned.filter((c) => c.mustRestFromCap).length,
            cooldown: stillUnassigned.filter((c) => c.mustRestFromCooldown).length,
            consecutive: stillUnassigned.filter((c) => c.mustRestFromConsecutive).length,
            forced: stillUnassigned.filter((c) => c.mustRestFromForced).length,
            noOptions: stillUnassigned.filter(
              (c) => !c.mustRest && c.availableOptions.length === 0,
            ).length,
            options: stillUnassigned.filter(
              (c) => !c.mustRest && c.availableOptions.length > 0,
            ).length,
          }
          pushComputeLog({
            date: ymd,
            level: "(all)",
            staff: "(summary)",
            code: "REST",
            severity: "warn",
            message: `Pre-REST summary: ${stillUnassigned.length} staff still unassigned → cap=${totals.cap}, cooldown=${totals.cooldown}, consecutive=${totals.consecutive}, forced=${totals.forced}, noLevelOptions=${totals.noOptions}, hadOptionsButSkipped=${totals.options}`,
          })
          // For the "had options but skipped" group, surface what codes were
          // available so the user can tell whether TP4/TP5 should have been
          // assigned. These are the highest-priority improvement target.
          stillUnassigned
            .filter((c) => !c.mustRest && c.availableOptions.length > 0)
            .forEach((ctx) => {
              const codes = ctx.availableOptions
                .map((o) => `${o.code}(p${o.priority})`)
                .join(", ")
              pushComputeLog({
                date: ymd,
                level: ctx.staff.level || "(unleveled)",
                staff: `${ctx.staff.staffNo || ""} - ${ctx.staff.fullName || ""}`.trim() || "-",
                code: "REST",
                severity: "warn",
                message: `Unassigned despite having options: [${codes}] — falling through to REST. Check level-chain or recent-code restrictions.`,
              })
            })
        }

        // Sort the still-unassigned set so that staff at levels which DO NOT
        // already have a REST today are processed first. This lets the
        // per-level "no double rest" rescue below kick in for whichever
        // staff would have caused a 2nd REST at the same level.
        const finalRestQueue = contexts
          .filter((ctx) => !ctx.hasFutureLeave && !assignedToday.has(ctx.token))
          .sort((a, b) => {
            const aLs = getLevelDayState(ymd, a.staff.level || "").restCount
            const bLs = getLevelDayState(ymd, b.staff.level || "").restCount
            if (aLs !== bLs) return aLs - bLs
            return 0
          })
        finalRestQueue.forEach((ctx) => {
          const levelState = getLevelDayState(ymd, ctx.staff.level || "")
          // Rotational mode: any staff not given a shift today is on REST.
          // OFF is reserved for Fixed-off mode only.
          const reason = restReasonFor(ctx)
          // Cooldown rule for ALL rest days: next shift must start at least
          // restHours after the staff's last worked shift end. We compute
          // this once here and apply to both branches. If the staff has no
          // recorded shift end (e.g. start of roster), restUntil is left
          // unchanged so the loop's natural rolling-cap behaviour governs.
          const restCooldownEnd = cooldownFromShiftEnd(ctx.state.lastShiftEndTime)

          // === Per-level "no double rest" rescue ===
          // Per spec: "staff on same level should not be OF same day".
          // If this level already has a REST today, attempt to assign ANY
          // working shift to keep the same-level OFF count at 1. The
          // rescue is layered with progressively softer relaxations:
          //
          //   Tier A (soft): respect cooldown + per-level uniqueness.
          //   Tier B (medium): bypass cooldown only.
          //   Tier C (hard last resort): bypass cooldown AND reuse codes
          //          already taken at the level / globally — two staff at
          //          the same level may share a code on this day. This is
          //          ONLY taken when the alternative is a 2nd same-level
          //          OFF, which the user has explicitly forbidden.
          //
          // The 35h hour cap and max-consecutive-days are NEVER bypassed
          // by the rescue — those remain hard labour-fairness limits.
          const levelHasRest = levelState.restCount > 0
          if (levelHasRest && !ctx.mustRestFromCap && !ctx.mustRestFromConsecutive && !ctx.mustRestFromForced) {
            const remainingHours = maxHoursCap - ctx.state.allocatedHours
            const capAndConsecutiveOk = (o: (typeof ctx.options)[number]) =>
              ctx.state.allocatedHours + o.durationHours <= maxHoursCap &&
              ctx.state.consecutiveWorkDays + 1 <= maxConsecutiveWorkDays &&
              o.durationHours <= remainingHours
            const sortByLowestPriority = (
              a: (typeof ctx.options)[number],
              b: (typeof ctx.options)[number],
            ) =>
              b.priority - a.priority ||
              a.startDateTime.getTime() - b.startDateTime.getTime()

            // Tier A: cooldown + uniqueness respected.
            let tier = ""
            let recoveryOptions = ctx.availableOptions
              .filter(capAndConsecutiveOk)
              .filter((o) => canUseOption(ctx, o))
              .sort(sortByLowestPriority)
            if (recoveryOptions.length > 0) {
              tier = "tier A (cooldown+uniqueness ok)"
            } else {
              // Tier B: bypass cooldown, keep uniqueness.
              recoveryOptions = ctx.options
                .filter(capAndConsecutiveOk)
                .filter((o) => canUseOption(ctx, o))
                .sort(sortByLowestPriority)
              if (recoveryOptions.length > 0) {
                tier = "tier B (cooldown bypassed)"
              } else {
                // Tier C last resort: bypass cooldown AND uniqueness.
                // Allow reusing codes the level / day already used so two
                // staff at the same level share a code, rather than
                // produce a 2nd same-level REST.
                recoveryOptions = ctx.options
                  .filter(capAndConsecutiveOk)
                  .sort(sortByLowestPriority)
                if (recoveryOptions.length > 0) {
                  tier = "tier C (code reuse — same-level shared)"
                }
              }
            }

            if (recoveryOptions.length > 0) {
              const picked = recoveryOptions[0]
              assignOption(
                ctx,
                picked,
                `Auto level-rest-rescue ${tier}`,
              )
              pushComputeLog({
                date: ymd,
                level: ctx.staff.level || "(unleveled)",
                staff: `${ctx.staff.staffNo || ""} - ${ctx.staff.fullName || ""}`.trim() || "-",
                code: picked.code,
                severity: "warn",
                message: `Same-level no-double-rest rescue ${tier}: assigned ${picked.code} (level had ${levelState.restCount} REST already)`,
              })
              return
            }
          }

          if (ctx.mustRest || ctx.state.allocatedHours >= maxHoursCap) {
              addRow(ctx.staff, ymd, "OF", `Rest Period [${reason}]`)
              if (restCooldownEnd) ctx.state.restUntil = restCooldownEnd
              pushComputeLog({
                date: ymd,
                level: ctx.staff.level || "(unleveled)",
                staff: `${ctx.staff.staffNo || ""} - ${ctx.staff.fullName || ""}`.trim() || "-",
                code: "REST",
                severity: "info",
                message: `Rest Period: ${reason}`,
              })
            } else {
              // No working shift could be assigned (constraint-blocked or no
              // remaining options at this level). Don't zero allocatedHours —
              // the rolling-window recompute drops today's REST naturally.
              addRow(ctx.staff, ymd, "OF", `Rotational Rest [${reason}]`)
              if (restCooldownEnd) ctx.state.restUntil = restCooldownEnd
              pushComputeLog({
                date: ymd,
                level: ctx.staff.level || "(unleveled)",
                staff: `${ctx.staff.staffNo || ""} - ${ctx.staff.fullName || ""}`.trim() || "-",
                code: "REST",
                severity: "warn",
                message: `Constraint rotational rest: ${reason}`,
              })
            }
            // Track that the staff just rested so tomorrow's Late-shift
            // sequencing rule can prefer them as a "just rested" candidate.
            ctx.state.lastShiftCode = "REST"
            ctx.state.lastShiftType = "rest"
            ctx.state.recentShiftCodes = ["REST", ...ctx.state.recentShiftCodes].slice(
              0,
              RECENT_CODE_WINDOW,
            )
            const staffToken = `${(ctx.staff.staffNo || "").trim()}::${(ctx.staff.fullName || "").trim()}`
            lastRestDateByStaff.set(staffToken, ymd)
            dayState.restCount += 1
            levelState.restCount += 1
            assignedToday.add(ctx.token)
          })

        const mustCodesToday = Array.from(
          new Set(
            shiftQueue
              .filter((s) => s.requiredToday && s.typeRank <= mustPriorityCutoff)
              .map((s) => s.code),
          ),
        )
        // Enforce required-before-rest: if required codes are still missing,
        // try converting same-day REST rows into required shifts first.
        const getDayRows = () => [...kept, ...generated].filter((r) => (r.date || "") === ymd)
        const getDutyRows = (rows: Entry[]) =>
          rows.filter((r) => {
            const code = (r.shiftCode || "").trim()
            return code && !isOffCode(code) && code !== "UNSET"
          })
        const getRestRows = (rows: Entry[]) =>
          rows.filter((r) => isOffCode((r.shiftCode || "").trim()))

        let dayRows = getDayRows()
        let dutyRows = getDutyRows(dayRows)
        let restRows = getRestRows(dayRows)

        const missingBeforeRestBackfill = mustCodesToday.filter(
          (code) => !dutyRows.some((r) => (r.shiftCode || "").trim() === code),
        )
        missingBeforeRestBackfill.forEach((missingCode) => {
          const baseCandidates = contexts
            .filter((ctx) => !ctx.hasFutureLeave)
            .map((ctx) => {
              const option = ctx.options.find((o) => o.code === missingCode)
              const dayGeneratedRow = generated.find(
                (r) =>
                  (r.date || "") === ymd &&
                  (r.staffNo || "").trim() === (ctx.staff.staffNo || "").trim() &&
                  (r.staffName || "").trim() === (ctx.staff.fullName || "").trim(),
              )
              return { ctx, option, dayGeneratedRow }
            })
            .filter((x) => Boolean(x.option) && Boolean(x.dayGeneratedRow))
            .filter((x) => isOffCode((x.dayGeneratedRow?.shiftCode || "").trim()))
            .map((x) => ({
              ctx: x.ctx,
              option: x.option as NonNullable<(typeof contexts)[number]["options"][number]>,
              dayGeneratedRow: x.dayGeneratedRow as Entry,
            }))
            .filter(({ ctx, option }) => canUseOption(ctx, option))
            .sort((a, b) => a.ctx.state.allocatedHours - b.ctx.state.allocatedHours)

          const cappedCandidates = baseCandidates
            .filter(
              ({ ctx, option }) =>
                ctx.state.allocatedHours + option.durationHours <= maxHoursCap &&
                ctx.state.consecutiveWorkDays + 1 <= maxConsecutiveWorkDays,
            )
          // Cap is ABSOLUTE: never replace REST with a duty shift if doing
          // so would push the staff past maxHoursCap or maxConsecutiveWorkDays.
          // The required shift simply stays unfilled and the warning below
          // surfaces it in the compute log.
          const picked = cappedCandidates[0]
          if (!picked) {
            const blocked = baseCandidates.length > 0
            pushComputeLog({
              date: ymd,
              level: "-",
              staff: "-",
              code: missingCode,
              severity: "error",
              message: blocked
                ? `Required shift ${missingCode} unfilled: REST candidates would exceed cap (max ${maxHoursCap}h / ${maxConsecutiveWorkDays}d)`
                : `Required shift ${missingCode} unfilled: no resting staff eligible`,
            })
            return
          }
          const levelState = getLevelDayState(ymd, picked.ctx.staff.level || "")

          picked.dayGeneratedRow.shiftCode = missingCode
          picked.dayGeneratedRow.location = "Required backfill from REST"

          picked.ctx.state.lastShiftCode = missingCode
          picked.ctx.state.lastShiftType = (
            shiftTypeByCode.get(missingCode) || ""
          ).toLowerCase()
          // Replace the REST entry that was about to be pushed for today —
          // we recorded REST in recentShiftCodes when we assigned REST, so
          // pop it and push the actual code.
          if (picked.ctx.state.recentShiftCodes[0] === "REST") {
            picked.ctx.state.recentShiftCodes = picked.ctx.state.recentShiftCodes.slice(1)
          }
          picked.ctx.state.recentShiftCodes = [
            missingCode,
            ...picked.ctx.state.recentShiftCodes,
          ].slice(0, RECENT_CODE_WINDOW)
          picked.ctx.state.lastShiftEndTime = picked.option.endDateTime
          picked.ctx.state.allocatedHours += picked.option.durationHours
          if (picked.ctx.state.allocatedHours >= maxHoursCap) {
            picked.ctx.state.restUntil = cooldownFromShiftEnd(picked.ctx.state.lastShiftEndTime)
          }

          dayState.usedCodes.add(missingCode)
          dayState.usedTimeKeys.add(picked.option.timeKey)
          levelState.usedCodes.add(missingCode)
          levelState.usedTimeKeys.add(picked.option.timeKey)
          dayState.restCount = Math.max(0, dayState.restCount - 1)
          levelState.restCount = Math.max(0, levelState.restCount - 1)

          pushComputeLog({
            date: ymd,
            level: picked.ctx.staff.level || "(unleveled)",
            staff: `${picked.ctx.staff.staffNo || ""} - ${picked.ctx.staff.fullName || ""}`.trim() || "-",
            code: missingCode,
            severity: "warn",
            message: `Required shift ${missingCode} backfilled by replacing REST`,
          })
        })

        dayRows = getDayRows()
        dutyRows = getDutyRows(dayRows)
        restRows = getRestRows(dayRows)

        const missingMustCodes = mustCodesToday.filter(
          (code) => !dutyRows.some((r) => (r.shiftCode || "").trim() === code),
        )
        if (missingMustCodes.length > 0) {
          pushComputeLog({
            date: ymd,
            level: "ALL",
            staff: "-",
            code: missingMustCodes.join(", "),
            severity: "error",
            message: `Daily required coverage missing for code(s): ${missingMustCodes.join(", ")}`,
          })
        }
        const dutyCoverage = dutyRows
          .map((r) => `${(r.shiftCode || "").trim()}=${`${r.staffNo || ""} - ${r.staffName || ""}`.trim()}`)
          .join(" | ")
        pushComputeLog({
          date: ymd,
          level: "ALL",
          staff: "-",
          code: `DUTY:${dutyRows.length}/REST:${restRows.length}`,
          severity: "info",
          message:
            dutyCoverage.length > 0
              ? `Daily coverage: ${dutyCoverage}`
              : "Daily coverage: no duty shift assigned",
        })
        if (restRows.length > 1) {
          const restNames = restRows.map((r) => `${r.staffNo || ""} - ${r.staffName || ""}`.trim())
          pushComputeLog({
            date: ymd,
            level: "ALL",
            staff: restNames.join(", "),
            code: "REST",
            severity: "warn",
            message: `Multiple staff on REST same day (${restRows.length}): ${restNames.join(", ")}`,
          })
        }
        const restByLevel = new Map<string, string[]>()
        restRows.forEach((r) => {
          const level = staffLevelByNo.get((r.staffNo || "").trim()) || "unleveled"
          if (!restByLevel.has(level)) restByLevel.set(level, [])
          restByLevel.get(level)!.push(`${r.staffNo || ""} - ${r.staffName || ""}`.trim())
        })
        restByLevel.forEach((names, level) => {
          if (names.length > 1) {
            pushComputeLog({
              date: ymd,
              level,
              staff: names.join(", "),
              code: "REST",
              severity: "warn",
              message: `Same-level same-day REST conflict (${names.length})`,
            })
          }
        })
      })
    }

    const finalRowsInScope = [...generated, ...kept].filter((r) => {
      const date = (r.date || "").trim()
      if (date < rosterSettings.fromDate || date > rosterSettings.toDate) return false
      return selectedRosterStaff.some(
        (s) =>
          (s.staffNo || "").trim() === (r.staffNo || "").trim() &&
          (s.fullName || "").trim() === (r.staffName || "").trim(),
      )
    })
    const hoursByCode = new Map<string, number>()
    store.shift.forEach((s) => {
      const code = (s.shiftCode || "").trim()
      if (!code) return
      hoursByCode.set(code, shiftHoursForCode(code))
    })
    selectedRosterStaff.forEach((staff) => {
      const sno = (staff.staffNo || "").trim()
      const sname = (staff.fullName || "").trim()
      const rows = finalRowsInScope
        .filter(
          (r) =>
            (r.staffNo || "").trim() === sno &&
            (r.staffName || "").trim() === sname,
        )
        .sort((a, b) => (a.date || "").localeCompare(b.date || ""))
      let streak = 0
      rows.forEach((row, idx) => {
        const code = (row.shiftCode || "").trim()
        const isDuty = Boolean(code && !isOffCode(code) && code !== "UNSET")
        streak = isDuty ? streak + 1 : 0
        if (streak > maxConsecutiveWorkDays) {
          pushComputeLog({
            date: row.date || "-",
            level: staff.level || "(unleveled)",
            staff: `${sno} - ${sname}`.trim(),
            code,
            severity: "error",
            message: `Consecutive workday exceeded: ${streak}d > max ${maxConsecutiveWorkDays}d`,
          })
        }
        let rollingHours = 0
        for (let i = Math.max(0, idx - lookbackDays); i <= idx; i++) {
          const hCode = (rows[i].shiftCode || "").trim()
          if (!hCode || isOffCode(hCode) || hCode === "UNSET") continue
          rollingHours += hoursByCode.get(hCode) || 0
        }
        if (rollingHours > maxHoursCap + 0.001) {
          pushComputeLog({
            date: row.date || "-",
            level: staff.level || "(unleveled)",
            staff: `${sno} - ${sname}`.trim(),
            code,
            severity: "error",
            message: `Rolling hours exceeded: ${rollingHours.toFixed(1)}h > cap ${maxHoursCap}h (window ${lookbackDays + 1}d)`,
          })
        }
      })
    })

    const sortedComputeLogs = computeLogs.sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date)
      return a.at.localeCompare(b.at)
    })

    setRosterComputeLogs(sortedComputeLogs)

    const unresolvedPriorityReport = unfilledPriorityReport.filter((u) => {
      const existsInFinalDay = [...kept, ...generated].some(
        (r) => (r.date || "") === u.date && ((r.shiftCode || "").trim() === u.code),
      )
      return !existsInFinalDay
    })

    if (rosterSettings.offType !== "Fixed" && unresolvedPriorityReport.length > 0) {
      const lines = unresolvedPriorityReport
        .map((u) => `  - ${u.date} | Level: ${u.level} | Shift: ${u.code} (${u.reason})`)
        .join("\n")
      const proceed = window.confirm(
        `${unresolvedPriorityReport.length} top-priority shift slot${
          unresolvedPriorityReport.length > 1 ? "s" : ""
        } could not be filled:\n\n${lines}\n\nGenerate and save the roster anyway?`,
      )
      if (!proceed) {
        const canceledLogs = [
          ...sortedComputeLogs,
          {
            id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
            at: new Date().toISOString(),
            date: "-",
            level: "-",
            staff: "-",
            code: "-",
            severity: "warn" as const,
            message: "Roster generation canceled by user after top-priority conflicts.",
          },
        ]
        setRosterComputeLogs(canceledLogs)
        void persistRosterAuditRun(runId, canceledLogs)
        return
      }
    }

    setStore((prev) => ({
      ...prev,
      roster: [...generated, ...kept],
    }))
    void persistRosterAuditRun(runId, sortedComputeLogs)
  }

  const clearGeneratedRoster = () => {
    const targetStaff = new Set(
      selectedRosterStaff.map((s) => `${s.staffNo}::${s.fullName}`),
    )

    setStore((prev) => ({
      ...prev,
      roster: prev.roster.filter((r) => {
        const inRange =
          (r.date || "") >= rosterSettings.fromDate && (r.date || "") <= rosterSettings.toDate
        const isTargetStaff = targetStaff.has(`${r.staffNo || ""}::${r.staffName || ""}`)
        return !(inRange && isTargetStaff)
      }),
    }))
  }

  const updateForm = (fieldKey: string, value: string) => {
    setForms((prev) => ({
      ...prev,
      [activeConfig.key]: {
        ...prev[activeConfig.key],
        [fieldKey]: value,
      },
    }))
  }

  const toggleMultiSelectValue = (fieldKey: string, option: string) => {
    const normalizeForField = (value: string) => {
      if (activeConfig.key === "levelShiftPriority" && fieldKey === "rosterPriorityNumbers") {
        const p = parseLeadingNumber(value)
        return p === null ? value.trim() : String(p)
      }
      return value.trim()
    }
    const currentValues = (currentForm[fieldKey] ?? "")
      .split(",")
      .map((item) => normalizeForField(item))
      .filter(Boolean)
    const optionValue = normalizeForField(option)
    const exists = currentValues.includes(optionValue)
    const nextValues = exists
      ? currentValues.filter((item) => item !== optionValue)
      : [...currentValues, optionValue]
    updateForm(fieldKey, nextValues.join(", "))
  }

  const submitActiveModule = () => {
    const resolvedForm = { ...currentForm }
    if (activeConfig.key === "levelShiftPriority") {
      const levelOptions = [...store.levels.map((entry) => entry.levelName?.trim()), ...store.staff.map((entry) => entry.level?.trim())]
        .filter((item): item is string => Boolean(item))
      if (!resolvedForm.levelName?.trim() && levelOptions.length > 0) {
        resolvedForm.levelName = Array.from(new Set(levelOptions))[0]
      }
      const priorityOptions = Array.from(
        new Set(
          store.shift
            .map((entry) => {
              const p = parseLeadingNumber(entry.rosterPriority || "")
              const code = (entry.shiftCode || "").trim()
              if (p === null || !code) return ""
              return `${p} - ${code}`
            })
            .filter(Boolean),
        ),
      )
      if (!resolvedForm.rosterPriorityNumbers?.trim() && priorityOptions.length > 0) {
        resolvedForm.rosterPriorityNumbers = priorityOptions[0]
      }
    }
    if (activeConfig.key === "userManagement") {
      const roleOptions = Array.from(
        new Set(
          store.roleManagement
            .map((entry) => (entry.roleName || "").trim())
            .filter((item): item is string => Boolean(item)),
        ),
      )
      if (roleOptions.length === 0) {
        window.alert('Please create at least one role in "Role Management" before creating users.')
        return
      }
      if (!resolvedForm.roleName?.trim()) {
        resolvedForm.roleName = roleOptions[0]
      }
    }

    const missing = activeConfig.fields.find(
      (field) =>
        field.required &&
        !resolvedForm[field.key]?.trim() &&
        !(activeConfig.key === "leaves" && field.key === "leavePolicyName"),
    )
    if (missing) {
      window.alert(`Please complete "${missing.label}" before saving.`)
      return
    }

    if (activeConfig.key === "userManagement") {
      if (resolvedForm.userType === "Dispatch Staff User" && !resolvedForm.dispatchStaff?.trim()) {
        window.alert('Please select "Dispatch Staff" for Dispatch Staff User.')
        return
      }
      if (resolvedForm.userType === "General User" && !resolvedForm.fullName?.trim()) {
        window.alert('Please complete "Full Name" for General User.')
        return
      }
    }

    if (
      activeConfig.key === "leaveAttendanceControl" &&
      resolvedForm.allowNoOfDaysWithoutDocument === "Yes" &&
      !resolvedForm.consecutiveDaysWithoutDocument?.trim()
    ) {
      window.alert('Please enter "Consecutive Days (Without Document)".')
      return
    }
    if (
      activeConfig.key === "rosterPriorityType" &&
      resolvedForm.requiredDaily === "No" &&
      !resolvedForm.requiredWeekdays?.trim()
    ) {
      window.alert('Please select "Week Days Required" when Required Daily is "No".')
      return
    }

    if (
      (activeConfig.key === "staff" || activeConfig.key === "crewDataBase") &&
      resolvedForm.activeStatus === "Inactive" &&
      (!resolvedForm.inactiveDate?.trim() || !resolvedForm.inactiveReason?.trim())
    ) {
      window.alert('Please provide "Inactive Date" and "Inactive Reason" for inactive record.')
      return
    }

    if (activeConfig.key === "staff") {
      const selectedLevel = (currentForm.level ?? "").trim()
      const createdLevel = (currentForm.newLevel ?? "").trim()
      const finalLevel = createdLevel || selectedLevel

      if (!finalLevel) {
        window.alert('Please select "Level" or enter "Create New Level".')
        return
      }

      const staffEntry: Entry = {
        id: editingEntryId ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        createdAt: new Date().toISOString(),
        ...currentForm,
        level: finalLevel,
        loginPassword:
          currentForm.loginPassword?.trim() ||
          (editingEntryId
            ? store.staff.find((s) => s.id === editingEntryId)?.loginPassword || currentForm.staffNo
            : currentForm.staffNo),
        forcePasswordChange:
          currentForm.forcePasswordChange?.trim() ||
          (editingEntryId
            ? store.staff.find((s) => s.id === editingEntryId)?.forcePasswordChange || "Yes"
            : "Yes"),
        newLevel: "",
      }

      setStore((prev) => ({
        ...prev,
        levels: createdLevel
          ? [
              ...prev.levels,
              {
                id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
                createdAt: new Date().toISOString(),
                levelName: createdLevel,
                levelDescription: "",
              },
            ].filter(
              (entry, index, arr) =>
                arr.findIndex(
                  (item) =>
                    item.levelName?.trim().toLowerCase() === entry.levelName?.trim().toLowerCase(),
                ) === index,
            )
          : prev.levels,
        [activeConfig.key]: editingEntryId
          ? prev[activeConfig.key].map((item) => (item.id === editingEntryId ? staffEntry : item))
          : [staffEntry, ...prev[activeConfig.key]],
      }))

      setForms((prev) => ({
        ...prev,
        [activeConfig.key]: buildEmptyForm(activeConfig),
      }))
      setStaffLookupQuery("")
      setEditingEntryId(null)
      setIsFormOpen(false)
      return
    }

    if (activeConfig.key === "crewDataBase") {
      const crewCode = normalizeCrewCode(resolvedForm.crewCode || "")
      const crewName = (resolvedForm.crewName || "").trim()
      if (!/^[A-Z]{4}$/.test(crewCode)) {
        window.alert('Crew Code must be exactly 4 letters (A-Z).')
        return
      }
      if (!crewName) {
        window.alert("Please enter Crew Name.")
        return
      }
      const duplicate = store.crewDataBase.find(
        (row) => normalizeCrewCode(row.crewCode || "") === crewCode && row.id !== editingEntryId,
      )
      if (duplicate) {
        window.alert(`Crew Code "${crewCode}" already exists.`)
        return
      }
      const crewEntry: Entry = {
        id: editingEntryId ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        createdAt: new Date().toISOString(),
        ...resolvedForm,
        crewCode,
        crewName,
      }
      setStore((prev) => ({
        ...prev,
        crewDataBase: editingEntryId
          ? prev.crewDataBase.map((item) => (item.id === editingEntryId ? crewEntry : item))
          : [crewEntry, ...prev.crewDataBase],
      }))
      setForms((prev) => ({
        ...prev,
        [activeConfig.key]: buildEmptyForm(activeConfig),
      }))
      setStaffLookupQuery("")
      setEditingEntryId(null)
      setIsFormOpen(false)
      return
    }

    if (activeConfig.key === "levels") {
      const levelName = (currentForm.levelName ?? "").trim()
      if (!levelName) {
        window.alert("Please enter Level Name.")
        return
      }

      const levelEntry: Entry = {
        id: editingEntryId ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        createdAt: new Date().toISOString(),
        ...currentForm,
        levelName,
      }

      setStore((prev) => ({
        ...prev,
        [activeConfig.key]: editingEntryId
          ? prev[activeConfig.key].map((item) => (item.id === editingEntryId ? levelEntry : item))
          : [levelEntry, ...prev[activeConfig.key]].filter(
              (entry, index, arr) =>
                arr.findIndex(
                  (item) =>
                    item.levelName?.trim().toLowerCase() === entry.levelName?.trim().toLowerCase(),
                ) === index,
            ),
      }))

      setForms((prev) => ({
        ...prev,
        [activeConfig.key]: buildEmptyForm(activeConfig),
      }))
      setStaffLookupQuery("")
      setEditingEntryId(null)
      setIsFormOpen(false)
      return
    }

    if (activeConfig.key === "levelShiftPriority") {
      const levelName = (currentForm.levelName ?? "").trim()
      const priorityOrder = (currentForm.priorityOrder ?? "").trim()
      const rosterPriorityNumbers = splitMultiValue(resolvedForm.rosterPriorityNumbers ?? "")
        .map((item) => parseLeadingNumber(item))
        .filter((n): n is number => n !== null)
        .map((n) => String(n))
      const finalLevelName = (resolvedForm.levelName ?? "").trim()
      const finalPriorityOrder = (resolvedForm.priorityOrder ?? "").trim()
      if (!finalLevelName || !finalPriorityOrder || rosterPriorityNumbers.length === 0) {
        window.alert("Please select Level, Priority Order, and at least one Roster Priority Number.")
        return
      }

      const entry: Entry = {
        id: editingEntryId ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        createdAt: new Date().toISOString(),
        ...resolvedForm,
        levelName: finalLevelName,
        priorityOrder: finalPriorityOrder,
        rosterPriorityNumbers: rosterPriorityNumbers.join(", "),
      }

      setStore((prev) => {
        const next = editingEntryId
          ? prev.levelShiftPriority.map((item) => (item.id === editingEntryId ? entry : item))
          : [entry, ...prev.levelShiftPriority]

        // keep one row per level + priority; latest edit/save wins
        const deduped: Entry[] = []
        const seen = new Set<string>()
        next.forEach((item) => {
          const key = `${normalizeText(item.levelName || "")}::${(item.priorityOrder || "").trim()}`
          if (seen.has(key)) return
          seen.add(key)
          deduped.push(item)
        })
        return {
          ...prev,
          levelShiftPriority: deduped,
        }
      })

      setForms((prev) => ({
        ...prev,
        [activeConfig.key]: buildEmptyForm(activeConfig),
      }))
      setStaffLookupQuery("")
      setEditingEntryId(null)
      setIsFormOpen(false)
      return
    }

    if (activeConfig.key === "rosterPriorityType") {
      const priorityType = (resolvedForm.priorityType ?? "").trim()
      const shiftCodes = splitMultiValue(resolvedForm.shiftCodes ?? "")
      const requiredDaily = (resolvedForm.requiredDaily ?? "").trim()
      const requiredWeekdays = splitMultiValue(resolvedForm.requiredWeekdays ?? "")
      if (!priorityType || shiftCodes.length === 0 || !requiredDaily) {
        window.alert("Please select priority type, shift selection, and required daily.")
        return
      }
      if (requiredDaily === "No" && requiredWeekdays.length === 0) {
        window.alert('Please select "Week Days Required".')
        return
      }
      const entry: Entry = {
        id: editingEntryId ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        createdAt: new Date().toISOString(),
        ...resolvedForm,
        priorityType,
        shiftCodes: shiftCodes.join(", "),
        requiredDaily,
        requiredWeekdays: requiredWeekdays.join(", "),
      }

      setStore((prev) => ({
        ...prev,
        rosterPriorityType: editingEntryId
          ? prev.rosterPriorityType.map((item) => (item.id === editingEntryId ? entry : item))
          : [entry, ...prev.rosterPriorityType],
      }))
      setForms((prev) => ({
        ...prev,
        [activeConfig.key]: buildEmptyForm(activeConfig),
      }))
      setStaffLookupQuery("")
      setEditingEntryId(null)
      setIsFormOpen(false)
      return
    }

    if (activeConfig.key === "publicHolidays") {
      if (currentForm.holidayStartDate > currentForm.holidayEndDate) {
        window.alert('"Holiday End Date" must be after or equal to "Holiday Start Date".')
        return
      }
    }

    if (activeConfig.key === "shift") {
      const shiftDuration = calculateShiftDuration(currentForm.startTime, currentForm.endTime)
      const shiftEntry: Entry = {
        id: editingEntryId ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        createdAt: new Date().toISOString(),
        ...currentForm,
        shiftIcon: "",
        shiftDuration,
      }

      setStore((prev) => ({
        ...prev,
        [activeConfig.key]: editingEntryId
          ? prev[activeConfig.key].map((item) => (item.id === editingEntryId ? shiftEntry : item))
          : [shiftEntry, ...prev[activeConfig.key]],
      }))

      setForms((prev) => ({
        ...prev,
        [activeConfig.key]: buildEmptyForm(activeConfig),
      }))
      setStaffLookupQuery("")
      setEditingEntryId(null)
      setIsFormOpen(false)
      return
    }

    if (activeConfig.key === "leaves") {
      const fromDate = currentForm.fromDate
      const toDate = currentForm.toDate
      const leavePolicyName = currentForm.leavePolicyName
      const availableLeavePolicies = store.leaveAttendanceControl
        .filter((item) => item.leaveAttendanceType === "Leave")
        .map((item) => item.leaveAttendanceName?.trim())
        .filter((item): item is string => Boolean(item))

      if (availableLeavePolicies.length === 0) {
        window.alert(
          "No Leave Policy found. Please create one in Leave Type.",
        )
        return
      }

      if (!leavePolicyName?.trim()) {
        window.alert("Please select a Leave Policy before saving.")
        return
      }

      if (!currentForm.staffNo?.trim() || !currentForm.staffName?.trim()) {
        window.alert("Please search and select a staff member.")
        return
      }
      if (fromDate > toDate) {
        window.alert('"To Date" must be after or equal to "From Date".')
        return
      }
      const policy = store.leaveAttendanceControl.find(
        (item) =>
          item.leaveAttendanceType === "Leave" &&
          item.leaveAttendanceName === leavePolicyName,
      )
      if (!policy) {
        window.alert("Selected leave policy was not found.")
        return
      }
      const leaveSummary = calculateLeaveSummary(fromDate, toDate, leavePolicyName)

      const leaveEntry: Entry = {
        id: editingEntryId ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        createdAt: new Date().toISOString(),
        ...currentForm,
        fromDate: leaveSummary.fromDate,
        toDate: leaveSummary.toDate,
        staffLookup: `${currentForm.staffNo} - ${currentForm.staffName}`,
        totalCalendarDays: leaveSummary.totalCalendarDays,
        excludedWeekdayDetails: leaveSummary.excludedWeekdayDetails,
        publicHolidayDays: leaveSummary.publicHolidayDays,
        chargeableLeaveDays: leaveSummary.chargeableLeaveDays,
      }

      setStore((prev) => ({
        ...prev,
        [activeConfig.key]: editingEntryId
          ? prev[activeConfig.key].map((entry) => (entry.id === editingEntryId ? leaveEntry : entry))
          : [leaveEntry, ...prev[activeConfig.key]],
      }))

      setForms((prev) => ({
        ...prev,
        [activeConfig.key]: buildEmptyForm(activeConfig),
      }))
      setStaffLookupQuery("")
      setEditingEntryId(null)
      setIsFormOpen(false)
      return
    }

    const entry: Entry = {
      id: editingEntryId ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      createdAt: new Date().toISOString(),
      ...resolvedForm,
    }

    setStore((prev) => ({
      ...prev,
      [activeConfig.key]: editingEntryId
        ? prev[activeConfig.key].map((item) => (item.id === editingEntryId ? entry : item))
        : [entry, ...prev[activeConfig.key]],
    }))

    setForms((prev) => ({
      ...prev,
      [activeConfig.key]: buildEmptyForm(activeConfig),
    }))
    setStaffLookupQuery("")
    setEditingEntryId(null)
    setIsFormOpen(false)
  }

  const getSelectOptions = (field: FieldConfig): string[] => {
    if (
      (activeConfig.key === "staff" && field.key === "level") ||
      (activeConfig.key === "levelShiftPriority" && field.key === "levelName")
    ) {
      const levels = [...store.levels.map((entry) => entry.levelName?.trim()), ...store.staff.map((entry) => entry.level?.trim())]
        .filter((item): item is string => Boolean(item))
      return Array.from(new Set(levels))
    }

    if (activeConfig.key === "leaves" && field.key === "leavePolicyName") {
      const names = store.leaveAttendanceControl
        .filter((entry) => entry.leaveAttendanceType === "Leave")
        .map((entry) => entry.leaveAttendanceName?.trim())
        .filter((item): item is string => Boolean(item))
      return Array.from(new Set(names))
    }
    if (activeConfig.key === "userManagement" && field.key === "roleName") {
      return Array.from(
        new Set(
          store.roleManagement
            .map((entry) => (entry.roleName || "").trim())
            .filter((item): item is string => Boolean(item)),
        ),
      )
    }
    if (activeConfig.key === "userManagement" && field.key === "dispatchStaff") {
      return store.staff
        .filter((staff) => (staff.activeStatus || "Active") === "Active")
        .map((staff) => `${staff.staffNo} - ${staff.fullName}`)
    }
    return field.options ?? []
  }

  const getMultiSelectOptions = (field: FieldConfig): string[] => {
    if (activeConfig.key === "shift" && field.key === "assignedLevels") {
      const levels = [...store.levels.map((entry) => entry.levelName?.trim()), ...store.staff.map((entry) => entry.level?.trim())]
        .filter((item): item is string => Boolean(item))
      return Array.from(new Set(levels))
    }
    if (activeConfig.key === "rosterPriorityType" && field.key === "shiftCodes") {
      return Array.from(
        new Set(
          store.shift
            .map((entry) => entry.shiftCode?.trim())
            .filter((item): item is string => Boolean(item)),
        ),
      ).sort()
    }
    if (activeConfig.key === "levelShiftPriority" && field.key === "rosterPriorityNumbers") {
      const byPriority = new Map<number, Set<string>>()
      store.shift.forEach((entry) => {
        const p = parseLeadingNumber(entry.rosterPriority || "")
        const code = (entry.shiftCode || "").trim()
        if (p === null || !code) return
        if (!byPriority.has(p)) byPriority.set(p, new Set())
        byPriority.get(p)!.add(code)
      })
      return Array.from(byPriority.entries())
        .sort((a, b) => a[0] - b[0])
        .map(([p, codes]) => `${p} - ${Array.from(codes).sort().join(", ")}`)
    }
    return field.options ?? []
  }

  useEffect(() => {
    const leavePolicyNames = store.leaveAttendanceControl
      .filter((entry) => entry.leaveAttendanceType === "Leave")
      .map((entry) => entry.leaveAttendanceName?.trim())
      .filter((item): item is string => Boolean(item))
    const uniqueLeavePolicyNames = Array.from(new Set(leavePolicyNames))

    if (uniqueLeavePolicyNames.length === 0) return
    if (forms.leaves.leavePolicyName?.trim()) return

    setForms((prev) => ({
      ...prev,
      leaves: {
        ...prev.leaves,
        leavePolicyName: uniqueLeavePolicyNames[0],
      },
    }))
  }, [store.leaveAttendanceControl, forms.leaves.leavePolicyName])

  useEffect(() => {
    const levelOptions = Array.from(
      new Set(
        [...store.levels.map((entry) => entry.levelName?.trim()), ...store.staff.map((entry) => entry.level?.trim())]
          .filter((item): item is string => Boolean(item)),
      ),
    )
    const priorityOptions = Array.from(
      new Set(
        store.shift
          .map((entry) => {
            const p = parseLeadingNumber(entry.rosterPriority || "")
            const code = (entry.shiftCode || "").trim()
            if (p === null || !code) return ""
            return `${p} - ${code}`
          })
          .filter(Boolean),
      ),
    )

    setForms((prev) => {
      const current = prev.levelShiftPriority
      const nextLevel = current.levelName?.trim() || levelOptions[0] || ""
      const nextRosterPriorityNumbers =
        current.rosterPriorityNumbers?.trim() || priorityOptions[0] || ""
      if (
        nextLevel === current.levelName &&
        nextRosterPriorityNumbers === current.rosterPriorityNumbers
      ) {
        return prev
      }
      return {
        ...prev,
        levelShiftPriority: {
          ...current,
          levelName: nextLevel,
          rosterPriorityNumbers: nextRosterPriorityNumbers,
        },
      }
    })
  }, [store.levels, store.staff, store.shift])

  useEffect(() => {
    const roleOptions = Array.from(
      new Set(
        store.roleManagement
          .map((entry) => (entry.roleName || "").trim())
          .filter((item): item is string => Boolean(item)),
      ),
    )
    if (roleOptions.length === 0) return
    setForms((prev) => {
      const currentRole = (prev.userManagement.roleName || "").trim()
      if (currentRole) return prev
      return {
        ...prev,
        userManagement: {
          ...prev.userManagement,
          roleName: roleOptions[0],
        },
      }
    })
  }, [store.roleManagement])

  const removeEntry = (moduleKey: ModuleKey, entryId: string) => {
    if (!window.confirm("Are you sure you want to delete this record?")) return
    setStore((prev) => ({
      ...prev,
      [moduleKey]: prev[moduleKey].filter((entry) => entry.id !== entryId),
    }))
  }

  const saveCrewLeavePolicy = () => {
    const policyName = (crewLeavePolicyForm.policyName || "").trim()
    const crewType = (crewLeavePolicyForm.crewType || "").trim()
    const workPatternDays = Number(crewLeavePolicyForm.workPatternDays || "0")
    const leavePatternDays = Number(crewLeavePolicyForm.leavePatternDays || "0")
    const thresholds: Record<string, string> = {}
    Object.entries(crewLeavePolicyThresholds).forEach(([code, value]) => {
      const n = Math.max(Number(value || "0"), 0)
      if (!Number.isFinite(n)) return
      thresholds[code] = String(n)
    })
    if (!policyName || !crewType) {
      window.alert("Policy Name and Crew Type are required.")
      return
    }
    if (!Number.isFinite(workPatternDays) || workPatternDays <= 0) {
      window.alert("Work Pattern Days must be greater than 0.")
      return
    }
    if (!Number.isFinite(leavePatternDays) || leavePatternDays <= 0) {
      window.alert("Leave Pattern Days must be greater than 0.")
      return
    }
    const entry: Entry = {
      id: editingRotationId ?? `clp_${Date.now()}_${Math.random().toString(16).slice(2, 6)}`,
      createdAt: new Date().toISOString(),
      recordType: "policy",
      policyName,
      crewType,
      workPatternDays: String(workPatternDays),
      leavePatternDays: String(leavePatternDays),
      codeThresholdsJson: JSON.stringify(thresholds),
      activeStatus: crewLeavePolicyForm.activeStatus || "Active",
    }
    setStore((prev) => ({
      ...prev,
      crewLeavePlanner: editingRotationId
        ? prev.crewLeavePlanner.map((row) => (row.id === editingRotationId ? entry : row))
        : [entry, ...prev.crewLeavePlanner],
    }))
    setCrewLeavePolicyForm((prev) => ({ ...prev, policyName: "" }))
    setEditingRotationId(null)
  }

  const editCrewLeavePolicy = (policyId: string) => {
    const p = crewLeavePolicies.find((row) => row.id === policyId)
    if (!p) return
    setCrewLeavePolicyForm({
      policyName: p.policyName || "",
      crewType: p.crewType || "Captain",
      workPatternDays: p.workPatternDays || "90",
      leavePatternDays: p.leavePatternDays || "30",
      codeThresholdsJson: p.codeThresholdsJson || "{}",
      activeStatus: p.activeStatus || "Active",
    })
    try {
      const parsed = JSON.parse(p.codeThresholdsJson || "{}") as Record<string, string>
      setCrewLeavePolicyThresholds(parsed || {})
    } catch {
      setCrewLeavePolicyThresholds({})
    }
    setEditingRotationId(p.id)
    setIsRotationModalOpen(true)
  }

  const deleteCrewLeavePolicy = (policyId: string) => {
    if (!window.confirm("Delete this rotation type?")) return
    setStore((prev) => ({
      ...prev,
      crewLeavePlanner: prev.crewLeavePlanner.filter((row) => row.id !== policyId),
    }))
  }

  const assignCrewLeavePolicy = () => {
    const crewCode = normalizeCrewCode(crewLeaveAssignmentForm.crewCode || "")
    const policyId = (crewLeaveAssignmentForm.policyId || "").trim()
    if (!crewCode || !policyId) {
      window.alert("Crew Code and Policy are required.")
      return
    }
    const policy = crewLeavePolicies.find((p) => p.id === policyId)
    if (!policy) {
      window.alert("Selected policy was not found.")
      return
    }
    const crew = store.crewDataBase.find((c) => normalizeCrewCode(c.crewCode || "") === crewCode)
    if (!crew) {
      window.alert("Selected crew was not found in Crew Data Base.")
      return
    }
    if ((crew.crewType || "").trim() !== (policy.crewType || "").trim()) {
      window.alert("Selected policy is not eligible for this crew type.")
      return
    }

    const entry: Entry = {
      id: `cla_${Date.now()}_${Math.random().toString(16).slice(2, 6)}`,
      createdAt: new Date().toISOString(),
      recordType: "assignment",
      crewCode,
      policyId,
      policyName: policy.policyName || "",
      crewType: crew.crewType || "",
    }
    setStore((prev) => {
      const kept = prev.crewLeavePlanner.filter(
        (row) =>
          !(
            (row.recordType || "").trim().toLowerCase() === "assignment" &&
            normalizeCrewCode(row.crewCode || "") === crewCode
          ),
      )
      return { ...prev, crewLeavePlanner: [entry, ...kept] }
    })
    setCrewLeaveAssignmentForm({ crewCode: "", policyId: "" })
  }

  const addCrewLeaveMark = () => {
    const crewCode = normalizeCrewCode(crewLeaveMarkForm.crewCode || "")
    const markType = (crewLeaveMarkForm.markType || "").trim().toUpperCase()
    let cycleNumber = (crewLeaveMarkForm.cycleNumber || "").trim()
    const fromDate = (crewLeaveMarkForm.fromDate || "").trim()
    const toDate = (crewLeaveMarkForm.toDate || "").trim() || fromDate
    if (!crewCode || !markType || !fromDate || !toDate) {
      window.alert("Crew Code, Mark Type, From Date and To Date are required.")
      return
    }
    if (toDate < fromDate) {
      window.alert('"To Date" must be after or equal to "From Date".')
      return
    }
    if (!["MC", "AB", "SD"].includes(markType)) {
      window.alert("Mark Type must be MC, AB or SD.")
      return
    }
    const noOfDays = getDateRangeInclusive(fromDate, toDate).length
    const crew = store.crewDataBase.find((c) => normalizeCrewCode(c.crewCode || "") === crewCode)
    if (!crew) {
      window.alert("Selected crew was not found.")
      return
    }
    if (!cycleNumber) {
      cycleNumber = findMatchedWorkCycle(crewCode, fromDate, toDate)
    }
    if (!cycleNumber) {
      window.alert("No matching work cycle found for selected dates. Select Deduction Cycle.")
      return
    }
    const entry: Entry = {
      id: editingCrewLeaveMarkId || `clm_${Date.now()}_${Math.random().toString(16).slice(2, 6)}`,
      createdAt:
        store.crewLeavePlanner.find((r) => r.id === editingCrewLeaveMarkId)?.createdAt || new Date().toISOString(),
      recordType: "mark",
      crewCode,
      crewName: crew.crewName || "",
      crewType: crew.crewType || "",
      markType,
      cycleNumber,
      noOfDays: String(noOfDays),
      fromDate,
      toDate,
      reason: (crewLeaveMarkForm.reason || "").trim(),
    }
    setStore((prev) => {
      const next = editingCrewLeaveMarkId
        ? prev.crewLeavePlanner.map((row) => (row.id === editingCrewLeaveMarkId ? entry : row))
        : [entry, ...prev.crewLeavePlanner]
      return { ...prev, crewLeavePlanner: next }
    })
    setEditingCrewLeaveMarkId(null)
    setCrewLeaveMarkForm((prev) => ({ ...prev, cycleNumber: "", reason: "" }))
  }

  const editCrewLeaveMark = (id: string) => {
    const row = crewLeaveMarks.find((m) => m.id === id)
    if (!row) return
    setCrewLeaveMarkForm({
      crewCode: normalizeCrewCode(row.crewCode || ""),
      markType: (row.markType || "MC").toUpperCase(),
      cycleNumber: (row.cycleNumber || "").trim(),
      fromDate: row.fromDate || toYmd(new Date()),
      toDate: row.toDate || row.fromDate || toYmd(new Date()),
      reason: row.reason || "",
    })
    setEditingCrewLeaveMarkId(id)
  }

  const deleteCrewLeaveMark = (id: string) => {
    if (!window.confirm("Delete this daily attendance record?")) return
    setStore((prev) => ({
      ...prev,
      crewLeavePlanner: prev.crewLeavePlanner.filter((row) => row.id !== id),
    }))
    if (editingCrewLeaveMarkId === id) {
      setEditingCrewLeaveMarkId(null)
      setCrewLeaveMarkForm({
        crewCode: "",
        markType: "MC",
        cycleNumber: "",
        fromDate: toYmd(new Date()),
        toDate: toYmd(new Date()),
        reason: "",
      })
    }
  }

  const setCrewOpsCodeDeduction = (code: "MC" | "AB" | "SD", allowed: boolean) => {
    setStore((prev) => {
      const existing = prev.crewLeavePlanner.find(
        (row) =>
          (row.recordType || "").trim().toLowerCase() === "codeconfig" &&
          (row.code || "").trim().toUpperCase() === code,
      )
      const entry: Entry = {
        id: existing?.id || `clc_${Date.now()}_${Math.random().toString(16).slice(2, 6)}`,
        createdAt: existing?.createdAt || new Date().toISOString(),
        recordType: "codeConfig",
        code,
        allowLeaveDeduction: allowed ? "Yes" : "No",
      }
      const next = existing
        ? prev.crewLeavePlanner.map((row) => (row.id === existing.id ? entry : row))
        : [entry, ...prev.crewLeavePlanner]
      return { ...prev, crewLeavePlanner: next }
    })
  }

  const saveCrewOpsCode = () => {
    const code = (crewOpsCodeForm.code || "").trim().toUpperCase()
    const name = (crewOpsCodeForm.name || "").trim()
    if (!code) {
      window.alert("Code is required.")
      return
    }
    if (!name) {
      window.alert("Code name is required.")
      return
    }
    setStore((prev) => {
      const existing = prev.crewLeavePlanner.find(
        (row) =>
          (row.recordType || "").trim().toLowerCase() === "codeconfig" &&
          (row.code || "").trim().toUpperCase() === code,
      )
      const entry: Entry = {
        id: existing?.id || `clc_${Date.now()}_${Math.random().toString(16).slice(2, 6)}`,
        createdAt: existing?.createdAt || new Date().toISOString(),
        recordType: "codeConfig",
        code,
        codeName: name,
        allowLeaveDeduction: crewOpsCodeForm.allowLeaveDeduction === "Yes" ? "Yes" : "No",
      }
      const next = existing
        ? prev.crewLeavePlanner.map((row) => (row.id === existing.id ? entry : row))
        : [entry, ...prev.crewLeavePlanner]
      return { ...prev, crewLeavePlanner: next }
    })
    setCrewOpsCodeForm({ code: "MC", name: "", allowLeaveDeduction: "Yes" })
  }

  const editCrewOpsCode = (id: string) => {
    const row = crewOpsCodeConfigs.find((x) => x.id === id)
    if (!row) return
    setCrewOpsCodeForm({
      code: (row.code || "").trim().toUpperCase() || "MC",
      name: row.codeName || "",
      allowLeaveDeduction: isYesLike(row.allowLeaveDeduction) ? "Yes" : "No",
    })
  }

  const deleteCrewOpsCode = (id: string) => {
    if (!window.confirm("Delete this attendance code?")) return
    setStore((prev) => ({
      ...prev,
      crewLeavePlanner: prev.crewLeavePlanner.filter((row) => row.id !== id),
    }))
  }

  const resetCrewDutyTypeForm = () => {
    setCrewDutyTypeForm({
      dutyType: "",
      rosterBarMap: "ROSTER DAY",
    })
    setEditingCrewDutyTypeId(null)
    setEditingCrewDutyTypeName("")
  }

  const saveCrewDutyType = () => {
    const dutyType = (crewDutyTypeForm.dutyType || "").trim()
    const rosterBarMap = crewDutyTypeForm.rosterBarMap === "REST" ? "REST" : "ROSTER DAY"
    if (!dutyType) {
      window.alert("Duty type is required.")
      return
    }
    setStore((prev) => {
      const existingById = editingCrewDutyTypeId
        ? prev.crewLeavePlanner.find((row) => row.id === editingCrewDutyTypeId)
        : undefined
      const existingByName = prev.crewLeavePlanner.find(
        (row) =>
          (row.recordType || "").trim().toLowerCase() === "dutytypeconfig" &&
          (row.dutyType || "").trim().toLowerCase() === dutyType.toLowerCase(),
      )
      const isEditing = Boolean(editingCrewDutyTypeName)
      if (existingByName && (!isEditing || existingByName.id !== existingById?.id)) {
        window.alert(`Duty type "${dutyType}" already exists.`)
        return prev
      }
      const existing = existingById
      const entry: Entry = {
        id: existing?.id || `cdt_${Date.now()}_${Math.random().toString(16).slice(2, 6)}`,
        createdAt: existing?.createdAt || new Date().toISOString(),
        recordType: "dutyTypeConfig",
        dutyType,
        rosterBarMap,
        isDefault: existing?.isDefault || "No",
      }
      const next = existing
        ? prev.crewLeavePlanner.map((row) => (row.id === existing.id ? entry : row))
        : [entry, ...prev.crewLeavePlanner]
      return { ...prev, crewLeavePlanner: next }
    })
    resetCrewDutyTypeForm()
  }

  const editCrewDutyType = (id: string) => {
    const row = crewDutyTypeConfigs.find((x) => x.id === id)
    if (!row) return
    setEditingCrewDutyTypeId(id.startsWith("default-duty-type-") ? null : id)
    setEditingCrewDutyTypeName(row.dutyType || "")
    setCrewDutyTypeForm({
      dutyType: row.dutyType || "Flight Duty",
      rosterBarMap: row.rosterBarMap === "REST" ? "REST" : "ROSTER DAY",
    })
  }

  const deleteCrewDutyType = (id: string) => {
    const target = crewDutyTypeConfigs.find((row) => row.id === id)
    const dutyType = (target?.dutyType || "").trim()
    if (!window.confirm("Delete this duty type?")) return
    if (id.startsWith("default-duty-type-")) {
      if (!dutyType) return
      setStore((prev) => ({
        ...prev,
        crewLeavePlanner: [
          {
            id: `cdt_hidden_${Date.now()}_${Math.random().toString(16).slice(2, 6)}`,
            createdAt: new Date().toISOString(),
            recordType: "dutyTypeConfig",
            dutyType,
            rosterBarMap: target?.rosterBarMap === "REST" ? "REST" : "ROSTER DAY",
            hidden: "Yes",
            isDefault: "Yes",
          },
          ...prev.crewLeavePlanner,
        ],
      }))
      if (editingCrewDutyTypeName === dutyType) resetCrewDutyTypeForm()
      return
    }
    setStore((prev) => ({
      ...prev,
      crewLeavePlanner: prev.crewLeavePlanner.filter((row) => row.id !== id),
    }))
    if (editingCrewDutyTypeId === id) resetCrewDutyTypeForm()
    if (editingCrewDutyTypeName && dutyType === editingCrewDutyTypeName) {
      resetCrewDutyTypeForm()
    }
  }

  const resetCrewDutyCodeForm = () => {
    setCrewDutyCodeForm({
      code: "",
      name: "",
      dutyType: "Flight Duty",
      startTime: "00:00",
      endTime: "00:00",
      color: "#0ea5a5",
    })
    setEditingCrewDutyCodeId(null)
  }

  const saveCrewDutyCode = () => {
    const code = (crewDutyCodeForm.code || "").trim().toUpperCase()
    const name = (crewDutyCodeForm.name || "").trim()
    const dutyType = (crewDutyCodeForm.dutyType || "Flight Duty").trim()
    const startTime = (crewDutyCodeForm.startTime || "").trim()
    const endTime = (crewDutyCodeForm.endTime || "").trim()
    const color = (crewDutyCodeForm.color || getDefaultDutyCodeColor(code)).trim()
    if (!code) {
      window.alert("Duty code is required.")
      return
    }
    if (!name) {
      window.alert("Duty name is required.")
      return
    }
    if (!startTime || !endTime) {
      window.alert("Default start and end time are required.")
      return
    }
    setStore((prev) => {
      const duplicate = prev.crewLeavePlanner.find(
        (row) =>
          (row.recordType || "").trim().toLowerCase() === "dutycodeconfig" &&
          row.id !== editingCrewDutyCodeId &&
          (row.code || "").trim().toUpperCase() === code,
      )
      if (duplicate) {
        window.alert(`Duty code "${code}" already exists.`)
        return prev
      }
      const existing = editingCrewDutyCodeId
        ? prev.crewLeavePlanner.find((row) => row.id === editingCrewDutyCodeId)
        : undefined
      const entry: Entry = {
        id: existing?.id || `cdc_${Date.now()}_${Math.random().toString(16).slice(2, 6)}`,
        createdAt: existing?.createdAt || new Date().toISOString(),
        recordType: "dutyCodeConfig",
        code,
        dutyName: name,
        dutyType,
        rosterBarMap: crewDutyTypeMap.get(dutyType.toLowerCase()) || getCrewDutyRosterBarMap(dutyType),
        defaultStartTime: startTime,
        defaultEndTime: endTime,
        color,
      }
      const next = existing
        ? prev.crewLeavePlanner.map((row) => (row.id === existing.id ? entry : row))
        : [entry, ...prev.crewLeavePlanner]
      return { ...prev, crewLeavePlanner: next }
    })
    resetCrewDutyCodeForm()
  }

  const editCrewDutyCode = (id: string) => {
    const row = crewDutyCodeConfigs.find((x) => x.id === id)
    if (!row) return
    setEditingCrewDutyCodeId(id)
    setCrewDutyCodeForm({
      code: (row.code || "").trim().toUpperCase(),
      name: row.dutyName || "",
      dutyType: row.dutyType || "Flight Duty",
      startTime: row.defaultStartTime || "00:00",
      endTime: row.defaultEndTime || "00:00",
      color: row.color || getDefaultDutyCodeColor(row.code || ""),
    })
  }

  const deleteCrewDutyCode = (id: string) => {
    if (!window.confirm("Delete this duty code?")) return
    setStore((prev) => ({
      ...prev,
      crewLeavePlanner: prev.crewLeavePlanner.filter((row) => row.id !== id),
    }))
    if (editingCrewDutyCodeId === id) resetCrewDutyCodeForm()
  }

  const generateCrewLeaveBlocks = () => {
    const selectedCrewCode = normalizeCrewCode(
      crewLeaveAssignmentForm.crewCode || crewLeaveGeneratorForm.crewCode || "",
    )
    if (!selectedCrewCode) {
      window.alert("Please select a crew member.")
      return
    }
    const years = Number(crewLeaveGeneratorForm.years || "0")
    if (!Number.isFinite(years) || years <= 0 || years > 10) {
      window.alert("Years must be between 1 and 10.")
      return
    }
    const targetCrew = store.crewDataBase.filter(
      (c) => normalizeCrewCode(c.crewCode || "") === selectedCrewCode,
    )
    if (targetCrew.length === 0) {
      window.alert("No crew found for generation.")
      return
    }

    const policyById = new Map<string, Entry>()
    crewLeavePolicies.forEach((p) => policyById.set(p.id, p))
    const assignmentByCrewCode = new Map<string, Entry>()
    crewLeaveAssignments.forEach((a) => assignmentByCrewCode.set(normalizeCrewCode(a.crewCode || ""), a))
    const selectedPolicyId = (crewLeaveAssignmentForm.policyId || "").trim()
    const selectedPolicy = selectedPolicyId ? policyById.get(selectedPolicyId) : undefined
    const replaceExisting = crewLeaveGeneratorForm.replaceExisting === "Yes"

    const generated: Entry[] = []
    let skippedNoRelease = 0
    let skippedNoPolicy = 0
    targetCrew.forEach((crew) => {
      const crewCode = normalizeCrewCode(crew.crewCode || "")
      const crewName = (crew.crewName || "").trim()
      const crewType = (crew.crewType || "").trim()
      const releaseDate = (crew.releaseDate || "").trim()
      if (!releaseDate) {
        skippedNoRelease += 1
        return
      }
      const assignment = assignmentByCrewCode.get(crewCode)
      const policy =
        selectedPolicy ||
        (assignment?.policyId ? policyById.get(assignment.policyId) : undefined) ||
        crewLeavePolicies.find(
          (p) =>
            (p.activeStatus || "").trim() === "Active" &&
            (p.crewType || "").trim() === crewType,
        )
      if (selectedPolicy && (selectedPolicy.crewType || "").trim() !== crewType) {
        skippedNoPolicy += 1
        return
      }
      const workDays = Number(policy?.workPatternDays || "0")
      const leaveDays = Number(policy?.leavePatternDays || "0")
      if (!policy || !Number.isFinite(workDays) || !Number.isFinite(leaveDays) || workDays <= 0 || leaveDays <= 0) {
        skippedNoPolicy += 1
        return
      }
      const horizonEnd = addDaysToYmd(releaseDate, Math.round(years * 365))
      let cursor = releaseDate
      let cycle = 1
      while (cursor < horizonEnd) {
        const workStart = cursor
        const workEnd = addDaysToYmd(workStart, workDays - 1)
        generated.push({
          id: `clb_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`,
          createdAt: new Date().toISOString(),
          recordType: "generatedBlock",
          crewCode,
          crewName,
          crewType,
          policyId: policy.id,
          policyName: policy.policyName || "",
          cycleNumber: String(cycle),
          blockType: "work",
          startDate: workStart,
          endDate: workEnd,
          originalStartDate: workStart,
          originalEndDate: workEnd,
          plannedDays: String(workDays),
        })
        const leaveStart = addDaysToYmd(workEnd, 1)
        const leaveEnd = addDaysToYmd(leaveStart, leaveDays - 1)
        generated.push({
          id: `clb_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`,
          createdAt: new Date().toISOString(),
          recordType: "generatedBlock",
          crewCode,
          crewName,
          crewType,
          policyId: policy.id,
          policyName: policy.policyName || "",
          cycleNumber: String(cycle),
          blockType: "leave",
          startDate: leaveStart,
          endDate: leaveEnd,
          originalStartDate: leaveStart,
          originalEndDate: leaveEnd,
          plannedDays: String(leaveDays),
        })
        cursor = addDaysToYmd(leaveEnd, 1)
        cycle += 1
      }
    })

    setStore((prev) => {
      const kept = replaceExisting
        ? prev.crewLeavePlanner.filter((row) => {
            if ((row.recordType || "").trim().toLowerCase() !== "generatedblock") return true
            if (!selectedCrewCode) return false
            return normalizeCrewCode(row.crewCode || "") !== selectedCrewCode
          })
        : prev.crewLeavePlanner
      return { ...prev, crewLeavePlanner: [...generated, ...kept] }
    })

    window.alert(
      `Generated ${generated.length} block(s).` +
        (skippedNoRelease > 0 ? `\nSkipped ${skippedNoRelease} crew (no Release Date).` : "") +
        (skippedNoPolicy > 0 ? `\nSkipped ${skippedNoPolicy} crew (no eligible active policy).` : ""),
    )
  }

  const generateCrewRosterBars = () => {
    const fromDate = (crewRosterGeneratorForm.fromDate || "").trim()
    const toDate = (crewRosterGeneratorForm.toDate || "").trim()
    if (!fromDate || !toDate || toDate < fromDate) {
      window.alert("Select a valid roster period.")
      return
    }
    const selectedCrewCode = normalizeCrewCode(crewRosterGeneratorForm.crewCode || "")
    const weekdayName = (ymd: string) =>
      new Date(`${ymd}T00:00:00`).toLocaleDateString("en-US", { weekday: "long" })
    const maxRecoveryRestGapHours = Number(crewRosterGeneratorForm.maxRecoveryRestGap || "0")
    if (!Number.isFinite(maxRecoveryRestGapHours) || maxRecoveryRestGapHours <= 0) {
      window.alert("Maximum Time Between Two Recovery Rest Period must be greater than 0.")
      return
    }
    const dutyCodeByCode = new Map<string, Entry>()
    crewDutyCodeConfigs.forEach((row) => {
      const code = (row.code || "").trim().toUpperCase()
      if (code) dutyCodeByCode.set(code, row)
    })
    const selectedDefaultDutyCode = (crewRosterGeneratorForm.defaultDutyCode || "FD").trim().toUpperCase()
    const defaultDutyConfig = dutyCodeByCode.get(selectedDefaultDutyCode)
    if (!defaultDutyConfig) {
      window.alert(`Default Duty Code ${selectedDefaultDutyCode || "FD"} is not found in Duty Code Manager.`)
      return
    }
    const selectedDefaultDutyType = (defaultDutyConfig.dutyType || "Flight Duty").trim()
    const selectedDefaultDutyMap =
      crewDutyTypeMap.get(selectedDefaultDutyType.toLowerCase()) ||
      defaultDutyConfig.rosterBarMap ||
      getCrewDutyRosterBarMap(selectedDefaultDutyType)
    if (selectedDefaultDutyMap === "REST") {
      window.alert("Default Duty Code must map to Roster Day, not Rest.")
      return
    }
    const maxDutyTimePerDayHours = Number(crewRosterGeneratorForm.maxDutyTimePerDay || "13")
    if (!Number.isFinite(maxDutyTimePerDayHours) || maxDutyTimePerDayHours <= 0) {
      window.alert("Maximum Duty Time / Day must be greater than 0.")
      return
    }
    const dutySequenceCodes = (crewRosterGeneratorForm.dutyCodeSequence || selectedDefaultDutyCode)
      .split(",")
      .map((code) => code.trim().toUpperCase())
      .filter(Boolean)
    if (dutySequenceCodes.length === 0) {
      window.alert("Enter at least one Duty Code in Duty Code Sequence.")
      return
    }
    const dutySequenceConfigs = dutySequenceCodes.map((code) => {
      const config = dutyCodeByCode.get(code)
      if (!config) return null
      const dutyType = (config.dutyType || "Flight Duty").trim()
      const map =
        crewDutyTypeMap.get(dutyType.toLowerCase()) ||
        config.rosterBarMap ||
        getCrewDutyRosterBarMap(dutyType)
      if (map === "REST") return null
      return { code, config, dutyType }
    })
    if (dutySequenceConfigs.some((item) => !item)) {
      window.alert("Duty Code Sequence can only include valid codes mapped to Roster Day.")
      return
    }
    const targetCrews = Array.from(
      new Map(
        crewGeneratedBlocks
          .filter((b) => {
            const crewCode = normalizeCrewCode(b.crewCode || "")
            if (!crewCode) return false
            if (selectedCrewCode && crewCode !== selectedCrewCode) return false
            const start = (b.startDate || "").trim()
            const end = (b.endDate || "").trim()
            return Boolean(start && end && overlapDaysInclusive(start, end, fromDate, toDate) > 0)
          })
          .map((b) => [
            normalizeCrewCode(b.crewCode || ""),
            { crewCode: normalizeCrewCode(b.crewCode || ""), crewName: b.crewName || "", crewType: b.crewType || "" },
          ]),
      ).values(),
    )
    const blocksByCrew = new Map<string, Entry[]>()
    crewGeneratedBlocks.forEach((b) => {
      const crewCode = normalizeCrewCode(b.crewCode || "")
      if (!crewCode) return
      const arr = blocksByCrew.get(crewCode) || []
      arr.push(b)
      blocksByCrew.set(crewCode, arr)
    })
    const existingRosterByCrew = new Map<string, Entry[]>()
    crewGeneratedRosterBars.forEach((r) => {
      const crewCode = normalizeCrewCode(r.crewCode || "")
      if (!crewCode) return
      const arr = existingRosterByCrew.get(crewCode) || []
      arr.push(r)
      existingRosterByCrew.set(crewCode, arr)
    })
    const actualLeaveByCrew = new Map<string, Entry[]>()
    store.crewLeavePlanner
      .filter((row) => (row.recordType || "").trim().toLowerCase() === "actualleavesegment")
      .forEach((row) => {
        const crewCode = normalizeCrewCode(row.crewCode || "")
        if (!crewCode) return
        const arr = actualLeaveByCrew.get(crewCode) || []
        arr.push(row)
        actualLeaveByCrew.set(crewCode, arr)
      })

    const newRosterRows: Entry[] = []
    targetCrews.forEach((crewRow) => {
      const crewCode = crewRow.crewCode
      const ovStart = fromDate
      const ovEnd = toDate
      const rosterWindowEnd = toDate
      const crewBlocks = blocksByCrew.get(crewCode) || []
      const cycleOneStart =
        crewBlocks
          .filter((blk) => (blk.blockType || "").trim().toLowerCase() === "work")
          .filter((blk) => (blk.cycleNumber || "").trim() === "1")
          .map((blk) => (blk.startDate || "").trim())
          .filter(Boolean)
          .sort()[0] || ovStart
      let cursor = cycleOneStart
      const prevRosterEnd = (existingRosterByCrew.get(crewCode) || [])
        .map((r) => (r.endDate || "").trim())
        .filter((d) => Boolean(d) && d < ovStart)
        .sort()
        .slice(-1)[0]
      if (prevRosterEnd) {
        const nextFromPrev = addDaysToYmd(prevRosterEnd, 1)
        if (nextFromPrev > cursor) cursor = nextFromPrev
      }
      const prevActualLeaveEnd = (actualLeaveByCrew.get(crewCode) || [])
        .map((r) => (r.endDate || "").trim())
        .filter((d) => Boolean(d) && d < ovStart)
        .sort()
        .slice(-1)[0]
      if (prevActualLeaveEnd) {
        const nextFromActualLeave = addDaysToYmd(prevActualLeaveEnd, 1)
        if (nextFromActualLeave > cursor) cursor = nextFromActualLeave
      }
      while (addDaysToYmd(cursor, 6) < ovStart) {
        cursor = addDaysToYmd(cursor, 7)
      }
      let nextRosterDayStartDt: Date | null = null
      if (
        selectedCrewCode &&
        crewRosterGeneratorForm.useFixedRestWeekday &&
        crewRosterGeneratorForm.fixedRestWeekdays.length > 0
      ) {
        const selectedRestDays = new Set(crewRosterGeneratorForm.fixedRestWeekdays)
        let firstRestDate = ovStart
        let guard = 0
        while (guard < 14) {
          if (selectedRestDays.has(weekdayName(firstRestDate))) break
          firstRestDate = addDaysToYmd(firstRestDate, 1)
          guard += 1
        }
        const firstRestStartDt = addHoursToDate(parseLocalDateTime(firstRestDate, "00:00"), -4.5)
        nextRosterDayStartDt = addHoursToDate(firstRestStartDt, -maxRecoveryRestGapHours)
      }
      if (!nextRosterDayStartDt) {
        nextRosterDayStartDt = parseLocalDateTime(cursor, "05:30")
      }
      const rosterWindowEndDt = parseLocalDateTime(ovEnd, "23:59")
      let cycleGuard = 0
      while (nextRosterDayStartDt <= rosterWindowEndDt && cycleGuard < 500) {
        const addRosterSegment = (
          segmentStart: Date,
          segmentEnd: Date,
          rosterType: "REST" | "ROSTER DAY",
          note: string,
        ): Date | null => {
          const segmentMaxEnd = rosterType === "ROSTER DAY" ? ovEnd : rosterWindowEnd
          const windowStartDt = parseLocalDateTime(ovStart, "00:00")
          const windowEndDt = parseLocalDateTime(segmentMaxEnd, "23:59")
          const effectiveStartDt = segmentStart > windowStartDt ? segmentStart : windowStartDt
          const effectiveEndDt = segmentEnd < windowEndDt ? segmentEnd : windowEndDt
          if (effectiveEndDt <= effectiveStartDt) return null
          const durationHours =
            (effectiveEndDt.getTime() - effectiveStartDt.getTime()) / (60 * 60 * 1000)
          if (durationHours < 2 / 60) return null
          const segmentStartMeta = formatDateToYmdHm(effectiveStartDt)
          const segmentEndMeta = formatDateToYmdHm(effectiveEndDt)
          const cycleNumberForSegment = (() => {
            const crewBlocks = blocksByCrew.get(crewCode) || []
            const day = segmentStartMeta.ymd
            const match = crewBlocks.find((blk) => {
              const s = (blk.startDate || "").trim()
              const e = (blk.endDate || "").trim()
              return s && e && day >= s && day <= e
            })
            return match?.cycleNumber || ""
          })()
          newRosterRows.push({
            id: `grb_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`,
            createdAt: new Date().toISOString(),
            recordType: "generatedRoster",
            crewCode,
            crewName: crewRow.crewName || "",
            crewType: crewRow.crewType || "",
            cycleNumber: cycleNumberForSegment,
            startDate: segmentStartMeta.ymd,
            endDate: segmentEndMeta.ymd,
            startTime: segmentStartMeta.hm,
            endTime: segmentEndMeta.hm,
            rosterType,
            sourceBlockId: "",
            maxRecoveryRestGap: crewRosterGeneratorForm.maxRecoveryRestGap,
            restPeriod: `${crewRosterGeneratorForm.restPeriod || "58"}h`,
            localNight: crewRosterGeneratorForm.localNight,
            localDay: crewRosterGeneratorForm.localDay,
            maxTimeBetweenRecoveryRestHours: String(maxRecoveryRestGapHours),
            fdpPattern: note,
          })
          if (rosterType !== "ROSTER DAY") return null
          let lastDutyEndDt: Date | null = null
          getDateRangeInclusive(segmentStartMeta.ymd, segmentEndMeta.ymd).forEach((dayDate) => {
            const dayYmd = toYmd(dayDate)
            let assignedHours = 0
            dutySequenceConfigs.forEach((item) => {
              if (!item) return
              const { code: dutyCode, config: dutyConfig, dutyType } = item
              const dutyStartTime = (dutyConfig.defaultStartTime || "00:00").trim()
              const dutyEndTime = (dutyConfig.defaultEndTime || "23:59").trim()
              const dutyStartDt = parseLocalDateTime(dayYmd, dutyStartTime)
              const dutyEndDt = parseLocalDateTime(
                dutyEndTime <= dutyStartTime && dutyEndTime !== "23:59" ? addDaysToYmd(dayYmd, 1) : dayYmd,
                dutyEndTime,
              )
              const clippedStartDt = dutyStartDt > effectiveStartDt ? dutyStartDt : effectiveStartDt
              const clippedEndDt = dutyEndDt < effectiveEndDt ? dutyEndDt : effectiveEndDt
              if (clippedEndDt <= clippedStartDt) return
              const dutyHours = (clippedEndDt.getTime() - clippedStartDt.getTime()) / (60 * 60 * 1000)
              if (assignedHours + dutyHours > maxDutyTimePerDayHours + 0.01) return
              assignedHours += dutyHours
              if (!lastDutyEndDt || clippedEndDt > lastDutyEndDt) lastDutyEndDt = clippedEndDt
              const dutyStartMeta = formatDateToYmdHm(clippedStartDt)
              const dutyEndMeta = formatDateToYmdHm(clippedEndDt)
              newRosterRows.push({
                id: `grb_duty_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`,
                createdAt: new Date().toISOString(),
                recordType: "generatedRoster",
                crewCode,
                crewName: crewRow.crewName || "",
                crewType: crewRow.crewType || "",
                cycleNumber: cycleNumberForSegment,
                startDate: dutyStartMeta.ymd,
                endDate: dutyEndMeta.ymd,
                startTime: dutyStartMeta.hm,
                endTime: dutyEndMeta.hm,
                rosterType: "DUTY CODE",
                dutyCode,
                dutyType,
                dutyColor: dutyConfig.color || getDefaultDutyCodeColor(dutyCode),
                sourceBlockId: "",
                maxDutyTimePerDayHours: String(maxDutyTimePerDayHours),
                fdpPattern: `Auto sequence ${dutySequenceCodes.join(", ")}`,
              })
            })
          })
          return lastDutyEndDt
        }

        const fdpStart = nextRosterDayStartDt
        const fdpEnd = addHoursToDate(fdpStart, maxRecoveryRestGapHours)
        const lastRosterDutyEndDt =
          addRosterSegment(fdpStart, fdpEnd, "ROSTER DAY", "Day-1 05:30 through Day-5 FDP") || fdpEnd
        const rosterDutyHours =
          (lastRosterDutyEndDt.getTime() - fdpStart.getTime()) / (60 * 60 * 1000)
        const reachedRecoveryLimit = rosterDutyHours >= maxRecoveryRestGapHours - 1
        if (!reachedRecoveryLimit) {
          break
        }

        const mergedRestStart = lastRosterDutyEndDt
        const restPeriodHours = Number(crewRosterGeneratorForm.restPeriod || "58") || 58
        const mergedRestEnd = addHoursToDate(mergedRestStart, restPeriodHours)
        addRosterSegment(
          mergedRestStart,
          mergedRestEnd,
          "REST",
          `Merged REST: ${restPeriodHours} hours from last roster duty end`,
        )

        nextRosterDayStartDt = mergedRestEnd
        cycleGuard += 1
      }
    })
    const targetCrewCodes = new Set(targetCrews.map((c) => c.crewCode))
    setStore((prev) => {
      const kept = prev.crewLeavePlanner.filter((row) => {
        if ((row.recordType || "").trim().toLowerCase() !== "generatedroster") return true
        const crewCode = normalizeCrewCode(row.crewCode || "")
        if (selectedCrewCode && crewCode !== selectedCrewCode) return true
        if (!selectedCrewCode && !targetCrewCodes.has(crewCode)) return true
        const start = (row.startDate || "").trim()
        const end = (row.endDate || "").trim()
        if (!start || !end) return true
        const overlapsWindow = overlapDaysInclusive(start, end, fromDate, toDate) > 0
        return !overlapsWindow
      })
      return { ...prev, crewLeavePlanner: [...newRosterRows, ...kept] }
    })
    setCrewLeaveConflictConfig((p) => ({
      ...p,
      fromDate: "",
      toDate: "",
      crewTypeFilter: "",
      severityFilter: "all",
    }))
    setCrewLeavePlannerTab("timeline")
    setIsCrewRosterGeneratorOpen(false)
    window.alert(`Generated ${newRosterRows.length} roster bar segment(s).`)
  }

  const clearCrewRosterBars = () => {
    const fromDate = (crewRosterGeneratorForm.fromDate || "").trim()
    const toDate = (crewRosterGeneratorForm.toDate || "").trim()
    if (!fromDate || !toDate || toDate < fromDate) {
      window.alert("Select a valid roster period to clear.")
      return
    }
    const selectedCrewCode = normalizeCrewCode(crewRosterGeneratorForm.crewCode || "")
    const crewLabel = selectedCrewCode || "all selected crew"
    if (!window.confirm(`Clear generated roster bars for ${crewLabel} from ${formatYmdToDdMmYy(fromDate)} to ${formatYmdToDdMmYy(toDate)}?`)) {
      return
    }
    let removedCount = 0
    setStore((prev) => {
      const nextCrewLeavePlanner = prev.crewLeavePlanner.filter((row) => {
        if ((row.recordType || "").trim().toLowerCase() !== "generatedroster") return true
        const crewCode = normalizeCrewCode(row.crewCode || "")
        if (selectedCrewCode && crewCode !== selectedCrewCode) return true
        const start = (row.startDate || "").trim()
        const end = (row.endDate || start).trim()
        if (!start || !end) return true
        const overlapsWindow = overlapDaysInclusive(start, end, fromDate, toDate) > 0
        if (overlapsWindow) removedCount += 1
        return !overlapsWindow
      })
      return { ...prev, crewLeavePlanner: nextCrewLeavePlanner }
    })
    setCrewLeavePlannerTab("timeline")
    window.alert(`Cleared ${removedCount} roster bar(s).`)
  }

  const normalizeCrewRosterExcelDate = (value: unknown): string => {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return toYmd(value)
    if (typeof value === "number" && value > 20000 && value < 80000) {
      const date = new Date(1899, 11, 30)
      date.setDate(date.getDate() + Math.floor(value))
      return toYmd(date)
    }
    return normalizeImportDate(String(value || ""))
  }

  const normalizeImportedCrewDutyCode = (value: unknown): string => {
    const raw = String(value || "").trim().toUpperCase().replace(/\s+/g, "")
    if (raw === "O" || raw === "OFF" || raw === "REST") return "OF"
    if (raw === "PL" || raw === "AL" || raw === "NP" || raw === "LEAVE") return "AL"
    return raw
  }

  const importCrewRosterExcel = async (file: File) => {
    try {
      await loadSheetJs()
    } catch (error) {
      window.alert(error instanceof Error ? error.message : "Failed to load Excel parser.")
      return
    }
    const XLSX = (window as SheetJSWindow).XLSX
    if (!XLSX) {
      window.alert("Excel parser is not available.")
      return
    }

    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true })
    const dutyCodeDefaults: Record<string, { name: string; dutyType: string; startTime: string; endTime: string }> = {
      FD: { name: "Flight Duty", dutyType: "Flight Duty", startTime: "00:00", endTime: "23:59" },
      GD: { name: "Ground Duty", dutyType: "Ground Duty", startTime: "00:00", endTime: "23:59" },
      OF: { name: "Off / Release", dutyType: "Release - OF", startTime: "19:30", endTime: "23:59" },
    }
    const existingDutyCodeMap = new Map<string, Entry>()
    crewDutyCodeConfigs.forEach((row) => {
      const code = (row.code || "").trim().toUpperCase()
      if (code) existingDutyCodeMap.set(code, row)
    })
    const allowedDutyCodes = new Set(Array.from(existingDutyCodeMap.keys()))
    const crewDbByCode = new Map<string, Entry>()
    store.crewDataBase.forEach((row) => {
      const code = normalizeCrewCode(row.crewCode || "")
      if (code) crewDbByCode.set(code, row)
    })
    const generatedBlocksByCrew = new Map<string, Entry[]>()
    crewGeneratedBlocks.forEach((row) => {
      const code = normalizeCrewCode(row.crewCode || "")
      if (!code) return
      const list = generatedBlocksByCrew.get(code) || []
      list.push(row)
      generatedBlocksByCrew.set(code, list)
    })
    const importedRows: Entry[] = []
    const importedActualLeaveRows: Entry[] = []
    const affectedCrewCodes = new Set<string>()
    const importedDates = new Set<string>()
    const skippedCodes = new Map<string, number>()
    const skippedOutsideRosterDay = new Map<string, number>()
    const restPeriodHours = Number(crewRosterGeneratorForm.restPeriod || "58") || 58
    const maxRecoveryRestGapHours = Number(crewRosterGeneratorForm.maxRecoveryRestGap || "110") || 110

    type ImportedRosterCell = {
      crewCode: string
      crewName: string
      crewType: string
      sheetName: string
      ymd: string
      code: string
      dutyType: string
      rosterBarMap: "REST" | "ROSTER DAY"
      startTime: string
      endTime: string
      color: string
      startDt: Date
      endDt: Date
    }

    type ImportedActualLeaveCell = {
      crewCode: string
      crewName: string
      crewType: string
      sheetName: string
      ymd: string
    }

    type ImportedBaseSegment = {
      crewCode: string
      crewName: string
      crewType: string
      sheetName: string
      rosterType: "REST" | "ROSTER DAY"
      startDt: Date
      endDt: Date
      dutyCode?: string
      dutyType?: string
      dutyColor?: string
      note: string
    }

    const findGeneratedBlockForDay = (crewCode: string, ymd: string) =>
      (generatedBlocksByCrew.get(crewCode) || []).find((block) => {
        const start = (block.startDate || "").trim()
        const end = (block.endDate || "").trim()
        return start && end && ymd >= start && ymd <= end
      })

    const pushImportedBaseSegment = (segment: ImportedBaseSegment) => {
      if (segment.endDt <= segment.startDt) return
      const startMeta = formatDateToYmdHm(segment.startDt)
      const endMeta = formatDateToYmdHm(segment.endDt)
      const matchingBlock = findGeneratedBlockForDay(segment.crewCode, startMeta.ymd)
      affectedCrewCodes.add(segment.crewCode)
      importedDates.add(startMeta.ymd)
      importedDates.add(endMeta.ymd)
      importedRows.push({
        id: `grb_imp_base_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`,
        createdAt: new Date().toISOString(),
        recordType: "generatedRoster",
        crewCode: segment.crewCode,
        crewName: segment.crewName,
        crewType: segment.crewType,
        cycleNumber: matchingBlock?.cycleNumber || "",
        startDate: startMeta.ymd,
        endDate: endMeta.ymd,
        startTime: startMeta.hm,
        endTime: endMeta.hm,
        rosterType: segment.rosterType,
        dutyCode: segment.dutyCode || "",
        dutyType: segment.dutyType || "",
        dutyColor: segment.dutyColor || "",
        sourceBlockId: matchingBlock?.id || "",
        maxRecoveryRestGap: String(maxRecoveryRestGapHours),
        restPeriod: `${restPeriodHours}h`,
        localNight: crewRosterGeneratorForm.localNight,
        localDay: crewRosterGeneratorForm.localDay,
        maxTimeBetweenRecoveryRestHours: String(maxRecoveryRestGapHours),
        fdpPattern: segment.note,
      })
    }

    workbook.SheetNames.forEach((sheetName) => {
      const ws = workbook.Sheets[sheetName]
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" })
      if (rows.length < 3) return
      const dateRow = rows[0] || []
      const dateColumns = dateRow
        .map((value, idx) => ({ idx, ymd: normalizeCrewRosterExcelDate(value) }))
        .filter((col) => col.idx >= 3 && col.ymd)
      if (dateColumns.length === 0) return

      rows.slice(2).forEach((rowValues) => {
        const crewCode = normalizeCrewCode(String(rowValues[2] || ""))
        if (!crewCode) return
        const crewDb = crewDbByCode.get(crewCode)
        const crewName = (crewDb?.crewName || String(rowValues[0] || "")).trim()
        const crewType = (crewDb?.crewType || sheetName || "").trim()
        const dutyCells: ImportedRosterCell[] = []
        const actualLeaveCells: ImportedActualLeaveCell[] = []

        dateColumns.forEach(({ idx, ymd }) => {
          const code = normalizeImportedCrewDutyCode(rowValues[idx])
          if (!code) return
          if (code === "AL") {
            actualLeaveCells.push({ crewCode, crewName, crewType, sheetName, ymd })
            return
          }
          if (!allowedDutyCodes.has(code)) {
            skippedCodes.set(code, (skippedCodes.get(code) || 0) + 1)
            return
          }
          const configured = existingDutyCodeMap.get(code)
          if (!configured) return
          const fallback = dutyCodeDefaults[code] || {
            name: configured.dutyName || code,
            dutyType: configured.dutyType || "Ground Duty",
            startTime: configured.defaultStartTime || "00:00",
            endTime: configured.defaultEndTime || "23:59",
          }
          const dutyType = configured?.dutyType || fallback.dutyType
          const rosterBarMap =
            crewDutyTypeMap.get(dutyType.trim().toLowerCase()) ||
            configured?.rosterBarMap ||
            getCrewDutyRosterBarMap(dutyType)
          const startTime =
            rosterBarMap === "REST" && (!configured?.defaultStartTime || configured.defaultStartTime === "00:00")
              ? fallback.startTime
              : configured?.defaultStartTime || fallback.startTime
          const endTime = configured?.defaultEndTime || fallback.endTime
          const dutyStartDt = parseLocalDateTime(ymd, startTime || "00:00")
          const dutyEndDt = parseLocalDateTime(
            startTime && endTime && endTime <= startTime && endTime !== "23:59" ? addDaysToYmd(ymd, 1) : ymd,
            endTime || "23:59",
          )
          dutyCells.push({
            crewCode,
            crewName,
            crewType,
            sheetName,
            ymd,
            code,
            dutyType,
            rosterBarMap: rosterBarMap === "REST" ? "REST" : "ROSTER DAY",
            startTime: startTime || "00:00",
            endTime: endTime || "23:59",
            color: configured.color || getDefaultDutyCodeColor(code),
            startDt: dutyStartDt,
            endDt: dutyEndDt,
          })
        })

        if (actualLeaveCells.length > 0) {
          const sortedLeaveCells = actualLeaveCells.sort((a, b) => a.ymd.localeCompare(b.ymd))
          let segmentStart = sortedLeaveCells[0].ymd
          let segmentEnd = sortedLeaveCells[0].ymd
          const pushActualLeaveSegment = (startDate: string, endDate: string) => {
            const matchingBlock = findGeneratedBlockForDay(crewCode, startDate)
            affectedCrewCodes.add(crewCode)
            importedDates.add(startDate)
            importedDates.add(endDate)
            importedActualLeaveRows.push({
              id: `als_imp_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`,
              createdAt: new Date().toISOString(),
              recordType: "actualLeaveSegment",
              crewCode,
              crewName,
              crewType,
              cycleNumber: matchingBlock?.cycleNumber || "",
              startDate,
              endDate,
              plannedDays: String(getDateRangeInclusive(startDate, endDate).length),
              sourceBlockId: matchingBlock?.id || "",
              overrideConflicts: "Yes",
              importedCode: "PL",
              fdpPattern: `Imported ${sheetName} PL as AL`,
            })
          }
          sortedLeaveCells.slice(1).forEach((cell) => {
            if (cell.ymd === addDaysToYmd(segmentEnd, 1)) {
              segmentEnd = cell.ymd
              return
            }
            pushActualLeaveSegment(segmentStart, segmentEnd)
            segmentStart = cell.ymd
            segmentEnd = cell.ymd
          })
          pushActualLeaveSegment(segmentStart, segmentEnd)
        }

        if (dutyCells.length === 0) return

        dutyCells.sort((a, b) => a.startDt.getTime() - b.startDt.getTime())
        const restSegments: ImportedBaseSegment[] = []
        const rosterDayCells = dutyCells.filter((cell) => cell.rosterBarMap !== "REST")
        const restCells = dutyCells.filter((cell) => cell.rosterBarMap === "REST")

        restCells.forEach((restCell) => {
          const previousCell = dutyCells
            .filter((cell) => cell.startDt < restCell.startDt)
            .sort((a, b) => b.startDt.getTime() - a.startDt.getTime())[0]
          if (previousCell?.rosterBarMap === "REST") return
          const previousDutyCell = dutyCells
            .filter((cell) => cell.rosterBarMap !== "REST" && cell.startDt < restCell.startDt)
            .sort((a, b) => b.startDt.getTime() - a.startDt.getTime())[0]
          const restStartDt = previousDutyCell ? previousDutyCell.endDt : restCell.startDt
          const restEndDt = addHoursToDate(restStartDt, restPeriodHours)
          restSegments.push({
            crewCode,
            crewName,
            crewType,
            sheetName,
            rosterType: "REST",
            startDt: restStartDt,
            endDt: restEndDt,
            dutyCode: restCell.code,
            dutyType: restCell.dutyType,
            dutyColor: restCell.color,
            note: `Imported ${sheetName} ${restCell.code}: ${restPeriodHours}h recovery rest`,
          })
        })

        restSegments.sort((a, b) => a.startDt.getTime() - b.startDt.getTime())
        restSegments.forEach(pushImportedBaseSegment)

        const rosterDaySegments: ImportedBaseSegment[] = []
        const pushRosterDaySegment = (startDt: Date, requestedEndDt: Date, note: string) => {
          const maxEndDt = addHoursToDate(startDt, maxRecoveryRestGapHours)
          const endDt = requestedEndDt < maxEndDt ? requestedEndDt : maxEndDt
          if (endDt <= startDt) return
          rosterDaySegments.push({
            crewCode,
            crewName,
            crewType,
            sheetName,
            rosterType: "ROSTER DAY",
            startDt,
            endDt,
            note,
          })
        }

        if (restSegments.length > 0) {
          const firstRest = restSegments[0]
          const dutiesBeforeFirstRest = rosterDayCells.filter((cell) => cell.startDt < firstRest.startDt)
          if (dutiesBeforeFirstRest.length > 0) {
            pushRosterDaySegment(
              dutiesBeforeFirstRest[0].startDt,
              firstRest.startDt,
              `Imported ${sheetName}: duty period until ${firstRest.dutyCode || "OF"}`,
            )
          }
          restSegments.forEach((restSegment, index) => {
            const nextRest = restSegments[index + 1]
            if (nextRest) {
              pushRosterDaySegment(
                restSegment.endDt,
                nextRest.startDt,
                `Imported ${sheetName}: roster day between recovery rest periods`,
              )
              return
            }
            const dutiesAfterRest = rosterDayCells.filter((cell) => cell.endDt > restSegment.endDt)
            if (dutiesAfterRest.length > 0) {
              const lastDutyEnd = dutiesAfterRest.reduce(
                (latest, cell) => (cell.endDt > latest ? cell.endDt : latest),
                dutiesAfterRest[0].endDt,
              )
              pushRosterDaySegment(
                restSegment.endDt,
                lastDutyEnd,
                `Imported ${sheetName}: duty period after ${restSegment.dutyCode || "OF"}`,
              )
            }
          })
        } else if (rosterDayCells.length > 0) {
          const firstDutyStart = rosterDayCells[0].startDt
          const lastDutyEnd = rosterDayCells.reduce(
            (latest, cell) => (cell.endDt > latest ? cell.endDt : latest),
            rosterDayCells[0].endDt,
          )
          pushRosterDaySegment(firstDutyStart, lastDutyEnd, `Imported ${sheetName}: roster day from duty codes`)
        }

        rosterDaySegments
          .sort((a, b) => a.startDt.getTime() - b.startDt.getTime())
          .forEach(pushImportedBaseSegment)

        rosterDayCells.forEach((cell) => {
          const matchingRosterDay = rosterDaySegments.find(
            (segment) => cell.startDt < segment.endDt && cell.endDt > segment.startDt,
          )
          if (!matchingRosterDay) {
            skippedOutsideRosterDay.set(cell.code, (skippedOutsideRosterDay.get(cell.code) || 0) + 1)
            return
          }
          const clippedStartDt = cell.startDt > matchingRosterDay.startDt ? cell.startDt : matchingRosterDay.startDt
          const clippedEndDt = cell.endDt < matchingRosterDay.endDt ? cell.endDt : matchingRosterDay.endDt
          if (clippedEndDt <= clippedStartDt) {
            skippedOutsideRosterDay.set(cell.code, (skippedOutsideRosterDay.get(cell.code) || 0) + 1)
            return
          }
          const startMeta = formatDateToYmdHm(clippedStartDt)
          const endMeta = formatDateToYmdHm(clippedEndDt)
          const matchingBlock = findGeneratedBlockForDay(crewCode, startMeta.ymd)
          affectedCrewCodes.add(crewCode)
          importedDates.add(startMeta.ymd)
          importedDates.add(endMeta.ymd)
          importedRows.push({
            id: `grb_imp_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`,
            createdAt: new Date().toISOString(),
            recordType: "generatedRoster",
            crewCode,
            crewName,
            crewType,
            cycleNumber: matchingBlock?.cycleNumber || "",
            startDate: startMeta.ymd,
            endDate: endMeta.ymd,
            startTime: startMeta.hm,
            endTime: endMeta.hm,
            rosterType: "DUTY CODE",
            dutyCode: cell.code,
            dutyType: cell.dutyType,
            dutyColor: cell.color,
            sourceBlockId: matchingBlock?.id || "",
            fdpPattern: `Imported ${cell.sheetName} ${cell.code}`,
          })
        })
      })
    })

    const allImportedRows = [...importedRows, ...importedActualLeaveRows]

    if (allImportedRows.length === 0) {
      const outsideRosterText = Array.from(skippedOutsideRosterDay.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8)
        .map(([code, count]) => `${code}: ${count}`)
        .join(", ")
      window.alert(
        outsideRosterText
          ? `No duty codes were imported because they did not fit inside imported Roster Day periods.\nSkipped: ${outsideRosterText}`
          : "No matching duty code or PL/AL/NP cells found in the selected Excel file. Check that the codes exist in Duty Code Manager.",
      )
      return
    }

    const sortedDates = Array.from(importedDates).sort()
    const fromDate = sortedDates[0]
    const toDate = sortedDates[sortedDates.length - 1]
    const hasImportedRosterRows = importedRows.length > 0
    setStore((prev) => {
      const kept = prev.crewLeavePlanner.filter((row) => {
        const recordType = (row.recordType || "").trim().toLowerCase()
        const crewCode = normalizeCrewCode(row.crewCode || "")
        if (!affectedCrewCodes.has(crewCode)) return true
        const start = (row.startDate || "").trim()
        const end = (row.endDate || start).trim()
        const overlapsImport = Boolean(start && end && overlapDaysInclusive(start, end, fromDate, toDate) > 0)
        if (!overlapsImport) return true
        if (recordType === "generatedroster") {
          if (hasImportedRosterRows) return false
          return (row.rosterType || "").trim().toUpperCase() !== "DUTY CODE"
        }
        if (recordType === "actualleavesegment") return false
        return true
      })
      return {
        ...prev,
        crewLeavePlanner: [...allImportedRows, ...kept],
      }
    })
    setCrewLeaveTimelineFilter(getCurrentYearRange())
    setCrewLeavePlannerTab("timeline")

    const skippedText = Array.from(skippedCodes.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)
      .map(([code, count]) => `${code || "(blank)"}: ${count}`)
      .join(", ")
    const outsideRosterText = Array.from(skippedOutsideRosterDay.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)
      .map(([code, count]) => `${code}: ${count}`)
      .join(", ")
    window.alert(
      `Crew roster import complete.\nImported roster cells: ${importedRows.length}\nImported PL/AL/NP leave segment(s): ${importedActualLeaveRows.length}\nCrew affected: ${affectedCrewCodes.size}\nPeriod: ${formatYmdToDdMmYy(
        fromDate,
      )} - ${formatYmdToDdMmYy(toDate)}${
        skippedText ? `\nSkipped other codes: ${skippedText}` : ""
      }${
        outsideRosterText
          ? `\nSkipped outside Roster Day / overlapping Rest: ${outsideRosterText}`
          : ""
      }`,
    )
  }

  const openRosterGeneratorForCrew = (crewCodeValue: string, fromDateValue?: string) => {
    const crewCode = normalizeCrewCode(crewCodeValue || "")
    const fromDate = (fromDateValue || "").trim()
    const range = getSixMonthRangeFromDate(fromDate || undefined)
    setCrewRosterGeneratorForm((prev) => ({
      ...prev,
      crewCode,
      fromDate: fromDate || range.fromDate,
      toDate: range.toDate,
    }))
    setCrewRosterGeneratorTab("setup")
    setIsCrewRosterGeneratorOpen(true)
  }

  const openActualLeaveRosterDetails = (rb: Entry) => {
    const crewCode = normalizeCrewCode(rb.crewCode || "")
    const cycleNumber = (rb.cycleNumber || "").trim()
    const sourceBlockId = (rb.sourceBlockId || "").trim()
    const matchingLeaveBlock =
      (sourceBlockId
        ? computedCrewGeneratedBlocks.find((row) => (row.id || "").trim() === sourceBlockId)
        : undefined) ||
      (cycleNumber
        ? computedCrewGeneratedBlocks.find(
            (row) =>
              normalizeCrewCode(row.crewCode || "") === crewCode &&
              (row.blockType || "").trim().toLowerCase() === "leave" &&
              (row.cycleNumber || "").trim() === cycleNumber,
          )
        : undefined) ||
      computedCrewGeneratedBlocks.find((row) => {
        if (normalizeCrewCode(row.crewCode || "") !== crewCode) return false
        if ((row.blockType || "").trim().toLowerCase() !== "leave") return false
        const start = (row.startDate || "").trim()
        const end = (row.endDate || start).trim()
        const alStart = (rb.startDate || "").trim()
        const alEnd = (rb.endDate || alStart).trim()
        return Boolean(start && end && alStart && alEnd && overlapDaysInclusive(start, end, alStart, alEnd) > 0)
      })
    const nextCycle = cycleNumber || (matchingLeaveBlock?.cycleNumber || "").trim()
    setLeaveDetailModal({
      open: true,
      blockId: matchingLeaveBlock?.id || "",
      actualLeaveSegmentId: (rb.id || "").trim(),
      crewCode,
      crewName: rb.crewName || matchingLeaveBlock?.crewName || "",
      cycleNumber: nextCycle,
      leaveStart: matchingLeaveBlock?.startDate || (rb.startDate || ""),
      leaveEnd: matchingLeaveBlock?.endDate || (rb.endDate || ""),
      policyName: matchingLeaveBlock?.policyName || "",
    })
    setLeaveDetailSelectedCycle(nextCycle)
    setLeaveDetailDraft({
      start: (rb.startDate || "").trim(),
      end: (rb.endDate || rb.startDate || "").trim(),
    })
  }

  const openRosterBarDetails = (
    rb: Entry,
    labelType: string,
    selectedStartDt?: Date,
    selectedEndDt?: Date,
  ) => {
    const start = (rb.startDate || "").trim()
    const end = (rb.endDate || "").trim()
    const startHm = (rb.startTime || "00:00").trim()
    const endHm = (rb.endTime || "23:59").trim()
    const barStartDt = parseLocalDateTime(start, startHm || "00:00")
    const barEndDt = parseLocalDateTime(end, endHm || "23:59")
    const detailStartDt = selectedStartDt && selectedStartDt > barStartDt ? selectedStartDt : barStartDt
    const detailEndDt = selectedEndDt && selectedEndDt < barEndDt ? selectedEndDt : barEndDt
    if (detailEndDt <= detailStartDt) return
    const startMeta = formatDateToYmdHm(detailStartDt)
    const endMeta = formatDateToYmdHm(detailEndDt)
    const durationHours = (detailEndDt.getTime() - detailStartDt.getTime()) / (60 * 60 * 1000)
    const crewCode = normalizeCrewCode(rb.crewCode || "")
    const dutyAssignments = crewGeneratedRosterBars
      .filter((row) => normalizeCrewCode(row.crewCode || "") === crewCode)
      .filter((row) => (row.rosterType || "").trim().toUpperCase() === "DUTY CODE")
      .map((row) => {
        const dutyStart = parseLocalDateTime(
          (row.startDate || "").trim(),
          (row.startTime || "00:00").trim() || "00:00",
        )
        const dutyEnd = parseLocalDateTime(
          (row.endDate || row.startDate || "").trim(),
          (row.endTime || "23:59").trim() || "23:59",
        )
        return { row, dutyStart, dutyEnd }
      })
      .filter(({ dutyStart, dutyEnd }) => dutyStart < detailEndDt && dutyEnd > detailStartDt)
      .sort((a, b) => a.dutyStart.getTime() - b.dutyStart.getTime())
      .map(({ row, dutyStart, dutyEnd }) => {
        const dutyStartMeta = formatDateToYmdHm(dutyStart)
        const dutyEndMeta = formatDateToYmdHm(dutyEnd)
        return {
          id: row.id,
          code: (row.dutyCode || "").trim().toUpperCase(),
          dutyType: row.dutyType || "",
          startDate: dutyStartMeta.ymd,
          startTime: dutyStartMeta.hm,
          endDate: dutyEndMeta.ymd,
          endTime: dutyEndMeta.hm,
        }
      })
    setRosterBarDetailModal({
      open: true,
      crewCode,
      crewName: rb.crewName || "",
      rosterType: rb.dutyCode ? `${String(rb.dutyCode).trim().toUpperCase()} - ${labelType}` : labelType,
      startDate: startMeta.ymd,
      startTime: startMeta.hm,
      endDate: endMeta.ymd,
      endTime: endMeta.hm,
      durationHours: durationHours.toFixed(1),
      daySegments: buildRosterDaySegments(detailStartDt, detailEndDt, labelType),
      dutyAssignments,
    })
  }

  const updateRosterDetailDutyAssignment = (
    index: number,
    field: "startDate" | "startTime" | "endDate" | "endTime",
    value: string,
  ) => {
    setRosterBarDetailModal((prev) => {
      const dutyAssignments = prev.dutyAssignments.map((row) => ({ ...row }))
      if (!dutyAssignments[index]) return prev
      dutyAssignments[index][field] = value
      if ((field === "endDate" || field === "endTime") && dutyAssignments[index + 1]) {
        dutyAssignments[index + 1].startDate = dutyAssignments[index].endDate
        dutyAssignments[index + 1].startTime = dutyAssignments[index].endTime
      }
      return { ...prev, dutyAssignments }
    })
  }

  const saveRosterDetailDutyAssignments = () => {
    if (!rosterBarDetailModal.open || rosterBarDetailModal.dutyAssignments.length === 0) return
    const rows = rosterBarDetailModal.dutyAssignments
    for (let i = 0; i < rows.length; i += 1) {
      const row = rows[i]
      if (!row.startDate || !row.startTime || !row.endDate || !row.endTime) {
        window.alert("Start and End date/time are required for all duty codes.")
        return
      }
      const startDt = parseLocalDateTime(row.startDate, row.startTime)
      const endDt = parseLocalDateTime(row.endDate, row.endTime)
      if (endDt <= startDt) {
        window.alert(`Duty code ${row.code || i + 1} end time must be after start time.`)
        return
      }
      const nextRow = rows[i + 1]
      if (nextRow) {
        const nextStartDt = parseLocalDateTime(nextRow.startDate, nextRow.startTime)
        if (nextStartDt < endDt) {
          window.alert(`Duty code ${nextRow.code || i + 2} starts before the previous duty ends.`)
          return
        }
      }
    }
    const updatesById = new Map(rows.map((row) => [row.id, row]))
    setStore((prev) => ({
      ...prev,
      crewLeavePlanner: prev.crewLeavePlanner.map((row) => {
        const update = updatesById.get(row.id)
        if (!update) return row
        return {
          ...row,
          startDate: update.startDate,
          startTime: update.startTime,
          endDate: update.endDate,
          endTime: update.endTime,
        }
      }),
    }))
    window.alert("Duty code timing updated.")
  }

  const crewLeaveBalances = useMemo(() => {
    const policyById = new Map<string, Entry>()
    crewLeavePolicies.forEach((p) => policyById.set(p.id, p))
    const assignmentByCrewCode = new Map<string, Entry>()
    crewLeaveAssignments.forEach((a) => assignmentByCrewCode.set(normalizeCrewCode(a.crewCode || ""), a))
    const marksByCrewCode = new Map<string, Entry[]>()
    crewLeaveMarks.forEach((m) => {
      const code = normalizeCrewCode(m.crewCode || "")
      const arr = marksByCrewCode.get(code) || []
      arr.push(m)
      marksByCrewCode.set(code, arr)
    })

    return store.crewDataBase.map((crew) => {
      const crewCode = normalizeCrewCode(crew.crewCode || "")
      const crewType = (crew.crewType || "").trim()
      const releaseDate = (crew.releaseDate || "").trim()
      const assignment = assignmentByCrewCode.get(crewCode)
      const policy =
        (assignment?.policyId ? policyById.get(assignment.policyId) : undefined) ||
        crewLeavePolicies.find(
          (p) =>
            (p.activeStatus || "").trim() === "Active" &&
            (p.crewType || "").trim() === crewType,
        )
      const workRotation = Number(policy?.workPatternDays || "0")
      const leaveRotation = Number(policy?.leavePatternDays || "0")
      const cycleDays = workRotation + leaveRotation
      const deductions = (marksByCrewCode.get(crewCode) || []).reduce((acc, mark) => {
        const d = Number(mark.noOfDays || "0")
        return acc + (Number.isFinite(d) ? d : 0)
      }, 0)
      const leaveBalance = Math.max(leaveRotation - deductions, 0)
      const renewalDate =
        releaseDate && cycleDays > 0
          ? toYmd(new Date(new Date(`${releaseDate}T00:00:00`).getTime() + cycleDays * 24 * 60 * 60 * 1000))
          : ""
      return {
        id: `bal_${crew.id}`,
        crewCode,
        crewName: crew.crewName || "",
        staffId: crew.employeeNo || "",
        crewType,
        releaseDate,
        policyName: policy?.policyName || "-",
        workRotation: workRotation > 0 ? String(workRotation) : "-",
        leaveRotation: leaveRotation > 0 ? String(leaveRotation) : "-",
        leaveDeduction: String(deductions),
        leaveBalance: String(leaveBalance),
        renewalDate: renewalDate || "-",
      }
    })
  }, [crewLeaveAssignments, crewLeaveMarks, crewLeavePolicies, store.crewDataBase])

  const editEntry = (moduleKey: ModuleKey, entry: Entry) => {
    const config = MODULES.find((item) => item.key === moduleKey)
    if (!config) return
    const nextForm = buildEmptyForm(config)
    config.fields.forEach((field) => {
      nextForm[field.key] = entry[field.key] ?? nextForm[field.key] ?? ""
    })

    setActiveModule(moduleKey)
    setForms((prev) => ({
      ...prev,
      [moduleKey]: nextForm,
    }))
    if (moduleKey === "leaves") {
      const staffLabel = entry.staffLookup || `${entry.staffNo || ""} - ${entry.staffName || ""}`.trim()
      setStaffLookupQuery(staffLabel)
    } else {
      setStaffLookupQuery("")
    }
    setEditingEntryId(entry.id)
    setIsFormOpen(true)
  }

  const findActualLeaveDropConflicts = (params: {
    crewCode: string
    startDate: string
    endDate: string
    sourceBlockId?: string
  }) => {
    const crewCode = normalizeCrewCode(params.crewCode || "")
    const startDate = (params.startDate || "").trim()
    const endDate = (params.endDate || "").trim()
    if (!crewCode || !startDate || !endDate) return []
    const generatedConflicts = crewGeneratedRosterBars
      .filter((row) => normalizeCrewCode(row.crewCode || "") === crewCode)
      .filter((row) => {
        const start = (row.startDate || "").trim()
        const end = (row.endDate || start).trim()
        return Boolean(start && end && overlapDaysInclusive(startDate, endDate, start, end) > 0)
      })
      .map((row) => ({
        type: (row.rosterType || "Roster").trim().toUpperCase(),
        code: (row.dutyCode || "").trim().toUpperCase(),
        start: (row.startDate || "").trim(),
        end: (row.endDate || row.startDate || "").trim(),
        startTime: (row.startTime || "00:00").trim(),
        endTime: (row.endTime || "23:59").trim(),
      }))
    const actualLeaveConflicts = store.crewLeavePlanner
      .filter((row) => (row.recordType || "").trim().toLowerCase() === "actualleavesegment")
      .filter((row) => normalizeCrewCode(row.crewCode || "") === crewCode)
      .filter((row) => (row.sourceBlockId || "").trim() !== (params.sourceBlockId || "").trim())
      .filter((row) => {
        const start = (row.startDate || "").trim()
        const end = (row.endDate || start).trim()
        return Boolean(start && end && overlapDaysInclusive(startDate, endDate, start, end) > 0)
      })
      .map((row) => ({
        type: "ACTUAL LEAVE",
        code: "",
        start: (row.startDate || "").trim(),
        end: (row.endDate || row.startDate || "").trim(),
        startTime: "00:00",
        endTime: "23:59",
      }))
    return [...generatedConflicts, ...actualLeaveConflicts]
  }

  const openActualLeaveDropModal = (params: {
    block: Entry
    startDate: string
    endDate: string
  }) => {
    const crewCode = normalizeCrewCode(params.block.crewCode || "")
    const sourceBlockId = (params.block.id || "").trim()
    const startDate = (params.startDate || "").trim()
    const endDate = (params.endDate || "").trim()
    if (!crewCode || !startDate || !endDate || endDate < startDate) return
    setPendingActualLeaveDrop({
      open: true,
      sourceBlockId,
      crewCode,
      crewName: params.block.crewName || "",
      cycleNumber: (params.block.cycleNumber || "").trim(),
      startDate,
      endDate,
      plannedDays: getDateRangeInclusive(startDate, endDate).length,
      conflicts: findActualLeaveDropConflicts({ crewCode, startDate, endDate, sourceBlockId }),
      override: false,
    })
  }

  const confirmActualLeaveDrop = () => {
    if (!pendingActualLeaveDrop.open) return
    if (!pendingActualLeaveDrop.startDate || !pendingActualLeaveDrop.endDate) {
      window.alert("Actual Leave Start and End are required.")
      return
    }
    if (pendingActualLeaveDrop.endDate < pendingActualLeaveDrop.startDate) {
      window.alert("Actual Leave End must be after or equal to Start.")
      return
    }
    if (pendingActualLeaveDrop.conflicts.length > 0 && !pendingActualLeaveDrop.override) {
      window.alert("Conflict found. Tick Override to save this Actual Leave.")
      return
    }
    addActualLeaveSegment({
      crewCode: pendingActualLeaveDrop.crewCode,
      crewName: pendingActualLeaveDrop.crewName,
      cycleNumber: pendingActualLeaveDrop.cycleNumber,
      startDate: pendingActualLeaveDrop.startDate,
      endDate: pendingActualLeaveDrop.endDate,
      sourceBlockId: pendingActualLeaveDrop.sourceBlockId,
      overrideConflicts: pendingActualLeaveDrop.override ? "Yes" : "No",
      clearConflictingDutyCodes: pendingActualLeaveDrop.override,
    })
    setPendingActualLeaveDrop((p) => ({ ...p, open: false }))
  }

  const openCrewDutyDropModal = (params: {
    code: string
    crewCode: string
    crewName: string
    date: string
  }) => {
    const code = (params.code || "").trim().toUpperCase()
    const crewCode = normalizeCrewCode(params.crewCode || "")
    const date = (params.date || "").trim()
    if (!code || !crewCode || !date) return
    const config = crewDutyCodeConfigs.find((row) => (row.code || "").trim().toUpperCase() === code)
    if (!config) {
      window.alert(`Duty code ${code} not found in Duty Code Manager.`)
      return
    }
    const dutyType = (config.dutyType || "").trim()
    const rosterBarMap =
      crewDutyTypeMap.get(dutyType.toLowerCase()) ||
      config.rosterBarMap ||
      getCrewDutyRosterBarMap(dutyType)
    const defaultStart =
      rosterBarMap === "REST" && (!config.defaultStartTime || config.defaultStartTime === "00:00")
        ? "19:30"
        : (config.defaultStartTime || "00:00").trim()
    const defaultEnd = (config.defaultEndTime || "23:59").trim()
    const startDt = parseLocalDateTime(date, defaultStart)
    const endDt =
      rosterBarMap === "REST"
        ? addHoursToDate(startDt, Number(crewRosterGeneratorForm.restPeriod || "58") || 58)
        : parseLocalDateTime(defaultEnd <= defaultStart && defaultEnd !== "23:59" ? addDaysToYmd(date, 1) : date, defaultEnd)
    const startMeta = formatDateToYmdHm(startDt)
    const endMeta = formatDateToYmdHm(endDt)
    const conflicts = crewGeneratedRosterBars
      .filter((row) => normalizeCrewCode(row.crewCode || "") === crewCode)
      .filter((row) => {
        const type = (row.rosterType || "").trim().toUpperCase()
        if (type !== "DUTY CODE" && type !== "REST") return false
        const rowStart = parseLocalDateTime((row.startDate || "").trim(), (row.startTime || "00:00").trim())
        const rowEnd = parseLocalDateTime((row.endDate || row.startDate || "").trim(), (row.endTime || "23:59").trim())
        return startDt < rowEnd && endDt > rowStart
      })
      .map((row) => ({
        id: row.id,
        type: (row.rosterType || "").trim().toUpperCase(),
        code: (row.dutyCode || "").trim().toUpperCase(),
        start: (row.startDate || "").trim(),
        end: (row.endDate || row.startDate || "").trim(),
        startTime: (row.startTime || "00:00").trim(),
        endTime: (row.endTime || "23:59").trim(),
      }))
    setPendingCrewDutyDrop({
      open: true,
      crewCode,
      crewName: params.crewName || "",
      code,
      dutyType,
      rosterBarMap: rosterBarMap === "REST" ? "REST" : "ROSTER DAY",
      startDate: startMeta.ymd,
      startTime: startMeta.hm,
      endDate: endMeta.ymd,
      endTime: endMeta.hm,
      conflicts,
      mode: conflicts.length ? "after" : "override",
    })
  }

  const confirmCrewDutyDrop = () => {
    if (!pendingCrewDutyDrop.open) return
    const config = crewDutyCodeConfigs.find(
      (row) => (row.code || "").trim().toUpperCase() === pendingCrewDutyDrop.code,
    )
    if (!config) return
    let startDt = parseLocalDateTime(pendingCrewDutyDrop.startDate, pendingCrewDutyDrop.startTime || "00:00")
    if (pendingCrewDutyDrop.mode === "after" && pendingCrewDutyDrop.conflicts.length > 0) {
      const latestEnd = pendingCrewDutyDrop.conflicts
        .map((c) => parseLocalDateTime(c.end, c.endTime || "23:59"))
        .sort((a, b) => b.getTime() - a.getTime())[0]
      if (latestEnd) startDt = latestEnd
    }
    const endDt =
      pendingCrewDutyDrop.rosterBarMap === "REST"
        ? addHoursToDate(startDt, Number(crewRosterGeneratorForm.restPeriod || "58") || 58)
        : parseLocalDateTime(
            pendingCrewDutyDrop.endTime <= pendingCrewDutyDrop.startTime &&
              pendingCrewDutyDrop.endTime !== "23:59"
              ? addDaysToYmd(formatDateToYmdHm(startDt).ymd, 1)
              : formatDateToYmdHm(startDt).ymd,
            pendingCrewDutyDrop.endTime || "23:59",
          )
    if (endDt <= startDt) {
      window.alert("Duty end must be after start.")
      return
    }
    const startMeta = formatDateToYmdHm(startDt)
    const endMeta = formatDateToYmdHm(endDt)
    const entry: Entry = {
      id: `grb_drop_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`,
      createdAt: new Date().toISOString(),
      recordType: "generatedRoster",
      crewCode: pendingCrewDutyDrop.crewCode,
      crewName: pendingCrewDutyDrop.crewName,
      crewType: "",
      cycleNumber: "",
      startDate: startMeta.ymd,
      endDate: endMeta.ymd,
      startTime: startMeta.hm,
      endTime: endMeta.hm,
      rosterType: pendingCrewDutyDrop.rosterBarMap === "REST" ? "REST" : "DUTY CODE",
      dutyCode: pendingCrewDutyDrop.code,
      dutyType: pendingCrewDutyDrop.dutyType,
      dutyColor: config.color || getDefaultDutyCodeColor(pendingCrewDutyDrop.code),
      fdpPattern: "Manual drag assignment",
    }
    setStore((prev) => {
      const kept = prev.crewLeavePlanner.filter((row) => {
        if (pendingCrewDutyDrop.mode !== "override") return true
        if ((row.recordType || "").trim().toLowerCase() !== "generatedroster") return true
        if (!pendingCrewDutyDrop.conflicts.some((c) => c.id === row.id)) return true
        return false
      })
      return { ...prev, crewLeavePlanner: [entry, ...kept] }
    })
    setPendingCrewDutyDrop((p) => ({ ...p, open: false }))
  }

  const addActualLeaveSegment = (params: {
    crewCode: string
    crewName?: string
    cycleNumber: string
    startDate: string
    endDate: string
    maxDays?: number
    sourceBlockId?: string
    overrideConflicts?: string
    clearConflictingDutyCodes?: boolean
  }) => {
    const crewCode = normalizeCrewCode(params.crewCode || "")
    const cycleNumber = (params.cycleNumber || "").trim()
    const startDate = (params.startDate || "").trim()
    const endDate = (params.endDate || "").trim()
    if (!crewCode) {
      window.alert("Select a crew first.")
      return
    }
    if (!startDate || !endDate) {
      window.alert("Actual Leave Start and End are required.")
      return
    }
    if (endDate < startDate) {
      window.alert("Actual Leave End must be after or equal to Start.")
      return
    }
    const requestedDays = getDateRangeInclusive(startDate, endDate).length
    if (params.maxDays != null && requestedDays > params.maxDays) {
      window.alert(`Selected range exceeds remaining balance (${params.maxDays} day(s)).`)
      return
    }
    const entry: Entry = {
      id: `als_${Date.now()}_${Math.random().toString(16).slice(2, 6)}`,
      createdAt: new Date().toISOString(),
      recordType: "actualLeaveSegment",
      crewCode,
      crewName: params.crewName || "",
      cycleNumber,
      startDate,
      endDate,
      plannedDays: String(requestedDays),
      sourceBlockId: params.sourceBlockId || "",
      overrideConflicts: params.overrideConflicts || "",
    }
    setStore((prev) => {
      const kept = prev.crewLeavePlanner.filter((row) => {
        const recordType = (row.recordType || "").trim().toLowerCase()
        const rowCrewCode = normalizeCrewCode(row.crewCode || "")
        if (recordType === "actualleavesegment" && rowCrewCode === crewCode) {
          const sameSource =
            params.sourceBlockId &&
            (row.sourceBlockId || "").trim() &&
            (row.sourceBlockId || "").trim() === params.sourceBlockId
          const sameCycle =
            !params.sourceBlockId &&
            (row.cycleNumber || "").trim() === cycleNumber &&
            (row.startDate || "").trim() === startDate &&
            (row.endDate || row.startDate || "").trim() === endDate
          if (sameSource || sameCycle) return false
        }
        if (!params.clearConflictingDutyCodes) return true
        if (recordType !== "generatedroster") return true
        if (rowCrewCode !== crewCode) return true
        if ((row.rosterType || "").trim().toUpperCase() !== "DUTY CODE") return true
        const rowStart = (row.startDate || "").trim()
        const rowEnd = (row.endDate || rowStart).trim()
        return !rowStart || !rowEnd || overlapDaysInclusive(startDate, endDate, rowStart, rowEnd) === 0
      })
      return { ...prev, crewLeavePlanner: [entry, ...kept] }
    })
  }

  const addLeavePatternActualSegment = (cycleNumber: string) => {
    const crewCode = normalizeCrewCode(leavePatternManagerCrewCode || "")
    const draft = leavePatternSegmentDraft[cycleNumber] || { start: "", end: "" }
    addActualLeaveSegment({
      crewCode,
      cycleNumber,
      startDate: draft.start || "",
      endDate: draft.end || "",
    })
    setLeavePatternSegmentDraft((prev) => ({ ...prev, [cycleNumber]: { start: "", end: "" } }))
  }

  const removeLeavePatternActualSegment = (segmentId: string) => {
    if (!window.confirm("Delete this actual leave segment?")) return
    setStore((prev) => ({
      ...prev,
      crewLeavePlanner: prev.crewLeavePlanner.filter((row) => row.id !== segmentId),
    }))
    if (editingLeaveSegmentId === segmentId) {
      setEditingLeaveSegmentId("")
      setEditingLeaveSegmentDraft({ cycleNumber: "", start: "", end: "" })
    }
  }

  const updateLeavePatternActualSegment = (segmentId: string) => {
    const segmentIdValue = (segmentId || "").trim()
    const nextCycleNumber = (editingLeaveSegmentDraft.cycleNumber || "").trim()
    const startDate = (editingLeaveSegmentDraft.start || "").trim()
    const endDate = (editingLeaveSegmentDraft.end || "").trim()
    if (!segmentIdValue) return
    if (!nextCycleNumber) {
      window.alert("Select the cycle this Actual Leave should adjust.")
      return
    }
    if (!startDate || !endDate) {
      window.alert("Actual Leave Start and End are required.")
      return
    }
    if (endDate < startDate) {
      window.alert("Actual Leave End must be after or equal to Start.")
      return
    }
    const existing = store.crewLeavePlanner.find((row) => row.id === segmentIdValue)
    const existingDays = Number(existing?.plannedDays || "0") || 0
    const existingCycle = (existing?.cycleNumber || "").trim()
    const requestedDays = getDateRangeInclusive(startDate, endDate).length
    const targetCycleRow = leaveDetailPatternRows.find((r) => (r.cycleNumber || "").trim() === nextCycleNumber)
    const targetCycleRemainingDays = Number(targetCycleRow?.rem || "0") || 0
    const availableDays = targetCycleRemainingDays + (existingCycle === nextCycleNumber ? existingDays : 0)
    if (requestedDays > availableDays) {
      window.alert(`Selected range exceeds remaining balance (${availableDays} day(s)).`)
      return
    }
    const crewCode = normalizeCrewCode(leaveDetailModal.crewCode || existing?.crewCode || "")
    const matchingLeaveBlock = leaveDetailGeneratedRows.find(
      (row) =>
        normalizeCrewCode(row.crewCode || "") === crewCode &&
        (row.blockType || "").trim().toLowerCase() === "leave" &&
        (row.cycleNumber || "").trim() === nextCycleNumber,
    )
    setStore((prev) => ({
      ...prev,
      crewLeavePlanner: prev.crewLeavePlanner.map((row) => {
        if (row.id !== segmentIdValue) return row
        if ((row.recordType || "").trim().toLowerCase() !== "actualleavesegment") return row
        return {
          ...row,
          cycleNumber: nextCycleNumber,
          sourceBlockId: matchingLeaveBlock?.id || row.sourceBlockId || "",
          startDate,
          endDate,
          plannedDays: String(requestedDays),
          updatedAt: new Date().toISOString(),
        }
      }),
    }))
    setLeaveDetailSelectedCycle(nextCycleNumber)
    setEditingLeaveSegmentId("")
    setEditingLeaveSegmentDraft({ cycleNumber: "", start: "", end: "" })
  }

  const assignActualLeaveSegmentToCycle = (segmentId: string, cycleNumberValue: string) => {
    const segmentIdValue = (segmentId || "").trim()
    const cycleNumber = (cycleNumberValue || "").trim()
    if (!segmentIdValue) {
      window.alert("Select an Actual Leave segment first.")
      return
    }
    if (!cycleNumber) {
      window.alert("Select the cycle this Actual Leave should adjust.")
      return
    }
    const crewCode = normalizeCrewCode(leaveDetailModal.crewCode || "")
    const matchingLeaveBlock = leaveDetailGeneratedRows.find(
      (row) =>
        normalizeCrewCode(row.crewCode || "") === crewCode &&
        (row.blockType || "").trim().toLowerCase() === "leave" &&
        (row.cycleNumber || "").trim() === cycleNumber,
    )
    setStore((prev) => ({
      ...prev,
      crewLeavePlanner: prev.crewLeavePlanner.map((row) => {
        if (row.id !== segmentIdValue) return row
        if ((row.recordType || "").trim().toLowerCase() !== "actualleavesegment") return row
        const startDate = (row.startDate || "").trim()
        const endDate = (row.endDate || startDate).trim()
        return {
          ...row,
          crewCode,
          crewName: row.crewName || leaveDetailModal.crewName || "",
          cycleNumber,
          sourceBlockId: matchingLeaveBlock?.id || row.sourceBlockId || "",
          plannedDays:
            startDate && endDate && endDate >= startDate
              ? String(getDateRangeInclusive(startDate, endDate).length)
              : row.plannedDays || "",
        }
      }),
    }))
    setLeaveDetailModal((p) => ({
      ...p,
      blockId: matchingLeaveBlock?.id || p.blockId,
      cycleNumber,
      leaveStart: matchingLeaveBlock?.startDate || p.leaveStart,
      leaveEnd: matchingLeaveBlock?.endDate || p.leaveEnd,
      policyName: matchingLeaveBlock?.policyName || p.policyName,
    }))
    setLeaveDetailSelectedCycle(cycleNumber)
  }

  const shiftLeaveBarAndNextWork = (
    leaveBlockId: string,
    deltaDays: number,
  ): { crewCode: string; cycleNumber: string; leaveDeltaDays: number; nextWorkDeltaDays: number } | null => {
    if (!leaveBlockId || !deltaDays) return null
    let result: { crewCode: string; cycleNumber: string; leaveDeltaDays: number; nextWorkDeltaDays: number } | null = null
    setStore((prev) => {
      const rows = [...prev.crewLeavePlanner]
      const leaveIdx = rows.findIndex((r) => r.id === leaveBlockId)
      if (leaveIdx < 0) return prev
      const leaveRow = rows[leaveIdx]
      if ((leaveRow.recordType || "").trim().toLowerCase() !== "generatedblock") return prev
      if ((leaveRow.blockType || "").trim().toLowerCase() !== "leave") return prev
      const oldLeaveStart = (leaveRow.startDate || "").trim()
      const oldLeaveEnd = (leaveRow.endDate || "").trim()
      if (!oldLeaveStart || !oldLeaveEnd) return prev

      const newLeaveStart = addDaysToYmd(oldLeaveStart, deltaDays)
      const newLeaveEnd = addDaysToYmd(oldLeaveEnd, deltaDays)
      rows[leaveIdx] = {
        ...leaveRow,
        originalStartDate: (leaveRow.originalStartDate || "").trim() || oldLeaveStart,
        originalEndDate: (leaveRow.originalEndDate || "").trim() || oldLeaveEnd,
        startDate: newLeaveStart,
        endDate: newLeaveEnd,
      }

      const crewCode = normalizeCrewCode(leaveRow.crewCode || "")
      const currentCycle = Number(leaveRow.cycleNumber || "0") || 0
      const nextCycle = String(currentCycle + 1)
      let nextWorkDeltaDays = 0

      const nextWorkIdx = rows.findIndex((r) => {
        if ((r.recordType || "").trim().toLowerCase() !== "generatedblock") return false
        if ((r.blockType || "").trim().toLowerCase() !== "work") return false
        if (normalizeCrewCode(r.crewCode || "") !== crewCode) return false
        return (r.cycleNumber || "").trim() === nextCycle
      })
      if (nextWorkIdx >= 0) {
        const nextWork = rows[nextWorkIdx]
        const oldWorkStart = (nextWork.startDate || "").trim()
        const oldWorkEnd = (nextWork.endDate || "").trim()
        const workDays = Number(nextWork.plannedDays || "0") || getDateRangeInclusive(oldWorkStart, oldWorkEnd).length
        const newWorkStart = addDaysToYmd(newLeaveEnd, 1)
        const newWorkEnd = workDays > 0 ? addDaysToYmd(newWorkStart, workDays - 1) : newWorkStart
        rows[nextWorkIdx] = {
          ...nextWork,
          originalStartDate: (nextWork.originalStartDate || "").trim() || oldWorkStart,
          originalEndDate: (nextWork.originalEndDate || "").trim() || oldWorkEnd,
          startDate: newWorkStart,
          endDate: newWorkEnd,
        }
        nextWorkDeltaDays = deltaDays
      }

      result = {
        crewCode,
        cycleNumber: String(currentCycle),
        leaveDeltaDays: deltaDays,
        nextWorkDeltaDays,
      }
      return { ...prev, crewLeavePlanner: rows }
    })
    return result
  }

  const resetGeneratedCycleToOriginal = (crewCodeValue: string, cycleNumber: string) => {
    const crewCode = normalizeCrewCode(crewCodeValue || "")
    const cycle = (cycleNumber || "").trim()
    if (!crewCode || !cycle) return
    setStore((prev) => {
      const nextRows = prev.crewLeavePlanner.map((row) => {
        if ((row.recordType || "").trim().toLowerCase() !== "generatedblock") return row
        if (normalizeCrewCode(row.crewCode || "") !== crewCode) return row
        if ((row.cycleNumber || "").trim() !== cycle) return row
        const originalStart = (row.originalStartDate || "").trim()
        const originalEnd = (row.originalEndDate || "").trim()
        if (!originalStart || !originalEnd) return row
        return {
          ...row,
          startDate: originalStart,
          endDate: originalEnd,
        }
      })
      return { ...prev, crewLeavePlanner: nextRows }
    })
  }


  const openInactivePopup = ({
    staffId,
    fromEdit,
    seedDate,
    seedReason,
  }: {
    staffId: string | null
    fromEdit: boolean
    seedDate?: string
    seedReason?: string
  }) => {
    setInactivePopupStaffId(staffId)
    setInactivePopupFromEdit(fromEdit)
    setInactiveDateDraft(seedDate || toYmd(new Date()))
    setInactiveReasonDraft(seedReason || "")
    setIsInactivePopupOpen(true)
  }

  const closeInactivePopup = () => {
    setIsInactivePopupOpen(false)
    setInactivePopupStaffId(null)
    setInactivePopupFromEdit(false)
    setInactiveDateDraft("")
    setInactiveReasonDraft("")
  }

  useEffect(() => {
    function onEsc(event: KeyboardEvent) {
      if (event.key !== "Escape") return
      if (isFormOpen) {
        setIsFormOpen(false)
        return
      }
      if (isRosterGeneratorOpen) {
        setIsRosterGeneratorOpen(false)
        return
      }
      if (isInactivePopupOpen) {
        closeInactivePopup()
      }
    }
    window.addEventListener("keydown", onEsc)
    return () => window.removeEventListener("keydown", onEsc)
  }, [isFormOpen, isRosterGeneratorOpen, isInactivePopupOpen])

  const confirmInactive = () => {
    if (!inactiveDateDraft.trim() || !inactiveReasonDraft.trim()) {
      window.alert("Inactive Date and Inactive Reason are required.")
      return
    }

    if (inactivePopupFromEdit) {
      updateForm("activeStatus", "Inactive")
      updateForm("inactiveDate", inactiveDateDraft)
      updateForm("inactiveReason", inactiveReasonDraft)
      closeInactivePopup()
      return
    }

    if (!inactivePopupStaffId) return
    setStore((prev) => ({
      ...prev,
      staff: prev.staff.map((item) =>
        item.id === inactivePopupStaffId
          ? {
              ...item,
              activeStatus: "Inactive",
              inactiveDate: inactiveDateDraft,
              inactiveReason: inactiveReasonDraft,
            }
          : item,
      ),
    }))
    closeInactivePopup()
  }

  const setStaffActive = (staffId: string) => {
    setStore((prev) => ({
      ...prev,
      staff: prev.staff.map((item) =>
        item.id === staffId
          ? { ...item, activeStatus: "Active", inactiveDate: "", inactiveReason: "" }
          : item,
      ),
    }))
  }

  const calculateLeaveSummary = (fromDateValue: string, toDateValue: string, leavePolicyName: string) => {
    const fromDate = normalizeDateOnly(fromDateValue)
    const toDate = normalizeDateOnly(toDateValue) || fromDate
    const dateRange = getDateRangeInclusive(fromDate, toDate)
    const weekdayMap = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    const policy = store.leaveAttendanceControl.find(
      (item) => item.leaveAttendanceType === "Leave" && item.leaveAttendanceName === leavePolicyName,
    )
    const weekdaysNotCounted = splitMultiValue(policy?.nonCountWeekdays || "")

    const holidaySet = new Set<string>()
    for (const holiday of store.publicHolidays) {
      const holidayRange = getDateRangeInclusive(holiday.holidayStartDate, holiday.holidayEndDate)
      for (const day of holidayRange) {
        holidaySet.add(toYmd(day))
      }
    }

    const excludedDates = new Set<string>()
    const excludedWeekdayDateKeys = new Set<string>()
    const govPublicHolidayDateKeys = new Set<string>()
    for (const day of dateRange) {
      const dayKey = toYmd(day)
      const weekdayName = weekdayMap[day.getDay()]
      if (weekdaysNotCounted.includes(weekdayName)) {
        excludedDates.add(dayKey)
        excludedWeekdayDateKeys.add(dayKey)
      }
      if (holidaySet.has(dayKey)) {
        govPublicHolidayDateKeys.add(dayKey)
        excludedDates.add(dayKey)
      }
    }

    const totalCalendarDays = dateRange.length
    const excludedWeekdayDays = excludedWeekdayDateKeys.size
    const excludedWeekdayDates = Array.from(excludedWeekdayDateKeys)
      .sort()
      .map((d) => formatYmdToDdMmYy(d))
    const effectiveGovPublicHolidayDateKeys = new Set(
      Array.from(govPublicHolidayDateKeys).filter((d) => !excludedWeekdayDateKeys.has(d)),
    )
    const publicHolidayDays = effectiveGovPublicHolidayDateKeys.size
    const govPublicHolidayDates = Array.from(effectiveGovPublicHolidayDateKeys)
      .sort()
      .map((d) => formatYmdToDdMmYy(d))
    const chargeableLeaveDays = Math.max(totalCalendarDays - excludedDates.size, 0)

    return {
      fromDate,
      toDate,
      totalCalendarDays: String(totalCalendarDays),
      excludedWeekdayDetails:
        excludedWeekdayDays > 0
          ? `${excludedWeekdayDays} day(s): ${excludedWeekdayDates.join(", ")}`
          : "0 day(s)",
      publicHolidayDays:
        publicHolidayDays > 0 ? `${publicHolidayDays} day(s): ${govPublicHolidayDates.join(", ")}` : "0 day(s)",
      chargeableLeaveDays: String(chargeableLeaveDays),
    }
  }

  const resetAllData = () => {
    if (!window.confirm("Reset all OCCfloat data? This action cannot be undone.")) return
    setStore(getEmptyStore())
    setForms(
      MODULES.reduce((acc, module) => {
        acc[module.key] = buildEmptyForm(module)
        return acc
      }, {} as Record<ModuleKey, Record<string, string>>),
    )
    window.localStorage.removeItem(STORAGE_KEY)
    setEditingEntryId(null)
    closeInactivePopup()
  }

  const getGovPublicHolidayDisplay = (leaveEntry: Entry): string => {
    const fromDate = (leaveEntry.fromDate || "").trim()
    const toDate = (leaveEntry.toDate || "").trim()
    if (!fromDate || !toDate) {
      return leaveEntry.publicHolidayDays || "0 day(s)"
    }
    const leaveRange = getDateRangeInclusive(fromDate, toDate)
    if (leaveRange.length === 0) {
      return leaveEntry.publicHolidayDays || "0 day(s)"
    }

    const holidaySet = new Set<string>()
    store.publicHolidays.forEach((holiday) => {
      const range = getDateRangeInclusive(holiday.holidayStartDate, holiday.holidayEndDate)
      range.forEach((day) => holidaySet.add(toYmd(day)))
    })

    const policy = store.leaveAttendanceControl.find(
      (item) =>
        item.leaveAttendanceType === "Leave" &&
        item.leaveAttendanceName === (leaveEntry.leavePolicyName || "").trim(),
    )
    const weekdaysNotCounted = new Set(splitMultiValue(policy?.nonCountWeekdays || ""))
    const weekdayMap = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

    // If PH and GH overlap on the same date, GH is not counted.
    const hitDates = leaveRange
      .filter((day) => holidaySet.has(toYmd(day)))
      .filter((day) => !weekdaysNotCounted.has(weekdayMap[day.getDay()]))
      .map((day) => toYmd(day))

    if (hitDates.length === 0) return "0 day(s)"
    return `${hitDates.length} day(s): ${hitDates
      .sort()
      .map((d) => formatYmdToDdMmYy(d))
      .join(", ")}`
  }

  const normalizeImportDate = (value: string): string => {
    const raw = (value || "").trim()
    if (!raw) return ""
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw
    const slash = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/)
    if (slash) {
      const dd = slash[1].padStart(2, "0")
      const mm = slash[2].padStart(2, "0")
      const yyyy = slash[3].length === 2 ? `20${slash[3]}` : slash[3]
      return `${yyyy}-${mm}-${dd}`
    }
    const dash = raw.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/)
    if (dash) {
      const dd = dash[1].padStart(2, "0")
      const mm = dash[2].padStart(2, "0")
      const yyyy = dash[3].length === 2 ? `20${dash[3]}` : dash[3]
      return `${yyyy}-${mm}-${dd}`
    }
    return ""
  }

  const downloadPublicHolidayCsvTemplate = () => {
    const template = [
      "holidayStartDate,holidayEndDate,holidayName",
      "2026-01-01,2026-01-01,New Year",
      "01/05/2026,01/05/2026,Labour Day",
    ].join("\n")
    const blob = new Blob([template], { type: "text/csv;charset=utf-8;" })
    const url = URL.createObjectURL(blob)
    const link = document.createElement("a")
    link.href = url
    link.download = "public-holidays-template.csv"
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    URL.revokeObjectURL(url)
  }

  const importCrewDatabaseCsv = async (file: File) => {
    const text = await file.text()
    const lines = text
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
    if (lines.length < 2) {
      window.alert("CSV file is empty.")
      return
    }

    const header = parseCsvLine(lines[0]).map((h) => normalizeText(h))
    const indexCrewCode = header.findIndex((h) => h === normalizeText("Crew Code"))
    const indexCrewName = header.findIndex((h) => h === normalizeText("Crew Name"))
    const indexCrewType = header.findIndex((h) => h === normalizeText("Crew Type"))
    const indexCrewCategory = header.findIndex((h) => h === normalizeText("Crew Category"))
    const indexJoiningDate = header.findIndex((h) => h === normalizeText("Joining Date"))
    const indexEmailId = header.findIndex((h) => h === normalizeText("Email ID"))
    const indexMobileNo = header.findIndex((h) => h === normalizeText("Mobile No"))
    const indexEmployeeNo = header.findIndex((h) => h === normalizeText("Employee No"))
    const indexGender = header.findIndex((h) => h === normalizeText("Gender"))

    if (indexCrewCode < 0 || indexCrewName < 0) {
      window.alert('Invalid CSV header. Required columns: "Crew Code, Crew Name".')
      return
    }

    const parsed: Entry[] = []
    const errors: string[] = []
    const seenCodes = new Set<string>()
    for (let i = 1; i < lines.length; i++) {
      const cells = parseCsvLine(lines[i])
      const crewCode = normalizeCrewCode(cells[indexCrewCode] || "")
      const crewName = (cells[indexCrewName] || "").trim()
      if (!crewCode || !crewName) {
        errors.push(`Row ${i + 1}: missing Crew Code or Crew Name.`)
        continue
      }
      if (!/^[A-Z]{4}$/.test(crewCode)) {
        errors.push(`Row ${i + 1}: Crew Code must be exactly 4 letters.`)
        continue
      }
      if (seenCodes.has(crewCode)) {
        errors.push(`Row ${i + 1}: duplicate Crew Code "${crewCode}" in CSV.`)
        continue
      }
      seenCodes.add(crewCode)
      const joiningDate = normalizeImportDate(cells[indexJoiningDate] || "")
      const genderRaw = (cells[indexGender] || "").trim().toUpperCase()
      const gender = genderRaw === "M" || genderRaw === "F" ? genderRaw : ""

      parsed.push({
        id: `crew_${Date.now()}_${i}_${Math.random().toString(16).slice(2, 6)}`,
        createdAt: new Date().toISOString(),
        crewCode,
        crewName,
        crewType: (cells[indexCrewType] || "").trim(),
        crewCategory: (cells[indexCrewCategory] || "").trim(),
        joiningDate,
        emailId: (cells[indexEmailId] || "").trim(),
        mobileNo: (cells[indexMobileNo] || "").trim(),
        employeeNo: (cells[indexEmployeeNo] || "").trim(),
        gender,
        activeStatus: "Active",
        inactiveDate: "",
        inactiveReason: "",
      })
    }

    if (parsed.length === 0) {
      window.alert(errors.length > 0 ? errors.slice(0, 5).join("\n") : "No valid rows found.")
      return
    }

    setStore((prev) => {
      const byCode = new Map<string, Entry>()
      prev.crewDataBase.forEach((row) => {
        const key = normalizeCrewCode(row.crewCode || "")
        if (key) byCode.set(key, row)
      })
      parsed.forEach((row) => {
        const key = normalizeCrewCode(row.crewCode || "")
        const existing = byCode.get(key)
        byCode.set(key, {
          ...(existing || {}),
          ...row,
          id: existing?.id || row.id,
          createdAt: existing?.createdAt || row.createdAt,
        })
      })
      return { ...prev, crewDataBase: Array.from(byCode.values()) }
    })

    const errorText =
      errors.length > 0 ? `\n\nSkipped ${errors.length} row(s):\n${errors.slice(0, 5).join("\n")}` : ""
    window.alert(`Crew DB updated. Imported/updated ${parsed.length} row(s).${errorText}`)
  }

  const importPublicHolidayCsv = async (file: File) => {
    const rawText = await file.text()
    const lines = rawText
      .split(/\r\n|\n|\r/)
      .map((line) => line.trim())
      .filter(Boolean)
    if (lines.length < 2) {
      window.alert("CSV has no data rows.")
      return
    }

    const header = parseCsvLine(lines[0]).map((h) => h.trim())
    const indexStart = header.findIndex((h) => normalizeText(h) === "holidaystartdate")
    const indexEnd = header.findIndex((h) => normalizeText(h) === "holidayenddate")
    const indexName = header.findIndex((h) => normalizeText(h) === "holidayname")
    if (indexStart < 0 || indexEnd < 0 || indexName < 0) {
      window.alert(
        'Invalid CSV header. Required columns: "holidayStartDate, holidayEndDate, holidayName".',
      )
      return
    }

    const parsed: Entry[] = []
    const errors: string[] = []
    for (let i = 1; i < lines.length; i++) {
      const cells = parseCsvLine(lines[i])
      const startDate = normalizeImportDate(cells[indexStart] || "")
      const endDate = normalizeImportDate(cells[indexEnd] || "")
      const holidayName = (cells[indexName] || "").trim()
      if (!startDate || !endDate || !holidayName) {
        errors.push(`Row ${i + 1}: missing/invalid start date, end date, or holiday name.`)
        continue
      }
      if (startDate > endDate) {
        errors.push(`Row ${i + 1}: start date is after end date.`)
        continue
      }
      parsed.push({
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        createdAt: new Date().toISOString(),
        holidayStartDate: startDate,
        holidayEndDate: endDate,
        holidayName,
      })
    }

    if (parsed.length === 0) {
      window.alert(`No valid rows imported.\n${errors.slice(0, 5).join("\n")}`)
      return
    }

    setStore((prev) => ({
      ...prev,
      publicHolidays: [...parsed, ...prev.publicHolidays],
    }))

    if (errors.length > 0) {
      window.alert(
        `Imported ${parsed.length} row(s) with ${errors.length} skipped row(s).\n${errors
          .slice(0, 5)
          .join("\n")}`,
      )
    } else {
      window.alert(`Imported ${parsed.length} public holiday row(s).`)
    }
  }

  const downloadRosterCsv = () => {
    if (!rosterExportRange.fromDate || !rosterExportRange.toDate) {
      window.alert("Please select export period.")
      return
    }
    if (rosterExportRange.fromDate > rosterExportRange.toDate) {
      window.alert('"Export To Date" must be after or equal to "Export From Date".')
      return
    }
    const exportDates = getDateRangeInclusive(rosterExportRange.fromDate, rosterExportRange.toDate).map((d) =>
      toYmd(d),
    )
    if (exportDates.length === 0) {
      window.alert("No dates in selected export period.")
      return
    }
    const roleByStaffToken = new Map<string, string>()
    store.staff.forEach((s) => {
      const token = `${(s.staffNo || "").trim()}::${(s.fullName || "").trim()}`
      roleByStaffToken.set(token, s.designation || "")
    })
    const headerCells = ["staffNo", "staffName", "level", "role", ...exportDates]
    const dataRows = rosterMatrixRows.map((row) => {
      const token = `${(row.staffNo || "").trim()}::${(row.staffName || "").trim()}`
      const role = roleByStaffToken.get(token) || ""
      const cells = [
        row.staffNo || "",
        row.staffName || "",
        row.level || "",
        role,
        ...exportDates.map((d) => row.cells[d] || "-"),
      ]
      return cells.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(",")
    })
    const csv = `\uFEFF${[headerCells.join(","), ...dataRows].join("\n")}`
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" })
    const url = URL.createObjectURL(blob)
    const link = document.createElement("a")
    link.href = url
    link.download = `roster-grid-${rosterExportRange.fromDate}_to_${rosterExportRange.toDate}.csv`
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    URL.revokeObjectURL(url)
  }

  const openRosterPrintView = () => {
    if (!rosterExportRange.fromDate || !rosterExportRange.toDate) {
      window.alert("Please select print period.")
      return
    }
    if (rosterExportRange.fromDate > rosterExportRange.toDate) {
      window.alert('"Print To Date" must be after or equal to "Print From Date".')
      return
    }
    const exportDates = getDateRangeInclusive(rosterExportRange.fromDate, rosterExportRange.toDate).map((d) =>
      toYmd(d),
    )
    if (exportDates.length === 0) {
      window.alert("No dates in selected print period.")
      return
    }

    const rows = rosterMatrixRows.map((row) => ({
      staffNo: row.staffNo || "",
      staffName: row.staffName || "",
      level: row.level || "",
      cells: row.cells,
    }))

    const codeCellHtml = (codeRaw: string) => {
      const code = (codeRaw || "-").toUpperCase()
      if (code === "-" || !code) return `<span class="muted">-</span>`
      if (isOffCode(code)) {
        return `<span class="code" style="background:${ROSTER_OFF_COLOR.bg};color:${ROSTER_OFF_COLOR.text};">OF</span>`
      }
      if (code === "AL") {
        return `<span class="code" style="background:${ROSTER_AL_COLOR.bg};color:${ROSTER_AL_COLOR.text};">AL</span>`
      }
      if (code === "GH") {
        return `<span class="code" style="background:${ROSTER_GH_COLOR.bg};color:${ROSTER_GH_COLOR.text};">GH</span>`
      }
      if (code === "PH") {
        return `<span class="code" style="background:${ROSTER_PH_COLOR.bg};color:${ROSTER_PH_COLOR.text};">PH</span>`
      }
      const color = rosterShiftColorsByCode.get(code)
      if (color) {
        return `<span class="code" style="background:${color.bg};color:${color.text};">${code}</span>`
      }
      return `<span class="code">${code}</span>`
    }

    const headerDates = exportDates
      .map(
        (d) =>
          `<th class="date-col"><div class="date-vert"><span class="date-text">${formatYmdToDdMmYy(d)}</span><span class="weekday">${formatYmdToWeekdayShort(d)}</span></div></th>`,
      )
      .join("")

    const bodyRows = rows
      .map((row) => {
        const dateCells = exportDates
          .map((d) => `<td>${codeCellHtml((row.cells[d] || "").trim())}</td>`)
          .join("")
        return `<tr><td class="sticky-left">${row.staffNo}</td><td class="sticky-left-2">${row.staffName}</td><td class="sticky-left-3">${row.level}</td>${dateCells}</tr>`
      })
      .join("")

    const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Roster Print View</title>
  <style>
    @page { size: A4 landscape; margin: 6mm; }
    body { font-family: Inter, Arial, sans-serif; margin: 0; color: #111827; }
    .wrap { padding: 4px; }
    .title { font-size: 16px; font-weight: 700; margin-bottom: 4px; }
    .sub { font-size: 11px; color: #4b5563; margin-bottom: 6px; }
    .toolbar { margin-bottom: 6px; }
    .toolbar button { padding: 6px 10px; border: 1px solid #cbd5e1; border-radius: 6px; background: #fff; cursor: pointer; }
    .table-wrap { overflow: auto; border: 1px solid #e5e7eb; border-radius: 8px; }
    table { border-collapse: collapse; min-width: 1200px; width: max-content; }
    th, td { border: 1px solid #e5e7eb; padding: 2px 2px; text-align: center; white-space: nowrap; font-size: 10px; }
    thead th { background: #f8fafc; position: sticky; top: 0; z-index: 5; }
    .weekday { font-size: 9px; color: #6b7280; }
    .date-col { width: 24px; min-width: 24px; max-width: 24px; padding: 1px 0; vertical-align: bottom; }
    .date-vert { display: inline-flex; flex-direction: column; align-items: center; gap: 2px; line-height: 1; }
    .date-text { writing-mode: vertical-rl; transform: rotate(180deg); }
    .sticky-left, .sticky-left-2, .sticky-left-3 { background: #fff; position: sticky; z-index: 4; text-align: left; }
    .sticky-left { left: 0; min-width: 68px; }
    .sticky-left-2 { left: 68px; min-width: 132px; }
    .sticky-left-3 { left: 200px; min-width: 62px; text-align: center; }
    thead .sticky-left, thead .sticky-left-2, thead .sticky-left-3 { background: #f8fafc; z-index: 6; }
    .code { display: inline-flex; align-items: center; justify-content: center; min-width: 1.5rem; height: 1rem; padding: 0 0.18rem; font-weight: 600; border-radius: 0; font-size: 9px; line-height: 1; }
    .muted { color: #9ca3af; }
    @media print {
      .toolbar { display: none; }
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .table-wrap { border: 0; border-radius: 0; overflow: visible; }
      table { min-width: 0; width: max-content; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="title">Roster</div>
    <div class="sub">Period: ${formatYmdToDdMmYy(rosterExportRange.fromDate)} to ${formatYmdToDdMmYy(rosterExportRange.toDate)}</div>
    <div class="toolbar"><button onclick="window.print()">Print</button></div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th class="sticky-left">Staff No</th>
            <th class="sticky-left-2">Staff Name</th>
            <th class="sticky-left-3">Level</th>
            ${headerDates}
          </tr>
        </thead>
        <tbody>
          ${bodyRows}
        </tbody>
      </table>
    </div>
  </div>
</body>
</html>`

    const blob = new Blob([html], { type: "text/html;charset=utf-8" })
    const url = URL.createObjectURL(blob)
    const a = document.createElement("a")
    a.href = url
    a.target = "_blank"
    a.rel = "noopener noreferrer"
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)

    // Fallback for environments that still block _blank opening.
    setTimeout(() => {
      if (!document.hidden) {
        window.location.href = url
      }
    }, 250)
  }

  const importRosterCsv = async (file: File) => {
    const rawText = await file.text()
    const lines = rawText
      .split(/\r\n|\n|\r/)
      .map((line) => line.trim())
      .filter(Boolean)
    if (lines.length < 2) {
      window.alert("Roster CSV has no data rows.")
      return
    }

    const headerRaw = parseCsvLine(lines[0]).map((h) => h.replace(/\uFEFF/g, "").trim())
    const header = headerRaw.map((h) => normalizeText(h))
    const indexStaffNo = header.findIndex((h) => h === "staffno")
    const indexStaffName = header.findIndex((h) => h === "staffname")
    const indexRole = header.findIndex((h) => h === "role")
    if (indexStaffNo < 0 || indexStaffName < 0) {
      window.alert('Invalid roster matrix CSV header. Required: "staffNo, staffName, ...date columns".')
      return
    }
    const dateColumns = headerRaw
      .map((h, idx) => ({ idx, ymd: normalizeImportDate(h) }))
      .filter((x) => x.ymd)
    if (dateColumns.length === 0) {
      window.alert("No date columns found in roster matrix file.")
      return
    }

    const today = toYmd(new Date())
    const errors: string[] = []
    let pastDateRejectCount = 0
    const leaveTypeCodes = new Set<string>(["AL"])
    store.leaveAttendanceControl
      .filter((p) => (p.leaveAttendanceType || "").trim().toLowerCase() === "leave")
      .forEach((p) => {
        const code = (p.typeCode || "").trim().toUpperCase()
        if (code) leaveTypeCodes.add(code)
      })
    const existingMap = new Map<string, Entry>()
    store.roster.forEach((row) => {
      const key = `${(row.date || "").trim()}::${(row.staffNo || "").trim()}::${(row.staffName || "").trim()}`
      existingMap.set(key, row)
    })
    const nextRosterMap = new Map<string, Entry>()
    store.roster.forEach((row) => {
      const key = `${(row.date || "").trim()}::${(row.staffNo || "").trim()}::${(row.staffName || "").trim()}`
      nextRosterMap.set(key, row)
    })
    const changeLogs: RosterChangeLog[] = []
    let parsedCells = 0

    for (let i = 1; i < lines.length; i++) {
      const cells = parseCsvLine(lines[i])
      const staffNo = (cells[indexStaffNo] || "").trim()
      const staffName = (cells[indexStaffName] || "").trim()
      const role = indexRole >= 0 ? (cells[indexRole] || "").trim() : ""
      if (!staffNo || !staffName) {
        errors.push(`Row ${i + 1}: missing staffNo/staffName.`)
        continue
      }
      dateColumns.forEach(({ idx, ymd }) => {
        const rawCode = (cells[idx] || "").trim()
        const code = rawCode.toUpperCase()
        const leaveKey = `${staffNo}::${ymd}`
        const existingLeaveCode = (leaveCodeByStaffAndDate.get(leaveKey) || "").trim().toUpperCase()
        if (!rawCode || rawCode === "-") {
          const key = `${ymd}::${staffNo}::${staffName}`
          const prev = existingMap.get(key)
          const prevCode = (prev?.shiftCode || "").trim().toUpperCase()
          if (existingLeaveCode || leaveTypeCodes.has(prevCode)) {
            errors.push(
              `Row ${i + 1}, ${formatYmdToDdMmYy(ymd)}: skipped because leave is allocated (${existingLeaveCode || prevCode}).`,
            )
            return
          }
          if (prev) {
            nextRosterMap.delete(key)
            changeLogs.push({
              id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
              at: new Date().toISOString(),
              requestedBy: authenticatedStaff?.fullName || authenticatedStaff?.staffNo || "System",
              date: ymd,
              staffNo,
              staffName,
              type: "Upload",
              action: "Removed",
              details: `Before: shift=${prev.shiftCode || "-"}, role=${prev.role || "-"}, location=${prev.location || "-"} -> After: -`,
              reason: "CSV Upload",
            })
          }
          return
        }
        if (ymd < today && !rosterUploadAdminMode) {
          pastDateRejectCount += 1
          errors.push(`Row ${i + 1}, ${formatYmdToDdMmYy(ymd)}: past date updates are not allowed.`)
          return
        }
        if (code === "PH" || code === "GH") {
          errors.push(`Row ${i + 1}, ${formatYmdToDdMmYy(ymd)}: ${code} is system-managed and cannot be uploaded.`)
          return
        }
        if (existingLeaveCode) {
          errors.push(
            `Row ${i + 1}, ${formatYmdToDdMmYy(ymd)}: skipped because leave is allocated (${existingLeaveCode}).`,
          )
          return
        }
        if (leaveTypeCodes.has(code)) {
          errors.push(
            `Row ${i + 1}, ${formatYmdToDdMmYy(ymd)}: leave code ${code} is protected and cannot be changed by upload.`,
          )
          return
        }

        parsedCells += 1
        const key = `${ymd}::${staffNo}::${staffName}`
        const prev = existingMap.get(key)
        const next: Entry = prev
          ? {
              ...prev,
              role: role || prev.role || "",
              shiftCode: rawCode,
              location: "Uploaded Grid",
            }
          : {
              id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
              createdAt: new Date().toISOString(),
              date: ymd,
              staffNo,
              staffName,
              role: role || "",
              shiftCode: rawCode,
              location: "Uploaded Grid",
            }
        nextRosterMap.set(key, next)

        const beforeText = prev
          ? `shift=${prev.shiftCode || "-"}, role=${prev.role || "-"}, location=${prev.location || "-"}`
          : "-"
        const afterText = `shift=${next.shiftCode || "-"}, role=${next.role || "-"}, location=${next.location || "-"}`
        if (!prev || beforeText !== afterText) {
          changeLogs.push({
            id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
            at: new Date().toISOString(),
            requestedBy: authenticatedStaff?.fullName || authenticatedStaff?.staffNo || "System",
            date: ymd,
            staffNo,
            staffName,
            type: "Upload",
            action: prev ? "Updated" : "Added",
            details: `Before: ${beforeText} -> After: ${afterText}`,
            reason: "CSV Upload",
          })
        }
      })
    }

    if (parsedCells === 0 && changeLogs.length === 0) {
      if (pastDateRejectCount > 0 && !rosterUploadAdminMode) {
        window.alert(
          `Import blocked: this file mostly contains past-date roster cells.\n` +
            `Enable "Administrator upload mode" in Roster Import/Export to upload past roster data.\n` +
            `Rejected past-date cells: ${pastDateRejectCount}`,
        )
        return
      }
      window.alert(`No valid roster cells imported.\n${errors.slice(0, 8).join("\n")}`)
      return
    }

    setStore((prev) => ({
      ...prev,
      roster: Array.from(nextRosterMap.values()),
    }))
    setRosterChangeLogs((prev) => [...changeLogs, ...prev].slice(0, 800))

    const importedCount = parsedCells
    const changedCount = changeLogs.length
    if (errors.length > 0) {
      window.alert(
        `Roster import complete.\nImported rows: ${importedCount}\nChanged rows: ${changedCount}\nSkipped rows: ${errors.length}\n${errors
          .slice(0, 6)
          .join("\n")}`,
      )
    } else {
      window.alert(`Roster import complete.\nImported rows: ${importedCount}\nChanged rows: ${changedCount}`)
    }
  }

  const applyRosterCellEdit = (
    staffNo: string,
    staffName: string,
    date: string,
    rawCode: string,
  ) => {
    const code = (rawCode || "").trim().toUpperCase()
    const key = `${date}::${staffNo}::${staffName}`
    const role =
      store.staff.find(
        (s) =>
          (s.staffNo || "").trim() === (staffNo || "").trim() &&
          (s.fullName || "").trim() === (staffName || "").trim(),
      )?.designation || ""
    const previousCode =
      rosterMatrixRows.find((r) => r.staffNo === staffNo && r.staffName === staffName)?.cells[date] || "-"
    setStore((prev) => {
      const nextMap = new Map<string, Entry>()
      prev.roster.forEach((row) => {
        const rowKey = `${(row.date || "").trim()}::${(row.staffNo || "").trim()}::${(row.staffName || "").trim()}`
        nextMap.set(rowKey, row)
      })
      if (!code || code === "-") {
        nextMap.delete(key)
      } else {
        const existing = nextMap.get(key)
        const nextEntry: Entry = existing
          ? { ...existing, shiftCode: code, location: "Manual Edit" }
          : {
              id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
              createdAt: new Date().toISOString(),
              date,
              staffNo,
              staffName,
              role,
              shiftCode: code,
              location: "Manual Edit",
            }
        nextMap.set(key, nextEntry)
      }
      return {
        ...prev,
        roster: Array.from(nextMap.values()),
      }
    })
    const manualLog: RosterChangeLog = {
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      at: new Date().toISOString(),
      requestedBy: authenticatedStaff?.fullName || authenticatedStaff?.staffNo || "System",
      date,
      staffNo,
      staffName,
      type: "Manual",
      action: "Updated",
      details: `Before: ${previousCode} -> After: ${code || "-"}`,
      reason: "Grid Edit Mode",
    }
    setRosterChangeLogs((prev) => [manualLog, ...prev].slice(0, 800))
  }

  const setRosterCodeForCell = (
    rows: Entry[],
    staffNo: string,
    staffName: string,
    date: string,
    code: string,
  ): Entry[] => {
    const key = `${date}::${staffNo}::${staffName}`
    const map = new Map<string, Entry>()
    rows.forEach((r) => {
      map.set(`${(r.date || "").trim()}::${(r.staffNo || "").trim()}::${(r.staffName || "").trim()}`, r)
    })
    if (!code || code === "-") {
      map.delete(key)
    } else {
      const role =
        store.staff.find(
          (s) =>
            (s.staffNo || "").trim() === (staffNo || "").trim() &&
            (s.fullName || "").trim() === (staffName || "").trim(),
        )?.designation || ""
      const existing = map.get(key)
      map.set(
        key,
        existing
          ? { ...existing, shiftCode: code, location: "Change Request" }
          : {
              id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
              createdAt: new Date().toISOString(),
              date,
              staffNo,
              staffName,
              role,
              shiftCode: code,
              location: "Change Request",
            },
      )
    }
    return Array.from(map.values())
  }

  const getAllowedRosterCodes = (): Set<string> => {
    const allowed = new Set<string>(["-", "OF", "OFF", "REST", "PH", "GH", "AL"])
    store.shift.forEach((s) => {
      const code = (s.shiftCode || "").trim().toUpperCase()
      if (code) allowed.add(code)
    })
    store.leaveAttendanceControl
      .filter((l) => normalizeText(l.leaveAttendanceType || "") === "leave")
      .forEach((l) => {
        const code = (l.typeCode || "").trim().toUpperCase()
        if (code) allowed.add(code)
      })
    return allowed
  }

  const openRosterChangePopup = (cell: {
    date: string
    staffNo: string
    staffName: string
    level: string
    code: string
  }) => {
    const today = toYmd(new Date())
    if (cell.date <= today) {
      window.alert("Roster changes are allowed only for future dates.")
      return
    }
    setRosterChangeCell(cell)
    setRosterChangeType("Shift Reassign")
    setRosterChangeToCode(cell.code === "-" ? "OF" : cell.code)
    setRosterChangeTargetStaff("")
    setRosterChangeReason("")
    setRosterChangeConflictToCode("")
    setIsRosterChangeModalOpen(true)
  }

  const submitRosterChangeRequest = () => {
    if (!rosterChangeCell) return
    if (!rosterChangeReason.trim()) {
      window.alert("State Change Reason is required.")
      return
    }
    if (rosterChangeType !== "Mutual Swap" && !rosterChangeToCode.trim()) {
      window.alert("Please select changed code.")
      return
    }
    if (rosterChangeType === "Mutual Swap" && !rosterChangeTargetStaff.trim()) {
      window.alert("Please select swap staff.")
      return
    }
    const [targetStaffNo, targetStaffName] = rosterChangeTargetStaff.split("::")
    const conflictStaffNo = rosterChangeConflict?.staffNo || ""
    const conflictStaffName = rosterChangeConflict?.staffName || ""
    const conflictFromCode =
      rosterChangeConflict && rosterChangeCell
        ? (rosterChangeConflict.cells[rosterChangeCell.date] || "").toUpperCase()
        : ""
    const conflictToCode = (rosterChangeConflictToCode || "").trim().toUpperCase()
    const req: RosterChangeRequest = {
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      createdAt: new Date().toISOString(),
      status: "Pending Approval",
      changeType: rosterChangeType,
      date: rosterChangeCell.date,
      staffNo: rosterChangeCell.staffNo,
      staffName: rosterChangeCell.staffName,
      level: rosterChangeCell.level,
      fromCode: rosterChangeCell.code,
      toCode: rosterChangeToCode.trim().toUpperCase(),
      targetStaffNo: (targetStaffNo || "").trim(),
      targetStaffName: (targetStaffName || "").trim(),
      conflictStaffNo,
      conflictStaffName,
      conflictFromCode,
      conflictToCode,
      reason: rosterChangeReason.trim(),
    }
    setRosterChangeRequests((prev) => [req, ...prev])
    const submitLog: RosterChangeLog = {
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      at: new Date().toISOString(),
      requestedBy: authenticatedStaff?.fullName || authenticatedStaff?.staffNo || "System",
      date: req.date,
      staffNo: req.staffNo,
      staffName: req.staffName,
      type: "Roster Manual Change",
      action: "Change Submitted",
      details:
        req.changeType === "Mutual Swap"
          ? `Pending: Swap with ${req.targetStaffNo} - ${req.targetStaffName}`
          : `Pending: ${req.fromCode || "-"} -> ${req.toCode || "-"}`,
      reason: req.reason || "-",
    }
    setRosterChangeLogs((prev) => [submitLog, ...prev].slice(0, 800))
    setIsRosterChangeModalOpen(false)
  }

  const approveRosterChangeRequest = (requestId: string) => {
    const req = rosterChangeRequests.find((r) => r.id === requestId)
    if (!req || req.status !== "Pending Approval") return
    const allowed = getAllowedRosterCodes()
    setStore((prev) => {
      let nextRows = [...prev.roster]
      if (req.changeType === "Mutual Swap") {
        const sourceCurrent = (req.fromCode || "-").toUpperCase()
        const targetCurrent = (req.toCode || "-").toUpperCase()
        nextRows = setRosterCodeForCell(nextRows, req.staffNo, req.staffName, req.date, targetCurrent)
        nextRows = setRosterCodeForCell(
          nextRows,
          req.targetStaffNo,
          req.targetStaffName,
          req.date,
          sourceCurrent,
        )
      } else {
        const nextCode = req.toCode.toUpperCase()
        if (!allowed.has(nextCode)) {
          window.alert(`Cannot approve. Code "${nextCode}" is not valid.`)
          return prev
        }
        if (req.conflictStaffNo && req.conflictStaffName) {
          if (req.conflictToCode) {
            if (!allowed.has(req.conflictToCode)) {
              window.alert(`Cannot approve. Conflict replacement code "${req.conflictToCode}" is not valid.`)
              return prev
            }
            nextRows = setRosterCodeForCell(
              nextRows,
              req.conflictStaffNo,
              req.conflictStaffName,
              req.date,
              req.conflictToCode,
            )
          } else {
            nextRows = setRosterCodeForCell(
              nextRows,
              req.conflictStaffNo,
              req.conflictStaffName,
              req.date,
              "-",
            )
          }
        }
        nextRows = setRosterCodeForCell(nextRows, req.staffNo, req.staffName, req.date, nextCode)
      }
      return { ...prev, roster: nextRows }
    })
    setRosterChangeRequests((prev) =>
      prev.map((r) => (r.id === requestId ? { ...r, status: "Approved" } : r)),
    )
    syncStaffPortalRequestStatus(requestId, "Approved")
    const portalLog: RosterChangeLog = {
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      at: new Date().toISOString(),
      requestedBy: authenticatedStaff?.fullName || authenticatedStaff?.staffNo || "System",
      date: req.date,
      staffNo: req.staffNo,
      staffName: req.staffName,
      type: "Roster Manual Change",
      action: "Approved Change",
      details:
        req.changeType === "Mutual Swap"
          ? `Swap with ${req.targetStaffNo} - ${req.targetStaffName}`
          : `Before: ${req.fromCode || "-"} -> After: ${req.toCode || "-"}`,
      reason: req.reason || "-",
    }
    setRosterChangeLogs((prev) => [portalLog, ...prev].slice(0, 800))
  }

  const rejectRosterChangeRequest = (requestId: string) => {
    const req = rosterChangeRequests.find((r) => r.id === requestId)
    setRosterChangeRequests((prev) =>
      prev.map((r) => (r.id === requestId ? { ...r, status: "Rejected" } : r)),
    )
    syncStaffPortalRequestStatus(requestId, "Rejected")
    if (!req) return
    const rejectLog: RosterChangeLog = {
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      at: new Date().toISOString(),
      requestedBy: authenticatedStaff?.fullName || authenticatedStaff?.staffNo || "System",
      date: req.date,
      staffNo: req.staffNo,
      staffName: req.staffName,
      type: "Roster Manual Change",
      action: "Rejected Change",
      details:
        req.changeType === "Mutual Swap"
          ? `Rejected: Swap with ${req.targetStaffNo} - ${req.targetStaffName}`
          : `Rejected: ${req.fromCode || "-"} -> ${req.toCode || "-"}`,
      reason: req.reason || "-",
    }
    setRosterChangeLogs((prev) => [rejectLog, ...prev].slice(0, 800))
  }

  const authenticatedUser = useMemo(
    () => store.userManagement.find((u) => u.id === authStaffId) ?? null,
    [store.userManagement, authStaffId],
  )
  const authenticatedStaff = useMemo(() => {
    if (!authenticatedUser) return null
    const dispatchToken = (authenticatedUser.dispatchStaff || "").trim()
    if (dispatchToken) {
      const direct =
        store.staff.find(
          (s) => `${(s.staffNo || "").trim()} - ${(s.fullName || "").trim()}` === dispatchToken,
        ) || null
      if (direct) return direct
    }
    const uname = (authenticatedUser.userName || "").trim().toLowerCase()
    return (
      store.staff.find((s) => (s.staffNo || "").trim().toLowerCase() === uname) ||
      store.staff.find((s) => (s.fullName || "").trim().toLowerCase() === uname) ||
      null
    )
  }, [authenticatedUser, store.staff])
  const authenticatedUserAccessFlows = useMemo(() => {
    if (!authenticatedUser) return null
    if (store.userManagement.length === 0) return null
    const roleName = (authenticatedUser.roleName || "").trim()
    if (!roleName) return new Set<string>()
    const roleRow =
      store.roleManagement.find((r) => normalizeText(r.roleName || "") === normalizeText(roleName)) ||
      null
    if (!roleRow) return new Set<string>()
    return new Set(splitMultiValue(roleRow.controlFlows || ""))
  }, [authenticatedUser, store.userManagement, store.roleManagement])
  const isAdminUser = useMemo(() => {
    const role = normalizeText(authenticatedUser?.roleName || "")
    const userType = normalizeText(authenticatedUser?.userType || "")
    return role.includes("admin") || userType.includes("admin")
  }, [authenticatedUser])
  const hasAccess = (flowName: string) => {
    if (isAdminUser) return true
    if (!authenticatedUserAccessFlows) return true
    return authenticatedUserAccessFlows.has(flowName)
  }
  const hasModuleAccess = (moduleKey: ModuleKey) => {
    const flow = MODULE_VIEW_FLOW[moduleKey]
    if (!flow) return true
    return hasAccess(flow)
  }
  const visibleNavGroups = useMemo(
    () =>
      NAV_GROUPS.map((g) => ({
        ...g,
        items: g.items.filter((it) => hasModuleAccess(it.key)),
      })).filter((g) => g.items.length > 0),
    [authenticatedUserAccessFlows],
  )
  const canRosterView = hasAccess("Roster > Roster > View")
  const canRosterGenerate = hasAccess("Roster > Roster > Generate")
  const canRosterDownload = hasAccess("Roster > Roster > Download")
  const canRosterUpload = hasAccess("Roster > Roster > Upload")
  const rosterChangeCodeOptions = useMemo(() => {
    return Array.from(getAllowedRosterCodes())
      .filter((c) => c !== "-")
      .sort()
  }, [store.shift, store.leaveAttendanceControl])
  const rosterChangeTargetStaffOptions = useMemo(() => {
    if (!rosterChangeCell) return []
    const sameLevel = rosterVisibleRows.filter(
      (r) =>
        (r.level || "").trim().toLowerCase() === (rosterChangeCell.level || "").trim().toLowerCase() &&
        !(r.staffNo === rosterChangeCell.staffNo && r.staffName === rosterChangeCell.staffName),
    )
    const other = rosterVisibleRows.filter(
      (r) =>
        (r.level || "").trim().toLowerCase() !== (rosterChangeCell.level || "").trim().toLowerCase() &&
        !(r.staffNo === rosterChangeCell.staffNo && r.staffName === rosterChangeCell.staffName),
    )
    return [...sameLevel, ...other].map((r) => ({
      value: `${r.staffNo}::${r.staffName}`,
      label: `${r.staffNo} - ${r.staffName} (${r.level})`,
    }))
  }, [rosterChangeCell, rosterVisibleRows])

  useEffect(() => {
    if (!hydrated) return

    const importPortalRosterRequests = () => {
      try {
        const raw = window.localStorage.getItem(STAFF_PORTAL_REQUESTS_KEY)
        const localRequests = raw ? (JSON.parse(raw) as StaffPortalRequest[]) : []
        const mirroredRequests = (((store as unknown as Record<string, unknown>).__staffPortalRequests) || []) as StaffPortalRequest[]
        const requests = Array.isArray(localRequests) ? [...localRequests] : []
        if (Array.isArray(mirroredRequests)) {
          const seen = new Set(requests.map((r) => r.id))
          mirroredRequests.forEach((r) => {
            if (!r?.id || seen.has(r.id)) return
            requests.push(r)
            seen.add(r.id)
          })
        }
        if (!Array.isArray(requests)) return
        const pendingRoster = requests.filter(
          (r) => r.type === "Roster Change" && r.status === "Pending Approval" && r.date,
        )

        const existingReqIds = new Set(rosterChangeRequests.map((r) => r.id))
        const existingLogIds = new Set(rosterChangeLogs.map((l) => l.id))
        const toAddReq: RosterChangeRequest[] = []
        const toAddLog: RosterChangeLog[] = []

        if (pendingRoster.length > 0) {
          pendingRoster.forEach((r) => {
            if (!r.id || existingReqIds.has(r.id)) return
            const level =
              store.staff.find(
                (s) =>
                  (s.staffNo || "").trim() === (r.staffNo || "").trim() &&
                  (s.fullName || "").trim() === (r.staffName || "").trim(),
              )?.level || ""
            const mappedReq: RosterChangeRequest = {
              id: r.id,
              createdAt: new Date().toISOString(),
              status: "Pending Approval",
              changeType: r.changeWithStaffNo ? "Mutual Swap" : "Shift Reassign",
              date: r.date || "",
              staffNo: r.staffNo || "",
              staffName: r.staffName || "",
              level,
              fromCode: (r.fromCode || "-").toUpperCase(),
              toCode: (r.toCode || "-").toUpperCase(),
              targetStaffNo: r.changeWithStaffNo || "",
              targetStaffName: r.changeWithStaffName || "",
              conflictStaffNo: "",
              conflictStaffName: "",
              conflictFromCode: "",
              conflictToCode: "",
              reason: r.reason || "Staff portal request",
            }
            toAddReq.push(mappedReq)

            const portalLogId = `portal-submit-${r.id}`
            if (!existingLogIds.has(portalLogId)) {
              toAddLog.push({
                id: portalLogId,
                at: new Date().toISOString(),
                requestedBy: `${r.staffNo || ""}${r.staffName ? ` - ${r.staffName}` : ""}`.trim() || "Staff Portal",
                date: mappedReq.date,
                staffNo: mappedReq.staffNo,
                staffName: mappedReq.staffName,
                type: "Roster Manual Change",
                action: "Change Submitted",
                details:
                  mappedReq.changeType === "Mutual Swap"
                    ? `Pending: Swap with ${mappedReq.targetStaffNo} - ${mappedReq.targetStaffName}`
                    : `Pending: ${mappedReq.fromCode || "-"} -> ${mappedReq.toCode || "-"}`,
                reason: mappedReq.reason || "-",
              })
            }
          })
        }

        if (toAddReq.length > 0) {
          setRosterChangeRequests((prev) => [...toAddReq, ...prev])
        }
        if (toAddLog.length > 0) {
          setRosterChangeLogs((prev) => [...toAddLog, ...prev].slice(0, 800))
        }

        // Import staff-portal leave requests into Staff Leaves queue.
        const leavePolicyNames = new Set(
          store.leaveAttendanceControl
            .filter((x) => normalizeText(x.leaveAttendanceType || "") === "leave")
            .map((x) => (x.leaveAttendanceName || "").trim())
            .filter(Boolean),
        )
        const leaveReqs = requests.filter((r) => {
          if (r.type !== "Leave") return false
          const status = normalizeText(r.status || "")
          if (status === "rejected" || status === "cancelled") return true
          if (status !== "pending approval" && status !== "approved" && status !== "pending") return false
          const policy = (r.leavePolicyName || "").trim()
          // If policy is empty, still import (fallback name is applied below)
          // so portal requests are never silently dropped.
          if (!policy) return true
          // If no leave policy is configured yet, still allow import so queue isn't blocked.
          if (leavePolicyNames.size === 0) return true
          return leavePolicyNames.has(policy)
        })
        if (leaveReqs.length > 0) {
          setStore((prev) => {
            const existingByReqId = new Map<string, Entry>()
            const existingByDedup = new Map<string, Entry>()
            prev.leaves.forEach((l) => {
              const reqId = (l.portalRequestId || "").trim()
              if (reqId) existingByReqId.set(reqId, l)
              const dedup = `${(l.staffNo || "").trim().toLowerCase()}::${normalizeText(
                l.leavePolicyName || "",
              )}::${(l.fromDate || "").trim()}::${(l.toDate || "").trim() || (l.fromDate || "").trim()}`
              existingByDedup.set(dedup, l)
            })

            const upserts = new Map<string, Entry>()
            leaveReqs.forEach((r) => {
              const fromDate = normalizeDateOnly(r.fromDate || "")
              const toDate = normalizeDateOnly(r.toDate || "") || fromDate
              if (!fromDate || !toDate) return
              const resolvedLeavePolicyName =
                (r.leavePolicyName || "").trim() ||
                Array.from(leavePolicyNames)[0] ||
                "Leave"
              const leaveSummary = calculateLeaveSummary(fromDate, toDate, resolvedLeavePolicyName)
              const dedup = `${(r.staffNo || "").trim().toLowerCase()}::${normalizeText(
                resolvedLeavePolicyName,
              )}::${fromDate}::${toDate}`
              const existing =
                (r.id ? existingByReqId.get(r.id) : null) || existingByDedup.get(dedup) || null
              const rowId =
                existing?.id || `leave_portal_${r.id || `${Date.now()}_${Math.random().toString(16).slice(2)}`}`
              const portalStatus = normalizeText(r.status || "")
              const mappedStatus =
                portalStatus === "approved"
                  ? "Approved"
                  : portalStatus === "rejected"
                    ? "Rejected"
                    : portalStatus === "cancelled"
                      ? "Rejected"
                      : "Pending"

              upserts.set(rowId, {
                ...existing,
                id: rowId,
                createdAt: new Date().toISOString(),
                staffNo: r.staffNo || existing?.staffNo || "",
                staffName: r.staffName || existing?.staffName || "",
                leavePolicyName: resolvedLeavePolicyName || existing?.leavePolicyName || "Leave",
                fromDate: leaveSummary.fromDate,
                toDate: leaveSummary.toDate,
                totalCalendarDays: leaveSummary.totalCalendarDays,
                excludedWeekdayDetails: leaveSummary.excludedWeekdayDetails,
                publicHolidayDays: leaveSummary.publicHolidayDays,
                chargeableLeaveDays: leaveSummary.chargeableLeaveDays,
                status: mappedStatus,
                remarks:
                  mappedStatus === "Rejected" && portalStatus === "cancelled"
                    ? "Cancelled by staff from portal"
                    : r.reason || existing?.remarks || "Leave request from staff portal",
                portalRequestId: r.id || existing?.portalRequestId || "",
              })
            })

            if (upserts.size === 0) return prev
            const nextLeaves = prev.leaves.map((l) => upserts.get(l.id) || l)
            upserts.forEach((entry, id) => {
              if (!nextLeaves.some((x) => x.id === id)) nextLeaves.unshift(entry)
            })
            return { ...prev, leaves: nextLeaves }
          })
        }

        // Import staff-portal duty mark requests into Attendance queue.
        const normAttendanceName = (v: string) =>
          (v || "")
            .trim()
            .toLowerCase()
            .replace(/[^a-z0-9]/g, "")
        const attendanceTypeNames = new Set(
          store.leaveAttendanceControl
            .filter((x) => (x.leaveAttendanceType || "").trim().toLowerCase() === "attendance")
            .map((x) => normAttendanceName(x.leaveAttendanceName || ""))
            .filter(Boolean),
        )
        // Fallback common duty-mark names to avoid missing imports due to
        // spelling/spacing variations in policy setup.
        const attendanceFallback = new Set([
          "medical",
          "sickleave",
          "familyresponsibleleave",
          "familyresposbileleave",
        ])
        const attendanceReqs = requests.filter((r) => {
          if (r.type !== "Leave" && r.type !== "Attendance") return false
          const status = (r.status || "").trim().toLowerCase()
          if (
            status !== "pending" &&
            status !== "pending approval" &&
            status !== "pending peer acceptance" &&
            status !== "pending document" &&
            status !== "document submitted"
          ) {
            return false
          }
          const name = normAttendanceName(r.dutyMarkType || r.leavePolicyName || "")
          if (!name) return false
          if (name.includes("medical") || name.includes("sick") || name.includes("family")) {
            return true
          }
          if (attendanceTypeNames.size === 0) return attendanceFallback.has(name)
          return attendanceTypeNames.has(name) || attendanceFallback.has(name)
        })
        const cancelledAttendanceReqIds = new Set(
          requests
            .filter((r) => (r.type === "Leave" || r.type === "Attendance") && r.status === "Cancelled")
            .map((r) => (r.id || "").trim())
            .filter(Boolean),
        )
        if (attendanceReqs.length > 0 || cancelledAttendanceReqIds.size > 0) {
          setStore((prev) => {
            const existingReqIds = new Set(prev.attendance.map((a) => (a.portalRequestId || "").trim()))
            const existingByReqId = new Map<string, Entry>()
            const existingByDedupKey = new Map<string, Entry>()
            prev.attendance.forEach((a) => {
              const id = (a.portalRequestId || "").trim()
              if (id) existingByReqId.set(id, a)
              const dedupKey = buildAttendanceDedupKey({
                staffNo: a.staffNo || "",
                attendanceType: a.attendanceType || a.leavePolicyName || "",
                fromDate: a.fromDate || "",
                toDate: a.toDate || a.fromDate || "",
                requestSource: a.requestSource || "Staff Portal",
              })
              if (dedupKey) existingByDedupKey.set(dedupKey, a)
            })
            const additionsById = new Map<string, Entry>()
            attendanceReqs.forEach((r) => {
              const submittedDate = (r.createdAt || "").slice(0, 10) || toYmd(new Date())
              const attendanceType = r.dutyMarkType || r.leavePolicyName || ""
              const fromDate = r.fromDate || ""
              const toDate = r.toDate || r.fromDate || ""
              const dedupKey = buildAttendanceDedupKey({
                staffNo: r.staffNo || "",
                attendanceType,
                fromDate,
                toDate,
                requestSource: "Staff Portal",
              })
              const existing =
                (r.id ? existingByReqId.get(r.id) : null) ||
                existingByDedupKey.get(dedupKey) ||
                null
              const rowId = existing?.id || `att_portal_${r.id || `${Date.now()}_${Math.random().toString(16).slice(2)}`}`
              const nextRow: Entry = {
                ...existing,
                id: rowId,
                createdAt: new Date().toISOString(),
                date: submittedDate,
                fromDate,
                toDate,
                noOfDays: r.noOfDays || "",
                staffNo: r.staffNo || "",
                staffName: r.staffName || "",
                attendanceType,
                status: "Pending",
                workflowStatus: "Pending",
                approvalStatus: "Pending",
                documentStatus: "Not Required",
                documentName: r.documentName || "",
                requestSource: "Staff Portal",
                portalRequestId: r.id || existing?.portalRequestId || "",
                remarks: r.reason || "Duty mark request from staff portal",
              }
              additionsById.set(nextRow.id, nextRow)
              if (r.id) existingReqIds.add(r.id)
            })
            const updated = prev.attendance.map((a) => {
              const direct = additionsById.get(a.id)
              if (direct) return direct
              const reqId = (a.portalRequestId || "").trim()
              if (!reqId) return a
              if (cancelledAttendanceReqIds.has(reqId)) {
                return {
                  ...a,
                  status: "Cancelled",
                  workflowStatus: "Cancelled",
                  approvalStatus: "Cancelled",
                  remarks: a.remarks || "Cancelled by staff before approval",
                }
              }
              const req = attendanceReqs.find((x) => x.id === reqId)
              if (!req) return a
              const nextStatus =
                req.status === "Pending Document"
                  ? "Pending Document Upload"
                  : req.status === "Document Submitted"
                    ? "Pending Approval"
                    : req.status || a.status
              return {
                ...a,
                status: nextStatus,
                workflowStatus: req.status || a.workflowStatus,
                documentStatus: req.documentName ? "Submitted" : a.documentStatus,
                documentName: req.documentName || a.documentName || "",
                documentData: req.documentData || a.documentData || "",
                remarks: a.remarks || req.reason || "",
              }
            })
            const merged = [...updated]
            additionsById.forEach((row, id) => {
              if (!merged.some((m) => m.id === id)) merged.unshift(row)
            })
            return { ...prev, attendance: merged }
          })
        }
      } catch {
        // Ignore malformed portal request payloads.
      }
    }

    importPortalRosterRequests()
    const intervalId = window.setInterval(importPortalRosterRequests, 3000)
    const onStorage = (event: StorageEvent) => {
      if (event.key === STAFF_PORTAL_REQUESTS_KEY) importPortalRosterRequests()
    }
    // Same-tab navigations don't fire the storage event, so also re-poll
    // whenever the tab becomes visible or regains focus — covers the case
    // of switching from /staff-portal back to / in the same tab.
    const onVisible = () => {
      if (document.visibilityState === "visible") importPortalRosterRequests()
    }
    const onFocus = () => importPortalRosterRequests()
    window.addEventListener("storage", onStorage)
    document.addEventListener("visibilitychange", onVisible)
    window.addEventListener("focus", onFocus)
    return () => {
      window.clearInterval(intervalId)
      window.removeEventListener("storage", onStorage)
      document.removeEventListener("visibilitychange", onVisible)
      window.removeEventListener("focus", onFocus)
    }
  }, [hydrated, rosterChangeRequests, rosterChangeLogs, store.staff, store.leaveAttendanceControl])

  const syncStaffPortalRequestStatus = (
    requestId: string,
    status:
      | "Pending Document"
      | "Pending Approval"
      | "Approved"
      | "Rejected"
      | "Document Submitted"
      | "Cancelled",
  ) => {
    try {
      const raw = window.localStorage.getItem(STAFF_PORTAL_REQUESTS_KEY)
      if (!raw) return
      const requests = JSON.parse(raw) as StaffPortalRequest[]
      if (!Array.isArray(requests)) return
      const next = requests.map((r) => (r.id === requestId ? { ...r, status } : r))
      window.localStorage.setItem(STAFF_PORTAL_REQUESTS_KEY, JSON.stringify(next))
      if (supabase && !disableRemoteSyncRef.current) {
        void (async () => {
          try {
            const { data } = await supabase
              .from(SUPABASE_STORE_TABLE)
              .select("payload")
              .eq("id", SUPABASE_STORE_ID)
              .maybeSingle()
            const payload =
              data?.payload && typeof data.payload === "object"
                ? ({ ...(data.payload as Record<string, unknown>) } as Record<string, unknown>)
                : {}
            const mirror = Array.isArray(payload.__staffPortalRequests)
              ? (payload.__staffPortalRequests as StaffPortalRequest[])
              : []
            const mirroredNext = mirror.map((r) => (r.id === requestId ? { ...r, status } : r))
            payload.__staffPortalRequests = mirroredNext
            await savePayloadJsonToSupabase(JSON.stringify(payload))
          } catch {
            // Non-blocking.
          }
        })()
      }
    } catch {
      // Ignore malformed portal request payloads.
    }
  }

  const markAttendanceSubmitDocument = (entryId: string) => {
    let reqId = ""
    setStore((prev) => ({
      ...prev,
      attendance: prev.attendance.map((a) => {
        if (a.id !== entryId) return a
        reqId = a.portalRequestId || ""
        return {
          ...a,
          workflowStatus: "Pending Document",
          documentStatus: "Requested",
          status: "Pending Document Upload",
        }
      }),
    }))
    if (reqId) syncStaffPortalRequestStatus(reqId, "Pending Document")
  }

  const markAttendanceProceedApproval = (entryId: string) => {
    let reqId = ""
    setStore((prev) => ({
      ...prev,
      attendance: prev.attendance.map((a) => {
        if (a.id !== entryId) return a
        reqId = a.portalRequestId || ""
        return {
          ...a,
          workflowStatus: "Pending Approval",
          approvalStatus: "Pending",
          status: "Pending Approval",
        }
      }),
    }))
    if (reqId) syncStaffPortalRequestStatus(reqId, "Pending Approval")
  }

  const approveAttendanceRequest = (entryId: string) => {
    let reqId = ""
    setStore((prev) => {
      const approvedEntry = prev.attendance.find((a) => a.id === entryId)
      const nextAttendance = prev.attendance.map((a) => {
        if (a.id !== entryId) return a
        reqId = a.portalRequestId || ""
        return {
          ...a,
          status: "Approved",
          workflowStatus: "Approved",
          approvalStatus: "Approved",
        }
      })

      let nextRoster = [...prev.roster]
      if (approvedEntry) {
        const markName = normalizeText(
          approvedEntry.attendanceType || approvedEntry.leavePolicyName || approvedEntry.type || "",
        )
        const resolvedCode = markName.includes("medical")
          ? "MC"
          : markName.includes("sick")
            ? "SL"
            : ""
        const fromDate = (approvedEntry.fromDate || "").trim()
        const toDate = (approvedEntry.toDate || "").trim() || fromDate
        const staffNo = (approvedEntry.staffNo || "").trim()
        const staffName = (approvedEntry.staffName || "").trim()
        if (resolvedCode && fromDate && toDate && staffNo) {
          getDateRangeInclusive(fromDate, toDate).forEach((d) => {
            const ymd = toYmd(d)
            nextRoster = setRosterCodeForCell(nextRoster, staffNo, staffName, ymd, resolvedCode)
          })
        }
      }

      return {
        ...prev,
        attendance: nextAttendance,
        roster: nextRoster,
      }
    })
    if (reqId) syncStaffPortalRequestStatus(reqId, "Approved")
  }

  const rejectAttendanceRequest = (entryId: string) => {
    const reason = window.prompt("Enter reject remarks", "")?.trim() || ""
    let reqId = ""
    setStore((prev) => ({
      ...prev,
      attendance: prev.attendance.map((a) => {
        if (a.id !== entryId) return a
        reqId = a.portalRequestId || ""
        return {
          ...a,
          status: "Rejected",
          workflowStatus: "Rejected",
          approvalStatus: "Rejected",
          remarks: reason || a.remarks || "",
        }
      }),
    }))
    if (reqId) syncStaffPortalRequestStatus(reqId, "Rejected")
  }

  const backfillAttendanceFromPortal = () => {
    try {
      const raw = window.localStorage.getItem(STAFF_PORTAL_REQUESTS_KEY)
      if (!raw) {
        window.alert(
          "No staff portal requests found in this browser's localStorage.\n\n" +
            "If you submitted the duty mark on a different device, the staff " +
            "portal data is not synced across devices. Open the staff portal " +
            "in this same browser and re-submit the duty mark.",
        )
        return
      }
      const requests = JSON.parse(raw) as StaffPortalRequest[]
      if (!Array.isArray(requests)) {
        window.alert("Invalid staff portal request payload.")
        return
      }
      const normalizeName = (v: string) =>
        (v || "")
          .trim()
          .toLowerCase()
          .replace(/[^a-z0-9]/g, "")
      const totalRequests = requests.length
      const leaveRequests = requests.filter((r) => r.type === "Leave")
      const attendanceRequests = requests.filter((r) => r.type === "Attendance")
      const candidates = requests.filter((r) => {
        if (r.type !== "Leave" && r.type !== "Attendance") return false
        const status = (r.status || "").trim().toLowerCase()
        if (status === "approved" || status === "rejected" || status === "cancelled") return false
        const n = normalizeName(r.dutyMarkType || r.leavePolicyName || "")
        return n.includes("medical") || n.includes("sick") || n.includes("family")
      })
      if (candidates.length === 0) {
        window.alert(
          `No pending duty mark requests found for attendance backfill.\n\n` +
            `Diagnostic:\n` +
            `  • Total staff-portal requests in storage: ${totalRequests}\n` +
            `  • Of which type=Leave: ${leaveRequests.length}\n` +
            `  • Of which type=Attendance: ${attendanceRequests.length}\n` +
            `  • Of which match duty-mark types (Medical / Sick / Family): 0\n\n` +
            `If you expected to see duty marks here, check that the request was ` +
            `submitted from the staff portal in this same browser.`,
        )
        return
      }
      let added = 0
      let alreadyImported = 0
      setStore((prev) => {
        const existingReqIds = new Set(prev.attendance.map((a) => (a.portalRequestId || "").trim()))
        const existingByDedupKey = new Map<string, Entry>()
        prev.attendance.forEach((a) => {
          const dedupKey = buildAttendanceDedupKey({
            staffNo: a.staffNo || "",
            attendanceType: a.attendanceType || a.leavePolicyName || "",
            fromDate: a.fromDate || "",
            toDate: a.toDate || a.fromDate || "",
            requestSource: a.requestSource || "Staff Portal",
          })
          if (dedupKey) existingByDedupKey.set(dedupKey, a)
        })
        const additions: Entry[] = []
        candidates.forEach((r) => {
          const dedupKey = buildAttendanceDedupKey({
            staffNo: r.staffNo || "",
            attendanceType: r.dutyMarkType || r.leavePolicyName || "",
            fromDate: r.fromDate || "",
            toDate: r.toDate || r.fromDate || "",
            requestSource: "Staff Portal",
          })
          if ((r.id && existingReqIds.has(r.id)) || existingByDedupKey.has(dedupKey)) {
            alreadyImported += 1
            return
          }
          const submittedDate = (r.createdAt || "").slice(0, 10) || toYmd(new Date())
          const existing = existingByDedupKey.get(dedupKey)
          additions.push({
            ...(existing || {}),
            id: existing?.id || `att_backfill_${r.id || `${Date.now()}_${Math.random().toString(16).slice(2)}`}`,
            createdAt: new Date().toISOString(),
            date: submittedDate,
            fromDate: r.fromDate || "",
            toDate: r.toDate || r.fromDate || "",
            noOfDays: r.noOfDays || "",
            staffNo: r.staffNo || "",
            staffName: r.staffName || "",
            attendanceType: r.dutyMarkType || r.leavePolicyName || "",
            status: "Pending",
            workflowStatus: "Pending",
            approvalStatus: "Pending",
            documentStatus: "Not Required",
            requestSource: "Staff Portal",
            portalRequestId: r.id,
            remarks: r.reason || "Duty mark request backfilled from staff portal",
          })
        })
        added = additions.length
        if (added === 0) return prev
        return { ...prev, attendance: [...additions, ...prev.attendance] }
      })
      window.setTimeout(() => {
        const lines: string[] = []
        lines.push(`Backfill summary:`)
        lines.push(`  • Duty mark candidates in portal: ${candidates.length}`)
        lines.push(`  • Already in attendance (skipped): ${alreadyImported}`)
        lines.push(`  • Newly added to attendance: ${added}`)
        if (added > 0) {
          lines.push("")
          lines.push("Refresh the Attendance grid to see the new entries.")
        }
        window.alert(lines.join("\n"))
      }, 50)
    } catch (err) {
      window.alert(
        `Failed to backfill attendance from staff portal requests.\n\n${
          err instanceof Error ? err.message : String(err)
        }`,
      )
    }
  }
  const rosterChangeConflict = useMemo(() => {
    if (!rosterChangeCell) return null
    if (rosterChangeType === "Mutual Swap") return null
    const nextCode = (rosterChangeToCode || "").trim().toUpperCase()
    if (!nextCode || nextCode === "-" || isOffCode(nextCode) || nextCode === "AL" || nextCode === "PH" || nextCode === "GH") {
      return null
    }
    return (
      rosterVisibleRows.find(
        (r) =>
          (r.cells[rosterChangeCell.date] || "").toUpperCase() === nextCode &&
          !(r.staffNo === rosterChangeCell.staffNo && r.staffName === rosterChangeCell.staffName),
      ) || null
    )
  }, [rosterChangeCell, rosterChangeToCode, rosterChangeType, rosterVisibleRows])

  const login = () => {
    const userName = loginStaffNo.trim()
    const password = loginPassword
    const user = store.userManagement.find(
      (u) =>
        (u.userName || "").trim().toLowerCase() === userName.toLowerCase() &&
        (u.status || "Active").trim().toLowerCase() === "active",
    )
    if (!user) {
      setLoginError("Invalid username or password.")
      return
    }
    const effectivePassword = (user.password || "").trim() || (user.userName || "").trim()
    if (password !== effectivePassword) {
      setLoginError("Invalid username or password.")
      return
    }
    setLoginError("")
    setAuthStaffId(user.id)
    window.localStorage.setItem(AUTH_STORAGE_KEY, user.id)
    setMustChangePassword(false)
  }

  const logout = () => {
    setAuthStaffId(null)
    setMustChangePassword(false)
    setLoginPassword("")
    setNewPassword("")
    setConfirmNewPassword("")
    setPasswordChangeError("")
    window.localStorage.removeItem(AUTH_STORAGE_KEY)
  }

  const changePasswordNow = () => {
    if (!authenticatedStaff) return
    const np = newPassword.trim()
    const cp = confirmNewPassword.trim()
    if (np.length < 4) {
      setPasswordChangeError("New password must be at least 4 characters.")
      return
    }
    if (np !== cp) {
      setPasswordChangeError("Password confirmation does not match.")
      return
    }
    setStore((prev) => ({
      ...prev,
      staff: prev.staff.map((s) =>
        s.id === authenticatedStaff.id
          ? { ...s, loginPassword: np, forcePasswordChange: "No" }
          : s,
      ),
    }))
    setPasswordChangeError("")
    setMustChangePassword(false)
    setNewPassword("")
    setConfirmNewPassword("")
  }

  // Deep-link support: read ?module=<key> on mount and switch to it.
  // Keep hooks above conditional returns to preserve hook order.
  useEffect(() => {
    if (typeof window === "undefined") return
    const params = new URLSearchParams(window.location.search)
    const requested = params.get("module")
    if (!requested) return
    const valid = MODULES.some((m) => m.key === requested)
    if (valid && hasModuleAccess(requested as ModuleKey) && requested !== activeModule) {
      setActiveModule(requested as ModuleKey)
    }
    const requestedCrewTab = params.get("crewTab")
    if (
      requestedCrewTab &&
      ["rotationTypes", "attendanceCodes", "dailyAttendance", "leavePatternManager", "generator", "blocks", "timeline", "conflicts"].includes(
        requestedCrewTab,
      )
    ) {
      setCrewLeavePlannerTab(requestedCrewTab as typeof crewLeavePlannerTab)
    }
    setCrewLeavePlannerFocusMode(params.get("crewView") === "masterOnly")
    // Run once on mount.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    if (hasModuleAccess(activeModule)) return
    const firstAllowed =
      visibleNavGroups.flatMap((g) => g.items).map((i) => i.key)[0] || "dashboard"
    if (firstAllowed !== activeModule) setActiveModule(firstAllowed)
  }, [activeModule, visibleNavGroups])

  useEffect(() => {
    if (typeof window === "undefined") return
    const params = new URLSearchParams(window.location.search)
    const sameModule = params.get("module") === activeModule
    const currentCrewTab = params.get("crewTab") || ""
    const nextCrewTab = activeModule === "crewLeavePlanner" ? crewLeavePlannerTab : ""
    if (sameModule && currentCrewTab === nextCrewTab) return
    params.set("module", activeModule)
    if (activeModule === "crewLeavePlanner") params.set("crewTab", crewLeavePlannerTab)
    else params.delete("crewTab")
    const next = `${window.location.pathname}?${params.toString()}${window.location.hash}`
    window.history.replaceState(null, "", next)
  }, [activeModule, crewLeavePlannerTab])

  if (!hydrated) {
    return <main className="p-4">Loading OCCfloat...</main>
  }

  if (!isAuthReady) {
    return <main className="p-4">Loading authentication...</main>
  }

  if (!authStaffId) {
    return (
      <main className="container py-5" style={{ maxWidth: 480 }}>
        <div className="card border shadow-sm">
          <div className="card-body p-4">
            <h1 className="h4 fw-semibold mb-1">OCC Dispatch Login</h1>
            <p className="text-muted small mb-4">
              Login using User Name and Password from User Management.
            </p>
            <div className="mb-3">
              <label className="form-label small fw-semibold">User Name</label>
              <input
                className="form-control"
                value={loginStaffNo}
                onChange={(e) => setLoginStaffNo(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") login()
                }}
              />
            </div>
            <div className="mb-3">
              <label className="form-label small fw-semibold">Password</label>
              <div className="input-group">
                <input
                  type={showLoginPassword ? "text" : "password"}
                  className="form-control"
                  value={loginPassword}
                  onChange={(e) => setLoginPassword(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") login()
                  }}
                />
                <button
                  type="button"
                  className="btn btn-outline-secondary"
                  onClick={() => setShowLoginPassword((v) => !v)}
                  aria-label={showLoginPassword ? "Hide password" : "Show password"}
                >
                  {showLoginPassword ? "Hide" : "Show"}
                </button>
              </div>
            </div>
            {loginError ? <div className="alert alert-danger py-2 small">{loginError}</div> : null}
            <Button className="w-100" onClick={login}>
              Login
            </Button>
          </div>
        </div>
      </main>
    )
  }

  const syncMap: Record<string, { label: string; cls: string }> = {
    loading: { label: "Loading...", cls: "bg-secondary-subtle text-secondary-emphasis" },
    saving: { label: "Saving...", cls: "bg-warning-subtle text-warning-emphasis" },
    synced: { label: "Synced", cls: "bg-success-subtle text-success-emphasis" },
    error: { label: "Error - Local", cls: "bg-danger-subtle text-danger-emphasis" },
    idle: { label: "Local only", cls: "bg-secondary-subtle text-secondary-emphasis" },
  }
  const syncInfo = syncMap[syncStatus] ?? syncMap.idle

  const switchModule = (key: ModuleKey) => {
    if (!hasModuleAccess(key)) return
    setActiveModule(key)
    setStaffLookupQuery("")
    setEditingEntryId(null)
    setIsFormOpen(false)
    setIsRosterGeneratorOpen(false)
    setIsRosterIOModalOpen(false)
    setIsSidebarOpen(false)
    closeInactivePopup()
  }

  const gridSx = {
    border: 0,
    "& .MuiDataGrid-columnHeaders": { backgroundColor: "var(--bs-tertiary-bg)" },
    "& .MuiDataGrid-cell, & .MuiDataGrid-columnHeader": {
      borderColor: "var(--bs-border-color)",
    },
    "& .MuiDataGrid-row:hover": { backgroundColor: "var(--bs-light-bg-subtle)" },
    "& .MuiDataGrid-footerContainer": { borderTopColor: "var(--bs-border-color)" },
  }

  return (
    <div className="container-fluid p-0">
      <div className="d-flex flex-column flex-lg-row min-vh-100">
        <aside
          className={
            "app-sidebar bg-body border-end order-2 order-lg-1 d-flex" +
            (isSidebarOpen ? " app-sidebar--open" : "") +
            (openGroupTitle ? " app-sidebar--expanded" : "")
          }
          onClick={(e) => e.stopPropagation()}
        >
          <div className="app-rail d-flex flex-column align-items-center py-3 gap-2 border-end">
            <button
              type="button"
              className={
                "app-brand mb-2 border-0 bg-transparent" +
                (activeModule === "dashboard" ? " app-rail__btn--active" : "")
              }
              onClick={() => {
                switchModule("dashboard")
                setOpenGroupTitle(null)
              }}
              title="Dashboard"
              aria-label="Dashboard"
            >
              OC
            </button>
            {visibleNavGroups.map((group) => {
              const GroupIcon = group.icon
              const groupActive = group.items.some((it) => it.key === activeModule)
              const isOpen = openGroupTitle === group.title
              return (
                <button
                  key={group.title}
                  type="button"
                  onClick={() =>
                    setOpenGroupTitle((prev) => (prev === group.title ? null : group.title))
                  }
                  title={group.title}
                  aria-label={group.title}
                  aria-expanded={isOpen}
                  className={
                    "btn btn-sm app-rail__btn d-inline-flex align-items-center justify-content-center" +
                    (isOpen || groupActive ? " app-rail__btn--active" : "")
                  }
                >
                  <GroupIcon size={20} aria-hidden="true" />
                </button>
              )
            })}
            <div className="mt-auto d-flex flex-column align-items-center gap-2 pb-2">
              <button
                type="button"
                className="btn btn-sm app-rail__btn d-inline-flex align-items-center justify-content-center"
                title={
                  authenticatedUser?.fullName ||
                  authenticatedUser?.userName ||
                  authenticatedStaff?.fullName ||
                  authenticatedStaff?.staffNo ||
                  "User"
                }
                aria-label="Current user"
              >
                <User size={18} aria-hidden="true" />
              </button>
              <button
                type="button"
                className="btn btn-sm app-rail__btn d-inline-flex align-items-center justify-content-center"
                title={`Sync: ${syncInfo.label}`}
                aria-label={`Sync status ${syncInfo.label}`}
              >
                <Activity
                  size={18}
                  aria-hidden="true"
                  color={
                    syncStatus === "error"
                      ? "#dc3545"
                      : syncStatus === "saving" || syncStatus === "loading"
                      ? "#f59f00"
                      : "#198754"
                  }
                />
              </button>
              <div title="Toggle theme" aria-label="Toggle theme">
                <ThemeToggle />
              </div>
              <button
                type="button"
                className="btn btn-sm app-rail__btn d-inline-flex align-items-center justify-content-center"
                title="Logout"
                aria-label="Logout"
                onClick={logout}
              >
                <LogOut size={18} aria-hidden="true" />
              </button>
            </div>
          </div>
          {openGroupTitle ? (
            <div className="app-submenu p-3 flex-grow-1">
              {(() => {
                const group = visibleNavGroups.find((g) => g.title === openGroupTitle)
                if (!group) return null
                const GroupIcon = group.icon
                return (
                  <>
                    <div className="d-flex align-items-center gap-2 mb-2">
                      <GroupIcon size={16} aria-hidden="true" />
                      <div className="fw-semibold">{group.title}</div>
                    </div>
                    <ul className="nav nav-pills flex-column">
                      {group.items.map((item) => {
                        const Icon = item.icon
                        const active = activeModule === item.key
                        return (
                          <li key={item.key} className="nav-item">
                            <button
                              type="button"
                              onClick={() => switchModule(item.key)}
                              className={
                                "nav-link w-100 d-flex align-items-center gap-2" +
                                (active ? " active" : "")
                              }
                            >
                              <Icon size={16} aria-hidden="true" />
                              <span className="text-truncate">{item.label}</span>
                            </button>
                          </li>
                        )
                      })}
                    </ul>
                  </>
                )
              })()}
            </div>
          ) : null}
        </aside>

        <main
          className="flex-grow-1 d-flex flex-column min-w-0 order-1 order-lg-2"
          style={{ overflowX: "hidden" }}
          onClick={() => setOpenGroupTitle(null)}
        >
          <div className="flex-grow-1 p-3 p-md-4" style={{ overflowX: "hidden" }}>
            <div className="d-flex d-lg-none align-items-center justify-content-between mb-3">
              <button
                type="button"
                className="btn btn-sm btn-outline-secondary"
                onClick={() => setIsSidebarOpen((v) => !v)}
                aria-label="Toggle navigation"
                aria-expanded={isSidebarOpen}
              >
                <span aria-hidden="true">&#9776;</span>
              </button>
              <div className="small text-muted text-truncate">{activeConfig.title}</div>
            </div>
            {activeConfig.key === "dashboard" ? (
            <div className="row g-3 mb-4">
              {dashboardStats.map((stat) => (
                <div key={stat.label} className="col-12 col-sm-6 col-lg-3">
                  <div className="card stat-card h-100 border">
                    <div className="card-body">
                      <div className="text-uppercase text-muted small fw-semibold mb-2" style={{ fontSize: "0.7rem", letterSpacing: "0.05em" }}>
                        {stat.label}
                      </div>
                      <div className="fs-2 fw-semibold">{stat.value}</div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
            ) : null}

            <div>
                {activeConfig.key === "dashboard" ? (
                  <div className="row g-3">
                    <div className="col-12 col-md-4">
                      <div className="card h-100 border">
                        <div className="card-body">
                          <h3 className="h6 fw-semibold mb-2">Staff</h3>
                          <p className="text-muted small mb-3">Manage staff, levels, leave types and holidays.</p>
                          <Button onClick={() => switchModule("staff")}>Open Staff</Button>
                        </div>
                      </div>
                    </div>
                    <div className="col-12 col-md-4">
                      <div className="card h-100 border">
                        <div className="card-body">
                          <h3 className="h6 fw-semibold mb-2">Roster</h3>
                          <p className="text-muted small mb-3">Generate and review shifts, rest, and priorities.</p>
                          <Button onClick={() => switchModule("roster")}>Open Roster</Button>
                        </div>
                      </div>
                    </div>
                    <div className="col-12 col-md-4">
                      <div className="card h-100 border">
                        <div className="card-body">
                          <h3 className="h6 fw-semibold mb-2">Operations</h3>
                          <p className="text-muted small mb-3">Track attendance, checklist and briefing updates.</p>
                          <Button onClick={() => switchModule("attendance")}>Open Operations</Button>
                        </div>
                      </div>
                    </div>
                  </div>
                ) : activeConfig.key === "roster" ? (
                  !canRosterView ? (
                    <div className="alert alert-warning mb-0">
                      You do not have access to view Roster module.
                    </div>
                  ) : (
                  <div className="d-flex flex-column gap-4">
                    <div className="d-flex flex-wrap align-items-center gap-2">
                      <Button onClick={() => setIsRosterGeneratorOpen(true)} disabled={!canRosterGenerate}>
                        Open Roster Generator
                      </Button>
                      <Button
                        variant="outline"
                        onClick={() => setIsRosterIOModalOpen(true)}
                        disabled={!canRosterDownload && !canRosterUpload}
                      >
                        Roster Import / Export
                      </Button>
                      <span className="small text-muted">
                        Logic: skips future leave blocks while generating.
                      </span>
                      <div className="form-check form-switch ms-1 mb-0">
                        <input
                          id="roster-direct-edit-mode"
                          type="checkbox"
                          className="form-check-input"
                          checked={rosterDirectEditMode}
                          onChange={(e) => setRosterDirectEditMode(e.target.checked)}
                        />
                        <label className="form-check-label small" htmlFor="roster-direct-edit-mode">
                          Edit Mode Active
                        </label>
                      </div>
                    </div>

                    <div className="card border">
                      <div className="card-header bg-body-tertiary d-flex flex-wrap align-items-center justify-content-between gap-2 py-2">
                        <div className="d-flex align-items-center gap-2">
                          <span className="rounded-circle bg-primary" style={{ width: 8, height: 8, display: "inline-block" }} />
                          <span className="fw-semibold small">Roster Compute Log</span>
                          <span className="badge text-bg-light border">{rosterComputeLogsInRange.length} entries</span>
                        </div>
                        <div className="d-flex flex-wrap align-items-center gap-2">
                          <input
                            type="date"
                            className="form-control form-control-sm"
                            style={{ width: 150 }}
                            value={rosterComputeRangeDraft.fromDate}
                            onChange={(e) =>
                              setRosterComputeRangeDraft((p) => ({ ...p, fromDate: e.target.value }))
                            }
                          />
                          <input
                            type="date"
                            className="form-control form-control-sm"
                            style={{ width: 150 }}
                            value={rosterComputeRangeDraft.toDate}
                            onChange={(e) =>
                              setRosterComputeRangeDraft((p) => ({ ...p, toDate: e.target.value }))
                            }
                          />
                          <Button
                            size="sm"
                            onClick={() => {
                              if (
                                rosterComputeRangeDraft.fromDate &&
                                rosterComputeRangeDraft.toDate &&
                                rosterComputeRangeDraft.fromDate > rosterComputeRangeDraft.toDate
                              ) {
                                window.alert('"To Date" must be after or equal to "From Date".')
                                return
                              }
                              setRosterComputeRange(rosterComputeRangeDraft)
                              if (!rosterComputeRangeDraft.fromDate || !rosterComputeRangeDraft.toDate) {
                                return
                              }
                              const rangeDates = getDateRangeInclusive(
                                rosterComputeRangeDraft.fromDate,
                                rosterComputeRangeDraft.toDate,
                              ).map((d) => toYmd(d))
                              const weekdayMap = [
                                "Sunday",
                                "Monday",
                                "Tuesday",
                                "Wednesday",
                                "Thursday",
                                "Friday",
                                "Saturday",
                              ]
                              const requiredRules: Array<{
                                code: string
                                requiredDaily: boolean
                                weekdays: Set<string>
                              }> = []
                              store.rosterPriorityType.forEach((entry) => {
                                const priorityType = normalizeText(entry.priorityType || "")
                                const isRequiredPriority =
                                  priorityType === "mandatory" || priorityType === "required"
                                if (!isRequiredPriority) return
                                const reqDaily = normalizeText(entry.requiredDaily || "") === "yes"
                                const weekdays = new Set(splitMultiValue(entry.requiredWeekdays || ""))
                                splitMultiValue(entry.shiftCodes || "").forEach((code) => {
                                  const c = (code || "").trim().toUpperCase()
                                  if (!c) return
                                  requiredRules.push({
                                    code: c,
                                    requiredDaily: reqDaily,
                                    weekdays,
                                  })
                                })
                              })

                              const nextLogs: RosterComputeLog[] = []
                              rangeDates.forEach((ymd) => {
                                const dayDate = new Date(`${ymd}T00:00:00`)
                                const weekdayName = weekdayMap[dayDate.getDay()]
                                const rowsForDate = store.roster.filter((r) => (r.date || "").trim() === ymd)
                                const dutyCount = rowsForDate.filter((r) => {
                                  const code = (r.shiftCode || "").trim().toUpperCase()
                                  return code && code !== "-" && !isOffCode(code)
                                }).length
                                nextLogs.push({
                                  id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
                                  at: new Date().toISOString(),
                                  date: ymd,
                                  level: "ALL",
                                  staff: "-",
                                  code: "-",
                                  severity: "info",
                                  message: `Coverage summary: ${rowsForDate.length} assignments, ${dutyCount} duty shifts`,
                                })

                                requiredRules.forEach((rule) => {
                                  const shouldCheck =
                                    rule.requiredDaily || rule.weekdays.has(weekdayName)
                                  if (!shouldCheck) return
                                  const hasCode = rowsForDate.some(
                                    (r) => (r.shiftCode || "").trim().toUpperCase() === rule.code,
                                  )
                                  if (!hasCode) {
                                    nextLogs.push({
                                      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
                                      at: new Date().toISOString(),
                                      date: ymd,
                                      level: "ALL",
                                      staff: "-",
                                      code: rule.code,
                                      severity: "error",
                                      message: `Required coverage missing for code ${rule.code}`,
                                    })
                                  }
                                })
                              })

                              if (nextLogs.length === 0) {
                                nextLogs.push({
                                  id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
                                  at: new Date().toISOString(),
                                  date: "-",
                                  level: "ALL",
                                  staff: "-",
                                  code: "-",
                                  severity: "info",
                                  message: "No conflicts found in selected range.",
                                })
                              }
                              setRosterComputeLogs(nextLogs)
                            }}
                          >
                            Compute
                          </Button>
                          <Button variant="ghost" size="sm" onClick={() => setRosterComputeLogs([])}>
                            Clear Log
                          </Button>
                        </div>
                      </div>
                      {rosterComputeLogsInRange.length === 0 ? (
                        <div className="card-body small text-muted">
                          No compute logs yet. Generate roster to view conflicts and warnings.
                        </div>
                      ) : (
                        <div style={{ height: "14rem" }}>
                          <DataGrid
                            rows={rosterComputeLogsInRange.map((log) => ({
                              id: log.id,
                              date: log.date && log.date !== "-" ? formatYmdToDdMmYy(log.date) : "-",
                              level: log.level || "-",
                              staff: log.staff || "-",
                              code: log.code || "-",
                              severity: log.severity,
                              message: log.message || "-",
                            }))}
                            columns={[
                              { field: "date", headerName: "Date", width: 96 },
                              { field: "level", headerName: "Level", width: 112 },
                              { field: "staff", headerName: "Staff", width: 176 },
                              { field: "code", headerName: "Code", width: 96 },
                              {
                                field: "severity",
                                headerName: "Type",
                                width: 100,
                                renderCell: (params: GridRenderCellParams) => (
                                  <span
                                    className={
                                      "badge " +
                                      (params.value === "error"
                                        ? "bg-danger-subtle text-danger-emphasis"
                                        : params.value === "warn"
                                          ? "bg-warning-subtle text-warning-emphasis"
                                          : "bg-info-subtle text-info-emphasis")
                                    }
                                  >
                                    {String(params.value || "").toUpperCase()}
                                  </span>
                                ),
                              },
                              { field: "message", headerName: "Message", minWidth: 420, flex: 1 },
                            ]}
                            hideFooter
                            disableRowSelectionOnClick
                            rowHeight={34}
                            sx={gridSx}
                          />
                        </div>
                      )}
                    </div>

                    <div className="card border">
                      <div className="card-header bg-body-tertiary d-flex flex-wrap align-items-center justify-content-between gap-2 py-2">
                        <div className="d-flex align-items-center gap-2">
                          <span className="rounded-circle bg-success" style={{ width: 8, height: 8, display: "inline-block" }} />
                          <span className="fw-semibold small">Roster Change Log</span>
                          <span className="badge text-bg-light border">{rosterChangeLogs.length} entries</span>
                        </div>
                        <Button variant="ghost" size="sm" onClick={() => setRosterChangeLogs([])}>
                          Clear Log
                        </Button>
                      </div>
                      {rosterChangeLogs.length === 0 ? (
                        <div className="card-body small text-muted">
                          No roster change logs yet.
                        </div>
                      ) : (
                        <div style={{ height: "14rem" }}>
                          <DataGrid
                            rows={rosterChangeLogs.map((log) => ({
                              id: log.id,
                              at: formatDateText(log.at),
                              requestedBy: log.requestedBy,
                              date: formatYmdToDdMmYy(log.date),
                              staffNo: log.staffNo,
                              staffName: log.staffName,
                              type: log.type,
                              action: log.action,
                              details: log.details,
                              reason: log.reason,
                            }))}
                            columns={[
                              { field: "at", headerName: "Timestamp", width: 170 },
                              { field: "requestedBy", headerName: "Requested User", width: 180 },
                              { field: "date", headerName: "Date", width: 96 },
                              { field: "staffNo", headerName: "Staff No", width: 120 },
                              { field: "staffName", headerName: "Staff Name", width: 220 },
                              {
                                field: "type",
                                headerName: "Type",
                                width: 120,
                                renderCell: (params: GridRenderCellParams) => {
                                  const v = String(params.value || "")
                                  const cls =
                                    v === "Manual"
                                      ? "bg-info-subtle text-info-emphasis"
                                      : v === "Upload"
                                        ? "bg-primary-subtle text-primary-emphasis"
                                        : "bg-secondary-subtle text-secondary-emphasis"
                                  return <span className={`badge ${cls}`}>{v}</span>
                                },
                              },
                              {
                                field: "action",
                                headerName: "Action",
                                width: 110,
                                renderCell: (params: GridRenderCellParams) => (
                                  <span
                                    className={
                                      "badge " +
                                      (params.value === "Added"
                                        ? "bg-success-subtle text-success-emphasis"
                                        : params.value === "Removed"
                                          ? "bg-danger-subtle text-danger-emphasis"
                                          : "bg-warning-subtle text-warning-emphasis")
                                    }
                                  >
                                    {String(params.value || "")}
                                  </span>
                                ),
                              },
                              { field: "details", headerName: "Change Details", minWidth: 360, flex: 1 },
                              { field: "reason", headerName: "Reason", minWidth: 220, flex: 1 },
                            ]}
                            hideFooter
                            disableRowSelectionOnClick
                            rowHeight={34}
                            sx={gridSx}
                          />
                        </div>
                      )}
                    </div>

                    {rosterDates.length === 0 || rosterMatrixRows.length === 0 ? (
                      <div className="empty-state">
                        <div className="fw-semibold mb-1">No roster data yet</div>
                        <div className="small mb-3">
                          Pick a date range above and run the roster generator to populate the grid.
                        </div>
                        <Button onClick={() => setIsRosterGeneratorOpen(true)}>
                          Open Roster Generator
                        </Button>
                      </div>
                    ) : (
                    <>
                    {rosterDistribution.codes.length > 0 ? (
                      <div className="d-flex flex-wrap align-items-center gap-2">
                        <span className="small fw-semibold text-muted me-1">
                          Distribution Filter:
                        </span>
                        {rosterDistribution.codes.map((code) => {
                          const excluded = rosterDistributionExcluded.has(code)
                          let badgeStyle: React.CSSProperties = {}
                          if (code === "OF") {
                            badgeStyle = {
                              background: ROSTER_OFF_COLOR.bg,
                              color: ROSTER_OFF_COLOR.text,
                            }
                          } else if (code === "AL") {
                            badgeStyle = { background: ROSTER_AL_COLOR.bg, color: ROSTER_AL_COLOR.text }
                          } else if (code === "GH") {
                            badgeStyle = {
                              background: ROSTER_GH_COLOR.bg,
                              color: ROSTER_GH_COLOR.text,
                            }
                          } else if (code === "PH") {
                            badgeStyle = {
                              background: ROSTER_PH_COLOR.bg,
                              color: ROSTER_PH_COLOR.text,
                            }
                          } else {
                            const color = rosterShiftColorsByCode.get(code)
                            if (color) {
                              badgeStyle = { background: color.bg, color: color.text }
                            }
                          }
                          return (
                            <button
                              key={`dist-filter-${code}`}
                              type="button"
                              onClick={() => {
                                setRosterDistributionExcluded((prev) => {
                                  const next = new Set(prev)
                                  if (next.has(code)) next.delete(code)
                                  else next.add(code)
                                  return next
                                })
                              }}
                              className={
                                "btn btn-sm border code-cell" +
                                (excluded ? " opacity-50" : "")
                              }
                              style={{
                                ...badgeStyle,
                                textDecoration: excluded ? "line-through" : "none",
                                cursor: "pointer",
                              }}
                              title={excluded ? `Show ${code}` : `Hide ${code}`}
                            >
                              {code}
                            </button>
                          )
                        })}
                        {rosterDistributionExcluded.size > 0 ? (
                          <button
                            type="button"
                            className="btn btn-sm btn-link p-0 ms-2 small"
                            onClick={() => setRosterDistributionExcluded(new Set())}
                          >
                            Show all
                          </button>
                        ) : null}
                      </div>
                    ) : null}
                    <div className="card border">
                      <div ref={rosterGridWrapRef} style={{ height: "calc(100vh - 300px)", minHeight: 560 }}>
                        <DataGrid
                          rows={rosterVisibleRows.map((row) => ({
                            id: row.staffId,
                            staffNo: row.staffNo,
                            staffName: row.staffName,
                            level: row.level,
                            leaveDays: rosterDates.reduce((acc, d) => {
                              const code = (row.cells[d] || "").toUpperCase()
                              return acc + (code === "AL" || code === "GH" || code === "PH" ? 1 : 0)
                            }, 0),
                            ...row.cells,
                          }))}
                          columns={[
                            {
                              field: "staffNo",
                              headerName: "Staff No",
                              width: 120,
                              sortable: true,
                              filterable: true,
                              headerClassName: "roster-sticky-1-header",
                              cellClassName: "roster-sticky-1-cell",
                            },
                            {
                              field: "staffName",
                              headerName: "Staff Name",
                              width: 240,
                              sortable: true,
                              filterable: true,
                              headerClassName: "roster-sticky-2-header",
                              cellClassName: "roster-sticky-2-cell",
                            },
                            {
                              field: "level",
                              headerName: "Level",
                              width: 120,
                              sortable: true,
                              filterable: true,
                              headerClassName: "roster-sticky-3-header",
                              cellClassName: "roster-sticky-3-cell",
                              renderCell: (params: GridRenderCellParams) => (
                                <span className="badge text-bg-light border">{String(params.value || "-")}</span>
                              ),
                            },
                            { field: "leaveDays", headerName: "Leave", width: 90, sortable: true, filterable: true },
                            ...rosterDates.map((date) => ({
                              field: date,
                              headerName: formatYmdToDdMmYy(date),
                              width: 62,
                              sortable: false,
                              filterable: false,
                              disableColumnMenu: true,
                              editable: true,
                              renderHeader: () => (
                                <div className="roster-date-vertical w-100">
                                  <span>{formatYmdToDdMmYy(date)}</span>
                                  <span className="text-muted" style={{ fontSize: "0.62rem" }}>
                                    {formatYmdToWeekdayShort(date)}
                                  </span>
                                </div>
                              ),
                              renderCell: (params: GridRenderCellParams) => {
                                const code = String(params.value || "-")
                                const rowStaffNo = String(params.row.staffNo || "").trim()
                                const rowStaffName = String(params.row.staffName || "").trim()
                                const rowLevel = String(params.row.level || "")
                                const isDutyMarkLocked = pendingDutyMarkLockKeys.has(`${rowStaffNo}::${date}`)
                                const lockTitle = "Duty mark pending approval. Roster code is locked."
                                if (code === "-") return <span className="text-body-tertiary">-</span>
                                if (isOffCode(code)) {
                                  const badge = (
                                    <span className="code-cell" style={{ background: ROSTER_OFF_COLOR.bg, color: ROSTER_OFF_COLOR.text }}>OF</span>
                                  )
                                  if (rosterDirectEditMode || isDutyMarkLocked) {
                                    return <span title={isDutyMarkLocked ? lockTitle : undefined}>{badge}</span>
                                  }
                                  return (
                                    <button
                                      type="button"
                                      className="border-0 bg-transparent p-0"
                                      onClick={() =>
                                        openRosterChangePopup({
                                          date,
                                          staffNo: rowStaffNo,
                                          staffName: rowStaffName,
                                          level: rowLevel,
                                          code: "OF",
                                        })
                                      }
                                    >
                                      {badge}
                                    </button>
                                  )
                                }
                                if (code === "AL") {
                                  const badge = <span className="code-cell" style={{ background: ROSTER_AL_COLOR.bg, color: ROSTER_AL_COLOR.text }}>{code}</span>
                                  if (rosterDirectEditMode || isDutyMarkLocked) {
                                    return <span title={isDutyMarkLocked ? lockTitle : undefined}>{badge}</span>
                                  }
                                  return (
                                    <button
                                      type="button"
                                      className="border-0 bg-transparent p-0"
                                      onClick={() =>
                                        openRosterChangePopup({
                                          date,
                                          staffNo: rowStaffNo,
                                          staffName: rowStaffName,
                                          level: rowLevel,
                                          code,
                                        })
                                      }
                                    >
                                      {badge}
                                    </button>
                                  )
                                }
                                if (code === "GH") {
                                  const badge = <span className="code-cell" style={{ background: ROSTER_GH_COLOR.bg, color: ROSTER_GH_COLOR.text }}>{code}</span>
                                  if (rosterDirectEditMode || isDutyMarkLocked) {
                                    return <span title={isDutyMarkLocked ? lockTitle : undefined}>{badge}</span>
                                  }
                                  return (
                                    <button
                                      type="button"
                                      className="border-0 bg-transparent p-0"
                                      onClick={() =>
                                        openRosterChangePopup({
                                          date,
                                          staffNo: rowStaffNo,
                                          staffName: rowStaffName,
                                          level: rowLevel,
                                          code,
                                        })
                                      }
                                    >
                                      {badge}
                                    </button>
                                  )
                                }
                                if (code === "PH") {
                                  const badge = <span className="code-cell" style={{ background: ROSTER_PH_COLOR.bg, color: ROSTER_PH_COLOR.text }}>{code}</span>
                                  if (rosterDirectEditMode || isDutyMarkLocked) {
                                    return <span title={isDutyMarkLocked ? lockTitle : undefined}>{badge}</span>
                                  }
                                  return (
                                    <button
                                      type="button"
                                      className="border-0 bg-transparent p-0"
                                      onClick={() =>
                                        openRosterChangePopup({
                                          date,
                                          staffNo: rowStaffNo,
                                          staffName: rowStaffName,
                                          level: rowLevel,
                                          code,
                                        })
                                      }
                                    >
                                      {badge}
                                    </button>
                                  )
                                }
                                const color = rosterShiftColorsByCode.get(code)
                                if (!color) return code
                                const badge = <span className="code-cell" style={{ background: color.bg, color: color.text }}>{code}</span>
                                if (rosterDirectEditMode || isDutyMarkLocked) {
                                  return <span title={isDutyMarkLocked ? lockTitle : undefined}>{badge}</span>
                                }
                                return (
                                  <button
                                    type="button"
                                    className="border-0 bg-transparent p-0"
                                    onClick={() =>
                                      openRosterChangePopup({
                                        date,
                                        staffNo: rowStaffNo,
                                        staffName: rowStaffName,
                                        level: rowLevel,
                                        code,
                                      })
                                    }
                                  >
                                    {badge}
                                  </button>
                                )
                              },
                            })) as GridColDef[],
                          ]}
                          disableRowSelectionOnClick
                          disableVirtualization
                          isCellEditable={(params) => {
                            if (!rosterDirectEditMode) return false
                            const field = String(params.field || "")
                            if (!/^\d{4}-\d{2}-\d{2}$/.test(field)) return false
                            const staffNo = String(params.row?.staffNo || "").trim()
                            if (pendingDutyMarkLockKeys.has(`${staffNo}::${field}`)) return false
                            const today = toYmd(new Date())
                            return field > today
                          }}
                          processRowUpdate={(newRow, oldRow) => {
                            if (!rosterDirectEditMode) return oldRow
                            const allowedCodes = getAllowedRosterCodes()
                            const staffNo = String(newRow.staffNo || "")
                            const staffName = String(newRow.staffName || "")
                            rosterDates.forEach((date) => {
                              const prevCode = String(oldRow?.[date] || "-").toUpperCase()
                              const nextCode = String(newRow?.[date] || "-").toUpperCase()
                              if (prevCode !== nextCode) {
                                if (!allowedCodes.has(nextCode)) {
                                  const allowedList = Array.from(allowedCodes).sort().join(", ")
                                  window.alert(
                                    `Invalid code "${nextCode}" for ${formatYmdToDdMmYy(date)}.\nAllowed codes: ${allowedList}`,
                                  )
                                  throw new Error(`Invalid roster code: ${nextCode}`)
                                }
                                applyRosterCellEdit(staffNo, staffName, date, nextCode)
                              }
                            })
                            return newRow
                          }}
                          onProcessRowUpdateError={() => {}}
                          rowHeight={34}
                          sx={{
                            ...gridSx,
                            "& .MuiDataGrid-cell": {
                              borderColor: "var(--bs-border-color)",
                              alignItems: "flex-end",
                              paddingBottom: "4px",
                            },
                            "& .MuiDataGrid-columnHeader[data-field^='20'], & .MuiDataGrid-cell[data-field^='20']": {
                              paddingLeft: "6px",
                              paddingRight: "6px",
                            },
                          }}
                        />
                      </div>
                    </div>
                    {rosterDistribution.codes.length > 0 ? (
                      <div className="card border">
                        <div className="card-header bg-body-tertiary py-2 small fw-semibold">
                          Distribution Count
                        </div>
                        <div style={{ height: 260 }}>
                          <DataGrid
                            rows={rosterDistribution.codes
                              .filter((c) => !rosterDistributionExcluded.has(c))
                              .map((code, distIndex) => {
                                const dateMap = rosterDistribution.counts.get(code)
                                const total = rosterDates.reduce((acc, date) => acc + (dateMap?.get(date) || 0), 0)
                                const row: Record<string, string | number> = { id: code, code, total, __idx: distIndex }
                                rosterDates.forEach((date) => {
                                  row[date] = dateMap?.get(date) || 0
                                })
                                return row
                              })}
                            columns={[
                              {
                                field: "code",
                                headerName: "Code",
                                minWidth: 120,
                                renderCell: (params: GridRenderCellParams) => {
                                  const code = String(params.value || "-")
                                  let badgeStyle: React.CSSProperties = {}
                                  if (code === "OF") badgeStyle = { background: ROSTER_OFF_COLOR.bg, color: ROSTER_OFF_COLOR.text }
                                  else if (code === "AL") {
                                    badgeStyle = { background: ROSTER_AL_COLOR.bg, color: ROSTER_AL_COLOR.text }
                                  } else if (code === "GH") badgeStyle = { background: ROSTER_GH_COLOR.bg, color: ROSTER_GH_COLOR.text }
                                  else if (code === "PH") badgeStyle = { background: ROSTER_PH_COLOR.bg, color: ROSTER_PH_COLOR.text }
                                  else {
                                    const color = rosterShiftColorsByCode.get(code)
                                    if (color) badgeStyle = { background: color.bg, color: color.text }
                                  }
                                  return <span className="code-cell" style={badgeStyle}>{code}</span>
                                },
                              },
                              { field: "total", headerName: "Total", width: 90 },
                              ...rosterDates.map((date) => ({
                                field: date,
                                headerName: formatYmdToWeekdayShort(date),
                                width: 88,
                                sortable: false,
                                filterable: false,
                                disableColumnMenu: true,
                              })) as GridColDef[],
                            ]}
                            disableRowSelectionOnClick
                            hideFooter
                            rowHeight={34}
                            sx={gridSx}
                          />
                        </div>
                      </div>
                    ) : null}
                    </>
                    )}
                  </div>
                  )
                ) : activeConfig.key === "rosterApprovals" ? (
                  <div className="d-flex flex-column gap-3">
                    <div className="card border">
                      <div className="card-header bg-body-tertiary py-2 small fw-semibold">
                        Roster Change Requests
                      </div>
                      <div style={{ height: "calc(100vh - 260px)", minHeight: 560 }}>
                        <DataGrid
                          rows={rosterChangeRequests}
                          columns={[
                            {
                              field: "createdAt",
                              headerName: "Requested",
                              width: 150,
                              valueFormatter: (params) => formatDateText(String(params || "")),
                            },
                            {
                              field: "status",
                              headerName: "Status",
                              width: 140,
                              renderCell: (params: GridRenderCellParams) => {
                                const status = String(params.value || "")
                                const cls =
                                  status === "Approved"
                                    ? "bg-success-subtle text-success-emphasis"
                                    : status === "Rejected"
                                      ? "bg-danger-subtle text-danger-emphasis"
                                      : "bg-warning-subtle text-warning-emphasis"
                                return <span className={`badge ${cls}`}>{status || "-"}</span>
                              },
                            },
                            { field: "changeType", headerName: "Type", width: 140 },
                            {
                              field: "date",
                              headerName: "Date",
                              width: 100,
                              valueFormatter: (params) => formatYmdToDdMmYy(String(params || "")),
                            },
                            { field: "staffNo", headerName: "Staff No", width: 100 },
                            { field: "staffName", headerName: "Staff Name", minWidth: 180, flex: 1 },
                            { field: "fromCode", headerName: "From", width: 80 },
                            { field: "toCode", headerName: "To", width: 80 },
                            { field: "reason", headerName: "Reason", minWidth: 200, flex: 1 },
                            {
                              field: "__actions",
                              headerName: "Actions",
                              width: 190,
                              sortable: false,
                              filterable: false,
                              renderCell: (params: GridRenderCellParams) => {
                                const r = params.row as RosterChangeRequest
                                if (r.status !== "Pending Approval") return <span className="text-muted small">-</span>
                                return (
                                  <div className="d-flex gap-2">
                                    <Button size="sm" onClick={() => approveRosterChangeRequest(r.id)}>Approve</Button>
                                    <Button size="sm" variant="destructive" onClick={() => rejectRosterChangeRequest(r.id)}>Reject</Button>
                                  </div>
                                )
                              },
                            },
                          ]}
                          disableRowSelectionOnClick
                          rowHeight={34}
                          sx={gridSx}
                          pageSizeOptions={[10, 20, 50, 100]}
                          initialState={{
                            pagination: {
                              paginationModel: { pageSize: 20, page: 0 },
                            },
                          }}
                        />
                      </div>
                    </div>
                  </div>
                ) : activeConfig.key === "opsControl" ? (
                  <OpsControlTimeline
                    flights={store.opsControl as unknown as FlightEntry[]}
                    setFlights={(next) =>
                      setStore((prev) => ({
                        ...prev,
                        opsControl: next as unknown as Entry[],
                      }))
                    }
                  />
                ) : activeConfig.key === "mxTypeManager" ? (
                  <OpsControlTimeline
                    view="mxTypeManager"
                    flights={store.opsControl as unknown as FlightEntry[]}
                    setFlights={(next) =>
                      setStore((prev) => ({
                        ...prev,
                        opsControl: next as unknown as Entry[],
                      }))
                    }
                  />
                ) : activeConfig.key === "aircraftReport" ? (
                  <AircraftReportModule
                    flights={store.opsControl as unknown as FlightEntry[]}
                    setFlights={(next) =>
                      setStore((prev) => ({
                        ...prev,
                        opsControl: next as unknown as Entry[],
                      }))
                    }
                  />
                ) : activeConfig.key === "crewLeavePlanner" ? (
                  <div className="d-flex flex-column gap-3">
                    {!crewLeavePlannerFocusMode ? (
                    <ul className="nav nav-tabs">
                      {[
                        { key: "rotationTypes", label: "Rotation Types" },
                        { key: "attendanceCodes", label: "Attendance Codes" },
                        { key: "dutyTypes", label: "Duty Type Manager" },
                        { key: "dutyCodes", label: "Duty Codes" },
                        { key: "dailyAttendance", label: "Daily Attendance Record" },
                        { key: "leavePatternManager", label: "Leave Pattern Manager" },
                        { key: "generator", label: "Generator" },
                        { key: "blocks", label: "Crew Leave Pattern" },
                        { key: "timeline", label: "Master Roster" },
                        { key: "conflicts", label: "Conflicts" },
                      ].map((tab) => {
                        const isConflicts = tab.key === "conflicts"
                        const conflictCount = isConflicts ? crewLeaveConflicts.length : 0
                        return (
                          <li key={tab.key} className="nav-item">
                            <button
                              type="button"
                              className={"nav-link" + (crewLeavePlannerTab === tab.key ? " active" : "")}
                              onClick={() =>
                                setCrewLeavePlannerTab(
                                  tab.key as
                                    | "rotationTypes"
                                    | "attendanceCodes"
                                    | "dutyTypes"
                                    | "dutyCodes"
                                    | "dailyAttendance"
                                    | "leavePatternManager"
                                    | "generator"
                                    | "blocks"
                                    | "timeline"
                                    | "conflicts",
                                )
                              }
                            >
                              {tab.label}
                              {isConflicts && conflictCount > 0 ? (
                                <span
                                  className="badge rounded-pill bg-danger ms-2"
                                  style={{ fontSize: "0.7rem" }}
                                >
                                  {conflictCount}
                                </span>
                              ) : null}
                            </button>
                          </li>
                        )
                      })}
                    </ul>
                    ) : null}

                    {crewLeavePlannerTab === "rotationTypes" ? (
                      <>
                        <div className="card border">
                          <div className="card-header bg-body-tertiary py-2 small fw-semibold d-flex align-items-center justify-content-between">
                            <span>Rotation Types</span>
                            <Button
                              size="sm"
                              onClick={() => {
                                setEditingRotationId(null)
                                setCrewLeavePolicyForm({
                                  policyName: "",
                                  crewType: "Captain",
                                  workPatternDays: "90",
                                  leavePatternDays: "30",
                                  codeThresholdsJson: "{}",
                                  activeStatus: "Active",
                                })
                                setCrewLeavePolicyThresholds(
                                  crewOpsDeductionCodes.reduce((acc, row) => {
                                    acc[row.code] = "0"
                                    return acc
                                  }, {} as Record<string, string>),
                                )
                                setIsRotationModalOpen(true)
                              }}
                            >
                              + New Rotation
                            </Button>
                          </div>
                          <div className="card-body">
                            <div className="text-muted small">
                              Create rotation types with the <strong>+ New Rotation</strong> popup.
                            </div>
                          </div>
                        </div>
                        <div className="card border">
                          <div style={{ height: "calc(100vh - 330px)", minHeight: 520 }}>
                            <DataGrid
                              rows={crewLeavePolicies}
                              columns={[
                                { field: "policyName", headerName: "Name", minWidth: 220, flex: 1 },
                                { field: "crewType", headerName: "Category", width: 170 },
                                { field: "workPatternDays", headerName: "Work Days", width: 120 },
                                { field: "leavePatternDays", headerName: "Leave Days", width: 120 },
                                {
                                  field: "codeThresholdsJson",
                                  headerName: "Code Thresholds",
                                  minWidth: 220,
                                  flex: 1,
                                  renderCell: (params: GridRenderCellParams) => {
                                    const raw = String(params.value || "{}")
                                    try {
                                      const parsed = JSON.parse(raw) as Record<string, string>
                                      const txt = Object.entries(parsed)
                                        .map(([k, v]) => `${k}:${v}`)
                                        .join(", ")
                                      return txt || "-"
                                    } catch {
                                      return "-"
                                    }
                                  },
                                },
                                {
                                  field: "__cycles",
                                  headerName: "Cycles / year",
                                  width: 120,
                                  valueGetter: (_v, row) => {
                                    const w = Number(row.workPatternDays || "0")
                                    const l = Number(row.leavePatternDays || "0")
                                    const c = w + l
                                    return c > 0 ? String(Math.round(365 / c)) : "-"
                                  },
                                },
                                {
                                  field: "__actions",
                                  headerName: "Actions",
                                  width: 160,
                                  sortable: false,
                                  filterable: false,
                                  renderCell: (params: GridRenderCellParams) => {
                                    const row = params.row as Entry
                                    return (
                                      <div className="d-flex gap-2">
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-link p-0"
                                          onClick={() => editCrewLeavePolicy(row.id)}
                                        >
                                          Edit
                                        </button>
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-link text-danger p-0"
                                          onClick={() => deleteCrewLeavePolicy(row.id)}
                                        >
                                          Delete
                                        </button>
                                      </div>
                                    )
                                  },
                                },
                              ]}
                              disableRowSelectionOnClick
                              rowHeight={34}
                              sx={gridSx}
                            />
                          </div>
                        </div>
                      </>
                    ) : null}

                    {crewLeavePlannerTab === "attendanceCodes" ? (
                      <div className="d-flex flex-column gap-3">
                        <div className="card border">
                          <div className="card-header bg-body-tertiary py-2 small fw-semibold">
                            Attendance Codes (Crew Ops)
                          </div>
                          <div className="card-body">
                            <div className="row g-2">
                              <div className="col-12 col-md-2">
                                <label className="form-label small fw-semibold">Code</label>
                                <input
                                  className="form-control text-uppercase"
                                  value={crewOpsCodeForm.code}
                                  onChange={(e) =>
                                    setCrewOpsCodeForm((p) => ({ ...p, code: e.target.value.toUpperCase() }))
                                  }
                                />
                              </div>
                              <div className="col-12 col-md-5">
                                <label className="form-label small fw-semibold">Name</label>
                                <input
                                  className="form-control"
                                  value={crewOpsCodeForm.name}
                                  onChange={(e) => setCrewOpsCodeForm((p) => ({ ...p, name: e.target.value }))}
                                  placeholder="e.g. Medical Certificate"
                                />
                              </div>
                              <div className="col-12 col-md-3">
                                <label className="form-label small fw-semibold">Allow Leave Deduction</label>
                                <select
                                  className="form-select"
                                  value={crewOpsCodeForm.allowLeaveDeduction}
                                  onChange={(e) =>
                                    setCrewOpsCodeForm((p) => ({ ...p, allowLeaveDeduction: e.target.value }))
                                  }
                                >
                                  <option value="Yes">Yes</option>
                                  <option value="No">No</option>
                                </select>
                              </div>
                              <div className="col-12 col-md-2 d-flex align-items-end">
                                <Button onClick={saveCrewOpsCode}>Save Code</Button>
                              </div>
                            </div>
                          </div>
                        </div>

                        <div className="card border">
                          <div style={{ height: "calc(100vh - 360px)", minHeight: 500 }}>
                            <DataGrid
                              rows={crewOpsCodeConfigs.map((row) => ({
                                id: row.id,
                                code: (row.code || "").trim().toUpperCase(),
                                codeName: row.codeName || "",
                                allowLeaveDeduction: isYesLike(row.allowLeaveDeduction) ? "Yes" : "No",
                              }))}
                              columns={[
                                { field: "code", headerName: "Code", width: 120 },
                                { field: "codeName", headerName: "Name", minWidth: 220, flex: 1 },
                                {
                                  field: "allowLeaveDeduction",
                                  headerName: "Allow Leave Deduction",
                                  width: 180,
                                },
                                {
                                  field: "__actions",
                                  headerName: "Actions",
                                  width: 150,
                                  sortable: false,
                                  filterable: false,
                                  renderCell: (params: GridRenderCellParams) => {
                                    const row = params.row as Entry
                                    return (
                                      <div className="d-flex gap-2">
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-link p-0"
                                          onClick={() => editCrewOpsCode(row.id)}
                                        >
                                          Edit
                                        </button>
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-link text-danger p-0"
                                          onClick={() => deleteCrewOpsCode(row.id)}
                                        >
                                          Delete
                                        </button>
                                      </div>
                                    )
                                  },
                                },
                              ]}
                              disableRowSelectionOnClick
                              rowHeight={34}
                              sx={gridSx}
                            />
                          </div>
                        </div>
                      </div>
                    ) : null}

                    {crewLeavePlannerTab === "dutyTypes" ? (
                      <div className="d-flex flex-column gap-3">
                        <div className="card border">
                          <div className="card-header bg-body-tertiary py-2 small fw-semibold">
                            Duty Type Manager
                          </div>
                          <div className="card-body">
                            <div className="row g-2">
                              <div className="col-12 col-md-5">
                                <label className="form-label small fw-semibold">Duty Type</label>
                                <input
                                  className="form-control"
                                  value={crewDutyTypeForm.dutyType}
                                  onChange={(e) =>
                                    setCrewDutyTypeForm((p) => ({ ...p, dutyType: e.target.value }))
                                  }
                                  placeholder="e.g. Flight Duty, Ground Duty, OF"
                                />
                              </div>
                              <div className="col-12 col-md-4">
                                <label className="form-label small fw-semibold">Map To Roster Bar</label>
                                <select
                                  className="form-select"
                                  value={crewDutyTypeForm.rosterBarMap}
                                  onChange={(e) =>
                                    setCrewDutyTypeForm((p) => ({ ...p, rosterBarMap: e.target.value }))
                                  }
                                >
                                  <option value="ROSTER DAY">Roster Day</option>
                                  <option value="REST">Rest Day</option>
                                </select>
                              </div>
                              <div className="col-12 col-md-3 d-flex align-items-end">
                                <div className="d-flex gap-2">
                                  <Button onClick={saveCrewDutyType}>
                                    {editingCrewDutyTypeName ? "Update" : "Save"}
                                  </Button>
                                  {(editingCrewDutyTypeName || crewDutyTypeForm.dutyType) ? (
                                    <Button variant="outline" onClick={resetCrewDutyTypeForm}>
                                      Cancel
                                    </Button>
                                  ) : null}
                                </div>
                              </div>
                            </div>
                            <div className="small text-muted mt-2">
                              Example: Flight Duty and Ground Duty map to Roster Day. OF / Release maps to Rest Day.
                            </div>
                          </div>
                        </div>

                        <div className="card border">
                          <div style={{ height: "calc(100vh - 360px)", minHeight: 500 }}>
                            <DataGrid
                              rows={crewDutyTypeConfigs.map((row) => ({
                                id: row.id,
                                dutyType: row.dutyType || "",
                                rosterBarMap: row.rosterBarMap === "REST" ? "REST" : "ROSTER DAY",
                                source: row.isDefault === "Yes" ? "Default" : "Saved",
                              }))}
                              columns={[
                                { field: "dutyType", headerName: "Duty Type", minWidth: 220, flex: 1 },
                                { field: "rosterBarMap", headerName: "Roster Bar Map", width: 170 },
                                { field: "source", headerName: "Source", width: 120 },
                                {
                                  field: "__actions",
                                  headerName: "Actions",
                                  width: 150,
                                  sortable: false,
                                  filterable: false,
                                  renderCell: (params: GridRenderCellParams) => {
                                    const row = params.row as Entry
                                    return (
                                      <div className="d-flex gap-2">
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-link p-0"
                                          onClick={() => editCrewDutyType(row.id)}
                                        >
                                          Edit
                                        </button>
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-link text-danger p-0"
                                          onClick={() => deleteCrewDutyType(row.id)}
                                        >
                                          Delete
                                        </button>
                                      </div>
                                    )
                                  },
                                },
                              ]}
                              disableRowSelectionOnClick
                              rowHeight={34}
                              sx={gridSx}
                            />
                          </div>
                        </div>
                      </div>
                    ) : null}

                    {crewLeavePlannerTab === "dutyCodes" ? (
                      <div className="d-flex flex-column gap-3">
                        <div className="card border">
                          <div className="card-header bg-body-tertiary py-2 small fw-semibold">
                            Duty Code Manager
                          </div>
                          <div className="card-body">
                            <div className="row g-2">
                              <div className="col-6 col-md-2">
                                <label className="form-label small fw-semibold">Duty Code</label>
                                <input
                                  className="form-control text-uppercase"
                                  value={crewDutyCodeForm.code}
                                  onChange={(e) =>
                                    setCrewDutyCodeForm((p) => ({ ...p, code: e.target.value.toUpperCase() }))
                                  }
                                  placeholder="e.g. D1"
                                />
                              </div>
                              <div className="col-12 col-md-3">
                                <label className="form-label small fw-semibold">Duty Name</label>
                                <input
                                  className="form-control"
                                  value={crewDutyCodeForm.name}
                                  onChange={(e) => setCrewDutyCodeForm((p) => ({ ...p, name: e.target.value }))}
                                  placeholder="e.g. Morning Duty"
                                />
                              </div>
                              <div className="col-12 col-md-2">
                                <label className="form-label small fw-semibold">Duty Type</label>
                                <select
                                  className="form-select"
                                  value={crewDutyCodeForm.dutyType}
                                  onChange={(e) => {
                                    const dutyType = e.target.value
                                    const dutyTypeConfig = crewDutyTypeConfigs.find((row) => row.dutyType === dutyType)
                                    const isRestType =
                                      dutyTypeConfig?.rosterBarMap === "REST" ||
                                      getCrewDutyRosterBarMap(dutyType) === "REST"
                                    setCrewDutyCodeForm((p) => ({
                                      ...p,
                                      dutyType,
                                      startTime:
                                        isRestType && (!p.startTime || p.startTime === "00:00")
                                          ? "19:30"
                                          : p.startTime,
                                      endTime:
                                        isRestType && (!p.endTime || p.endTime === "00:00")
                                          ? "23:59"
                                          : p.endTime,
                                    }))
                                  }}
                                >
                                  {crewDutyTypeConfigs.map((row) => (
                                    <option key={row.id} value={row.dutyType || ""}>
                                      {row.dutyType || "-"} / {row.rosterBarMap === "REST" ? "Rest Day" : "Roster Day"}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              <div className="col-6 col-md-2">
                                <label className="form-label small fw-semibold">Default Start</label>
                                <input
                                  type="time"
                                  className="form-control"
                                  value={crewDutyCodeForm.startTime}
                                  onChange={(e) =>
                                    setCrewDutyCodeForm((p) => ({ ...p, startTime: e.target.value }))
                                  }
                                />
                              </div>
                              <div className="col-6 col-md-2">
                                <label className="form-label small fw-semibold">Default End</label>
                                <input
                                  type="time"
                                  className="form-control"
                                  value={crewDutyCodeForm.endTime}
                                  onChange={(e) =>
                                    setCrewDutyCodeForm((p) => ({ ...p, endTime: e.target.value }))
                                  }
                                />
                              </div>
                              <div className="col-6 col-md-1">
                                <label className="form-label small fw-semibold">Color</label>
                                <input
                                  type="color"
                                  className="form-control form-control-color w-100"
                                  value={crewDutyCodeForm.color}
                                  onChange={(e) =>
                                    setCrewDutyCodeForm((p) => ({ ...p, color: e.target.value }))
                                  }
                                  title="Duty code bar color"
                                />
                              </div>
                              <div className="col-12 col-md-1 d-flex align-items-end">
                                <div className="d-flex gap-2">
                                  <Button onClick={saveCrewDutyCode}>
                                    {editingCrewDutyCodeId ? "Update" : "Save"}
                                  </Button>
                                  {editingCrewDutyCodeId ? (
                                    <Button variant="outline" onClick={resetCrewDutyCodeForm}>
                                      Cancel
                                    </Button>
                                  ) : null}
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>

                        <div className="card border">
                          <div style={{ height: "calc(100vh - 360px)", minHeight: 500 }}>
                            <DataGrid
                              rows={crewDutyCodeConfigs.map((row) => ({
                                id: row.id,
                                code: (row.code || "").trim().toUpperCase(),
                                dutyName: row.dutyName || "",
                                dutyType: row.dutyType || "Flight Duty",
                                rosterBarMap:
                                  crewDutyTypeMap.get((row.dutyType || "Flight Duty").trim().toLowerCase()) ||
                                  row.rosterBarMap ||
                                  getCrewDutyRosterBarMap(row.dutyType || "Flight Duty"),
                                defaultStartTime: row.defaultStartTime || "",
                                defaultEndTime: row.defaultEndTime || "",
                                color: row.color || getDefaultDutyCodeColor(row.code || ""),
                                duration:
                                  row.defaultStartTime && row.defaultEndTime
                                    ? calculateShiftDuration(row.defaultStartTime, row.defaultEndTime)
                                    : "-",
                              }))}
                              columns={[
                                { field: "code", headerName: "Duty Code", width: 130 },
                                { field: "dutyName", headerName: "Duty Name", minWidth: 190, flex: 1 },
                                { field: "dutyType", headerName: "Duty Type", width: 140 },
                                { field: "rosterBarMap", headerName: "Roster Bar Map", width: 140 },
                                {
                                  field: "color",
                                  headerName: "Color",
                                  width: 100,
                                  renderCell: (params: GridRenderCellParams) => (
                                    <span
                                      className="d-inline-flex align-items-center gap-2"
                                      title={String(params.value || "")}
                                    >
                                      <span
                                        style={{
                                          width: 18,
                                          height: 18,
                                          borderRadius: 3,
                                          background: String(params.value || "#64748b"),
                                          border: "1px solid #cbd5e1",
                                        }}
                                      />
                                      <span className="small">{String(params.value || "")}</span>
                                    </span>
                                  ),
                                },
                                { field: "defaultStartTime", headerName: "Default Start", width: 140 },
                                { field: "defaultEndTime", headerName: "Default End", width: 140 },
                                { field: "duration", headerName: "Duration", width: 120 },
                                {
                                  field: "__actions",
                                  headerName: "Actions",
                                  width: 150,
                                  sortable: false,
                                  filterable: false,
                                  renderCell: (params: GridRenderCellParams) => {
                                    const row = params.row as Entry
                                    return (
                                      <div className="d-flex gap-2">
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-link p-0"
                                          onClick={() => editCrewDutyCode(row.id)}
                                        >
                                          Edit
                                        </button>
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-link text-danger p-0"
                                          onClick={() => deleteCrewDutyCode(row.id)}
                                        >
                                          Delete
                                        </button>
                                      </div>
                                    )
                                  },
                                },
                              ]}
                              disableRowSelectionOnClick
                              rowHeight={34}
                              sx={gridSx}
                            />
                          </div>
                        </div>
                      </div>
                    ) : null}

                    {crewLeavePlannerTab === "dailyAttendance" ? (
                      <div className="d-flex flex-column gap-3">
                        <div className="card border">
                          <div className="card-header bg-body-tertiary py-2 small fw-semibold">
                            Daily Attendance Record
                          </div>
                          <div className="card-body">
                            <div className="row g-2">
                              <div className="col-12 col-md-4">
                                <label className="form-label small fw-semibold">Crew</label>
                                <select
                                  className="form-select"
                                  value={crewLeaveMarkForm.crewCode}
                                  onChange={(e) => setCrewLeaveMarkForm((p) => ({ ...p, crewCode: e.target.value }))}
                                >
                                  <option value="">Select Crew</option>
                                  {store.crewDataBase.map((c) => (
                                    <option key={c.id} value={normalizeCrewCode(c.crewCode || "")}>
                                      {normalizeCrewCode(c.crewCode || "")} - {c.crewName || ""}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              <div className="col-6 col-md-2">
                                <label className="form-label small fw-semibold">Code</label>
                                <select
                                  className="form-select"
                                  value={crewLeaveMarkForm.markType}
                                  onChange={(e) => setCrewLeaveMarkForm((p) => ({ ...p, markType: e.target.value }))}
                                >
                                  {crewOpsCodeConfigs.map((row) => (
                                    <option key={row.id} value={(row.code || "").trim().toUpperCase()}>
                                      {(row.code || "").trim().toUpperCase()} - {row.codeName || ""}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              <div className="col-6 col-md-2">
                                <label className="form-label small fw-semibold">From Date</label>
                                <input
                                  type="date"
                                  className="form-control"
                                  value={crewLeaveMarkForm.fromDate}
                                  onChange={(e) => setCrewLeaveMarkForm((p) => ({ ...p, fromDate: e.target.value }))}
                                />
                              </div>
                              <div className="col-6 col-md-2">
                                <label className="form-label small fw-semibold">To Date</label>
                                <input
                                  type="date"
                                  className="form-control"
                                  value={crewLeaveMarkForm.toDate}
                                  onChange={(e) => setCrewLeaveMarkForm((p) => ({ ...p, toDate: e.target.value }))}
                                />
                              </div>
                              <div className="col-6 col-md-2">
                                <label className="form-label small fw-semibold">Deduction Cycle</label>
                                <select
                                  className="form-select"
                                  value={crewLeaveMarkForm.cycleNumber}
                                  onChange={(e) => setCrewLeaveMarkForm((p) => ({ ...p, cycleNumber: e.target.value }))}
                                >
                                  <option value="">Auto (by date overlap)</option>
                                  {dailyAttendanceCycleOptions.map((cycle) => (
                                    <option key={cycle} value={cycle}>
                                      Cycle {cycle}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              <div className="col-6 col-md-2 d-flex align-items-end">
                                <div className="d-flex gap-2">
                                  <Button onClick={addCrewLeaveMark}>
                                    {editingCrewLeaveMarkId ? "Update Record" : "Add Record"}
                                  </Button>
                                  {editingCrewLeaveMarkId ? (
                                    <Button
                                      variant="outline"
                                      onClick={() => {
                                        setEditingCrewLeaveMarkId(null)
                                        setCrewLeaveMarkForm({
                                          crewCode: "",
                                          markType: "MC",
                                          cycleNumber: "",
                                          fromDate: toYmd(new Date()),
                                          toDate: toYmd(new Date()),
                                          reason: "",
                                        })
                                      }}
                                    >
                                      Cancel
                                    </Button>
                                  ) : null}
                                </div>
                              </div>
                              <div className="col-12">
                                <label className="form-label small fw-semibold">Reason</label>
                                <input
                                  className="form-control"
                                  value={crewLeaveMarkForm.reason}
                                  onChange={(e) => setCrewLeaveMarkForm((p) => ({ ...p, reason: e.target.value }))}
                                />
                              </div>
                            </div>
                          </div>
                        </div>
                        <div className="card border">
                          <div style={{ height: "calc(100vh - 390px)", minHeight: 480 }}>
                            <DataGrid
                              rows={crewLeaveMarks.map((m) => ({
                                id: m.id,
                                crewCode: m.crewCode || "",
                                crewName: m.crewName || "",
                                markType: m.markType || "",
                                cycleNumber: m.cycleNumber || "",
                                fromDate: m.fromDate || "",
                                toDate: m.toDate || "",
                                noOfDays: m.noOfDays || "",
                                reason: m.reason || "",
                              }))}
                              columns={[
                                { field: "crewCode", headerName: "Crew Code", width: 120 },
                                { field: "crewName", headerName: "Crew Name", minWidth: 170, flex: 1 },
                                { field: "markType", headerName: "Code", width: 90 },
                                { field: "cycleNumber", headerName: "Cycle", width: 90 },
                                { field: "fromDate", headerName: "From", width: 120 },
                                { field: "toDate", headerName: "To", width: 120 },
                                { field: "noOfDays", headerName: "Days", width: 90 },
                                { field: "reason", headerName: "Reason", minWidth: 200, flex: 1 },
                                {
                                  field: "__actions",
                                  headerName: "Actions",
                                  width: 150,
                                  sortable: false,
                                  filterable: false,
                                  renderCell: (params: GridRenderCellParams) => {
                                    const row = params.row as Entry
                                    return (
                                      <div className="d-flex gap-2">
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-link p-0"
                                          onClick={() => editCrewLeaveMark(row.id)}
                                        >
                                          Edit
                                        </button>
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-link text-danger p-0"
                                          onClick={() => deleteCrewLeaveMark(row.id)}
                                        >
                                          Delete
                                        </button>
                                      </div>
                                    )
                                  },
                                },
                              ]}
                              disableRowSelectionOnClick
                              rowHeight={34}
                              sx={gridSx}
                            />
                          </div>
                        </div>
                      </div>
                    ) : null}

                    {crewLeavePlannerTab === "leavePatternManager" ? (
                      <div className="d-flex flex-column gap-3">
                        <div className="card border">
                          <div className="card-header bg-body-tertiary py-2 small fw-semibold">
                            Leave Pattern Manager
                          </div>
                          <div className="card-body">
                            <div className="row g-2">
                              <div className="col-12 col-md-5">
                                <label className="form-label small fw-semibold">Search (Crew Code / Staff # / Name)</label>
                                <input
                                  className="form-control"
                                  value={leavePatternManagerSearch}
                                  onChange={(e) => setLeavePatternManagerSearch(e.target.value)}
                                  placeholder="Type crew code or staff number"
                                />
                              </div>
                              <div className="col-12 col-md-7">
                                <label className="form-label small fw-semibold">Crew</label>
                                <select
                                  className="form-select"
                                  value={leavePatternManagerCrewCode}
                                  onChange={(e) => setLeavePatternManagerCrewCode(e.target.value)}
                                >
                                  {leavePatternManagerCrewMatches.length ? null : (
                                    <option value="">No matching crew</option>
                                  )}
                                  {leavePatternManagerCrewMatches.map((c) => (
                                    <option key={c.id} value={normalizeCrewCode(c.crewCode || "")}>
                                      {normalizeCrewCode(c.crewCode || "")} - {c.crewName || ""} | Staff # {c.employeeNo || "-"}
                                    </option>
                                  ))}
                                </select>
                              </div>
                            </div>
                          </div>
                        </div>
                        <div className="card border">
                          <div className="card-body p-0">
                            <div style={{ overflow: "auto" }}>
                              <table className="table table-sm mb-0 align-middle">
                                <thead>
                                  <tr>
                                    <th rowSpan={2} style={{ minWidth: 90, background: "#f8fafc" }}>Cycle</th>
                                    <th colSpan={4} className="text-center" style={{ background: "#dbeafe", color: "#1e3a8a" }}>Work Pattern</th>
                                    <th colSpan={4} className="text-center" style={{ background: "#dcfce7", color: "#166534" }}>Leave Pattern</th>
                                    <th colSpan={2} className="text-center" style={{ background: "#e5e7eb", color: "#111827" }}>Actual Leave (Split)</th>
                                  </tr>
                                  <tr>
                                    <th style={{ minWidth: 120 }}>Start</th>
                                    <th style={{ minWidth: 120 }}>End</th>
                                    <th style={{ minWidth: 100 }}>No of Days</th>
                                    <th style={{ minWidth: 90 }}>Adjust</th>
                                    <th style={{ minWidth: 120 }}>Start</th>
                                    <th style={{ minWidth: 120 }}>End</th>
                                    <th style={{ minWidth: 100 }}>No of Days</th>
                                    <th style={{ minWidth: 90 }}>REM</th>
                                    <th style={{ minWidth: 260 }}>Start</th>
                                    <th style={{ minWidth: 260 }}>End</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {leavePatternManagerRows.length ? (
                                    leavePatternManagerRows.map((r) => {
                                      const cycleKey = String(r.cycleNumber || "")
                                      const segments = leavePatternSegmentsByCycle.get(cycleKey) || []
                                      const expanded = leavePatternExpandedByCycle[cycleKey] === true
                                      const visibleSegments = expanded ? segments : segments.slice(0, 1)
                                      return (
                                      <tr key={r.id}>
                                        <td className="fw-semibold">{r.cycleLabel.replace("RT", "Cycle ")}</td>
                                        <td>{formatYmdToDdMmYy(r.wpStart)}</td>
                                        <td>{formatYmdToDdMmYy(r.wpEnd)}</td>
                                        <td>{r.wpDays}</td>
                                        <td className={Number(r.adjust || "0") > 0 ? "text-danger fw-semibold" : ""}>{r.adjust}</td>
                                        <td>{formatYmdToDdMmYy(r.lpStart)}</td>
                                        <td>{formatYmdToDdMmYy(r.lpEnd)}</td>
                                        <td>{r.lpDays}</td>
                                        <td
                                          className="fw-semibold text-primary"
                                          style={{ cursor: "pointer", textDecoration: "underline" }}
                                          onClick={() => {
                                            const crewCode = normalizeCrewCode(leavePatternManagerCrewCode || "")
                                            if (!crewCode) return
                                            const leaveBlock =
                                              leaveBlockByCrewCycle.get(`${crewCode}::${String(r.cycleNumber || "").trim()}`)
                                            setLeaveDetailModal({
                                              open: true,
                                              blockId: leaveBlock?.id || "",
                                              crewCode,
                                              crewName: leaveBlock?.crewName || "",
                                              cycleNumber: String(r.cycleNumber || "").trim(),
                                              leaveStart: r.lpStart || "",
                                              leaveEnd: r.lpEnd || "",
                                              policyName: leaveBlock?.policyName || "",
                                            })
                                            setLeaveDetailSelectedCycle(String(r.cycleNumber || "").trim())
                                            setLeaveDetailDraft({ start: r.lpStart || "", end: r.lpEnd || "" })
                                          }}
                                          title="Click to set Actual Leave Dates for this cycle"
                                        >
                                          {r.rem}
                                        </td>
                                        <td>
                                          <div className="d-flex flex-column gap-1">
                                            {segments.length > 1 ? (
                                              <button
                                                type="button"
                                                className="btn btn-sm btn-outline-secondary align-self-start py-0 px-2"
                                                onClick={() =>
                                                  setLeavePatternExpandedByCycle((prev) => ({
                                                    ...prev,
                                                    [cycleKey]: !expanded,
                                                  }))
                                                }
                                              >
                                                {expanded ? "Collapse" : `Expand (${segments.length})`}
                                              </button>
                                            ) : null}
                                            {visibleSegments.map((s) => (
                                              <div key={s.id} className="d-flex align-items-center justify-content-between border rounded px-2 py-1">
                                                <span>{formatYmdToDdMmYy(s.startDate || "")}</span>
                                                <button
                                                  type="button"
                                                  className="btn btn-sm btn-link text-danger p-0"
                                                  onClick={() => removeLeavePatternActualSegment(s.id)}
                                                >
                                                  Delete
                                                </button>
                                              </div>
                                            ))}
                                            <input
                                              type="date"
                                              className="form-control form-control-sm"
                                              value={leavePatternSegmentDraft[cycleKey]?.start || ""}
                                              onChange={(e) =>
                                                setLeavePatternSegmentDraft((prev) => ({
                                                  ...prev,
                                                  [cycleKey]: {
                                                    ...(prev[cycleKey] || { start: "", end: "" }),
                                                    start: e.target.value,
                                                  },
                                                }))
                                              }
                                            />
                                          </div>
                                        </td>
                                        <td>
                                          <div className="d-flex flex-column gap-1">
                                            {visibleSegments.map((s) => (
                                              <div key={`${s.id}_end`} className="border rounded px-2 py-1">
                                                {formatYmdToDdMmYy(s.endDate || "")}
                                              </div>
                                            ))}
                                            <div className="d-flex gap-1">
                                              <input
                                                type="date"
                                                className="form-control form-control-sm"
                                                value={leavePatternSegmentDraft[cycleKey]?.end || ""}
                                                onChange={(e) =>
                                                  setLeavePatternSegmentDraft((prev) => ({
                                                    ...prev,
                                                    [cycleKey]: {
                                                      ...(prev[cycleKey] || { start: "", end: "" }),
                                                      end: e.target.value,
                                                    },
                                                  }))
                                                }
                                              />
                                              <button
                                                type="button"
                                                className="btn btn-sm btn-primary"
                                                onClick={() => addLeavePatternActualSegment(cycleKey)}
                                              >
                                                Add
                                              </button>
                                            </div>
                                          </div>
                                        </td>
                                      </tr>
                                      )
                                    })
                                  ) : (
                                    <tr>
                                      <td colSpan={11} className="text-muted">
                                        No generated leave pattern found for selected crew.
                                      </td>
                                    </tr>
                                  )}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        </div>
                      </div>
                    ) : null}

                    {crewLeavePlannerTab === "generator" ? (
                      <div className="card border">
                        <div className="card-header bg-body-tertiary py-2 small fw-semibold">Generate Rotation</div>
                        <div className="card-body">
                          <div className="alert alert-info py-2 small mb-3">
                            Rotation Start Date is taken from <strong>Crew Data Base → Release Date</strong>. Employment Date uses
                            <strong> Joining Date</strong>.
                          </div>
                          <div className="row g-2 mb-2">
                            <div className="col-12 col-md-4">
                              <label className="form-label small fw-semibold">Crew Member</label>
                            <select
                              className="form-select"
                              value={crewLeaveAssignmentForm.crewCode}
                              onChange={(e) =>
                                setCrewLeaveAssignmentForm((p) => ({ ...p, crewCode: e.target.value, policyId: "" }))
                              }
                            >
                                <option value="">Select Crew</option>
                                {crewWithReleaseDate.map((c) => (
                                  <option key={c.id} value={normalizeCrewCode(c.crewCode || "")}>
                                    {normalizeCrewCode(c.crewCode || "")} - {c.crewName || ""} ({c.crewType || "-"})
                                  </option>
                                ))}
                              </select>
                            </div>
                            {selectedCrewGeneratedSummary ? (
                              <div className="col-12">
                                {selectedCrewGeneratedSummary.exists ? (
                                  <div className="alert alert-warning py-2 small mb-1">
                                    Rotation already generated for <strong>{selectedCrewGeneratedSummary.crewCode}</strong>.
                                    Blocks: <strong>{selectedCrewGeneratedSummary.totalBlocks}</strong>, Cycles:{" "}
                                    <strong>{selectedCrewGeneratedSummary.cycleCount}</strong>, Range:{" "}
                                    <strong>
                                      {selectedCrewGeneratedSummary.firstDate} to {selectedCrewGeneratedSummary.lastDate}
                                    </strong>
                                    . Use <strong>Replace existing blocks = Yes</strong> to overwrite.
                                  </div>
                                ) : (
                                  <div className="alert alert-success py-2 small mb-1">
                                    No generated rotation found for <strong>{selectedCrewGeneratedSummary.crewCode}</strong>.
                                  </div>
                                )}
                              </div>
                            ) : null}
                            <div className="col-12 col-md-6">
                              <label className="form-label small fw-semibold">Rotation Type</label>
                              <select
                                className="form-select"
                                value={crewLeaveAssignmentForm.policyId}
                                onChange={(e) => setCrewLeaveAssignmentForm((p) => ({ ...p, policyId: e.target.value }))}
                              >
                                <option value="">Select Policy</option>
                                {crewLeavePolicies
                                  .filter((p) => (p.activeStatus || "").trim() === "Active")
                                  .filter((p) => {
                                    if (!crewLeaveAssignmentForm.crewCode) return true
                                    const crew = store.crewDataBase.find(
                                      (c) => normalizeCrewCode(c.crewCode || "") === crewLeaveAssignmentForm.crewCode,
                                    )
                                    return (p.crewType || "").trim() === (crew?.crewType || "").trim()
                                  })
                                  .map((p) => (
                                    <option key={p.id} value={p.id}>
                                      {p.policyName || "-"} ({p.workPatternDays || "-"} / {p.leavePatternDays || "-"})
                                    </option>
                                  ))}
                              </select>
                            </div>
                            <div className="col-12 col-md-2 d-flex align-items-end">
                              <Button onClick={assignCrewLeavePolicy}>Assign</Button>
                            </div>
                          </div>

                          <div className="row g-2">
                            <div className="col-6 col-md-3">
                              <label className="form-label small fw-semibold">Years to Generate</label>
                              <input
                                type="number"
                                min={1}
                                max={10}
                                className="form-control"
                                value={crewLeaveGeneratorForm.years}
                                onChange={(e) => setCrewLeaveGeneratorForm((p) => ({ ...p, years: e.target.value }))}
                              />
                            </div>
                            <div className="col-6 col-md-4">
                              <label className="form-label small fw-semibold">Replace existing blocks</label>
                              <select
                                className="form-select"
                                value={crewLeaveGeneratorForm.replaceExisting}
                                onChange={(e) =>
                                  setCrewLeaveGeneratorForm((p) => ({ ...p, replaceExisting: e.target.value }))
                                }
                              >
                                <option value="Yes">Yes</option>
                                <option value="No">No</option>
                              </select>
                            </div>
                            <div className="col-12 col-md-5 d-flex align-items-end">
                              <Button onClick={generateCrewLeaveBlocks}>Generate</Button>
                            </div>
                          </div>

                        </div>
                      </div>
                    ) : null}

                    {isRotationModalOpen ? (
                      <div
                        className="modal show d-block"
                        tabIndex={-1}
                        style={{ background: "rgba(0,0,0,0.45)" }}
                        onClick={() => setIsRotationModalOpen(false)}
                      >
                        <div className="modal-dialog modal-dialog-centered" onClick={(e) => e.stopPropagation()}>
                          <div className="modal-content border-0 shadow" style={{ borderRadius: 14 }}>
                            <div className="modal-header border-0 pb-0">
                              <h5 className="modal-title">New Rotation Type</h5>
                              <button
                                type="button"
                                className="btn btn-sm btn-light rounded-circle"
                                onClick={() => setIsRotationModalOpen(false)}
                              >
                                <X size={14} />
                              </button>
                            </div>
                            <div className="modal-body">
                              <div className="mb-2">
                                <label className="form-label small fw-semibold">Name</label>
                                <input
                                  className="form-control"
                                  value={crewLeavePolicyForm.policyName}
                                  onChange={(e) => setCrewLeavePolicyForm((p) => ({ ...p, policyName: e.target.value }))}
                                />
                              </div>
                              <div className="mb-2">
                                <label className="form-label small fw-semibold">Category</label>
                                <select
                                  className="form-select"
                                  value={crewLeavePolicyForm.crewType}
                                  onChange={(e) => setCrewLeavePolicyForm((p) => ({ ...p, crewType: e.target.value }))}
                                >
                                  {(crewTypes.length > 0 ? crewTypes : ["Captain", "First Officer", "Cabin Crew"]).map((t) => (
                                    <option key={t} value={t}>{t}</option>
                                  ))}
                                </select>
                              </div>
                              <div className="row g-2 mb-2">
                                <div className="col-6">
                                  <label className="form-label small fw-semibold">Work Days</label>
                                  <input
                                    type="number"
                                    min={1}
                                    className="form-control"
                                    value={crewLeavePolicyForm.workPatternDays}
                                    onChange={(e) => setCrewLeavePolicyForm((p) => ({ ...p, workPatternDays: e.target.value }))}
                                  />
                                </div>
                                <div className="col-6">
                                  <label className="form-label small fw-semibold">Leave Days</label>
                                  <input
                                    type="number"
                                    min={1}
                                    className="form-control"
                                    value={crewLeavePolicyForm.leavePatternDays}
                                    onChange={(e) => setCrewLeavePolicyForm((p) => ({ ...p, leavePatternDays: e.target.value }))}
                                  />
                                </div>
                              </div>
                              <div className="row g-2 mb-2">
                                {crewOpsDeductionCodes.length === 0 ? (
                                  <div className="col-12 small text-muted">
                                    No allowed deduction codes found. Enable codes in Attendance Codes tab first.
                                  </div>
                                ) : (
                                  crewOpsDeductionCodes.map((row) => (
                                    <div className="col-6" key={row.code}>
                                      <label className="form-label small fw-semibold">
                                        {row.code} ({row.name}) Threshold / Cycle
                                      </label>
                                      <input
                                        type="number"
                                        min={0}
                                        className="form-control"
                                        value={crewLeavePolicyThresholds[row.code] || "0"}
                                        onChange={(e) =>
                                          setCrewLeavePolicyThresholds((prev) => ({
                                            ...prev,
                                            [row.code]: e.target.value,
                                          }))
                                        }
                                      />
                                    </div>
                                  ))
                                )}
                              </div>
                              <div className="small text-muted mb-2">
                                If cycle total exceeds threshold, the excess is deducted from leave days.
                              </div>
                              {rotationYearPreview ? (
                                <div className="border rounded p-2 bg-body-tertiary mb-2">
                                  <div className="small fw-semibold mb-1">Year fit preview</div>
                                  <div className="small">
                                    {rotationYearPreview.work} work + {rotationYearPreview.leave} leave = {rotationYearPreview.cycle}-day cycle
                                  </div>
                                  <div className="small text-muted">
                                    365 ÷ {rotationYearPreview.cycle} = {rotationYearPreview.cyclesPerYear.toFixed(2)} cycles/yr (nearest: {rotationYearPreview.nearest})
                                  </div>
                                  <div className="small">
                                    With {rotationYearPreview.nearest} cycles: {rotationYearPreview.covered} days covered,{" "}
                                    <span className={rotationYearPreview.diff > 0 ? "text-success" : rotationYearPreview.diff < 0 ? "text-danger" : ""}>
                                      {rotationYearPreview.diff > 0 ? "+" : ""}
                                      {rotationYearPreview.diff} day vs year
                                    </span>
                                  </div>
                                  <div className="small text-muted">
                                    Leave per year ≈ {rotationYearPreview.leavePerYear.toFixed(1)} days · leave share {rotationYearPreview.leaveShare.toFixed(1)}%
                                  </div>
                                </div>
                              ) : null}
                              <div>
                                <div className="small fw-semibold mb-1">Quick presets</div>
                                <div className="d-flex flex-wrap gap-1">
                                  {rotationQuickPresets.map((p) => {
                                    const cycle = p.work + p.leave
                                    const diff = Math.round((365 / cycle)) * cycle - 365
                                    return (
                                      <button
                                        key={`${p.work}-${p.leave}`}
                                        type="button"
                                        className="btn btn-sm btn-outline-secondary"
                                        onClick={() =>
                                          setCrewLeavePolicyForm((prev) => ({
                                            ...prev,
                                            workPatternDays: String(p.work),
                                            leavePatternDays: String(p.leave),
                                          }))
                                        }
                                      >
                                        {p.work}/{p.leave} · {Math.round(365 / cycle)}/yr · {diff > 0 ? "+" : ""}
                                        {diff}d
                                      </button>
                                    )
                                  })}
                                </div>
                              </div>
                            </div>
                            <div className="modal-footer border-0 pt-0">
                              <Button variant="outline" onClick={() => setIsRotationModalOpen(false)}>
                                Cancel
                              </Button>
                              <Button
                                onClick={() => {
                                  saveCrewLeavePolicy()
                                  setIsRotationModalOpen(false)
                                }}
                              >
                                {editingRotationId ? "Update" : "Save"}
                              </Button>
                            </div>
                          </div>
                        </div>
                      </div>
                    ) : null}

                    {crewLeavePlannerTab === "blocks" ? (
                      <div className="card border">
                        <div className="card-header bg-body-tertiary py-2 small fw-semibold">Crew Leave Pattern — Planned vs Actual</div>
                        <div className="card-body pb-0">
                        <div className="alert alert-info py-2 small">
                          This view shows only <strong>Leave Pattern</strong> blocks. Deduction and Actual days are applied here.
                        </div>
                        <div className="row g-2">
                          <div className="col-12 col-md-4">
                            <label className="form-label small fw-semibold">Crew</label>
                            <select
                              className="form-select"
                              value={crewLeaveBlocksFilter.crewCode}
                              onChange={(e) => setCrewLeaveBlocksFilter((p) => ({ ...p, crewCode: e.target.value }))}
                            >
                              <option value="">All</option>
                              {store.crewDataBase.map((c) => (
                                <option key={c.id} value={normalizeCrewCode(c.crewCode || "")}>
                                  {normalizeCrewCode(c.crewCode || "")} - {c.crewName || ""}
                                </option>
                              ))}
                            </select>
                          </div>
                          <div className="col-6 col-md-3">
                            <label className="form-label small fw-semibold">From</label>
                            <input
                              type="date"
                              className="form-control"
                              value={crewLeaveBlocksFilter.fromDate}
                              onChange={(e) => setCrewLeaveBlocksFilter((p) => ({ ...p, fromDate: e.target.value }))}
                            />
                          </div>
                          <div className="col-6 col-md-3">
                            <label className="form-label small fw-semibold">To</label>
                            <input
                              type="date"
                              className="form-control"
                              value={crewLeaveBlocksFilter.toDate}
                              onChange={(e) => setCrewLeaveBlocksFilter((p) => ({ ...p, toDate: e.target.value }))}
                            />
                          </div>
                        </div>
                      </div>
                      <div style={{ height: "calc(100vh - 330px)", minHeight: 520 }}>
                        <DataGrid
                          rows={computedCrewGeneratedBlocks
                            .filter((b) => {
                              const crewCode = normalizeCrewCode(b.crewCode || "")
                              const startDate = (b.startDate || "").trim()
                              const endDate = (b.endDate || "").trim()
                              if ((b.blockType || "").toLowerCase() !== "leave") return false
                              if (crewLeaveBlocksFilter.crewCode && crewCode !== crewLeaveBlocksFilter.crewCode) return false
                              if (crewLeaveBlocksFilter.fromDate && endDate < crewLeaveBlocksFilter.fromDate) return false
                              if (crewLeaveBlocksFilter.toDate && startDate > crewLeaveBlocksFilter.toDate) return false
                              return true
                            })
                            .map((b) => ({
                            id: b.id,
                            crewCode: b.crewCode || "",
                            crewName: b.crewName || "",
                            crewType: b.crewType || "",
                            policyName: b.policyName || "",
                            cycleNumber: b.cycleNumber || "",
                            blockType: b.blockType || "",
                            startDate: b.startDate || "",
                            endDate: b.endDate || "",
                            plannedDays: b.plannedDays || "",
                            actualDays: b.actualDays || b.plannedDays || "0",
                            deductionDays: b.deductionDays || "0",
                            delta: b.delta || "0",
                          }))}
                          columns={[
                            { field: "crewCode", headerName: "Crew Code", width: 120 },
                            { field: "crewName", headerName: "Crew Name", minWidth: 180, flex: 1 },
                            { field: "crewType", headerName: "Crew Type", width: 160 },
                            { field: "policyName", headerName: "Rotation Type", minWidth: 180, flex: 1 },
                            { field: "cycleNumber", headerName: "Cycle", width: 90 },
                            { field: "blockType", headerName: "Block", width: 100 },
                            { field: "startDate", headerName: "Start Date", width: 120 },
                            { field: "endDate", headerName: "End Date", width: 120 },
                            { field: "plannedDays", headerName: "Planned", width: 90 },
                            { field: "actualDays", headerName: "Actual", width: 90 },
                            { field: "deductionDays", headerName: "Deduction", width: 100 },
                            { field: "delta", headerName: "Δ", width: 70 },
                          ]}
                          disableRowSelectionOnClick
                          rowHeight={34}
                          sx={gridSx}
                        />
                      </div>
                    </div>
                    ) : null}

                    {crewLeavePlannerTab === "timeline" ? (
                      <div>
                        <div className="d-flex flex-wrap align-items-end justify-content-between gap-2 mb-2">
                          <div className="d-flex flex-wrap align-items-end gap-2">
                            <div style={{ width: 170 }}>
                              <label className="form-label small fw-semibold mb-1">From Date</label>
                              <input
                                type="date"
                                className="form-control form-control-sm"
                                value={crewLeaveTimelineFilter.fromDate}
                                onChange={(e) =>
                                  setCrewLeaveTimelineFilter((p) => ({ ...p, fromDate: e.target.value }))
                                }
                              />
                            </div>
                            <div style={{ width: 170 }}>
                              <label className="form-label small fw-semibold mb-1">To Date</label>
                              <input
                                type="date"
                                className="form-control form-control-sm"
                                value={crewLeaveTimelineFilter.toDate}
                                onChange={(e) =>
                                  setCrewLeaveTimelineFilter((p) => ({ ...p, toDate: e.target.value }))
                                }
                              />
                            </div>
                            <div style={{ width: 190 }}>
                              <label className="form-label small fw-semibold mb-1">Crew Type</label>
                              <select
                                className="form-select form-select-sm"
                                value={crewLeaveTimelineCrewTypeFilter}
                                onChange={(e) => setCrewLeaveTimelineCrewTypeFilter(e.target.value)}
                              >
                                <option value="">All Crew Types</option>
                                {crewTimelineCrewTypes.map((t) => (
                                  <option key={t} value={t}>
                                    {t}
                                  </option>
                                ))}
                              </select>
                            </div>
                            <button
                              type="button"
                              className="btn btn-sm btn-outline-dark"
                              onClick={() => {
                                const range = getCurrentSixMonthRange()
                                setCrewRosterGeneratorForm((p) => ({
                                  ...p,
                                  fromDate: range.fromDate,
                                  toDate: range.toDate,
                                }))
                                setCrewRosterGeneratorTab("setup")
                                setIsCrewRosterGeneratorOpen(true)
                              }}
                            >
                              Roster Generator
                            </button>
                            <button
                              type="button"
                              className="btn btn-sm btn-outline-primary"
                              onClick={() => crewRosterExcelInputRef.current?.click()}
                            >
                              Import Crew Roster Excel
                            </button>
                            <input
                              ref={crewRosterExcelInputRef}
                              type="file"
                              accept=".xlsx,.xls"
                              className="d-none"
                              onChange={async (e) => {
                                const inputEl = e.currentTarget
                                const file = e.target.files?.[0]
                                if (!file) return
                                await importCrewRosterExcel(file)
                                if (inputEl) inputEl.value = ""
                              }}
                            />
                            <button
                              type="button"
                              className={"btn btn-sm " + (showTimelineWpLpSummary ? "btn-primary" : "btn-outline-secondary")}
                              onClick={() => setShowTimelineWpLpSummary((v) => !v)}
                            >
                              WP/LP Summary {showTimelineWpLpSummary ? "On" : "Off"}
                            </button>
                            <button
                              type="button"
                              className={"btn btn-sm " + (showTimelineWorkBars ? "btn-primary" : "btn-outline-secondary")}
                              onClick={() => setShowTimelineWorkBars((v) => !v)}
                            >
                              WP Bar {showTimelineWorkBars ? "On" : "Off"}
                            </button>
                            <button
                              type="button"
                              className={"btn btn-sm " + (showTimelineLeaveBars ? "btn-primary" : "btn-outline-secondary")}
                              onClick={() => setShowTimelineLeaveBars((v) => !v)}
                            >
                              LP Bar {showTimelineLeaveBars ? "On" : "Off"}
                            </button>
                          </div>
                          <div className="d-flex align-items-center gap-2">
                            <label className="small fw-semibold mb-0">
                              Timeline Zoom ({crewLeaveTimelineDayWidth}px/day)
                            </label>
                            <input
                              type="range"
                              min={18}
                              max={144}
                              step={2}
                              className="form-range mb-0"
                              style={{ width: 150 }}
                              value={crewLeaveTimelineDayWidth}
                              onChange={(e) => setCrewLeaveTimelineDayWidth(Number(e.target.value) || 36)}
                            />
                            <button
                              type="button"
                              className="btn btn-sm btn-outline-primary"
                              onClick={() => {
                                const url = new URL(window.location.href)
                                url.searchParams.set("module", "crewLeavePlanner")
                                url.searchParams.set("crewTab", "timeline")
                                url.searchParams.set("crewView", "masterOnly")
                                window.open(url.toString(), "_blank", "noopener,noreferrer")
                              }}
                            >
                              Open In New Tab
                            </button>
                          </div>
                        </div>
                        <div className="d-flex flex-wrap align-items-center gap-2 mb-2 p-2 border rounded bg-body-tertiary">
                          <span className="small fw-semibold text-muted me-1">Drag Duty / Rest Code:</span>
                          {crewDutyCodeConfigs.length ? (
                            crewDutyCodeConfigs.map((row) => {
                              const code = (row.code || "").trim().toUpperCase()
                              if (!code) return null
                              const dutyType = (row.dutyType || "").trim()
                              const map =
                                crewDutyTypeMap.get(dutyType.toLowerCase()) ||
                                row.rosterBarMap ||
                                getCrewDutyRosterBarMap(dutyType)
                              const color =
                                row.color ||
                                (map === "REST" ? ROSTER_OFF_COLOR.bg : getDefaultDutyCodeColor(code))
                              return (
                                <span
                                  key={`drag-duty-${row.id}`}
                                  draggable
                                  onDragStart={(e) => {
                                    e.dataTransfer.setData("application/x-crew-duty-code", code)
                                    e.dataTransfer.effectAllowed = "copy"
                                  }}
                                  className="badge rounded-pill"
                                  title={`${code} - ${row.dutyName || dutyType || ""} (${map === "REST" ? "Rest" : "Roster Day"})`}
                                  style={{
                                    cursor: "grab",
                                    background: color,
                                    color: "#ffffff",
                                    border: "1px solid rgba(15,23,42,0.25)",
                                    padding: "7px 10px",
                                    fontSize: 12,
                                  }}
                                >
                                  {code} {map === "REST" ? "Rest" : "Duty"}
                                </span>
                              )
                            })
                          ) : (
                            <span className="small text-muted">Create duty codes in Duty Code Manager first.</span>
                          )}
                        </div>
                          {!crewLeaveTimelineView.crewRows.length || !crewLeaveTimelineView.dates.length ? (
                            <div className="small text-muted">
                              No generated blocks for selected date range. Generate rotation first, then open timeline.
                            </div>
                          ) : (
                            <div
                              ref={crewLeaveTimelineScrollRef}
                              style={{
                                overflow: "auto",
                                maxHeight: "72vh",
                                border: "1px solid #e5e7eb",
                                borderRadius: 8,
                              }}
                            >
                              <div style={{ minWidth: 980 }}>
                                {(() => {
                                  const todayYmd = toYmd(new Date())
                                  const monthSegments: Array<{ key: string; label: string; count: number }> = []
                                  crewLeaveTimelineView.dates.forEach((date) => {
                                    const key = date.slice(0, 7)
                                    const d = new Date(`${date}T00:00:00`)
                                    const label = d.toLocaleDateString("en-US", { month: "short", year: "numeric" })
                                    const last = monthSegments[monthSegments.length - 1]
                                    if (last && last.key === key) {
                                      last.count += 1
                                    } else {
                                      monthSegments.push({ key, label, count: 1 })
                                    }
                                  })
                                  return (
                                    <>
                                      <div
                                        className="border-bottom"
                                        style={{
                                          display: "grid",
                                          gridTemplateColumns: "220px auto",
                                          width: 220 + crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                          minWidth: 220 + crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                          position: "sticky",
                                          top: 0,
                                          zIndex: 30,
                                          height: 32,
                                          background: "var(--bs-body-bg)",
                                        }}
                                      >
                                        <div
                                          style={{
                                            width: 220,
                                            minWidth: 220,
                                            position: "sticky",
                                            left: 0,
                                            zIndex: 30,
                                            background: "#ffffff",
                                            padding: "8px 10px",
                                            borderRight: "1px solid #e5e7eb",
                                            boxShadow: "2px 0 0 #e5e7eb",
                                            fontWeight: 700,
                                          }}
                                        >
                                          Crew Code
                                        </div>
                                        <div className="d-flex">
                                          {monthSegments.map((seg) => (
                                            <div
                                              key={seg.key}
                                              style={{
                                                width: seg.count * crewLeaveTimelineDayWidth,
                                                minWidth: seg.count * crewLeaveTimelineDayWidth,
                                                height: 32,
                                                borderRight: "1px solid #e2e8f0",
                                                display: "flex",
                                                justifyContent: "center",
                                                alignItems: "center",
                                                background: "#f8fafc",
                                                color: "#334155",
                                                fontSize: 10,
                                                fontWeight: 700,
                                                letterSpacing: "0.03em",
                                                padding: "2px 0",
                                              }}
                                            >
                                              <span>{seg.label}</span>
                                            </div>
                                          ))}
                                        </div>
                                      </div>
                                      <div
                                        className="border-bottom"
                                        style={{
                                          display: "grid",
                                          gridTemplateColumns: "220px auto",
                                          width: 220 + crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                          minWidth: 220 + crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                          position: "sticky",
                                          top: 32,
                                          zIndex: 29,
                                          height: 34,
                                          background: "var(--bs-body-bg)",
                                        }}
                                      >
                                        <div
                                          style={{
                                            width: 220,
                                            minWidth: 220,
                                            position: "sticky",
                                            left: 0,
                                            zIndex: 30,
                                            background: "#ffffff",
                                            padding: "8px 10px",
                                            borderRight: "1px solid #e5e7eb",
                                            boxShadow: "2px 0 0 #e5e7eb",
                                            fontWeight: 700,
                                          }}
                                        >
                                          Date / Day
                                        </div>
                                        <div className="d-flex">
                                          {crewLeaveTimelineView.dates.map((date, idx) => {
                                            const isToday = date === todayYmd
                                            return (
                                              <div
                                                key={date}
                                                style={{
                                                  width: crewLeaveTimelineDayWidth,
                                                  minWidth: crewLeaveTimelineDayWidth,
                                                  height: 34,
                                                  fontSize: 10,
                                                  textAlign: "center",
                                                  padding: "4px 0",
                                                  borderRight: crewTimelineMonthBoundaryIndexes.has(idx)
                                                    ? "2px solid #94a3b8"
                                                    : "1px solid #eef2f7",
                                                  color: isToday ? "#0f172a" : "#64748b",
                                                  background: isToday ? "#fff7ed" : "transparent",
                                                  fontWeight: isToday ? 700 : 500,
                                                }}
                                              >
                                                <div>{date.slice(8, 10)}</div>
                                                {crewLeaveTimelineDayWidth > 24 ? (
                                                  <div>{formatYmdToWeekdayShort(date)}</div>
                                                ) : null}
                                              </div>
                                            )
                                          })}
                                        </div>
                                      </div>
                                    </>
                                  )
                                })()}

                                {showTimelineWpLpSummary
                                  ? crewTypeDailySummaryMap.map(([crewType, dayMap], summaryIdx) => (
                                  <div key={`summary-${crewType}`}>
                                    <div
                                      className="border-bottom"
                                      style={{
                                        display: "grid",
                                        gridTemplateColumns: "220px auto",
                                        width: 220 + crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                        minWidth: 220 + crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                        minHeight: 30,
                                        background: "#f8fafc",
                                        position: "sticky",
                                        top: 66 + summaryIdx * 60,
                                        zIndex: 28,
                                      }}
                                    >
                                      <div
                                        style={{
                                          width: 220,
                                          minWidth: 220,
                                          position: "sticky",
                                          left: 0,
                                          zIndex: 20,
                                          background: "#f8fafc",
                                          padding: "5px 10px",
                                          borderRight: "1px solid #e5e7eb",
                                          boxShadow: "2px 0 0 #e5e7eb",
                                          fontWeight: 700,
                                          fontSize: 12,
                                        }}
                                      >
                                        {crewType} - WP
                                      </div>
                                      <div className="d-flex">
                                        {crewLeaveTimelineView.dates.map((date, idx) => {
                                          const v = dayMap.get(date) || { wp: 0, lp: 0 }
                                          return (
                                            <div
                                              key={`${crewType}-wp-${date}`}
                                              title={`${crewType} | ${date} | WP ${v.wp}`}
                                              style={{
                                                width: crewLeaveTimelineDayWidth,
                                                minWidth: crewLeaveTimelineDayWidth,
                                                borderRight: crewTimelineMonthBoundaryIndexes.has(idx)
                                                  ? "2px solid #94a3b8"
                                                  : "1px solid #eef2f7",
                                                fontSize: crewLeaveTimelineDayWidth > 24 ? 10 : 9,
                                                textAlign: "center",
                                                lineHeight: 1.2,
                                                paddingTop: 5,
                                                color: "#1e3a8a",
                                                fontWeight: 700,
                                              }}
                                            >
                                              {v.wp}
                                            </div>
                                          )
                                        })}
                                      </div>
                                    </div>
                                    <div
                                      className="border-bottom"
                                      style={{
                                        display: "grid",
                                        gridTemplateColumns: "220px auto",
                                        width: 220 + crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                        minWidth: 220 + crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                        minHeight: 30,
                                        background: "#f0fdf4",
                                        position: "sticky",
                                        top: 96 + summaryIdx * 60,
                                        zIndex: 28,
                                      }}
                                    >
                                      <div
                                        style={{
                                          width: 220,
                                          minWidth: 220,
                                          position: "sticky",
                                          left: 0,
                                          zIndex: 20,
                                          background: "#f0fdf4",
                                          padding: "5px 10px",
                                          borderRight: "1px solid #e5e7eb",
                                          boxShadow: "2px 0 0 #e5e7eb",
                                          fontWeight: 700,
                                          fontSize: 12,
                                        }}
                                      >
                                        {crewType} - LP
                                      </div>
                                      <div className="d-flex">
                                        {crewLeaveTimelineView.dates.map((date, idx) => {
                                          const v = dayMap.get(date) || { wp: 0, lp: 0 }
                                          const lpOverLimit = v.lp >= 5
                                          return (
                                            <div
                                              key={`${crewType}-lp-${date}`}
                                              title={`${crewType} | ${date} | LP ${v.lp}`}
                                              style={{
                                                width: crewLeaveTimelineDayWidth,
                                                minWidth: crewLeaveTimelineDayWidth,
                                                borderRight: crewTimelineMonthBoundaryIndexes.has(idx)
                                                  ? "2px solid #94a3b8"
                                                  : "1px solid #eef2f7",
                                                fontSize: crewLeaveTimelineDayWidth > 24 ? 10 : 9,
                                                textAlign: "center",
                                                lineHeight: 1.2,
                                                paddingTop: 5,
                                                color: lpOverLimit ? "#ffffff" : "#166534",
                                                background: lpOverLimit ? "#dc2626" : "transparent",
                                                fontWeight: 700,
                                              }}
                                            >
                                              {v.lp}
                                            </div>
                                          )
                                        })}
                                      </div>
                                    </div>
                                  </div>
                                ))
                                  : null}

                                {crewLeaveTimelineView.crewRows.map((row) => {
                                  const showAnyPatternBars = showTimelineWorkBars || showTimelineLeaveBars
                                  const timelineRowHeight = showAnyPatternBars ? 42 : 30
                                  const rosterBarTop = showAnyPatternBars ? 22 : 3
                                  const rosterBarHeight = showAnyPatternBars ? 16 : timelineRowHeight - 6
                                  return (
                                  <div
                                    key={row.crewCode}
                                    className="border-bottom"
                                    style={{
                                      display: "grid",
                                      gridTemplateColumns: "220px auto",
                                      width: 220 + crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                      minWidth: 220 + crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                      minHeight: timelineRowHeight,
                                    }}
                                  >
                                    <div
                                      style={{
                                        width: 220,
                                        minWidth: 220,
                                        position: "sticky",
                                        left: 0,
                                        zIndex: 20,
                                        background: "#ffffff",
                                        padding: "3px 8px",
                                        borderRight: "1px solid #e5e7eb",
                                        boxShadow: "2px 0 0 #e5e7eb",
                                        fontWeight: 600,
                                        fontSize: 12,
                                        display: "flex",
                                        alignItems: "center",
                                        gap: 6,
                                      }}
                                    >
                                      <button
                                        type="button"
                                        className="btn btn-link p-0 text-start fw-semibold"
                                        style={{ fontSize: 12, textDecoration: "underline", flexShrink: 0 }}
                                        onClick={() =>
                                          openRosterGeneratorForCrew(
                                            row.crewCode,
                                            (crewLeaveTimelineFilter.fromDate || "").trim(),
                                          )
                                        }
                                      >
                                        {row.crewCode}
                                      </button>
                                      <span
                                        className="text-muted text-truncate"
                                        style={{ fontSize: 11, lineHeight: 1.1, minWidth: 0 }}
                                      >
                                        {row.crewName || "-"}
                                      </span>
                                    </div>
                                    <div
                                      data-crew-timeline-row="true"
                                      data-crew-code={row.crewCode}
                                      onDragOver={(e) => {
                                        if (e.dataTransfer.types.includes("application/x-crew-duty-code")) {
                                          e.preventDefault()
                                          e.dataTransfer.dropEffect = "copy"
                                        }
                                      }}
                                      onDrop={(e) => {
                                        const code = e.dataTransfer.getData("application/x-crew-duty-code")
                                        if (!code) return
                                        e.preventDefault()
                                        e.stopPropagation()
                                        const rect = (e.currentTarget as HTMLDivElement).getBoundingClientRect()
                                        const x = Math.max(0, e.clientX - rect.left)
                                        const idx = Math.floor(x / crewLeaveTimelineDayWidth)
                                        const selectedDate =
                                          crewLeaveTimelineView.dates[
                                            Math.min(
                                              Math.max(idx, 0),
                                              Math.max(crewLeaveTimelineView.dates.length - 1, 0),
                                            )
                                          ] || (crewLeaveTimelineFilter.fromDate || "").trim()
                                        openCrewDutyDropModal({
                                          code,
                                          crewCode: row.crewCode,
                                          crewName: row.crewName || "",
                                          date: selectedDate,
                                        })
                                      }}
                                      onClick={(e) => {
                                        const rect = (e.currentTarget as HTMLDivElement).getBoundingClientRect()
                                        const x = Math.max(0, e.clientX - rect.left)
                                        const idx = Math.floor(x / crewLeaveTimelineDayWidth)
                                        const selectedDate =
                                          crewLeaveTimelineView.dates[
                                            Math.min(
                                              Math.max(idx, 0),
                                              Math.max(crewLeaveTimelineView.dates.length - 1, 0),
                                            )
                                          ] || (crewLeaveTimelineFilter.fromDate || "").trim()
                                        openRosterGeneratorForCrew(row.crewCode, selectedDate)
                                      }}
                                      style={{
                                        position: "relative",
                                        width: crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                        minWidth: crewLeaveTimelineView.totalDays * crewLeaveTimelineDayWidth,
                                        backgroundImage: "linear-gradient(to right, rgba(148, 163, 184, 0.18) 1px, transparent 1px)",
                                        backgroundSize: `${crewLeaveTimelineDayWidth}px 100%`,
                                        backgroundRepeat: "repeat",
                                      }}
                                    >
                                      {Array.from(crewTimelineMonthBoundaryIndexes).map((idx) => (
                                        <div
                                          key={`month-boundary-${row.crewCode}-${idx}`}
                                          style={{
                                            position: "absolute",
                                            left: (idx + 1) * crewLeaveTimelineDayWidth - 1,
                                            top: 0,
                                            bottom: 0,
                                            width: 2,
                                            background: "#94a3b8",
                                            opacity: 0.7,
                                            pointerEvents: "none",
                                            zIndex: 2,
                                          }}
                                        />
                                      ))}
                                      {(() => {
                                        const today = toYmd(new Date())
                                        const viewStart = (crewLeaveTimelineFilter.fromDate || "").trim()
                                        const viewEnd = (crewLeaveTimelineFilter.toDate || "").trim()
                                        if (!viewStart || !viewEnd) return null
                                        if (today < viewStart || today > viewEnd) return null
                                        const todayOffset = overlapDaysInclusive(
                                          viewStart,
                                          addDaysToYmd(today, -1),
                                          viewStart,
                                          viewEnd,
                                        )
                                        return (
                                          <div
                                            title={`Today: ${today}`}
                                            style={{
                                              position: "absolute",
                                              left: todayOffset * crewLeaveTimelineDayWidth,
                                              top: 0,
                                              bottom: 0,
                                              width: crewLeaveTimelineDayWidth,
                                              background: "rgba(239, 68, 68, 0.22)",
                                              borderLeft: "2px solid #dc2626",
                                              borderRight: "2px solid #dc2626",
                                              zIndex: 6,
                                              pointerEvents: "none",
                                            }}
                                          />
                                        )
                                      })()}
                                      {row.blocks.map((block) => {
                                        const start = (block.startDate || "").trim()
                                        const end = (block.endDate || "").trim()
                                        const viewStart = (crewLeaveTimelineFilter.fromDate || "").trim()
                                        const viewEnd = (crewLeaveTimelineFilter.toDate || "").trim()
                                        const overlapStart = start > viewStart ? start : viewStart
                                        const overlapEnd = end < viewEnd ? end : viewEnd
                                        const leftDays = overlapDaysInclusive(viewStart, addDaysToYmd(overlapStart, -1), viewStart, viewEnd)
                                        const spanDays = overlapDaysInclusive(overlapStart, overlapEnd, viewStart, viewEnd)
                                        const isLeave = (block.blockType || "").toLowerCase() === "leave"
                                        if (isLeave && !showTimelineLeaveBars) return null
                                        if (!isLeave && !showTimelineWorkBars) return null
                                        const deductionDays = isLeave
                                          ? deductionByGeneratedBlockId.get((block.id || "").trim()) || 0
                                          : 0
                                        const sourceBlockId = (block.id || "").trim()
                                        const cycleKey = `${normalizeCrewCode(block.crewCode || "")}::${(block.cycleNumber || "").trim()}`
                                        const hasAssignedActualLeave =
                                          isLeave &&
                                          row.rosterBars.some((rb) => {
                                            if ((rb.rosterType || "").trim().toUpperCase() !== "ACTUAL LEAVE") return false
                                            if ((rb.sourceBlockId || "").trim() && (rb.sourceBlockId || "").trim() === sourceBlockId) {
                                              return true
                                            }
                                            const rbCycleKey = `${normalizeCrewCode(rb.crewCode || "")}::${(rb.cycleNumber || "").trim()}`
                                            return rbCycleKey === cycleKey
                                          })
                                        const rotationKey = (block.policyName || "").trim().toUpperCase() || "DEFAULT"
                                        const palette = [
                                          { workBg: "#dbeafe", workBorder: "#60a5fa", workText: "#1e3a8a", leaveBg: "#dcfce7", leaveBorder: "#4ade80", leaveText: "#166534" },
                                          { workBg: "#fef3c7", workBorder: "#f59e0b", workText: "#92400e", leaveBg: "#ffedd5", leaveBorder: "#fb923c", leaveText: "#9a3412" },
                                          { workBg: "#e0e7ff", workBorder: "#818cf8", workText: "#3730a3", leaveBg: "#f3e8ff", leaveBorder: "#c084fc", leaveText: "#6b21a8" },
                                          { workBg: "#ccfbf1", workBorder: "#14b8a6", workText: "#115e59", leaveBg: "#ecfccb", leaveBorder: "#84cc16", leaveText: "#3f6212" },
                                          { workBg: "#fee2e2", workBorder: "#f87171", workText: "#991b1b", leaveBg: "#ffe4e6", leaveBorder: "#fb7185", leaveText: "#9f1239" },
                                        ]
                                        const seed = rotationKey.split("").reduce((acc, ch) => acc + ch.charCodeAt(0), 0)
                                        const tone = palette[seed % palette.length]
                                        if (spanDays <= 0) return null
                                        return (
                                          <div
                                            key={block.id}
                                            title={`${(block.blockType || "").toUpperCase()} | ${start} to ${end}`}
                                            onClick={(e) => {
                                              e.stopPropagation()
                                              if (Date.now() - lastLeaveBarDragAtRef.current < 220) return
                                              if (leaveBarDragRef.current?.moved) return
                                              if (!isLeave) return
                                              setLeaveDetailModal({
                                                open: true,
                                                blockId: (block.id || "").trim(),
                                                crewCode: normalizeCrewCode(block.crewCode || ""),
                                                crewName: block.crewName || "",
                                                cycleNumber: (block.cycleNumber || "").trim(),
                                                leaveStart: start,
                                                leaveEnd: end,
                                                policyName: block.policyName || "",
                                              })
                                              setLeaveDetailSelectedCycle((block.cycleNumber || "").trim())
                                              setLeaveDetailDraft({ start, end })
                                            }}
                                            onMouseDown={(e) => {
                                              if (!isLeave) return
                                              e.preventDefault()
                                              e.stopPropagation()
                                              const drag = {
                                                blockId: (block.id || "").trim(),
                                                startX: e.clientX,
                                                startY: e.clientY,
                                                moved: false,
                                              }
                                              leaveBarDragRef.current = drag
                                              setDraggingLeaveBlockId((block.id || "").trim())
                                              const onMove = (ev: MouseEvent) => {
                                                if (!leaveBarDragRef.current) return
                                                const dx = ev.clientX - leaveBarDragRef.current.startX
                                                const dy = ev.clientY - leaveBarDragRef.current.startY
                                                if (Math.abs(dx) > 4 || Math.abs(dy) > 4) leaveBarDragRef.current.moved = true
                                              }
                                              const onUp = (ev: MouseEvent) => {
                                                const ref = leaveBarDragRef.current
                                                leaveBarDragRef.current = null
                                                setDraggingLeaveBlockId("")
                                                window.removeEventListener("mousemove", onMove)
                                                window.removeEventListener("mouseup", onUp)
                                                if (!ref) return
                                                const dx = ev.clientX - ref.startX
                                                const dy = ev.clientY - ref.startY
                                                const dropTarget = document
                                                  .elementFromPoint(ev.clientX, ev.clientY)
                                                  ?.closest("[data-crew-timeline-row]") as HTMLElement | null
                                                if (ref.moved && dropTarget && Math.abs(dy) > 8) {
                                                  const dropCrewCode = normalizeCrewCode(dropTarget.dataset.crewCode || "")
                                                  const sourceCrewCode = normalizeCrewCode(block.crewCode || "")
                                                  if (dropCrewCode && dropCrewCode !== sourceCrewCode) {
                                                    window.alert("Actual Leave can only be assigned to the same crew row.")
                                                    return
                                                  }
                                                  const rect = dropTarget.getBoundingClientRect()
                                                  const x = Math.max(0, ev.clientX - rect.left)
                                                  const idx = Math.floor(x / crewLeaveTimelineDayWidth)
                                                  const dropDate =
                                                    crewLeaveTimelineView.dates[
                                                      Math.min(
                                                        Math.max(idx, 0),
                                                        Math.max(crewLeaveTimelineView.dates.length - 1, 0),
                                                      )
                                                    ] || start
                                                  const leaveDays = getDateRangeInclusive(start, end).length
                                                  const actualEndDate = addDaysToYmd(dropDate, Math.max(leaveDays - 1, 0))
                                                  openActualLeaveDropModal({
                                                    block,
                                                    startDate: dropDate,
                                                    endDate: actualEndDate,
                                                  })
                                                  return
                                                }
                                                let deltaDays = Math.round(dx / crewLeaveTimelineDayWidth)
                                                if (ref.moved && deltaDays === 0) deltaDays = dx >= 0 ? 1 : -1
                                                if (!ref.moved || deltaDays === 0) return
                                                setPendingLeaveBarAdjustment({
                                                  blockId: ref.blockId,
                                                  crewCode: normalizeCrewCode(block.crewCode || ""),
                                                  crewName: block.crewName || "",
                                                  cycleNumber: (block.cycleNumber || "").trim(),
                                                  policyName: block.policyName || "",
                                                  leaveStart: addDaysToYmd(start, deltaDays),
                                                  leaveEnd: addDaysToYmd(end, deltaDays),
                                                  deltaDays,
                                                })
                                                setLeaveDetailModal({
                                                  open: true,
                                                  blockId: (block.id || "").trim(),
                                                  crewCode: normalizeCrewCode(block.crewCode || ""),
                                                  crewName: block.crewName || "",
                                                  cycleNumber: (block.cycleNumber || "").trim(),
                                                  leaveStart: addDaysToYmd(start, deltaDays),
                                                  leaveEnd: addDaysToYmd(end, deltaDays),
                                                  policyName: block.policyName || "",
                                                })
                                                setLeaveDetailSelectedCycle((block.cycleNumber || "").trim())
                                                setLeaveDetailDraft({
                                                  start: addDaysToYmd(start, deltaDays),
                                                  end: addDaysToYmd(end, deltaDays),
                                                })
                                              }
                                              window.addEventListener("mousemove", onMove)
                                              window.addEventListener("mouseup", onUp)
                                            }}
                                            style={{
                                              position: "absolute",
                                              left: leftDays * crewLeaveTimelineDayWidth + 1,
                                              top: 2,
                                              height: 16,
                                              width: Math.max(spanDays * crewLeaveTimelineDayWidth - 2, 10),
                                              borderRadius: 0,
                                              background: isLeave && hasAssignedActualLeave ? "#fef08a" : isLeave ? "#facc15" : tone.workBg,
                                              border: `1px solid ${
                                                isLeave && hasAssignedActualLeave ? "#a3a3a3" : isLeave ? "#ca8a04" : tone.workBorder
                                              }`,
                                              color: isLeave ? "#111827" : tone.workText,
                                              fontSize: 10,
                                              fontWeight: 700,
                                              padding: "0 5px",
                                              whiteSpace: "nowrap",
                                              overflow: "hidden",
                                              textOverflow: "ellipsis",
                                              cursor: isLeave
                                                ? draggingLeaveBlockId === (block.id || "").trim()
                                                  ? "grabbing"
                                                  : "grab"
                                                : "default",
                                            }}
                                          >
                                            {isLeave && hasAssignedActualLeave
                                              ? `LP ${start} → ${end} (${spanDays}d) | Leave Assigned`
                                              : `${(block.blockType || "").toUpperCase()} ${start} → ${end} (${spanDays}d)`}
                                            {isLeave && deductionDays > 0 ? (
                                              <span
                                                style={{
                                                  position: "absolute",
                                                  right: 2,
                                                  top: 2,
                                                  background: "#dc2626",
                                                  color: "#ffffff",
                                                  fontWeight: 800,
                                                  fontSize: 10,
                                                  lineHeight: 1,
                                                  padding: "2px 4px",
                                                  borderRadius: 2,
                                                  border: "1px solid #7f1d1d",
                                                }}
                                              >
                                                D-{deductionDays}
                                              </span>
                                            ) : null}
                                          </div>
                                        )
                                      })}
                                      {row.rosterBars.map((rb) => {
                                        const start = (rb.startDate || "").trim()
                                        const end = (rb.endDate || "").trim()
                                        const startHm = (rb.startTime || "00:00").trim()
                                        const endHm = (rb.endTime || "23:59").trim()
                                        const viewStart = (crewLeaveTimelineFilter.fromDate || "").trim()
                                        const viewEnd = (crewLeaveTimelineFilter.toDate || "").trim()
                                        const overlapStart = start > viewStart ? start : viewStart
                                        const overlapEnd = end < viewEnd ? end : viewEnd
                                        const leftDays = overlapDaysInclusive(
                                          viewStart,
                                          addDaysToYmd(overlapStart, -1),
                                          viewStart,
                                          viewEnd,
                                        )
                                        const spanDays = overlapDaysInclusive(overlapStart, overlapEnd, viewStart, viewEnd)
                                        if (spanDays <= 0) return null
                                        const viewStartDt = parseLocalDateTime(viewStart, "00:00")
                                        const barStartDt = parseLocalDateTime(start, startHm || "00:00")
                                        const barEndDt = parseLocalDateTime(end, endHm || "23:59")
                                        const msPerDay = 24 * 60 * 60 * 1000
                                        const leftFloatDays = Math.max(
                                          0,
                                          (barStartDt.getTime() - viewStartDt.getTime()) / msPerDay,
                                        )
                                        const spanFloatDays = Math.max(
                                          0.2,
                                          (barEndDt.getTime() - barStartDt.getTime()) / msPerDay,
                                        )
                                        const spanHours = Math.max(
                                          0.1,
                                          (barEndDt.getTime() - barStartDt.getTime()) / (60 * 60 * 1000),
                                        )
                                        const rType = (rb.rosterType || "").trim().toUpperCase()
                                        const isRest = rType === "REST"
                                        const isRosterDay = rType === "ROSTER DAY"
                                        const isDutyCodeBar = rType === "DUTY CODE"
                                        const isActualLeave = rType === "ACTUAL LEAVE"
                                        const isPlannedLeave = rType === "PLANNED LEAVE"
                                        const overlappingActualLeaveBars = isActualLeave
                                          ? []
                                          : row.rosterBars
                                              .filter((other) => (other.rosterType || "").trim().toUpperCase() === "ACTUAL LEAVE")
                                              .map((other) => {
                                                const otherStart = (other.startDate || "").trim()
                                                const otherEnd = (other.endDate || otherStart).trim()
                                                return {
                                                  startDt: parseLocalDateTime(otherStart, "00:00"),
                                                  endDt: parseLocalDateTime(otherEnd, "23:59"),
                                                }
                                              })
                                              .filter((other) => barStartDt < other.endDt && barEndDt > other.startDt)
                                              .sort((a, b) => a.startDt.getTime() - b.startDt.getTime())
                                        let renderStartDt = barStartDt
                                        let renderEndDt = barEndDt
                                        if (!isActualLeave && overlappingActualLeaveBars.length > 0) {
                                          const firstAl = overlappingActualLeaveBars[0]
                                          if (barStartDt < firstAl.startDt) {
                                            renderEndDt = firstAl.startDt
                                          } else {
                                            const lastAl = overlappingActualLeaveBars[overlappingActualLeaveBars.length - 1]
                                            renderStartDt = lastAl.endDt
                                          }
                                        }
                                        if (isRosterDay && overlappingActualLeaveBars.length === 0) {
                                          const nextActualLeave = row.rosterBars
                                            .filter((other) => (other.rosterType || "").trim().toUpperCase() === "ACTUAL LEAVE")
                                            .map((other) => {
                                              const otherStart = (other.startDate || "").trim()
                                              const otherEnd = (other.endDate || otherStart).trim()
                                              return {
                                                startDt: parseLocalDateTime(otherStart, "00:00"),
                                                endDt: parseLocalDateTime(otherEnd, "23:59"),
                                              }
                                            })
                                            .filter((other) => other.startDt > barEndDt)
                                            .sort((a, b) => a.startDt.getTime() - b.startDt.getTime())[0]
                                          if (nextActualLeave) {
                                            const hasRestBeforeLeave = row.rosterBars.some((other) => {
                                              if ((other.rosterType || "").trim().toUpperCase() !== "REST") return false
                                              const otherStart = (other.startDate || "").trim()
                                              const otherEnd = (other.endDate || otherStart).trim()
                                              const otherStartHm = (other.startTime || "00:00").trim()
                                              const otherEndHm = (other.endTime || "23:59").trim()
                                              const restStartDt = parseLocalDateTime(otherStart, otherStartHm || "00:00")
                                              const restEndDt = parseLocalDateTime(otherEnd, otherEndHm || "23:59")
                                              return restStartDt < nextActualLeave.startDt && restEndDt > barEndDt
                                            })
                                            if (!hasRestBeforeLeave) {
                                              const viewEndDt = parseLocalDateTime(viewEnd, "23:59")
                                              renderEndDt = nextActualLeave.startDt < viewEndDt ? nextActualLeave.startDt : viewEndDt
                                            }
                                          }
                                        }
                                        if (!isActualLeave && renderEndDt <= renderStartDt) return null
                                        const renderLeftFloatDays = Math.max(
                                          0,
                                          (renderStartDt.getTime() - viewStartDt.getTime()) / msPerDay,
                                        )
                                        const renderSpanFloatDays = Math.max(
                                          0.08,
                                          (renderEndDt.getTime() - renderStartDt.getTime()) / msPerDay,
                                        )
                                        const hasActualLeaveOverlap = overlappingActualLeaveBars.length > 0
                                        const labelType = isActualLeave
                                          ? "ACTUAL LEAVE"
                                          : isPlannedLeave
                                            ? "LEAVE PATTERN"
                                            : isDutyCodeBar
                                              ? "DUTY CODE"
                                              : isRest
                                                ? "REST"
                                                : "ROSTER DAY"
                                        const dutyCode = (rb.dutyCode || "").trim().toUpperCase()
                                        const dutyColor = dutyCode
                                          ? crewDutyCodeColorMap.get(dutyCode) || rb.dutyColor || getDefaultDutyCodeColor(dutyCode)
                                          : ""
                                        const hasDutyCodeOverlay =
                                          !isDutyCodeBar &&
                                          !isRest &&
                                          row.rosterBars.some((other) => {
                                            if ((other.rosterType || "").trim().toUpperCase() !== "DUTY CODE") return false
                                            const otherStart = (other.startDate || "").trim()
                                            const otherEnd = (other.endDate || otherStart).trim()
                                            return overlapDaysInclusive(start, end, otherStart, otherEnd) > 0
                                          })
                                        const barLabel = isActualLeave
                                          ? "AL"
                                          : isPlannedLeave
                                            ? "LP"
                                            : dutyCode || (isRest ? "OF" : "")
                                        const barTitle = `${barLabel ? `${barLabel} | ` : ""}${labelType} | ${start} ${startHm} to ${end} ${endHm}`
                                        const getSelectedRosterDateTime = (clientX: number, rect: DOMRect) => {
                                          const ratio = Math.min(Math.max((clientX - rect.left) / Math.max(rect.width, 1), 0), 1)
                                          return new Date(
                                            renderStartDt.getTime() +
                                              (renderEndDt.getTime() - renderStartDt.getTime()) * ratio,
                                          )
                                        }
                                        return (
                                          <div
                                            key={rb.id}
                                            title={barTitle}
                                            onClick={(e) => {
                                              e.stopPropagation()
                                              if (Date.now() - lastRosterBarSelectionAtRef.current < 220) return
                                              if (isActualLeave) {
                                                openActualLeaveRosterDetails(rb)
                                                return
                                              }
                                              openRosterBarDetails(rb, labelType, renderStartDt, renderEndDt)
                                            }}
                                            onMouseDown={(e) => {
                                              e.preventDefault()
                                              e.stopPropagation()
                                              const target = e.currentTarget
                                              const rect = target.getBoundingClientRect()
                                              const startX = e.clientX
                                              const selectionStartDt = getSelectedRosterDateTime(e.clientX, rect)
                                              let moved = false
                                              const onMove = (ev: MouseEvent) => {
                                                if (Math.abs(ev.clientX - startX) > 4) moved = true
                                              }
                                              const onUp = (ev: MouseEvent) => {
                                                window.removeEventListener("mousemove", onMove)
                                                window.removeEventListener("mouseup", onUp)
                                                if (!moved) return
                                                lastRosterBarSelectionAtRef.current = Date.now()
                                                const selectionEndDt = getSelectedRosterDateTime(ev.clientX, rect)
                                                const selectedStartDt =
                                                  selectionStartDt <= selectionEndDt ? selectionStartDt : selectionEndDt
                                                const selectedEndDt =
                                                  selectionStartDt <= selectionEndDt ? selectionEndDt : selectionStartDt
                                                openRosterBarDetails(rb, labelType, selectedStartDt, selectedEndDt)
                                              }
                                              window.addEventListener("mousemove", onMove)
                                              window.addEventListener("mouseup", onUp)
                                            }}
                                            style={{
                                              position: "absolute",
                                              left: Math.max(leftDays, renderLeftFloatDays) * crewLeaveTimelineDayWidth + 1,
                                              top: isDutyCodeBar ? rosterBarTop + 2 : rosterBarTop,
                                              height: isDutyCodeBar ? Math.max(rosterBarHeight - 4, 10) : rosterBarHeight,
                                              width: Math.max(renderSpanFloatDays * crewLeaveTimelineDayWidth - 2, 8),
                                              borderRadius: 0,
                                              backgroundColor: isActualLeave
                                                ? "#6b7280"
                                                : isPlannedLeave
                                                  ? "#eab308"
                                                  : isDutyCodeBar
                                                    ? dutyColor || "#0ea5a5"
                                                    : hasDutyCodeOverlay
                                                      ? "#dcfce7"
                                                      : dutyColor || (isRest ? "#334155" : "#0ea5a5"),
                                              border: `1px solid ${
                                                isActualLeave
                                                  ? "#4b5563"
                                                  : isPlannedLeave
                                                    ? "#a16207"
                                                    : hasDutyCodeOverlay && !isDutyCodeBar
                                                      ? "#86efac"
                                                      : dutyColor || (isRest ? "#1e293b" : "#0f766e")
                                              }`,
                                              color: hasDutyCodeOverlay && !isDutyCodeBar ? "#14532d" : "#ffffff",
                                              fontSize: 9,
                                              fontWeight: 700,
                                              padding: "1px 4px",
                                              whiteSpace: "nowrap",
                                              overflow: "hidden",
                                              textOverflow: "ellipsis",
                                              cursor: "crosshair",
                                              zIndex: isActualLeave ? 16 : isDutyCodeBar ? 12 : 4,
                                            }}
                                          >
                                            {barLabel}
                                          </div>
                                        )
                                      })}
                                    </div>
                                  </div>
                                  )
                                })}
                              </div>
                            </div>
                          )}
                      </div>
                    ) : null}

                    {crewLeavePlannerTab === "conflicts" ? (
                      <div className="d-flex flex-column gap-3">
                        <div className="card border">
                          <div className="card-header bg-body-tertiary py-2 small fw-semibold d-flex align-items-center justify-content-between">
                            <span>Conflict Detection</span>
                            <span className="small text-muted">
                              Generated blocks scanned: {crewGeneratedBlocks.length}
                            </span>
                          </div>
                          <div className="card-body">
                            <div className="row g-2 mb-2">
                              <div className="col-6 col-md-2">
                                <label className="form-label small fw-semibold">Default limit / day</label>
                                <input
                                  type="number"
                                  min={0}
                                  className="form-control"
                                  value={crewLeaveConflictConfig.defaultThreshold}
                                  onChange={(e) =>
                                    setCrewLeaveConflictConfig((p) => ({
                                      ...p,
                                      defaultThreshold: Number(e.target.value) || 0,
                                    }))
                                  }
                                />
                                <div className="small text-muted">
                                  Max concurrent crew on leave per type
                                </div>
                              </div>
                              <div className="col-6 col-md-2">
                                <label className="form-label small fw-semibold">Crew Type</label>
                                <select
                                  className="form-select"
                                  value={crewLeaveConflictConfig.crewTypeFilter}
                                  onChange={(e) =>
                                    setCrewLeaveConflictConfig((p) => ({
                                      ...p,
                                      crewTypeFilter: e.target.value,
                                    }))
                                  }
                                >
                                  <option value="">All</option>
                                  {crewTypes.map((t) => (
                                    <option key={t} value={t}>
                                      {t}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              <div className="col-6 col-md-2">
                                <label className="form-label small fw-semibold">Severity</label>
                                <select
                                  className="form-select"
                                  value={crewLeaveConflictConfig.severityFilter}
                                  onChange={(e) =>
                                    setCrewLeaveConflictConfig((p) => ({
                                      ...p,
                                      severityFilter: e.target.value as
                                        | "all"
                                        | "high"
                                        | "medium"
                                        | "low",
                                    }))
                                  }
                                >
                                  <option value="all">All</option>
                                  <option value="high">High</option>
                                  <option value="medium">Medium</option>
                                  <option value="low">Low</option>
                                </select>
                              </div>
                              <div className="col-6 col-md-3">
                                <label className="form-label small fw-semibold">From</label>
                                <input
                                  type="date"
                                  className="form-control"
                                  value={crewLeaveConflictConfig.fromDate}
                                  onChange={(e) =>
                                    setCrewLeaveConflictConfig((p) => ({
                                      ...p,
                                      fromDate: e.target.value,
                                    }))
                                  }
                                />
                              </div>
                              <div className="col-6 col-md-3">
                                <label className="form-label small fw-semibold">To</label>
                                <input
                                  type="date"
                                  className="form-control"
                                  value={crewLeaveConflictConfig.toDate}
                                  onChange={(e) =>
                                    setCrewLeaveConflictConfig((p) => ({
                                      ...p,
                                      toDate: e.target.value,
                                    }))
                                  }
                                />
                              </div>
                            </div>

                            {crewTypes.length > 0 ? (
                              <div className="border rounded p-2 bg-body-tertiary">
                                <div className="small fw-semibold mb-2">
                                  Per-type limit override (blank = use default)
                                </div>
                                <div className="row g-2">
                                  {crewTypes.map((t) => (
                                    <div key={t} className="col-6 col-md-3">
                                      <label className="form-label small">{t}</label>
                                      <input
                                        type="number"
                                        min={0}
                                        className="form-control form-control-sm"
                                        placeholder={`Default: ${crewLeaveConflictConfig.defaultThreshold}`}
                                        value={
                                          crewLeaveConflictConfig.thresholdsByType[t] !== undefined
                                            ? String(crewLeaveConflictConfig.thresholdsByType[t])
                                            : ""
                                        }
                                        onChange={(e) => {
                                          const v = e.target.value
                                          setCrewLeaveConflictConfig((p) => {
                                            const next = { ...p.thresholdsByType }
                                            if (v === "") {
                                              delete next[t]
                                            } else {
                                              next[t] = Number(v) || 0
                                            }
                                            return { ...p, thresholdsByType: next }
                                          })
                                        }}
                                      />
                                    </div>
                                  ))}
                                </div>
                              </div>
                            ) : null}
                          </div>
                        </div>

                        <div className="row g-2">
                          <div className="col-4">
                            <div className="card border h-100">
                              <div className="card-body py-2">
                                <div className="small text-muted">High severity</div>
                                <div className="h4 mb-0 text-danger">
                                  {crewLeaveConflictsCountsBySeverity.high}
                                </div>
                              </div>
                            </div>
                          </div>
                          <div className="col-4">
                            <div className="card border h-100">
                              <div className="card-body py-2">
                                <div className="small text-muted">Medium severity</div>
                                <div className="h4 mb-0 text-warning">
                                  {crewLeaveConflictsCountsBySeverity.medium}
                                </div>
                              </div>
                            </div>
                          </div>
                          <div className="col-4">
                            <div className="card border h-100">
                              <div className="card-body py-2">
                                <div className="small text-muted">Low severity</div>
                                <div className="h4 mb-0 text-info">
                                  {crewLeaveConflictsCountsBySeverity.low}
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>

                        <div className="card border">
                          <div className="card-header bg-body-tertiary py-2 small fw-semibold d-flex align-items-center justify-content-between">
                            <span>Detected conflicts ({crewLeaveConflicts.length})</span>
                          </div>
                          <div style={{ height: "calc(100vh - 360px)", minHeight: 520 }}>
                            <DataGrid
                              rows={crewLeaveConflicts}
                              columns={[
                                {
                                  field: "severity",
                                  headerName: "Severity",
                                  width: 110,
                                  renderCell: (params: GridRenderCellParams) => {
                                    const sev = String(params.value || "")
                                    const cls =
                                      sev === "high"
                                        ? "bg-danger-subtle text-danger-emphasis"
                                        : sev === "medium"
                                          ? "bg-warning-subtle text-warning-emphasis"
                                          : "bg-info-subtle text-info-emphasis"
                                    return (
                                      <span className={"badge rounded-pill " + cls}>
                                        {sev.toUpperCase()}
                                      </span>
                                    )
                                  },
                                },
                                { field: "type", headerName: "Type", width: 130 },
                                { field: "crewCode", headerName: "Crew", minWidth: 140, flex: 1 },
                                { field: "crewName", headerName: "Name", minWidth: 160, flex: 1 },
                                { field: "crewType", headerName: "Category", width: 140 },
                                { field: "startDate", headerName: "Start", width: 120 },
                                { field: "endDate", headerName: "End", width: 120 },
                                {
                                  field: "message",
                                  headerName: "Details",
                                  minWidth: 280,
                                  flex: 2,
                                },
                              ]}
                              disableRowSelectionOnClick
                              rowHeight={34}
                              sx={gridSx}
                            />
                          </div>
                        </div>
                      </div>
                    ) : null}
                  </div>
                ) : isCrewOperationBlankModule ? (
                  <div className="card border">
                    <div className="card-body py-5 text-center">
                      <h3 className="h6 fw-semibold mb-2">{activeConfig.title}</h3>
                      <p className="text-muted mb-0">
                        Blank screen ready. This module is independent and not linked to current modules.
                      </p>
                    </div>
                  </div>
                ) : (
                  <div className="d-flex flex-column gap-3">
                    <div className="d-flex flex-wrap align-items-center gap-2">
                      {activeConfig.key !== "attendance" ? (
                        <Button
                        onClick={() => {
                          setForms((prev) => ({
                            ...prev,
                            [activeConfig.key]: buildEmptyForm(activeConfig),
                          }))
                          setStaffLookupQuery("")
                          setEditingEntryId(null)
                          setIsFormOpen(true)
                          closeInactivePopup()
                        }}
                      >
                        + Create {activeConfig.title} Entry
                      </Button>
                      ) : null}
                      {activeConfig.key === "attendance" ? (
                        <Button variant="outline" onClick={backfillAttendanceFromPortal}>
                          Backfill Portal Duty Marks
                        </Button>
                      ) : null}
                      {activeConfig.key === "publicHolidays" ? (
                        <>
                          <Button variant="outline" onClick={downloadPublicHolidayCsvTemplate}>
                            Download CSV Template
                          </Button>
                          <label className="btn btn-outline-secondary mb-0">
                            <span>Import CSV</span>
                            <input
                              type="file"
                              accept=".csv,text/csv"
                              className="d-none"
                              onChange={async (e) => {
                                const file = e.target.files?.[0]
                                if (!file) return
                                await importPublicHolidayCsv(file)
                                e.currentTarget.value = ""
                              }}
                            />
                          </label>
                        </>
                      ) : null}
                      {activeConfig.key === "crewDataBase" ? (
                        <>
                          <input
                            type="text"
                            className="form-control"
                            style={{ maxWidth: 360 }}
                            placeholder="Search Crew: Code, Name, Staff ID"
                            value={crewDbSearch}
                            onChange={(e) => setCrewDbSearch(e.target.value)}
                          />
                          <label className="btn btn-outline-secondary mb-0">
                            <span>Upload Crew CSV</span>
                            <input
                              type="file"
                              accept=".csv,text/csv"
                              className="d-none"
                              onChange={async (e) => {
                                const file = e.target.files?.[0]
                                if (!file) return
                                await importCrewDatabaseCsv(file)
                                e.currentTarget.value = ""
                              }}
                            />
                          </label>
                        </>
                      ) : null}
                    </div>

                    <div className="card border">
                      <div style={{ height: "calc(100vh - 240px)", minHeight: 560 }}>
                        <DataGrid
                          rows={filteredActiveRows}
                          disableVirtualization={activeConfig.key === "crewDataBase"}
                          columns={[
                            ...activeConfig.columns.map((column) => ({
                              field: column,
                              headerName: labelize(column),
                              minWidth:
                                activeConfig.key === "leaves" && column === "staffNo"
                                  ? 130
                                  : activeConfig.key === "leaves" && column === "staffName"
                                    ? 220
                                    : activeConfig.key === "crewDataBase" && column === "crewCode"
                                      ? 130
                                      : activeConfig.key === "crewDataBase" && column === "crewName"
                                        ? 220
                                        : activeConfig.key === "crewDataBase" && column === "crewType"
                                          ? 180
                                    : 150,
                              width:
                                activeConfig.key === "leaves" && column === "staffNo"
                                  ? 130
                                  : activeConfig.key === "leaves" && column === "staffName"
                                    ? 220
                                    : activeConfig.key === "crewDataBase" && column === "crewCode"
                                      ? 130
                                      : activeConfig.key === "crewDataBase" && column === "crewName"
                                        ? 220
                                        : activeConfig.key === "crewDataBase" && column === "crewType"
                                          ? 180
                                    : undefined,
                              flex:
                                activeConfig.key === "leaves" &&
                                (column === "staffNo" || column === "staffName")
                                  ? 0
                                  : activeConfig.key === "crewDataBase" &&
                                    (column === "crewCode" || column === "crewName" || column === "crewType")
                                    ? 0
                                  : 1,
                              headerClassName:
                                activeConfig.key === "leaves" && column === "staffNo"
                                  ? "leaves-sticky-1-header"
                                  : activeConfig.key === "leaves" && column === "staffName"
                                    ? "leaves-sticky-2-header"
                                    : activeConfig.key === "crewDataBase" && column === "crewCode"
                                      ? "crewdb-sticky-1-header"
                                      : activeConfig.key === "crewDataBase" && column === "crewName"
                                        ? "crewdb-sticky-2-header"
                                        : activeConfig.key === "crewDataBase" && column === "crewType"
                                          ? "crewdb-sticky-3-header"
                                    : undefined,
                              cellClassName:
                                activeConfig.key === "leaves" && column === "staffNo"
                                  ? "leaves-sticky-1-cell"
                                  : activeConfig.key === "leaves" && column === "staffName"
                                    ? "leaves-sticky-2-cell"
                                    : activeConfig.key === "crewDataBase" && column === "crewCode"
                                      ? "crewdb-sticky-1-cell"
                                      : activeConfig.key === "crewDataBase" && column === "crewName"
                                        ? "crewdb-sticky-2-cell"
                                        : activeConfig.key === "crewDataBase" && column === "crewType"
                                          ? "crewdb-sticky-3-cell"
                                    : undefined,
                              renderCell: (params: GridRenderCellParams) => {
                                const entry = params.row as Entry
                                if (
                                  (activeConfig.key === "staff" && column === "avatar") ||
                                  (activeConfig.key === "shift" && column === "shiftIcon")
                                ) {
                                  if (activeConfig.key === "shift" && column === "shiftIcon") {
                                    const Icon = getShiftLucideIcon(entry)
                                    return (
                                      <div
                                        className="d-flex align-items-center justify-content-center rounded-circle"
                                        style={{
                                          width: 34,
                                          height: 34,
                                          border: "2px solid rgb(29, 78, 216)",
                                          color: "rgb(29, 78, 216)",
                                          background: "rgb(255, 255, 255)",
                                          flex: "0 0 auto",
                                        }}
                                      >
                                        <Icon size={16} aria-hidden="true" />
                                      </div>
                                    )
                                  }
                                  const src = String(entry[column] || "").trim()
                                  if (!src) return <span className="text-muted">-</span>
                                  return (
                                    <img
                                      src={src}
                                      alt={column === "avatar" ? "Avatar" : "Shift Icon"}
                                      style={{
                                        width: 32,
                                        height: 32,
                                        objectFit: "cover",
                                        borderRadius: "50%",
                                        border: "1px solid var(--bs-border-color)",
                                      }}
                                    />
                                  )
                                }
                                if (column === "activeStatus") {
                                  return (
                                    <span
                                      className={
                                        "badge rounded-pill " +
                                        (entry[column] === "Active"
                                          ? "bg-success-subtle text-success-emphasis"
                                          : "bg-danger-subtle text-danger-emphasis")
                                      }
                                    >
                                      {entry[column] || "-"}
                                    </span>
                                  )
                                }
                                if (activeConfig.key === "leaves" && column === "publicHolidayDays") {
                                  return getGovPublicHolidayDisplay(entry)
                                }
                                if (
                                  activeConfig.key === "attendance" &&
                                  (column === "status" ||
                                    column === "workflowStatus" ||
                                    column === "approvalStatus" ||
                                    column === "documentStatus")
                                ) {
                                  return (
                                    <span className="badge rounded-pill" style={attendanceStatusStyle(String(entry[column] || ""))}>
                                      {String(entry[column] || "-")}
                                    </span>
                                  )
                                }
                                if (entry[column] === undefined || entry[column] === null || entry[column] === "") {
                                  return <span className="text-muted">-</span>
                                }
                                const raw = String(entry[column])
                                return isDateLikeColumn(column) ? formatDateText(raw) : raw
                              },
                            })) as GridColDef[],
                            {
                              field: "__actions",
                              headerName: "Actions",
                              minWidth: activeConfig.key === "staff" ? 230 : 170,
                              sortable: false,
                              filterable: false,
                              renderCell: (params: GridRenderCellParams) => {
                                const entry = params.row as Entry
                                return activeConfig.key === "staff" ? (
                                  <div className="d-flex align-items-center gap-2">
                                    <Button variant="ghost" size="sm" onClick={() => editEntry(activeConfig.key, entry)}>
                                      Edit
                                    </Button>
                                    <div className="form-check form-switch mb-0">
                                      <input
                                        type="checkbox"
                                        className="form-check-input"
                                        id={`active-${entry.id}`}
                                        checked={entry.activeStatus === "Active"}
                                        onChange={(e) => {
                                          if (e.target.checked) {
                                            setStaffActive(entry.id)
                                            return
                                          }
                                          openInactivePopup({
                                            staffId: entry.id,
                                            fromEdit: false,
                                            seedDate: entry.inactiveDate,
                                            seedReason: entry.inactiveReason,
                                          })
                                        }}
                                      />
                                      <label className="form-check-label small" htmlFor={`active-${entry.id}`}>
                                        Active
                                      </label>
                                    </div>
                                  </div>
                                ) : activeConfig.key === "attendance" ? (
                                  <div className="d-flex gap-2">
                                    <Button variant="ghost" size="sm" onClick={() => editEntry(activeConfig.key, entry)}>
                                      Edit
                                    </Button>
                                    <Button
                                      variant="outline"
                                      size="sm"
                                      onClick={() => markAttendanceSubmitDocument(entry.id)}
                                    >
                                      Set Upload Document
                                    </Button>
                                    <Button
                                      variant="outline"
                                      size="sm"
                                      onClick={() => markAttendanceProceedApproval(entry.id)}
                                    >
                                      Set Pending Approval
                                    </Button>
                                    <Button
                                      size="sm"
                                      onClick={() => approveAttendanceRequest(entry.id)}
                                    >
                                      Approve
                                    </Button>
                                    <Button
                                      variant="destructive"
                                      size="sm"
                                      onClick={() => rejectAttendanceRequest(entry.id)}
                                    >
                                      Reject
                                    </Button>
                                  </div>
                                ) : (
                                  <div className="d-flex gap-2">
                                    <Button variant="ghost" size="sm" onClick={() => editEntry(activeConfig.key, entry)}>
                                      Edit
                                    </Button>
                                    <Button variant="destructive" size="sm" onClick={() => removeEntry(activeConfig.key, entry.id)}>
                                      Delete
                                    </Button>
                                  </div>
                                )
                              },
                            },
                          ]}
                          disableRowSelectionOnClick
                          rowHeight={38}
                          sx={gridSx}
                          pageSizeOptions={[10, 20, 50, 100]}
                          initialState={{
                            pagination: {
                              paginationModel: { pageSize: 20, page: 0 },
                            },
                          }}
                        />
                      </div>
                    </div>
                  </div>
                )}
            </div>
          </div>
        </main>
      </div>

      {isRosterChangeModalOpen && rosterChangeCell ? (
        <>
          <div className="app-modal-backdrop" />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1058 }}>
            <div className="modal-dialog modal-dialog-centered">
              <div className="modal-content">
                <div className="modal-header">
                  <h5 className="modal-title fw-semibold">Roster Change Request</h5>
                  <button
                    type="button"
                    className="btn-close"
                    aria-label="Close"
                    onClick={() => setIsRosterChangeModalOpen(false)}
                  />
                </div>
                <div className="modal-body">
                  <div className="small text-muted mb-2">
                    {rosterChangeCell.staffNo} - {rosterChangeCell.staffName} | {formatYmdToDdMmYy(rosterChangeCell.date)} | Current: {rosterChangeCell.code}
                  </div>
                  <div className="mb-3">
                    <label className="form-label small fw-semibold">Change Type</label>
                    <select
                      className="form-select"
                      value={rosterChangeType}
                      onChange={(e) => setRosterChangeType(e.target.value as RosterChangeRequest["changeType"])}
                    >
                      <option>Mutual Swap</option>
                      <option>Shift Reassign</option>
                      <option>State Override</option>
                    </select>
                  </div>
                  {rosterChangeType === "Mutual Swap" ? (
                    <div className="mb-3">
                      <label className="form-label small fw-semibold">Swap With</label>
                      <select
                        className="form-select"
                        value={rosterChangeTargetStaff}
                        onChange={(e) => setRosterChangeTargetStaff(e.target.value)}
                      >
                        <option value="">Select staff...</option>
                        {rosterChangeTargetStaffOptions.map((opt) => (
                          <option key={opt.value} value={opt.value}>
                            {opt.label}
                          </option>
                        ))}
                      </select>
                    </div>
                  ) : (
                    <div className="mb-3">
                      <label className="form-label small fw-semibold">Changed Code</label>
                      <select
                        className="form-select"
                        value={rosterChangeToCode}
                        onChange={(e) => setRosterChangeToCode(e.target.value)}
                      >
                        {rosterChangeCodeOptions.map((code) => (
                          <option key={code} value={code}>
                            {code}
                          </option>
                        ))}
                      </select>
                      {rosterChangeConflict ? (
                        <div className="mt-2 border rounded p-2 bg-light-subtle">
                          <div className="small fw-semibold text-danger-emphasis mb-1">
                            Conflict: code {rosterChangeToCode.toUpperCase()} is assigned to{" "}
                            {rosterChangeConflict.staffNo} - {rosterChangeConflict.staffName}
                          </div>
                          <label className="form-label small fw-semibold mb-1">
                            Change Other Staff Shift To
                          </label>
                          <select
                            className="form-select form-select-sm"
                            value={rosterChangeConflictToCode}
                            onChange={(e) => setRosterChangeConflictToCode(e.target.value)}
                          >
                            <option value="">Remove other staff code</option>
                            {rosterChangeCodeOptions.map((code) => (
                              <option key={`conflict-${code}`} value={code}>
                                {code}
                              </option>
                            ))}
                          </select>
                        </div>
                      ) : null}
                    </div>
                  )}
                  <div className="mb-1">
                    <label className="form-label small fw-semibold">State Change Reason</label>
                    <textarea
                      className="form-control"
                      rows={3}
                      value={rosterChangeReason}
                      onChange={(e) => setRosterChangeReason(e.target.value)}
                    />
                  </div>
                  <div className="form-text">Request will be submitted as Pending Approval.</div>
                </div>
                <div className="modal-footer">
                  <Button variant="outline" onClick={() => setIsRosterChangeModalOpen(false)}>
                    Cancel
                  </Button>
                  <Button onClick={submitRosterChangeRequest}>Pass to Approval</Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}

      {isFormOpen ? (
        <>
          <div className="app-modal-backdrop" />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1055 }}>
            <div className="modal-dialog modal-lg modal-dialog-centered modal-dialog-scrollable">
              <div className="modal-content">
                <div className="modal-header">
                  <div>
                    <h5 className="modal-title fw-semibold">
                      {editingEntryId ? `Edit ${activeConfig.title} Entry` : `Create ${activeConfig.title} Entry`}
                    </h5>
                    <p className="small text-muted mb-0">{activeConfig.description}</p>
                  </div>
                  <button
                    type="button"
                    className="btn-close"
                    aria-label="Close"
                    onClick={() => {
                      setIsFormOpen(false)
                      setEditingEntryId(null)
                      closeInactivePopup()
                    }}
                  />
                </div>
                <div className="modal-body">
                  {activeConfig.key === "attendance" ? (
                    <div className="mb-3 border rounded p-2 bg-body-tertiary">
                      <div className="small fw-semibold mb-2">Request Workflow</div>
                      <Box sx={{ maxWidth: 420 }}>
                        <Stepper
                          activeStep={getAttendanceStepIndexFromForm(currentForm)}
                          orientation="vertical"
                          sx={workflowStepperSx}
                        >
                          <Step>
                            <StepLabel>Pending Request</StepLabel>
                          </Step>
                          <Step>
                            <StepLabel>Document Stage</StepLabel>
                          </Step>
                          <Step>
                            <StepLabel>Approval Stage</StepLabel>
                          </Step>
                          <Step>
                            <StepLabel>
                              {(currentForm.status || "").trim().toLowerCase() === "rejected"
                                ? "Final: Rejected"
                                : "Final: Approved"}
                            </StepLabel>
                          </Step>
                        </Stepper>
                      </Box>
                    </div>
                  ) : null}
                  <div className="row g-3">
                    {activeConfig.fields.map((field) => {
                      if (
                        field.key === "consecutiveDaysWithoutDocument" &&
                        currentForm.allowNoOfDaysWithoutDocument !== "Yes"
                      ) {
                        return null
                      }
                      if (
                        field.key === "allowNoOfDaysWithoutDocument" &&
                        currentForm.documentRequired === "No"
                      ) {
                        return null
                      }
                      if (
                        activeConfig.key === "rosterPriorityType" &&
                        field.key === "requiredWeekdays" &&
                        currentForm.requiredDaily !== "No"
                      ) {
                        return null
                      }
                      if (field.key === "inactiveDate" && currentForm.activeStatus !== "Inactive") {
                        return null
                      }
                      if (field.key === "inactiveReason" && currentForm.activeStatus !== "Inactive") {
                        return null
                      }
                      if (
                        activeConfig.key === "userManagement" &&
                        field.key === "dispatchStaff" &&
                        currentForm.userType !== "Dispatch Staff User"
                      ) {
                        return null
                      }
                      if (
                        activeConfig.key === "userManagement" &&
                        field.key === "fullName" &&
                        currentForm.userType === "Dispatch Staff User"
                      ) {
                        return null
                      }
                      if (activeConfig.key === "shift" && field.key === "shiftIcon") {
                        return null
                      }

                      return (
                        <div key={field.key} className="col-12 col-md-6">
                          <label className="form-label small fw-semibold">
                            {field.label}
                            {field.required ? <span className="text-danger ms-1">*</span> : null}
                          </label>
                          {(activeConfig.key === "staff" || activeConfig.key === "crewDataBase") &&
                          field.key === "activeStatus" ? (
                            <div className="border rounded p-2">
                              <div className="form-check form-switch mb-0">
                                <input
                                  type="checkbox"
                                  className="form-check-input"
                                  id="staff-active-form"
                                  checked={currentForm.activeStatus === "Active"}
                                  onChange={(e) => {
                                    if (e.target.checked) {
                                      updateForm("activeStatus", "Active")
                                      updateForm("inactiveDate", "")
                                      updateForm("inactiveReason", "")
                                      return
                                    }
                                    openInactivePopup({
                                      staffId: null,
                                      fromEdit: true,
                                      seedDate: currentForm.inactiveDate,
                                      seedReason: currentForm.inactiveReason,
                                    })
                                  }}
                                />
                                <label className="form-check-label small" htmlFor="staff-active-form">
                                  {currentForm.activeStatus === "Active" ? "Active" : "Inactive"}
                                </label>
                              </div>
                            </div>
                          ) : field.type === "checkbox" ? (
                            <div className="border rounded p-2">
                              <div className="form-check form-switch mb-0">
                                <input
                                  type="checkbox"
                                  className="form-check-input"
                                  id={`field-${activeConfig.key}-${field.key}`}
                                  checked={(currentForm[field.key] || "No") === "Yes"}
                                  onChange={(e) =>
                                    updateForm(field.key, e.target.checked ? "Yes" : "No")
                                  }
                                />
                                <label
                                  className="form-check-label small"
                                  htmlFor={`field-${activeConfig.key}-${field.key}`}
                                >
                                  {(currentForm[field.key] || "No") === "Yes" ? "Yes" : "No"}
                                </label>
                              </div>
                            </div>
                          ) : field.type === "staffLookup" ? (
                            <div className="border rounded p-2">
                              <input
                                type="text"
                                list="staff-lookup-options"
                                className="form-control"
                                placeholder="Search by staff no or full name"
                                value={staffLookupQuery}
                                onChange={(e) => {
                                  const value = e.target.value
                                  setStaffLookupQuery(value)
                                  const matched = store.staff.find(
                                    (staff) => `${staff.staffNo} - ${staff.fullName}` === value,
                                  )
                                  if (matched) {
                                    updateForm("staffLookup", value)
                                    updateForm("staffNo", matched.staffNo)
                                    updateForm("staffName", matched.fullName)
                                  } else {
                                    updateForm("staffLookup", value)
                                    updateForm("staffNo", "")
                                    updateForm("staffName", "")
                                  }
                                }}
                              />
                              <datalist id="staff-lookup-options">
                                {staffLookupResults.map((staff) => (
                                  <option key={staff.id} value={`${staff.staffNo} - ${staff.fullName}`} />
                                ))}
                              </datalist>
                              <div className="form-text small">
                                Type to search, then pick from dropdown suggestions.
                              </div>
                            </div>
                          ) : field.type === "select" ? (
                            <select
                              className="form-select"
                              value={currentForm[field.key] ?? ""}
                              onChange={(e) => {
                                updateForm(field.key, e.target.value)
                                if (activeConfig.key === "userManagement" && field.key === "userType") {
                                  if (e.target.value === "Dispatch Staff User") {
                                    updateForm("fullName", "")
                                  } else {
                                    updateForm("dispatchStaff", "")
                                  }
                                }
                                if (activeConfig.key === "userManagement" && field.key === "dispatchStaff") {
                                  const parts = e.target.value.split(" - ")
                                  updateForm("fullName", parts.slice(1).join(" - ") || "")
                                }
                                if (
                                  field.key === "allowNoOfDaysWithoutDocument" &&
                                  e.target.value !== "Yes"
                                ) {
                                  updateForm("consecutiveDaysWithoutDocument", "")
                                }
                                if (field.key === "documentRequired" && e.target.value === "No") {
                                  updateForm("allowNoOfDaysWithoutDocument", "No")
                                  updateForm("consecutiveDaysWithoutDocument", "")
                                }
                                if (field.key === "activeStatus" && e.target.value === "Active") {
                                  updateForm("inactiveDate", "")
                                  updateForm("inactiveReason", "")
                                }
                              }}
                            >
                              {getSelectOptions(field).length === 0 ? (
                                <option value="">No options available</option>
                              ) : (
                                getSelectOptions(field).map((option) => (
                                  <option key={option} value={option}>{option}</option>
                                ))
                              )}
                            </select>
                          ) : field.type === "multiselect" ? (
                            <div className="border rounded p-2">
                              <div className="text-uppercase text-muted small fw-semibold mb-2" style={{ fontSize: "0.7rem", letterSpacing: "0.05em" }}>
                                Select one or more
                              </div>
                              <div className="d-flex flex-wrap gap-2">
                                {getMultiSelectOptions(field).map((option) => {
                                  const normalizeForField = (value: string) => {
                                    if (
                                      activeConfig.key === "levelShiftPriority" &&
                                      field.key === "rosterPriorityNumbers"
                                    ) {
                                      const p = parseLeadingNumber(value)
                                      return p === null ? value.trim() : String(p)
                                    }
                                    return value.trim()
                                  }
                                  const selectedValues = (currentForm[field.key] ?? "")
                                    .split(",")
                                    .map((item) => normalizeForField(item))
                                    .filter(Boolean)
                                  const checked = selectedValues.includes(normalizeForField(option))
                                  return (
                                    <label
                                      key={option}
                                      className={
                                        "btn btn-sm " +
                                        (checked ? "btn-primary" : "btn-outline-secondary") +
                                        " d-inline-flex align-items-center gap-2 mb-0"
                                      }
                                    >
                                      <input
                                        type="checkbox"
                                        className="form-check-input mt-0"
                                        checked={checked}
                                        onChange={() => toggleMultiSelectValue(field.key, option)}
                                        style={{ marginLeft: 0 }}
                                      />
                                      <span>{option}</span>
                                    </label>
                                  )
                                })}
                              </div>
                              {getMultiSelectOptions(field).length === 0 ? (
                                <div className="form-text small">
                                  No levels found yet. Create staff levels first.
                                </div>
                              ) : null}
                            </div>
                          ) : field.type === "textarea" ? (
                            <textarea
                              className="form-control"
                              rows={3}
                              value={currentForm[field.key] ?? ""}
                              onChange={(e) => updateForm(field.key, e.target.value)}
                              placeholder={`Enter ${field.label.toLowerCase()}`}
                            />
                          ) : field.type === "image" ? (
                            <div className="border rounded p-2">
                              <input
                                type="file"
                                accept="image/*"
                                className="form-control"
                                onChange={(e) => {
                                  const file = e.target.files?.[0]
                                  if (!file) return
                                  const reader = new FileReader()
                                  reader.onload = () => {
                                    updateForm(field.key, String(reader.result || ""))
                                  }
                                  reader.readAsDataURL(file)
                                }}
                              />
                              {currentForm[field.key] ? (
                                <div className="mt-2 d-flex align-items-center gap-2">
                                  <img
                                    src={currentForm[field.key]}
                                    alt="Preview"
                                    style={{
                                      width: 40,
                                      height: 40,
                                      objectFit: "cover",
                                      borderRadius: "50%",
                                      border: "1px solid var(--bs-border-color)",
                                    }}
                                  />
                                  <Button
                                    type="button"
                                    variant="outline"
                                    size="sm"
                                    onClick={() => updateForm(field.key, "")}
                                  >
                                    Remove
                                  </Button>
                                </div>
                              ) : null}
                            </div>
                          ) : (
                            <input
                              type={field.type}
                              className="form-control"
                              value={currentForm[field.key] ?? ""}
                              onChange={(e) => updateForm(field.key, e.target.value)}
                            />
                          )}
                        </div>
                      )
                    })}
                  </div>
                </div>
                <div className="modal-footer">
                  <Button
                    variant="outline"
                    onClick={() => {
                      setIsFormOpen(false)
                      setEditingEntryId(null)
                      closeInactivePopup()
                    }}
                  >
                    Cancel
                  </Button>
                  <Button onClick={submitActiveModule}>
                    {editingEntryId ? "Update Entry" : `Save ${activeConfig.title} Entry`}
                  </Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}

      {leaveDetailModal.open ? (
        <>
          <div className="app-modal-backdrop" style={{ zIndex: 1068 }} />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1069 }}>
            <div className="modal-dialog modal-xl modal-dialog-scrollable">
              <div className="modal-content">
                <div className="modal-header">
                  <h5 className="modal-title fw-semibold">
                    Leave Details - {leaveDetailModal.crewCode}
                    {leaveDetailModal.crewName ? ` | ${leaveDetailModal.crewName}` : ""}
                  </h5>
                  <button
                    type="button"
                    className="btn-close"
                    aria-label="Close"
                    onClick={() => {
                      setLeaveDetailModal((p) => ({ ...p, open: false }))
                      setLeaveDetailSelectedCycle("")
                      setPendingLeaveBarAdjustment(null)
                      setEditingLeaveSegmentId("")
                      setEditingLeaveSegmentDraft({ cycleNumber: "", start: "", end: "" })
                    }}
                  />
                </div>
                <div className="modal-body d-flex flex-column gap-3">
                  {lastLeaveBarAdjustment &&
                  lastLeaveBarAdjustment.crewCode === normalizeCrewCode(leaveDetailModal.crewCode || "") &&
                  lastLeaveBarAdjustment.cycleNumber === leaveDetailActiveCycleNumber ? (
                    <div className="alert alert-warning py-2 small mb-0">
                      Adjusted Days: Leave Pattern{" "}
                      <strong>
                        {lastLeaveBarAdjustment.leaveDeltaDays >= 0 ? "+" : ""}
                        {lastLeaveBarAdjustment.leaveDeltaDays}
                      </strong>
                      {" | "}Next Work Pattern{" "}
                      <strong>
                        {lastLeaveBarAdjustment.nextWorkDeltaDays >= 0 ? "+" : ""}
                        {lastLeaveBarAdjustment.nextWorkDeltaDays}
                      </strong>
                    </div>
                  ) : null}
                  {pendingLeaveBarAdjustment &&
                  pendingLeaveBarAdjustment.crewCode === normalizeCrewCode(leaveDetailModal.crewCode || "") &&
                  pendingLeaveBarAdjustment.cycleNumber === leaveDetailActiveCycleNumber ? (
                    <div className="alert alert-info py-2 small mb-0">
                      Pending Drag Adjustment: Leave Pattern{" "}
                      <strong>
                        {pendingLeaveBarAdjustment.deltaDays >= 0 ? "+" : ""}
                        {pendingLeaveBarAdjustment.deltaDays}
                      </strong>
                      {" day(s). Use Apply Drag Adjustment to confirm."}
                    </div>
                  ) : null}
                  <div className="card border">
                    <div className="card-body py-2">
                      <div className="row g-2 small">
                        <div className="col-md-3"><strong>Cycle:</strong> {leaveDetailModal.cycleNumber || "-"}</div>
                        <div className="col-md-3"><strong>Rotation:</strong> {leaveDetailModal.policyName || "-"}</div>
                        <div className="col-md-3"><strong>Leave Start:</strong> {leaveDetailModal.leaveStart || "-"}</div>
                        <div className="col-md-3"><strong>Leave End:</strong> {leaveDetailModal.leaveEnd || "-"}</div>
                      </div>
                    </div>
                  </div>

                  <div className="card border">
                    <div className="card-header py-2 small fw-semibold">Set Actual Leave (Mapped to this leave cycle)</div>
                    <div className="card-body">
                      <div className="row g-2 align-items-start">
                        <div className="col-12 col-md-3 d-flex flex-column">
                          <label className="form-label small fw-semibold">Cycle</label>
                          <select
                            className="form-select form-select-sm"
                            style={{ maxWidth: 220 }}
                            value={leaveDetailActiveCycleNumber}
                            onChange={(e) => {
                              const nextCycle = e.target.value
                              setLeaveDetailSelectedCycle(nextCycle)
                              const row = leaveDetailPatternRows.find((r) => (r.cycleNumber || "").trim() === nextCycle)
                              const cycleStart = (row?.lpStart || "").trim()
                              const cycleEnd = (row?.lpEnd || "").trim()
                              setLeaveDetailDraft({ start: cycleStart, end: cycleEnd || cycleStart })
                            }}
                          >
                            {leaveDetailPatternRows.map((r) => (
                              <option key={r.id} value={r.cycleNumber}>
                                {r.cycleLabel}
                              </option>
                            ))}
                          </select>
                          <div className="small text-muted mt-1" style={{ minHeight: 18 }}>
                            <span className="text-muted">REM / Leave Balance: </span>
                            <strong>{leaveDetailCycleRemainingDays}</strong>
                          </div>
                        </div>
                        <div className="col-12 col-md-3 d-flex flex-column">
                          <label className="form-label small fw-semibold">Actual Leave Start</label>
                          <input
                            type="date"
                            className="form-control form-control-sm"
                            style={{ maxWidth: 220 }}
                            value={leaveDetailDraft.start}
                            onChange={(e) => {
                              const start = e.target.value
                              setLeaveDetailDraft((prev) => ({ ...prev, start }))
                            }}
                          />
                          <div style={{ minHeight: 18 }} />
                        </div>
                        <div className="col-12 col-md-3 d-flex flex-column">
                          <label className="form-label small fw-semibold">Actual Leave End</label>
                          <input
                            type="date"
                            className="form-control form-control-sm"
                            style={{ maxWidth: 220 }}
                            value={leaveDetailDraft.end}
                            onChange={(e) =>
                              setLeaveDetailDraft((prev) => ({ ...prev, end: e.target.value }))
                            }
                          />
                          <div className="small text-muted mt-1" style={{ minHeight: 18 }}>
                            Remaining balance in this cycle: {leaveDetailCycleRemainingDays} day
                            {leaveDetailCycleRemainingDays === 1 ? "" : "s"}
                          </div>
                        </div>
                        <div className="col-12 col-md-3 d-flex justify-content-md-end align-items-start pt-4">
                          <Button
                            onClick={() =>
                              addActualLeaveSegment({
                                crewCode: leaveDetailModal.crewCode,
                                cycleNumber: leaveDetailActiveCycleNumber,
                                startDate: leaveDetailDraft.start,
                                endDate: leaveDetailDraft.end,
                                maxDays: leaveDetailCycleRemainingDays,
                              })
                            }
                          >
                            Save Actual Leave Segment
                          </Button>
                        </div>
                      </div>
                      <div className="mt-3">
                        <div className="small fw-semibold mb-1">Saved Segments (this cycle)</div>
                        <div className="small text-muted mb-2">
                          Cycle: {leaveDetailActiveCycleNumber || "-"} | Deduction: {leaveDetailActiveCycleRow?.adjust || "0"} day
                          {(leaveDetailActiveCycleRow?.adjust || "0") === "1" ? "" : "s"} | REM: {leaveDetailCycleRemainingDays} day
                          {leaveDetailCycleRemainingDays === 1 ? "" : "s"}
                        </div>
                        <div className="table-responsive mb-2">
                          <table className="table table-sm mb-0 align-middle" style={{ fontSize: 12 }}>
                            <thead>
                              <tr>
                                <th style={{ minWidth: 84 }}>Cycle</th>
                                <th style={{ minWidth: 90 }}>Code</th>
                                <th style={{ minWidth: 90 }}>Marked</th>
                                <th style={{ minWidth: 90 }}>Threshold</th>
                                <th style={{ minWidth: 100 }}>Deducted</th>
                              </tr>
                            </thead>
                            <tbody>
                              {leaveDetailCycleDeductionRows.length ? (
                                leaveDetailCycleDeductionRows.map((r) => (
                                  <tr key={`${leaveDetailActiveCycleNumber}_${r.code}`}>
                                    <td>{leaveDetailActiveCycleNumber || "-"}</td>
                                    <td>{r.code}</td>
                                    <td>{r.markedDays}</td>
                                    <td>{r.threshold}</td>
                                    <td className={r.deductedDays > 0 ? "text-danger fw-semibold" : ""}>{r.deductedDays}</td>
                                  </tr>
                                ))
                              ) : (
                                <tr>
                                  <td colSpan={5} className="text-muted">No deduction code records in this cycle.</td>
                                </tr>
                              )}
                            </tbody>
                          </table>
                        </div>
                        <div className="table-responsive">
                          {leaveDetailSegmentsForActiveCycle.length ? (
                            <table className="table table-sm mb-0 align-middle" style={{ fontSize: 12 }}>
                              <thead>
                                <tr>
                                  <th style={{ minWidth: 84 }}>Cycle</th>
                                  <th style={{ minWidth: 108 }}>Start</th>
                                  <th style={{ minWidth: 108 }}>End</th>
                                  <th style={{ minWidth: 80 }}>Days</th>
                                  <th style={{ minWidth: 90 }}>Deduction</th>
                                  <th style={{ minWidth: 70 }} />
                                </tr>
                              </thead>
                              <tbody>
                                {leaveDetailSegmentsForActiveCycle.map((s) => {
                                  const isEditingSegment = editingLeaveSegmentId === s.id
                                  const editCycle = isEditingSegment
                                    ? editingLeaveSegmentDraft.cycleNumber
                                    : s.cycleNumber || leaveDetailActiveCycleNumber || ""
                                  const editStart = isEditingSegment ? editingLeaveSegmentDraft.start : s.startDate || ""
                                  const editEnd = isEditingSegment ? editingLeaveSegmentDraft.end : s.endDate || ""
                                  const editDays =
                                    editStart && editEnd && editEnd >= editStart
                                      ? getDateRangeInclusive(editStart, editEnd).length
                                      : Number(s.plannedDays || "0") || 0
                                  return (
                                    <tr key={s.id}>
                                      <td>
                                        {isEditingSegment ? (
                                          <select
                                            className="form-select form-select-sm"
                                            value={editCycle}
                                            onChange={(e) =>
                                              setEditingLeaveSegmentDraft((p) => ({
                                                ...p,
                                                cycleNumber: e.target.value,
                                              }))
                                            }
                                          >
                                            {leaveDetailPatternRows.map((r) => (
                                              <option key={`edit-seg-cycle-${s.id}-${r.id}`} value={r.cycleNumber}>
                                                {r.cycleLabel}
                                              </option>
                                            ))}
                                          </select>
                                        ) : (
                                          s.cycleNumber || leaveDetailActiveCycleNumber || "-"
                                        )}
                                      </td>
                                      <td>
                                        {isEditingSegment ? (
                                          <input
                                            type="date"
                                            className="form-control form-control-sm"
                                            value={editStart}
                                            onChange={(e) =>
                                              setEditingLeaveSegmentDraft((p) => ({ ...p, start: e.target.value }))
                                            }
                                          />
                                        ) : (
                                          formatYmdToDdMmYy(s.startDate || "")
                                        )}
                                      </td>
                                      <td>
                                        {isEditingSegment ? (
                                          <input
                                            type="date"
                                            className="form-control form-control-sm"
                                            value={editEnd}
                                            onChange={(e) =>
                                              setEditingLeaveSegmentDraft((p) => ({ ...p, end: e.target.value }))
                                            }
                                          />
                                        ) : (
                                          formatYmdToDdMmYy(s.endDate || "")
                                        )}
                                      </td>
                                      <td>{isEditingSegment ? editDays : s.plannedDays || "0"}</td>
                                      <td>{leaveDetailActiveCycleRow?.adjust || "0"}</td>
                                      <td className="text-end">
                                        {isEditingSegment ? (
                                          <div className="d-flex justify-content-end gap-2">
                                            <button
                                              type="button"
                                              className="btn btn-sm btn-link p-0"
                                              onClick={() => updateLeavePatternActualSegment(s.id)}
                                            >
                                              Save
                                            </button>
                                            <button
                                              type="button"
                                              className="btn btn-sm btn-link text-muted p-0"
                                              onClick={() => {
                                                setEditingLeaveSegmentId("")
                                                setEditingLeaveSegmentDraft({ cycleNumber: "", start: "", end: "" })
                                              }}
                                            >
                                              Cancel
                                            </button>
                                          </div>
                                        ) : (
                                          <div className="d-flex justify-content-end gap-2">
                                            <button
                                              type="button"
                                              className="btn btn-sm btn-link p-0"
                                              onClick={() => {
                                                setEditingLeaveSegmentId(s.id)
                                                setEditingLeaveSegmentDraft({
                                                  cycleNumber: s.cycleNumber || leaveDetailActiveCycleNumber || "",
                                                  start: s.startDate || "",
                                                  end: s.endDate || s.startDate || "",
                                                })
                                              }}
                                            >
                                              Edit
                                            </button>
                                            <button
                                              type="button"
                                              className="btn btn-sm btn-link text-danger p-0"
                                              onClick={() => removeLeavePatternActualSegment(s.id)}
                                            >
                                              Delete
                                            </button>
                                          </div>
                                        )}
                                      </td>
                                    </tr>
                                  )
                                })}
                              </tbody>
                            </table>
                          ) : (
                            <div className="small text-muted">No actual leave segments yet.</div>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>

                  {leaveDetailUnassignedActualLeaveSegments.length ? (
                    <div className="card border border-warning-subtle">
                      <div className="card-header py-2 small fw-semibold d-flex align-items-center justify-content-between">
                        <span>Unassigned AL Log</span>
                        <span className="badge text-bg-warning">{leaveDetailUnassignedActualLeaveSegments.length}</span>
                      </div>
                      <div className="card-body p-0">
                        <div className="table-responsive">
                          <table className="table table-sm mb-0 align-middle" style={{ fontSize: 12 }}>
                            <thead>
                              <tr>
                                <th>Start</th>
                                <th>End</th>
                                <th>Days</th>
                                <th>Current Cycle</th>
                                <th style={{ width: 140 }} />
                              </tr>
                            </thead>
                            <tbody>
                              {leaveDetailUnassignedActualLeaveSegments.map((seg) => (
                                <tr key={`unassigned-al-${seg.id}`}>
                                  <td>{formatYmdToDdMmYy(seg.startDate || "")}</td>
                                  <td>{formatYmdToDdMmYy(seg.endDate || seg.startDate || "")}</td>
                                  <td>{seg.plannedDays || "0"}</td>
                                  <td>{seg.cycleNumber || "Not assigned"}</td>
                                  <td className="text-end">
                                    <Button
                                      size="sm"
                                      variant="outline"
                                      onClick={() => {
                                        setLeaveDetailModal((p) => ({
                                          ...p,
                                          actualLeaveSegmentId: seg.id,
                                          cycleNumber: (seg.cycleNumber || "").trim(),
                                          leaveStart: seg.startDate || p.leaveStart,
                                          leaveEnd: seg.endDate || seg.startDate || p.leaveEnd,
                                        }))
                                        setLeaveDetailSelectedCycle((seg.cycleNumber || "").trim())
                                        setLeaveDetailDraft({
                                          start: seg.startDate || "",
                                          end: seg.endDate || seg.startDate || "",
                                        })
                                      }}
                                    >
                                      Select
                                    </Button>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  ) : null}

                  <div className="card border">
                    <div className="card-header py-2 small fw-semibold">Leave Pattern Manager View (Crew)</div>
                    <div className="card-body p-0">
                      <div style={{ maxHeight: 320, overflow: "auto" }}>
                        <table className="table table-sm mb-0 align-middle" style={{ fontSize: 11 }}>
                          <thead>
                            <tr>
                              <th rowSpan={2} style={{ minWidth: 78, background: "#f8fafc", padding: "4px 6px" }}>Cycle</th>
                              <th colSpan={4} className="text-center" style={{ background: "#dbeafe", color: "#1e3a8a", padding: "4px 6px" }}>Work Pattern</th>
                              <th colSpan={4} className="text-center" style={{ background: "#dcfce7", color: "#166534", padding: "4px 6px" }}>Leave Pattern</th>
                            </tr>
                            <tr>
                              <th style={{ minWidth: 104, padding: "4px 6px" }}>Start</th>
                              <th style={{ minWidth: 104, padding: "4px 6px" }}>End</th>
                              <th style={{ minWidth: 78, padding: "4px 6px" }}>No of Days</th>
                              <th style={{ minWidth: 112, padding: "4px 6px" }}>Adjust</th>
                              <th style={{ minWidth: 104, padding: "4px 6px" }}>Start</th>
                              <th style={{ minWidth: 104, padding: "4px 6px" }}>End</th>
                              <th style={{ minWidth: 78, padding: "4px 6px" }}>No of Days</th>
                              <th style={{ minWidth: 70, padding: "4px 6px" }}>REM</th>
                              <th style={{ minWidth: 92, padding: "4px 6px" }}>Reset</th>
                            </tr>
                          </thead>
                          <tbody>
                            {leaveDetailPatternRows.length ? (
                              leaveDetailPatternRows.map((r) => (
                                <tr
                                  key={r.id}
                                  className={(r.cycleNumber || "").trim() === leaveDetailActiveCycleNumber ? "table-warning" : ""}
                                >
                                  <td className="fw-semibold">{r.cycleLabel}</td>
                                  <td>{formatYmdToDdMmYy(r.wpStart)}</td>
                                  <td>{formatYmdToDdMmYy(r.wpEnd)}</td>
                                  <td>{r.wpDays}</td>
                                  <td className={Number(r.adjust || "0") > 0 ? "text-danger fw-semibold" : ""}>
                                    {(() => {
                                      const w = diffDaysYmd(r.wpOriginalStart || "", r.wpStart || "")
                                      const l = diffDaysYmd(r.lpOriginalStart || "", r.lpStart || "")
                                      const wpAdj = Number.isFinite(w) ? w : 0
                                      const lpAdj = Number.isFinite(l) ? l : 0
                                      const ded = Number(r.adjust || "0") || 0
                                      return `D-${ded} | W${wpAdj >= 0 ? "+" : ""}${wpAdj} | L${lpAdj >= 0 ? "+" : ""}${lpAdj}`
                                    })()}
                                  </td>
                                  <td>{formatYmdToDdMmYy(r.lpStart)}</td>
                                  <td>{formatYmdToDdMmYy(r.lpEnd)}</td>
                                  <td>{r.lpDays}</td>
                                  <td className="fw-semibold">{r.rem}</td>
                                  <td>
                                    <Button
                                      size="sm"
                                      variant="outline"
                                      onClick={() =>
                                        resetGeneratedCycleToOriginal(
                                          leaveDetailModal.crewCode,
                                          String(r.cycleNumber || "").trim(),
                                        )
                                      }
                                    >
                                      Reset
                                    </Button>
                                  </td>
                                </tr>
                              ))
                            ) : (
                              <tr>
                                <td colSpan={10} className="text-muted">No generated leave pattern rows for this crew.</td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
                <div className="modal-footer">
                  {pendingLeaveBarAdjustment &&
                  pendingLeaveBarAdjustment.crewCode === normalizeCrewCode(leaveDetailModal.crewCode || "") &&
                  pendingLeaveBarAdjustment.cycleNumber === leaveDetailActiveCycleNumber ? (
                    <Button
                      onClick={() => {
                        const adj = shiftLeaveBarAndNextWork(
                          pendingLeaveBarAdjustment.blockId,
                          pendingLeaveBarAdjustment.deltaDays,
                        )
                        if (adj) {
                          setLastLeaveBarAdjustment(adj)
                          lastLeaveBarDragAtRef.current = Date.now()
                        }
                        setPendingLeaveBarAdjustment(null)
                      }}
                    >
                      Apply Drag Adjustment
                    </Button>
                  ) : null}
                  {pendingLeaveBarAdjustment &&
                  pendingLeaveBarAdjustment.crewCode === normalizeCrewCode(leaveDetailModal.crewCode || "") &&
                  pendingLeaveBarAdjustment.cycleNumber === leaveDetailActiveCycleNumber ? (
                    <Button variant="outline" onClick={() => setPendingLeaveBarAdjustment(null)}>
                      Cancel Pending Drag
                    </Button>
                  ) : null}
                  {lastLeaveBarAdjustment &&
                  lastLeaveBarAdjustment.crewCode === normalizeCrewCode(leaveDetailModal.crewCode || "") &&
                  lastLeaveBarAdjustment.cycleNumber === leaveDetailActiveCycleNumber ? (
                    <Button
                      variant="outline"
                      onClick={() => {
                        resetGeneratedCycleToOriginal(leaveDetailModal.crewCode, leaveDetailActiveCycleNumber)
                        setLastLeaveBarAdjustment(null)
                      }}
                    >
                      Cancel Last Drag
                    </Button>
                  ) : null}
                  <Button
                    variant="outline"
                    onClick={() => {
                      setLeaveDetailModal((p) => ({ ...p, open: false }))
                      setLeaveDetailSelectedCycle("")
                      setPendingLeaveBarAdjustment(null)
                      setEditingLeaveSegmentId("")
                      setEditingLeaveSegmentDraft({ cycleNumber: "", start: "", end: "" })
                    }}
                  >
                    Close
                  </Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}

      {isCrewRosterGeneratorOpen ? (
        <>
          <div className="app-modal-backdrop" style={{ zIndex: 1066 }} />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1067 }}>
            <div className="modal-dialog modal-lg modal-dialog-centered">
              <div className="modal-content">
                <div className="modal-header">
                  <h5 className="modal-title fw-semibold">Roster Generator</h5>
                  <button
                    type="button"
                    className="btn-close"
                    aria-label="Close"
                    onClick={() => setIsCrewRosterGeneratorOpen(false)}
                  />
                </div>
                <div className="modal-body d-flex flex-column gap-3">
                  <ul className="nav nav-tabs">
                    <li className="nav-item">
                      <button
                        type="button"
                        className={"nav-link" + (crewRosterGeneratorTab === "setup" ? " active" : "")}
                        onClick={() => setCrewRosterGeneratorTab("setup")}
                      >
                        Setup
                      </button>
                    </li>
                    <li className="nav-item">
                      <button
                        type="button"
                        className={"nav-link" + (crewRosterGeneratorTab === "rules" ? " active" : "")}
                        onClick={() => setCrewRosterGeneratorTab("rules")}
                      >
                        Roster Rule Set
                      </button>
                    </li>
                  </ul>

                  {crewRosterGeneratorTab === "setup" ? (
                    <div className="row g-2">
                      <div className="col-6 col-md-3">
                        <label className="form-label small fw-semibold">From Date</label>
                        <input
                          type="date"
                          className="form-control form-control-sm"
                          value={crewRosterGeneratorForm.fromDate}
                          onChange={(e) =>
                            setCrewRosterGeneratorForm((p) => ({ ...p, fromDate: e.target.value }))
                          }
                        />
                      </div>
                      <div className="col-6 col-md-3">
                        <label className="form-label small fw-semibold">To Date</label>
                        <input
                          type="date"
                          className="form-control form-control-sm"
                          value={crewRosterGeneratorForm.toDate}
                          onChange={(e) =>
                            setCrewRosterGeneratorForm((p) => ({ ...p, toDate: e.target.value }))
                          }
                        />
                      </div>
                      <div className="col-12 col-md-4">
                        <label className="form-label small fw-semibold">Crew</label>
                        <select
                          className="form-select form-select-sm"
                          value={crewRosterGeneratorForm.crewCode}
                          onChange={(e) =>
                            setCrewRosterGeneratorForm((p) => ({ ...p, crewCode: e.target.value }))
                          }
                        >
                          <option value="">All Crew</option>
                          {crewWithReleaseDate.map((c) => (
                            <option key={c.id} value={normalizeCrewCode(c.crewCode || "")}>
                              {normalizeCrewCode(c.crewCode || "")} - {c.crewName || ""}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div className="col-12 col-md-2">
                        <label className="form-label small fw-semibold">Default Duty Code</label>
                        <select
                          className="form-select form-select-sm"
                          value={crewRosterGeneratorForm.defaultDutyCode}
                          onChange={(e) =>
                            setCrewRosterGeneratorForm((p) => ({
                              ...p,
                              defaultDutyCode: e.target.value,
                              dutyCodeSequence:
                                !p.dutyCodeSequence || p.dutyCodeSequence === p.defaultDutyCode
                                  ? e.target.value
                                  : p.dutyCodeSequence,
                            }))
                          }
                        >
                          {crewDutyCodeConfigs
                            .filter((row) => {
                              const code = (row.code || "").trim().toUpperCase()
                              if (!code) return false
                              const dutyType = (row.dutyType || "").trim()
                              const map =
                                crewDutyTypeMap.get(dutyType.toLowerCase()) ||
                                row.rosterBarMap ||
                                getCrewDutyRosterBarMap(dutyType)
                              return map !== "REST"
                            })
                            .map((row) => {
                              const code = (row.code || "").trim().toUpperCase()
                              return (
                                <option key={`generator-duty-${row.id}`} value={code}>
                                  {code}
                                </option>
                              )
                            })}
                        </select>
                      </div>
                      <div className="col-12">
                        <div className="small text-muted">
                          Roster Day uses the selected default duty code from Duty Code Manager. Fixed rest days control the
                          48h recovery rest placement.
                        </div>
                      </div>
                      <div className="col-12">
                        <label className="d-flex align-items-center gap-2 small fw-semibold mb-1">
                          <input
                            type="checkbox"
                            className="form-check-input mt-0"
                            checked={crewRosterGeneratorForm.useFixedRestWeekday}
                            onChange={(e) =>
                              setCrewRosterGeneratorForm((p) => ({
                                ...p,
                                useFixedRestWeekday: e.target.checked,
                              }))
                            }
                          />
                          Fix 48h Rest Period To Specific Weekday (Individual Crew)
                        </label>
                        {crewRosterGeneratorForm.useFixedRestWeekday ? (
                          <div className="d-flex flex-wrap gap-3 small">
                            {[
                              "Monday",
                              "Tuesday",
                              "Wednesday",
                              "Thursday",
                              "Friday",
                              "Saturday",
                              "Sunday",
                            ].map((day) => (
                              <label key={day} className="d-flex align-items-center gap-1">
                                <input
                                  type="checkbox"
                                  className="form-check-input mt-0"
                                  checked={crewRosterGeneratorForm.fixedRestWeekdays.includes(day)}
                                  onChange={(e) =>
                                    setCrewRosterGeneratorForm((p) => ({
                                      ...p,
                                      fixedRestWeekdays: e.target.checked
                                        ? Array.from(new Set([...p.fixedRestWeekdays, day]))
                                        : p.fixedRestWeekdays.filter((d) => d !== day),
                                    }))
                                  }
                                />
                                <span>{day}</span>
                              </label>
                            ))}
                          </div>
                        ) : null}
                      </div>
                    </div>
                  ) : (
                    <div className="row g-2">
                      <div className="col-12 col-md-6">
                        <label className="form-label small fw-semibold">
                          Maximum Time Between Two Recovery Rest Period
                        </label>
                        <input
                          className="form-control form-control-sm"
                          value={crewRosterGeneratorForm.maxRecoveryRestGap}
                          onChange={(e) =>
                            setCrewRosterGeneratorForm((p) => ({
                              ...p,
                              maxRecoveryRestGap: e.target.value,
                            }))
                          }
                        />
                      </div>
                      <div className="col-12 col-md-6">
                        <label className="form-label small fw-semibold">Maximum Duty Time / Day</label>
                        <input
                          className="form-control form-control-sm"
                          value={crewRosterGeneratorForm.maxDutyTimePerDay}
                          onChange={(e) =>
                            setCrewRosterGeneratorForm((p) => ({
                              ...p,
                              maxDutyTimePerDay: e.target.value,
                            }))
                          }
                        />
                        <div className="form-text small">Example: 13 hours. Multiple duty codes stop when this limit is reached.</div>
                      </div>
                      <div className="col-12 col-md-6">
                        <label className="form-label small fw-semibold">Rest Period</label>
                        <input
                          className="form-control form-control-sm"
                          value={crewRosterGeneratorForm.restPeriod}
                          onChange={(e) =>
                            setCrewRosterGeneratorForm((p) => ({ ...p, restPeriod: e.target.value }))
                          }
                        />
                      </div>
                      <div className="col-12">
                        <label className="form-label small fw-semibold">Duty Code Sequence</label>
                        <input
                          className="form-control form-control-sm"
                          placeholder="CB,FD,DP"
                          value={crewRosterGeneratorForm.dutyCodeSequence}
                          onChange={(e) =>
                            setCrewRosterGeneratorForm((p) => ({
                              ...p,
                              dutyCodeSequence: e.target.value,
                            }))
                          }
                        />
                        <div className="form-text small">
                          Use comma-separated Duty Code Manager codes mapped to Roster Day. Each code keeps its default start/end time.
                        </div>
                      </div>
                      <div className="col-12 col-md-6">
                        <label className="form-label small fw-semibold">Local Night</label>
                        <input
                          className="form-control form-control-sm"
                          value={crewRosterGeneratorForm.localNight}
                          onChange={(e) =>
                            setCrewRosterGeneratorForm((p) => ({ ...p, localNight: e.target.value }))
                          }
                        />
                      </div>
                      <div className="col-12 col-md-6">
                        <label className="form-label small fw-semibold">Local Day</label>
                        <input
                          className="form-control form-control-sm"
                          value={crewRosterGeneratorForm.localDay}
                          onChange={(e) =>
                            setCrewRosterGeneratorForm((p) => ({ ...p, localDay: e.target.value }))
                          }
                        />
                      </div>
                    </div>
                  )}
                </div>
                <div className="modal-footer">
                  <Button variant="destructive" onClick={clearCrewRosterBars}>
                    Clear Roster
                  </Button>
                  <Button variant="outline" onClick={() => setIsCrewRosterGeneratorOpen(false)}>
                    Cancel
                  </Button>
                  <Button onClick={generateCrewRosterBars}>Generate Roster Bars</Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}

      {pendingCrewDutyDrop.open ? (
        <>
          <div className="app-modal-backdrop" style={{ zIndex: 1068 }} />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1069 }}>
            <div className="modal-dialog modal-lg modal-dialog-centered">
              <div className="modal-content">
                <div className="modal-header">
                  <div>
                    <h5 className="modal-title fw-semibold">Assign Duty / Rest Code</h5>
                    <div className="small text-muted">
                      {pendingCrewDutyDrop.crewCode} {pendingCrewDutyDrop.crewName ? `- ${pendingCrewDutyDrop.crewName}` : ""}
                    </div>
                  </div>
                  <button
                    type="button"
                    className="btn-close"
                    aria-label="Close"
                    onClick={() => setPendingCrewDutyDrop((p) => ({ ...p, open: false }))}
                  />
                </div>
                <div className="modal-body">
                  <div className="row g-2 mb-3">
                    <div className="col-6 col-md-3">
                      <div className="small text-muted">Code</div>
                      <div className="fw-semibold">{pendingCrewDutyDrop.code}</div>
                    </div>
                    <div className="col-6 col-md-3">
                      <div className="small text-muted">Map</div>
                      <div className="fw-semibold">
                        {pendingCrewDutyDrop.rosterBarMap === "REST" ? "Rest Day" : "Roster Day"}
                      </div>
                    </div>
                    <div className="col-6 col-md-3">
                      <label className="form-label small fw-semibold">Start</label>
                      <div className="d-flex gap-1">
                        <input
                          type="date"
                          className="form-control form-control-sm"
                          value={pendingCrewDutyDrop.startDate}
                          onChange={(e) => setPendingCrewDutyDrop((p) => ({ ...p, startDate: e.target.value }))}
                        />
                        <input
                          type="time"
                          className="form-control form-control-sm"
                          value={pendingCrewDutyDrop.startTime}
                          onChange={(e) => setPendingCrewDutyDrop((p) => ({ ...p, startTime: e.target.value }))}
                        />
                      </div>
                    </div>
                    <div className="col-6 col-md-3">
                      <label className="form-label small fw-semibold">End</label>
                      <div className="d-flex gap-1">
                        <input
                          type="date"
                          className="form-control form-control-sm"
                          value={pendingCrewDutyDrop.endDate}
                          onChange={(e) => setPendingCrewDutyDrop((p) => ({ ...p, endDate: e.target.value }))}
                        />
                        <input
                          type="time"
                          className="form-control form-control-sm"
                          value={pendingCrewDutyDrop.endTime}
                          onChange={(e) => setPendingCrewDutyDrop((p) => ({ ...p, endTime: e.target.value }))}
                        />
                      </div>
                    </div>
                  </div>
                  {pendingCrewDutyDrop.conflicts.length ? (
                    <div className="alert alert-warning py-2">
                      <div className="fw-semibold mb-2">Conflict with existing assigned code</div>
                      <div className="table-responsive mb-2">
                        <table className="table table-sm mb-0">
                          <thead>
                            <tr>
                              <th>Type</th>
                              <th>Code</th>
                              <th>Start</th>
                              <th>End</th>
                            </tr>
                          </thead>
                          <tbody>
                            {pendingCrewDutyDrop.conflicts.map((c) => (
                              <tr key={c.id}>
                                <td>{c.type}</td>
                                <td>{c.code || "-"}</td>
                                <td>{formatYmdToDdMmYy(c.start)} {c.startTime}</td>
                                <td>{formatYmdToDdMmYy(c.end)} {c.endTime}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <div className="d-flex flex-wrap gap-3 small">
                        <label className="d-flex align-items-center gap-2">
                          <input
                            type="radio"
                            name="crew-duty-drop-mode"
                            checked={pendingCrewDutyDrop.mode === "after"}
                            onChange={() => setPendingCrewDutyDrop((p) => ({ ...p, mode: "after" }))}
                          />
                          Assign after existing duty end
                        </label>
                        <label className="d-flex align-items-center gap-2">
                          <input
                            type="radio"
                            name="crew-duty-drop-mode"
                            checked={pendingCrewDutyDrop.mode === "override"}
                            onChange={() => setPendingCrewDutyDrop((p) => ({ ...p, mode: "override" }))}
                          />
                          Override conflicting assigned code
                        </label>
                      </div>
                    </div>
                  ) : (
                    <div className="alert alert-success py-2 mb-0">No assigned duty/rest conflict found.</div>
                  )}
                </div>
                <div className="modal-footer">
                  <Button variant="outline" onClick={() => setPendingCrewDutyDrop((p) => ({ ...p, open: false }))}>
                    Cancel
                  </Button>
                  <Button onClick={confirmCrewDutyDrop}>Assign Code</Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}

      {pendingActualLeaveDrop.open ? (
        <>
          <div className="app-modal-backdrop" style={{ zIndex: 1068 }} />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1069 }}>
            <div className="modal-dialog modal-lg modal-dialog-centered">
              <div className="modal-content">
                <div className="modal-header">
                  <div>
                    <h5 className="modal-title fw-semibold">Assign Actual Leave From Roster Drop</h5>
                    <div className="small text-muted">
                      {pendingActualLeaveDrop.crewCode} {pendingActualLeaveDrop.crewName ? `- ${pendingActualLeaveDrop.crewName}` : ""}
                      {pendingActualLeaveDrop.cycleNumber ? ` | Cycle ${pendingActualLeaveDrop.cycleNumber}` : ""}
                    </div>
                  </div>
                  <button
                    type="button"
                    className="btn-close"
                    aria-label="Close"
                    onClick={() => setPendingActualLeaveDrop((p) => ({ ...p, open: false }))}
                  />
                </div>
                <div className="modal-body">
                  <div className="row g-2 align-items-end mb-3">
                    <div className="col-6 col-md-4">
                      <label className="form-label small fw-semibold">Actual Leave Start</label>
                      <input
                        type="date"
                        className="form-control form-control-sm"
                        value={pendingActualLeaveDrop.startDate}
                        onChange={(e) => {
                          const startDate = e.target.value
                          const plannedDays = Math.max(pendingActualLeaveDrop.plannedDays || 1, 1)
                          const endDate = startDate ? addDaysToYmd(startDate, plannedDays - 1) : ""
                          setPendingActualLeaveDrop((p) => ({
                            ...p,
                            startDate,
                            endDate,
                            conflicts: findActualLeaveDropConflicts({
                              crewCode: p.crewCode,
                              startDate,
                              endDate,
                              sourceBlockId: p.sourceBlockId,
                            }),
                            override: false,
                          }))
                        }}
                      />
                    </div>
                    <div className="col-6 col-md-4">
                      <label className="form-label small fw-semibold">Actual Leave End</label>
                      <input
                        type="date"
                        className="form-control form-control-sm"
                        value={pendingActualLeaveDrop.endDate}
                        onChange={(e) => {
                          const endDate = e.target.value
                          setPendingActualLeaveDrop((p) => ({
                            ...p,
                            endDate,
                            plannedDays:
                              p.startDate && endDate && endDate >= p.startDate
                                ? getDateRangeInclusive(p.startDate, endDate).length
                                : p.plannedDays,
                            conflicts: findActualLeaveDropConflicts({
                              crewCode: p.crewCode,
                              startDate: p.startDate,
                              endDate,
                              sourceBlockId: p.sourceBlockId,
                            }),
                            override: false,
                          }))
                        }}
                      />
                    </div>
                    <div className="col-12 col-md-4">
                      <div className="border rounded p-2 bg-body-tertiary">
                        <div className="small text-muted">Leave Days</div>
                        <div className="fw-semibold">{pendingActualLeaveDrop.plannedDays || 0} day(s)</div>
                      </div>
                    </div>
                  </div>

                  {pendingActualLeaveDrop.conflicts.length > 0 ? (
                    <div className="alert alert-warning py-2">
                      <div className="fw-semibold mb-1">Conflict found in selected range</div>
                      <div className="table-responsive">
                        <table className="table table-sm mb-2 align-middle">
                          <thead>
                            <tr>
                              <th>Type</th>
                              <th>Code</th>
                              <th>Start</th>
                              <th>End</th>
                            </tr>
                          </thead>
                          <tbody>
                            {pendingActualLeaveDrop.conflicts.slice(0, 12).map((conflict, idx) => (
                              <tr key={`${conflict.type}-${conflict.start}-${idx}`}>
                                <td>{conflict.type}</td>
                                <td>{conflict.code || "-"}</td>
                                <td>
                                  {formatYmdToDdMmYy(conflict.start)} {conflict.startTime}
                                </td>
                                <td>
                                  {formatYmdToDdMmYy(conflict.end)} {conflict.endTime}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <label className="d-flex align-items-center gap-2 small fw-semibold">
                        <input
                          type="checkbox"
                          className="form-check-input mt-0"
                          checked={pendingActualLeaveDrop.override}
                          onChange={(e) =>
                            setPendingActualLeaveDrop((p) => ({ ...p, override: e.target.checked }))
                          }
                        />
                        Override conflict and save Actual Leave
                      </label>
                    </div>
                  ) : (
                    <div className="alert alert-success py-2 mb-0">
                      No roster conflict found for this Actual Leave range.
                    </div>
                  )}
                </div>
                <div className="modal-footer">
                  <Button variant="outline" onClick={() => setPendingActualLeaveDrop((p) => ({ ...p, open: false }))}>
                    Cancel
                  </Button>
                  <Button
                    onClick={confirmActualLeaveDrop}
                    disabled={pendingActualLeaveDrop.conflicts.length > 0 && !pendingActualLeaveDrop.override}
                  >
                    Save Actual Leave
                  </Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}

      {rosterBarDetailModal.open ? (
        <>
          <div className="app-modal-backdrop" style={{ zIndex: 1064 }} />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1065 }}>
            <div className="modal-dialog modal-xl modal-dialog-centered">
              <div className="modal-content">
                <div className="modal-header">
                  <h5 className="modal-title fw-semibold">Roster Bar Details</h5>
                  <button
                    type="button"
                    className="btn-close"
                    aria-label="Close"
                    onClick={() => setRosterBarDetailModal((p) => ({ ...p, open: false }))}
                  />
                </div>
                <div className="modal-body">
                  <div className="small">
                    <div><strong>Crew:</strong> {rosterBarDetailModal.crewCode} {rosterBarDetailModal.crewName ? `| ${rosterBarDetailModal.crewName}` : ""}</div>
                    <div><strong>Type:</strong> {rosterBarDetailModal.rosterType}</div>
                    <div><strong>Start:</strong> {rosterBarDetailModal.startDate} {rosterBarDetailModal.startTime}</div>
                    <div><strong>End:</strong> {rosterBarDetailModal.endDate} {rosterBarDetailModal.endTime}</div>
                    <div><strong>Duration:</strong> {rosterBarDetailModal.durationHours}h</div>
                    <div>
                      <strong>Duty Code Assigned:</strong>{" "}
                      {rosterBarDetailModal.dutyAssignments.length
                        ? rosterBarDetailModal.dutyAssignments.map((d) => d.code).filter(Boolean).join(", ")
                        : "-"}
                    </div>
                  </div>
                  <div className="mt-3">
                    <div className="small fw-semibold mb-1">Assigned Duty Codes</div>
                    <div style={{ maxHeight: 180, overflow: "auto", border: "1px solid #e5e7eb", borderRadius: 6 }}>
                      <table className="table table-sm mb-0">
                        <thead>
                          <tr>
                            <th>Code</th>
                            <th>Duty Type</th>
                            <th style={{ minWidth: 170 }}>Start</th>
                            <th style={{ minWidth: 170 }}>End</th>
                          </tr>
                        </thead>
                        <tbody>
                          {rosterBarDetailModal.dutyAssignments.length ? (
                            rosterBarDetailModal.dutyAssignments.map((d, idx) => (
                              <tr key={d.id || `${d.code}-${d.startDate}-${idx}`}>
                                <td className="fw-semibold">{d.code || "-"}</td>
                                <td>{d.dutyType || "-"}</td>
                                <td>
                                  <div className="d-flex gap-1">
                                    <input
                                      type="date"
                                      className="form-control form-control-sm"
                                      value={d.startDate}
                                      onChange={(e) =>
                                        updateRosterDetailDutyAssignment(idx, "startDate", e.target.value)
                                      }
                                    />
                                    <input
                                      type="time"
                                      className="form-control form-control-sm"
                                      value={d.startTime}
                                      onChange={(e) =>
                                        updateRosterDetailDutyAssignment(idx, "startTime", e.target.value)
                                      }
                                    />
                                  </div>
                                </td>
                                <td>
                                  <div className="d-flex gap-1">
                                    <input
                                      type="date"
                                      className="form-control form-control-sm"
                                      value={d.endDate}
                                      onChange={(e) =>
                                        updateRosterDetailDutyAssignment(idx, "endDate", e.target.value)
                                      }
                                    />
                                    <input
                                      type="time"
                                      className="form-control form-control-sm"
                                      value={d.endTime}
                                      onChange={(e) =>
                                        updateRosterDetailDutyAssignment(idx, "endTime", e.target.value)
                                      }
                                    />
                                  </div>
                                </td>
                              </tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={4} className="text-muted">No duty code assigned in this selected range.</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                  <div className="mt-3">
                    <div className="small fw-semibold mb-1">By Day Timing</div>
                    <div style={{ maxHeight: 420, overflow: "auto", border: "1px solid #e5e7eb", borderRadius: 6 }}>
                      <table className="table table-sm mb-0">
                        <thead>
                          <tr>
                            <th>Day</th>
                            <th>Type</th>
                            <th>Rest From</th>
                            <th>Rest To</th>
                            <th>Roster Day From</th>
                            <th>Roster Day To</th>
                          </tr>
                        </thead>
                        <tbody>
                          {rosterBarDetailModal.daySegments.length ? (
                            rosterBarDetailModal.daySegments.map((seg, idx) => (
                              <tr key={`${seg.day}-${idx}`}>
                                <td>{seg.day}</td>
                                <td>{seg.type}</td>
                                <td>{seg.restFrom || "-"}</td>
                                <td>{seg.restTo || "-"}</td>
                                <td>{seg.rosterFrom || "-"}</td>
                                <td>{seg.rosterTo || "-"}</td>
                              </tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={6} className="text-muted">No day segments.</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
                <div className="modal-footer">
                  <Button
                    variant="outline"
                    onClick={() => setRosterBarDetailModal((p) => ({ ...p, open: false }))}
                  >
                    Close
                  </Button>
                  <Button
                    onClick={saveRosterDetailDutyAssignments}
                    disabled={rosterBarDetailModal.dutyAssignments.length === 0}
                  >
                    Save Duty Times
                  </Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}

      {mustChangePassword ? (
        <>
          <div className="app-modal-backdrop" />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1060 }}>
            <div className="modal-dialog modal-dialog-centered">
              <div className="modal-content">
                <div className="modal-header">
                  <h5 className="modal-title fw-semibold">Change Password</h5>
                </div>
                <div className="modal-body">
                  <p className="small text-muted mb-3">
                    First login detected. Please change your password now.
                  </p>
                  <div className="mb-3">
                    <label className="form-label small fw-semibold">New Password</label>
                    <input
                      type="password"
                      className="form-control"
                      value={newPassword}
                      onChange={(e) => setNewPassword(e.target.value)}
                    />
                  </div>
                  <div className="mb-1">
                    <label className="form-label small fw-semibold">Confirm New Password</label>
                    <input
                      type="password"
                      className="form-control"
                      value={confirmNewPassword}
                      onChange={(e) => setConfirmNewPassword(e.target.value)}
                    />
                  </div>
                  {passwordChangeError ? (
                    <div className="alert alert-danger py-2 small mt-2 mb-0">{passwordChangeError}</div>
                  ) : null}
                </div>
                <div className="modal-footer">
                  <Button onClick={changePasswordNow}>Save Password</Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}

      {isInactivePopupOpen ? (
        <>
          <div className="app-modal-backdrop" style={{ zIndex: 1060 }} />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1065 }}>
            <div className="modal-dialog modal-dialog-centered">
              <div className="modal-content">
                <div className="modal-header">
                  <div className="d-flex align-items-start gap-2">
                    <span
                      className="d-inline-flex align-items-center justify-content-center rounded-circle bg-warning-subtle text-warning-emphasis"
                      style={{ width: 36, height: 36, flexShrink: 0 }}
                    >
                      !
                    </span>
                    <div>
                      <h5 className="modal-title fw-semibold mb-0">Set Staff Inactive</h5>
                      <p className="small text-muted mb-0 mt-1">
                        Please provide inactive date and reason.
                      </p>
                    </div>
                  </div>
                  <button
                    type="button"
                    className="btn-close"
                    aria-label="Close"
                    onClick={closeInactivePopup}
                  />
                </div>
                <div className="modal-body">
                  <div className="mb-3">
                    <label className="form-label small fw-semibold">
                      Inactive Date <span className="text-danger">*</span>
                    </label>
                    <input
                      type="date"
                      className="form-control"
                      value={inactiveDateDraft}
                      onChange={(e) => setInactiveDateDraft(e.target.value)}
                    />
                  </div>
                  <div className="mb-0">
                    <label className="form-label small fw-semibold">
                      Inactive Reason <span className="text-danger">*</span>
                    </label>
                    <textarea
                      className="form-control"
                      rows={3}
                      value={inactiveReasonDraft}
                      onChange={(e) => setInactiveReasonDraft(e.target.value)}
                      placeholder="Enter reason"
                    />
                  </div>
                </div>
                <div className="modal-footer">
                  <Button variant="outline" onClick={closeInactivePopup}>
                    Cancel
                  </Button>
                  <Button onClick={confirmInactive}>Save Inactive</Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}

      {isRosterIOModalOpen ? (
        <>
          <div className="app-modal-backdrop" />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1055 }}>
            <div className="modal-dialog modal-md modal-dialog-centered">
              <div className="modal-content">
                <div className="modal-header">
                  <div>
                    <h5 className="modal-title fw-semibold">Roster Import / Export</h5>
                    <p className="small text-muted mb-0">
                      Download roster matrix by period or upload updated matrix file.
                    </p>
                  </div>
                  <button
                    type="button"
                    className="btn-close"
                    aria-label="Close"
                    onClick={() => setIsRosterIOModalOpen(false)}
                  />
                </div>
                <div className="modal-body">
                  <div className="row g-3">
                    <div className="col-12 col-md-6">
                      <label className="form-label small fw-semibold">Export From</label>
                      <input
                        type="date"
                        className="form-control"
                        value={rosterExportRange.fromDate}
                        onChange={(e) =>
                          setRosterExportRange((prev) => ({ ...prev, fromDate: e.target.value }))
                        }
                      />
                    </div>
                    <div className="col-12 col-md-6">
                      <label className="form-label small fw-semibold">Export To</label>
                      <input
                        type="date"
                        className="form-control"
                        value={rosterExportRange.toDate}
                        onChange={(e) =>
                          setRosterExportRange((prev) => ({ ...prev, toDate: e.target.value }))
                        }
                      />
                    </div>
                    <div className="col-12">
                      <Button
                        variant="outline"
                        onClick={downloadRosterCsv}
                        className="w-100"
                        disabled={!canRosterDownload}
                      >
                        Download Roster Matrix (Excel CSV)
                      </Button>
                    </div>
                    <div className="col-12">
                      <Button
                        variant="outline"
                        onClick={openRosterPrintView}
                        className="w-100"
                        disabled={!canRosterDownload}
                      >
                        Open Print View (Landscape)
                      </Button>
                    </div>
                    <div className="col-12">
                      <label className="form-check mb-0">
                        <input
                          type="checkbox"
                          className="form-check-input"
                          checked={rosterUploadAdminMode}
                          onChange={(e) => setRosterUploadAdminMode(e.target.checked)}
                          disabled={!canRosterUpload}
                        />
                        <span className="form-check-label small ms-1">
                          Administrator upload (allow past dates)
                        </span>
                      </label>
                    </div>
                    <div className="col-12">
                      <Button
                        variant="outline"
                        className="w-100"
                        onClick={() => rosterUploadInputRef.current?.click()}
                        disabled={!canRosterUpload}
                      >
                        Upload Roster Matrix (Excel CSV)
                      </Button>
                        <input
                          ref={rosterUploadInputRef}
                          type="file"
                          accept=".csv,text/csv"
                          className="d-none"
                          onChange={async (e) => {
                            const inputEl = e.currentTarget
                            const file = e.target.files?.[0]
                            if (!file) return
                            await importRosterCsv(file)
                            if (inputEl) inputEl.value = ""
                            setIsRosterIOModalOpen(false)
                          }}
                        />
                    </div>
                  </div>
                </div>
                <div className="modal-footer">
                  <Button variant="outline" onClick={() => setIsRosterIOModalOpen(false)}>
                    Close
                  </Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}

      {isRosterGeneratorOpen ? (
        <>
          <div className="app-modal-backdrop" />
          <div className="modal d-block" tabIndex={-1} role="dialog" style={{ zIndex: 1055 }}>
            <div className="modal-dialog modal-lg modal-dialog-centered modal-dialog-scrollable">
              <div className="modal-content">
                <div className="modal-header">
                  <div>
                    <h5 className="modal-title fw-semibold">Roster Generator</h5>
                    <p className="small text-muted mb-0">
                      Configure roster settings, then generate or clear within selected scope.
                    </p>
                  </div>
                  <button
                    type="button"
                    className="btn-close"
                    aria-label="Close"
                    onClick={() => setIsRosterGeneratorOpen(false)}
                  />
                </div>
                <div className="modal-body">
                  <div className="row g-3">
                    <div className="col-12 col-md-4">
                      <label className="form-label small fw-semibold">From Date</label>
                      <input
                        type="date"
                        className="form-control"
                        value={rosterSettings.fromDate}
                        min={toYmd(new Date())}
                        onChange={(e) =>
                          setRosterSettings((p) => ({ ...p, fromDate: e.target.value }))
                        }
                      />
                    </div>
                    <div className="col-12 col-md-4">
                      <label className="form-label small fw-semibold">To Date</label>
                      <input
                        type="date"
                        className="form-control"
                        value={rosterSettings.toDate}
                        min={toYmd(new Date())}
                        onChange={(e) =>
                          setRosterSettings((p) => ({ ...p, toDate: e.target.value }))
                        }
                      />
                    </div>
                    <div className="col-12 col-md-4">
                      <label className="form-label small fw-semibold">Staff Selection</label>
                      <select
                        className="form-select"
                        value={rosterSettings.staffId}
                        onChange={(e) =>
                          setRosterSettings((p) => ({ ...p, staffId: e.target.value }))
                        }
                      >
                        <option value="ALL">All Active Staff</option>
                        {activeStaff
                          .slice()
                          .sort((a, b) => {
                            const noA = a.staffNo || ""
                            const noB = b.staffNo || ""
                            return (
                              noA.localeCompare(noB) ||
                              (a.fullName || "").localeCompare(b.fullName || "")
                            )
                          })
                          .map((staff) => (
                            <option key={getStaffSelectionToken(staff)} value={getStaffSelectionToken(staff)}>
                              {staff.staffNo} - {staff.fullName}
                            </option>
                          ))}
                      </select>
                    </div>
                    <div className="col-12 col-md-4">
                      <label className="form-label small fw-semibold">Set Off</label>
                      <select
                        className="form-select"
                        value={rosterSettings.offType}
                        onChange={(e) =>
                          setRosterSettings((p) => ({ ...p, offType: e.target.value }))
                        }
                      >
                        <option value="Fixed">Fixed</option>
                        <option value="Rotational">Rotational</option>
                      </select>
                    </div>
                    {rosterSettings.offType === "Fixed" ? (
                      <div className="col-12 col-md-8">
                        <label className="form-label small fw-semibold">Fixed Weekday Off</label>
                        <div className="border rounded p-2 d-flex flex-wrap gap-2">
                          {[
                            "Monday",
                            "Tuesday",
                            "Wednesday",
                            "Thursday",
                            "Friday",
                            "Saturday",
                            "Sunday",
                          ].map((day) => {
                            const selected = splitMultiValue(rosterSettings.fixedOffWeekdays)
                            const checked = selected.includes(day)
                            return (
                              <label
                                key={day}
                                className={
                                  "btn btn-sm " +
                                  (checked ? "btn-primary" : "btn-outline-secondary") +
                                  " d-inline-flex align-items-center gap-2 mb-0"
                                }
                              >
                                <input
                                  type="checkbox"
                                  className="form-check-input mt-0"
                                  checked={checked}
                                  onChange={() => {
                                    const next = checked
                                      ? selected.filter((x) => x !== day)
                                      : [...selected, day]
                                    setRosterSettings((p) => ({
                                      ...p,
                                      fixedOffWeekdays: next.join(", "),
                                    }))
                                  }}
                                  style={{ marginLeft: 0 }}
                                />
                                <span>{day}</span>
                              </label>
                            )
                          })}
                        </div>
                      </div>
                    ) : (
                      <div className="col-12 col-md-8">
                        <div className="row g-3">
                          <div className="col-12 col-md-6">
                            <label className="form-label small fw-semibold">Work Day Duration (Hours)</label>
                            <input
                              type="number"
                              className="form-control"
                              value={rosterSettings.maxShiftHours}
                              onChange={(e) =>
                                setRosterSettings((p) => ({ ...p, maxShiftHours: e.target.value }))
                              }
                            />
                          </div>
                          <div className="col-12 col-md-6">
                            <label className="form-label small fw-semibold">Rotational Rest Period (Hours)</label>
                            <input
                              type="number"
                              className="form-control"
                              value={rosterSettings.rotationalRestHours}
                              onChange={(e) =>
                                setRosterSettings((p) => ({ ...p, rotationalRestHours: e.target.value }))
                              }
                            />
                          </div>
                        </div>
                      </div>
                    )}
                    <div className="col-12">
                      <label className="form-label small fw-semibold">Shift Code (By Assigned Levels)</label>
                      <select
                        className="form-select"
                        value={rosterSettings.shiftCode}
                        onChange={(e) =>
                          setRosterSettings((p) => ({ ...p, shiftCode: e.target.value }))
                        }
                      >
                        <option value="">Auto (Best Priority)</option>
                        {rosterEligibleShiftCodes.map((code) => (
                          <option key={code} value={code}>{code}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                </div>
                <div className="modal-footer">
                  <Button variant="outline" onClick={clearGeneratedRoster}>
                    Clear Generated Roster
                  </Button>
                  <Button
                    onClick={() => {
                      generateRoster()
                      setIsRosterGeneratorOpen(false)
                    }}
                  >
                    Generate
                  </Button>
                </div>
              </div>
            </div>
          </div>
        </>
      ) : null}
    </div>
  )
}
