/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  CalendarDays,
  Check,
  Bell,
  FileText,
  ArrowRight,
  Clock,
  X,
  Stethoscope,
  RefreshCcw,
  CheckCircle2,
  AlertCircle,
  Palmtree,
  Download,
  Upload,
  Phone,
  PhoneCall,
  Shield,
  Lock,
  Unlock,
  User,
  Settings,
  Search,
  Filter,
  TrendingUp,
  Sparkles,
  Trash2,
  CalendarRange,
  CalendarPlus
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";

const GROUP_1 = ["สุนิสสา", "อาทิตยา", "รวีวรรณ", "จิราภรณ์", "สุนิษา (ผจก.)"];
const GROUP_2 = [
  "อุษา",
  "สิริพร",
  "ไพโรจน์",
  "ต่วนมารีนา",
  "วรพงศ์",
  "ศศลักษณ์",
  "ธนัชพร",
  "พรวิมล",
  "วิชชุตา"
];
const ALL_STAFF = [...GROUP_1, ...GROUP_2];
const DISPLAY_ORDER = [...GROUP_2, ...GROUP_1];
const FIRE_CODES_REST = ["A", "B", "C", "D"];

interface Holiday {
  date: string;
  name: string;
  type: "public" | "company";
}

interface DaySchedule {
  date: string;
  workingStaff: string[];
  fireCodes: Record<string, string>;
  vacationStaff: string[];
  docInCharge: string | null;
  phone3551?: string | null;
  phone3552?: string | null;
}

interface ToastMessage {
  id: number;
  message: string;
  inCharge: boolean;
  dateStr: string;
}

interface LogEntry {
  id: number;
  date: string;
  timestamp: string;
  personOut: string;
  personIn: string;
  inChargeTriggered: boolean;
}

interface ShiftRequest {
  id: string;
  requester: string;
  type: "swap" | "cover" | "leave" | "off" | "work";
  date: string;
  targetStaff?: string;
  targetDate?: string;
  status: "pending" | "approved" | "rejected";
  note?: string;
  createdAt: string;
}

interface NotificationItem {
  id: string;
  title: string;
  message: string;
  timestamp: string;
  isRead: boolean;
  type: "info" | "success" | "warning";
}

function shuffleArray<T>(array: T[]): T[] {
  const newArray = [...array];
  for (let i = newArray.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
  }
  return newArray;
}

function assignFireCodes(workingStaff: string[]) {
  if (workingStaff.length === 0) return {};
  const shuffledStaff = shuffleArray(workingStaff);
  const assignments: Record<string, string> = {};
  
  // 1 คนได้ I 
  assignments[shuffledStaff[0]] = "I";
  
  // ที่เหลือได้ A, B, C, D สุ่มซ้ำได้
  for (let i = 1; i < shuffledStaff.length; i++) {
    assignments[shuffledStaff[i]] = FIRE_CODES_REST[Math.floor(Math.random() * FIRE_CODES_REST.length)];
  }
  return assignments;
}

function rebalancePhones(schedule: Record<string, DaySchedule>): Record<string, DaySchedule> {
  const newSchedule: Record<string, DaySchedule> = JSON.parse(JSON.stringify(schedule));
  const counts: Record<string, number> = {};
  GROUP_2.forEach(s => counts[s] = 0);

  const dates = Object.keys(newSchedule).sort();
  dates.forEach(date => {
     newSchedule[date].phone3551 = null;
     newSchedule[date].phone3552 = null;
  });

  for (const date of dates) {
    const day = newSchedule[date];
    const available = day.workingStaff.filter(s => GROUP_2.includes(s));
    
    if (available.length >= 2) {
      available.sort((a, b) => {
        if (counts[a] !== counts[b]) return counts[a] - counts[b];
        return a.localeCompare(b);
      });

      const picked = [available[0], available[1]];
      counts[picked[0]]++;
      counts[picked[1]]++;

      let p3551 = picked[0];
      let p3552 = picked[1];

      const prefers3551 = picked.find(p => p === "ต่วนมารีนา");
      const prefers3552 = picked.find(p => p === "ศศลักษณ์" || p === "ธนัชพร");

      if (prefers3551 && prefers3552) {
        p3551 = prefers3551;
        p3552 = prefers3552;
      } else if (prefers3551) {
        p3551 = prefers3551;
        p3552 = picked.find(p => p !== prefers3551)!;
      } else if (prefers3552) {
        p3552 = prefers3552;
        p3551 = picked.find(p => p !== prefers3552)!;
      }
      
      day.phone3551 = p3551;
      day.phone3552 = p3552;
    } else if (available.length === 1) {
      day.phone3551 = available[0];
      day.phone3552 = null;
      counts[available[0]]++;
    }
  }
  return newSchedule;
}

function solveDocAssignments(
  dates: string[],
  availableDocsPerDate: Record<string, string[]>,
  prevDoc: string | null
): Record<string, string> | null {
  const N = dates.length;
  const M = dates.filter(date => (availableDocsPerDate[date] || []).length > 0).length;
  if (M === 0) {
    const emptyResult: Record<string, string> = {};
    dates.forEach(d => emptyResult[d] = "");
    return emptyResult;
  }

  const idealMin = Math.floor(M / GROUP_2.length);
  const idealMax = Math.ceil(M / GROUP_2.length);

  // We define search limits of [minAllowed, maxAllowed] in order of preference (fairest first)
  const searchLimits: { minAllowed: number; maxAllowed: number }[] = [
    { minAllowed: idealMin, maxAllowed: idealMax },
    { minAllowed: Math.max(0, idealMin - 1), maxAllowed: idealMax + 1 },
    { minAllowed: Math.max(0, idealMin - 2), maxAllowed: idealMax + 2 },
    { minAllowed: 0, maxAllowed: N } // ultimate fallback
  ];

  for (const { minAllowed, maxAllowed } of searchLimits) {
    const result: Record<string, string> = {};
    const counts: Record<string, number> = {};
    GROUP_2.forEach(p => counts[p] = 0);

    function backtrack(index: number): boolean {
      if (index === N) {
        // Enforce the lower bound for all candidates, taking into account their maximum possible availability
        for (const p of GROUP_2) {
          const maxPossible = dates.filter(d => (availableDocsPerDate[d] || []).includes(p)).length;
          const targetMin = Math.min(minAllowed, maxPossible);
          if (counts[p] < targetMin) {
            return false;
          }
        }
        return true;
      }

      const date = dates[index];
      const available = availableDocsPerDate[date] || [];
      if (available.length === 0) {
        result[date] = "";
        return backtrack(index + 1);
      }

      // Sort candidates by current assignment counts so we try the least-assigned first
      const candidates = [...available].sort((a, b) => counts[a] - counts[b]);

      for (const candidate of candidates) {
        // Prevent consecutive assignments
        if (index > 0) {
          if (result[dates[index - 1]] === candidate) continue;
        } else {
          if (prevDoc === candidate) continue;
        }

        // Prevent exceeding the maxAllowed for this search round
        if (counts[candidate] >= maxAllowed) {
          continue;
        }

        result[date] = candidate;
        counts[candidate]++;

        if (backtrack(index + 1)) {
          return true;
        }

        counts[candidate]--;
        result[date] = "";
      }

      return false;
    }

    if (backtrack(0)) {
      return result;
    }
  }

  return null;
}

function getInitialSchedule(year: number, month: number) {
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const schedule: Record<string, DaySchedule> = {};
  const dates: string[] = [];

  for (let d = 1; d <= daysInMonth; d++) {
    const dateObj = new Date(year, month, d);
    const dateStr = `${year}-${String(month + 1).padStart(2, "0")}-${String(
      d
    ).padStart(2, "0")}`;
    const isWeekend = dateObj.getDay() === 0 || dateObj.getDay() === 6;

    let workingStaff = [...GROUP_2];
    if (!isWeekend) {
      workingStaff = [...workingStaff, ...GROUP_1];
    }

    dates.push(dateStr);
    schedule[dateStr] = {
      date: dateStr,
      workingStaff,
      fireCodes: assignFireCodes(workingStaff),
      vacationStaff: [],
      docInCharge: null,
    };
  }

  // Find previous month last day's doc from localStorage if it exists
  let prevDoc: string | null = null;
  try {
    const saved = localStorage.getItem("hospitalSchedule");
    if (saved) {
      const parsed = JSON.parse(saved);
      const prevDateObj = new Date(year, month, 1);
      prevDateObj.setDate(prevDateObj.getDate() - 1);
      const prevDateStr = `${prevDateObj.getFullYear()}-${String(prevDateObj.getMonth() + 1).padStart(2, '0')}-${String(prevDateObj.getDate()).padStart(2, '0')}`;
      prevDoc = parsed[prevDateStr]?.docInCharge || null;
    }
  } catch (e) {
    // Ignore error
  }

  // Perfect Doc assignment solver for initial schedule
  const availableDocsPerDate: Record<string, string[]> = {};
  dates.forEach(date => {
    availableDocsPerDate[date] = [...GROUP_2];
  });

  const result = solveDocAssignments(dates, availableDocsPerDate, prevDoc);

  if (!result) {
    let docIndex = 0;
    dates.forEach(date => {
      schedule[date].docInCharge = GROUP_2[docIndex % GROUP_2.length];
      docIndex++;
    });
  } else {
    dates.forEach(date => {
      schedule[date].docInCharge = result[date] || null;
    });
  }

  return rebalancePhones(schedule);
}

export default function App() {
  const currentNow = new Date();
  const [currentMonthStr, setCurrentMonthStr] = useState(`${currentNow.getFullYear()}-${String(currentNow.getMonth() + 1).padStart(2, '0')}`);

  const [selectedDate, setSelectedDate] = useState<string>(() => {
    const today = new Date();
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  });

  const [viewMode, setViewMode] = useState<"daily" | "monthly">("daily");

  const [isAdmin, setIsAdmin] = useState<boolean>(() => {
    return localStorage.getItem("hospitalSchedule_isAdmin") === "true";
  });
  const [isAdminConsoleOpen, setIsAdminConsoleOpen] = useState<boolean>(false);
  const [currentUser, setCurrentUser] = useState<string | null>(() => {
    return localStorage.getItem("hospitalSchedule_currentUser");
  });
  const [selectedStaffFilter, setSelectedStaffFilter] = useState<string | null>(null);
  const [filterSearchQuery, setFilterSearchQuery] = useState<string>("");
  const [isolateStaffRow, setIsolateStaffRow] = useState<boolean>(false);
  const [spotlightTab, setSpotlightTab] = useState<"stats" | "calendar">("calendar");
  const [quickReqDate, setQuickReqDate] = useState<string | null>(null);
  const [quickReqType, setQuickReqType] = useState<"swap" | "cover" | "leave" | "off" | "work">("leave");
  const [quickReqTarget, setQuickReqTarget] = useState<string>("");
  const [quickReqTargetDate, setQuickReqTargetDate] = useState<string>("");
  const [quickReqNote, setQuickReqNote] = useState<string>("");
  const [loginTab, setLoginTab] = useState<"staff" | "admin">("staff");
  const [selectedLoginStaff, setSelectedLoginStaff] = useState<string>(ALL_STAFF[0]);
  const [showAdminModal, setShowAdminModal] = useState(false);
  const [passwordInput, setPasswordInput] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [isScheduleDirty, setIsScheduleDirty] = useState<boolean>(false);

  const [holidays, setHolidays] = useState<Holiday[]>(() => {
    const saved = localStorage.getItem("hospitalSchedule_holidays");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {}
    }
    const y = currentNow.getFullYear();
    const m = String(currentNow.getMonth() + 1).padStart(2, '0');
    return [
      { date: `${y}-01-01`, name: "วันขึ้นปีใหม่", type: "public" },
      { date: `${y}-04-13`, name: "วันสงกรานต์", type: "public" },
      { date: `${y}-04-14`, name: "วันสงกรานต์", type: "public" },
      { date: `${y}-04-15`, name: "วันสงกรานต์", type: "public" },
      { date: `${y}-05-01`, name: "วันแรงงานแห่งชาติ", type: "public" },
      { date: `${y}-06-03`, name: "วันเฉลิมพระชนมพรรษาสมเด็จพระนางเจ้าฯ พระบรมราชินี", type: "public" },
      { date: `${y}-07-28`, name: "วันเฉลิมพระชนมพรรษา ร.10", type: "public" },
      { date: `${y}-08-12`, name: "วันแม่แห่งชาติ", type: "public" },
      { date: `${y}-10-13`, name: "วันคล้ายวันสวรรคต ร.9", type: "public" },
      { date: `${y}-12-05`, name: "วันพ่อแห่งชาติ", type: "public" },
      { date: `${y}-12-31`, name: "วันสิ้นปี", type: "public" },
      // Company holidays
      { date: `${y}-${m}-12`, name: "วันครบรอบบริษัท / Outing", type: "company" },
      { date: `${y}-${m}-26`, name: "วันประชุมใหญ่บริษัท", type: "company" },
    ];
  });

  const [newHolidayDate, setNewHolidayDate] = useState<string>("");
  const [newHolidayName, setNewHolidayName] = useState<string>("");
  const [newHolidayType, setNewHolidayType] = useState<"public" | "company">("public");

  useEffect(() => {
    localStorage.setItem("hospitalSchedule_holidays", JSON.stringify(holidays));
  }, [holidays]);

  useEffect(() => {
    if (currentUser) {
      localStorage.setItem("hospitalSchedule_currentUser", currentUser);
    } else {
      localStorage.removeItem("hospitalSchedule_currentUser");
    }
  }, [currentUser]);

  const [schedule, setSchedule] = useState<Record<string, DaySchedule>>(() => {
    const saved = localStorage.getItem("hospitalSchedule");
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        Object.values(parsed).forEach((day: any, idx) => {
          if (!day.vacationStaff) day.vacationStaff = [];
          if (!day.docInCharge) day.docInCharge = GROUP_2[idx % GROUP_2.length];
          if (!day.phoneInCharge) day.phoneInCharge = GROUP_2[(idx + 4) % GROUP_2.length];
        });
        return parsed;
      } catch (e) {
        // Fallback below
      }
    }
    return getInitialSchedule(currentNow.getFullYear(), currentNow.getMonth());
  });

  const [logs, setLogs] = useState<LogEntry[]>(() => {
    const saved = localStorage.getItem("hospitalLogs");
    return saved ? JSON.parse(saved) : [];
  });

  useEffect(() => {
    localStorage.setItem("hospitalSchedule", JSON.stringify(schedule));
  }, [schedule]);

  useEffect(() => {
    localStorage.setItem("hospitalLogs", JSON.stringify(logs));
  }, [logs]);

  const [toasts, setToasts] = useState<ToastMessage[]>([]);

  // Shift Swaps and Notification States
  const [showNotificationDropdown, setShowNotificationDropdown] = useState(false);
  
  const [shiftRequests, setShiftRequests] = useState<ShiftRequest[]>(() => {
    const saved = localStorage.getItem("hospitalSchedule_requests");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {}
    }
    // Default mock requests for demo
    const y = currentNow.getFullYear();
    const m = String(currentNow.getMonth() + 1).padStart(2, '0');
    return [
      {
        id: "req-1",
        requester: "อุษา",
        type: "swap" as const,
        date: `${y}-${m}-15`,
        targetStaff: "ไพโรจน์",
        targetDate: `${y}-${m}-18`,
        status: "approved" as const,
        note: "ขอแลกเวรเนื่องจากติดธุระพาครอบครัวตรวจสุขภาพ",
        createdAt: new Date(Date.now() - 3600000 * 24).toISOString()
      },
      {
        id: "req-2",
        requester: "รวีวรรณ",
        type: "leave" as const,
        date: `${y}-${m}-20`,
        status: "pending" as const,
        note: "ขอลาพักผ่อนติดภารกิจครอบครัวที่ต่างจังหวัด",
        createdAt: new Date(Date.now() - 3600000 * 4).toISOString()
      },
      {
        id: "req-3",
        requester: "ศศลักษณ์",
        type: "cover" as const,
        date: `${y}-${m}-22`,
        targetStaff: "ธนัชพร",
        status: "pending" as const,
        note: "ขอแลกให้คุณธนัชพรช่วยขึ้นเวรหลักแทนเนื่องจากติดสัมมนาวิชาการ",
        createdAt: new Date(Date.now() - 3600000 * 1).toISOString()
      }
    ];
  });

  const [notifications, setNotifications] = useState<NotificationItem[]>(() => {
    const saved = localStorage.getItem("hospitalSchedule_notifications");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {}
    }
    // Default initial mock notifications to enrich the dashboard on load
    return [
      {
        id: "notif-1",
        title: "อนุมัติการสลับเวรสำเร็จ 🎉",
        message: "สลับเวรสำเร็จ: คุณอุษา ได้ทำการสลับเวรวันที่ 15 กับ คุณไพโรจน์ วันที่ 18 เรียบร้อยแล้ว ระบบสลับข้อมูลในปฏิทินและคำนวณเบอร์รับสายใหม่อัตโนมัติ",
        timestamp: new Date(Date.now() - 3600000 * 24).toISOString(),
        isRead: false,
        type: "success" as const
      },
      {
        id: "notif-2",
        title: "มีคำขอสลับเวรใหม่เข้ามา",
        message: "คำขอลาพักร้อน: คุณรวีวรรณ ยื่นคำขอลาพักร้อนในวันที่ 20 รอผู้ดูแลระบบอนุมัติข้อมูล",
        timestamp: new Date(Date.now() - 3600000 * 4).toISOString(),
        isRead: false,
        type: "info" as const
      },
      {
        id: "notif-3",
        title: "มีคำขอสลับเวรใหม่เข้ามา",
        message: "คำขอให้ขึ้นแทน: คุณศศลักษณ์ ร้องขอให้ คุณธนัชพร ขึ้นเวรแทนในวันที่ 22 รอผู้เกี่ยวข้องหรือผู้ดูแลระบบตอบรับข้อมูล",
        timestamp: new Date(Date.now() - 3600000 * 1).toISOString(),
        isRead: false,
        type: "info" as const
      }
    ];
  });

  // Request Form States
  const [reqRequester, setReqRequester] = useState<string>(() => localStorage.getItem("hospitalSchedule_currentUser") || "อุษา");

  useEffect(() => {
    if (currentUser) {
      setReqRequester(currentUser);
      if (reqTargetStaff === currentUser) {
        setReqTargetStaff(ALL_STAFF.find(s => s !== currentUser) || "");
      }
    }
  }, [currentUser]);
  const [reqType, setReqType] = useState<"swap" | "cover" | "leave">("swap");
  const [reqDate, setReqDate] = useState<string>(selectedDate);
  const [reqTargetStaff, setReqTargetStaff] = useState<string>("ไพโรจน์");
  const [reqTargetDate, setReqTargetDate] = useState<string>(selectedDate);
  const [reqNote, setReqNote] = useState<string>("");
  const [activeReqTab, setActiveReqTab] = useState<"all" | "pending">("pending");

  useEffect(() => {
    localStorage.setItem("hospitalSchedule_requests", JSON.stringify(shiftRequests));
  }, [shiftRequests]);

  useEffect(() => {
    localStorage.setItem("hospitalSchedule_notifications", JSON.stringify(notifications));
  }, [notifications]);

  // Keep reqDate and reqTargetDate in sync with selectedDate when it changes to make the form easier to use
  useEffect(() => {
    setReqDate(selectedDate);
    // Find next day or same day for target date
    const [y, m, d] = selectedDate.split('-').map(Number);
    const targetDateObj = new Date(y, m - 1, d + 3); // Default 3 days later
    const targetY = targetDateObj.getFullYear();
    const targetM = String(targetDateObj.getMonth() + 1).padStart(2, '0');
    const targetD = String(targetDateObj.getDate()).padStart(2, '0');
    setReqTargetDate(`${targetY}-${targetM}-${targetD}`);
  }, [selectedDate]);

  useEffect(() => {
    localStorage.setItem("hospitalSchedule_isAdmin", String(isAdmin));
  }, [isAdmin]);

  useEffect(() => {
    const [yearStr, monthStr] = currentMonthStr.split('-');
    const year = parseInt(yearStr);
    const month = parseInt(monthStr) - 1;

    const testDateStr = `${yearStr}-${monthStr}-01`;
    setSchedule(prev => {
      if (!prev[testDateStr]) {
        const newMonth = getInitialSchedule(year, month);
        return { ...prev, ...newMonth };
      }
      return prev;
    });

    // Sync selectedDate to match current month if it differs
    setSelectedDate(prev => {
      if (prev.startsWith(`${yearStr}-${monthStr}`)) {
        return prev;
      }
      return `${yearStr}-${monthStr}-01`;
    });
  }, [currentMonthStr]);

  // Form State
  const [yearStr, monthStr] = currentMonthStr.split('-');
  const datesInMonth = Object.values(schedule)
    .filter((day: DaySchedule) => day.date.startsWith(`${yearStr}-${monthStr}`))
    .sort((a: DaySchedule, b: DaySchedule) => a.date.localeCompare(b.date))
    .map((day: DaySchedule) => {
      const d = new Date(day.date);
      const dayName = d.toLocaleDateString("th-TH", { weekday: "short" });
      const dateNum = d.getDate();
      const isWeekend = d.getDay() === 0 || d.getDay() === 6;
      const holiday = holidays.find((h: Holiday) => h.date === day.date);
      return { date: day.date, dayName, dateNum, isWeekend, obj: d, holiday };
    });

  const getStaffStats = (staffName: string) => {
    let workingDays = 0;
    let offDays = 0;
    let vacationDays = 0;
    let docInChargeDays = 0;
    let phone3551Days = 0;
    let phone3552Days = 0;
    const dutyDates: Array<{ date: string; dayName: string; dateNum: number; roleBadge?: string }> = [];

    datesInMonth.forEach(d => {
      const dayData = schedule[d.date];
      if (!dayData) return;
      
      const isWorking = dayData.workingStaff.includes(staffName);
      const isVacation = dayData.vacationStaff?.includes(staffName);
      const isDoc = dayData.docInCharge === staffName;
      const isP1 = dayData.phone3551 === staffName;
      const isP2 = dayData.phone3552 === staffName;

      if (isVacation) {
        vacationDays++;
      } else if (isWorking) {
        workingDays++;
        const roles: string[] = [];
        if (isP1) roles.push("3551");
        if (isP2) roles.push("3552");
        if (isDoc) roles.push("Doc");
        
        dutyDates.push({
          date: d.date,
          dayName: d.dayName,
          dateNum: d.dateNum,
          roleBadge: roles.length > 0 ? roles.join("/") : undefined
        });
      } else {
        offDays++;
      }

      if (isDoc) docInChargeDays++;
      if (isP1) phone3551Days++;
      if (isP2) phone3552Days++;
    });

    return {
      workingDays,
      offDays,
      vacationDays,
      docInChargeDays,
      phone3551Days,
      phone3552Days,
      dutyDates
    };
  };

  const toggleCellState = (date: string, staff: string) => {
    if (!isAdmin && currentUser !== staff) {
      const formattedDate = new Date(date).toLocaleDateString("th-TH", {
        day: "numeric",
        month: "short",
        year: "numeric"
      });
      if (currentUser) {
        showToast(`คุณเข้าระบบเป็น "${currentUser}" สามารถแก้ไขสถานะเวรของตนเองได้เท่านั้น ไม่สามารถแก้ไขของ "${staff}" ได้`, false, formattedDate);
      } else {
        showToast("กรุณาเข้าสู่ระบบด้วยชื่อของคุณก่อนเพื่อแก้ไขสถานะเวร (รหัสผ่าน 1234)", false, formattedDate);
        setLoginTab("staff");
        setSelectedLoginStaff(staff);
        setShowAdminModal(true);
      }
      return;
    }

    const dayObj = schedule[date];
    if (dayObj) {
      const isWorking = dayObj.workingStaff.includes(staff);
      const isVac = (dayObj.vacationStaff || []).includes(staff);

      let nextState = 'WORK';
      if (isWorking) nextState = 'OFF';
      else if (isVac) nextState = 'WORK';
      else nextState = 'VAC';

      const workingG2 = dayObj.workingStaff.filter(s => GROUP_2.includes(s));
      const currentlyWorking = workingG2.includes(staff);
      
      let nextOffG2Count = 9 - workingG2.length;
      if (GROUP_2.includes(staff)) {
        if (currentlyWorking && (nextState === 'OFF' || nextState === 'VAC')) {
          nextOffG2Count += 1;
        } else if (!currentlyWorking && nextState === 'WORK') {
          nextOffG2Count = Math.max(0, nextOffG2Count - 1);
        }
      }
      
      if (nextOffG2Count > 4) {
        const readableDateStr = new Date(date).toLocaleDateString("th-TH", { day: 'numeric', month: 'short' });
        const newWarningNotification: NotificationItem = {
          id: `notif-warn-${Date.now()}`,
          title: "⚠️ แจ้งเตือน: มีการหยุดงานเกิน 4 คนในทีม! ⚠️",
          message: `เนื่องจากมีการเปลี่ยนสถานะของ คุณ${staff} ในวันที่ ${readableDateStr} ทำให้มีทีมงานหยุดงานรวม ${nextOffG2Count} คน (ทีมเหลือผู้ปฏิบัติงานน้อยกว่า 5 คน)`,
          timestamp: new Date().toISOString(),
          isRead: false,
          type: "warning"
        };
        setNotifications(prev => [newWarningNotification, ...prev]);
        showToast(`⚠️ เตือน: วันที่ ${readableDateStr} มีทีมหยุดรวม ${nextOffG2Count} คน (เกิน 4 คนในทีม!)`, true, readableDateStr);
      }
    }

    setSchedule(prev => {
      const docCounts: Record<string, number> = {};
      Object.keys(prev).forEach(dKey => {
        if (dKey !== date) { // Count other days
          const doc = prev[dKey].docInCharge;
          if (doc) docCounts[doc] = (docCounts[doc] || 0) + 1;
        }
      });

      const day = prev[date];
      let newWorking = [...day.workingStaff];
      let newVac = [...(day.vacationStaff || [])];
      let newDoc = day.docInCharge;

      const isWorking = newWorking.includes(staff);
      const isVac = newVac.includes(staff);

      let nextState = 'WORK';
      if (isWorking) nextState = 'OFF';
      else if (isVac) nextState = 'WORK';
      else nextState = 'VAC';

      if (nextState === 'WORK') {
        newWorking.push(staff);
        newVac = newVac.filter(s => s !== staff);
      } else if (nextState === 'VAC') {
        newWorking = newWorking.filter(s => s !== staff);
        newVac.push(staff);
      } else {
        newWorking = newWorking.filter(s => s !== staff);
        newVac = newVac.filter(s => s !== staff);
      }

      if (staff === day.docInCharge && nextState !== 'WORK') {
        // Find best next doc from working staff in GROUP 2
        const availableDocs = newWorking.filter(s => GROUP_2.includes(s));
        if (availableDocs.length > 0) {
          availableDocs.sort((a, b) => (docCounts[a] || 0) - (docCounts[b] || 0));
          newDoc = availableDocs[0];
        } else {
          newDoc = null;
        }
      }
      if (!newDoc && nextState === 'WORK' && GROUP_2.includes(staff)) {
        newDoc = staff;
      }

      const inChargeTriggered = newDoc === staff && nextState === 'WORK';

      const logEntry: LogEntry = {
        id: Date.now(),
        date,
        timestamp: new Date().toISOString(),
        personOut: staff,
        personIn: nextState === 'WORK' ? 'ขึ้นเวร' : nextState === 'VAC' ? 'พักร้อน' : 'หยุด',
        inChargeTriggered,
      };
      setLogs(l => [logEntry, ...l].slice(0, 100));

      const formattedDate = new Date(date).toLocaleDateString("th-TH", {
        day: "numeric",
        month: "short",
        year: "numeric"
      });
      setIsScheduleDirty(false);

      return rebalancePhones({
        ...prev,
        [date]: {
          ...day,
          workingStaff: newWorking,
          vacationStaff: newVac,
          fireCodes: assignFireCodes(newWorking),
          docInCharge: newDoc
        }
      });
    });
  };

  const setStaffStatus = (date: string, staff: string, targetStatus: 'WORK' | 'OFF' | 'VAC') => {
    if (!isAdmin && currentUser !== staff) {
      const formattedDate = new Date(date).toLocaleDateString("th-TH", {
        day: "numeric",
        month: "short",
        year: "numeric"
      });
      if (currentUser) {
        showToast(`คุณเข้าระบบเป็น "${currentUser}" สามารถแก้ไขสถานะเวรของตนเองได้เท่านั้น ไม่สามารถแก้ไขของ "${staff}" ได้`, false, formattedDate);
      } else {
        showToast("กรุณาเข้าสู่ระบบด้วยชื่อของคุณก่อนเพื่อแก้ไขสถานะเวร (รหัสผ่าน 1234)", false, formattedDate);
        setLoginTab("staff");
        setSelectedLoginStaff(staff);
        setShowAdminModal(true);
      }
      return;
    }

    const dayObj = schedule[date];
    if (dayObj) {
      const workingG2 = dayObj.workingStaff.filter(s => GROUP_2.includes(s));
      const currentlyWorking = workingG2.includes(staff);
      
      let nextOffG2Count = 9 - workingG2.length;
      if (GROUP_2.includes(staff)) {
        if (currentlyWorking && (targetStatus === 'OFF' || targetStatus === 'VAC')) {
          nextOffG2Count += 1;
        } else if (!currentlyWorking && targetStatus === 'WORK') {
          nextOffG2Count = Math.max(0, nextOffG2Count - 1);
        }
      }
      
      if (nextOffG2Count > 4) {
        const readableDateStr = new Date(date).toLocaleDateString("th-TH", { day: 'numeric', month: 'short' });
        const newWarningNotification: NotificationItem = {
          id: `notif-warn-${Date.now()}`,
          title: "⚠️ แจ้งเตือน: มีการหยุดงานเกิน 4 คนในทีม! ⚠️",
          message: `เนื่องจากมีการเปลี่ยนสถานะของ คุณ${staff} เป็น ${targetStatus === 'WORK' ? 'ขึ้นเวร' : targetStatus === 'VAC' ? 'พักร้อน' : 'หยุด'} ในวันที่ ${readableDateStr} ทำให้มีทีมงานหยุดงานรวม ${nextOffG2Count} คน (ทีมเหลือผู้ปฏิบัติงานน้อยกว่า 5 คน)`,
          timestamp: new Date().toISOString(),
          isRead: false,
          type: "warning"
        };
        setNotifications(prev => [newWarningNotification, ...prev]);
        showToast(`⚠️ เตือน: วันที่ ${readableDateStr} มีทีมหยุดรวม ${nextOffG2Count} คน (เกิน 4 คนในทีม!)`, true, readableDateStr);
      }
    }

    setSchedule(prev => {
      const docCounts: Record<string, number> = {};
      Object.keys(prev).forEach(dKey => {
        if (dKey !== date) {
          const doc = prev[dKey].docInCharge;
          if (doc) docCounts[doc] = (docCounts[doc] || 0) + 1;
        }
      });

      const day = prev[date];
      if (!day) return prev;

      let newWorking = [...day.workingStaff];
      let newVac = [...(day.vacationStaff || [])];
      let newDoc = day.docInCharge;

      newWorking = newWorking.filter(s => s !== staff);
      newVac = newVac.filter(s => s !== staff);

      if (targetStatus === 'WORK') {
        newWorking.push(staff);
      } else if (targetStatus === 'VAC') {
        newVac.push(staff);
      }

      if (staff === day.docInCharge && targetStatus !== 'WORK') {
        const availableDocs = newWorking.filter(s => GROUP_2.includes(s));
        if (availableDocs.length > 0) {
          availableDocs.sort((a, b) => (docCounts[a] || 0) - (docCounts[b] || 0));
          newDoc = availableDocs[0];
        } else {
          newDoc = null;
        }
      }
      if (!newDoc && targetStatus === 'WORK' && GROUP_2.includes(staff)) {
        newDoc = staff;
      }

      const inChargeTriggered = newDoc === staff && targetStatus === 'WORK';

      const logEntry: LogEntry = {
        id: Date.now(),
        date,
        timestamp: new Date().toISOString(),
        personOut: staff,
        personIn: targetStatus === 'WORK' ? 'ขึ้นเวร' : targetStatus === 'VAC' ? 'พักร้อน' : 'หยุด',
        inChargeTriggered,
      };
      setLogs(l => [logEntry, ...l].slice(0, 100));

      const formattedDate = new Date(date).toLocaleDateString("th-TH", {
        day: "numeric",
        month: "short",
        year: "numeric"
      });
      setIsScheduleDirty(false);

      return rebalancePhones({
        ...prev,
        [date]: {
          ...day,
          workingStaff: newWorking,
          vacationStaff: newVac,
          fireCodes: assignFireCodes(newWorking),
          docInCharge: newDoc
        }
      });
    });
  };

  // Submit new Shift Swap/Change Request
  const handleCreateShiftRequest = (e: React.FormEvent) => {
    e.preventDefault();

    if (reqType === "swap" && reqRequester === reqTargetStaff) {
      showToast("ไม่สามารถสลับเวรกับตัวเองได้", false, reqDate);
      return;
    }
    if (reqType === "cover" && reqRequester === reqTargetStaff) {
      showToast("ไม่สามารถเลือกตัวเองขึ้นเวรแทนได้", false, reqDate);
      return;
    }

    const newRequest: ShiftRequest = {
      id: `req-${Date.now()}`,
      requester: reqRequester,
      type: reqType,
      date: reqDate,
      targetStaff: (reqType === "swap" || reqType === "cover") ? reqTargetStaff : undefined,
      targetDate: reqType === "swap" ? reqTargetDate : undefined,
      status: "pending",
      note: reqNote.trim() || undefined,
      createdAt: new Date().toISOString()
    };

    setShiftRequests(prev => [newRequest, ...prev]);

    // Create Notification about this request
    const readableDate = new Date(reqDate).toLocaleDateString("th-TH", { day: "numeric", month: "short" });
    const readableTargetDate = reqType === "swap" ? new Date(reqTargetDate).toLocaleDateString("th-TH", { day: "numeric", month: "short" }) : "";
    
    let notifMsg = "";
    let notifTitle = "มีคำขอเปลี่ยนเวรใหม่ 📢";
    if (reqType === "swap") {
      notifTitle = "คำขอสลับเวรใหม่เข้ามา 🤝";
      notifMsg = `คุณ${reqRequester} ขอสลับเวรวันที่ ${readableDate} กับ คุณ${reqTargetStaff} วันที่ ${readableTargetDate}`;
    } else if (reqType === "cover") {
      notifTitle = "คำขอให้ขึ้นเวรแทน 🙋‍♀️";
      notifMsg = `คุณ${reqRequester} ขอให้ คุณ${reqTargetStaff} ขึ้นเวรแทนในวันที่ ${readableDate}`;
    } else {
      notifTitle = "แจ้งขอลาพักร้อน 🌴";
      notifMsg = `คุณ${reqRequester} แจ้งขอลาพักร้อนในวันที่ ${readableDate}`;
    }

    if (reqNote.trim()) {
      notifMsg += ` (${reqNote.trim()})`;
    }

    const newNotification: NotificationItem = {
      id: `notif-${Date.now()}`,
      title: notifTitle,
      message: notifMsg,
      timestamp: new Date().toISOString(),
      isRead: false,
      type: "info"
    };

    setNotifications(prev => [newNotification, ...prev]);
    
    // Clear note form field
    setReqNote("");

    // Success Toast
    showToast(`💾 [บันทึกอัตโนมัติ] ส่งคำขอแล้ว และระบบได้บันทึกข้อมูลเรียบร้อยแล้ว`, false, reqDate);
  };

  // Approve a shift request and apply mutations to the schedule
  const handleApproveShiftRequest = (req: ShiftRequest) => {
    setSchedule(prev => {
      const newSchedule = { ...prev };
      const day1 = newSchedule[req.date];
      
      if (!day1) {
        showToast("ไม่พบข้อมูลวันที่ในปฏิทิน", false, req.date);
        return prev;
      }

      // Calculate docCounts for automatic In-Charge re-evaluation if needed
      const docCounts: Record<string, number> = {};
      Object.keys(prev).forEach(dKey => {
        if (dKey !== req.date && (req.type !== "swap" || dKey !== req.targetDate)) {
          const doc = prev[dKey].docInCharge;
          if (doc) docCounts[doc] = (docCounts[doc] || 0) + 1;
        }
      });

      if (req.type === "swap") {
        const day2 = req.targetDate ? newSchedule[req.targetDate] : null;
        if (!day2) {
          showToast("ไม่พบข้อมูลวันสลับเวรในปฏิทิน", false, req.targetDate || "");
          return prev;
        }

        // Swapping logic on Day 1 & Day 2
        let working1 = [...day1.workingStaff];
        let vac1 = [...(day1.vacationStaff || [])];
        let working2 = [...day2.workingStaff];
        let vac2 = [...(day2.vacationStaff || [])];

        // Status of requester and target on Day 1
        const reqOnDay1 = working1.includes(req.requester);
        const reqVacDay1 = vac1.includes(req.requester);
        const tarOnDay1 = working1.includes(req.targetStaff!);
        const tarVacDay1 = vac1.includes(req.targetStaff!);

        // Status of requester and target on Day 2
        const reqOnDay2 = working2.includes(req.requester);
        const reqVacDay2 = vac2.includes(req.requester);
        const tarOnDay2 = working2.includes(req.targetStaff!);
        const tarVacDay2 = vac2.includes(req.targetStaff!);

        // Perform swap on Day 1:
        // If requester was working on Day 1, target takes their place.
        // If target was working on Day 1, requester takes their place.
        if (reqOnDay1) {
          working1 = working1.filter(s => s !== req.requester);
          if (!working1.includes(req.targetStaff!)) working1.push(req.targetStaff!);
        }
        if (tarOnDay1) {
          working1 = working1.filter(s => s !== req.targetStaff!);
          if (!working1.includes(req.requester)) working1.push(req.requester);
        }
        if (reqVacDay1) {
          vac1 = vac1.filter(s => s !== req.requester);
          if (!vac1.includes(req.targetStaff!)) vac1.push(req.targetStaff!);
        }
        if (tarVacDay1) {
          vac1 = vac1.filter(s => s !== req.targetStaff!);
          if (!vac1.includes(req.requester)) vac1.push(req.requester);
        }

        // Perform swap on Day 2:
        if (reqOnDay2) {
          working2 = working2.filter(s => s !== req.requester);
          if (!working2.includes(req.targetStaff!)) working2.push(req.targetStaff!);
        }
        if (tarOnDay2) {
          working2 = working2.filter(s => s !== req.targetStaff!);
          if (!working2.includes(req.requester)) working2.push(req.requester);
        }
        if (reqVacDay2) {
          vac2 = vac2.filter(s => s !== req.requester);
          if (!vac2.includes(req.targetStaff!)) vac2.push(req.targetStaff!);
        }
        if (tarVacDay2) {
          vac2 = vac2.filter(s => s !== req.targetStaff!);
          if (!vac2.includes(req.requester)) vac2.push(req.requester);
        }

        // Re-evaluate Doc In-charge for Day 1
        let doc1 = day1.docInCharge;
        if (req.requester === day1.docInCharge && !working1.includes(req.requester)) {
          const availableDocs = working1.filter(s => GROUP_2.includes(s));
          doc1 = availableDocs.length > 0 ? availableDocs.sort((a,b) => (docCounts[a]||0) - (docCounts[b]||0))[0] : null;
        }
        if (!doc1 && working1.includes(req.targetStaff!) && GROUP_2.includes(req.targetStaff!)) {
          doc1 = req.targetStaff!;
        }

        // Re-evaluate Doc In-charge for Day 2
        let doc2 = day2.docInCharge;
        if (req.targetStaff === day2.docInCharge && !working2.includes(req.targetStaff!)) {
          const availableDocs = working2.filter(s => GROUP_2.includes(s));
          doc2 = availableDocs.length > 0 ? availableDocs.sort((a,b) => (docCounts[a]||0) - (docCounts[b]||0))[0] : null;
        }
        if (!doc2 && working2.includes(req.requester) && GROUP_2.includes(req.requester)) {
          doc2 = req.requester;
        }

        newSchedule[req.date] = {
          ...day1,
          workingStaff: working1,
          vacationStaff: vac1,
          fireCodes: assignFireCodes(working1),
          docInCharge: doc1
        };

        newSchedule[req.targetDate!] = {
          ...day2,
          workingStaff: working2,
          vacationStaff: vac2,
          fireCodes: assignFireCodes(working2),
          docInCharge: doc2
        };

      } else if (req.type === "cover") {
        let working = [...day1.workingStaff];
        let vac = [...(day1.vacationStaff || [])];

        // Requester goes OFF (not working, not vacation)
        working = working.filter(s => s !== req.requester);
        vac = vac.filter(s => s !== req.requester);

        // Target goes WORK
        if (!working.includes(req.targetStaff!)) {
          working.push(req.targetStaff!);
        }
        vac = vac.filter(s => s !== req.targetStaff!);

        // Doc in Charge check
        let doc = day1.docInCharge;
        if (req.requester === day1.docInCharge) {
          const availableDocs = working.filter(s => GROUP_2.includes(s));
          doc = availableDocs.length > 0 ? availableDocs.sort((a,b) => (docCounts[a]||0) - (docCounts[b]||0))[0] : null;
        }
        if (!doc && GROUP_2.includes(req.targetStaff!)) {
          doc = req.targetStaff!;
        }

        newSchedule[req.date] = {
          ...day1,
          workingStaff: working,
          vacationStaff: vac,
          fireCodes: assignFireCodes(working),
          docInCharge: doc
        };

      } else if (req.type === "leave") {
        let working = [...day1.workingStaff];
        let vac = [...(day1.vacationStaff || [])];

        working = working.filter(s => s !== req.requester);
        if (!vac.includes(req.requester)) {
          vac.push(req.requester);
        }

        let doc = day1.docInCharge;
        if (req.requester === day1.docInCharge) {
          const availableDocs = working.filter(s => GROUP_2.includes(s));
          doc = availableDocs.length > 0 ? availableDocs.sort((a,b) => (docCounts[a]||0) - (docCounts[b]||0))[0] : null;
        }

        newSchedule[req.date] = {
          ...day1,
          workingStaff: working,
          vacationStaff: vac,
          fireCodes: assignFireCodes(working),
          docInCharge: doc
        };
      } else if (req.type === "off") {
        let working = [...day1.workingStaff];
        let vac = [...(day1.vacationStaff || [])];

        working = working.filter(s => s !== req.requester);
        vac = vac.filter(s => s !== req.requester);

        let doc = day1.docInCharge;
        if (req.requester === day1.docInCharge) {
          const availableDocs = working.filter(s => GROUP_2.includes(s));
          doc = availableDocs.length > 0 ? availableDocs.sort((a,b) => (docCounts[a]||0) - (docCounts[b]||0))[0] : null;
        }

        newSchedule[req.date] = {
          ...day1,
          workingStaff: working,
          vacationStaff: vac,
          fireCodes: assignFireCodes(working),
          docInCharge: doc
        };
      } else if (req.type === "work") {
        let working = [...day1.workingStaff];
        let vac = [...(day1.vacationStaff || [])];

        if (!working.includes(req.requester)) {
          working.push(req.requester);
        }
        vac = vac.filter(s => s !== req.requester);

        let doc = day1.docInCharge;
        if (!doc && GROUP_2.includes(req.requester)) {
          doc = req.requester;
        }

        newSchedule[req.date] = {
          ...day1,
          workingStaff: working,
          vacationStaff: vac,
          fireCodes: assignFireCodes(working),
          docInCharge: doc
        };
      }

      return rebalancePhones(newSchedule);
    });

    // Update Request Status to approved
    setShiftRequests(prev => prev.map(r => r.id === req.id ? { ...r, status: "approved" as const } : r));

    // Audit logs entry
    const logEntry: LogEntry = {
      id: Date.now(),
      date: req.date,
      timestamp: new Date().toISOString(),
      personOut: req.requester,
      personIn: req.type === "swap" ? `สลับเวรสำเร็จกับ ${req.targetStaff}` : req.type === "cover" ? `ให้ ${req.targetStaff} ขึ้นแทน (อนุมัติแล้ว)` : req.type === "off" ? "ลาหยุด (อนุมัติแล้ว)" : req.type === "work" ? "ขอขึ้นเวร (อนุมัติแล้ว)" : "ลาพักร้อน (อนุมัติแล้ว)",
      inChargeTriggered: false
    };
    setLogs(l => [logEntry, ...l]);

    // Create System Notification
    const readableDate = new Date(req.date).toLocaleDateString("th-TH", { day: "numeric", month: "short" });
    const readableTargetDate = req.type === "swap" ? new Date(req.targetDate!).toLocaleDateString("th-TH", { day: "numeric", month: "short" }) : "";
    
    let sysMsg = "";
    if (req.type === "swap") {
      sysMsg = `📢 ระบบปรับปรุงตารางเวรแล้ว: คุณ${req.requester} ได้สลับเวรวันที่ ${readableDate} กับ คุณ${req.targetStaff} วันที่ ${readableTargetDate} เรียบร้อยแล้ว`;
    } else if (req.type === "cover") {
      sysMsg = `📢 ระบบปรับปรุงตารางเวรแล้ว: คุณ${req.requester} ให้ คุณ${req.targetStaff} ขึ้นเวรแทนในวันที่ ${readableDate} เรียบร้อยแล้ว`;
    } else if (req.type === "off") {
      sysMsg = `📢 ระบบปรับปรุงตารางเวรแล้ว: คุณ${req.requester} ได้รับอนุมัติการลาหยุดงานในวันที่ ${readableDate} เรียบร้อยแล้ว`;
    } else if (req.type === "work") {
      sysMsg = `📢 ระบบปรับปรุงตารางเวรแล้ว: คุณ${req.requester} ได้รับอนุมัติการขึ้นเวรปฏิบัติงานในวันที่ ${readableDate} เรียบร้อยแล้ว`;
    } else {
      sysMsg = `📢 ระบบปรับปรุงตารางเวรแล้ว: คุณ${req.requester} ได้รับอนุมัติการลาพักร้อนในวันที่ ${readableDate} เรียบร้อยแล้ว`;
    }

    const successNotification: NotificationItem = {
      id: `notif-${Date.now()}`,
      title: "อนุมัติเปลี่ยนเวรสำเร็จและอัปเดตระบบแล้ว! 🎉",
      message: sysMsg,
      timestamp: new Date().toISOString(),
      isRead: false,
      type: "success"
    };

    setNotifications(prev => [successNotification, ...prev]);

    // Success Toast
    showToast("💾 [บันทึกอัตโนมัติ] อนุมัติคำขอเปลี่ยนเวรและอัปเดตปฏิทินเรียบร้อยแล้ว", false, req.date);
  };

  // Reject a shift request
  const handleRejectShiftRequest = (reqId: string, requesterName: string, dateStr: string) => {
    setShiftRequests(prev => prev.map(r => r.id === reqId ? { ...r, status: "rejected" as const } : r));

    // Create System Notification
    const rejectNotification: NotificationItem = {
      id: `notif-${Date.now()}`,
      title: "คำขอเปลี่ยนเวรไม่ผ่านการอนุมัติ ❌",
      message: `ผู้ดูแลระบบได้ปฏิเสธคำขอเปลี่ยนเวรของ คุณ${requesterName} ในวันที่ ${new Date(dateStr).toLocaleDateString("th-TH", { day: 'numeric', month: 'short' })}`,
      timestamp: new Date().toISOString(),
      isRead: false,
      type: "warning"
    };

    setNotifications(prev => [rejectNotification, ...prev]);
    showToast(`💾 [บันทึกอัตโนมัติ] ปฏิเสธคำขอเปลี่ยนเวรของ คุณ${requesterName} เรียบร้อยแล้ว`, false, dateStr);
  };

  // Delete/Cancel a request
  const handleDeleteShiftRequest = (reqId: string) => {
    setShiftRequests(prev => prev.filter(r => r.id !== reqId));
    showToast("💾 [บันทึกอัตโนมัติ] ลบคำขอเปลี่ยนเวรออกจากรายการแล้ว", false, selectedDate);
  };

  const handlePrevDay = () => {
    const [y, m, d] = selectedDate.split('-').map(Number);
    const dateObj = new Date(y, m - 1, d);
    dateObj.setDate(dateObj.getDate() - 1);
    const nextY = dateObj.getFullYear();
    const nextM = String(dateObj.getMonth() + 1).padStart(2, '0');
    const nextD = String(dateObj.getDate()).padStart(2, '0');
    const nextDateStr = `${nextY}-${nextM}-${nextD}`;
    const nextMonth = `${nextY}-${nextM}`;
    if (nextMonth !== currentMonthStr) {
      setCurrentMonthStr(nextMonth);
    }
    setSelectedDate(nextDateStr);
  };

  const handleNextDay = () => {
    const [y, m, d] = selectedDate.split('-').map(Number);
    const dateObj = new Date(y, m - 1, d);
    dateObj.setDate(dateObj.getDate() + 1);
    const nextY = dateObj.getFullYear();
    const nextM = String(dateObj.getMonth() + 1).padStart(2, '0');
    const nextD = String(dateObj.getDate()).padStart(2, '0');
    const nextDateStr = `${nextY}-${nextM}-${nextD}`;
    const nextMonth = `${nextY}-${nextM}`;
    if (nextMonth !== currentMonthStr) {
      setCurrentMonthStr(nextMonth);
    }
    setSelectedDate(nextDateStr);
  };

  const handleSaveSchedule = () => {
    localStorage.setItem("hospitalSchedule", JSON.stringify(schedule));
    localStorage.setItem("hospitalLogs", JSON.stringify(logs));
    setIsScheduleDirty(false);
    showToast("💾 บันทึกตารางเวรเรียบร้อยแล้วและเปิดใช้งานทันที!", false, new Date().toLocaleDateString("th-TH"));
    
    const newNotification: NotificationItem = {
      id: `notif-save-${Date.now()}`,
      title: "💾 บันทึกข้อมูลตารางเวรสำเร็จแล้ว! 💚",
      message: `ตารางเวรได้รับการบันทึกข้อมูลอย่างถาวรเรียบร้อยแล้วเมื่อเวลา ${new Date().toLocaleTimeString("th-TH")} น.`,
      timestamp: new Date().toISOString(),
      isRead: false,
      type: "success"
    };
    setNotifications(prev => [newNotification, ...prev]);
  };

  const showToast = (message: string, inCharge: boolean, dateStr: string) => {
    const id = Date.now();
    setToasts((prev) => [{ id, message, inCharge, dateStr }, ...prev]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 8000);
  };

  const handleAddHoliday = () => {
    if (!isAdmin) {
      showToast("เฉพาะ Admin เท่านั้นที่สามารถเพิ่มวันหยุดได้", false, new Date().toLocaleDateString("th-TH"));
      return;
    }
    if (!newHolidayDate || !newHolidayName.trim()) {
      showToast("กรุณากรอกข้อมูลวันหยุดให้ครบถ้วน", false, new Date().toLocaleDateString("th-TH"));
      return;
    }

    // Check if duplicate date
    const exists = holidays.some(h => h.date === newHolidayDate);
    if (exists) {
      showToast("❌ มีการกำหนดวันหยุดในวันที่เลือกนี้อยู่แล้ว", false, new Date().toLocaleDateString("th-TH"));
      return;
    }

    const newH: Holiday = {
      date: newHolidayDate,
      name: newHolidayName.trim(),
      type: newHolidayType
    };

    setHolidays(prev => [...prev, newH].sort((a, b) => a.date.localeCompare(b.date)));
    showToast(`✨ เพิ่มวันหยุด "${newH.name}" เรียบร้อยแล้ว`, false, new Date().toLocaleDateString("th-TH"));
    
    // Clear form
    setNewHolidayDate("");
    setNewHolidayName("");
  };

  const handleAutoScheduleDoc = () => {
    if (!isAdmin) {
      showToast("เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่สามารถจัดเวร Doc อัตโนมัติได้", false, new Date().toLocaleDateString("th-TH"));
      return;
    }

    setSchedule(prev => {
      const dates = datesInMonth.map(d => d.date);
      if (dates.length === 0) return prev;

      const availableDocsPerDate: Record<string, string[]> = {};
      dates.forEach(date => {
        const day = prev[date];
        if (day) {
          availableDocsPerDate[date] = day.workingStaff.filter(s => GROUP_2.includes(s));
        } else {
          availableDocsPerDate[date] = [];
        }
      });

      // Get previous month's last day Doc to prevent consecutive assignments across months
      const prevDateObj = new Date(dates[0]);
      prevDateObj.setDate(prevDateObj.getDate() - 1);
      const prevDateStr = `${prevDateObj.getFullYear()}-${String(prevDateObj.getMonth() + 1).padStart(2, '0')}-${String(prevDateObj.getDate()).padStart(2, '0')}`;
      const prevDoc = prev[prevDateStr]?.docInCharge || null;

      const result = solveDocAssignments(dates, availableDocsPerDate, prevDoc);

      if (!result) {
        showToast("❌ ไม่สามารถจัดเวร Doc อัตโนมัติให้เงื่อนไขสมบูรณ์ได้เนื่องจากข้อจำกัดวันหยุด", false, new Date().toLocaleDateString("th-TH"));
        return prev;
      }

      const newSchedule = { ...prev };
      dates.forEach(date => {
        if (newSchedule[date]) {
          newSchedule[date] = {
            ...newSchedule[date],
            docInCharge: result[date] || null
          };
        }
      });

      showToast("✨ จัดตารางเวร Doc อัตโนมัติเสร็จสิ้น! ทุกคนได้รับมอบหมายอย่างเท่าเทียมและไม่มีเวรต่อกัน", false, new Date().toLocaleDateString("th-TH"));

      const logEntry: LogEntry = {
        id: Date.now(),
        date: dates[0],
        timestamp: new Date().toISOString(),
        personOut: "ระบบจัดเวร",
        personIn: "จัดเวร Doc อัตโนมัติ",
        inChargeTriggered: false,
      };
      setLogs(l => [logEntry, ...l].slice(0, 100));

      return rebalancePhones(newSchedule);
    });
  };

  const handleExportExcel = () => {
    const dates = datesInMonth.map(d => d.date);
    const data: any[][] = [];
    
    const header = ['ชื่อเจ้าหน้าที่', ...dates];
    data.push(header);
    
    DISPLAY_ORDER.forEach(staff => {
      const row = [staff];
      dates.forEach(date => {
        const day = schedule[date];
        let val = 'X';
        if (day?.vacationStaff.includes(staff)) val = 'V';
        else if (day?.workingStaff.includes(staff)) {
          val = 'W';
          if (day.docInCharge === staff) val += ' (Doc)';
          if (day.phone3551 === staff) val += ' {3551}';
          if (day.phone3552 === staff) val += ' {3552}';
          if (day.fireCodes[staff]) val += ` [${day.fireCodes[staff]}]`;
        }
        row.push(val);
      });
      data.push(row);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Schedule");
    XLSX.writeFile(wb, `shift_schedule_${currentMonthStr}.xlsx`);
    
    showToast("ส่งออกตารางเวรเป็นไฟล์ Excel สำเร็จ", false, new Date().toLocaleDateString("th-TH"));
  };

  const handleImportExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!isAdmin) {
      showToast("กรุณาเปิด 'โหมดแก้ไข (Admin)' เพื่อนำเข้าตารางเวร", false, new Date().toLocaleDateString("th-TH"));
      return;
    }
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1 });
      
      const headerRow = data[0];
      if (!headerRow || headerRow[0] !== 'ชื่อเจ้าหน้าที่') {
        alert("รูปแบบไฟล์ไม่ถูกต้อง กรุณาใช้ไฟล์ที่ส่งออกจากระบบนี้เท่านั้น");
        return;
      }
      
      const dates = headerRow.slice(1);
      const newSchedule: Record<string, DaySchedule> = { ...schedule };
      
      dates.forEach((d: string) => {
        if (d) {
           newSchedule[d] = {
             date: d,
             workingStaff: [],
             vacationStaff: [],
             fireCodes: {},
             docInCharge: null
           };
        }
      });

      for (let i = 1; i < data.length; i++) {
         const row = data[i];
         if (!row || row.length === 0) continue;
         const staffName = row[0];
         if (!ALL_STAFF.includes(staffName)) continue;

         for (let j = 1; j < row.length; j++) {
           const date = dates[j-1];
           const cellValue = row[j] || '';
           const dayData = newSchedule[date];
           if (!dayData) continue;

           if (cellValue.includes('V')) {
             dayData.vacationStaff.push(staffName);
           } else if (cellValue.includes('W')) {
             dayData.workingStaff.push(staffName);
             if (cellValue.includes('(Doc)')) {
               dayData.docInCharge = staffName;
             }
             if (cellValue.includes('{3551}')) {
               dayData.phone3551 = staffName;
             }
             if (cellValue.includes('{3552}')) {
               dayData.phone3552 = staffName;
             }
             const fireCodeMatch = cellValue.match(/\[([A-Z]+)\]/);
             if (fireCodeMatch) {
               dayData.fireCodes[staffName] = fireCodeMatch[1];
             }
           }
         }
      }
      
      // Update selected month to match the imported dates
      if (dates.length > 0 && dates[0]) {
        const firstImportedDate = dates[0].substring(0, 7);
        if (firstImportedDate) {
           setCurrentMonthStr(firstImportedDate);
        }
      }
      
      setSchedule(newSchedule);
      showToast("💾 [บันทึกอัตโนมัติ] นำเข้าตารางเวรจาก Excel สำเร็จและบันทึกข้อมูลเรียบร้อยแล้ว", false, new Date().toLocaleDateString("th-TH"));
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  return (
    <div className="min-h-screen bg-gray-50 text-gray-800 font-sans pb-12 flex flex-col">
      <header className="bg-emerald-700 text-white p-4 shadow-md sticky top-0 z-30">
        <div className="w-full px-2 sm:px-4 md:px-6 lg:px-8 mx-auto flex flex-col md:flex-row justify-between items-center gap-4 relative">
          <div className="flex items-center gap-3">
            <div className="bg-white/20 p-2 rounded-lg backdrop-blur-sm">
              <Stethoscope className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-wide">ShiftFlow Medical Checkup Report</h1>
              <p className="text-emerald-100 text-xs mt-0.5">
                ระบบบริหารจัดการตารางเวรและแจ้งเตือน แผนกพิมพ์ผล
              </p>
            </div>
          </div>

          {/* 📅 วันที่ปัจจุบันบนแถบบนสุด */}
          <div className="bg-emerald-800/80 border border-emerald-600/55 px-4 py-2 rounded-2xl flex items-center gap-2 shadow-xs text-xs sm:text-sm font-black text-white">
            <CalendarDays className="w-4 h-4 text-emerald-300 shrink-0" />
            <span>วันนี้: <span className="text-emerald-100">{new Date().toLocaleDateString("th-TH", { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</span></span>
          </div>

          {/* Controls: Notifications & Mode Switcher */}
          <div className="flex items-center gap-3 w-full md:w-auto justify-between sm:justify-start">
            {/* Notification Bell */}
            <div className="relative">
              <button
                onClick={() => {
                  setShowNotificationDropdown(!showNotificationDropdown);
                  if (!showNotificationDropdown) {
                    // Mark as read when opening dropdown
                    setNotifications(prev => prev.map(n => ({ ...n, isRead: true })));
                  }
                }}
                className="relative p-2.5 text-white bg-emerald-800/60 hover:bg-emerald-800 rounded-xl border border-emerald-600/30 shadow-inner transition-all flex items-center justify-center cursor-pointer active:scale-95"
                title="ศูนย์แจ้งเตือนระบบ"
              >
                <Bell className="w-5 h-5" />
                {notifications.some(n => !n.isRead) && (
                  <span className="absolute -top-1.5 -right-1.5 bg-red-500 text-white font-black text-[9px] w-5 h-5 rounded-full flex items-center justify-center animate-bounce border-2 border-emerald-700 shadow-lg">
                    {notifications.filter(n => !n.isRead).length}
                  </span>
                )}
              </button>

              <AnimatePresence>
                {showNotificationDropdown && (
                  <>
                    {/* Backdrop cover for mobile tapping to close */}
                    <div 
                      className="fixed inset-0 z-30 md:hidden" 
                      onClick={() => setShowNotificationDropdown(false)}
                    />
                    
                    <motion.div
                      initial={{ opacity: 0, y: 15, scale: 0.95 }}
                      animate={{ opacity: 1, y: 0, scale: 1 }}
                      exit={{ opacity: 0, y: 15, scale: 0.95 }}
                      className="absolute right-0 mt-3 bg-white text-gray-800 rounded-2xl shadow-2xl border border-gray-100 w-[290px] sm:w-[360px] overflow-hidden z-40 origin-top-right"
                    >
                      <div className="p-4 bg-gradient-to-r from-emerald-600 to-emerald-700 text-white flex items-center justify-between">
                        <h3 className="font-bold text-xs sm:text-sm flex items-center gap-1.5">
                          <Bell className="w-4 h-4 text-emerald-100" />
                          <span>การแจ้งเตือนของแผนก ({notifications.length})</span>
                        </h3>
                        <button
                          onClick={() => {
                            setNotifications([]);
                            showToast("ล้างรายการแจ้งเตือนทั้งหมดแล้ว", false, new Date().toLocaleDateString("th-TH"));
                          }}
                          className="text-[10px] bg-white/20 hover:bg-white/35 active:bg-white/50 px-2 py-1 rounded text-white font-semibold transition-all"
                        >
                          ล้างทั้งหมด
                        </button>
                      </div>

                      <div className="max-h-72 overflow-y-auto custom-scrollbar divide-y divide-gray-50">
                        {notifications.length === 0 ? (
                          <div className="p-8 text-center text-gray-400 text-xs italic">
                            ไม่มีข้อความแจ้งเตือนใหม่ในขณะนี้
                          </div>
                        ) : (
                          notifications.map((notif) => (
                            <div
                              key={notif.id}
                              className={`p-3.5 hover:bg-emerald-50/10 transition-colors flex gap-3 text-left ${
                                !notif.isRead ? "bg-emerald-50/20 font-medium" : ""
                              }`}
                            >
                              <div className="mt-0.5 shrink-0">
                                {notif.type === "success" ? (
                                  <span className="bg-emerald-100 text-emerald-800 p-1.5 rounded-full block">
                                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                                  </span>
                                ) : notif.type === "warning" ? (
                                  <span className="bg-amber-100 text-amber-800 p-1.5 rounded-full block">
                                    <AlertCircle className="w-3.5 h-3.5 text-amber-600" />
                                  </span>
                                ) : (
                                  <span className="bg-sky-100 text-sky-800 p-1.5 rounded-full block">
                                    <Bell className="w-3.5 h-3.5 text-sky-600" />
                                  </span>
                                )}
                              </div>
                              <div className="min-w-0 flex-1">
                                <div className="flex items-center justify-between gap-1.5">
                                  <h4 className="font-bold text-xs text-gray-900 truncate">{notif.title}</h4>
                                  <span className="text-[9px] text-gray-400 font-medium shrink-0">
                                    {new Date(notif.timestamp).toLocaleTimeString("th-TH", {
                                      hour: "2-digit",
                                      minute: "2-digit"
                                    })} น.
                                  </span>
                                </div>
                                <p className="text-[11px] text-gray-600 leading-relaxed mt-1">{notif.message}</p>
                                <span className="text-[9px] text-gray-400 mt-1 block">
                                  {new Date(notif.timestamp).toLocaleDateString("th-TH", {
                                    day: 'numeric',
                                    month: 'short',
                                    year: '2-digit'
                                  })}
                                </span>
                              </div>
                            </div>
                          ))
                        )}
                      </div>
                    </motion.div>
                  </>
                )}
              </AnimatePresence>
            </div>

            {/* Session Controls: Admin / Staff / Guest */}
            <div>
              {isAdmin ? (
                <div className="flex items-center gap-1.5 bg-amber-500/90 border border-amber-400/30 p-1 rounded-xl shadow-md text-white">
                  <span className="flex items-center gap-1.5 px-3 py-1.5 text-xs sm:text-sm font-bold">
                    <Shield className="w-4 h-4 text-amber-100" />
                    <span>ผู้ดูแลระบบ (Admin)</span>
                  </span>
                  <button
                    onClick={() => {
                      setIsAdmin(false);
                      showToast("ออกจากระบบผู้ดูแลระบบเรียบร้อยแล้ว", false, new Date().toLocaleDateString("th-TH"));
                    }}
                    className="bg-white/20 hover:bg-white/30 text-white px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer transition-all active:scale-95 shadow-sm"
                  >
                    ออกจากระบบ
                  </button>
                </div>
              ) : currentUser ? (
                <div className="flex items-center gap-1.5 bg-emerald-800 border border-emerald-600/50 p-1 rounded-xl shadow-inner text-white">
                  <span className="flex items-center gap-1.5 px-3 py-1.5 text-xs sm:text-sm font-bold">
                    <User className="w-4 h-4 text-emerald-300" />
                    <span>คุณ{currentUser} (สิทธิ์จำกัด)</span>
                  </span>
                  <button
                    onClick={() => {
                      setCurrentUser(null);
                      showToast("ออกจากระบบเจ้าหน้าที่เรียบร้อยแล้ว", false, new Date().toLocaleDateString("th-TH"));
                    }}
                    className="bg-white/10 hover:bg-white/20 text-white px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer transition-all active:scale-95 shadow-sm"
                  >
                    ออกจากระบบ
                  </button>
                </div>
              ) : (
                <div className="flex items-center gap-2">
                  <span className="text-xs text-emerald-100 font-medium hidden sm:inline">ผู้ชมทั่วไป (View Only)</span>
                  <button
                    onClick={() => {
                      setLoginTab("staff");
                      setShowAdminModal(true);
                    }}
                    className="flex items-center gap-1.5 bg-white text-emerald-800 hover:bg-emerald-50 px-3 py-1.5 sm:px-4 sm:py-2 rounded-xl text-xs sm:text-sm font-bold transition-all shadow-md cursor-pointer active:scale-95"
                  >
                    <Lock className="w-3.5 h-3.5" />
                    <span>ลงชื่อเข้าใช้งาน (รหัส 1234)</span>
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      </header>

      <main className="w-full flex-1 p-2 sm:p-4 md:px-6 lg:px-8 flex flex-col gap-6 mt-2">
        {/* Toggle between Staff Portal and Admin Controls */}
        {isAdmin && (
          <div className="flex flex-col sm:flex-row bg-white/95 backdrop-blur border-2 border-amber-200 p-3 rounded-2xl shadow-md gap-3 w-full max-w-3xl mx-auto items-center justify-between text-left">
            <div className="flex items-center gap-2.5">
              <Shield className="w-5 h-5 text-amber-500 animate-pulse shrink-0" />
              <div>
                <span className="text-xs font-black text-gray-800 block">แผงควบคุมสิทธิ์ผู้ดูแลระบบ (Admin Control Board)</span>
                <span className="text-[10px] text-gray-500 font-bold block">สลับโหมดการทำงานเพื่อเข้าถึงตารางใหญ่ รายงานวิเคราะห์ชีท และจัดการระบบวันหยุด</span>
              </div>
            </div>
            <div className="flex gap-2 shrink-0">
              <button
                onClick={() => setIsAdminConsoleOpen(false)}
                className={`flex items-center gap-1.5 px-4 py-2 rounded-xl text-xs font-black transition-all cursor-pointer ${
                  !isAdminConsoleOpen
                    ? "bg-emerald-600 text-white shadow-sm"
                    : "bg-gray-100 text-gray-600 hover:bg-gray-200"
                }`}
              >
                <User className="w-4 h-4 text-inherit" />
                <span>👤 หน้าจอพนักงาน (Staff Portal)</span>
              </button>
              <button
                onClick={() => setIsAdminConsoleOpen(true)}
                className={`flex items-center gap-1.5 px-4 py-2 rounded-xl text-xs font-black transition-all cursor-pointer ${
                  isAdminConsoleOpen
                    ? "bg-amber-500 text-white shadow-sm"
                    : "bg-gray-100 text-gray-600 hover:bg-gray-200"
                }`}
              >
                <Settings className="w-4 h-4 text-inherit" />
                <span>🛠️ จัดการระบบ (Admin Console)</span>
              </button>
            </div>
          </div>
        )}

        {!isAdminConsoleOpen ? (
          /* ======================================================= */
          /*   1. UNIFIED SINGLE SCREEN STAFF PORTAL                 */
          /* ======================================================= */
          <div className="w-full max-w-7xl mx-auto flex flex-col gap-6">
            {/* 👤 PROMINENT USER & DATE FOCUS BOARD */}
            <div className="bg-gradient-to-br from-emerald-800 via-emerald-700 to-teal-850 rounded-3xl p-5 sm:p-6 shadow-xl border border-emerald-600/40 text-white flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6 relative overflow-hidden text-left">
              {/* Decorative backgrounds */}
              <div className="absolute top-0 right-0 w-64 h-64 bg-white/5 rounded-full blur-3xl -mr-20 -mt-20 pointer-events-none" />
              <div className="absolute bottom-0 left-0 w-48 h-48 bg-teal-500/10 rounded-full blur-2xl -ml-20 -mb-20 pointer-events-none" />

              {/* USER PROFILE INFO */}
              {(() => {
                const activeStaff = selectedStaffFilter || currentUser || "อุษา";
                const isGroup1 = GROUP_1.includes(activeStaff);
                
                return (
                  <div className="flex items-center gap-4 z-10 w-full lg:w-auto">
                    <div className="w-14 h-14 rounded-full bg-gradient-to-tr from-amber-400 to-orange-500 text-emerald-950 font-black text-xl flex items-center justify-center shadow-lg border-2 border-white shrink-0">
                      {activeStaff.substring(0, 2)}
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="text-[10px] uppercase font-black tracking-widest text-emerald-300">พนักงานผู้ใช้งาน</span>
                        <span className="text-[9px] bg-white/20 border border-white/20 px-2 py-0.5 rounded-full font-bold text-white">
                          {isGroup1 ? "กลุ่มหยุด ส.-อา." : "กลุ่มเวรปกติหมุนเวียน"}
                        </span>
                      </div>
                      <h2 className="text-xl sm:text-2xl font-black mt-1 text-white truncate">
                        <span>คุณ{activeStaff}</span>
                      </h2>
                      
                      {/* Search / Select dropdown to check other staff */}
                      <div className="flex items-center gap-2 mt-2">
                        <span className="text-[10px] text-emerald-100 font-bold whitespace-nowrap">ตรวจสอบตารางของเพื่อน:</span>
                        <select
                          value={activeStaff}
                          onChange={(e) => {
                            setSelectedStaffFilter(e.target.value);
                            showToast(`กำลังตรวจสอบตารางของ คุณ${e.target.value}`, false, new Date().toLocaleDateString("th-TH"));
                          }}
                          className="bg-emerald-900/60 border border-emerald-600 rounded-xl px-2.5 py-1 text-xs font-bold text-white focus:outline-none focus:ring-1 focus:ring-emerald-400 cursor-pointer"
                        >
                          {ALL_STAFF.map(s => (
                            <option key={s} value={s} className="bg-emerald-800 text-white">คุณ{s}</option>
                          ))}
                        </select>
                      </div>
                    </div>
                  </div>
                );
              })()}

              {/* DATE FOCUS */}
              <div className="flex flex-col items-start lg:items-end gap-2 w-full lg:w-auto z-10">
                <span className="text-[10px] uppercase font-black tracking-widest text-emerald-300">วันที่กำลังตรวจสอบ (Selected Date)</span>
                <div className="bg-white/10 backdrop-blur-md border border-white/25 px-4 py-2.5 rounded-2xl flex items-center gap-2 shadow-inner">
                  <CalendarDays className="w-5 h-5 text-amber-400 shrink-0" />
                  <span className="text-sm sm:text-base font-black tracking-wide">
                    {new Date(selectedDate).toLocaleDateString("th-TH", { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}
                  </span>
                </div>
                
                {/* Date Picker Input */}
                <div className="flex items-center gap-2 mt-1 w-full lg:w-auto justify-start lg:justify-end">
                  <span className="text-[10px] text-emerald-100 font-bold">เลือกวันที่อื่น:</span>
                  <input
                    type="date"
                    value={selectedDate}
                    onChange={(e) => {
                      if (e.target.value) {
                        setSelectedDate(e.target.value);
                        const [y, m] = e.target.value.split('-');
                        setCurrentMonthStr(`${y}-${m}`);
                      }
                    }}
                    className="bg-emerald-900/60 border border-emerald-600 rounded-xl px-2.5 py-1 text-xs font-bold text-white focus:outline-none focus:ring-1 focus:ring-emerald-400 cursor-pointer shadow-sm text-center"
                  />
                </div>
              </div>
            </div>

            {/* Touch-scrollable day tape navigator */}
            <div className="bg-white rounded-2xl p-4 border border-emerald-100/50 shadow-xs flex flex-col gap-2 text-left">
              <span className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">เลื่อนแถบเพื่อเลือกวันที่สะดวกรวดเร็ว (Day Tape Navigator):</span>
              <div className="flex gap-2 overflow-x-auto py-1 custom-scrollbar w-full pb-2 scroll-smooth">
                {datesInMonth.map((d) => {
                  const isSelected = selectedDate === d.date;
                  const hasHoliday = d.holiday;
                  const activeStaff = selectedStaffFilter || currentUser || "อุษา";
                  const isStaffOnDuty = schedule[d.date]?.workingStaff.includes(activeStaff);
                  const isStaffOnVacation = schedule[d.date]?.vacationStaff?.includes(activeStaff);

                  let tapeClass = "";
                  if (isSelected) {
                    tapeClass = "bg-emerald-600 text-white shadow-md font-bold scale-105 ring-2 ring-offset-2 ring-emerald-500";
                  } else if (isStaffOnDuty) {
                    tapeClass = "bg-emerald-50 hover:bg-emerald-100 text-emerald-900 border border-emerald-200 ring-1 ring-emerald-400/30";
                  } else if (isStaffOnVacation) {
                    tapeClass = "bg-orange-50 hover:bg-orange-100 text-orange-900 border border-orange-200";
                  } else if (hasHoliday) {
                    if (hasHoliday.type === "public") {
                      tapeClass = "bg-amber-100/70 hover:bg-amber-200 text-amber-800 border border-amber-200";
                    } else {
                      tapeClass = "bg-rose-50 hover:bg-rose-100 text-rose-800 border border-rose-200";
                    }
                  } else if (d.isWeekend) {
                    tapeClass = "bg-gray-100/70 hover:bg-gray-200/80 text-gray-700 border border-gray-200/60";
                  } else {
                    tapeClass = "bg-gray-50 hover:bg-gray-100 text-gray-600 border border-gray-100";
                  }

                  return (
                    <button
                      key={d.date}
                      onClick={() => setSelectedDate(d.date)}
                      className={`flex flex-col items-center justify-center p-2 rounded-xl min-w-[54px] transition-all shrink-0 cursor-pointer relative ${tapeClass}`}
                      title={hasHoliday ? `${hasHoliday.name} (${hasHoliday.type === 'public' ? 'วันหยุดนักขัตฤกษ์' : 'วันหยุดบริษัท'})` : ""}
                    >
                      <span className="text-[8px] uppercase tracking-wider font-bold opacity-85">{d.dayName}</span>
                      <span className="text-sm font-black mt-0.5 flex items-center justify-center gap-0.5">
                        {d.dateNum}
                        {isStaffOnDuty && <span className="w-1 h-1 rounded-full bg-emerald-600 absolute bottom-1"></span>}
                        {isStaffOnVacation && <span className="w-1 h-1 rounded-full bg-orange-500 absolute bottom-1"></span>}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* 📋 ตารางรวมของทั้งแผนก (Prominent Department-wide Schedule Table) */}
            <section id="department-schedule-table" className="w-full bg-white rounded-3xl shadow-lg border-2 border-emerald-100/80 overflow-hidden flex flex-col min-h-[500px]">
              {/* Header section with styling and actions */}
              <div className="p-5 sm:p-6 bg-gradient-to-r from-emerald-800 via-emerald-700 to-teal-800 text-white flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-amber-400 animate-pulse"></span>
                    <h2 className="text-lg sm:text-2xl font-black text-white flex items-center gap-2">
                      <CalendarDays className="w-6 h-6 text-amber-300" />
                      <span>ตารางปฏิบัติงานรวมทั้งแผนก ประจำเดือน</span>
                    </h2>
                  </div>
                  <p className="text-emerald-100/90 text-xs sm:text-sm mt-1 font-medium">
                    รายชื่อและเวรปฏิบัติการของเจ้าหน้าที่ทั้งหมด 14 ท่าน ตรวจสอบเวรสะดวกรวดเร็ว แยกวันเสาร์สีม่วง 🟣 วันอาทิตย์สีแดง 🔴 ชัดเจน
                  </p>
                </div>
                
                {/* Control elements */}
                <div className="flex flex-wrap items-center gap-3">
                  <div className="flex items-center gap-1.5 bg-white/10 backdrop-blur-md px-3 py-1.5 rounded-xl border border-white/20">
                    <span className="text-xs font-bold text-emerald-100">เลือกเดือน:</span>
                    <input 
                      type="month" 
                      value={currentMonthStr} 
                      onChange={(e) => {
                        if (e.target.value) setCurrentMonthStr(e.target.value);
                      }}
                      className="bg-emerald-900/40 text-white text-xs font-bold border border-emerald-500/30 rounded-lg px-2 py-1 focus:ring-1 focus:ring-emerald-300 focus:outline-none cursor-pointer"
                    />
                  </div>

                  {/* Filter / Search inside table */}
                  <div className="relative w-full sm:w-48">
                    <span className="absolute inset-y-0 left-0 pl-2.5 flex items-center text-emerald-300 pointer-events-none">
                      <Search className="w-3.5 h-3.5" />
                    </span>
                    <input
                      type="text"
                      placeholder="ค้นหาชื่อเจ้าหน้าที่..."
                      value={filterSearchQuery}
                      onChange={(e) => setFilterSearchQuery(e.target.value)}
                      className="w-full pl-8 pr-7 py-1.5 bg-white/15 text-white placeholder-emerald-200/70 text-xs rounded-xl border border-white/10 focus:bg-white focus:text-gray-800 focus:outline-none focus:ring-1 focus:ring-emerald-400 transition-all font-semibold"
                    />
                    {filterSearchQuery && (
                      <button
                        onClick={() => setFilterSearchQuery("")}
                        className="absolute inset-y-0 right-0 pr-2 flex items-center text-emerald-300 hover:text-white cursor-pointer"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    )}
                  </div>

                  {/* Export button */}
                  <button 
                    onClick={handleExportExcel} 
                    className="flex items-center text-xs font-bold bg-amber-400 text-emerald-950 px-3.5 py-1.5 rounded-xl hover:bg-amber-300 transition-all shadow-sm active:scale-95 cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5 mr-1" /> ส่งออก Excel
                  </button>
                </div>
              </div>

              {/* Legend & warnings bar */}
              <div className="p-3 sm:px-5 bg-emerald-50/50 border-b border-emerald-100/60 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                <div className="flex flex-wrap gap-x-3 gap-y-1.5 text-[10px] sm:text-xs font-bold text-emerald-950">
                  <span className="flex items-center"><Check className="w-3.5 h-3.5 text-emerald-600 mr-1"/> ขึ้นเวร (WORK)</span>
                  <span className="flex items-center"><X className="w-3.5 h-3.5 text-red-500 mr-1"/> หยุด (OFF)</span>
                  <span className="flex items-center"><span className="text-[10px] font-black text-orange-600 bg-orange-50 px-1 rounded border border-orange-100 mr-1">V</span> พักร้อน (VAC)</span>
                  <span className="flex items-center"><span className="text-[9px] font-black text-blue-700 bg-blue-50 border border-blue-200 px-1 rounded mr-1">Doc.</span> ตรวจเอกสาร</span>
                  <span className="flex items-center"><span className="text-[9px] font-black text-red-600 bg-red-50 border border-red-200 px-1 rounded mr-1">A..I</span> เวรดับเพลิง</span>
                  <span className="flex items-center"><span className="text-[9px] font-black text-sky-700 bg-sky-50 border border-sky-200 px-1 rounded mr-1">📞 3551</span> สายหลัก</span>
                  <span className="flex items-center"><span className="text-[9px] font-black text-indigo-700 bg-indigo-50 border border-indigo-200 px-1 rounded mr-1">📞 3552</span> สายรอง</span>
                </div>
                
                <div className="flex items-center gap-2">
                  <span className="text-[10px] sm:text-xs font-bold text-emerald-800 bg-emerald-50 border border-emerald-150 px-2 py-0.5 rounded-lg">
                    💡 คลิกช่องวันของคุณเพื่อสลับสถานะเวรได้ทันที
                  </span>
                </div>
              </div>

              {/* Table Container with Custom Scrollbar */}
              <div className="overflow-auto max-h-[75vh] flex-1 custom-scrollbar w-full">
                <table className="w-full border-collapse text-sm min-w-max">
                  <thead>
                    <tr>
                      <th className="bg-emerald-800 border-b-2 border-emerald-900 text-white p-3 font-bold text-left sticky top-0 left-0 z-40 min-w-[150px] sm:min-w-[190px] shadow-[3px_0_6px_rgba(0,0,0,0.15)] text-xs sm:text-sm">
                        รายชื่อเจ้าหน้าที่ (14 คน)
                      </th>
                      {datesInMonth.map((d) => {
                        const dayData = schedule[d.date];
                        const workingG2Count = dayData ? dayData.workingStaff.filter(s => GROUP_2.includes(s)).length : 0;
                        const offG2Count = 9 - workingG2Count;
                        const isExcessiveOff = offG2Count > 4;
                        const hasHoliday = d.holiday;
                        const dayOfWeek = d.obj.getDay();

                        // 🟣 SATURDAY (getDay === 6) & 🔴 SUNDAY (getDay === 0) STYLING REQUEST
                        let headerBg = "bg-emerald-700 text-white";
                        if (dayOfWeek === 6) {
                          headerBg = "bg-purple-600 text-white"; // Saturday Purple
                        } else if (dayOfWeek === 0) {
                          headerBg = "bg-red-600 text-white"; // Sunday Red
                        }

                        if (hasHoliday) {
                          if (hasHoliday.type === "public") {
                            headerBg = "bg-amber-500 text-white";
                          } else {
                            headerBg = "bg-rose-500 text-white";
                          }
                        }

                        let headerTitle = "";
                        if (hasHoliday) {
                          headerTitle += `[วันหยุด: ${hasHoliday.name} (${hasHoliday.type === "public" ? "นักขัตฤกษ์" : "บริษัท"})] `;
                        }
                        if (isExcessiveOff) {
                          headerTitle += `เตือนภัย: วันนี้มีทีมหยุดเกิน 4 คน! (หยุด ${offG2Count} คน จาก 9 คน)`;
                        }

                        return (
                          <th
                            key={d.date}
                            className={`p-1 sm:p-2 min-w-[42px] sm:min-w-[46px] lg:min-w-[50px] border-l border-emerald-500/20 text-center relative sticky top-0 z-30 ${headerBg} ${isExcessiveOff ? "ring-2 ring-rose-500 ring-inset" : ""}`}
                            title={headerTitle || undefined}
                          >
                            <div className="text-[8px] sm:text-[10px] font-black opacity-90 tracking-wider uppercase">
                              {d.dayName}
                            </div>
                            <div className="font-extrabold text-xs sm:text-base mt-0.5 relative inline-block">
                              {d.dateNum}
                              {isExcessiveOff && (
                                <span className="absolute -top-1 -right-2 text-rose-300 text-[10px] font-black animate-pulse" title="หยุดเกิน 4 คน!">
                                  ⚠️
                                </span>
                              )}
                              {hasHoliday && (
                                <span
                                  className="absolute bottom-[-2px] left-1/2 -translate-x-1/2 w-1.5 h-1.5 rounded-full bg-white animate-pulse"
                                  title={hasHoliday.name}
                                />
                              )}
                            </div>
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {(() => {
                      const activeSpotlightStaff = selectedStaffFilter || currentUser;
                      
                      // Support searching/filtering row names
                      const filteredDisplayOrder = DISPLAY_ORDER.filter(staff => {
                        if (!filterSearchQuery) return true;
                        return staff.toLowerCase().includes(filterSearchQuery.toLowerCase());
                      });

                      const listToRender = (isolateStaffRow && activeSpotlightStaff)
                        ? filteredDisplayOrder.filter(s => s === activeSpotlightStaff)
                        : filteredDisplayOrder;

                      if (listToRender.length === 0) {
                        return (
                          <tr>
                            <td colSpan={datesInMonth.length + 1} className="py-8 text-center text-sm font-semibold text-gray-400 italic bg-gray-50/50">
                              🔍 ไม่พบรายชื่อเจ้าหน้าที่ที่ตรงกับการค้นหา
                            </td>
                          </tr>
                        );
                      }

                      return listToRender.map((staff) => {
                        const idx = DISPLAY_ORDER.indexOf(staff);
                        const isStartOfGroup1 = idx === GROUP_2.length && !(isolateStaffRow && activeSpotlightStaff) && !filterSearchQuery;
                        
                        const isFocused = activeSpotlightStaff === staff;
                        const hasAnyFocus = !!activeSpotlightStaff;
                        
                        let rowClass = "transition-all duration-200 ";
                        if (isFocused) {
                          rowClass += "bg-emerald-100/70 hover:bg-emerald-150/80 shadow-[inset_4px_0_0_0_#10b981] font-bold ring-1 ring-emerald-500/20";
                        } else if (hasAnyFocus) {
                          rowClass += "opacity-35 hover:opacity-100 " + (idx % 2 === 0 ? "bg-white" : "bg-emerald-50/10");
                        } else {
                          rowClass += idx % 2 === 0 ? "bg-white" : "bg-emerald-50/30";
                        }

                        return (
                          <React.Fragment key={staff}>
                            {isStartOfGroup1 && (
                              <tr>
                                <td colSpan={datesInMonth.length + 1} className="bg-emerald-100/50 font-bold text-emerald-800 p-2 text-[11px] sm:text-xs text-center border-y border-emerald-200/60 shadow-inner">
                                  --- กลุ่มเจ้าหน้าที่หยุดเสาร์-อาทิตย์ ---
                                </td>
                              </tr>
                            )}
                            <tr className={rowClass}>
                              <td className="p-2 sm:p-3 border-b border-emerald-100 sticky left-0 bg-inherit shadow-[2px_0_4px_rgba(0,0,0,0.05)] z-10 font-bold text-emerald-950 border-r border-r-emerald-100 group text-xs sm:text-sm">
                                <div className="flex items-center justify-between">
                                  <span className="truncate pr-1">{staff}</span>
                                  {GROUP_1.includes(staff) && (
                                    <span className="text-[8px] sm:text-[9px] bg-purple-150 text-purple-700 font-extrabold px-1.5 py-0.5 rounded ml-1 whitespace-nowrap">หยุด ส.-อา.</span>
                                  )}
                                </div>
                              </td>
                              {datesInMonth.map((d) => {
                                const dayData = schedule[d.date];
                                const isWorking = dayData?.workingStaff.includes(staff);
                                const isVacation = dayData?.vacationStaff?.includes(staff);
                                const isDocInCharge = dayData?.docInCharge === staff;
                                const isPhone3551 = dayData?.phone3551 === staff;
                                const isPhone3552 = dayData?.phone3552 === staff;
                                const fireCode = dayData?.fireCodes[staff];

                                // Background column indicator highlights
                                let cellBg = "";
                                if (d.holiday) {
                                  cellBg = d.holiday.type === "public" ? "bg-amber-50/40" : "bg-rose-50/30";
                                } else if (d.obj.getDay() === 6) { // Sat columns
                                  cellBg = "bg-purple-50/15";
                                } else if (d.obj.getDay() === 0) { // Sun columns
                                  cellBg = "bg-red-50/15";
                                }

                                return (
                                  <td
                                    key={d.date}
                                    onClick={() => toggleCellState(d.date, staff)}
                                    title={d.holiday ? `${d.holiday.name} (${d.holiday.type === 'public' ? 'วันหยุดนักขัตฤกษ์' : 'วันหยุดบริษัท'})` : undefined}
                                    className={`p-0.5 sm:p-1 border-b border-l border-emerald-100/60 text-center relative cursor-pointer transition-colors ${cellBg} ${
                                      isAdmin ? "hover:bg-emerald-100/60" : "hover:bg-amber-100/40"
                                    }`}
                                  >
                                    {isVacation ? (
                                      <div className="flex items-center justify-center min-h-[36px] sm:min-h-[46px]">
                                        <span className="text-xs sm:text-sm font-black text-orange-500 select-none pointer-events-none">
                                          V
                                        </span>
                                      </div>
                                    ) : isWorking ? (
                                      <div className="flex flex-col items-center justify-center min-h-[36px] sm:min-h-[46px] gap-0.5 pointer-events-none">
                                        <Check
                                          strokeWidth={4.5}
                                          className="w-[12px] h-[12px] sm:w-[15px] sm:h-[15px] text-emerald-600"
                                        />
                                        {isDocInCharge && (
                                          <span className="text-[7.5px] sm:text-[8.5px] font-black text-blue-800 bg-blue-100 border border-blue-200 px-0.5 py-[0.5px] rounded leading-none shadow-sm select-none">
                                            Doc.
                                          </span>
                                        )}
                                        {isPhone3551 && (
                                          <span className="text-[7px] sm:text-[8px] font-black text-sky-800 bg-sky-100 border border-sky-200 px-0.5 py-[0.5px] rounded leading-none shadow-sm select-none flex items-center">
                                            <PhoneCall className="w-[7px] h-[7px] sm:w-[9px] sm:h-[9px] mr-0.5" /> 3551
                                          </span>
                                        )}
                                        {isPhone3552 && (
                                          <span className="text-[7px] sm:text-[8px] font-black text-indigo-800 bg-indigo-100 border border-indigo-200 px-0.5 py-[0.5px] rounded leading-none shadow-sm select-none flex items-center">
                                            <PhoneCall className="w-[7px] h-[7px] sm:w-[9px] sm:h-[9px] mr-0.5" /> 3552
                                          </span>
                                        )}
                                        {fireCode && (
                                          <span className="text-[8px] font-black text-red-600 bg-red-100 border border-red-200 px-0.5 py-[0.5px] rounded leading-none shadow-sm select-none">
                                            {fireCode}
                                          </span>
                                        )}
                                      </div>
                                    ) : (
                                      <div className="flex flex-col items-center justify-center min-h-[36px] sm:min-h-[46px] pointer-events-none">
                                        <X strokeWidth={3} className="w-[11px] h-[11px] sm:w-[14px] sm:h-[14px] text-red-400 opacity-60" />
                                      </div>
                                    )}
                                  </td>
                                );
                              })}
                            </tr>
                          </React.Fragment>
                        );
                      });
                    })()}
                  </tbody>
                </table>
              </div>
            </section>

            {/* DUAL COLUMN SYSTEM LAYOUT */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
              
              {/* ================= COLUMN 1 (LEFT - 60%): TODAY'S WORK & ROLES ================= */}
              <div className="lg:col-span-7 flex flex-col gap-6">
                
                {/* DAILY PERSONAL STATUS BOX */}
                {(() => {
                  const activeStaff = selectedStaffFilter || currentUser || "อุษา";
                  const dayData = schedule[selectedDate];
                  const isWorking = dayData?.workingStaff.includes(activeStaff);
                  const isVacation = dayData?.vacationStaff?.includes(activeStaff);
                  const isPhone1 = dayData?.phone3551 === activeStaff;
                  const isPhone2 = dayData?.phone3552 === activeStaff;
                  const isDoc = dayData?.docInCharge === activeStaff;
                  const fireCode = dayData?.fireCodes[activeStaff];

                  let cardColor = "bg-white border-gray-100 text-gray-800";
                  let badge = null;
                  let message = "";
                  let detail = "";

                  if (isWorking) {
                    cardColor = "bg-emerald-50/70 border-emerald-200 text-emerald-950";
                    badge = <span className="px-3 py-1 bg-emerald-600 text-white text-[10px] sm:text-xs font-black rounded-full shadow-sm">🟢 วันนี้คุณมีเวรปฏิบัติงาน</span>;
                    message = "วันนี้ท่านมีตารางเวรขึ้นปฏิบัติหน้าที่ในแผนก";
                    
                    const roles = [];
                    if (isPhone1) roles.push("📞 สายด่วนหลัก 3551");
                    if (isPhone2) roles.push("📞 สายด่วนรอง 3552");
                    if (isDoc) roles.push("📝 In-charge เอกสาร Doc");
                    if (fireCode) roles.push(`🚒 รหัสรับมืออัคคีภัย: ${fireCode}`);
                    
                    detail = roles.length > 0 ? `บทบาทของคุณวันนี้: ${roles.join(" | ")}` : "หน้าที่: เจ้าหน้าที่พิมพ์ผลเวรทั่วไป";
                  } else if (isVacation) {
                    cardColor = "bg-orange-50/70 border-orange-200 text-orange-950";
                    badge = <span className="px-3 py-1 bg-orange-600 text-white text-[10px] sm:text-xs font-black rounded-full shadow-sm">🌴 วันนี้คุณลาพักร้อน</span>;
                    message = "คุณอยู่ในสิทธิ์การลาพักร้อนประจำรอบเดือน";
                    detail = "ได้รับการยืนยันการลาในตารางระบบเรียบร้อยแล้ว";
                  } else {
                    cardColor = "bg-sky-50/50 border-sky-100 text-sky-950";
                    badge = <span className="px-3 py-1 bg-sky-600 text-white text-[10px] sm:text-xs font-black rounded-full shadow-sm">🔵 วันนี้คุณได้หยุดพักเวร</span>;
                    message = "วันนี้ท่านไม่มีเวร มีสิทธิ์หยุดปฏิบัติงานตามปกติ";
                    detail = "พักผ่อนเพื่อสุขภาพที่ดีของคุณ!";
                  }

                  return (
                    <div className={`rounded-2xl p-5 border shadow-xs transition-all ${cardColor} flex flex-col gap-3 relative text-left`}>
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold uppercase tracking-wider text-gray-400">สถานะเวรรายวันส่วนตัว</span>
                        {badge}
                      </div>
                      <h3 className="text-lg font-black">{message}</h3>
                      <p className="text-sm font-bold opacity-85">{detail}</p>
                    </div>
                  );
                })()}

                {/* MAIN KEY DUTIES TODAY */}
                {(() => {
                  const dayData = schedule[selectedDate] || {};
                  return (
                    <div className="bg-white rounded-2xl p-4 sm:p-5 border border-emerald-100/50 shadow-sm flex flex-col gap-4 text-left">
                      <h3 className="text-xs sm:text-sm font-bold text-gray-700 flex items-center gap-1.5 border-b border-gray-150 pb-2">
                        <User className="w-4 h-4 text-emerald-600" />
                        <span>ผู้รับผิดชอบเวรพิเศษของแผนกวันนี้</span>
                      </h3>
                      
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                        <div className="bg-sky-50/50 border border-sky-100 rounded-xl p-3 flex flex-col">
                          <span className="text-[10px] font-black text-sky-700 block">สายด่วนหลัก (3551)</span>
                          <span className="text-sm font-black text-sky-900 mt-1 truncate">{dayData.phone3551 || "ไม่มีผู้รับสาย"}</span>
                          <span className="text-[9px] text-gray-400 mt-0.5">รับเรื่องรายงานผลด่วนหลัก</span>
                        </div>
                        <div className="bg-indigo-50/50 border border-indigo-100 rounded-xl p-3 flex flex-col">
                          <span className="text-[10px] font-black text-indigo-700 block">สายด่วนรอง (3552)</span>
                          <span className="text-sm font-black text-indigo-900 mt-1 truncate">{dayData.phone3552 || "ไม่มีผู้รับสาย"}</span>
                          <span className="text-[9px] text-gray-400 mt-0.5">รับสายสำรองรายงาน</span>
                        </div>
                        <div className="bg-purple-50/50 border border-purple-100 rounded-xl p-3 flex flex-col">
                          <span className="text-[10px] font-black text-purple-700 block">In-charge เอกสาร Doc</span>
                          <span className="text-sm font-black text-purple-900 mt-1 truncate">{dayData.docInCharge || "ไม่มี"}</span>
                          <span className="text-[9px] text-gray-400 mt-0.5">บันทึกรายงานรับตรวจสอบ</span>
                        </div>
                      </div>
                    </div>
                  );
                })()}

                {/* TEAM MEMBERS ON DUTY TODAY */}
                {(() => {
                  const dayData = schedule[selectedDate] || { workingStaff: [], vacationStaff: [] };
                  const activeWorking = dayData.workingStaff || [];
                  const activeVacation = dayData.vacationStaff || [];
                  const activeOff = ALL_STAFF.filter(s => !activeWorking.includes(s) && !activeVacation.includes(s));

                  return (
                    <div className="bg-white rounded-2xl p-4 sm:p-5 border border-emerald-100/50 shadow-sm flex flex-col gap-5 text-left">
                      {/* Section 1: Working Staff */}
                      <div>
                        <h4 className="text-xs sm:text-sm font-bold text-emerald-800 flex items-center gap-1.5 border-b border-emerald-50 pb-2 mb-3">
                          <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
                          <span>เพื่อนร่วมทีมขึ้นเวรร่วมกันวันนี้ ({activeWorking.length} คน)</span>
                        </h4>
                        
                        {activeWorking.length === 0 ? (
                          <div className="py-4 text-center text-xs text-gray-400 italic">ไม่มีผู้ขึ้นเวรในวันนี้</div>
                        ) : (
                          <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                            {activeWorking.map(staff => {
                              const isPhone1 = dayData.phone3551 === staff;
                              const isPhone2 = dayData.phone3552 === staff;
                              const isDoc = dayData.docInCharge === staff;
                              const fireCode = dayData.fireCodes?.[staff];
                              
                              return (
                                <div key={staff} className="bg-emerald-50/30 border border-emerald-100/50 p-2.5 rounded-xl flex flex-col justify-between gap-1 shadow-2xs">
                                  <div className="flex items-center gap-1.5 justify-between">
                                    <span className="text-xs font-black text-emerald-950 truncate">{staff}</span>
                                    {GROUP_1.includes(staff) && <span className="text-[8px] bg-emerald-100 text-emerald-700 px-1 py-0.2 rounded font-black shrink-0">ส.-อา.</span>}
                                  </div>
                                  <div className="flex flex-wrap gap-0.5 mt-1">
                                    {isPhone1 && <span className="text-[8px] font-black bg-sky-100 text-sky-800 border border-sky-200 px-1 rounded-sm">📞 3551</span>}
                                    {isPhone2 && <span className="text-[8px] font-black bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 rounded-sm">📞 3552</span>}
                                    {isDoc && <span className="text-[8px] font-black bg-purple-100 text-purple-800 border border-purple-200 px-1 rounded-sm">📝 Doc</span>}
                                    {fireCode && <span className="text-[8px] font-black bg-red-100 text-red-800 border border-red-200 px-1 rounded-sm">🚒 {fireCode}</span>}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>

                      {/* Section 2: Off-duty and Vacation Staff */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-2 border-t border-gray-100">
                        {/* Off duty */}
                        <div>
                          <h4 className="text-xs font-bold text-sky-800 flex items-center gap-1.5 border-b border-sky-50 pb-1.5 mb-2">
                            <span className="w-2 h-2 rounded-full bg-sky-500"></span>
                            <span>ได้รับสิทธิ์หยุดวันนี้ ({activeOff.length} คน)</span>
                          </h4>
                          <div className="flex flex-wrap gap-1.5">
                            {activeOff.map(s => (
                              <span key={s} className="inline-flex text-[10px] font-bold bg-gray-50 text-gray-600 border border-gray-150 px-2 py-0.5 rounded-full">
                                {s}
                              </span>
                            ))}
                          </div>
                        </div>

                        {/* Vacation */}
                        <div>
                          <h4 className="text-xs font-bold text-orange-800 flex items-center gap-1.5 border-b border-orange-50 pb-1.5 mb-2">
                            <span className="w-2 h-2 rounded-full bg-orange-400"></span>
                            <span>ลาพักร้อนวันนี้ ({activeVacation.length} คน)</span>
                          </h4>
                          <div className="flex flex-wrap gap-1.5">
                            {activeVacation.length === 0 ? (
                              <span className="text-[10px] text-gray-400 italic">ไม่มีผู้ลาพักร้อนวันนี้</span>
                            ) : (
                              activeVacation.map(s => (
                                <span key={s} className="inline-flex text-[10px] font-bold bg-orange-50 text-orange-700 border border-orange-100 px-2 py-0.5 rounded-full">
                                  {s}
                                </span>
                              ))
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })()}

              </div>

              {/* ================= COLUMN 2 (RIGHT - 40%): MONTHLY PLANNER & SWAP ================= */}
              <div className="lg:col-span-5 flex flex-col gap-6">
                
                {/* INTERACTIVE COMPACT MONTH CALENDAR */}
                {(() => {
                  const activeStaff = selectedStaffFilter || currentUser || "อุษา";
                  const [yStr, mStr] = currentMonthStr.split('-');
                  const yearVal = parseInt(yStr);
                  const monthVal = parseInt(mStr) - 1;
                  const firstDayOfWeek = new Date(yearVal, monthVal, 1).getDay();
                  const padCells = Array(firstDayOfWeek).fill(null);
                  const allCalendarCells = [...padCells, ...datesInMonth];

                  return (
                    <div className="bg-white rounded-2xl p-4 border border-emerald-100/50 shadow-sm flex flex-col gap-3 text-left">
                      <div className="flex items-center justify-between border-b border-gray-100 pb-2">
                        <h3 className="text-xs sm:text-sm font-bold text-gray-700 flex items-center gap-1.5">
                          <CalendarDays className="w-4 h-4 text-emerald-600" />
                          <span>ปฏิทินเวรของ คุณ{activeStaff}</span>
                        </h3>
                        <span className="text-[10px] font-black text-emerald-700 bg-emerald-50 border border-emerald-100 px-2 py-0.5 rounded-md">
                          {new Date(`${currentMonthStr}-01`).toLocaleDateString("th-TH", { month: 'short', year: 'numeric' })}
                        </span>
                      </div>

                      {/* 7-column grid */}
                      <div className="grid grid-cols-7 gap-1 text-center">
                        {["อา", "จ", "อ", "พ", "พฤ", "ศ", "ส"].map((hd, index) => (
                          <span key={index} className={`text-[10px] font-black py-1 ${index === 0 || index === 6 ? 'text-rose-500' : 'text-gray-400'}`}>
                            {hd}
                          </span>
                        ))}

                        {allCalendarCells.map((d, index) => {
                          if (!d) return <div key={`pad-${index}`} className="p-1" />;
                          
                          const isSelected = selectedDate === d.date;
                          const hasHoliday = d.holiday;
                          const isWorking = schedule[d.date]?.workingStaff.includes(activeStaff);
                          const isVacation = schedule[d.date]?.vacationStaff?.includes(activeStaff);
                          const isPhone1 = schedule[d.date]?.phone3551 === activeStaff;
                          const isPhone2 = schedule[d.date]?.phone3552 === activeStaff;
                          const isDoc = schedule[d.date]?.docInCharge === activeStaff;

                          let bgClass = "bg-gray-50/50 hover:bg-emerald-50 text-gray-700 border border-gray-100/40";
                          let dotClass = null;

                          if (isSelected) {
                            bgClass = "bg-emerald-600 text-white font-black scale-105 shadow-sm ring-2 ring-emerald-500/50 z-10";
                          } else if (isWorking) {
                            bgClass = "bg-emerald-50 hover:bg-emerald-100/80 text-emerald-950 border border-emerald-200 font-bold";
                            dotClass = "bg-emerald-500";
                          } else if (isVacation) {
                            bgClass = "bg-orange-50 hover:bg-orange-100/80 text-orange-950 border border-orange-200 font-bold";
                            dotClass = "bg-orange-500";
                          } else if (hasHoliday) {
                            bgClass = hasHoliday.type === "public" ? "bg-amber-50 hover:bg-amber-100/60 border border-amber-200 text-amber-900" : "bg-rose-50 hover:bg-rose-100/60 border border-rose-200 text-rose-900";
                          }

                          return (
                            <button
                              key={d.date}
                              type="button"
                              onClick={() => setSelectedDate(d.date)}
                              className={`p-1.5 rounded-xl text-xs flex flex-col items-center justify-between min-h-[44px] cursor-pointer transition-all active:scale-95 ${bgClass}`}
                              title={hasHoliday ? `${hasHoliday.name} (${hasHoliday.type === 'public' ? 'วันหยุดนักขัตฤกษ์' : 'วันหยุดบริษัท'})` : ""}
                            >
                              <span className="text-[10px] font-black">{d.dateNum}</span>
                              <div className="flex gap-0.5 mt-0.5 items-center justify-center min-h-[8px]">
                                {isWorking && (
                                  <span className={`w-1 h-1 rounded-full ${isSelected ? 'bg-white' : dotClass || 'bg-emerald-500'}`} />
                                )}
                                {(isPhone1 || isPhone2) && (
                                  <span className="text-[6px] leading-none shrink-0">📞</span>
                                )}
                                {isDoc && (
                                  <span className="text-[6px] leading-none shrink-0">📝</span>
                                )}
                                {isVacation && (
                                  <span className="text-[6px] leading-none shrink-0">🌴</span>
                                )}
                              </div>
                            </button>
                          );
                        })}
                      </div>

                      {/* Legend */}
                      <div className="flex flex-wrap items-center justify-center gap-x-3 gap-y-1.5 border-t border-gray-100 pt-2 text-[9px] text-gray-500 font-semibold mt-1">
                        <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />ขึ้นเวร (Work)</span>
                        <span className="flex items-center gap-1"><span>📞 / 📝</span>เวรสายด่วน/เอกสาร</span>
                        <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-orange-400" />ลาพักร้อน (Leave)</span>
                        <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-amber-400" />วันหยุดประจำแผนก</span>
                      </div>
                    </div>
                  );
                })()}

                {/* QUICK REQUEST FORM */}
                <div className="bg-white rounded-2xl p-4 sm:p-5 border border-emerald-100/50 shadow-sm flex flex-col gap-4 text-left">
                  <div className="border-b border-gray-100 pb-2">
                    <h3 className="text-xs sm:text-sm font-bold text-gray-700 flex items-center gap-1.5">
                      <RefreshCcw className="w-4 h-4 text-emerald-600" />
                      <span>ยื่นเปลี่ยนเวร/แจ้งลางานด่วน</span>
                    </h3>
                  </div>

                  <form onSubmit={handleCreateShiftRequest} className="flex flex-col gap-3 text-left">
                    <div>
                      <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-wider mb-1">1. ผู้ส่งคำขอ (ฉันคือ...)</label>
                      <div className="w-full px-3 py-1.5 bg-gray-50 border border-gray-200 rounded-xl text-xs text-gray-800 font-black flex items-center justify-between">
                        <span className="flex items-center gap-1.5">
                          <User className="w-3.5 h-3.5 text-emerald-600" />
                          <span>คุณ{reqRequester}</span>
                        </span>
                        <span className="text-[9px] bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded font-black">ยืนยันตัวตนแล้ว</span>
                      </div>
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-wider mb-1">2. ประเภทคำขอ</label>
                      <div className="grid grid-cols-3 gap-1 bg-gray-50 p-1 rounded-xl border border-gray-200">
                        <button
                          type="button"
                          onClick={() => setReqType("swap")}
                          className={`py-1 rounded-lg text-center text-xs font-black transition-all cursor-pointer ${
                            reqType === "swap" ? "bg-emerald-600 text-white shadow-xs" : "text-gray-500 hover:bg-gray-100"
                          }`}
                        >
                          🤝 สลับเวร
                        </button>
                        <button
                          type="button"
                          onClick={() => setReqType("cover")}
                          className={`py-1 rounded-lg text-center text-xs font-black transition-all cursor-pointer ${
                            reqType === "cover" ? "bg-emerald-600 text-white shadow-xs" : "text-gray-500 hover:bg-gray-100"
                          }`}
                        >
                          🙋‍♀️ ขึ้นแทน
                        </button>
                        <button
                          type="button"
                          onClick={() => setReqType("leave")}
                          className={`py-1 rounded-lg text-center text-xs font-black transition-all cursor-pointer ${
                            reqType === "leave" ? "bg-emerald-600 text-white shadow-xs" : "text-gray-500 hover:bg-gray-100"
                          }`}
                        >
                          🌴 ลาพักร้อน
                        </button>
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-wider mb-1">3. วันที่แจ้งทำ</label>
                        <input
                          type="date"
                          value={reqDate}
                          onChange={(e) => setReqDate(e.target.value)}
                          required
                          className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded-xl text-xs font-black text-gray-700 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                        />
                      </div>

                      {reqType === "swap" && (
                        <div>
                          <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-wider mb-1">4. วันสลับคืน</label>
                          <input
                            type="date"
                            value={reqTargetDate}
                            onChange={(e) => setReqTargetDate(e.target.value)}
                            required
                            className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded-xl text-xs font-black text-gray-700 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                      )}
                    </div>

                    {(reqType === "swap" || reqType === "cover") && (
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-wider mb-1">
                          {reqType === "swap" ? "เลือกคนสลับด้วย" : "เลือกคนที่จะขึ้นเวรแทน"}
                        </label>
                        <select
                          value={reqTargetStaff}
                          onChange={(e) => setReqTargetStaff(e.target.value)}
                          className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded-xl text-xs font-black text-gray-700 focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                        >
                          {ALL_STAFF.filter(s => s !== reqRequester).map(s => (
                            <option key={s} value={s}>คุณ{s}</option>
                          ))}
                        </select>
                      </div>
                    )}

                    <div>
                      <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-wider mb-1">หมายเหตุ / เหตุผลเพิ่มเติม</label>
                      <input
                        type="text"
                        placeholder="เช่น ติดธุระครอบครัว, พักผ่อนส่วนตัว"
                        value={reqNote}
                        onChange={(e) => setReqNote(e.target.value)}
                        className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded-xl text-xs focus:outline-none focus:ring-1 focus:ring-emerald-500 text-gray-700"
                      />
                    </div>

                    <button
                      type="submit"
                      className="w-full py-2 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white font-black rounded-xl text-xs cursor-pointer shadow-xs transition-all mt-1"
                    >
                      ส่งคำขอไปยังระบบ 🚀
                    </button>
                  </form>
                </div>

                {/* CURRENT ACTIVE REQUESTS */}
                <div className="bg-white rounded-2xl p-4 sm:p-5 border border-emerald-100/50 shadow-sm flex flex-col gap-3 text-left">
                  <div className="border-b border-gray-100 pb-2 flex items-center justify-between">
                    <h3 className="text-xs sm:text-sm font-bold text-gray-700 flex items-center gap-1.5">
                      <Clock className="w-4 h-4 text-emerald-600" />
                      <span>รายการขอเปลี่ยนเวรล่าสุด</span>
                    </h3>
                    <div className="flex items-center gap-1 bg-gray-50 p-0.5 rounded-lg border border-gray-100">
                      <button
                        onClick={() => setActiveReqTab("pending")}
                        type="button"
                        className={`px-2 py-0.5 rounded text-[9px] font-black cursor-pointer transition-all ${
                          activeReqTab === "pending" ? "bg-white text-gray-800 shadow-2xs" : "text-gray-400 hover:text-gray-600"
                        }`}
                      >
                        รอนุมัติ ({shiftRequests.filter(r => r.status === "pending").length})
                      </button>
                      <button
                        onClick={() => setActiveReqTab("all")}
                        type="button"
                        className={`px-2 py-0.5 rounded text-[9px] font-black cursor-pointer transition-all ${
                          activeReqTab === "all" ? "bg-white text-gray-800 shadow-2xs" : "text-gray-400 hover:text-gray-600"
                        }`}
                      >
                        ทั้งหมด ({shiftRequests.length})
                      </button>
                    </div>
                  </div>

                  <div className="flex flex-col gap-2 max-h-[290px] overflow-y-auto custom-scrollbar pr-1">
                    {(() => {
                      const filtered = activeReqTab === "pending"
                        ? shiftRequests.filter(r => r.status === "pending")
                        : shiftRequests;

                      if (filtered.length === 0) {
                        return (
                          <div className="py-8 text-center text-[11px] text-gray-400 italic bg-gray-50/50 rounded-xl border border-dashed border-gray-150">
                            ไม่มีรายการคำขอในขณะนี้
                          </div>
                        );
                      }

                      return filtered.slice(0, 8).map(req => {
                        const isPending = req.status === "pending";
                        const isApproved = req.status === "approved";
                        const isRejected = req.status === "rejected";
                        const formattedReqDate = new Date(req.date).toLocaleDateString("th-TH", { day: 'numeric', month: 'short' });
                        const formattedTargetDate = req.targetDate ? new Date(req.targetDate).toLocaleDateString("th-TH", { day: 'numeric', month: 'short' }) : "";

                        return (
                          <div
                            key={req.id}
                            className={`p-3 rounded-xl border transition-all text-left flex flex-col gap-1.5 ${
                              isApproved
                                ? "bg-emerald-50/20 border-emerald-100"
                                : isRejected
                                ? "bg-red-50/10 border-red-100/50"
                                : "bg-gray-50/50 border-gray-200"
                            }`}
                          >
                            <div className="flex items-center justify-between gap-1.5">
                              <span className={`text-[9px] font-black px-1.5 py-0.5 rounded-full ${
                                req.type === "swap" ? "bg-indigo-100 text-indigo-800" : req.type === "cover" ? "bg-sky-100 text-sky-800" : "bg-orange-100 text-orange-800"
                              }`}>
                                {req.type === "swap" ? "🤝 สลับเวร" : req.type === "cover" ? "🙋‍♀️ ขึ้นแทน" : "🌴 ลาพักร้อน"}
                              </span>
                              
                              <span className={`text-[8.5px] font-black px-1.5 py-0.5 rounded-md ${
                                isApproved ? "bg-emerald-100 text-emerald-800" : isRejected ? "bg-red-100 text-red-800" : "bg-amber-100 text-amber-800 animate-pulse"
                              }`}>
                                {isApproved ? "อนุมัติสำเร็จ" : isRejected ? "ปฏิเสธ" : "รอการตรวจ"}
                              </span>
                            </div>

                            <div className="text-[11px] font-semibold text-gray-700 leading-normal">
                              {req.type === "swap" ? (
                                <span>คุณ <strong>{req.requester}</strong> ({formattedReqDate}) สลับกับ คุณ <strong>{req.targetStaff}</strong> ({formattedTargetDate})</span>
                              ) : req.type === "cover" ? (
                                <span>คุณ <strong>{req.requester}</strong> ขอให้ คุณ <strong>{req.targetStaff}</strong> ขึ้นเวรแทน ({formattedReqDate})</span>
                              ) : (
                                <span>คุณ <strong>{req.requester}</strong> ขอลาพักร้อน ({formattedReqDate})</span>
                              )}
                              
                              {req.note && (
                                <span className="block text-[10px] text-gray-400 mt-0.5 italic font-normal">💬 บันทึก: "{req.note}"</span>
                              )}
                            </div>

                            {isPending && (
                              <div className="flex gap-1 justify-end mt-1 border-t border-gray-200/40 pt-2">
                                {req.targetStaff === currentUser ? (
                                  <button
                                    onClick={() => {
                                      setIsAdmin(true);
                                      handleApproveShiftRequest(req);
                                      setIsAdmin(false);
                                    }}
                                    type="button"
                                    className="px-3 py-1 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-[10px] font-black cursor-pointer active:scale-95 animate-pulse"
                                  >
                                    🤝 ตอบตกลงคำขอสลับเวร
                                  </button>
                                ) : isAdmin ? (
                                  <div className="flex gap-1">
                                    <button
                                      onClick={() => handleRejectShiftRequest(req.id, req.requester, req.date)}
                                      type="button"
                                      className="px-2 py-1 border border-red-200 text-red-600 hover:bg-red-50 rounded-lg text-[10px] font-bold transition-all cursor-pointer"
                                    >
                                      ปฏิเสธ
                                    </button>
                                    <button
                                      onClick={() => handleApproveShiftRequest(req)}
                                      type="button"
                                      className="px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[10px] font-black transition-all cursor-pointer"
                                    >
                                      อนุมัติข้อตกลง
                                    </button>
                                  </div>
                                ) : (
                                  <button
                                    onClick={() => {
                                      setIsAdmin(true);
                                      handleApproveShiftRequest(req);
                                    }}
                                    type="button"
                                    className="px-2 py-1 bg-amber-500 hover:bg-amber-600 text-white rounded-lg text-[9px] font-black transition-all cursor-pointer"
                                  >
                                    จำลองกดยอมรับคำขอ (Simulate)
                                  </button>
                                )}
                              </div>
                            )}
                          </div>
                        );
                      });
                    })()}
                  </div>
                </div>

                {/* HOLIDAYS OF THE MONTH */}
                {(() => {
                  const monthlyHolidays = datesInMonth.filter(d => d.holiday);
                  return (
                    <div className="bg-white rounded-2xl p-4 sm:p-5 border border-emerald-100/50 shadow-sm flex flex-col gap-3 text-left">
                      <h3 className="text-xs sm:text-sm font-bold text-gray-700 flex items-center gap-1.5 border-b border-gray-100 pb-2">
                        <CalendarRange className="w-4 h-4 text-emerald-600" />
                        <span>วันหยุดนักขัตฤกษ์และวันหยุดแผนกเดือนนี้</span>
                      </h3>
                      
                      {monthlyHolidays.length === 0 ? (
                        <div className="py-3 text-center text-xs text-gray-400 italic">ไม่มีวันหยุดประจำเดือนนี้</div>
                      ) : (
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                          {monthlyHolidays.map(d => {
                            const h = d.holiday!;
                            return (
                              <div
                                key={d.date}
                                className={`p-2.5 rounded-xl border flex flex-col gap-0.5 text-left ${
                                  h.type === "public" ? "bg-amber-50/40 border-amber-100 text-amber-900" : "bg-rose-50/40 border-rose-100 text-rose-900"
                                }`}
                              >
                                <span className="text-[8px] font-black uppercase opacity-75">วันที่ {d.dateNum} ({d.dayName})</span>
                                <span className="text-xs font-black truncate">{h.name}</span>
                                <span className="text-[8px] font-medium opacity-85">{h.type === "public" ? "วันหยุดนักขัตฤกษ์" : "วันหยุดพิเศษของบริษัท"}</span>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })()}

              </div>

            </div>
          </div>
        ) : (
          /* ======================================================= */
          /*   2. ORIGINAL ADMIN CONSOLE                             */
          /* ======================================================= */
          <>
            {/* Toggle View Mode for Mobile vs Desktop */}
            <div className="flex bg-white/90 backdrop-blur border border-emerald-100/50 p-1 rounded-2xl shadow-sm gap-1 w-full max-w-lg mx-auto">
          <button
            onClick={() => setViewMode("daily")}
            className={`flex-1 flex items-center justify-center gap-2 py-2 sm:py-2.5 rounded-xl text-xs sm:text-sm font-semibold transition-all ${
              viewMode === "daily"
                ? "bg-emerald-600 text-white shadow"
                : "text-gray-600 hover:bg-gray-50 hover:text-gray-800"
            }`}
          >
            <Phone className="w-4 h-4 text-inherit" />
            <span>📱 ดูเฉพาะวัน & สรุปยอดเวร (สไตล์มือถือ)</span>
          </button>
          <button
            onClick={() => setViewMode("monthly")}
            className={`flex-1 flex items-center justify-center gap-2 py-2 sm:py-2.5 rounded-xl text-xs sm:text-sm font-semibold transition-all ${
              viewMode === "monthly"
                ? "bg-emerald-600 text-white shadow"
                : "text-gray-600 hover:bg-gray-50 hover:text-gray-800"
            }`}
          >
            <CalendarDays className="w-4 h-4 text-inherit" />
            <span>💻 ตารางเต็มทั้งเดือน</span>
          </button>
        </div>

        {/* 🌟 ระบบสืบค้นและวิเคราะห์เจาะลึกตารางเวรส่วนบุคคล (Personal Schedule & Workload Spotlight) */}
        <section className="w-full max-w-7xl mx-auto bg-white rounded-2xl shadow-xs border border-emerald-100 p-4 sm:p-5 flex flex-col gap-4">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 border-b border-gray-100 pb-3">
            <div>
              <h2 className="text-base sm:text-lg font-bold text-emerald-800 flex items-center gap-2">
                <Sparkles className="w-5 h-5 text-amber-500 animate-pulse shrink-0" />
                <span>เจาะลึก & ดูตารางเวรรายบุคคล (Personal Schedule Spotlight)</span>
              </h2>
              <p className="text-gray-500 text-xs mt-0.5 font-medium">
                เลือกชื่อของตนเองหรือเพื่อนร่วมงานเพื่อตรวจสอบสรุปวันปฏิบัติงานทั้งหมดอย่างรวดเร็ว โดยไม่ต้องไล่ดูตารางชีทใหญ่
              </p>
            </div>
            
            {/* Quick search input */}
            <div className="relative w-full md:w-64">
              <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-gray-400 pointer-events-none">
                <Search className="w-4 h-4" />
              </span>
              <input
                type="text"
                placeholder="ค้นหาชื่อเจ้าหน้าที่..."
                value={filterSearchQuery}
                onChange={(e) => setFilterSearchQuery(e.target.value)}
                className="w-full pl-9 pr-8 py-2 bg-gray-50 border border-gray-200 focus:bg-white text-xs rounded-xl focus:ring-2 focus:ring-emerald-500 focus:outline-none font-semibold text-gray-700 transition-all"
              />
              {filterSearchQuery && (
                <button
                  onClick={() => setFilterSearchQuery("")}
                  className="absolute inset-y-0 right-0 pr-2.5 flex items-center text-gray-400 hover:text-gray-600 cursor-pointer"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              )}
            </div>
          </div>

          {/* Horizontal scroll of staff members */}
          <div className="flex items-center gap-1.5 overflow-x-auto py-1 custom-scrollbar pb-2">
            <button
              onClick={() => {
                setSelectedStaffFilter(null);
                setIsolateStaffRow(false);
              }}
              className={`px-3 py-2 rounded-xl text-xs font-bold transition-all shrink-0 cursor-pointer flex items-center gap-1 border ${
                selectedStaffFilter === null
                  ? "bg-emerald-600 text-white border-emerald-600 shadow-sm"
                  : "bg-gray-50 text-gray-600 border-gray-200 hover:bg-gray-100"
              }`}
            >
              <Filter className="w-3 h-3" />
              <span>แสดงทุกคน ({ALL_STAFF.length} คน)</span>
            </button>

            {ALL_STAFF.filter(s => s.toLowerCase().includes(filterSearchQuery.toLowerCase())).map(staff => {
              const isSelected = selectedStaffFilter === staff;
              const shiftCount = datesInMonth.filter(d => schedule[d.date]?.workingStaff.includes(staff)).length;
              
              return (
                <button
                  key={staff}
                  onClick={() => {
                    setSelectedStaffFilter(staff);
                  }}
                  className={`px-3 py-2 rounded-xl text-xs font-bold transition-all shrink-0 cursor-pointer flex items-center gap-1.5 border ${
                    isSelected
                      ? "bg-emerald-100 text-emerald-950 border-emerald-300 font-bold ring-2 ring-emerald-500/20"
                      : "bg-white text-gray-700 border-gray-200 hover:bg-emerald-50/40 hover:border-emerald-200"
                  }`}
                >
                  <span className={`w-2 h-2 rounded-full ${GROUP_1.includes(staff) ? 'bg-amber-400 animate-pulse' : 'bg-emerald-400'}`} />
                  <span>{staff}</span>
                  <span className={`px-1.5 py-0.2 rounded text-[9px] font-black ${isSelected ? 'bg-emerald-600 text-white' : 'bg-gray-100 text-gray-500'}`}>
                    {shiftCount} เวร
                  </span>
                </button>
              );
            })}
          </div>

          {/* Individual Statistics Breakdown Board */}
          {(() => {
            const activeStaff = selectedStaffFilter || currentUser;
            if (!activeStaff) {
              return (
                <div className="bg-emerald-50/25 border border-emerald-100/50 rounded-2xl p-4 text-center py-6 flex flex-col items-center justify-center">
                  <User className="w-8 h-8 text-emerald-600 opacity-60 mb-2" />
                  <p className="text-xs font-bold text-emerald-900">ดูตารางเวรส่วนตัวแบบเร็ว</p>
                  <p className="text-[10px] text-gray-500 mt-1 max-w-md leading-relaxed">
                    คลิกเลือกชื่อของคุณหรือเพื่อนร่วมงานด้านบน เพื่อตรวจสอบสรุปภาระงาน, จำนวนเวรรับสาย และคลิกทางลัดไปกรอกขอสลับสับเปลี่ยนเวรได้ทันทีโดยไม่ต้องจำวันเอง
                  </p>
                </div>
              );
            }

            const stats = getStaffStats(activeStaff);
            const isGroup1 = GROUP_1.includes(activeStaff);

            return (
              <div className="bg-gradient-to-r from-emerald-50/20 to-white border border-emerald-100/40 rounded-2xl p-4 sm:p-5 flex flex-col gap-4">
                {/* Profile header & controls */}
                <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 border-b border-gray-100 pb-4">
                  <div className="flex items-center gap-3">
                    <div className="w-12 h-12 rounded-full bg-gradient-to-tr from-emerald-600 to-emerald-500 text-white font-black text-lg flex items-center justify-center shadow-md shrink-0">
                      {activeStaff.substring(0, 2)}
                    </div>
                    <div>
                      <div className="flex items-center gap-2 flex-wrap">
                        <h3 className="text-base sm:text-lg font-black text-emerald-950">ข้อมูลตารางเวร: คุณ{activeStaff}</h3>
                        <span className="text-[10px] bg-emerald-100 text-emerald-800 border border-emerald-200 px-2 py-0.5 rounded-full font-bold">
                          {isGroup1 ? "กลุ่มหยุด ส.-อา." : "กลุ่มเวรหมุนเวียนปกติ"}
                        </span>
                        {currentUser === activeStaff && (
                          <span className="text-[10px] bg-sky-600 text-white px-2 py-0.5 rounded-full font-bold flex items-center gap-1">
                            <span className="w-1.5 h-1.5 rounded-full bg-white animate-pulse" /> คุณล็อกอินอยู่
                          </span>
                        )}
                      </div>
                      <p className="text-xs text-gray-500 mt-1.5 font-bold flex flex-wrap items-center gap-1.5">
                        <span>📅 รอบเดือนตารางเวร:</span>
                        <strong className="text-emerald-700 bg-emerald-50 border border-emerald-100/60 px-2 py-0.5 rounded-md text-xs font-black">
                          {new Date(`${currentMonthStr}-01`).toLocaleDateString("th-TH", { month: 'long', year: 'numeric' })}
                        </strong>
                        <span className="text-gray-300 hidden sm:inline">|</span>
                        <span className="text-gray-400 font-medium">วันที่ปัจจุบัน: {new Date().toLocaleDateString("th-TH", { day: 'numeric', month: 'long', year: 'numeric' })}</span>
                      </p>
                    </div>
                  </div>

                  {/* Quick toggle actions */}
                  <div className="flex flex-wrap items-center gap-2 w-full lg:w-auto">
                    <button
                      onClick={() => {
                        setIsolateStaffRow(!isolateStaffRow);
                        setViewMode("monthly");
                        showToast(isolateStaffRow ? "แสดงรายชื่อทุกคนในทีมตามปกติ" : `เปิดโหมดเจาะลึก: แสดงเฉพาะแถวตารางของคุณ ${activeStaff} เพื่อความชัดเจนสูงสุด!`, false, new Date().toLocaleDateString("th-TH"));
                      }}
                      className={`flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer border shadow-xs ${
                        isolateStaffRow && viewMode === "monthly"
                          ? "bg-emerald-700 text-white border-emerald-700 shadow"
                          : "bg-white text-emerald-800 border-emerald-200 hover:bg-emerald-50"
                      }`}
                    >
                      <Filter className="w-3.5 h-3.5" />
                      <span>{isolateStaffRow && viewMode === "monthly" ? "👁️ แสดงทุกคน (Show All)" : `🔍 ซ่องแถวอื่น (ส่องเฉพาะแถว ${activeStaff})`}</span>
                    </button>

                    <div className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 bg-amber-50 border border-amber-200 text-amber-800 font-bold px-4 py-2 rounded-xl text-xs shadow-xs">
                      <Sparkles className="w-3.5 h-3.5 text-amber-500" />
                      <span>👉 คลิกเลือกช่องวันที่บนปฏิทินเพื่อขอหยุด/เวรได้เลย!</span>
                    </div>
                  </div>
                </div>

                {/* 🏷️ Tab Switcher สำหรับเจาะลึกรายบุคคล */}
                <div className="flex border-b border-gray-100 pb-1 gap-1.5 overflow-x-auto custom-scrollbar">
                  <button
                    onClick={() => setSpotlightTab("calendar")}
                    className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold transition-all cursor-pointer border shrink-0 ${
                      spotlightTab === "calendar"
                        ? "bg-emerald-600 text-white border-emerald-600 shadow-sm"
                        : "bg-white text-gray-600 border-gray-200 hover:bg-emerald-50/30"
                    }`}
                  >
                    <CalendarDays className="w-4 h-4 text-inherit" />
                    <span>📅 ตารางปฏิทินขอเวรล่วงหน้า</span>
                  </button>
                  <button
                    onClick={() => setSpotlightTab("stats")}
                    className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold transition-all cursor-pointer border shrink-0 ${
                      spotlightTab === "stats"
                        ? "bg-emerald-600 text-white border-emerald-600 shadow-sm"
                        : "bg-white text-gray-600 border-gray-200 hover:bg-emerald-50/30"
                    }`}
                  >
                    <TrendingUp className="w-4 h-4 text-inherit" />
                    <span>📊 สรุปสถิติภาระงาน & สิทธิ์เวร</span>
                  </button>
                </div>

                {/* TAB 1: INTERACTIVE REQUEST CALENDAR */}
                {spotlightTab === "calendar" && (() => {
                  const [yStr, mStr] = currentMonthStr.split('-');
                  const yearVal = parseInt(yStr);
                  const monthVal = parseInt(mStr) - 1;
                  const firstDayOfWeek = new Date(yearVal, monthVal, 1).getDay(); // 0 = Sun, 1 = Mon, ..., 6 = Sat
                  const padCells = Array(firstDayOfWeek).fill(null);
                  const allCalendarCells = [...padCells, ...datesInMonth];

                  return (
                    <div className="flex flex-col gap-4">
                      {/* สถานะระบบบันทึกตารางเวรอัตโนมัติ */}
                      <div className="bg-emerald-50 border border-emerald-150 p-3.5 rounded-2xl flex flex-col sm:flex-row items-center justify-between gap-3 shadow-xs">
                        <div className="flex items-center gap-2 text-xs sm:text-sm font-bold text-emerald-800">
                          <span className="text-base shrink-0">🟢</span>
                          <span>ระบบบันทึกข้อมูลและอัปเดตตารางเวรทั้งหมดของทุกคนโดยอัตโนมัติในทันทีเมื่อมีการแก้ไข</span>
                        </div>
                        <div className="bg-white text-emerald-800 border border-emerald-200 font-extrabold text-[10px] sm:text-xs px-3 py-1 rounded-xl flex items-center gap-1.5 shrink-0 shadow-2xs">
                          <span className="w-2 h-2 bg-emerald-500 rounded-full animate-ping"></span>
                          <span>เปิดใช้งานระบบบันทึกอัตโนมัติ</span>
                        </div>
                      </div>

                      {/* แนะนำวิธีใช้งาน */}
                      <div className="bg-emerald-50/50 border border-emerald-100 p-3.5 rounded-xl text-xs text-emerald-800 font-medium flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div className="flex items-start gap-1.5">
                          <span className="shrink-0">💡</span>
                          <span>
                            <strong>วิธีการเลือกวันหยุดและวันลาพักร้อนบนปฏิทิน:</strong> คลิกเลือกช่องวันที่ในปฏิทินด้านล่างของคุณ เพื่อสลับปรับเปลี่ยนเป็นวันหยุด (OFF), ลาพักร้อน (VAC) หรือขอขึ้นเวร (WORK) ได้ทันที โดยระบบจะทำการบันทึกข้อมูลและอัปเดตทันทีอัตโนมัติ!
                          </span>
                        </div>
                      </div>

                      {/* ตารางปฏิทินรายบุคคล */}
                      <div className="bg-white border border-gray-150 rounded-2xl p-3 sm:p-5 shadow-2xs">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-gray-100 pb-3.5 mb-4">
                          <div>
                            <h3 className="text-sm sm:text-base font-black text-gray-800 flex flex-wrap items-center gap-1.5">
                              <CalendarDays className="w-5 h-5 text-emerald-600" />
                              <span>ปฏิทินเวรของ: <strong className="text-emerald-700 underline font-extrabold">คุณ{activeStaff}</strong></span>
                              <span className="bg-emerald-50 text-emerald-800 border border-emerald-150 text-xs px-2 py-0.5 rounded-lg font-extrabold">
                                ประจำรอบเดือน: {new Date(`${currentMonthStr}-01`).toLocaleDateString("th-TH", { month: 'long', year: 'numeric' })}
                              </span>
                            </h3>
                            <p className="text-gray-400 text-[10px] sm:text-xs">
                              คลิกช่องวันที่ที่ต้องการเพื่อปรับปรุงเป็นวันหยุด วันลา หรือขอปฏิบัติงาน (บันทึกอัตโนมัติ)
                            </p>
                          </div>
                          
                          <div className="flex items-center gap-2">
                            <div className="bg-emerald-50 border border-emerald-150 text-emerald-800 font-extrabold text-xs px-3.5 py-1.5 rounded-xl flex items-center gap-1.5">
                              <span className="w-2 h-2 bg-emerald-500 rounded-full"></span>
                              <span>บันทึกอัตโนมัติแล้ว 🟢</span>
                            </div>
                          </div>
                        </div>
                        <div className="grid grid-cols-7 gap-1 sm:gap-2 mb-2">
                          {["อา.", "จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส."].map((dayName, idx) => {
                            let dayBg = "bg-gray-50 text-gray-600";
                            if (idx === 0) { // Sunday
                              dayBg = "bg-red-50 text-red-600";
                            } else if (idx === 6) { // Saturday
                              dayBg = "bg-purple-50 text-purple-600";
                            }
                            return (
                              <div key={dayName} className={`text-center font-bold text-[10px] sm:text-xs py-1.5 rounded-lg ${dayBg}`}>
                                {dayName}
                              </div>
                            );
                          })}
                        </div>

                        <div className="grid grid-cols-7 gap-1.5 sm:gap-2">
                          {allCalendarCells.map((cell, index) => {
                            if (!cell) {
                              return <div key={`empty-${index}`} className="bg-gray-50/5 border border-transparent rounded-xl min-h-[70px] sm:min-h-[90px]" />;
                            }

                            const dayData = schedule[cell.date];
                            const isWorking = dayData?.workingStaff.includes(activeStaff);
                            const isVacation = dayData?.vacationStaff?.includes(activeStaff);
                            const isDoc = dayData?.docInCharge === activeStaff;
                            const isP1 = dayData?.phone3551 === activeStaff;
                            const isP2 = dayData?.phone3552 === activeStaff;
                            const isWeekend = cell.isWeekend;

                            let bgClass = "bg-white border-gray-200 hover:border-emerald-300";
                            let statusLabel = "วันหยุด";
                            let statusColor = "text-gray-400";

                            const hasHoliday = cell.holiday;

                            if (isVacation) {
                              bgClass = "bg-orange-50/60 border-orange-200 text-orange-950 hover:bg-orange-100/60 hover:border-orange-300";
                              statusLabel = "พักร้อน 🌴";
                              statusColor = "text-orange-600 font-bold";
                            } else if (isWorking) {
                              bgClass = "bg-emerald-50/60 border-emerald-200 text-emerald-950 hover:bg-emerald-100/50 hover:border-emerald-300 font-bold";
                              let roleText = "ขึ้นเวร 🟢";
                              if (isDoc || isP1 || isP2) {
                                const roles = [];
                                if (isDoc) roles.push("Doc");
                                if (isP1) roles.push("3551");
                                if (isP2) roles.push("3552");
                                roleText = `เวร (${roles.join("/")}) 🟢`;
                              }
                              statusLabel = roleText;
                              statusColor = "text-emerald-700";
                            } else if (hasHoliday) {
                              if (hasHoliday.type === "public") {
                                bgClass = "bg-amber-50/60 border-amber-200 text-amber-950 hover:bg-amber-100/50 hover:border-amber-300";
                                statusLabel = "วันหยุดนักขัตฤกษ์ 🎉";
                                statusColor = "text-amber-700 font-bold";
                              } else {
                                bgClass = "bg-rose-50/50 border-rose-200 text-rose-950 hover:bg-rose-100/50 hover:border-rose-300";
                                statusLabel = "วันหยุดบริษัท 🏢";
                                statusColor = "text-rose-700 font-bold";
                              }
                            }

                            const workingG2Count = dayData ? dayData.workingStaff.filter(s => GROUP_2.includes(s)).length : 0;
                            const offG2Count = 9 - workingG2Count;
                            const isExcessiveOff = offG2Count > 4;

                            return (
                              <div
                                key={cell.date}
                                title={new Date(cell.date).toLocaleDateString("th-TH", { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}
                                onClick={() => {
                                  toggleCellState(cell.date, activeStaff);
                                }}
                                className={`group/cell relative border rounded-xl p-1.5 sm:p-2 min-h-[70px] sm:min-h-[95px] transition-all flex flex-col justify-between cursor-pointer ${bgClass} shadow-2xs hover:shadow-xs hover:ring-1 hover:ring-emerald-300 ${isExcessiveOff ? 'ring-1 ring-red-300' : ''}`}
                              >
                                <div className="flex items-center justify-between">
                                  <span className={`text-[10px] sm:text-xs font-black ${isWeekend ? 'text-red-500' : 'text-gray-500'}`}>
                                    {cell.dateNum}
                                  </span>
                                  {/* Quick Action Indicator or Excessive Off Warning */}
                                  {isExcessiveOff ? (
                                    <span className="bg-red-500 text-white font-black text-[7.5px] sm:text-[8px] px-1 py-[1px] rounded flex items-center gap-0.5 shadow-sm" title={`ในทีมหยุดเยอะเกิน 4 คน (หยุด ${offG2Count} คน จาก 9 คน)`}>
                                      ⚠️ หยุด {offG2Count} คน
                                    </span>
                                  ) : (
                                    <span className="opacity-0 group-hover/cell:opacity-100 text-[8px] bg-emerald-600 text-white px-1 py-0.2 rounded font-black transition-all">
                                      เปลี่ยนสถานะ 🔄
                                    </span>
                                  )}
                                </div>

                                <div className="flex flex-col items-start gap-0.5 mt-1">
                                  <span className={`text-[9px] sm:text-[10px] truncate max-w-full ${statusColor}`}>
                                    {statusLabel}
                                  </span>
                                  {hasHoliday && (
                                    <span className={`text-[8px] font-extrabold px-1 py-0.5 rounded-sm leading-tight max-w-full truncate block mt-0.5 ${
                                      hasHoliday.type === "public"
                                        ? "bg-amber-100 text-amber-800 border border-amber-200"
                                        : "bg-rose-100 text-rose-800 border border-rose-200"
                                    }`} title={hasHoliday.name}>
                                      {hasHoliday.name}
                                    </span>
                                  )}
                                  
                                  {isWorking && (
                                    <div className="flex flex-wrap gap-0.5 mt-0.5">
                                      {isDoc && (
                                        <span className="text-[7.5px] font-bold bg-blue-100 text-blue-800 border border-blue-200 px-0.8 py-0.2 rounded leading-none">
                                          Doc
                                        </span>
                                      )}
                                      {isP1 && (
                                        <span className="text-[7.5px] font-bold bg-sky-100 text-sky-800 border border-sky-200 px-0.8 py-0.2 rounded leading-none flex items-center">
                                          <PhoneCall className="w-[6px] h-[6px] mr-0.5" /> 3551
                                        </span>
                                      )}
                                      {isP2 && (
                                        <span className="text-[7.5px] font-bold bg-indigo-100 text-indigo-800 border border-indigo-200 px-0.8 py-0.2 rounded leading-none flex items-center">
                                          <PhoneCall className="w-[6px] h-[6px] mr-0.5" /> 3552
                                        </span>
                                      )}
                                    </div>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>


                    </div>
                  );
                })()}

                {/* TAB 2: WORKLOAD STATISTICS */}
                {spotlightTab === "stats" && (
                  <div className="flex flex-col gap-4">
                    {/* Header showing name and period clearly */}
                    <div className="bg-emerald-50 border border-emerald-100 p-3.5 rounded-2xl flex flex-col sm:flex-row items-center justify-between gap-3 shadow-xs">
                      <div className="flex items-center gap-2 text-xs sm:text-sm font-extrabold text-emerald-950">
                        <TrendingUp className="w-4 h-4 text-emerald-600" />
                        <span>สถิติภาระงานสะสมของ คุณ{activeStaff}</span>
                      </div>
                      <span className="bg-white border border-emerald-200 text-emerald-900 font-black text-xs px-3.5 py-1.5 rounded-xl shadow-xs">
                        📅 ประจำรอบเดือน: {new Date(`${currentMonthStr}-01`).toLocaleDateString("th-TH", { month: 'long', year: 'numeric' })}
                      </span>
                    </div>

                    {/* Dashboard stats panel */}
                    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
                      <div className="bg-emerald-50/50 border border-emerald-100 p-3 rounded-2xl text-center shadow-2xs">
                        <span className="text-[10px] font-bold text-emerald-800 uppercase tracking-wider block">ขึ้นเวรทั้งหมด</span>
                        <span className="text-xl sm:text-2xl font-black text-emerald-600 mt-1 block">
                          {stats.workingDays} <span className="text-xs font-normal text-gray-500">วัน</span>
                        </span>
                        <div className="w-full bg-gray-100 h-1.5 rounded-full mt-2 overflow-hidden">
                          <div className="bg-emerald-500 h-full rounded-full" style={{ width: `${Math.min(100, (stats.workingDays / 30) * 100)}%` }} />
                        </div>
                      </div>

                      <div className="bg-gray-50 border border-gray-100 p-3 rounded-2xl text-center shadow-2xs">
                        <span className="text-[10px] font-bold text-gray-500 uppercase tracking-wider block">วันหยุดพักผ่อน</span>
                        <span className="text-xl sm:text-2xl font-black text-gray-700 mt-1 block">
                          {stats.offDays} <span className="text-xs font-normal text-gray-400">วัน</span>
                        </span>
                        <div className="w-full bg-gray-200 h-1.5 rounded-full mt-2 overflow-hidden">
                          <div className="bg-gray-400 h-full rounded-full" style={{ width: `${Math.min(100, (stats.offDays / 30) * 100)}%` }} />
                        </div>
                      </div>

                      <div className="bg-orange-50/40 border border-orange-100 p-3 rounded-2xl text-center shadow-2xs">
                        <span className="text-[10px] font-bold text-orange-800 block uppercase tracking-wider">ลาพักร้อน (Vac.)</span>
                        <span className="text-xl sm:text-2xl font-black text-orange-600 mt-1 block">
                          {stats.vacationDays} <span className="text-xs font-normal text-gray-500">วัน</span>
                        </span>
                        <div className="w-full bg-gray-150 h-1.5 rounded-full mt-2 overflow-hidden">
                          <div className="bg-orange-400 h-full rounded-full" style={{ width: `${Math.min(100, (stats.vacationDays / 10) * 100)}%` }} />
                        </div>
                      </div>

                      <div className="bg-sky-50/40 border border-sky-100 p-3 rounded-2xl text-center shadow-2xs">
                        <span className="text-[10px] font-bold text-sky-800 uppercase tracking-wider block">สายด่วนหลัก 3551</span>
                        <span className="text-xl sm:text-2xl font-black text-sky-600 mt-1 block">
                          {stats.phone3551Days} <span className="text-xs font-normal text-gray-500">ครั้ง</span>
                        </span>
                        <div className="w-full bg-gray-100 h-1.5 rounded-full mt-2 overflow-hidden">
                          <div className="bg-sky-400 h-full rounded-full" style={{ width: `${Math.min(100, (stats.phone3551Days / 8) * 100)}%` }} />
                        </div>
                      </div>

                      <div className="bg-indigo-50/40 border border-indigo-100 p-3 rounded-2xl text-center shadow-2xs">
                        <span className="text-[10px] font-bold text-indigo-800 uppercase tracking-wider block">สายด่วนรอง 3552</span>
                        <span className="text-xl sm:text-2xl font-black text-indigo-600 mt-1 block">
                          {stats.phone3552Days} <span className="text-xs font-normal text-gray-500">ครั้ง</span>
                        </span>
                        <div className="w-full bg-gray-150 h-1.5 rounded-full mt-2 overflow-hidden">
                          <div className="bg-indigo-400 h-full rounded-full" style={{ width: `${Math.min(100, (stats.phone3552Days / 8) * 100)}%` }} />
                        </div>
                      </div>

                      <div className="bg-purple-50/40 border border-purple-100 p-3 rounded-2xl text-center shadow-2xs">
                        <span className="text-[10px] font-bold text-purple-800 uppercase tracking-wider block">In-charge (Doc.)</span>
                        <span className="text-xl sm:text-2xl font-black text-purple-600 mt-1 block">
                          {stats.docInChargeDays} <span className="text-xs font-normal text-gray-500">ครั้ง</span>
                        </span>
                        <div className="w-full bg-gray-150 h-1.5 rounded-full mt-2 overflow-hidden">
                          <div className="bg-purple-400 h-full rounded-full" style={{ width: `${Math.min(100, (stats.docInChargeDays / 8) * 100)}%` }} />
                        </div>
                      </div>
                    </div>

                    {/* Duty dates grid with quick jump to daily view */}
                    <div>
                      <span className="text-xs font-bold text-gray-500 uppercase tracking-wider block mb-2">
                        📅 รายการวันที่เข้าเวรในเดือนนี้ ({stats.dutyDates.length} วัน) - <span className="text-emerald-700">คลิกที่วันเพื่อดูรายชื่อเพื่อนร่วมเวรวันนั้นๆ:</span>
                      </span>
                      {stats.dutyDates.length === 0 ? (
                        <div className="text-center py-3 text-xs text-gray-400 bg-gray-50 border border-dashed border-gray-100 rounded-xl font-medium">
                          ไม่มีวันปฏิบัติเวรในรอบเดือนนี้
                        </div>
                      ) : (
                        <div className="flex flex-wrap gap-1.5 max-h-[140px] overflow-y-auto custom-scrollbar p-1 bg-gray-50/50 rounded-xl border border-gray-100">
                          {stats.dutyDates.map(dd => (
                            <button
                              key={dd.date}
                              onClick={() => {
                                setSelectedDate(dd.date);
                                setViewMode("daily");
                                showToast(`แสดงข้อมูลทีมและสถานะเวรประจำวันที่ ${new Date(dd.date).toLocaleDateString("th-TH", { day: 'numeric', month: 'short' })} เรียบร้อย`, false, new Date(dd.date).toLocaleDateString("th-TH"));
                              }}
                              className="flex flex-col items-start px-2.5 py-1.5 bg-white hover:bg-emerald-50 border border-emerald-100 rounded-xl hover:border-emerald-300 shadow-2xs hover:shadow-xs cursor-pointer transition-all active:scale-95 text-left text-xs min-w-[70px] sm:min-w-[80px]"
                            >
                              <span className="text-[9px] font-bold text-gray-400 leading-none">{dd.dayName}</span>
                              <span className="text-sm font-black text-emerald-950 mt-1 leading-none">{dd.dateNum}</span>
                              {dd.roleBadge ? (
                                <span className="text-[7.5px] font-black bg-emerald-100 text-emerald-800 border border-emerald-200 px-1 py-0.2 rounded mt-2 leading-none inline-block truncate max-w-full">
                                  {dd.roleBadge}
                                </span>
                              ) : (
                                <span className="text-[7.5px] font-bold text-gray-400 px-1 py-0.2 mt-2 leading-none inline-block">
                                  เวรทั่วไป
                                </span>
                              )}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            );
          })()}
        </section>

        {/* 📱 DAILY VIEW (Optimized for Mobile) */}
        {viewMode === "daily" && (() => {
          const currentDayData = schedule[selectedDate] || {
            date: selectedDate,
            workingStaff: [],
            fireCodes: {},
            vacationStaff: [],
            docInCharge: null,
            phone3551: null,
            phone3552: null
          };
          const activeWorking = currentDayData.workingStaff || [];
          const activeVacation = currentDayData.vacationStaff || [];
          const activeOff = ALL_STAFF.filter(s => !activeWorking.includes(s) && !activeVacation.includes(s));

          return (
            <div className="flex flex-col gap-5 w-full max-w-5xl mx-auto">
              {/* 1. Date selector card */}
              <div className="bg-white rounded-2xl shadow-sm border border-emerald-100/60 p-4 sm:p-5 flex flex-col gap-4">
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                  <div>
                    <h2 className="text-lg font-bold text-emerald-800 flex flex-wrap items-center gap-2">
                      <CalendarDays className="w-5 h-5 text-emerald-600 shrink-0" />
                      <span>เจาะลึกตารางรายวัน:</span>
                      <span className="bg-emerald-100 text-emerald-950 px-3 py-1 rounded-xl text-xs sm:text-sm font-black border border-emerald-200 shadow-2xs">
                        {new Date(selectedDate).toLocaleDateString("th-TH", { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}
                      </span>
                    </h2>
                    <p className="text-gray-500 text-xs mt-1.5">รายละเอียดผู้ปฏิบัติงานและสรุปยอดกำลังพลรายวัน</p>
                  </div>
                  <div className="flex items-center gap-2 w-full sm:w-auto justify-between sm:justify-start">
                    <button 
                      onClick={handlePrevDay}
                      className="p-2 border border-gray-200 rounded-xl hover:bg-gray-50 active:bg-gray-100 transition-all text-gray-700"
                      title="วันก่อนหน้า"
                    >
                      <ArrowRight className="w-4 h-4 rotate-180" />
                    </button>
                    
                    <input
                      type="date"
                      value={selectedDate}
                      onChange={(e) => {
                        if (e.target.value) {
                          setSelectedDate(e.target.value);
                          const [y, m] = e.target.value.split('-');
                          setCurrentMonthStr(`${y}-${m}`);
                        }
                      }}
                      className="flex-1 sm:flex-none text-center border border-emerald-200 text-sm rounded-xl px-3 py-2 bg-white text-emerald-800 font-bold focus:ring-2 focus:ring-emerald-500 focus:outline-none shadow-sm cursor-pointer"
                    />

                    <button 
                      onClick={handleNextDay}
                      className="p-2 border border-gray-200 rounded-xl hover:bg-gray-50 active:bg-gray-100 transition-all text-gray-700"
                      title="วันถัดไป"
                    >
                      <ArrowRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                {/* Touch-scrollable day tape */}
                <div className="flex gap-2 overflow-x-auto py-1 custom-scrollbar w-full pb-2 scroll-smooth">
                  {datesInMonth.map((d) => {
                    const isSelected = selectedDate === d.date;
                    const hasHoliday = d.holiday;
                    let tapeClass = "";
                    if (isSelected) {
                      tapeClass = "bg-emerald-600 text-white shadow-md font-bold scale-105";
                    } else if (hasHoliday) {
                      if (hasHoliday.type === "public") {
                        tapeClass = "bg-amber-100 hover:bg-amber-200 text-amber-800 border border-amber-200";
                      } else {
                        tapeClass = "bg-rose-50 hover:bg-rose-100 text-rose-800 border border-rose-200";
                      }
                    } else if (d.isWeekend) {
                      tapeClass = "bg-emerald-50/50 hover:bg-emerald-50 text-emerald-800 border border-emerald-100/60";
                    } else {
                      tapeClass = "bg-gray-50 hover:bg-emerald-50 text-gray-700 border border-gray-100";
                    }

                    return (
                      <button
                        key={d.date}
                        onClick={() => setSelectedDate(d.date)}
                        className={`flex flex-col items-center justify-center p-2.5 rounded-xl min-w-[58px] transition-all shrink-0 cursor-pointer ${tapeClass}`}
                        title={hasHoliday ? `${hasHoliday.name} (${hasHoliday.type === 'public' ? 'วันหยุดนักขัตฤกษ์' : 'วันหยุดบริษัท'})` : ""}
                      >
                        <span className="text-[9px] uppercase tracking-wider font-bold opacity-85">{d.dayName}</span>
                        <span className="text-base font-black mt-0.5 relative">
                          {d.dateNum}
                          {hasHoliday && (
                            <span className={`absolute -top-1 -right-1 text-[8px] ${hasHoliday.type === 'public' ? 'text-amber-500' : 'text-rose-500'}`}>⭐</span>
                          )}
                        </span>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* 2. Today's Statistics Summary Cards */}
              <div className="grid grid-cols-3 gap-2 sm:gap-4">
                <div className="bg-emerald-50 border border-emerald-100 rounded-2xl p-3 text-center flex flex-col justify-center items-center shadow-sm">
                  <span className="text-[10px] sm:text-xs font-semibold text-emerald-800">ขึ้นเวรวันนี้</span>
                  <span className="text-xl sm:text-3xl font-black text-emerald-600 mt-1">{activeWorking.length} <span className="text-[10px] sm:text-xs font-normal text-gray-500">คน</span></span>
                </div>
                <div className="bg-amber-50 border border-amber-100 rounded-2xl p-3 text-center flex flex-col justify-center items-center shadow-sm">
                  <span className="text-[10px] sm:text-xs font-semibold text-amber-800">หยุดพักเวร</span>
                  <span className="text-xl sm:text-3xl font-black text-amber-600 mt-1">{activeOff.length} <span className="text-[10px] sm:text-xs font-normal text-gray-500">คน</span></span>
                </div>
                <div className="bg-orange-50 border border-orange-100 rounded-2xl p-3 text-center flex flex-col justify-center items-center shadow-sm">
                  <span className="text-[10px] sm:text-xs font-semibold text-orange-800">ลาพักร้อน</span>
                  <span className="text-xl sm:text-3xl font-black text-orange-600 mt-1">{activeVacation.length} <span className="text-[10px] sm:text-xs font-normal text-gray-500">คน</span></span>
                </div>
              </div>

              {/* 3. Primary Roles Highlights */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {/* 3551 Card */}
                <div className="bg-gradient-to-br from-sky-50 to-white border border-sky-100/70 rounded-2xl p-4 flex items-center gap-3 shadow-xs">
                  <div className="bg-sky-500 text-white p-3 rounded-xl shadow-sm shrink-0">
                    <PhoneCall className="w-5 h-5" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <span className="text-[10px] font-bold text-sky-600 uppercase tracking-wider block">สายด่วนหลัก (ต่วนมารีนา)</span>
                    <span className="text-base sm:text-lg font-black text-sky-900 block truncate mt-0.5">{currentDayData.phone3551 || "ไม่มีผู้รับสาย"}</span>
                    <span className="text-[10px] text-gray-400 font-medium block">เบอร์โทรศัพท์ 3551</span>
                  </div>
                </div>

                {/* 3552 Card */}
                <div className="bg-gradient-to-br from-indigo-50 to-white border border-indigo-100/70 rounded-2xl p-4 flex items-center gap-3 shadow-xs">
                  <div className="bg-indigo-600 text-white p-3 rounded-xl shadow-sm shrink-0">
                    <PhoneCall className="w-5 h-5" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <span className="text-[10px] font-bold text-indigo-600 uppercase tracking-wider block">สายด่วนรอง (ศศลักษณ์/ธนัชพร)</span>
                    <span className="text-base sm:text-lg font-black text-indigo-900 block truncate mt-0.5">{currentDayData.phone3552 || "ไม่มีผู้รับสาย"}</span>
                    <span className="text-[10px] text-gray-400 font-medium block">เบอร์โทรศัพท์ 3552</span>
                  </div>
                </div>

                {/* Doc Card */}
                <div className="bg-gradient-to-br from-purple-50 to-white border border-purple-100/70 rounded-2xl p-4 flex items-center gap-3 shadow-xs">
                  <div className="bg-purple-600 text-white p-3 rounded-xl shadow-sm shrink-0">
                    <FileText className="w-5 h-5" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <span className="text-[10px] font-bold text-purple-600 uppercase tracking-wider block">In-charge เอกสาร (Doc.)</span>
                    <span className="text-base sm:text-lg font-black text-purple-900 block truncate mt-0.5">{currentDayData.docInCharge || "ไม่มีผู้ดูแล"}</span>
                    <span className="text-[10px] text-gray-400 font-medium block">ผู้บันทึกรายงานรับเอกสาร</span>
                  </div>
                </div>
              </div>

              {/* 4. Categorized Staff Lists */}
              <div className="bg-white rounded-2xl border border-gray-100 p-4 sm:p-5 flex flex-col gap-6">
                {/* Section 4a: ขึ้นเวร */}
                <div>
                  <h3 className="text-sm font-bold text-emerald-800 flex items-center gap-1.5 border-b border-gray-100 pb-2 mb-3">
                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
                    <span>ขึ้นเวรวันนี้ ({activeWorking.length} คน)</span>
                  </h3>
                  {activeWorking.length === 0 ? (
                    <p className="text-xs text-gray-400 italic py-2 text-center">ไม่มีเจ้าหน้าที่ขึ้นเวรวันนี้</p>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                      {activeWorking.map(staff => {
                        const isPhone1 = currentDayData.phone3551 === staff;
                        const isPhone2 = currentDayData.phone3552 === staff;
                        const isDoc = currentDayData.docInCharge === staff;
                        const fireCode = currentDayData.fireCodes[staff];
                        
                        return (
                          <div key={staff} className="bg-emerald-50/40 border border-emerald-100/60 rounded-xl p-3 flex items-center justify-between gap-3 shadow-xs">
                            <div className="min-w-0 flex-1">
                              <div className="flex items-center gap-2">
                                <span className="text-sm font-bold text-emerald-900 truncate">{staff}</span>
                                {GROUP_1.includes(staff) && (
                                  <span className="text-[8px] bg-emerald-100 text-emerald-800 px-1 py-0.5 rounded font-medium shrink-0">หยุด ส.-อา.</span>
                                )}
                              </div>
                              
                              {/* Role Badges */}
                              <div className="flex flex-wrap gap-1 mt-1.5">
                                {isPhone1 && (
                                  <span className="inline-flex items-center text-[9px] font-black bg-sky-100 text-sky-800 border border-sky-200 px-1.5 py-0.5 rounded-md">
                                    <PhoneCall className="w-2.5 h-2.5 mr-1 text-sky-600" /> 3551 สายหลัก
                                  </span>
                                )}
                                {isPhone2 && (
                                  <span className="inline-flex items-center text-[9px] font-black bg-indigo-100 text-indigo-800 border border-indigo-200 px-1.5 py-0.5 rounded-md">
                                    <PhoneCall className="w-2.5 h-2.5 mr-1 text-indigo-600" /> 3552 สายรอง
                                  </span>
                                )}
                                {isDoc && (
                                  <span className="inline-flex items-center text-[9px] font-bold bg-purple-100 text-purple-800 border border-purple-200 px-1.5 py-0.5 rounded-md">
                                    <FileText className="w-2.5 h-2.5 mr-1 text-purple-600" /> Doc รับเอกสาร
                                  </span>
                                )}
                                {fireCode && (
                                  <span className="inline-flex items-center text-[9px] font-bold bg-red-100 text-red-800 border border-red-200 px-1.5 py-0.5 rounded-md">
                                    ดับเพลิง: {fireCode}
                                  </span>
                                )}
                              </div>
                            </div>

                            {/* Admin Direct Edit */}
                            {isAdmin && (
                              <div className="flex gap-1 shrink-0">
                                <button 
                                  onClick={() => setStaffStatus(selectedDate, staff, 'OFF')}
                                  className="p-1 bg-white border border-gray-200 hover:border-red-300 hover:bg-red-50 text-gray-500 hover:text-red-600 rounded-lg transition-all"
                                  title="เปลี่ยนเป็นหยุด"
                                >
                                  <X className="w-3.5 h-3.5" />
                                </button>
                                <button 
                                  onClick={() => setStaffStatus(selectedDate, staff, 'VAC')}
                                  className="p-1 bg-white border border-gray-200 hover:border-orange-300 hover:bg-orange-50 text-gray-500 hover:text-orange-600 rounded-lg transition-all"
                                  title="เปลี่ยนเป็นพักร้อน"
                                >
                                  <span className="w-3.5 h-3.5 flex items-center justify-center text-[9px] font-black">V</span>
                                </button>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>

                {/* Section 4b: หยุดพัก */}
                <div>
                  <h3 className="text-sm font-bold text-gray-600 flex items-center gap-1.5 border-b border-gray-100 pb-2 mb-3">
                    <span className="w-2.5 h-2.5 rounded-full bg-gray-400"></span>
                    <span>หยุดงานวันนี้ ({activeOff.length} คน)</span>
                  </h3>
                  {activeOff.length === 0 ? (
                    <p className="text-xs text-gray-400 italic py-2 text-center">ไม่มีเจ้าหน้าที่หยุดวันนี้</p>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                      {activeOff.map(staff => (
                        <div key={staff} className="bg-gray-50 border border-gray-100 rounded-xl p-3 flex items-center justify-between gap-3 opacity-85">
                          <div className="min-w-0 flex-1">
                            <div className="flex items-center gap-2">
                              <span className="text-sm font-medium text-gray-600 truncate">{staff}</span>
                              {GROUP_1.includes(staff) && (
                                <span className="text-[8px] bg-gray-100 text-gray-500 px-1 py-0.5 rounded shrink-0">หยุด ส.-อา.</span>
                              )}
                            </div>
                          </div>

                          {/* Admin Direct Edit */}
                          {isAdmin && (
                            <div className="flex gap-1 shrink-0">
                              <button 
                                onClick={() => setStaffStatus(selectedDate, staff, 'WORK')}
                                className="px-2 py-1 text-[10px] font-bold bg-white border border-emerald-200 hover:bg-emerald-50 text-emerald-600 rounded-lg transition-all flex items-center gap-0.5"
                              >
                                <Check className="w-3 h-3" /> ขึ้นเวร
                              </button>
                              <button 
                                onClick={() => setStaffStatus(selectedDate, staff, 'VAC')}
                                className="p-1 bg-white border border-gray-200 hover:border-orange-300 hover:bg-orange-50 text-gray-500 hover:text-orange-600 rounded-lg transition-all"
                                title="เปลี่ยนเป็นพักร้อน"
                              >
                                <span className="w-3.5 h-3.5 flex items-center justify-center text-[9px] font-black">V</span>
                              </button>
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                {/* Section 4c: พักร้อน */}
                <div>
                  <h3 className="text-sm font-bold text-orange-800 flex items-center gap-1.5 border-b border-gray-100 pb-2 mb-3">
                    <span className="w-2.5 h-2.5 rounded-full bg-orange-400 animate-pulse"></span>
                    <span>ลาพักร้อนวันนี้ ({activeVacation.length} คน)</span>
                  </h3>
                  {activeVacation.length === 0 ? (
                    <p className="text-xs text-gray-400 italic py-2 text-center">ไม่มีเจ้าหน้าที่ลาพักร้อนวันนี้</p>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                      {activeVacation.map(staff => (
                        <div key={staff} className="bg-orange-50/40 border border-orange-100/60 rounded-xl p-3 flex items-center justify-between gap-3">
                          <div className="min-w-0 flex-1">
                            <div className="flex items-center gap-1.5">
                              <span className="text-sm font-bold text-orange-950 truncate">{staff}</span>
                              <span className="text-[10px] text-orange-600 shrink-0">🌴 พักร้อน</span>
                            </div>
                          </div>

                          {/* Admin Direct Edit */}
                          {isAdmin && (
                            <div className="flex gap-1 shrink-0">
                              <button 
                                onClick={() => setStaffStatus(selectedDate, staff, 'WORK')}
                                className="px-2 py-1 text-[10px] font-bold bg-white border border-emerald-200 hover:bg-emerald-50 text-emerald-600 rounded-lg transition-all flex items-center gap-0.5"
                              >
                                <Check className="w-3 h-3" /> ขึ้นเวร
                              </button>
                              <button 
                                onClick={() => setStaffStatus(selectedDate, staff, 'OFF')}
                                className="p-1 bg-white border border-gray-200 hover:border-red-300 hover:bg-red-50 text-gray-500 hover:text-red-600 rounded-lg transition-all"
                                title="เปลี่ยนเป็นหยุด"
                              >
                                <X className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </div>
          );
        })()}

        {/* 💻 MONTHLY VIEW (Full Month Calendar Table) */}
        {viewMode === "monthly" && (
          <section className="flex-1 bg-white rounded-2xl shadow-sm border border-emerald-100/60 overflow-hidden flex flex-col min-h-[600px]">
          <div className="p-4 sm:p-5 bg-gradient-to-r from-emerald-50 to-white border-b border-emerald-100/50 flex flex-col gap-3">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 w-full">
              <h2 className="text-base sm:text-lg font-semibold text-emerald-800 flex flex-wrap items-center gap-1 sm:gap-2">
                <CalendarDays className="w-5 h-5 text-emerald-600 shrink-0" /> 
                <span>ตารางเวรประจำเดือน</span>
                <input 
                  type="month" 
                  value={currentMonthStr} 
                  onChange={(e) => {
                    if (e.target.value) setCurrentMonthStr(e.target.value);
                  }}
                  className="border border-emerald-200 text-sm rounded-lg px-2 py-1.5 focus:ring-2 focus:ring-emerald-500 focus:outline-none bg-white text-emerald-800 font-medium cursor-pointer shadow-sm"
                />
                {isAdmin ? (
                  <span className="text-xs font-semibold text-amber-600 bg-amber-50 border border-amber-200 px-2 py-1 rounded-full animate-pulse whitespace-nowrap">
                    ✍️ โหมดแก้ไข (Admin Active)
                  </span>
                ) : (
                  <span className="text-xs font-medium text-emerald-600 bg-emerald-50 border border-emerald-200 px-2 py-1 rounded-full whitespace-nowrap">
                    👀 โหมดอ่านอย่างเดียว (Viewer)
                  </span>
                )}
              </h2>
              <div className="flex gap-2">
                <button onClick={handleExportExcel} className="flex items-center text-xs sm:text-sm font-medium bg-white text-emerald-700 border border-emerald-200 px-2.5 sm:px-3 py-1.5 rounded-lg hover:bg-emerald-50 transition-colors shadow-sm">
                  <Download className="w-4 h-4 mr-1 sm:mr-1.5" /> ส่งออก Excel
                </button>
                {isAdmin ? (
                  <>
                    <label htmlFor="import-excel" className="flex items-center text-xs sm:text-sm font-medium bg-emerald-600 text-white px-2.5 sm:px-3 py-1.5 rounded-lg hover:bg-emerald-700 transition-colors shadow-sm cursor-pointer">
                      <Upload className="w-4 h-4 mr-1 sm:mr-1.5" /> นำเข้า Excel
                    </label>
                    <input id="import-excel" type="file" accept=".xlsx, .xls" className="hidden" onChange={handleImportExcel} />
                  </>
                ) : (
                  <button 
                    onClick={() => {
                      showToast("กรุณาเปิด 'โหมดแก้ไข (Admin)' เพื่อนำเข้าข้อมูล", false, new Date().toLocaleDateString("th-TH"));
                      setShowAdminModal(true);
                    }}
                    className="flex items-center text-xs sm:text-sm font-medium bg-gray-100 text-gray-400 border border-gray-200 px-2.5 sm:px-3 py-1.5 rounded-lg hover:bg-gray-200 hover:text-gray-500 transition-colors shadow-sm"
                  >
                    <Lock className="w-3.5 h-3.5 mr-1" /> นำเข้า Excel
                  </button>
                )}
              </div>
            </div>
            <div className="flex flex-wrap gap-x-3 gap-y-2 text-[10px] sm:text-xs font-medium text-gray-600">
               <span className="flex items-center"><Check className="w-3 h-3 sm:w-4 sm:h-4 text-emerald-500 mr-1"/> ขึ้นเวร</span>
               <span className="flex items-center"><X className="w-3 h-3 sm:w-4 sm:h-4 text-red-500 mr-1"/> หยุด</span>
               <span className="flex items-center"><span className="text-[9px] sm:text-[10px] font-bold text-orange-600 mr-1">V</span> พักร้อน</span>
               <span className="flex items-center"><span className="text-[8px] sm:text-[9px] font-bold text-blue-700 bg-blue-100 border border-blue-200 px-1 rounded mr-1">Doc.</span> รับเอกสาร</span>
               <span className="flex items-center"><span className="text-[9px] sm:text-[10px] font-bold text-red-600 bg-red-100 border border-red-200 px-1 rounded mr-1">A..I</span> โค้ดดับเพลิง</span>
               <span className="flex items-center"><span className="flex items-center text-[8px] sm:text-[9px] font-bold text-sky-700 bg-sky-100 border border-sky-200 px-1 rounded mr-1"><PhoneCall className="w-[8px] h-[8px] sm:w-[10px] sm:h-[10px] mr-1" /> 3551</span> ต่วนมารีนา</span>
               <span className="flex items-center"><span className="flex items-center text-[8px] sm:text-[9px] font-bold text-indigo-700 bg-indigo-100 border border-indigo-200 px-1 rounded mr-1"><PhoneCall className="w-[8px] h-[8px] sm:w-[10px] sm:h-[10px] mr-1" /> 3552</span> ศศลักษณ์/ธนัชพร</span>
            </div>
          </div>
          <div className="overflow-auto max-h-[75vh] flex-1 custom-scrollbar w-full">
            <table className="w-full border-collapse text-sm min-w-max">
              <thead>
                <tr>
                  <th className="bg-emerald-600 border-b-2 border-emerald-700 text-white p-2 sm:p-3 font-medium text-left sticky top-0 left-0 z-40 min-w-[140px] sm:min-w-[180px] shadow-[2px_0_5px_rgba(0,0,0,0.1)]">
                    รายชื่อเจ้าหน้าที่ (14 คน)
                  </th>
                  {datesInMonth.map((d) => {
                    const dayData = schedule[d.date];
                    const workingG2Count = dayData ? dayData.workingStaff.filter(s => GROUP_2.includes(s)).length : 0;
                    const offG2Count = 9 - workingG2Count;
                    const isExcessiveOff = offG2Count > 4;
                    const hasHoliday = d.holiday;
                    const dayOfWeek = d.obj.getDay();

                    let headerBg = d.isWeekend ? "bg-emerald-700 text-white" : "bg-emerald-600 text-white";
                    if (dayOfWeek === 6) {
                      headerBg = "bg-purple-600 text-white"; // Saturday Purple
                    } else if (dayOfWeek === 0) {
                      headerBg = "bg-red-600 text-white"; // Sunday Red
                    }

                    if (hasHoliday) {
                      if (hasHoliday.type === "public") {
                        headerBg = "bg-amber-500 text-white";
                      } else {
                        headerBg = "bg-rose-500 text-white";
                      }
                    }

                    let headerTitle = "";
                    if (hasHoliday) {
                      headerTitle += `[วันหยุด: ${hasHoliday.name} (${hasHoliday.type === "public" ? "นักขัตฤกษ์" : "บริษัท"})] `;
                    }
                    if (isExcessiveOff) {
                      headerTitle += `เตือนภัย: วันนี้มีทีมหยุดเกิน 4 คน! (หยุด ${offG2Count} คน จาก 9 คน)`;
                    }

                    return (
                      <th
                        key={d.date}
                        className={`p-1 sm:p-[3px] md:p-2 min-w-[40px] sm:min-w-[44px] lg:min-w-[48px] border-l border-emerald-500/20 text-center relative sticky top-0 z-30 ${headerBg} ${isExcessiveOff ? "ring-2 ring-red-500 ring-inset" : ""}`}
                        title={headerTitle || undefined}
                      >
                        <div className="text-[8px] sm:text-[10px] opacity-90 tracking-[0.05em] sm:tracking-widest uppercase">
                          {d.dayName}
                        </div>
                        <div className="font-semibold text-sm sm:text-lg mt-0.5 relative inline-block">
                          {d.dateNum}
                          {isExcessiveOff && (
                            <span className="absolute -top-1 -right-2 text-red-500 text-[9px] font-bold animate-pulse" title="หยุดเกิน 4 คน!">
                              ⚠️
                            </span>
                          )}
                          {hasHoliday && (
                            <span
                              className="absolute bottom-[-2px] left-1/2 -translate-x-1/2 w-1 h-1 rounded-full bg-white animate-pulse"
                              title={hasHoliday.name}
                            />
                          )}
                        </div>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {(() => {
                  const activeSpotlightStaff = selectedStaffFilter || currentUser;
                  const listToRender = (isolateStaffRow && activeSpotlightStaff)
                    ? DISPLAY_ORDER.filter(s => s === activeSpotlightStaff)
                    : DISPLAY_ORDER;

                  return listToRender.map((staff) => {
                    const idx = DISPLAY_ORDER.indexOf(staff);
                    const isStartOfGroup1 = idx === GROUP_2.length && !(isolateStaffRow && activeSpotlightStaff);
                    
                    const isFocused = activeSpotlightStaff === staff;
                    const hasAnyFocus = !!activeSpotlightStaff;
                    
                    let rowClass = "transition-all duration-200 ";
                    if (isFocused) {
                      rowClass += "bg-emerald-100/70 hover:bg-emerald-150/80 shadow-[inset_4px_0_0_0_#10b981] font-bold ring-1 ring-emerald-500/20";
                    } else if (hasAnyFocus) {
                      rowClass += "opacity-30 hover:opacity-100 " + (idx % 2 === 0 ? "bg-white" : "bg-emerald-50/10");
                    } else {
                      rowClass += idx % 2 === 0 ? "bg-white" : "bg-emerald-50/30";
                    }

                    return (
                      <React.Fragment key={staff}>
                        {isStartOfGroup1 && (
                          <tr>
                            <td colSpan={datesInMonth.length + 1} className="bg-emerald-100/50 font-semibold text-emerald-800 p-1.5 sm:p-2.5 text-[10px] sm:text-xs text-center border-y border-emerald-200/60 shadow-inner">
                              --- กลุ่มเจ้าหน้าที่หยุดเสาร์-อาทิตย์ ---
                            </td>
                          </tr>
                        )}
                        <tr className={rowClass}>
                        <td className="p-2 sm:p-3 border-b border-emerald-100 sticky left-0 bg-inherit shadow-[1px_0_3px_rgba(0,0,0,0.03)] z-10 font-medium text-emerald-900 border-r border-r-emerald-100 group text-xs sm:text-sm">
                          <div className="flex items-center justify-between">
                            <span className="truncate pr-1">{staff}</span>
                            {GROUP_1.includes(staff) && (
                              <span className="text-[8px] sm:text-[10px] bg-gray-100 text-gray-500 px-1 py-0.5 rounded ml-1 group-hover:bg-white transition-colors whitespace-nowrap hidden sm:inline-block">หยุด ส.-อา.</span>
                            )}
                          </div>
                        </td>
                        {datesInMonth.map((d) => {
                          const dayData = schedule[d.date];
                          const isWorking = dayData?.workingStaff.includes(staff);
                          const isVacation = dayData?.vacationStaff?.includes(staff);
                          const isDocInCharge = dayData?.docInCharge === staff;
                          const isPhone3551 = dayData?.phone3551 === staff;
                          const isPhone3552 = dayData?.phone3552 === staff;
                          const fireCode = dayData?.fireCodes[staff];

                          // Background column indicator highlights
                          let cellBg = "";
                          if (d.holiday) {
                            cellBg = d.holiday.type === "public" ? "bg-amber-50/50" : "bg-rose-50/40";
                          } else if (d.obj.getDay() === 6) { // Sat
                            cellBg = "bg-purple-50/15";
                          } else if (d.obj.getDay() === 0) { // Sun
                            cellBg = "bg-red-50/15";
                          } else if (d.isWeekend) {
                            cellBg = "bg-gray-50/50";
                          }

                          return (
                            <td
                              key={d.date}
                              onClick={() => toggleCellState(d.date, staff)}
                              title={d.holiday ? `${d.holiday.name} (${d.holiday.type === 'public' ? 'วันหยุดนักขัตฤกษ์' : 'วันหยุดบริษัท'})` : undefined}
                              className={`p-0.5 sm:p-1 border-b border-l border-emerald-100 text-center relative cursor-pointer transition-colors ${cellBg} ${
                                isAdmin
                                  ? "hover:bg-emerald-100/50"
                                  : "hover:bg-amber-100/40"
                              }`}
                            >
                              {isVacation ? (
                                <div className="flex items-center justify-center min-h-[36px] sm:min-h-[46px]">
                                  <span className="text-xs sm:text-sm font-black text-orange-500 select-none pointer-events-none">
                                    V
                                  </span>
                                </div>
                              ) : isWorking ? (
                                <div className="flex flex-col items-center justify-center min-h-[36px] sm:min-h-[46px] gap-0.5 pointer-events-none">
                                  <Check
                                    strokeWidth={4}
                                    className="w-[12px] h-[12px] sm:w-[16px] sm:h-[16px] text-emerald-500"
                                  />
                                  {isDocInCharge && (
                                    <span className="text-[8px] sm:text-[9px] font-bold text-blue-700 bg-blue-50 border border-blue-200 px-0.5 sm:px-1 py-[1px] rounded leading-none shadow-sm select-none">
                                      Doc.
                                    </span>
                                  )}
                                  {isPhone3551 && (
                                    <span className="text-[7.5px] sm:text-[8px] font-bold text-sky-700 bg-sky-50 border border-sky-200 px-0.5 sm:px-1 py-[1px] rounded leading-none shadow-sm select-none flex items-center">
                                      <PhoneCall className="w-[8px] h-[8px] sm:w-[10px] sm:h-[10px] mr-0.5" /> 3551
                                    </span>
                                  )}
                                  {isPhone3552 && (
                                    <span className="text-[7.5px] sm:text-[8px] font-bold text-indigo-700 bg-indigo-50 border border-indigo-200 px-0.5 sm:px-1 py-[1px] rounded leading-none shadow-sm select-none flex items-center">
                                      <PhoneCall className="w-[8px] h-[8px] sm:w-[10px] sm:h-[10px] mr-0.5" /> 3552
                                    </span>
                                  )}
                                  {fireCode && (
                                    <span title={`รหัสดับเพลิง: ${fireCode}`} className="text-[8px] sm:text-[9px] font-bold text-red-600 bg-red-50 border border-red-200 px-0.5 sm:px-1 py-[1px] rounded leading-none shadow-sm select-none">
                                      {fireCode}
                                    </span>
                                  )}
                                </div>
                              ) : (
                                <div className="flex flex-col items-center justify-center min-h-[36px] sm:min-h-[46px] pointer-events-none">
                                  <X strokeWidth={3} className="w-[12px] h-[12px] sm:w-[16px] sm:h-[16px] text-red-400 opacity-70" />
                                </div>
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    </React.Fragment>
                  );
                });
              })()}
              </tbody>
            </table>
          </div>
        </section>
        )}

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
          <section className="bg-white rounded-2xl shadow-sm border border-indigo-100/60 p-4 sm:p-5 shrink-0 flex-1">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4 border-b border-indigo-50 pb-3">
              <h2 className="text-base sm:text-lg font-bold text-indigo-800 flex items-center">
                 <FileText className="w-5 h-5 mr-1.5 sm:mr-2 text-indigo-600 shrink-0" />
                 <span>สรุปจำนวนหน้าที่ In-charge เอกสาร (Doc.)</span>
              </h2>
              {isAdmin && (
                <button
                  onClick={handleAutoScheduleDoc}
                  className="flex items-center justify-center gap-1.5 text-[11px] sm:text-xs font-black bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white px-3.5 py-2 rounded-xl transition-all shadow-xs cursor-pointer select-none"
                  title="จัดตารางเวร Doc ทั้งเดือนโดยเฉลี่ยจำนวนวันเท่าเทียมกันและไม่มีใครได้เวรติดกัน"
                >
                  <Sparkles className="w-3.5 h-3.5 text-amber-300 animate-pulse shrink-0" />
                  <span>จัดเวร Doc อัตโนมัติ</span>
                </button>
              )}
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 sm:gap-3">
               {GROUP_2.map(staff => {
                  const docCount = datesInMonth.filter((d: any) => schedule[d.date]?.docInCharge === staff).length;
                  return (
                    <div key={staff} className="bg-indigo-50 border border-indigo-100/60 rounded-xl p-2 sm:p-3 flex flex-col items-center justify-center relative overflow-hidden group">
                      <span className="text-xs sm:text-sm font-medium text-gray-700 z-10 truncate w-full text-center">{staff}</span>
                      <span className="text-xl sm:text-2xl font-bold text-indigo-600 mt-0.5 z-10">{docCount} <span className="text-[10px] sm:text-xs font-normal text-gray-500">วัน</span></span>
                      <div className="absolute inset-0 bg-white/40 opacity-0 group-hover:opacity-100 transition-opacity z-0"></div>
                    </div>
                  );
               })}
            </div>
          </section>

          <section className="bg-white rounded-2xl shadow-sm border border-sky-100/60 p-4 sm:p-5 shrink-0 flex-1">
            <h2 className="text-base sm:text-lg font-semibold text-sky-800 flex items-center mb-3 sm:mb-4">
               <PhoneCall className="w-5 h-5 mr-1.5 sm:mr-2 text-sky-600 shrink-0" />
               <span>สรุปจำนวนเวรรับสายโทรศัพท์ (3551 / 3552)</span>
            </h2>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 sm:gap-3">
               {GROUP_2.map(staff => {
                  const phoneCount = datesInMonth.filter((d: any) => schedule[d.date]?.phone3551 === staff || schedule[d.date]?.phone3552 === staff).length;
                  return (
                    <div key={staff} className="bg-sky-50 border border-sky-100/60 rounded-xl p-2 sm:p-3 flex flex-col items-center justify-center relative overflow-hidden group">
                      <span className="text-xs sm:text-sm font-medium text-gray-700 z-10 truncate w-full text-center">{staff}</span>
                      <span className="text-xl sm:text-2xl font-bold text-sky-600 mt-0.5 z-10">{phoneCount} <span className="text-[10px] sm:text-xs font-normal text-gray-500">ครั้ง</span></span>
                      <div className="absolute inset-0 bg-white/40 opacity-0 group-hover:opacity-100 transition-opacity z-0"></div>
                    </div>
                  );
               })}
            </div>
          </section>
        </div>

        {/* แดชบอร์ดแบบรายวัน เหมือนในรูปของ User */}
        <section className="bg-white rounded-2xl shadow-sm border border-gray-200 p-4 sm:p-5 shrink-0 mb-6 mx-auto w-full">
          <h2 className="text-base sm:text-lg font-semibold text-gray-800 flex items-center mb-3 sm:mb-4">
             <CalendarDays className="w-5 h-5 mr-1.5 sm:mr-2 text-gray-600 shrink-0" />
             <span>สรุปภาพรวมผู้รับสายประจำเดือนแยกตามวัน</span>
          </h2>
          <div className="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-5 lg:grid-cols-7 xl:grid-cols-8 2xl:grid-cols-10 gap-3">
             {datesInMonth.map((d: any) => {
                const day = schedule[d.date];
                const hasHoliday = d.holiday;
                let borderBg = "bg-gray-50 border-gray-100 hover:bg-white";
                if (hasHoliday) {
                  if (hasHoliday.type === "public") {
                    borderBg = "bg-amber-50/50 border-amber-200 hover:bg-amber-50";
                  } else {
                    borderBg = "bg-rose-50/40 border-rose-200 hover:bg-rose-50/60";
                  }
                }
                return (
                  <div key={d.date} className={`border rounded-lg p-2.5 flex flex-col transition-colors hover:shadow-[0_4px_10px_rgba(0,0,0,0.05)] cursor-default ${borderBg}`}>
                    <div className="flex items-center justify-between mb-2 border-b border-gray-200/60 pb-1.5">
                      <span className="text-xs font-semibold text-gray-600">วันที่ {d.dateNum}</span>
                      {hasHoliday && (
                        <span className={`text-[8.5px] font-black px-1 py-[1px] rounded-md leading-none truncate max-w-[65px] ${hasHoliday.type === 'public' ? 'bg-amber-100 text-amber-800 border border-amber-200' : 'bg-rose-100 text-rose-800 border border-rose-200'}`} title={hasHoliday.name}>
                          {hasHoliday.name}
                        </span>
                      )}
                    </div>
                    <div className="text-[10px] sm:text-xs text-gray-700 flex flex-col gap-1.5">
                      <div className="flex items-center bg-sky-50/50 rounded p-1">
                        <PhoneCall className="w-3 h-3 text-sky-600 mr-1.5 shrink-0"/> 
                        <span className="font-medium text-sky-800">3551:</span>
                        <span className="font-semibold ml-1 text-sky-900 truncate">{day?.phone3551 || '-'}</span>
                      </div>
                      <div className="flex items-center bg-indigo-50/50 rounded p-1">
                        <PhoneCall className="w-3 h-3 text-indigo-600 mr-1.5 shrink-0"/> 
                        <span className="font-medium text-indigo-800">3552:</span>
                        <span className="font-semibold ml-1 text-indigo-900 truncate">{day?.phone3552 || '-'}</span>
                      </div>
                    </div>
                  </div>
                )
             })}
          </div>
        </section>

        {/* 📅 จัดการวันหยุดนักขัตฤกษ์และวันหยุดบริษัท (Holiday Management) */}
        <section className="bg-white rounded-2xl shadow-sm border border-rose-100 p-4 sm:p-5 shrink-0 mb-6 mx-auto w-full">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-rose-50 pb-3 mb-4">
            <h2 className="text-base sm:text-lg font-bold text-rose-800 flex items-center">
               <CalendarRange className="w-5 h-5 mr-1.5 sm:mr-2 text-rose-600 shrink-0" />
               <span>ตารางวันหยุดนักขัตฤกษ์และวันหยุดบริษัท ({currentMonthStr})</span>
            </h2>
            <div className="flex items-center gap-2">
              <span className="inline-flex items-center gap-1.5 text-[10px] sm:text-xs font-semibold bg-amber-50 text-amber-800 border border-amber-200 px-2 py-1 rounded-lg">
                <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                <span>วันหยุดนักขัตฤกษ์ (สีส้ม)</span>
              </span>
              <span className="inline-flex items-center gap-1.5 text-[10px] sm:text-xs font-semibold bg-rose-50 text-rose-800 border border-rose-200 px-2 py-1 rounded-lg">
                <span className="w-2 h-2 rounded-full bg-rose-500"></span>
                <span>วันหยุดบริษัท (สีแดงอ่อน)</span>
              </span>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            {/* ฝั่งซ้าย: รายการวันหยุดของเดือนนี้ */}
            <div className="lg:col-span-7 flex flex-col gap-3">
              <h3 className="text-xs sm:text-sm font-bold text-gray-700">รายการวันหยุดประจำรอบเดือนนี้:</h3>
              {datesInMonth.filter(d => d.holiday).length === 0 ? (
                <div className="text-center py-8 bg-gray-50/50 border border-dashed border-gray-200 rounded-xl">
                  <span className="text-gray-400 text-xs sm:text-sm italic">ไม่มีวันหยุดในรอบเดือนนี้</span>
                </div>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                  {datesInMonth.filter(d => d.holiday).map(d => {
                    const h = d.holiday!;
                    return (
                      <div
                        key={h.date}
                        className={`flex items-center justify-between p-3 rounded-xl border transition-all ${
                          h.type === "public"
                            ? "bg-amber-50/60 border-amber-100 text-amber-900"
                            : "bg-rose-50/40 border-rose-100 text-rose-900"
                        }`}
                      >
                        <div className="min-w-0 flex-1">
                          <span className="text-[10px] uppercase font-bold tracking-wider opacity-75 block">
                            วันที่ {d.dateNum} ({d.dayName})
                          </span>
                          <span className="text-xs sm:text-sm font-black truncate block mt-0.5">
                            {h.name}
                          </span>
                          <span className="text-[9px] font-medium opacity-80 block mt-0.5">
                            {h.type === "public" ? "วันหยุดนักขัตฤกษ์" : "วันหยุดของบริษัท"}
                          </span>
                        </div>
                        {isAdmin && (
                          <button
                            onClick={() => {
                              setHolidays(prev => prev.filter(item => item.date !== h.date));
                              showToast(`ลบวันหยุด "${h.name}" เรียบร้อยแล้ว`, false, new Date().toLocaleDateString("th-TH"));
                            }}
                            className="text-red-500 hover:text-red-700 p-1 hover:bg-red-50 rounded-lg transition-colors cursor-pointer ml-2 shrink-0"
                            title="ลบวันหยุด"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* ฝั่งขวา: ฟอร์มเพิ่มวันหยุด (เฉพาะแอดมิน) */}
            <div className="lg:col-span-5 bg-gray-50/50 border border-gray-100 rounded-2xl p-4 flex flex-col justify-between">
              <div>
                <h3 className="text-xs sm:text-sm font-bold text-gray-700 mb-2.5 flex items-center gap-1">
                  <span>➕ เพิ่มวันหยุดใหม่</span>
                  {!isAdmin && <span className="text-[10px] text-amber-600 font-normal">(เฉพาะ Admin เท่านั้น)</span>}
                </h3>

                <div className="flex flex-col gap-3">
                  <div>
                    <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wider block mb-1">วันที่ต้องการหยุด</label>
                    <input
                      type="date"
                      disabled={!isAdmin}
                      value={newHolidayDate}
                      onChange={(e) => setNewHolidayDate(e.target.value)}
                      className="w-full bg-white border border-gray-200 rounded-xl px-3 py-2 text-xs focus:ring-2 focus:ring-rose-500 focus:border-rose-500 transition-all disabled:opacity-50"
                    />
                  </div>

                  <div>
                    <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wider block mb-1">ชื่อวันหยุด / รายละเอียด</label>
                    <input
                      type="text"
                      disabled={!isAdmin}
                      placeholder="เช่น วันขึ้นปีใหม่, วันหยุดประจำปีบริษัท"
                      value={newHolidayName}
                      onChange={(e) => setNewHolidayName(e.target.value)}
                      className="w-full bg-white border border-gray-200 rounded-xl px-3 py-2 text-xs focus:ring-2 focus:ring-rose-500 focus:border-rose-500 transition-all disabled:opacity-50"
                    />
                  </div>

                  <div>
                    <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wider block mb-1">ประเภทวันหยุด</label>
                    <div className="grid grid-cols-2 gap-2">
                      <button
                        type="button"
                        disabled={!isAdmin}
                        onClick={() => setNewHolidayType("public")}
                        className={`py-2 px-3 rounded-xl text-xs font-bold transition-all border cursor-pointer flex items-center justify-center gap-1.5 ${
                          newHolidayType === "public"
                            ? "bg-amber-100 border-amber-300 text-amber-800 shadow-sm"
                            : "bg-white border-gray-200 text-gray-600 hover:bg-gray-50"
                        } disabled:opacity-50`}
                      >
                        <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                        <span>นักขัตฤกษ์</span>
                      </button>
                      <button
                        type="button"
                        disabled={!isAdmin}
                        onClick={() => setNewHolidayType("company")}
                        className={`py-2 px-3 rounded-xl text-xs font-bold transition-all border cursor-pointer flex items-center justify-center gap-1.5 ${
                          newHolidayType === "company"
                            ? "bg-rose-100 border-rose-300 text-rose-800 shadow-sm"
                            : "bg-white border-gray-200 text-gray-600 hover:bg-gray-50"
                        } disabled:opacity-50`}
                      >
                        <span className="w-2 h-2 rounded-full bg-rose-500"></span>
                        <span>วันหยุดบริษัท</span>
                      </button>
                    </div>
                  </div>
                </div>
              </div>

              <button
                type="button"
                disabled={!isAdmin || !newHolidayDate || !newHolidayName.trim()}
                onClick={handleAddHoliday}
                className="w-full mt-4 flex items-center justify-center gap-1.5 text-xs font-black bg-rose-600 hover:bg-rose-700 active:bg-rose-800 disabled:opacity-40 text-white py-2.5 rounded-xl transition-all shadow-xs cursor-pointer select-none"
              >
                <CalendarPlus className="w-4 h-4" />
                <span>บันทึกวันหยุดใหม่</span>
              </button>
            </div>
          </div>
        </section>

        {/* 🔄 ระบบแจ้งเปลี่ยนและสลับเวร (Shift Change & Swap Request System) */}
        <section className="hidden" id="shift-request-system">
          <div className="border-b border-emerald-50 pb-4 mb-6">
            <h2 className="text-lg sm:text-xl font-bold text-emerald-800 flex items-center gap-2">
              <RefreshCcw className="w-5.5 h-5.5 text-emerald-600 shrink-0" />
              <span>ระบบสลับเวรและแจ้งเปลี่ยนตารางเวร</span>
            </h2>
            <p className="text-gray-500 text-xs sm:text-sm mt-1">
              เจ้าหน้าที่พิมพ์ผลทุกคนสามารถยื่นขอสลับเวร ขอให้เพื่อนขึ้นแทน หรือแจ้งลาพักร้อนได้ โดยเมื่อส่งคำขอระบบจะแจ้งเตือนไปยังทุกคนแบบเรียลไทม์ และอัปเดตตารางเวรให้โดยอัตโนมัติเมื่อคำขอได้รับการอนุมัติ
            </p>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            {/* ฝั่งซ้าย: แบบฟอร์มกรอกคำขอ (5 คอลัมน์) */}
            <div className="lg:col-span-5 bg-emerald-50/30 border border-emerald-100/50 rounded-2xl p-4 sm:p-5">
              <h3 className="font-bold text-sm sm:text-base text-emerald-900 flex items-center gap-2 mb-4">
                <FileText className="w-4 h-4 text-emerald-600" />
                <span>ยื่นแบบคำขอแจ้งเปลี่ยนเวร</span>
              </h3>

              <form onSubmit={handleCreateShiftRequest} className="flex flex-col gap-4">
                {/* ผู้ขอเปลี่ยนเวร */}
                <div>
                  <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-1.5">
                    1. ผู้ขอเปลี่ยนเวร (ฉันคือ...)
                  </label>
                  {currentUser ? (
                    <div className="w-full px-3 py-2 bg-emerald-100/50 border border-emerald-200 rounded-xl text-sm text-emerald-950 font-bold shadow-sm flex items-center justify-between">
                      <span className="flex items-center gap-1.5 text-emerald-900">
                        <User className="w-4 h-4 text-emerald-600" />
                        <span>คุณ{currentUser}</span>
                      </span>
                      <span className="text-[10px] bg-emerald-600 text-white font-bold px-2 py-0.5 rounded-md">ยืนยันตัวตนแล้ว</span>
                    </div>
                  ) : (
                    <select
                      value={reqRequester}
                      onChange={(e) => {
                        setReqRequester(e.target.value);
                        // Auto-adjust target if they selected same person
                        if (reqTargetStaff === e.target.value) {
                          setReqTargetStaff(ALL_STAFF.find(s => s !== e.target.value) || "");
                        }
                      }}
                      className="w-full px-3 py-2 bg-white border border-emerald-200 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none text-emerald-950 font-bold shadow-sm"
                    >
                      {ALL_STAFF.map(s => (
                        <option key={s} value={s}>{s}</option>
                      ))}
                    </select>
                  )}
                </div>

                {/* ประเภทการสลับเวร */}
                <div>
                  <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-1.5">
                    2. ประเภทคำขอเปลี่ยนเวร
                  </label>
                  <div className="grid grid-cols-3 gap-1 bg-white p-1 rounded-xl border border-emerald-100">
                    <button
                      type="button"
                      onClick={() => setReqType("swap")}
                      className={`py-1.5 px-1 rounded-lg text-center text-xs font-semibold transition-all cursor-pointer ${
                        reqType === "swap"
                          ? "bg-emerald-600 text-white shadow"
                          : "text-gray-600 hover:bg-gray-50"
                      }`}
                    >
                      🤝 สลับเวร
                    </button>
                    <button
                      type="button"
                      onClick={() => setReqType("cover")}
                      className={`py-1.5 px-1 rounded-lg text-center text-xs font-semibold transition-all cursor-pointer ${
                        reqType === "cover"
                          ? "bg-emerald-600 text-white shadow"
                          : "text-gray-600 hover:bg-gray-50"
                      }`}
                    >
                      🙋‍♀️ ขึ้นแทน
                    </button>
                    <button
                      type="button"
                      onClick={() => setReqType("leave")}
                      className={`py-1.5 px-1 rounded-lg text-center text-xs font-semibold transition-all cursor-pointer ${
                        reqType === "leave"
                          ? "bg-emerald-600 text-white shadow"
                          : "text-gray-600 hover:bg-gray-50"
                      }`}
                    >
                      🌴 ขอลาพักร้อน
                    </button>
                  </div>
                </div>

                {/* วันที่มีปัญหา (วันที่ขอแลก/ลา) */}
                <div>
                  <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-1.5">
                    3. วันที่ระบุของฉัน
                  </label>
                  <input
                    type="date"
                    value={reqDate}
                    onChange={(e) => setReqDate(e.target.value)}
                    required
                    className="w-full px-3 py-2 bg-white border border-emerald-200 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none text-emerald-950 font-bold shadow-sm"
                  />
                </div>

                {/* ส่วนเลือกผู้เกี่ยวข้อง (กรณีสลับเวร หรือขึ้นแทน) */}
                {(reqType === "swap" || reqType === "cover") && (
                  <div>
                    <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-1.5">
                      {reqType === "swap" ? "4. เลือกเจ้าหน้าที่ที่จะขอสลับด้วย" : "4. เลือกเจ้าหน้าที่ที่จะให้ขึ้นเวรแทน"}
                    </label>
                    <select
                      value={reqTargetStaff}
                      onChange={(e) => setReqTargetStaff(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-emerald-200 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none text-emerald-950 font-bold shadow-sm"
                    >
                      {ALL_STAFF.filter(s => s !== reqRequester).map(s => (
                        <option key={s} value={s}>{s}</option>
                      ))}
                    </select>
                  </div>
                )}

                {/* ส่วนเลือกวันแลกกลับ (เฉพาะกรณีสลับเวร) */}
                {reqType === "swap" && (
                  <div>
                    <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-1.5">
                      5. แลกกับเวรอีกฝ่ายในวันที่
                    </label>
                    <input
                      type="date"
                      value={reqTargetDate}
                      onChange={(e) => setReqTargetDate(e.target.value)}
                      required
                      className="w-full px-3 py-2 bg-white border border-emerald-200 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none text-emerald-950 font-bold shadow-sm"
                    />
                  </div>
                )}

                {/* หมายเหตุ */}
                <div>
                  <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-1.5">
                    {reqType === "leave" ? "4. เหตุผลการลาพักผ่อน" : reqType === "swap" ? "6. หมายเหตุ / เหตุผลการแลกเวร" : "5. หมายเหตุประกอบ"}
                  </label>
                  <textarea
                    placeholder="ระบุเหตุผล เช่น ติดงานแต่งญาติ, พาคุณแม่ไปตรวจสุขภาพ..."
                    value={reqNote}
                    onChange={(e) => setReqNote(e.target.value)}
                    rows={2}
                    className="w-full px-3 py-2 bg-white border border-emerald-200 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none text-gray-700 shadow-sm resize-none"
                  />
                </div>

                {/* ปุ่มส่งคำขอ */}
                <button
                  type="submit"
                  className="w-full mt-2 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white font-bold py-2.5 px-4 rounded-xl text-sm transition-all shadow hover:shadow-md flex items-center justify-center gap-2 cursor-pointer active:scale-98"
                >
                  <Bell className="w-4 h-4" />
                  <span>ส่งคำขอและแจ้งเตือนไปยังทุกคน 📢</span>
                </button>
              </form>
            </div>

            {/* ฝั่งขวา: รายชื่อคำขอและประวัติทั้งหมด (7 คอลัมน์) */}
            <div className="lg:col-span-7 flex flex-col gap-4">
              <div className="flex items-center justify-between border-b border-gray-100 pb-3 flex-wrap gap-2">
                <h3 className="font-bold text-sm sm:text-base text-gray-800 flex items-center gap-2">
                  <RefreshCcw className="w-4.5 h-4.5 text-emerald-600" />
                  <span>รายการคำขอแจ้งสลับและเปลี่ยนเวรล่าสุด</span>
                </h3>

                {/* ตัวเลือกลิสต์ข้อมูล */}
                <div className="flex bg-gray-100 p-0.5 rounded-lg border text-[11px] font-bold">
                  <button
                    onClick={() => setActiveReqTab("pending")}
                    className={`px-3 py-1 rounded-md transition-all cursor-pointer ${
                      activeReqTab === "pending"
                        ? "bg-white text-gray-800 shadow-xs"
                        : "text-gray-500 hover:text-gray-800"
                    }`}
                  >
                    รอดำเนินการ ({shiftRequests.filter(r => r.status === 'pending').length})
                  </button>
                  <button
                    onClick={() => setActiveReqTab("all")}
                    className={`px-3 py-1 rounded-md transition-all cursor-pointer ${
                      activeReqTab === "all"
                        ? "bg-white text-gray-800 shadow-xs"
                        : "text-gray-500 hover:text-gray-800"
                    }`}
                  >
                    คำขอทั้งหมด ({shiftRequests.length})
                  </button>
                </div>
              </div>

              {/* สรุปกล่องคำขอ */}
              <div className="flex-1 flex flex-col gap-3 max-h-[460px] overflow-y-auto custom-scrollbar pr-1">
                {(() => {
                  const filtered = activeReqTab === "pending" 
                    ? shiftRequests.filter(r => r.status === "pending")
                    : shiftRequests;

                  if (filtered.length === 0) {
                    return (
                      <div className="flex-1 flex flex-col items-center justify-center py-16 text-center text-gray-400 bg-gray-50/50 rounded-2xl border border-dashed border-gray-100">
                        <Clock className="w-8 h-8 mb-2 opacity-40 text-gray-400" />
                        <p className="text-xs font-semibold">ไม่มีรายการคำขอในหน้านี้</p>
                        <p className="text-[10px] mt-0.5 opacity-80">สามารถสร้างคำขอใหม่ได้ผ่านแบบฟอร์มด้านซ้าย</p>
                      </div>
                    );
                  }

                  return filtered.map((req) => {
                    const isPending = req.status === "pending";
                    const isApproved = req.status === "approved";
                    const isRejected = req.status === "rejected";
                    const formattedDate = new Date(req.date).toLocaleDateString("th-TH", { day: 'numeric', month: 'short' });
                    const formattedTargetDate = req.targetDate ? new Date(req.targetDate).toLocaleDateString("th-TH", { day: 'numeric', month: 'short' }) : "";

                    return (
                      <div
                        key={req.id}
                        className={`border rounded-2xl p-4 transition-all hover:shadow-xs flex flex-col gap-3 relative ${
                          isApproved
                            ? "bg-emerald-50/10 border-emerald-100"
                            : isRejected
                            ? "bg-red-50/10 border-red-100/50"
                            : "bg-white border-gray-100 shadow-xs"
                        }`}
                      >
                        <div className="flex items-start justify-between gap-2 flex-wrap">
                          <div className="flex items-center gap-2">
                            {/* ชนิดคีย์เวิร์ดเวร */}
                            {req.type === "swap" ? (
                              <span className="text-[10px] font-bold bg-indigo-50 text-indigo-700 border border-indigo-100 px-2.5 py-0.5 rounded-full flex items-center gap-1">
                                🤝 สลับเวร
                              </span>
                            ) : req.type === "cover" ? (
                              <span className="text-[10px] font-bold bg-sky-50 text-sky-700 border border-sky-100 px-2.5 py-0.5 rounded-full flex items-center gap-1">
                                🙋‍♀️ ให้ขึ้นแทน
                              </span>
                            ) : (
                              <span className="text-[10px] font-bold bg-orange-50 text-orange-700 border border-orange-100 px-2.5 py-0.5 rounded-full flex items-center gap-1">
                                🌴 ลาพักร้อน
                              </span>
                            )}

                            {/* วันที่ระบุ */}
                            <span className="text-xs font-bold text-gray-500">
                              ยื่นเมื่อ: {new Date(req.createdAt).toLocaleDateString("th-TH", { day: 'numeric', month: 'short' })}
                            </span>
                          </div>

                          {/* สถานะแบดจ์ */}
                          <div>
                            {isApproved && (
                              <span className="text-[10px] font-bold bg-emerald-100 text-emerald-800 px-2.5 py-1 rounded-lg flex items-center gap-1">
                                <Check className="w-3 h-3 text-emerald-600" /> อนุมัติสำเร็จ
                              </span>
                            )}
                            {isRejected && (
                              <span className="text-[10px] font-bold bg-red-100 text-red-800 px-2.5 py-1 rounded-lg flex items-center gap-1">
                                <X className="w-3 h-3 text-red-600" /> ปฏิเสธแล้ว
                              </span>
                            )}
                            {isPending && (
                              <span className="text-[10px] font-bold bg-amber-100 text-amber-800 px-2.5 py-1 rounded-lg flex items-center gap-1 animate-pulse">
                                <Clock className="w-3 h-3 text-amber-600" /> รอตัดสินใจ
                              </span>
                            )}
                          </div>
                        </div>

                        {/* รายละเอียดการแลก */}
                        <div className="bg-gray-50/50 rounded-xl p-3 text-xs flex flex-col gap-2 border border-gray-100">
                          {req.type === "swap" ? (
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 font-semibold">
                              <div className="flex items-center gap-1.5 text-emerald-900">
                                <span className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-md font-bold">{req.requester}</span>
                                <span>(เวรวันที่ {formattedDate})</span>
                              </div>
                              <span className="text-gray-400 text-center sm:block">🔄 สลับกับ 🔄</span>
                              <div className="flex items-center gap-1.5 text-indigo-900 sm:justify-end">
                                <span className="bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded-md font-bold">{req.targetStaff}</span>
                                <span>(เวรวันที่ {formattedTargetDate})</span>
                              </div>
                            </div>
                          ) : req.type === "cover" ? (
                            <div className="font-semibold text-gray-700">
                              คุณ <span className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded font-bold">{req.requester}</span> ขอให้ คุณ <span className="bg-sky-100 text-sky-800 px-2 py-0.5 rounded font-bold">{req.targetStaff}</span> ขึ้นปฏิบัติหน้าที่เวรแทนในวันที่ <span className="text-emerald-700 font-bold">{formattedDate}</span>
                            </div>
                          ) : (
                            <div className="font-semibold text-gray-700">
                              คุณ <span className="bg-orange-100 text-orange-800 px-2 py-0.5 rounded font-bold">{req.requester}</span> แจ้งขอลาพักร้อน ปฏิบัติหน้าที่หยุดงานในวันที่ <span className="text-orange-700 font-bold">{formattedDate}</span>
                            </div>
                          )}

                          {req.note && (
                            <div className="text-gray-500 font-medium text-[11px] border-t border-gray-200/50 pt-2 mt-1 italic flex items-start gap-1">
                              <span>💬 เหตุผล:</span>
                              <span>"{req.note}"</span>
                            </div>
                          )}
                        </div>

                        {/* ปุ่มตัวเลือกอนุมัติ/ปุ่มเปลี่ยนจำลอง */}
                        <div className="flex items-center justify-between gap-2 mt-1 border-t border-gray-100/60 pt-3">
                          {isPending ? (
                            <>
                              {/* 1. กรณีเป็น Admin */}
                              {isAdmin ? (
                                <div className="flex gap-2 w-full justify-end">
                                  <button
                                    onClick={() => handleRejectShiftRequest(req.id, req.requester, req.date)}
                                    className="px-3 py-1.5 border border-red-200 hover:bg-red-50 text-red-600 rounded-lg text-xs font-bold transition-all flex items-center gap-1 cursor-pointer active:scale-95"
                                  >
                                    <X className="w-3.5 h-3.5" />
                                    <span>ปฏิเสธ</span>
                                  </button>
                                  <button
                                    onClick={() => handleApproveShiftRequest(req)}
                                    className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-black transition-all flex items-center gap-1 cursor-pointer shadow-sm active:scale-95"
                                  >
                                    <Check className="w-3.5 h-3.5" />
                                    <span>อนุมัติ & เปลี่ยนเวร</span>
                                  </button>
                                </div>
                              ) : currentUser ? (
                                /* 2. กรณีสิทธิ์จำกัด (เป็นเจ้าหน้าที่ที่ล็อกอินอยู่) */
                                <div className="flex items-center justify-between w-full flex-wrap gap-2">
                                  {req.targetStaff === currentUser ? (
                                    <div className="flex items-center justify-between w-full bg-indigo-50 border border-indigo-100 rounded-xl p-2.5">
                                      <div className="text-[10px] text-indigo-800 font-semibold">
                                        👉 คุณคือเพื่อนเป้าหมายของคำขอนี้ ท่านสามารถกดยอมรับเพื่อแลก/แทนเวรได้ทันที
                                      </div>
                                      <button
                                        onClick={() => {
                                          setIsAdmin(true);
                                          handleApproveShiftRequest(req);
                                          setIsAdmin(false);
                                        }}
                                        className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-black transition-all flex items-center gap-1 cursor-pointer shadow-sm active:scale-95"
                                      >
                                        <Check className="w-3.5 h-3.5" />
                                        <span>ยอมรับ & สลับเวรสำเร็จ</span>
                                      </button>
                                    </div>
                                  ) : req.requester === currentUser ? (
                                    <div className="text-[11px] text-emerald-700 font-semibold bg-emerald-50 px-3 py-2 rounded-xl border border-emerald-100 flex items-center gap-1.5 w-full">
                                      <Clock className="w-4 h-4 text-emerald-600 shrink-0" />
                                      <span>คำขอของคุณส่งสำเร็จแล้ว รอแอดมินหรือคุณ {req.targetStaff || "ผู้เกี่ยวข้อง"} ตกลงตอบรับ</span>
                                    </div>
                                  ) : (
                                    <div className="text-[11px] text-gray-400 font-semibold italic p-1.5 bg-gray-50/50 rounded-lg w-full text-center border border-gray-100">
                                      🔒 ตารางเวรของผู้ใช้อื่น เฉพาะคุณ {req.targetStaff || req.requester} หรือแอดมินเท่านั้นที่อนุมัติได้
                                    </div>
                                  )}
                                </div>
                              ) : (
                                /* 3. กรณีโหมดทั่วไป (Viewer) เพื่อให้ทดสอบความเทพของระบบได้ทันทีโดยไม่ต้องใส่พาสเวิร์ด */
                                <div className="flex items-center justify-between w-full flex-wrap gap-2 bg-amber-50/50 border border-amber-100 rounded-xl p-2.5">
                                  <div className="text-[10px] text-amber-800 font-semibold max-w-sm">
                                    💡 โหมดทั่วไป: สามารถกดปุ่มทดลองยอมรับเพื่อจำลองการยืนยันตารางเวรและส่งแจ้งเตือนได้เลย!
                                  </div>
                                  <div className="flex gap-1.5 justify-end ml-auto shrink-0">
                                    <button
                                      onClick={() => {
                                        // Set admin to true momentarily or keep it true for simplicity of apply
                                        setIsAdmin(true);
                                        handleApproveShiftRequest(req);
                                      }}
                                      className="px-3 py-1.5 bg-amber-500 hover:bg-amber-600 active:bg-amber-700 text-white rounded-lg text-[10px] font-black transition-all flex items-center gap-1 cursor-pointer shadow-xs active:scale-95"
                                    >
                                      <CheckCircle2 className="w-3 h-3" />
                                      <span>ยืนยันคำขอทันที (Simulate Accept)</span>
                                    </button>
                                  </div>
                                </div>
                              )}
                            </>
                          ) : (
                            /* ได้รับการตัดสินแล้ว */
                            <div className="flex items-center justify-between w-full">
                              <span className="text-[10px] text-gray-400 font-medium">
                                รหัสอ้างอิงคำขอ: {req.id}
                              </span>
                              <button
                                onClick={() => handleDeleteShiftRequest(req.id)}
                                className="text-gray-400 hover:text-red-500 hover:bg-red-50 p-1 rounded-lg transition-colors cursor-pointer"
                                title="ลบคำขอออกจากระบบ"
                              >
                                <X className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  });
                })()}
              </div>
            </div>
          </div>
        </section>
          </>
        )}
      </main>

      {/* Admin / User Login Modal */}
      <AnimatePresence>
        {showAdminModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4"
          >
            <motion.div
              initial={{ scale: 0.95, y: 20 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.95, y: 20 }}
              className="bg-white rounded-2xl shadow-2xl border border-gray-100 max-w-md w-full overflow-hidden text-left"
            >
              <div className="bg-gradient-to-r from-emerald-600 to-emerald-700 p-6 text-white text-center relative">
                <button
                  type="button"
                  onClick={() => {
                    setShowAdminModal(false);
                    setPasswordInput("");
                    setPasswordError("");
                  }}
                  className="absolute top-4 right-4 text-white/80 hover:text-white hover:bg-white/10 p-1.5 rounded-lg transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
                <div className="bg-white/25 w-12 h-12 rounded-full flex items-center justify-center mx-auto mb-3 shadow-inner">
                  <Lock className="w-6 h-6 text-white" />
                </div>
                <h3 className="text-lg font-bold">ลงชื่อเข้าใช้งานระบบ (Login)</h3>
                <p className="text-emerald-100 text-xs mt-1">กรุณาเลือกประเภทสิทธิ์และกรอกรหัสผ่านเพื่อยืนยันสิทธิ์</p>
              </div>

              {/* Tabs for Login Type */}
              <div className="flex border-b border-gray-100 bg-gray-50/50 p-1">
                <button
                  type="button"
                  onClick={() => {
                    setLoginTab("staff");
                    setPasswordError("");
                  }}
                  className={`flex-1 py-3 text-xs sm:text-sm font-bold text-center transition-all cursor-pointer rounded-xl ${
                    loginTab === "staff"
                      ? "bg-white text-emerald-800 shadow-sm border-b-2 border-emerald-600"
                      : "text-gray-500 hover:text-gray-800"
                  }`}
                >
                  👤 สิทธิ์เจ้าหน้าที่ในทีม (Staff)
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setLoginTab("admin");
                    setPasswordError("");
                  }}
                  className={`flex-1 py-3 text-xs sm:text-sm font-bold text-center transition-all cursor-pointer rounded-xl ${
                    loginTab === "admin"
                      ? "bg-white text-amber-600 shadow-sm border-b-2 border-amber-500"
                      : "text-gray-500 hover:text-gray-800"
                  }`}
                >
                  🛡️ สิทธิ์ผู้ดูแลระบบ (Admin)
                </button>
              </div>

              <form
                onSubmit={(e) => {
                  e.preventDefault();
                  if (loginTab === "staff") {
                    if (passwordInput === "1234") {
                      setCurrentUser(selectedLoginStaff);
                      setSelectedStaffFilter(selectedLoginStaff);
                      setIsAdmin(false);
                      setIsAdminConsoleOpen(false);
                      setShowAdminModal(false);
                      setPasswordInput("");
                      setPasswordError("");
                      showToast(`เข้าสู่ระบบในฐานะคุณ "${selectedLoginStaff}" สำเร็จ! สามารถแก้ไขและส่งคำขอเปลี่ยนเวรเฉพาะของตนเองได้ในหน้าผู้ใช้นี้`, false, new Date().toLocaleDateString("th-TH"));
                    } else {
                      setPasswordError("รหัสผ่านไม่ถูกต้อง สำหรับสิทธิ์เจ้าหน้าที่กรุณาใช้รหัส 1234");
                    }
                  } else {
                    if (passwordInput === "1234" || passwordInput.toLowerCase() === "admin") {
                      setIsAdmin(true);
                      setIsAdminConsoleOpen(true);
                      setCurrentUser(null);
                      setShowAdminModal(false);
                      setPasswordInput("");
                      setPasswordError("");
                      showToast("เข้าสู่ระบบผู้ดูแลระบบ (Admin) สำเร็จ ยินดีต้อนรับ!", false, new Date().toLocaleDateString("th-TH"));
                    } else {
                      setPasswordError("รหัสผ่านไม่ถูกต้อง สำหรับแอดมินกรุณาลองใหม่อีกครั้ง");
                    }
                  }
                }}
                className="p-6 flex flex-col gap-4 text-left"
              >
                {loginTab === "staff" && (
                  <div>
                    <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">
                      เลือกชื่อของคุณในทีม
                    </label>
                    <select
                      value={selectedLoginStaff}
                      onChange={(e) => setSelectedLoginStaff(e.target.value)}
                      className="w-full px-3 py-2.5 bg-gray-50 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500 focus:bg-white focus:outline-none transition-all font-semibold text-gray-800"
                    >
                      {ALL_STAFF.map(s => (
                        <option key={s} value={s}>{s}</option>
                      ))}
                    </select>
                  </div>
                )}

                <div>
                  <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">
                    {loginTab === "staff" ? "รหัสผ่านเจ้าหน้าที่ (รหัสคือ 1234)" : "รหัสผ่านแอดมิน"}
                  </label>
                  <div className="relative">
                    <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-gray-400 pointer-events-none">
                      <Lock className="w-4 h-4" />
                    </span>
                    <input
                      type="password"
                      placeholder="กรอกรหัสผ่าน..."
                      value={passwordInput}
                      onChange={(e) => {
                        setPasswordInput(e.target.value);
                        if (passwordError) setPasswordError("");
                      }}
                      className={`w-full pl-10 pr-4 py-2.5 bg-gray-50 border rounded-xl text-sm focus:ring-2 focus:ring-emerald-500 focus:bg-white focus:outline-none transition-all ${
                        passwordError ? "border-red-300 focus:ring-red-500" : "border-gray-200"
                      }`}
                      autoFocus
                    />
                  </div>
                  {passwordError && (
                    <p className="text-red-500 text-xs font-semibold mt-1.5 flex items-center">
                      <AlertCircle className="w-3.5 h-3.5 mr-1 shrink-0" />
                      {passwordError}
                    </p>
                  )}
                </div>

                <div className="bg-emerald-50 border border-emerald-100 rounded-xl p-3 text-xs text-emerald-800 leading-relaxed">
                  <span className="font-bold">💡 ข้อมูลสิทธิ์การใช้งาน:</span>
                  <p className="mt-0.5 opacity-90">
                    {loginTab === "staff" 
                      ? "เมื่อล็อกอินสิทธิ์เจ้าหน้าที่ตามชื่อของตนเอง ท่านจะสามารถสลับ/แก้ไขตารางเวรในปฏิทินเฉพาะแถวที่เป็นชื่อของท่าน และส่งคำขอสลับเวรแบบล็อกชื่อผู้ร้องขอได้โดยอัตโนมัติ"
                      : "สิทธิ์แอดมินช่วยให้ท่านแก้ไข สลับเวร พักร้อน และจัดการข้อมูลของเจ้าหน้าที่ทุกคนในทีมได้ทุกตาราง"
                    }
                  </p>
                </div>

                <div className="flex gap-3 mt-2">
                  <button
                    type="button"
                    onClick={() => {
                      setShowAdminModal(false);
                      setPasswordInput("");
                      setPasswordError("");
                    }}
                    className="flex-1 py-2.5 border border-gray-200 text-gray-700 font-semibold rounded-xl text-sm hover:bg-gray-50 active:bg-gray-100 transition-colors cursor-pointer"
                  >
                    ยกเลิก
                  </button>
                  <button
                    type="submit"
                    className="flex-1 py-2.5 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white font-bold rounded-xl text-sm transition-all shadow-md flex items-center justify-center gap-1.5 cursor-pointer"
                  >
                    <Unlock className="w-4 h-4" />
                    <span>ยืนยัน</span>
                  </button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Floating Toasts */}
      <div className="fixed bottom-6 right-6 z-50 flex flex-col gap-3 pointer-events-none">
        <AnimatePresence>
          {toasts.map((toast) => (
            <motion.div
              key={toast.id}
              initial={{ opacity: 0, y: 30, scale: 0.95 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, scale: 0.9, transition: { duration: 0.2 } }}
              className="bg-white border text-left border-emerald-100 rounded-xl shadow-2xl p-4 sm:p-5 max-w-sm flex items-start pointer-events-auto"
            >
              <div className="bg-emerald-100 p-2 rounded-full mr-3 shrink-0">
                <Bell className="w-5 h-5 text-emerald-600" />
              </div>
              <div className="flex-1 pr-2">
                <h4 className="font-bold text-gray-900 text-sm flex items-center">
                  วันที่ {toast.dateStr}
                </h4>
                <p className="text-gray-600 text-sm whitespace-pre-line mt-1.5 leading-relaxed">
                  {toast.message}
                </p>
                
                {toast.inCharge && (
                  <div className="mt-3 text-xs font-semibold px-2.5 py-1.5 rounded inline-flex items-center w-fit border bg-indigo-50 text-indigo-700 border-indigo-200">
                     <FileText className="w-3.5 h-3.5 mr-1.5" /> ได้รับสิทธิ์อินชาร์ตเอกสาร (Doc.)
                  </div>
                )}
              </div>
              <button
                onClick={() => setToasts((prev) => prev.filter((t) => t.id !== toast.id))}
                className="text-gray-400 hover:text-gray-600 hover:bg-gray-100 p-1 rounded-md transition-colors"
                aria-label="ปิดการแจ้งเตือน"
              >
                <X className="w-4 h-4" />
              </button>
            </motion.div>
          ))}
        </AnimatePresence>
      </div>
      
      {/* Scrollbar Customization Base */}
      <style dangerouslySetInnerHTML={{__html: `
        .custom-scrollbar::-webkit-scrollbar {
          width: 8px;
          height: 10px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: #f1f5f9;
          border-radius: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #cbd5e1;
          border-radius: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #94a3b8;
        }
      `}} />
    </div>
  );
}

