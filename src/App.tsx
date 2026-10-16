import { useEffect, useMemo, useRef, useState } from "react";
import {
  Activity, AlertTriangle, BookOpen, Clock3, Check, CheckCircle2, ChevronDown,
  ClipboardCheck, CloudUpload, Database, Download, Eye, FileCheck2, FileDown,
  FileText, Grid2X2, GraduationCap, Home, Layers3, Menu, MoreHorizontal, Pencil,
  Plus, RefreshCw, Search, Settings2, ShieldCheck, Shuffle, Smartphone, Sparkles,
  Table2, Trash2, Upload, X
} from "lucide-react";

type QuestionType = "multiple_choice" | "true_false" | "short_answer" | "essay";
type Level = "Biết" | "Hiểu" | "Vận dụng";
type MixMode = "g0" | "g1" | "g2" | "g3";
type Page = "overview" | "questions" | "matrix" | "mixer" | "validation" | "exports" | "online" | "settings";
type Option = { id: string; text: string };
type Question = {
  id: string; number: number; type: QuestionType; topic: string; level: Level;
  text: string; options: Option[]; answer: string | string[] | null; points: number; group: MixMode;
};
type NavItem = { key: Page; label: string; icon: typeof Home };
type NavGroup = { name: string; items: NavItem[] };

const TYPE_LABELS: Record<QuestionType, string> = {
  multiple_choice: "Trắc nghiệm", true_false: "Đúng / Sai",
  short_answer: "Trả lời ngắn", essay: "Tự luận"
};
const PAGE_TITLES: Record<Page, string> = {
  overview: "Tổng quan", questions: "Ngân hàng câu hỏi", matrix: "Ma trận & đặc tả",
  mixer: "Trộn đề thi", validation: "Kiểm định đề", exports: "Xuất hồ sơ",
  online: "Thi trực tuyến", settings: "Cài đặt"
};
const NAV: NavGroup[] = [
  { name: "KHÔNG GIAN LÀM VIỆC", items: [
    { key: "overview", label: "Tổng quan", icon: Home },
    { key: "questions", label: "Ngân hàng câu hỏi", icon: Database },
    { key: "matrix", label: "Ma trận & đặc tả", icon: Grid2X2 }
  ]},
  { name: "TẠO ĐỀ THI", items: [
    { key: "mixer", label: "Trộn đề thi", icon: Shuffle },
    { key: "validation", label: "Kiểm định đề", icon: ShieldCheck },
    { key: "exports", label: "Xuất hồ sơ", icon: FileDown }
  ]},
  { name: "MỞ RỘNG", items: [
    { key: "online", label: "Thi trực tuyến", icon: GraduationCap },
    { key: "settings", label: "Cài đặt", icon: Settings2 }
  ]}
];
const SEED: Question[] = [
  { id: "Q-HOA-001", number: 1, type: "multiple_choice", topic: "Este – Lipit", level: "Biết",
    text: "Este đơn chức, mạch hở có công thức chung là",
    options: [{id:"a",text:"CₙH₂ₙO₂ (n ≥ 2)"},{id:"b",text:"CₙH₂ₙ₊₂O"},{id:"c",text:"CₙH₂ₙ₋₂O₂"},{id:"d",text:"CₙH₂ₙO"}],
    answer: "a", points: 0.25, group: "g3" },
  { id: "Q-HOA-002", number: 2, type: "multiple_choice", topic: "Este – Lipit", level: "Hiểu",
    text: "Chất nào sau đây có phản ứng tráng bạc?",
    options: [{id:"e",text:"CH₃COOCH₃"},{id:"f",text:"HCOOCH₃"},{id:"g",text:"CH₃COOC₂H₅"},{id:"h",text:"C₂H₅COOCH₃"}],
    answer: "f", points: 0.25, group: "g1" },
  { id: "Q-HOA-003", number: 3, type: "true_false", topic: "Este – Lipit", level: "Hiểu",
    text: "Xét các phát biểu về chất béo.",
    options: [{id:"i",text:"Chất béo là trieste của glycerol với axit béo."},{id:"j",text:"Dầu ăn thường chứa nhiều gốc axit béo không no."},{id:"k",text:"Phản ứng xà phòng hóa tạo glycerol và muối."},{id:"l",text:"Mọi chất béo đều tan tốt trong nước."}],
    answer: ["i","j","k"], points: 1, group: "g0" },
  { id: "Q-HOA-004", number: 4, type: "short_answer", topic: "Este – Lipit", level: "Vận dụng",
    text: "Thủy phân hoàn toàn 0,1 mol este đơn chức X bằng NaOH vừa đủ. Số mol NaOH đã phản ứng là bao nhiêu?",
    options: [], answer: "0,1", points: 0.5, group: "g1" },
  { id: "Q-HOA-005", number: 5, type: "essay", topic: "Este – Lipit", level: "Vận dụng",
    text: "Trình bày phương pháp hóa học phân biệt ethyl acetate và acetic acid. Viết phương trình hóa học minh họa.",
    options: [], answer: null, points: 1, group: "g0" },
  { id: "Q-HOA-006", number: 6, type: "multiple_choice", topic: "Cân bằng hóa học", level: "Hiểu",
    text: "Khi tăng áp suất, cân bằng khí nào chuyển dịch theo chiều thuận?",
    options: [{id:"m",text:"N₂(g) + 3H₂(g) ⇌ 2NH₃(g)"},{id:"n",text:"H₂(g) + I₂(g) ⇌ 2HI(g)"},{id:"o",text:"H₂(g) ⇌ 2H(g)"},{id:"p",text:"CaCO₃(s) ⇌ CaO(s) + CO₂(g)"}],
    answer: "m", points: 0.25, group: "g3" }
];

function shuffle<T>(items: T[]): T[] {
  const result = [...items];
  for (let i = result.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [result[i], result[j]] = [result[j], result[i]];
  }
  return result;
}
function downloadJson(filename: string, data: unknown): void {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url; anchor.download = filename; anchor.click();
  URL.revokeObjectURL(url);
}
function loadSavedQuestions(): Question[] {
  try {
    const raw = localStorage.getItem("ymp-questions");
    if (!raw) return SEED;
    const parsed: unknown = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as Question[] : SEED;
  } catch { return SEED; }
}
function Pill({ type }: { type: QuestionType }) {
  return <span className={`pill ${type}`}>{TYPE_LABELS[type]}</span>;
