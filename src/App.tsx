import { useEffect, useMemo, useRef, useState } from "react";
import {
  Activity, AlertTriangle, BookOpen, Clock3, Check, CheckCircle2, ChevronDown,
  ClipboardCheck, CloudUpload, Database, Download, Eye, FileCheck2, FileDown,
  FileText, Grid2X2, GraduationCap, Home, Layers3, Menu, MoreHorizontal, Pencil,
  Plus, RefreshCw, Search, Settings2, ShieldCheck, Shuffle, Smartphone, Sparkles,
  Table2, Trash2, Upload, X
} from "lucide-react";
import { apiFetch } from "./services/api";

type QuestionType = "multiple_choice" | "true_false" | "short_answer" | "essay";
type Level = "Biết" | "Hiểu" | "Vận dụng";
type MixMode = "g0" | "g1" | "g2" | "g3";
type Page = "overview" | "questions" | "matrix" | "mixer" | "validation" | "exports" | "online" | "settings";
type Option = { id: string; text: string };
type Question = {
  id: string; number: number; type: QuestionType; topic: string; level: Level;
  text: string; options: Option[]; answer: string | string[] | null; points: number; group: MixMode;
  section?: string; section_title?: string; source_number?: number; has_image?: boolean; has_table?: boolean;
};
type ParsedSection = { id: string; title: string; question_count: number; question_numbers?: number[]; type: string; group?: MixMode; image_count: number; table_count: number };
type ParsedDocx = { source_filename: string; question_count: number; image_count: number; table_count: number; sections?: string[]; section_summary: ParsedSection[]; questions: Question[]; warnings: string[] };
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
}
function Stat({ icon: Icon, label, value, note }: { icon: typeof Home; label: string; value: string; note: string }) {
  return <article className="stat"><div>{label}<span><Icon size={18}/></span></div><b>{value}</b><small>{note}</small></article>;
}
function PanelHead({ title, desc, action, onAction }: { title: string; desc: string; action?: string; onAction?: () => void }) {
  return <header className="panelhead"><div><h2>{title}</h2><p>{desc}</p></div>{action && <button onClick={onAction}>{action} →</button>}</header>;
}
function Notice({ title, text, warning = false }: { title: string; text: string; warning?: boolean }) {
  return <div className={`notice ${warning ? "warn" : ""}`}><AlertTriangle size={19}/><div><b>{title}</b><p>{text}</p></div></div>;
}

export default function App() {
  const [page, setPage] = useState<Page>("overview");
  const [menuOpen, setMenuOpen] = useState(false);
  const [questions, setQuestions] = useState<Question[]>(loadSavedQuestions);
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState<"all" | QuestionType>("all");
  const [mode, setMode] = useState<MixMode>("g3");
  const [questionCount, setQuestionCount] = useState(6);
  const [codeCount, setCodeCount] = useState(4);
  const [examCode, setExamCode] = useState("001");
  const [mixedCodes, setMixedCodes] = useState<Question[][]>([]);
  const [activeCode, setActiveCode] = useState(0);
  const [sourceDocx, setSourceDocx] = useState<File | null>(null);
  const [parsedDocx, setParsedDocx] = useState<ParsedDocx | null>(null);
  const [sectionModes, setSectionModes] = useState<Record<string, MixMode>>({});
  const [headerConfig, setHeaderConfig] = useState({
    department: "SỞ GD&ĐT ĐỒNG NAI", school: "THPT THỐNG NHẤT B",
    exam_title: "ĐỀ KIỂM TRA", grade: "12", school_year: "2026-2027",
    subject: "", duration: ""
  });
  const [isImporting, setIsImporting] = useState(false);
  const [isExportingSource, setIsExportingSource] = useState(false);
  const [toast, setToast] = useState("");
  const [modalOpen, setModalOpen] = useState(false);
  const [editing, setEditing] = useState<Question | null>(null);
  const [draft, setDraft] = useState({ text: "", topic: "Chưa phân loại", type: "multiple_choice" as QuestionType, level: "Biết" as Level, points: "0.25" });
  const fileInput = useRef<HTMLInputElement>(null);

  useEffect(() => { try { localStorage.setItem("ymp-questions", JSON.stringify(questions)); } catch { setToast("Không thể lưu trên trình duyệt này."); } }, [questions]);
  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(""), 4000);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const filteredQuestions = useMemo(() => questions.filter(q =>
    `${q.text} ${q.topic} ${q.id}`.toLowerCase().includes(search.toLowerCase()) &&
    (filter === "all" || q.type === filter)
  ), [questions, search, filter]);
  const currentCode = mixedCodes[activeCode] ?? [];

  function openQuestionEditor(question?: Question) {
    setEditing(question ?? null);
    setDraft(question
      ? { text: question.text, topic: question.topic, type: question.type, level: question.level, points: String(question.points) }
      : { text: "", topic: "Chưa phân loại", type: "multiple_choice", level: "Biết", points: "0.25" });
    setModalOpen(true);
  }
  function saveQuestion() {
    if (!draft.text.trim()) { setToast("Vui lòng nhập nội dung câu hỏi."); return; }
    const points = Number(draft.points);
    if (!Number.isFinite(points) || points < 0) { setToast("Điểm phải là số không âm."); return; }
    if (editing) {
      setQuestions(prev => prev.map(q => q.id === editing.id ? { ...q, ...draft, points } : q));
    } else {
      const id = `Q-${Date.now().toString(36).toUpperCase()}`;
      const options: Option[] = draft.type === "multiple_choice"
        ? ["Phương án A", "Phương án B", "Phương án C", "Phương án D"].map((text, i) => ({ id: `${id}-O${i + 1}`, text }))
        : [];
      setQuestions(prev => [...prev, { id, number: prev.length + 1, ...draft, points, options, answer: null, group: "g0" }]);
    }
    setModalOpen(false);
    setToast("Đã lưu câu hỏi. Vui lòng xác nhận đáp án riêng trước khi phát hành.");
  }
  function createMixedCodes() {
    if (questions.length === 0) { setToast("Chưa có câu hỏi. Hãy tải đề Word gốc trước."); return; }
    const amount = sourceDocx ? questions.length : Math.min(Math.max(1, questionCount), questions.length);
    const output: Question[][] = [];
    for (let c = 0; c < Math.min(20, Math.max(1, codeCount)); c++) {
      let list = questions.slice(0, amount).map(q => ({
        ...q,
        options: !sourceDocx && (mode === "g2" || mode === "g3") && q.type === "multiple_choice" ? shuffle(q.options) : [...q.options]
      }));
      if (sourceDocx) {
        // Respect the marker-derived or teacher-overridden mode independently for each part.
        const sections = [...new Set(list.map(q => q.section ?? "P0"))];
        list = sections.flatMap(section => {
          const sectionMode = sectionModes[section] ?? list.find(q => (q.section ?? "P0") === section)?.group ?? "g0";
          const sectionQuestions = list.filter(q => (q.section ?? "P0") === section);
          return sectionMode === "g1" || sectionMode === "g3" ? shuffle(sectionQuestions) : sectionQuestions;
        });
      } else if (mode === "g1" || mode === "g3") list = shuffle(list);
      output.push(list.map((q, i) => ({ ...q, number: i + 1 })));
    }
    setMixedCodes(output); setActiveCode(0); setPage("mixer");
    setToast(sourceDocx
      ? `Đã tạo ${output.length} thứ tự mã đề từ DOCX gốc. Dùng nút Tải DOCX để xuất từng mã; hình/công thức lấy trực tiếp từ tệp nguồn.`
      : `Đã tạo ${output.length} mã đề xem trước.`);
  }
  async function importFile(file: File) {
    setIsImporting(true);
    try {
      if (file.name.toLowerCase().endsWith(".json")) {
        const parsed: unknown = JSON.parse(await file.text());
        const imported = Array.isArray(parsed) ? parsed : (parsed && typeof parsed === "object" && "questions" in parsed ? (parsed as { questions?: unknown }).questions : null);
        if (!Array.isArray(imported)) throw new Error("Tệp JSON cần là một mảng câu hỏi hoặc có thuộc tính questions.");
        const normalized = imported.map((item, i) => {
          const q = item as Partial<Question>;
          return { ...q, id: q.id || `Q-IMPORT-${Date.now()}-${i}`, number: i + 1,
            type: q.type ?? "multiple_choice", topic: q.topic ?? "Chưa phân loại", level: q.level ?? "Biết",
            text: q.text ?? "", options: q.options ?? [], answer: q.answer ?? null, points: q.points ?? 0.25, group: q.group ?? "g0" } as Question;
        });
        setSourceDocx(null); setParsedDocx(null); setQuestions(normalized); setMixedCodes([]);
        setToast(`Đã nhập ${normalized.length} câu hỏi từ JSON.`);
      } else if (file.name.toLowerCase().endsWith(".docx")) {
        const form = new FormData(); form.append("file", file);
        const response = await apiFetch("/api/exams/parse", { method: "POST", body: form });
        const result = await response.json() as Omit<ParsedDocx, "source_filename"> & { source_filename?: string };
        if (!result.questions?.length) {
          const warnings = result.warnings?.join(" ") ?? "";
          throw new Error(`Chưa nhận diện được câu hỏi trong DOCX. ${warnings} Tệp vẫn chưa được đưa vào ngân hàng.`);
        }
        const normalized = result.questions.map((q, i) => ({
          ...q, id: q.id || `Q-DOCX-${Date.now()}-${i}`, number: i + 1,
          type: q.type ?? "multiple_choice", topic: q.topic ?? "Chưa phân loại", level: q.level ?? "Biết",
          text: q.text ?? "", options: q.options ?? [], answer: q.answer ?? null, points: q.points ?? 0, group: (q.group ?? "g0") as MixMode
        }));
        setQuestions(normalized); // A source DOCX replaces demo data; do not mix unrelated seed questions.
        setSourceDocx(file);
        setParsedDocx({ ...result, source_filename: result.source_filename ?? file.name, questions: normalized });
        setMode("g0");
        setSectionModes(Object.fromEntries((result.section_summary ?? []).map((section: ParsedSection) => [section.id, section.group ?? "g0"])) as Record<string, MixMode>);
        setMixedCodes([]); setActiveCode(0);
        setQuestionCount(normalized.length);
        setPage("mixer");
        setToast(`Đã nhận ${normalized.length} câu trong ${result.section_summary?.length ?? result.sections?.length ?? 0} phần. Đề gốc được giữ để xuất DOCX không dựng lại hình/công thức.`);
      } else {
        throw new Error("Chỉ nhận DOCX hoặc JSON. PDF/OCR chưa được tích hợp.");
      }
    } catch (error) {
      setToast(error instanceof Error ? error.message : "Không thể nhập tệp.");
    } finally {
      setIsImporting(false);
      if (fileInput.current) fileInput.current.value = "";
    }
  }
  async function exportSourceDocx() {
    if (!sourceDocx || !mixedCodes[activeCode]) {
      setToast("Hãy tải DOCX gốc và tạo thứ tự mã đề trước.");
      return;
    }
    setIsExportingSource(true);
    try {
      const order = mixedCodes[activeCode].map(q => q.id);
      const form = new FormData(); form.append("file", sourceDocx);
      const code = examCode.trim() || String(activeCode + 1).padStart(3, "0");
      const header = { ...headerConfig, exam_code: code };
      const answers = currentCode.map((q, i) => ({ number: i + 1, id: q.id, type: q.type, answer: getAnswerLabel(q), points: q.points }));
      const params = new URLSearchParams({ question_order: JSON.stringify(order), header: JSON.stringify(header), section_modes: JSON.stringify(sectionModes), answers: JSON.stringify(answers), exam_code: code, seed: String(activeCode + 1) });
      const response = await apiFetch(`/api/exams/export-bundle?${params.toString()}`, { method: "POST", body: form });
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `Young-Mix-Pro-ma-${(examCode.trim() || String(activeCode + 1).padStart(3, "0")).replace(/[^a-zA-Z0-9_-]/g, "-")}.zip`;
      document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
      setToast(`Đã tải ZIP gồm đề Word và đáp án Excel cho mã đề ${examCode.trim() || String(activeCode + 1).padStart(3, "0")}. Hãy đối chiếu đáp án và kiểm tra bố cục trước khi phát hành.`);
    } catch (error) {
      setToast(error instanceof Error ? error.message : "Không thể xuất DOCX từ đề gốc.");
    } finally {
      setIsExportingSource(false);
    }
  }
  function exportBank() {
    downloadJson("young-mix-pro-bank.json", { version: 1, exportedAt: new Date().toISOString(), questions });
  }
  function getAnswerLabel(q: Question): string {
    if (typeof q.answer === "string" && q.type === "multiple_choice") {
      const index = q.options.findIndex(option => option.id === q.answer);
      return index >= 0 ? String.fromCharCode(65 + index) : "Chưa xác nhận";
    }
    if (Array.isArray(q.answer)) return q.answer.join(", ");
    return q.answer ?? "Chưa xác nhận";
  }

  return <div className="app">
    <aside className={`sidebar ${menuOpen ? "open" : ""}`}>
      <div className="brand"><span className="brandmark"><Layers3/></span><span><b>Young Mix <i>Pro</i></b><small>EXAM WORKSPACE</small></span><button className="icon close" onClick={() => setMenuOpen(false)} aria-label="Đóng menu"><X/></button></div>
      <button className="workspace" onClick={() => setPage("settings")}><span>YM</span><b>Không gian giáo viên<small>Gói cá nhân · Bản thử nghiệm</small></b><ChevronDown size={15}/></button>
      {NAV.map(group => <div className="navgroup" key={group.name}><label>{group.name}</label>{group.items.map(item => {
        const Icon = item.icon;
        return <button key={item.key} className={`navitem ${page === item.key ? "active" : ""}`} onClick={() => { setPage(item.key); setMenuOpen(false); }}>
          <Icon size={18}/>{item.label}{item.key === "validation" && <small>!</small>}
        </button>;
      })}</div>)}
      <div className="sidebottom"><div className="helper"><Sparkles size={18}/><b>Trợ lý ra đề</b><p>AI phân loại câu hỏi sẽ được kết nối ở giai đoạn tiếp theo.</p></div><div className="profile"><span>GV</span><b>Giáo viên<small>Chế độ demo</small></b><MoreHorizontal/></div></div>
    </aside>
    {menuOpen && <button className="scrim" onClick={() => setMenuOpen(false)} aria-label="Đóng menu"/>}
    <main>
      <header className="topbar"><div className="crumb"><button className="icon burger" onClick={() => setMenuOpen(true)} aria-label="Mở menu"><Menu/></button><span>Không gian làm việc</span> / <b>{PAGE_TITLES[page]}</b></div><div className="topright"><span className="saved"><i/> Đã lưu trên thiết bị</span><span className="avatar">GV</span></div></header>
      <div className="content">
        {page === "overview" && <>
          <div className="heading"><div><label>KHÔNG GIAN LÀM VIỆC</label><h1>Chào thầy, sẵn sàng tạo đề mới?</h1><p>Quản lý câu hỏi, xây dựng ma trận và tạo nhiều mã đề trong một không gian.</p></div><button className="btn primary" onClick={() => setPage("mixer")}><Plus size={17}/> Tạo đề thi mới</button></div>
          <div className="stats"><Stat icon={Database} label="Ngân hàng câu hỏi" value={String(questions.length).padStart(2,"0")} note="Câu hỏi trên thiết bị"/><Stat icon={FileCheck2} label="Mã đề đã tạo" value={String(mixedCodes.length).padStart(2,"0")} note="Trong phiên hiện tại"/><Stat icon={AlertTriangle} label="Cần rà soát" value={String(questions.filter(q => !q.answer).length).padStart(2,"0")} note="Câu chưa có đáp án"/><Stat icon={GraduationCap} label="Thi trực tuyến" value="—" note="Chưa kết nối backend thi"/></div>
          <div className="dashgrid">
            <section className="panel"><PanelHead title="Câu hỏi gần đây" desc="Những câu hỏi trong ngân hàng hiện tại" action="Xem tất cả" onAction={() => setPage("questions")}/><div className="tablewrap"><table><thead><tr><th>NỘI DUNG</th><th>DẠNG CÂU</th><th>MỨC ĐỘ</th><th>ĐIỂM</th><th/></tr></thead><tbody>{questions.slice(0,5).map(q => <tr key={q.id}><td><small className="qid">{q.id}</small><b className="questiontext">{q.text}</b><small className="muted">{q.topic}</small></td><td><Pill type={q.type}/></td><td><span className="level">{q.level}</span></td><td>{q.points}</td><td><button className="icon" onClick={() => openQuestionEditor(q)} aria-label="Sửa câu hỏi"><Pencil size={15}/></button></td></tr>)}</tbody></table></div><div className="panelfoot">{questions.length} câu hỏi trong ngân hàng <button onClick={() => setPage("questions")}>Mở danh sách →</button></div></section>
            <section className="panel quick"><PanelHead title="Bắt đầu nhanh" desc="Chọn tác vụ cần thực hiện"/>{[
              { icon: CloudUpload, title: "Nhập đề từ tệp", desc: "JSON hoặc DOCX", action: () => fileInput.current?.click() },
              { icon: Grid2X2, title: "Tạo ma trận đề", desc: "Nội dung và mức độ", action: () => setPage("matrix") },
              { icon: Shuffle, title: "Trộn nhiều mã đề", desc: "Đảo câu và phương án", action: () => setPage("mixer") },
              { icon: ClipboardCheck, title: "Kiểm định đề thi", desc: "Rà soát trước phát hành", action: () => setPage("validation") }
            ].map(item => { const Icon = item.icon; return <button className="quickrow" key={item.title} onClick={item.action}><span className="quickicon"><Icon size={19}/></span><span><b>{item.title}</b><small>{item.desc}</small></span><strong>→</strong></button>; })}</section>
          </div>
          <section className="panel roadmap"><PanelHead title="Quy trình tạo đề" desc="Tiến trình gợi ý cho một bộ đề hoàn chỉnh"/><div className="steps">{["Ngân hàng câu hỏi","Thiết lập trộn","Kiểm định đề","Xuất hồ sơ"].map((text,i)=><div className={i===0?"done":i===1?"now":""} key={text}><span>{i===0?<Check size={15}/>:i+1}</span><b>{text}</b><small>{["Đã có dữ liệu mẫu","Chọn quy tắc và số mã","Rà soát đáp án, nội dung","Word, PDF, Excel"][i]}</small></div>)}</div></section>
        </>}

        {page === "questions" && <>
          <div className="heading"><div><label>QUẢN LÝ HỌC LIỆU</label><h1>Ngân hàng câu hỏi</h1><p>Tìm kiếm, phân loại và chỉnh sửa câu hỏi.</p></div><div className="headbuttons"><button className="btn secondary" onClick={exportBank}><Download size={16}/> Xuất JSON</button><button className="btn primary" onClick={() => openQuestionEditor()}><Plus size={17}/> Thêm câu hỏi</button></div></div>
          <div className="filters"><div className="search"><Search size={17}/><input placeholder="Tìm câu hỏi, mã hoặc chủ đề..." value={search} onChange={e => setSearch(e.target.value)}/></div><select value={filter} onChange={e => setFilter(e.target.value as "all" | QuestionType)}><option value="all">Tất cả dạng câu</option>{Object.entries(TYPE_LABELS).map(([key,label]) => <option value={key} key={key}>{label}</option>)}</select><button className="btn secondary" onClick={() => fileInput.current?.click()}><Upload size={16}/> Nhập tệp</button></div>
          <div className="qgrid">{filteredQuestions.map(q => <article className="qcard" key={q.id}><div className="qtop"><small className="qid">{q.id}</small><span><button className="icon" onClick={() => openQuestionEditor(q)} aria-label="Sửa"><Pencil size={15}/></button><button className="icon" onClick={() => { setQuestions(prev => prev.filter(item => item.id !== q.id)); setToast("Đã xóa câu hỏi."); }} aria-label="Xóa"><Trash2 size={15}/></button></span></div><p>{q.text}</p>{q.options.length > 0 && <div className="options">{q.options.map((option,i) => <span key={option.id}><b>{String.fromCharCode(65+i)}.</b> {option.text}</span>)}</div>}<div className="qfoot"><Pill type={q.type}/><span className="level">{q.topic}</span><span className="level">{q.level}</span><small>{q.points} điểm</small></div></article>)}{filteredQuestions.length === 0 && <div className="empty">Không tìm thấy câu hỏi phù hợp.</div>}</div>
        </>}

        {page === "matrix" && <>
          <div className="heading"><div><label>THIẾT KẾ ĐÁNH GIÁ</label><h1>Ma trận & bản đặc tả</h1><p>Khung giao diện để phát triển ma trận theo nội dung, mức độ và dạng câu.</p></div><button className="btn primary" onClick={() => setToast("Ma trận hiện là bản minh họa; chức năng chỉnh sửa và liên kết câu hỏi chưa hoàn thiện.")}><Plus size={17}/> Tạo ma trận</button></div>
          <Notice title="Khung ma trận đang ở chế độ minh họa" text="MVP chưa đồng bộ ma trận với câu hỏi. Cần liên kết từng dòng ma trận với câu hỏi và kiểm tra tổng điểm trước khi dùng chính thức."/>
          <section className="panel"><PanelHead title="Ma trận minh họa — Hóa học 12" desc="Ví dụ giao diện, không phải ma trận chính thức"/><div className="tablewrap"><table className="matrix"><thead><tr><th>CHỦ ĐỀ / NỘI DUNG</th><th>BIẾT</th><th>HIỂU</th><th>VẬN DỤNG</th><th>TỔNG</th><th>TỈ LỆ</th></tr></thead><tbody>{[
            {topic:"Este – Lipit",know:2,understand:2,apply:1},{topic:"Cân bằng hóa học",know:1,understand:2,apply:1}
          ].map(row => <tr key={row.topic}><td>{row.topic}</td><td>{row.know}</td><td>{row.understand}</td><td>{row.apply}</td><td>{row.know+row.understand+row.apply}</td><td>{Math.round((row.know+row.understand+row.apply)/9*100)}%</td></tr>)}<tr><td><b>Tổng minh họa</b></td><td>3</td><td>4</td><td>2</td><td>9</td><td>100%</td></tr></tbody></table></div></section>
          <div className="twocol"><section className="panel"><PanelHead title="Bản đặc tả" desc="Yêu cầu cần đạt và cấu trúc câu hỏi"/><div className="empty"><BookOpen size={24}/><b>Chưa liên kết bản đặc tả</b><small>Liên kết với nhóm câu hỏi sẽ được bổ sung.</small></div></section><section className="panel"><PanelHead title="Kiểm tra phân bổ" desc="Các quy tắc dự kiến"/><div className="checkline"><CheckCircle2/> Tổng câu theo chủ đề <b>Ví dụ</b></div><div className="checkline"><CheckCircle2/> Tổng điểm toàn bài <b>Ví dụ</b></div><div className="checkline warning"><AlertTriangle/> Đối chiếu câu hỏi thực tế <b>Chưa nối</b></div></section></div>
        </>}

        {page === "mixer" && <>
          <div className="heading"><div><label>TẠO NHIỀU PHIÊN BẢN</label><h1>Trộn đề thi</h1><p>Thiết lập quy tắc đảo câu và phương án, sau đó tạo mã đề xem trước.</p></div><span className="chip"><Shuffle size={14}/> Engine MVP</span></div>
          {parsedDocx && sourceDocx && <section className="panel source-docx-panel">
            <div className="source-docx-head"><div><span className="source-docx-badge"><FileCheck2 size={16}/> ĐỀ WORD GỐC ĐÃ NHẬN</span><h2>{parsedDocx.source_filename}</h2><p>Đang dùng chính DOCX này làm nguồn xuất; không dựng lại hình ảnh và công thức từ bản xem trước.</p></div><button className="btn secondary" onClick={() => fileInput.current?.click()} disabled={isImporting}><Upload size={15}/> Đổi đề gốc</button></div>
            <div className="source-docx-stats"><div><b>{parsedDocx.question_count}</b><small>Câu nhận diện</small></div><div><b>{parsedDocx.section_summary?.length ?? 0}</b><small>Phần nhận diện</small></div><div><b>{parsedDocx.image_count}</b><small>Hình/đồ họa</small></div><div><b>{parsedDocx.table_count}</b><small>Bảng</small></div></div>
            <div className="source-section-list">{(parsedDocx.section_summary ?? []).map(section => <div key={section.id}><b>{section.title || section.id}</b><span>{section.question_count} câu · {section.image_count} hình · {section.table_count} bảng</span></div>)}</div>
            {parsedDocx.warnings?.length > 0 && <p className="source-docx-warning"><AlertTriangle size={15}/> {parsedDocx.warnings[0]}</p>}
          </section>}
          {!sourceDocx && <section className="panel source-upload-panel"><div><CloudUpload size={25}/><h2>Bắt đầu bằng đề Word gốc</h2><p>Tải DOCX mẫu lên để nhận diện các phần, câu hỏi, hình và bảng; sau đó tạo thứ tự trộn và xuất lại từ chính file gốc.</p></div><button className="btn primary" onClick={() => fileInput.current?.click()} disabled={isImporting}><Upload size={16}/>{isImporting ? "Đang nhận diện..." : "Tải đề Word gốc"}</button></section>}
          <div className="mixlayout"><section className="panel mixsettings"><PanelHead title="Cấu hình trộn" desc={sourceDocx ? "Áp dụng cho đề Word gốc đã tải" : "Áp dụng cho ngân hàng hiện tại"}/>
            <details className="header-settings" open>
              <summary><b>Cài đặt tiêu đề xuất DOCX</b></summary>
              <label className="field">Tên kỳ kiểm tra<input value={headerConfig.exam_title} onChange={e => setHeaderConfig(prev => ({...prev, exam_title:e.target.value}))}/></label>
              <div className="fieldrow"><label className="field">Sở GD&ĐT<input value={headerConfig.department} onChange={e => setHeaderConfig(prev => ({...prev, department:e.target.value}))}/></label><label className="field">Trường<input value={headerConfig.school} onChange={e => setHeaderConfig(prev => ({...prev, school:e.target.value}))}/></label></div>
              <div className="fieldrow"><label className="field">Môn thi<input placeholder="Nhập đúng môn của đề gốc" value={headerConfig.subject} onChange={e => setHeaderConfig(prev => ({...prev, subject:e.target.value}))}/></label><label className="field">Thời gian<input placeholder="Ví dụ: 90 phút" value={headerConfig.duration} onChange={e => setHeaderConfig(prev => ({...prev, duration:e.target.value}))}/></label></div>
              <div className="fieldrow"><label className="field">Khối/lớp<input value={headerConfig.grade} onChange={e => setHeaderConfig(prev => ({...prev, grade:e.target.value}))}/></label><label className="field">Năm học<input value={headerConfig.school_year} onChange={e => setHeaderConfig(prev => ({...prev, school_year:e.target.value}))}/></label></div>
            </details>
            {sourceDocx && parsedDocx?.section_summary?.length ? <div className="section-mode-settings"><b>Quy tắc theo từng phần (ưu tiên marker g0–g3 trong đề gốc)</b><p>g0 giữ nguyên · g1 đảo câu · g2 đảo phương án · g3 đảo cả câu và phương án. Có thể đổi từng phần trước khi tạo mã đề.</p>{parsedDocx.section_summary.map(section => <label className="field" key={section.id}>{section.title || section.id}<select value={sectionModes[section.id] ?? section.group ?? "g0"} onChange={e => setSectionModes(prev => ({...prev,[section.id]:e.target.value as MixMode}))}><option value="g0">g0 — Giữ nguyên</option><option value="g1">g1 — Đảo câu</option><option value="g2">g2 — Đảo phương án</option><option value="g3">g3 — Đảo câu và phương án</option></select></label>)}</div> : <><label className="field">Quy tắc trộn</label><div className="modes">{[
              {value:"g0" as MixMode,title:"Giữ nguyên",desc:"Không đảo câu, không đảo đáp án"},
              {value:"g1" as MixMode,title:"Đảo câu",desc:"Đổi thứ tự câu hỏi"},
              {value:"g2" as MixMode,title:"Đảo đáp án",desc:"Đổi phương án trắc nghiệm"},
              {value:"g3" as MixMode,title:"Đảo cả hai",desc:"Đảo câu và phương án"}
            ].map(item => <button key={item.value} className={mode===item.value?"chosen":""} onClick={() => setMode(item.value)}><span className="radio">{mode===item.value && <i/>}</span><span><b>{item.title}</b><small>{item.desc}</small></span>{mode===item.value && <Check size={15}/>}</button>)}</div></>}
            <div className="fieldrow"><label className="field">Số câu<input type="number" value={sourceDocx ? questions.length : questionCount} min="1" max={Math.max(1,questions.length)} disabled={Boolean(sourceDocx)} onChange={e => setQuestionCount(Math.max(1,Math.min(Math.max(1,questions.length),Number(e.target.value)||1)))}/><small>{sourceDocx ? "Giữ toàn bộ câu để bảo toàn cấu trúc đề gốc" : `Tối đa ${questions.length} câu`}</small></label><label className="field">Số mã đề<input type="number" value={codeCount} min="1" max="20" onChange={e => setCodeCount(Math.max(1,Math.min(20,Number(e.target.value)||1)))}/><small>Tối đa 20 mã trong MVP</small></label></div>
            <div className="warningbox"><AlertTriangle size={16}/> {sourceDocx ? "Mỗi phần áp dụng g0–g3 theo marker trong đề gốc hoặc lựa chọn của giáo viên. Đảo phương án chỉ an toàn khi các phương án là các đoạn A./B./C./D. riêng biệt; hãy kiểm tra DOCX và đáp án sau khi xuất." : "Đối chiếu công thức, hình ảnh và đáp án với file gốc trước khi phát hành."}</div><button className="btn primary full" onClick={createMixedCodes}><Shuffle size={17}/> {sourceDocx ? "Tạo thứ tự trộn từ đề gốc" : "Tạo mã đề xem trước"}</button>
          </section>
          <section className="panel preview"><PanelHead title="Xem trước mã đề" desc={mixedCodes.length ? `${mixedCodes.length} mã đề đã tạo` : "Chưa có mã đề trong phiên này"} action={mixedCodes.length ? "Tải JSON" : undefined} onAction={() => downloadJson(`ma-de-${String(activeCode+1).padStart(3,"0")}.json`,{code:String(activeCode+1).padStart(3,"0"),questions:currentCode})}/>
            {mixedCodes.length > 0 ? <>{sourceDocx && <div className="source-export-bar"><div><b>Xuất đề Word + đáp án Excel</b><small>Mã đề có thể chỉnh sửa trước khi xuất; hai file được đóng gói trong một ZIP.</small><label className="field">Mã đề xuất<input value={examCode} onChange={e => setExamCode(e.target.value)} maxLength={32} placeholder={String(activeCode + 1).padStart(3, "0")} /></label></div><button className="btn primary" onClick={() => void exportSourceDocx()} disabled={isExportingSource}><Download size={15}/>{isExportingSource ? "Đang đóng gói..." : "Tải ZIP đề + đáp án"}</button></div>}<div className="tabs">{mixedCodes.map((_,i) => <button className={i===activeCode?"on":""} key={`code-${i}`} onClick={() => setActiveCode(i)}>Mã {String(i+1).padStart(3,"0")}</button>)}</div>{sourceDocx ? <div className="source-preview-note"><FileText size={24}/><b>Bản xem trước văn bản không đại diện đầy đủ cho DOCX</b><p>Công thức, hình, biểu đồ và bảng được giữ trong file Word gốc. Dùng nút “Tải ZIP đề + đáp án” để nhận file Word và bảng đáp án Excel; mở DOCX để kiểm tra bố cục thật.</p><div><span>{currentCode.length} câu</span><span>{parsedDocx?.image_count ?? 0} hình/đồ họa nguồn</span><span>{parsedDocx?.table_count ?? 0} bảng nguồn</span></div></div> : <div className="paper"><div className="paperhead"><div><b>TRƯỜNG THPT …………</b><small>ĐỀ THI CHÍNH THỨC</small></div><div><b>KIỂM TRA ĐỊNH KỲ</b><small>Môn: Hóa học · 45 phút</small></div></div><p className="studentline">Họ và tên: .................................... Lớp: ........</p><p className="code">Mã đề: {String(activeCode+1).padStart(3,"0")}</p>{currentCode.map((q,i) => <div className="examq" key={q.id}><p><b>Câu {i+1}.</b> {q.text}</p>{q.type==="multiple_choice" && <div className="examopts">{q.options.map((option,j) => <span key={option.id}><b>{String.fromCharCode(65+j)}.</b> {option.text}</span>)}</div>}{q.type==="true_false" && <div className="examopts tf">{q.options.map((option,j) => <span key={option.id}><b>{String.fromCharCode(97+j)})</b> {option.text}</span>)}</div>}{q.type==="short_answer" && <p className="answerline">................................................................................</p>}{q.type==="essay" && <p className="answerline">................................................................................<br/>................................................................................</p>}</div>)}<div className="end">— HẾT —</div></div>}<div className="answerkey"><b><ShieldCheck size={16}/> Đáp án tham khảo — giáo viên</b><div>{currentCode.map((q,i) => <span key={q.id}>C{i+1}: <b>{getAnswerLabel(q)}</b></span>)}</div><small>Đáp án mẫu chưa được xác minh tự động.</small></div></> : <div className="empty previewempty"><Eye size={26}/><b>Bản xem trước sẽ xuất hiện tại đây</b><small>Chọn quy tắc trộn rồi tạo mã đề.</small><div className="paperfake"><i/><i/><i/><i/></div></div>}
          </section></div>
        </>}

        {page === "validation" && <>
          <div className="heading"><div><label>ĐẢM BẢO CHẤT LƯỢNG</label><h1>Kiểm định đề thi</h1><p>Rà soát cấu trúc, nội dung và đáp án trước khi phát hành.</p></div><button className="btn primary" onClick={() => setToast("Đã chạy kiểm tra cơ bản. Định dạng DOCX/PDF nâng cao chưa được kiểm tra.")}><RefreshCw size={16}/> Chạy kiểm tra</button></div>
          <Notice warning title="Chưa phải chứng nhận đề thi chính thức" text="Kiểm tra ngữ nghĩa đáp án, công thức, hình ảnh và bố cục cần các bộ xử lý riêng. Luôn đối chiếu file gốc."/>
          <section className="panel"><PanelHead title="Danh sách kiểm tra" desc="Trạng thái dữ liệu hiện tại"/>{[
            {title:"Mã định danh câu hỏi",desc:"Mỗi câu có ID để truy vết.",ok:questions.every(q=>Boolean(q.id))},
            {title:"Nội dung câu hỏi",desc:"Phát hiện câu hỏi trống.",ok:questions.every(q=>Boolean(q.text.trim()))},
            {title:"Phương án trắc nghiệm",desc:"Cần ít nhất hai phương án.",ok:questions.filter(q=>q.type==="multiple_choice").every(q=>q.options.length>=2)},
            {title:"Đáp án đã xác nhận",desc:"Giáo viên cần rà soát đáp án.",ok:false},
            {title:"Công thức, hình ảnh và bảng",desc:"Cần đối chiếu trực quan file gốc.",ok:false},
            {title:"Tệp DOCX/PDF cuối",desc:"Chưa có pipeline render so sánh.",ok:false}
          ].map(item => <div className="checkrow" key={item.title}><span className={item.ok?"ok":"pending"}>{item.ok?<CheckCircle2/>:<AlertTriangle/>}</span><div><b>{item.title}</b><small>{item.desc}</small></div><strong className={item.ok?"green":"orange"}>{item.ok?"Đạt cơ bản":"Cần rà soát"}</strong></div>)}</section>
        </>}

        {page === "exports" && <>
          <div className="heading"><div><label>HOÀN TẤT BỘ ĐỀ</label><h1>Xuất hồ sơ</h1><p>Tải dữ liệu câu hỏi và mã đề để sao lưu hoặc kiểm thử.</p></div></div>
          <div className="exportgrid">{[
            {id:"bank",icon:FileText,title:"Ngân hàng câu hỏi",desc:"Sao lưu câu hỏi thành JSON.",label:"Tải dữ liệu",action:exportBank},
            {id:"mixed",icon:FileCheck2,title:"Mã đề đã trộn",desc:"Tải cấu trúc mã đề đang chọn.",label:"Tải dữ liệu",action:() => mixedCodes.length ? downloadJson(`ma-de-${activeCode+1}.json`,currentCode) : setToast("Chưa có mã đề để xuất.")},
            {id:"excel",icon:Table2,title:"Đáp án Excel",desc:"Xuất đáp án theo mã đề đang chọn trong bộ ZIP.",label:"Đi tới trộn đề",action:() => setPage("mixer")},
            {id:"word",icon:FileDown,title:"Đề Word + đáp án Excel",desc:"Đóng gói đề DOCX và bảng đáp án XLSX thành một ZIP.",label:"Đi tới xuất đề",action:() => setPage("mixer")}
          ].map(item => { const Icon = item.icon; return <article className="exportcard" key={item.id}><Icon size={25}/><h3>{item.title}</h3><p>{item.desc}</p><button className="btn secondary full" onClick={item.action}><Download size={15}/> {item.label}</button></article>; })}</div>
          <Notice title="Không giả lập file thi chính thức" text="Bản xem trước hiện tại dùng để kiểm thử. Xuất Word/PDF/Excel cần được xác minh với các file mẫu thực tế."/>
        </>}

        {page === "online" && <>
          <div className="heading"><div><label>LỘ TRÌNH PHÁT TRIỂN</label><h1>Thi trực tuyến</h1><p>Khung tính năng để giao đề đã trộn cho học sinh làm trên điện thoại và máy tính.</p></div><span className="chip"><Smartphone size={15}/> Giai đoạn tiếp theo</span></div>
          <section className="onlinehero"><div><span><Sparkles size={14}/> PHIÊN BẢN TIẾP THEO</span><h2>Từ ngân hàng câu hỏi đến lớp học trực tuyến</h2><p>Dùng chung nền dữ liệu cho đề in và đề online, không nhập lại câu hỏi và đáp án.</p><button className="btn light" onClick={() => setToast("Đã ghi nhận yêu cầu: giao bài theo lớp, mã đề riêng, hẹn giờ, tự lưu và tự chấm.")}>Lưu yêu cầu phát triển <Check size={16}/></button></div><div className="devices"><div className="laptop"><i/><i/><i/><i/></div><div className="phone"><i/><i/><i/></div></div></section>
          <div className="features">{[
            {icon:Clock3,title:"Giới hạn thời gian",desc:"Giờ mở bài, thời lượng và hạn nộp."},
            {icon:RefreshCw,title:"Tự lưu câu trả lời",desc:"Khôi phục bài khi mất kết nối."},
            {icon:Shuffle,title:"Mã đề riêng",desc:"Lưu thứ tự câu và phương án."},
            {icon:Activity,title:"Báo cáo kết quả",desc:"Tổng hợp điểm và phân tích câu hỏi."}
          ].map(item => { const Icon = item.icon; return <article className="feature" key={item.title}><Icon size={22}/><b>{item.title}</b><p>{item.desc}</p><small>Dự kiến</small></article>; })}</div>
          <Notice warning title="Chưa phải hệ thống thi trực tuyến thực" text="Cần backend, tài khoản học sinh, lưu bài phía máy chủ, chấm điểm và bảo vệ đáp án trước khi dùng với học sinh thật."/>
        </>}

        {page === "settings" && <>
          <div className="heading"><div><label>CẤU HÌNH HỆ THỐNG</label><h1>Cài đặt</h1><p>Thông tin trường và dữ liệu trên thiết bị.</p></div></div>
          <section className="panel settings"><div><h3>Thông tin trường</h3><p>Dùng cho đầu trang đề thi khi bộ xuất tài liệu hoàn thiện.</p></div><label className="field">Tên trường<input defaultValue="Trường THPT ……………"/></label><label className="field">Sở / cơ quan quản lý<input defaultValue="SỞ GIÁO DỤC VÀ ĐÀO TẠO …………"/></label><label className="field">Tên giáo viên<input defaultValue="Giáo viên"/></label><button className="btn primary" onClick={() => setToast("Cài đặt hồ sơ hiện chỉ là mẫu giao diện.")}>Lưu hồ sơ mẫu</button><hr/><h3>Dữ liệu trên thiết bị</h3><p>Ngân hàng đang lưu trong localStorage của trình duyệt này, chưa đồng bộ máy chủ.</p><button className="btn secondary" onClick={exportBank}><Download size={16}/> Sao lưu JSON</button><button className="btn danger" onClick={() => { if (window.confirm("Xóa toàn bộ câu hỏi trên thiết bị? Hãy sao lưu trước.")) { setQuestions([]); setMixedCodes([]); setToast("Đã xóa dữ liệu cục bộ."); } }}><Trash2 size={16}/> Xóa dữ liệu cục bộ</button></section>
        </>}
      </div>
      <footer>Young Mix Pro <b>v0.1 MVP</b><span>Responsive workspace · Dành cho giáo viên</span></footer>
    </main>
    <input hidden type="file" ref={fileInput} accept=".json,.docx" onChange={event => { const file = event.target.files?.[0]; if (file) void importFile(file); }}/>
    {toast && <div className="toast"><CheckCircle2 size={18}/>{toast}<button onClick={() => setToast("")} aria-label="Đóng thông báo"><X size={15}/></button></div>}
    {modalOpen && <div className="modalback" onClick={event => { if (event.target === event.currentTarget) setModalOpen(false); }}><section className="modal">
      <header><div><h2>{editing ? "Chỉnh sửa câu hỏi" : "Thêm câu hỏi"}</h2><p>Đáp án không được tự suy luận; hãy xác nhận riêng.</p></div><button className="icon" onClick={() => setModalOpen(false)} aria-label="Đóng"><X/></button></header>
      <label className="field">Nội dung câu hỏi<textarea rows={4} value={draft.text} onChange={event => setDraft(prev => ({...prev,text:event.target.value}))} placeholder="Nhập nội dung câu hỏi..."/></label>
      <div className="fieldrow"><label className="field">Dạng câu<select value={draft.type} onChange={event => setDraft(prev => ({...prev,type:event.target.value as QuestionType}))}>{Object.entries(TYPE_LABELS).map(([key,label]) => <option value={key} key={key}>{label}</option>)}</select></label><label className="field">Mức độ<select value={draft.level} onChange={event => setDraft(prev => ({...prev,level:event.target.value as Level}))}><option>Biết</option><option>Hiểu</option><option>Vận dụng</option></select></label></div>
      <div className="fieldrow"><label className="field">Chủ đề<input value={draft.topic} onChange={event => setDraft(prev => ({...prev,topic:event.target.value}))}/></label><label className="field">Điểm<input type="number" min="0" step="0.25" value={draft.points} onChange={event => setDraft(prev => ({...prev,points:event.target.value}))}/></label></div>
      <footer><button className="btn secondary" onClick={() => setModalOpen(false)}>Hủy</button><button className="btn primary" onClick={saveQuestion}><Check size={16}/> Lưu câu hỏi</button></footer>
    </section></div>}
  </div>;
}
