const configuredBase = (import.meta.env.VITE_API_BASE_URL ?? "").trim();
export const API_BASE_URL = configuredBase.replace(/\/+$/, "");

export class ApiError extends Error {
  status?: number;
  constructor(message: string, status?: number) {
    super(message);
    this.name = "ApiError";
    this.status = status;
  }
}

async function readError(response: Response): Promise<string> {
  try {
    const body: unknown = await response.json();
    if (body && typeof body === "object" && "detail" in body) {
      const detail = (body as { detail?: unknown }).detail;
      if (typeof detail === "string") return detail;
      if (Array.isArray(detail)) return detail.map((item) => {
        if (item && typeof item === "object" && "msg" in item) return String((item as { msg: unknown }).msg);
        return String(item);
      }).join("; ");
    }
  } catch {
    // A proxy may return a non-JSON error page.
  }
  return `Máy chủ trả về lỗi HTTP ${response.status}.`;
}

export async function apiFetch(path: string, init: RequestInit = {}): Promise<Response> {
  let response: Response;
  try {
    response = await fetch(`${API_BASE_URL}${path}`, init);
  } catch {
    const target = API_BASE_URL || "cùng domain với frontend (VITE_API_BASE_URL đang trống)";
    throw new ApiError(`Không kết nối được backend tại ${target}. Kiểm tra backend có đang chạy và VITE_API_BASE_URL đã cấu hình đúng chưa.`);
  }
  if (!response.ok) throw new ApiError(await readError(response), response.status);
  return response;
}

export async function checkApiHealth(): Promise<{ ok: boolean; service: string; version: string }> {
  return (await apiFetch("/api/health")).json();
}
