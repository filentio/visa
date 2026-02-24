import type {
  Company,
  DownloadResponse,
  GeneratePackageRequest,
  GeneratePackageResponse,
  JobStatusResponse,
  PackageResponse,
} from "./types";

export class ApiError extends Error {
  status?: number;
  constructor(message: string, status?: number) {
    super(message);
    this.name = "ApiError";
    this.status = status;
  }
}

const BASE_URL = (import.meta.env.VITE_API_BASE_URL as string | undefined)?.replace(/\/+$/, "");

function requireBaseUrl(): string {
  if (!BASE_URL) {
    throw new ApiError("VITE_API_BASE_URL is not set. Create `.env` or use `.env.example`.");
  }
  return BASE_URL;
}

async function requestJson<T>(path: string, init?: RequestInit & { timeoutMs?: number }): Promise<T> {
  const baseUrl = requireBaseUrl();
  const url = `${baseUrl}${path.startsWith("/") ? "" : "/"}${path}`;

  const controller = new AbortController();
  const timeoutMs = init?.timeoutMs ?? 12_000;
  const t = window.setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(url, {
      ...init,
      headers: {
        "Content-Type": "application/json",
        ...(init?.headers ?? {}),
      },
      signal: controller.signal,
    });

    const contentType = res.headers.get("content-type") || "";
    const isJson = contentType.includes("application/json");
    const body = isJson ? await res.json().catch(() => null) : await res.text().catch(() => "");

    if (!res.ok) {
      const msg =
        (typeof body === "object" && body && ("detail" in body || "message" in body)
          ? String((body as any).detail ?? (body as any).message)
          : typeof body === "string" && body
            ? body
            : `Request failed: ${res.status}`) || `Request failed: ${res.status}`;
      throw new ApiError(msg, res.status);
    }

    return body as T;
  } catch (e: any) {
    if (e?.name === "AbortError") {
      throw new ApiError("API request timeout");
    }
    if (e instanceof ApiError) throw e;
    throw new ApiError(e?.message ? String(e.message) : "Network error");
  } finally {
    window.clearTimeout(t);
  }
}

export async function listCompanies(): Promise<Company[]> {
  return requestJson<Company[]>("/companies", { method: "GET" });
}

export async function generatePackage(req: GeneratePackageRequest): Promise<GeneratePackageResponse> {
  return requestJson<GeneratePackageResponse>("/packages/generate", {
    method: "POST",
    body: JSON.stringify(req),
    timeoutMs: 30_000,
  });
}

export async function getJob(jobId: string): Promise<JobStatusResponse> {
  return requestJson<JobStatusResponse>(`/jobs/${encodeURIComponent(jobId)}`, { method: "GET" });
}

export async function getPackage(packageId: string): Promise<PackageResponse> {
  return requestJson<PackageResponse>(`/packages/${encodeURIComponent(packageId)}`, { method: "GET" });
}

export async function getDownloadUrl(packageId: string): Promise<string> {
  const res: any = await requestJson<any>(`/packages/${encodeURIComponent(packageId)}/download`, { method: "GET" });
  if (typeof res === "string") return res;
  if (res && typeof res === "object" && typeof res.url === "string") return res.url;
  throw new ApiError("Unexpected download response");
}

export async function presignFile(key: string): Promise<string> {
  const res = await requestJson<DownloadResponse>(`/files/presign?key=${encodeURIComponent(key)}`, { method: "GET" });
  return res.url;
}

