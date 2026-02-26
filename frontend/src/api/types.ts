export type GeneratePackageRequest = {
  client: {
    full_name: string;
    passport_no: string;
    dob: string; // YYYY-MM-DD
    mrz_line1?: string | null;
    mrz_line2?: string | null;
    issuing_country?: string | null; // ISO3 fallback if MRZ not provided
  };
  company_id: string;
  job: {
    position: string;
    salary_rub: number;
    currency: "USD" | "AED";
    fx_source: "manual" | "cbr";
    fx_rate?: number | null; // required if manual
  };
  templates: {
    contract_template: "договор" | "договор2";
    insurance_template: "страховка" | "РГС";
    bank_template: "т-банк 2 (6 мес) $";
    salary_template: "Salary упрошенная";
  };
};

export type GeneratePackageResponse = {
  job_id: string;
  package_id: string;
};

export type JobStatusResponse = {
  job_id: string;
  status: "queued" | "running" | "done" | "error";
  error_message?: string | null;
  started_at?: string | null;
  finished_at?: string | null;
  version?: number | null;
};

export type PackageResponse = {
  package_id: string;
  status: "draft" | "generated" | "error";
  version_counter: number;
  client: {
    client_id: string;
    full_name: string;
    passport_masked: string;
    dob: string;
    issuing_country?: string | null;
  };
  company: {
    company_id: string;
    name: string;
  };
  documents: Array<{
    doc_type: "contract" | "bank_statement" | "insurance" | "salary" | "bundle" | "other";
    version: number;
    file_key: string;
    created_at: string;
    presigned_url?: string | null;
  }>;
};

export type DownloadResponse = {
  url: string;
};

export type Company = {
  id: string;
  name: string;
};

export type ClientSearchItem = {
  client_id: string;
  full_name: string;
  passport_masked: string;
  dob: string;
  issuing_country?: string | null;
  created_at: string;
};

export type ClientDetail = {
  client_id: string;
  full_name: string;
  passport_masked: string;
  dob: string;
  issuing_country?: string | null;
  country_display: string;
  created_at: string;
};

export type ClientPackageItem = {
  package_id: string;
  status: "generated" | "error" | "draft";
  version_counter: number;
  company: { company_id: string; name: string };
  created_at: string;
  updated_at: string;
};

export type RegenerateResponse = {
  job_id: string;
  package_id: string;
};

