import { useEffect, useMemo, useState } from "react";
import type { GeneratePackageRequest, JobStatusResponse, PackageResponse } from "../api/types";
import { ApiError, generatePackage, getDownloadUrl, getJob, getPackage } from "../api/client";
import { CompanySelect } from "../components/CompanySelect";
import { RemoteRoleSelect } from "../components/RemoteRoleSelect";
import { JobStatus } from "../components/JobStatus";
import { DocumentLinks } from "../components/DocumentLinks";

function emptyRequest(): GeneratePackageRequest {
  return {
    client: {
      full_name: "",
      passport_no: "",
      dob: "",
      mrz_line1: null,
      mrz_line2: null,
      issuing_country: null,
    },
    company_id: "",
    job: {
      position: "Senior business analyst",
      salary_rub: 840000,
      currency: "USD",
      fx_source: "manual",
      fx_rate: 79.75,
    },
    templates: {
      contract_template: "договор",
      insurance_template: "страховка",
      bank_template: "т-банк 2 (6 мес) $",
      salary_template: "Salary упрошенная",
    },
  };
}

export function CreatePackage() {
  const [req, setReq] = useState<GeneratePackageRequest>(() => emptyRequest());
  const [formError, setFormError] = useState<string | null>(null);
  const [apiError, setApiError] = useState<string | null>(null);

  const [jobId, setJobId] = useState<string | null>(null);
  const [packageId, setPackageId] = useState<string | null>(null);

  const [job, setJob] = useState<JobStatusResponse | null>(null);
  const [pkg, setPkg] = useState<PackageResponse | null>(null);

  const isPolling = !!jobId && (!job || job.status === "queued" || job.status === "running");

  const canSubmit = useMemo(() => {
    if (!req.client.full_name.trim()) return false;
    if (!req.client.passport_no.trim()) return false;
    if (!req.client.dob.trim()) return false;
    if (!req.company_id.trim()) return false;
    if (!req.job.position.trim()) return false;
    if (!Number.isFinite(req.job.salary_rub) || req.job.salary_rub <= 0) return false;
    if (req.job.fx_source === "manual") {
      if (req.job.fx_rate == null) return false;
      if (!Number.isFinite(req.job.fx_rate) || req.job.fx_rate <= 0) return false;
    }
    return true;
  }, [req]);

  useEffect(() => {
    if (!jobId) return;
    let alive = true;
    let intervalId: number | null = null;

    const tick = async () => {
      try {
        const j = await getJob(jobId);
        if (!alive) return;
        setJob(j);
        if (j.status === "done" && packageId) {
          const p = await getPackage(packageId);
          if (!alive) return;
          setPkg(p);
        }
        if (j.status === "done" || j.status === "error") {
          if (intervalId != null) {
            window.clearInterval(intervalId);
            intervalId = null;
          }
        }
      } catch (e) {
        if (!alive) return;
        const msg = e instanceof ApiError ? e.message : "Failed to fetch job status";
        setApiError(msg);
      }
    };

    tick();
    intervalId = window.setInterval(() => {
      tick();
    }, 2000);

    return () => {
      alive = false;
      if (intervalId != null) window.clearInterval(intervalId);
    };
  }, [jobId, packageId]);

  const onReset = () => {
    setReq(emptyRequest());
    setFormError(null);
    setApiError(null);
    setJobId(null);
    setPackageId(null);
    setJob(null);
    setPkg(null);
  };

  const onSubmit = async () => {
    setFormError(null);
    setApiError(null);
    setPkg(null);

    if (!canSubmit) {
      setFormError("Please fill all required fields.");
      return;
    }
    try {
      const res = await generatePackage(req);
      setJobId(res.job_id);
      setPackageId(res.package_id);
      setJob(null);
    } catch (e) {
      const msg = e instanceof ApiError ? e.message : "Generate failed";
      setApiError(msg);
    }
  };

  const onDownloadZip = async () => {
    if (!packageId) return;
    setApiError(null);
    try {
      const url = await getDownloadUrl(packageId);
      window.open(url, "_blank", "noopener,noreferrer");
    } catch (e) {
      const msg = e instanceof ApiError ? e.message : "Download failed";
      setApiError(msg);
    }
  };

  return (
    <div className="container">
      <div className="header">
        <div>
          <h1>Package generator</h1>
          <div className="hint">Create package → track job → download ZIP/PDF</div>
        </div>
        <button className="btn btn-secondary" onClick={onReset} disabled={isPolling}>
          New package
        </button>
      </div>

      {formError && <div className="error">{formError}</div>}
      {apiError && <div className="error">{apiError}</div>}

      <div className="grid2">
        <div className="panel">
          <div className="section-title">Client</div>

          <div className="field">
            <label className="label">Full name *</label>
            <input
              className="input"
              value={req.client.full_name}
              onChange={(e) => setReq({ ...req, client: { ...req.client, full_name: e.target.value } })}
              disabled={isPolling}
            />
          </div>

          <div className="field">
            <label className="label">Passport no *</label>
            <input
              className="input"
              value={req.client.passport_no}
              onChange={(e) => setReq({ ...req, client: { ...req.client, passport_no: e.target.value } })}
              disabled={isPolling}
            />
          </div>

          <div className="field">
            <label className="label">DOB *</label>
            <input
              className="input"
              type="date"
              value={req.client.dob}
              onChange={(e) => setReq({ ...req, client: { ...req.client, dob: e.target.value } })}
              disabled={isPolling}
            />
          </div>

          <div className="field">
            <label className="label">MRZ line 1 (optional)</label>
            <textarea
              className="input"
              rows={2}
              value={req.client.mrz_line1 ?? ""}
              onChange={(e) => setReq({ ...req, client: { ...req.client, mrz_line1: e.target.value || null } })}
              disabled={isPolling}
            />
          </div>

          <div className="field">
            <label className="label">MRZ line 2 (optional)</label>
            <textarea
              className="input"
              rows={2}
              value={req.client.mrz_line2 ?? ""}
              onChange={(e) => setReq({ ...req, client: { ...req.client, mrz_line2: e.target.value || null } })}
              disabled={isPolling}
            />
          </div>

          <div className="field">
            <label className="label">Issuing country ISO3 (optional fallback)</label>
            <input
              className="input"
              placeholder="RUS"
              value={req.client.issuing_country ?? ""}
              onChange={(e) =>
                setReq({ ...req, client: { ...req.client, issuing_country: e.target.value.trim() || null } })
              }
              disabled={isPolling}
            />
          </div>
        </div>

        <div className="panel">
          <div className="section-title">Company & job</div>

          <CompanySelect
            value={req.company_id}
            onChange={(companyId) => setReq({ ...req, company_id: companyId })}
            disabled={isPolling}
          />

          <RemoteRoleSelect
            value={req.job.position}
            onChange={(v) => setReq({ ...req, job: { ...req.job, position: v } })}
            disabled={isPolling}
          />

          <div className="field">
            <label className="label">Salary (RUB) *</label>
            <input
              className="input"
              type="number"
              value={req.job.salary_rub}
              onChange={(e) => setReq({ ...req, job: { ...req.job, salary_rub: Number(e.target.value) } })}
              disabled={isPolling}
            />
          </div>

          <div className="grid2">
            <div className="field">
              <label className="label">Currency</label>
              <select
                className="input"
                value={req.job.currency}
                onChange={(e) => {
                  const v = e.target.value;
                  if (v === "USD" || v === "AED") setReq({ ...req, job: { ...req.job, currency: v } });
                }}
                disabled={isPolling}
              >
                <option value="USD">USD</option>
                <option value="AED">AED</option>
              </select>
            </div>
            <div className="field">
              <label className="label">FX source</label>
              <select
                className="input"
                value={req.job.fx_source}
                onChange={(e) => {
                  const v = e.target.value;
                  if (v === "manual" || v === "cbr") setReq({ ...req, job: { ...req.job, fx_source: v } });
                }}
                disabled={isPolling}
              >
                <option value="manual">manual</option>
                <option value="cbr">cbr</option>
              </select>
            </div>
          </div>

          {req.job.fx_source === "manual" && (
            <div className="field">
              <label className="label">FX rate *</label>
              <input
                className="input"
                type="number"
                step="0.0001"
                value={req.job.fx_rate ?? ""}
                onChange={(e) =>
                  setReq({ ...req, job: { ...req.job, fx_rate: e.target.value ? Number(e.target.value) : null } })
                }
                disabled={isPolling}
              />
            </div>
          )}
        </div>
      </div>

      <div className="panel">
        <div className="section-title">Templates</div>

        <div className="grid2">
          <div className="field">
            <label className="label">Contract</label>
            <select
              className="input"
              value={req.templates.contract_template}
              onChange={(e) =>
                setReq({
                  ...req,
                  templates: {
                    ...req.templates,
                    contract_template: e.target.value === "договор2" ? "договор2" : "договор",
                  },
                })
              }
              disabled={isPolling}
            >
              <option value="договор">договор</option>
              <option value="договор2">договор2</option>
            </select>
          </div>

          <div className="field">
            <label className="label">Insurance</label>
            <select
              className="input"
              value={req.templates.insurance_template}
              onChange={(e) =>
                setReq({
                  ...req,
                  templates: {
                    ...req.templates,
                    insurance_template: e.target.value === "РГС" ? "РГС" : "страховка",
                  },
                })
              }
              disabled={isPolling}
            >
              <option value="страховка">страховка</option>
              <option value="РГС">РГС</option>
            </select>
          </div>
        </div>

        <div className="grid2">
          <div className="field">
            <label className="label">Bank template (fixed)</label>
            <input className="input" value={req.templates.bank_template} disabled />
          </div>
          <div className="field">
            <label className="label">Salary template (fixed)</label>
            <input className="input" value={req.templates.salary_template} disabled />
          </div>
        </div>

        <div className="row mt">
          <button className="btn" onClick={onSubmit} disabled={!canSubmit || isPolling}>
            {isPolling ? "Generating…" : "Generate"}
          </button>
          {packageId && job && job.status === "done" && (
            <button className="btn btn-secondary" onClick={onDownloadZip}>
              Download ZIP
            </button>
          )}
          {jobId && (
            <div className="hint">
              Job: <code>{jobId}</code>
            </div>
          )}
        </div>
      </div>

      <JobStatus job={job} jobId={jobId} />

      {job?.status === "done" && packageId && (
        <div className="panel">
          <div className="section-title">Result</div>
          <div className="row">
            <button className="btn" onClick={onDownloadZip}>
              Download ZIP
            </button>
            <button
              className="btn btn-secondary"
              onClick={async () => {
                if (!packageId) return;
                setApiError(null);
                try {
                  const p = await getPackage(packageId);
                  setPkg(p);
                } catch (e) {
                  setApiError(e instanceof ApiError ? e.message : "Failed to load package");
                }
              }}
            >
              Refresh documents
            </button>
          </div>
          {pkg ? (
            <div className="mt">
              <DocumentLinks pkg={pkg} />
            </div>
          ) : (
            <div className="hint mt">Loading package…</div>
          )}
        </div>
      )}
    </div>
  );
}

