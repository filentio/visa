import { useEffect, useMemo, useState } from "react";
import { Link, useParams } from "react-router-dom";
import type { ClientDetail, ClientPackageItem, JobStatusResponse, PackageResponse } from "../api/types";
import { ApiError, getClient, getClientPackages, getDownloadUrl, getJob, getPackage, presignFile, regeneratePackage } from "../api/client";

type PackageState = {
  expanded: boolean;
  loading: boolean;
  error: string | null;
  pkg: PackageResponse | null;
  regenJobId: string | null;
  regenJob: JobStatusResponse | null;
};

function defaultPackageState(): PackageState {
  return {
    expanded: false,
    loading: false,
    error: null,
    pkg: null,
    regenJobId: null,
    regenJob: null,
  };
}

function statusBadge(status: ClientPackageItem["status"]) {
  if (status === "generated") return "badge badge-green";
  if (status === "error") return "badge badge-red";
  return "badge badge-gray";
}

function groupByVersion(pkg: PackageResponse) {
  const m = new Map<number, PackageResponse["documents"]>();
  for (const d of pkg.documents) {
    const list = m.get(d.version) ?? [];
    list.push(d);
    m.set(d.version, list);
  }
  const versions = Array.from(m.keys()).sort((a, b) => b - a);
  return { versions, map: m };
}

export function ClientDetails() {
  const { clientId } = useParams();
  const [client, setClient] = useState<ClientDetail | null>(null);
  const [packages, setPackages] = useState<ClientPackageItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [states, setStates] = useState<Record<string, PackageState>>({});

  const refresh = async () => {
    if (!clientId) return;
    setLoading(true);
    setError(null);
    try {
      const [c, pkgs] = await Promise.all([getClient(clientId), getClientPackages(clientId)]);
      setClient(c);
      setPackages(pkgs);
    } catch (e) {
      setError(e instanceof ApiError ? e.message : "Failed to load client");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    refresh();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [clientId]);

  // Poll regenerate jobs (per package) every 2 seconds while any is running/queued.
  useEffect(() => {
    const active = Object.entries(states).filter(([, st]) => st.regenJobId && (!st.regenJob || st.regenJob.status === "queued" || st.regenJob.status === "running"));
    if (active.length === 0) return;
    let alive = true;

    const tick = async () => {
      const updates: Array<[string, JobStatusResponse]> = [];
      for (const [pkgId, st] of active) {
        try {
          const j = await getJob(st.regenJobId!);
          updates.push([pkgId, j]);
        } catch {
          // ignore transient errors; page shows last known state
        }
      }
      if (!alive) return;
      if (updates.length === 0) return;
      setStates((prev) => {
        const next = { ...prev };
        for (const [pkgId, j] of updates) {
          const cur = next[pkgId];
          if (!cur) continue;
          next[pkgId] = { ...cur, regenJob: j };
        }
        return next;
      });

      const finished = updates.filter(([, j]) => j.status === "done" || j.status === "error");
      if (finished.length > 0) {
        // refresh package list so version_counter updates
        refresh();
        // refresh expanded package details after done
        for (const [pkgId, j] of finished) {
          if (j.status === "done") {
            const st = states[pkgId];
            if (st?.expanded) {
              try {
                const p = await getPackage(pkgId);
                if (!alive) return;
                setStates((prev) => ({ ...prev, [pkgId]: { ...(prev[pkgId] ?? st), pkg: p } }));
              } catch {
                // ignore
              }
            }
          }
        }
      }
    };

    tick();
    const id = window.setInterval(tick, 2000);
    return () => {
      alive = false;
      window.clearInterval(id);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [states]);

  const ensureState = (packageId: string) =>
    setStates((prev) => ({
      ...prev,
      [packageId]: prev[packageId] ?? defaultPackageState(),
    }));

  const toggleExpand = async (packageId: string) => {
    ensureState(packageId);
    const wasExpanded = states[packageId]?.expanded ?? false;
    setStates((prev) => {
      const cur = prev[packageId] ?? defaultPackageState();
      return { ...prev, [packageId]: { ...cur, expanded: !cur.expanded } };
    });
    if (!wasExpanded) {
      // opening: load package details
      setStates((prev) => {
        const cur = prev[packageId] ?? defaultPackageState();
        return { ...prev, [packageId]: { ...cur, loading: true, error: null } };
      });
      try {
        const p = await getPackage(packageId);
        setStates((prev) => {
          const cur = prev[packageId] ?? defaultPackageState();
          return { ...prev, [packageId]: { ...cur, pkg: p, loading: false } };
        });
      } catch (e) {
        setStates((prev) => {
          const cur = prev[packageId] ?? defaultPackageState();
          return {
            ...prev,
            [packageId]: { ...cur, loading: false, error: e instanceof ApiError ? e.message : "Failed to load package" },
          };
        });
      }
    }
  };

  const onDownloadZip = async (packageId: string) => {
    try {
      const url = await getDownloadUrl(packageId);
      window.open(url, "_blank", "noopener,noreferrer");
    } catch (e) {
      setError(e instanceof ApiError ? e.message : "Download failed");
    }
  };

  const onRegenerate = async (p: ClientPackageItem) => {
    ensureState(p.package_id);
    setStates((prev) => {
      const cur = prev[p.package_id] ?? defaultPackageState();
      return { ...prev, [p.package_id]: { ...cur, error: null } };
    });
    try {
      const res = await regeneratePackage(p.package_id);
      setStates((prev) => ({
        ...prev,
        [p.package_id]: { ...(prev[p.package_id] ?? defaultPackageState()), regenJobId: res.job_id, regenJob: null },
      }));
    } catch (e) {
      setStates((prev) => ({
        ...prev,
        [p.package_id]: { ...(prev[p.package_id] ?? defaultPackageState()), error: e instanceof ApiError ? e.message : "Regenerate failed" },
      }));
    }
  };

  const list = useMemo(() => packages, [packages]);

  return (
    <div className="container">
      <div className="header">
        <div>
          <h1>Client</h1>
          <div className="hint">
            <Link className="link" to="/clients">
              ← Back to clients
            </Link>
          </div>
        </div>
        <button className="btn btn-secondary" onClick={refresh} disabled={loading}>
          Refresh
        </button>
      </div>

      {error && <div className="error">{error}</div>}

      {loading ? (
        <div className="hint">Loading…</div>
      ) : !client ? (
        <div className="hint">Client not found.</div>
      ) : (
        <>
          <div className="panel">
            <div className="section-title">Profile</div>
            <div className="grid2">
              <div>
                <div className="hint">Full name</div>
                <div>{client.full_name}</div>
                <div className="hint mt">Passport</div>
                <div className="mono">{client.passport_masked}</div>
              </div>
              <div>
                <div className="hint">DOB</div>
                <div className="mono">{client.dob}</div>
                <div className="hint mt">Issuing</div>
                <div className="mono">
                  {client.issuing_country ?? "—"} / {client.country_display}
                </div>
                <div className="hint mt">Created</div>
                <div className="mono">{client.created_at}</div>
              </div>
            </div>
          </div>

          <div className="panel mt">
            <div className="section-title">Packages</div>
            {list.length === 0 ? (
              <div className="hint">No packages for this client yet.</div>
            ) : (
              <div className="pkglist">
                {list.map((p) => {
                  const st = states[p.package_id];
                  const regenRunning = !!st?.regenJobId && (!st.regenJob || st.regenJob.status === "queued" || st.regenJob.status === "running");
                  const nextVer = p.version_counter + 1;
                  return (
                    <div key={p.package_id} className="pkgitem">
                      <div className="row row-between">
                        <div>
                          <div className="hint">{p.company.name}</div>
                          <div className="mono">{p.package_id}</div>
                        </div>
                        <div className={statusBadge(p.status)}>{p.status}</div>
                      </div>

                      <div className="grid2 mt">
                        <div>
                          <div className="hint">Version</div>
                          <div className="badge badge-gray">v{p.version_counter}</div>
                        </div>
                        <div>
                          <div className="hint">Created / Updated</div>
                          <div className="mono">{p.created_at}</div>
                          <div className="mono">{p.updated_at}</div>
                        </div>
                      </div>

                      {st?.error && <div className="error mt">{st.error}</div>}
                      {st?.regenJob && st.regenJob.status === "error" && (
                        <div className="error mt">{st.regenJob.error_message || "Regenerate failed"}</div>
                      )}
                      {st?.regenJobId && (
                        <div className="hint mt">
                          Regen job: <code>{st.regenJobId}</code> {st.regenJob ? `(${st.regenJob.status})` : ""}
                        </div>
                      )}

                      <div className="row mt">
                        <button className="btn btn-secondary" onClick={() => toggleExpand(p.package_id)}>
                          {st?.expanded ? "Close package" : "Open package"}
                        </button>
                        <button className="btn" onClick={() => onDownloadZip(p.package_id)}>
                          Download latest ZIP
                        </button>
                        <button className="btn" onClick={() => onRegenerate(p)} disabled={regenRunning}>
                          Regenerate v{nextVer}
                        </button>
                      </div>

                      {st?.expanded && (
                        <div className="mt">
                          {st.loading ? (
                            <div className="hint">Loading package…</div>
                          ) : st.error ? (
                            <div className="error">{st.error}</div>
                          ) : st.pkg ? (
                            <PackageHistory pkg={st.pkg} />
                          ) : (
                            <div className="hint">—</div>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

function PackageHistory({ pkg }: { pkg: PackageResponse }) {
  const { versions, map } = groupByVersion(pkg);
  const latest = versions[0];

  const [resolving, setResolving] = useState<Record<string, boolean>>({});

  const openUrl = async (doc: PackageResponse["documents"][number]) => {
    if (doc.presigned_url) {
      window.open(doc.presigned_url, "_blank", "noopener,noreferrer");
      return;
    }
    setResolving((p) => ({ ...p, [doc.file_key]: true }));
    try {
      const url = await presignFile(doc.file_key);
      window.open(url, "_blank", "noopener,noreferrer");
    } catch {
      // ignore
    } finally {
      setResolving((p) => ({ ...p, [doc.file_key]: false }));
    }
  };

  const labelMap: Record<PackageResponse["documents"][number]["doc_type"], string> = {
    contract: "Contract",
    bank_statement: "Bank statement",
    insurance: "Insurance",
    salary: "Salary",
    bundle: "Bundle ZIP",
    other: "Other",
  };

  return (
    <div className="panel">
      <div className="section-title">Documents history</div>
      {versions.length === 0 ? (
        <div className="hint">No documents yet.</div>
      ) : (
        <div className="verlist">
          {versions.map((v) => {
            const docs = (map.get(v) ?? []).slice().sort((a, b) => a.doc_type.localeCompare(b.doc_type));
            return (
              <div key={v} className="ver">
                <div className="row row-between">
                  <div className="row">
                    <div className="badge badge-gray">v{v}</div>
                    {v === latest && <div className="badge badge-blue">latest</div>}
                  </div>
                </div>
                <div className="docs mt">
                  {docs.map((d) => (
                    <div key={d.file_key} className="docrow">
                      <div className="hint">{labelMap[d.doc_type] ?? d.doc_type}</div>
                      <div className="row">
                        <button className="btn btn-secondary" onClick={() => openUrl(d)} disabled={!!resolving[d.file_key]}>
                          {resolving[d.file_key] ? "Getting link…" : "Download"}
                        </button>
                        <span className="mono">{d.created_at}</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

