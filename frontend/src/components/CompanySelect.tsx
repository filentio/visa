/* eslint-disable react-hooks/set-state-in-effect */
import { useEffect, useMemo, useState } from "react";
import type { Company } from "../api/types";
import { ApiError, listCompanies } from "../api/client";

type Props = {
  value: string;
  onChange: (companyId: string) => void;
  disabled?: boolean;
};

export function CompanySelect({ value, onChange, disabled }: Props) {
  const [companies, setCompanies] = useState<Company[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let alive = true;
    setLoading(true);
    setError(null);
    listCompanies()
      .then((items) => {
        if (!alive) return;
        setCompanies(items);
      })
      .catch((e) => {
        if (!alive) return;
        const msg = e instanceof ApiError ? e.message : "Failed to load companies";
        setError(msg);
      })
      .finally(() => {
        if (!alive) return;
        setLoading(false);
      });
    return () => {
      alive = false;
    };
  }, []);

  const options = useMemo(() => companies, [companies]);

  return (
    <div className="field">
      <label className="label">Company</label>
      {loading ? (
        <div className="hint">Loading companies…</div>
      ) : error ? (
        <div className="error">Companies error: {error}</div>
      ) : (
        <select className="input" value={value} onChange={(e) => onChange(e.target.value)} disabled={disabled}>
          <option value="">Select company…</option>
          {options.map((c) => (
            <option key={c.id} value={c.id}>
              {c.name}
            </option>
          ))}
        </select>
      )}
    </div>
  );
}

