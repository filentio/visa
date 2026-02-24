/* eslint-disable react-hooks/set-state-in-effect */
import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import type { ClientSearchItem } from "../api/types";
import { ApiError, searchClients } from "../api/client";

export function ClientsList() {
  const navigate = useNavigate();
  const [query, setQuery] = useState("");
  const [items, setItems] = useState<ClientSearchItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let alive = true;
    setLoading(true);
    setError(null);

    const id = window.setTimeout(() => {
      searchClients(query)
        .then((res) => {
          if (!alive) return;
          setItems(res);
        })
        .catch((e) => {
          if (!alive) return;
          setError(e instanceof ApiError ? e.message : "Failed to load clients");
        })
        .finally(() => {
          if (!alive) return;
          setLoading(false);
        });
    }, 300);

    return () => {
      alive = false;
      window.clearTimeout(id);
    };
  }, [query]);

  const rows = useMemo(() => items, [items]);

  return (
    <div className="container">
      <div className="header">
        <div>
          <h1>Clients</h1>
          <div className="hint">Search by full name or passport (server returns masked passport only)</div>
        </div>
      </div>

      <div className="panel">
        <div className="field">
          <label className="label">Search</label>
          <input className="input" value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Ivanov / 87555" />
        </div>

        {error && <div className="error">Error: {error}</div>}
        {loading ? (
          <div className="hint">Loading…</div>
        ) : rows.length === 0 ? (
          <div className="hint">No clients found.</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table className="table">
              <thead>
                <tr>
                  <th>Full name</th>
                  <th>Passport</th>
                  <th>DOB</th>
                  <th>Issuing</th>
                  <th>Created</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((c) => (
                  <tr
                    key={c.client_id}
                    className="tr-link"
                    onClick={() => navigate(`/clients/${c.client_id}`)}
                    role="button"
                    tabIndex={0}
                  >
                    <td>{c.full_name}</td>
                    <td className="mono">{c.passport_masked}</td>
                    <td className="mono">{c.dob}</td>
                    <td className="mono">{c.issuing_country ?? "—"}</td>
                    <td className="mono">{c.created_at}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

