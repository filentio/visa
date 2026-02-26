import type { PackageResponse } from "../api/types";

type Props = {
  pkg: PackageResponse;
};

function latestByType(pkg: PackageResponse) {
  const map = new Map<string, PackageResponse["documents"][number]>();
  for (const d of pkg.documents) {
    const key = d.doc_type;
    const cur = map.get(key);
    if (!cur || d.version > cur.version) map.set(key, d);
  }
  return map;
}

function DocRow({
  label,
  doc,
}: {
  label: string;
  doc?: PackageResponse["documents"][number];
}) {
  if (!doc) {
    return (
      <div className="docrow">
        <div className="hint">{label}</div>
        <div>—</div>
      </div>
    );
  }
  return (
    <div className="docrow">
      <div className="hint">{label}</div>
      <div className="row">
        <div className="badge badge-gray">v{doc.version}</div>
        {doc.presigned_url ? (
          <a className="link" href={doc.presigned_url} target="_blank" rel="noreferrer">
            Download
          </a>
        ) : (
          <span className="mono">{doc.file_key}</span>
        )}
      </div>
    </div>
  );
}

export function DocumentLinks({ pkg }: Props) {
  const latest = latestByType(pkg);
  const contract = latest.get("contract");
  const bank = latest.get("bank_statement");
  const insurance = latest.get("insurance");
  const salary = latest.get("salary");
  const bundle = latest.get("bundle");

  return (
    <div className="panel">
      <div className="row row-between">
        <div>
          <div className="hint">Package</div>
          <div className="mono">{pkg.package_id}</div>
        </div>
        <div className={pkg.status === "generated" ? "badge badge-green" : pkg.status === "error" ? "badge badge-red" : "badge badge-gray"}>
          {pkg.status}
        </div>
      </div>

      <div className="grid2 mt">
        <div>
          <div className="hint">Client</div>
          <div>{pkg.client.full_name}</div>
          <div className="hint">{pkg.client.passport_masked}</div>
        </div>
        <div>
          <div className="hint">Company</div>
          <div>{pkg.company.name}</div>
        </div>
      </div>

      <div className="hr" />

      <div className="hint">Documents (latest)</div>
      <div className="docs">
        <DocRow label="Contract PDF" doc={contract} />
        <DocRow label="Bank statement PDF" doc={bank} />
        <DocRow label="Insurance PDF" doc={insurance} />
        <DocRow label="Salary certificate PDF" doc={salary} />
        <DocRow label="Bundle ZIP" doc={bundle} />
      </div>
    </div>
  );
}

