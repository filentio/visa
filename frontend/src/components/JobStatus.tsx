import type { JobStatusResponse } from "../api/types";

type Props = {
  job: JobStatusResponse | null;
  jobId?: string | null;
};

function badgeClass(status: JobStatusResponse["status"]) {
  switch (status) {
    case "queued":
      return "badge badge-gray";
    case "running":
      return "badge badge-blue";
    case "done":
      return "badge badge-green";
    case "error":
      return "badge badge-red";
  }
}

export function JobStatus({ job, jobId }: Props) {
  if (!jobId) return null;
  if (!job) return <div className="panel">Job created: <code>{jobId}</code></div>;

  return (
    <div className="panel">
      <div className="row row-between">
        <div>
          <div className="hint">Job</div>
          <div>
            <code>{job.job_id}</code>
          </div>
        </div>
        <div className={badgeClass(job.status)}>{job.status}</div>
      </div>

      <div className="grid2 mt">
        <div>
          <div className="hint">Version</div>
          <div>{job.version ?? "—"}</div>
        </div>
        <div>
          <div className="hint">Started</div>
          <div>{job.started_at ?? "—"}</div>
        </div>
        <div>
          <div className="hint">Finished</div>
          <div>{job.finished_at ?? "—"}</div>
        </div>
      </div>

      {job.status === "error" && (
        <div className="error mt">
          {job.error_message || "Job failed"}
        </div>
      )}
    </div>
  );
}

