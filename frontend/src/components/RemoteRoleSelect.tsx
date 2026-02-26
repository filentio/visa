import { useMemo } from "react";

type Props = {
  value: string;
  onChange: (value: string) => void;
  disabled?: boolean;
};

const PRESETS = [
  "Senior business analyst",
  "Business analyst",
  "Project manager",
  "Product manager",
  "Software engineer",
  "QA engineer",
  "DevOps engineer",
  "Data analyst",
  "Designer",
  "Account manager",
];

export function RemoteRoleSelect({ value, onChange, disabled }: Props) {
  const preset = useMemo(() => (PRESETS.includes(value) ? value : "Other"), [value]);

  return (
    <div className="field">
      <label className="label">Position</label>
      <div className="row">
        <select
          className="input"
          value={preset}
          onChange={(e) => {
            const v = e.target.value;
            if (v === "Other") onChange("");
            else onChange(v);
          }}
          disabled={disabled}
        >
          {PRESETS.map((p) => (
            <option key={p} value={p}>
              {p}
            </option>
          ))}
          <option value="Other">Other…</option>
        </select>
        <input
          className="input"
          placeholder="Other position"
          value={preset === "Other" ? value : ""}
          onChange={(e) => onChange(e.target.value)}
          disabled={disabled || preset !== "Other"}
        />
      </div>
    </div>
  );
}

