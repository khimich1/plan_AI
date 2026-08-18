import { useState, type CSSProperties, type FormEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Spinner } from "@/shared/ui/Spinner";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import {
  useCreateDriverMutation,
  useGsmDriversQuery,
  usePatchDriverMutation,
} from "@/features/gsm/hooks/useGsmQueries";

const sectionStyle: CSSProperties = {
  display: "grid",
  gap: "0.75rem",
  padding: "1rem",
  borderRadius: 12,
  border: "1px solid #e4e7ec",
  background: "#ffffff",
};

const thStyle: CSSProperties = { padding: "0.5rem", textAlign: "left", borderBottom: "1px solid #eaecf0" };
const tdStyle: CSSProperties = { padding: "0.5rem", borderBottom: "1px solid #f2f4f7", verticalAlign: "middle" };

const emptyForm = {
  full_name: "",
  license_number: "",
  personnel_number: "",
};

export const DriversRegistryView = () => {
  const driversQuery = useGsmDriversQuery(true);
  const createMutation = useCreateDriverMutation();
  const patchMutation = usePatchDriverMutation();
  const [form, setForm] = useState(emptyForm);
  const [formError, setFormError] = useState<string | null>(null);
  const [actionError, setActionError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);

  const drivers = driversQuery.data ?? [];

  const onCreate = async (event: FormEvent) => {
    event.preventDefault();
    const full_name = form.full_name.trim();
    const license_number = form.license_number.trim();
    if (!full_name || !license_number) {
      setFormError("Укажите ФИО и номер удостоверения.");
      return;
    }
    setFormError(null);
    setActionError(null);
    try {
      await createMutation.mutateAsync({
        full_name,
        license_number,
        personnel_number: form.personnel_number.trim() || null,
      });
      setForm(emptyForm);
      setInfo(`Водитель «${full_name}» добавлен.`);
    } catch (err) {
      setFormError(formatGsmError(err));
    }
  };

  const archiveDriver = async (id: number, name: string) => {
    setActionError(null);
    setInfo(null);
    try {
      await patchMutation.mutateAsync({ id, payload: { is_active: false } });
      setInfo(`Водитель «${name}» архивирован.`);
    } catch (err) {
      setActionError(formatGsmError(err));
    }
  };

  if (driversQuery.isLoading) {
    return (
      <div style={sectionStyle} aria-label="Водители ГСМ">
        <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Водители</h2>
        <Spinner />
      </div>
    );
  }

  if (driversQuery.error) {
    return (
      <div style={sectionStyle} aria-label="Водители ГСМ">
        <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Водители</h2>
        <Alert tone="error">{formatGsmError(driversQuery.error)}</Alert>
      </div>
    );
  }

  return (
    <div style={sectionStyle} aria-label="Водители ГСМ">
      <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Водители ({drivers.length})</h2>
      {info && <Alert tone="success">{info}</Alert>}
      {actionError && <Alert tone="error">{actionError}</Alert>}

      <form
        onSubmit={(e) => void onCreate(e)}
        style={{ display: "grid", gap: "0.65rem", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))" }}
      >
        <FieldWrapper label="ФИО">
          <Input
            value={form.full_name}
            onChange={(e) => setForm((f) => ({ ...f, full_name: e.target.value }))}
            placeholder="Фамилия Имя Отчество"
          />
        </FieldWrapper>
        <FieldWrapper label="Удостоверение">
          <Input
            value={form.license_number}
            onChange={(e) => setForm((f) => ({ ...f, license_number: e.target.value }))}
            placeholder="44 21 846315"
          />
        </FieldWrapper>
        <FieldWrapper label="Табельный №">
          <Input
            value={form.personnel_number}
            onChange={(e) => setForm((f) => ({ ...f, personnel_number: e.target.value }))}
            placeholder="необязательно"
          />
        </FieldWrapper>
        <div style={{ display: "flex", alignItems: "end" }}>
          <Button type="submit" disabled={createMutation.isPending}>
            {createMutation.isPending ? "Добавление…" : "Добавить"}
          </Button>
        </div>
      </form>
      {formError && <Alert tone="error">{formError}</Alert>}

      {drivers.length === 0 ? (
        <Alert tone="info">Водители не найдены.</Alert>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.92rem" }}>
            <thead>
              <tr>
                <th style={thStyle}>ФИО</th>
                <th style={thStyle}>Удостоверение</th>
                <th style={thStyle}>Табельный</th>
                <th style={thStyle}>Действия</th>
              </tr>
            </thead>
            <tbody>
              {drivers.map((driver) => (
                <tr key={driver.id}>
                  <td style={tdStyle}>{driver.full_name}</td>
                  <td style={tdStyle}>{driver.license_number}</td>
                  <td style={tdStyle}>{driver.personnel_number ?? "—"}</td>
                  <td style={tdStyle}>
                    <Button
                      type="button"
                      variant="danger"
                      onClick={() => void archiveDriver(driver.id, driver.full_name)}
                      disabled={patchMutation.isPending}
                    >
                      Архив
                    </Button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};
