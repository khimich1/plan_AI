import { useEffect, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Spinner } from "@/shared/ui/Spinner";
import { isPlanVersionConflict } from "@/shared/lib/planConflict";
import {
  useSaveWorkCalendarMutation,
  useWorkCalendarQuery,
} from "@/features/production/hooks/useProductionQueries";

const sortDates = (dates: string[]) => [...dates].sort();

const formatDate = (iso: string) => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso;
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y}`;
};

export const WorkCalendarEditor = () => {
  const query = useWorkCalendarQuery();
  const saveMutation = useSaveWorkCalendarMutation();

  const [holidays, setHolidays] = useState<string[]>([]);
  const [workdays, setWorkdays] = useState<string[]>([]);
  const [newHoliday, setNewHoliday] = useState("");
  const [newWorkday, setNewWorkday] = useState("");

  useEffect(() => {
    if (query.data) {
      setHolidays(sortDates(query.data.extra_holidays ?? []));
      setWorkdays(sortDates(query.data.extra_workdays ?? []));
    }
  }, [query.data]);

  const addHoliday = () => {
    if (!newHoliday) return;
    if (!holidays.includes(newHoliday)) {
      setHolidays(sortDates([...holidays, newHoliday]));
    }
    setNewHoliday("");
  };

  const addWorkday = () => {
    if (!newWorkday) return;
    if (!workdays.includes(newWorkday)) {
      setWorkdays(sortDates([...workdays, newWorkday]));
    }
    setNewWorkday("");
  };

  const removeHoliday = (iso: string) => {
    setHolidays(holidays.filter((d) => d !== iso));
  };

  const removeWorkday = (iso: string) => {
    setWorkdays(workdays.filter((d) => d !== iso));
  };

  const handleSave = () => {
    saveMutation.mutate({
      extra_holidays: holidays,
      extra_workdays: workdays,
    });
  };

  if (query.isLoading) {
    return (
      <Card title="Производственный календарь">
        <div style={{ display: "flex", gap: "0.5rem" }}>
          <Spinner /> Загрузка…
        </div>
      </Card>
    );
  }

  if (query.isError) {
    return (
      <Card title="Производственный календарь">
        <Alert tone="error">Не удалось загрузить календарь.</Alert>
      </Card>
    );
  }

  return (
    <Card
      title="Производственный календарь"
      subtitle="Дополнительные выходные и рабочие дни, которые переопределяют стандартные выходные (суббота/воскресенье)."
    >
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: "1.25rem" }}>
        <section>
          <h3 style={{ marginTop: 0 }}>Дополнительные выходные</h3>
          <FieldWrapper label="Добавить дату">
            <div style={{ display: "flex", gap: "0.5rem" }}>
              <Input
                type="date"
                value={newHoliday}
                onChange={(e) => setNewHoliday(e.target.value)}
              />
              <Button variant="secondary" onClick={addHoliday} disabled={!newHoliday}>
                Добавить
              </Button>
            </div>
          </FieldWrapper>
          <ul style={{ listStyle: "none", padding: 0, margin: "0.75rem 0 0" }}>
            {holidays.length === 0 && <li style={{ color: "#475467" }}>Пусто</li>}
            {holidays.map((iso) => (
              <li
                key={iso}
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  padding: "0.5rem 0.75rem",
                  borderBottom: "1px solid #e4e7ec",
                }}
              >
                <span>{formatDate(iso)}</span>
                <Button variant="ghost" onClick={() => removeHoliday(iso)}>
                  Удалить
                </Button>
              </li>
            ))}
          </ul>
        </section>

        <section>
          <h3 style={{ marginTop: 0 }}>Дополнительные рабочие дни</h3>
          <FieldWrapper label="Добавить дату">
            <div style={{ display: "flex", gap: "0.5rem" }}>
              <Input
                type="date"
                value={newWorkday}
                onChange={(e) => setNewWorkday(e.target.value)}
              />
              <Button variant="secondary" onClick={addWorkday} disabled={!newWorkday}>
                Добавить
              </Button>
            </div>
          </FieldWrapper>
          <ul style={{ listStyle: "none", padding: 0, margin: "0.75rem 0 0" }}>
            {workdays.length === 0 && <li style={{ color: "#475467" }}>Пусто</li>}
            {workdays.map((iso) => (
              <li
                key={iso}
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  padding: "0.5rem 0.75rem",
                  borderBottom: "1px solid #e4e7ec",
                }}
              >
                <span>{formatDate(iso)}</span>
                <Button variant="ghost" onClick={() => removeWorkday(iso)}>
                  Удалить
                </Button>
              </li>
            ))}
          </ul>
        </section>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: "1.25rem", gap: "0.75rem", alignItems: "center" }}>
        {saveMutation.isSuccess && <span style={{ color: "#067647" }}>Сохранено ✓</span>}
        {saveMutation.isError && !isPlanVersionConflict(saveMutation.error) && (
          <span style={{ color: "#b42318" }}>Ошибка сохранения</span>
        )}
        <Button onClick={handleSave} disabled={saveMutation.isPending}>
          {saveMutation.isPending ? "Сохранение…" : "Сохранить"}
        </Button>
      </div>
    </Card>
  );
};
