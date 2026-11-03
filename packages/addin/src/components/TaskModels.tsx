import { useEffect, useState } from "react";
import { Button, Caption1, Card, Dropdown, Field, Option, makeStyles, tokens } from "@fluentui/react-components";
import { ChevronDown16Regular, ChevronRight16Regular } from "@fluentui/react-icons";
import type { TaskView } from "@miai/shared";
import { useBackend } from "../backend";
import { useI18n } from "../i18n";

const useStyles = makeStyles({
  card: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
  warn: { color: tokens.colorPaletteDarkOrangeForeground1 },
});

interface Props {
  /** Model pr. opgave-id (Settings.taskModels). */
  value: Record<string, string>;
  /** Modeller på den aktive endpoint. */
  models: string[];
  onChange: (next: Record<string, string>) => void;
}

/**
 * Model pr. opgave (#13): fx en lille, hurtig model til grammatik og en større til formler og diagrammer.
 * Motoren vælger: forespørgsel → model pr. opgave → standardmodel → første model fra endpointen.
 */
export function TaskModels({ value, models, onChange }: Props) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const api = useBackend();
  const [tasks, setTasks] = useState<TaskView[]>([]);
  const [open, setOpen] = useState(Object.keys(value).length > 0);
  const chosen = Object.keys(value).length;

  useEffect(() => {
    api.tasks().then(setTasks, () => setTasks([]));
  }, [api]);

  const set = (id: string, model: string) => {
    const next = { ...value };
    if (model) next[id] = model;
    else delete next[id];
    onChange(next);
  };

  return (
    <>
      <div>
        <Button appearance="subtle" icon={open ? <ChevronDown16Regular /> : <ChevronRight16Regular />} onClick={() => setOpen(!open)}>
          {t("taskModels")}
          {chosen > 0 ? ` (${chosen})` : ""}
        </Button>
      </div>
      {open && (
        <Card className={s.card}>
          <Caption1>{t("taskModelsHint")}</Caption1>
          {tasks.map((task) => {
            const current = value[task.id] ?? "";
            // En valgt model, som endpointen ikke (længere) har, vises stadig – med en advarsel.
            const missing = !!current && models.length > 0 && !models.includes(current);
            const options = current && !models.includes(current) ? [current, ...models] : models;
            return (
              <Field
                key={task.id}
                label={task.label[lang]}
                validationState={missing ? "warning" : "none"}
                validationMessage={missing ? t("taskModelMissing") : undefined}
              >
                <Dropdown value={current || t("useDefaultModel")} selectedOptions={[current]} onOptionSelect={(_, d) => set(task.id, d.optionValue ?? "")}>
                  <Option value="">{t("useDefaultModel")}</Option>
                  {options.map((m) => (
                    <Option key={m} value={m}>
                      {m}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
            );
          })}
          {models.length === 0 && <Caption1 className={s.warn}>{t("taskModelsNoModels")}</Caption1>}
        </Card>
      )}
    </>
  );
}
