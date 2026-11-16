import { useEffect, useState } from "react";
import { Button, Caption1, Card, Dropdown, Field, Option, makeStyles, tokens } from "@fluentui/react-components";
import { ChevronDown16Regular, ChevronRight16Regular } from "@fluentui/react-icons";
import { MODEL_GROUPS, modelGroupOf, type TaskView } from "@miai/shared";
import { useBackend } from "../backend";
import { useI18n } from "../i18n";

const useStyles = makeStyles({
  group: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  card: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
  warn: { color: tokens.colorPaletteDarkOrangeForeground1 },
});

interface Props {
  /** Model pr. opgavetype (Settings.groupModels). */
  groups: Record<string, string>;
  /** Undtagelser for enkelte opgaver (Settings.taskModels). */
  tasks: Record<string, string>;
  /** Modeller på den aktive endpoint. */
  models: string[];
  onChange: (next: { groupModels: Record<string, string>; taskModels: Record<string, string> }) => void;
}

/**
 * Model pr. opgavetype (#13): fx en lille, hurtig model til at rette og omskrive og en større til
 * agenten. Undtagelser for enkelte opgaver ligger sammenfoldet nedenunder.
 * Motoren vælger: forespørgsel → opgavens undtagelse → typen → standardmodel → første fra endpointen.
 */
export function TaskModels({ groups, tasks, models, onChange }: Props) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const api = useBackend();
  const [taskList, setTaskList] = useState<TaskView[]>([]);
  const [open, setOpen] = useState(Object.keys(tasks).length > 0);
  const exceptions = Object.keys(tasks).length;

  useEffect(() => {
    api.tasks().then(setTaskList, () => setTaskList([]));
  }, [api]);

  const update = (map: Record<string, string>, key: string, model: string) => {
    const next = { ...map };
    if (model) next[key] = model;
    else delete next[key];
    return next;
  };
  const groupLabel = (taskId: string) => MODEL_GROUPS.find((g) => g.id === modelGroupOf(taskId))!.label[lang];

  /** Rullemenu med modellerne; en valgt model, som endpointen ikke (længere) har, vises med en advarsel. */
  const picker = (key: string, label: string, hint: string | undefined, current: string, empty: string, set: (model: string) => void) => {
    const missing = !!current && models.length > 0 && !models.includes(current);
    const options = current && !models.includes(current) ? [current, ...models] : models;
    return (
      <Field
        key={key}
        label={label}
        hint={missing ? undefined : hint}
        validationState={missing ? "warning" : "none"}
        validationMessage={missing ? t("taskModelMissing") : undefined}
      >
        <Dropdown aria-label={label} value={current || empty} selectedOptions={[current]} onOptionSelect={(_, d) => set(d.optionValue ?? "")}>
          <Option value="">{empty}</Option>
          {options.map((m) => (
            <Option key={m} value={m}>
              {m}
            </Option>
          ))}
        </Dropdown>
      </Field>
    );
  };

  return (
    <>
      <div className={s.group}>
        <Caption1>{t("groupModelsHint")}</Caption1>
        {MODEL_GROUPS.map((g) =>
          picker(g.id, g.label[lang], g.hint[lang], groups[g.id] ?? "", t("useDefaultModel"), (m) =>
            onChange({ groupModels: update(groups, g.id, m), taskModels: tasks }),
          ),
        )}
        {models.length === 0 && <Caption1 className={s.warn}>{t("taskModelsNoModels")}</Caption1>}
      </div>
      <div>
        <Button appearance="subtle" icon={open ? <ChevronDown16Regular /> : <ChevronRight16Regular />} onClick={() => setOpen(!open)}>
          {t("taskModels")}
          {exceptions > 0 ? ` (${exceptions})` : ""}
        </Button>
      </div>
      {open && (
        <Card className={s.card}>
          <Caption1>{t("taskModelsHint")}</Caption1>
          {taskList.map((task) =>
            picker(task.id, task.label[lang], groupLabel(task.id), tasks[task.id] ?? "", t("useGroupModel"), (m) =>
              onChange({ groupModels: groups, taskModels: update(tasks, task.id, m) }),
            ),
          )}
        </Card>
      )}
    </>
  );
}
