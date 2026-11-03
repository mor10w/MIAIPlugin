import { useCallback, useEffect, useRef, useState } from "react";
import {
  Badge,
  Button,
  Card,
  Checkbox,
  Dropdown,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Option,
  Spinner,
  Subtitle2,
  Textarea,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { UserTaskFile as UserTaskSchema, localize, type UserTaskFile } from "@miai/shared";
import { useBackend } from "../backend";
import type { UserTaskInfo } from "../backend/types";
import { useI18n } from "../i18n";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  row: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center", flexWrap: "wrap" },
  listRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    alignItems: "center",
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  listLabel: { flex: 1, minWidth: 0, overflowWrap: "anywhere" },
  card: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
  twoCol: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: tokens.spacingHorizontalS },
  paramCard: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
  mono: { fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200 },
});

/** Redigerbar form for én brugerdefineret opgave (F14). */
interface FormState {
  id: string;
  labelDa: string;
  labelEn: string;
  descDa: string;
  descEn: string;
  hosts: ("word" | "excel" | "powerpoint")[];
  input: "selection" | "selectionOrDocument" | "prompt";
  format: "plain" | "markdown" | "tsv";
  applyModes: FileApplyMode[];
  temperature: string;
  system: string;
  user: string;
  params: ParamEdit[];
}

/** De anvendelsesmuligheder, som en brugeropgave kan have (formula-output er ikke tilladt i filer). */
type FileApplyMode = UserTaskFile["applyModes"][number];

interface ParamEdit {
  name: string;
  labelDa: string;
  labelEn: string;
  kind: "text" | "select";
  /** Én værdi pr. linje, evt. "værdi | etikette". */
  optionsText: string;
  defaultValue: string;
}

const emptyForm = (): FormState => ({
  id: "",
  labelDa: "",
  labelEn: "",
  descDa: "",
  descEn: "",
  hosts: [],
  input: "prompt",
  format: "plain",
  applyModes: ["insertAfter"],
  temperature: "",
  system: "",
  user: "",
  params: [],
});

/** "værdi | etikette" pr. linje → options (mangler etikette = værdien). */
function parseOptions(text: string): { value: string; label: string }[] {
  return text
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean)
    .map((l) => {
      const [value, ...rest] = l.split("|");
      return { value: value.trim(), label: (rest.join("|") || "").trim() || value.trim() };
    });
}

function specToForm(spec: UserTaskFile): FormState {
  return {
    id: spec.id,
    labelDa: spec.label.da,
    labelEn: spec.label.en ?? "",
    descDa: spec.description?.da ?? "",
    descEn: spec.description?.en ?? "",
    hosts: [...spec.hosts],
    input: spec.input,
    format: spec.output.format,
    applyModes: [...spec.applyModes],
    temperature: spec.temperature !== undefined ? String(spec.temperature) : "",
    system: spec.prompt.system,
    user: spec.prompt.user,
    params: spec.params.map((p) => ({
      name: p.name,
      labelDa: p.label.da,
      labelEn: p.label.en ?? "",
      kind: p.kind,
      optionsText: (p.options ?? []).map((o) => (o.label.da === o.value ? o.value : `${o.value} | ${o.label.da}`)).join("\n"),
      defaultValue: p.default ?? "",
    })),
  };
}

function formToSpec(f: FormState): UserTaskFile {
  return {
    id: f.id.trim(),
    label: { da: f.labelDa, en: f.labelEn || undefined },
    description: f.descDa ? { da: f.descDa, en: f.descEn || undefined } : undefined,
    hosts: [...f.hosts],
    input: f.input,
    output: { format: f.format },
    params: f.params.map((p) => ({
      name: p.name.trim(),
      label: { da: p.labelDa, en: p.labelEn || undefined },
      kind: p.kind,
      options: p.kind === "select" ? parseOptions(p.optionsText).map((o) => ({ value: o.value, label: { da: o.label } })) : undefined,
      default: p.defaultValue || undefined,
    })),
    applyModes: [...f.applyModes],
    temperature: f.temperature !== "" && !Number.isNaN(Number(f.temperature)) ? Number(f.temperature) : undefined,
    prompt: { system: f.system, user: f.user },
  };
}

export function TasksTab() {
  const { t, lang } = useI18n();
  const s = useStyles();
  const api = useBackend();
  const [tasks, setTasks] = useState<UserTaskInfo[]>();
  const [form, setForm] = useState<FormState>();
  const [msg, setMsg] = useState<{ intent: "success" | "error"; text: string }>();
  const fileRef = useRef<HTMLInputElement>(null);

  const refresh = useCallback(() => api.userTasks().then(setTasks, (e) => setMsg({ intent: "error", text: (e as Error).message })), [api]);
  useEffect(() => {
    void refresh();
  }, [refresh]);

  const patch = (p: Partial<FormState>) => setForm((f) => ({ ...f!, ...p }));
  const patchParam = (i: number, p: Partial<ParamEdit>) => setForm((f) => ({ ...f!, params: f!.params.map((x, j) => (j === i ? { ...x, ...p } : x)) }));

  const appLabel = { word: t("appWord"), excel: t("appExcel"), powerpoint: t("appPowerPoint") };

  const save = async () => {
    if (!form) return;
    const parsed = UserTaskSchema.safeParse(formToSpec(form));
    if (!parsed.success) {
      setMsg({ intent: "error", text: parsed.error.issues.map((i) => `${i.path.join(".")}: ${i.message}`).join("\n") });
      return;
    }
    try {
      await api.saveUserTask(parsed.data);
      setForm(undefined);
      setMsg({ intent: "success", text: t("saved") });
      void refresh();
    } catch (e) {
      setMsg({ intent: "error", text: (e as Error).message });
    }
  };

  /** Eksportér opgaven som JSON-fil – deling sker ved at kopiere filen/mappen. */
  const exportTask = (info: UserTaskInfo) => {
    const blob = new Blob([JSON.stringify(info.spec, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${info.spec.id}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const del = async (id: string) => {
    if (!window.confirm(t("confirmDeleteTask"))) return;
    try {
      await api.deleteUserTask(id);
      if (form?.id === id) setForm(undefined);
      void refresh();
    } catch (e) {
      setMsg({ intent: "error", text: (e as Error).message });
    }
  };

  /** Importér en fil: udfyld redigeringsfeltet, så brugeren kan gennemse og gemme. */
  const onImportFile = async (file: File) => {
    try {
      const parsed = UserTaskSchema.safeParse(JSON.parse(await file.text()));
      if (!parsed.success) throw new Error(parsed.error.issues.map((i) => `${i.path.join(".")}: ${i.message}`).join("\n"));
      setForm(specToForm(parsed.data));
      setMsg({ intent: "success", text: t("imported") });
    } catch (e) {
      setMsg({ intent: "error", text: (e as Error).message });
    }
  };

  const toggle = <T,>(list: T[], v: T): T[] => (list.includes(v) ? list.filter((x) => x !== v) : [...list, v]);

  return (
    <div className={s.root}>
      <div className={s.row}>
        <Button
          appearance="primary"
          onClick={() => {
            setForm(emptyForm());
            setMsg(undefined);
          }}
        >
          {t("newTask")}
        </Button>
        <Button onClick={() => fileRef.current?.click()}>{t("importTask")}</Button>
        <input
          ref={fileRef}
          type="file"
          accept=".json,application/json"
          hidden
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) void onImportFile(f);
            e.currentTarget.value = "";
          }}
        />
      </div>

      {msg && (
        <MessageBar intent={msg.intent}>
          <MessageBarBody>{msg.text}</MessageBarBody>
        </MessageBar>
      )}

      {form ? (
        <Card className={s.card}>
          <Subtitle2>{t("editTask")}</Subtitle2>
          <div className={s.twoCol}>
            <Field label={t("taskFieldId")}>
              <Input
                className={s.mono}
                value={form.id}
                disabled={tasks?.some((x) => x.spec.id === form.id && x.source === "machine")}
                onChange={(_, d) => patch({ id: d.value })}
              />
            </Field>
            <Field label={t("temperature")}>
              <Input type="number" min={0} max={2} step={0.1} value={form.temperature} onChange={(_, d) => patch({ temperature: d.value })} />
            </Field>
          </div>
          <div className={s.twoCol}>
            <Field label={t("taskLabelDa")}>
              <Input value={form.labelDa} onChange={(_, d) => patch({ labelDa: d.value })} />
            </Field>
            <Field label={t("taskLabelEn")}>
              <Input value={form.labelEn} onChange={(_, d) => patch({ labelEn: d.value })} />
            </Field>
          </div>
          <div className={s.twoCol}>
            <Field label={`${t("taskDescDa")} (valgfri)`}>
              <Input value={form.descDa} onChange={(_, d) => patch({ descDa: d.value })} />
            </Field>
            <Field label={`${t("taskDescEn")} (valgfri)`}>
              <Input value={form.descEn} onChange={(_, d) => patch({ descEn: d.value })} />
            </Field>
          </div>
          <Field label={t("taskApps")}>
            <div className={s.row}>
              {(["word", "excel", "powerpoint"] as const).map((h) => (
                <Checkbox key={h} label={appLabel[h]} checked={form.hosts.includes(h)} onChange={() => patch({ hosts: toggle(form.hosts, h) })} />
              ))}
            </div>
          </Field>
          <div className={s.twoCol}>
            <Field label={t("inputKind")}>
              <Dropdown
                value={{ selection: t("inputSelection"), selectionOrDocument: t("inputSelOrDoc"), prompt: t("inputPrompt") }[form.input]}
                selectedOptions={[form.input]}
                onOptionSelect={(_, d) => patch({ input: d.optionValue as FormState["input"] })}
              >
                <Option value="selection">{t("inputSelection")}</Option>
                <Option value="selectionOrDocument">{t("inputSelOrDoc")}</Option>
                <Option value="prompt">{t("inputPrompt")}</Option>
              </Dropdown>
            </Field>
            <Field label={t("outputFormat")}>
              <Dropdown
                value={{ plain: t("formatPlain"), markdown: t("formatMarkdown"), tsv: t("formatTsv") }[form.format]}
                selectedOptions={[form.format]}
                onOptionSelect={(_, d) => patch({ format: d.optionValue as FormState["format"] })}
              >
                <Option value="plain">{t("formatPlain")}</Option>
                <Option value="markdown">{t("formatMarkdown")}</Option>
                <Option value="tsv">{t("formatTsv")}</Option>
              </Dropdown>
            </Field>
          </div>
          <Field label={t("applyModesLabel")}>
            <div className={s.row}>
              {(["replace", "trackChanges", "insertAfter", "comment", "cells"] as const).map((m) => (
                <Checkbox
                  key={m}
                  label={t(`apply_${m}`)}
                  checked={form.applyModes.includes(m)}
                  onChange={() => patch({ applyModes: toggle(form.applyModes, m) })}
                />
              ))}
            </div>
          </Field>

          <Subtitle2>{t("paramsLabel")}</Subtitle2>
          {form.params.map((p, i) => (
            <Card key={i} className={s.paramCard}>
              <div className={s.twoCol}>
                <Field label={t("paramName")}>
                  <Input className={s.mono} value={p.name} onChange={(_, d) => patchParam(i, { name: d.value })} />
                </Field>
                <Field label={t("paramKind")}>
                  <Dropdown
                    value={p.kind === "text" ? t("kindText") : t("kindSelect")}
                    selectedOptions={[p.kind]}
                    onOptionSelect={(_, d) => patchParam(i, { kind: d.optionValue as ParamEdit["kind"] })}
                  >
                    <Option value="text">{t("kindText")}</Option>
                    <Option value="select">{t("kindSelect")}</Option>
                  </Dropdown>
                </Field>
              </div>
              <div className={s.twoCol}>
                <Field label={t("paramLabelDa")}>
                  <Input value={p.labelDa} onChange={(_, d) => patchParam(i, { labelDa: d.value })} />
                </Field>
                <Field label={t("paramLabelEn")}>
                  <Input value={p.labelEn} onChange={(_, d) => patchParam(i, { labelEn: d.value })} />
                </Field>
              </div>
              {p.kind === "select" && (
                <Field label={t("paramOptions")}>
                  <Textarea rows={3} className={s.mono} value={p.optionsText} onChange={(_, d) => patchParam(i, { optionsText: d.value })} />
                </Field>
              )}
              <div className={s.row}>
                <Field label={t("paramDefault")}>
                  <Input value={p.defaultValue} onChange={(_, d) => patchParam(i, { defaultValue: d.value })} />
                </Field>
                <Button appearance="subtle" onClick={() => setForm({ ...form, params: form.params.filter((_, j) => j !== i) })}>
                  {t("remove")}
                </Button>
              </div>
            </Card>
          ))}
          <Button
            size="small"
            onClick={() => patch({ params: [...form.params, { name: "", labelDa: "", labelEn: "", kind: "text", optionsText: "", defaultValue: "" }] })}
          >
            {t("paramAdd")}
          </Button>

          <Field label={t("systemPrompt")}>
            <Textarea rows={3} className={s.mono} value={form.system} onChange={(_, d) => patch({ system: d.value })} />
          </Field>
          <Field label={t("userTemplate")} hint={t("templateHint")}>
            <Textarea rows={6} className={s.mono} value={form.user} onChange={(_, d) => patch({ user: d.value })} />
          </Field>
          <div className={s.row}>
            <Button appearance="primary" onClick={() => void save()}>
              {t("save")}
            </Button>
            <Button appearance="subtle" onClick={() => setForm(undefined)}>
              {t("cancel")}
            </Button>
          </div>
        </Card>
      ) : (
        <>
          <Subtitle2>{t("userTasks")}</Subtitle2>
          {!tasks ? (
            <Spinner />
          ) : tasks.length === 0 ? (
            <Card>{t("noTasks")}</Card>
          ) : (
            <Card>
              {tasks.map((info) => (
                <div key={info.spec.id + info.source} className={s.listRow}>
                  <span className={s.listLabel}>{localize(info.spec.label, lang)}</span>
                  <Badge appearance="tint" color={info.source === "machine" ? "informative" : "subtle"}>
                    {t(info.source === "machine" ? "sourceMachine" : "sourceUser")}
                  </Badge>
                  {info.source === "user" && (
                    <>
                      <Button
                        size="small"
                        onClick={() => {
                          setForm(specToForm(info.spec));
                          setMsg(undefined);
                        }}
                      >
                        {t("editTask")}
                      </Button>
                      <Button size="small" appearance="subtle" onClick={() => void del(info.spec.id)}>
                        {t("deleteTask")}
                      </Button>
                    </>
                  )}
                  <Button size="small" appearance="secondary" onClick={() => exportTask(info)}>
                    {t("exportTask")}
                  </Button>
                </div>
              ))}
            </Card>
          )}
        </>
      )}
    </div>
  );
}
