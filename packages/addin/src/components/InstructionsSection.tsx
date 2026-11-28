import { Caption1, Field, Textarea, makeStyles, tokens } from "@fluentui/react-components";
import { estimateTokens, INSTRUCTION_HOSTS, MAX_INSTRUCTION_CHARS, type InstructionHost, type Policy, type Settings } from "@miai/shared";
import { useI18n } from "../i18n";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  host: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  muted: { color: tokens.colorNeutralForeground3 },
});

const APP: Record<InstructionHost, string> = { word: "Word", excel: "Excel", powerpoint: "PowerPoint", outlook: "Outlook" };

interface Props {
  settings: Settings;
  setSettings: (s: Settings) => void;
  policy: Policy;
  /** Programmet, task pane er åbnet i – dets felt vises først. */
  host?: string;
}

/**
 * Faste instruktioner pr. program (#78). Organisationens (fra politikken) vises skrivebeskyttet
 * over brugerens egne, i den rækkefølge modellen får dem. Antal tegn og et tokenestimat viser,
 * hvad instruktionerne fylder i hver forespørgsel.
 */
export function InstructionsSection({ settings, setSettings, policy, host }: Props) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const nf = new Intl.NumberFormat(lang === "da" ? "da-DK" : "en-GB");
  const hosts = [...INSTRUCTION_HOSTS].sort((a, b) => Number(b === host) - Number(a === host));
  return (
    <div className={s.root}>
      <Caption1 className={s.muted}>{t("instructionsIntro")}</Caption1>
      {hosts.map((h) => {
        const org = policy.instructions?.[h]?.trim() ?? "";
        const own = settings.instructions[h];
        return (
          <div key={h} className={s.host}>
            {org && (
              <Field label={t("instructionsOrg", { app: APP[h] })} hint={t("instructionsOrgHint", { tokens: nf.format(estimateTokens(org)) })}>
                <Textarea readOnly resize="vertical" rows={3} value={org} />
              </Field>
            )}
            <Field
              label={t("instructionsFor", { app: APP[h] })}
              hint={t("instructionsCount", { n: nf.format(own.length), max: nf.format(MAX_INSTRUCTION_CHARS), tokens: nf.format(estimateTokens(own)) })}
            >
              <Textarea
                resize="vertical"
                rows={3}
                maxLength={MAX_INSTRUCTION_CHARS}
                placeholder={t("instructionsPlaceholder")}
                value={own}
                onChange={(_, d) => setSettings({ ...settings, instructions: { ...settings.instructions, [h]: d.value.slice(0, MAX_INSTRUCTION_CHARS) } })}
              />
            </Field>
          </div>
        );
      })}
    </div>
  );
}
