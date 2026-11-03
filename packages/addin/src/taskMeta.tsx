import type { FluentIcon } from "@fluentui/react-icons";
import {
  ArrowRepeatAll20Regular,
  BotSparkle20Regular,
  CalculatorMultiple20Regular,
  DataBarVertical20Regular,
  DocumentEdit20Regular,
  Flowchart20Regular,
  MathFormula20Regular,
  Megaphone20Regular,
  PuzzlePiece20Regular,
  SlideAdd20Regular,
  Sparkle20Regular,
  TextBulletListSquare20Regular,
  TextBulletListTree20Regular,
  TextEditStyle20Regular,
  TextGrammarWand20Regular,
  TextProofingTools20Regular,
  Translate20Regular,
} from "@fluentui/react-icons";
import type { TaskView } from "@miai/shared";
import type { TaskCategory } from "./theme";

/** Ikon og kategori pr. indbygget opgave (#17). Brugerdefinerede opgaver får "custom". */
const META: Record<string, { icon: FluentIcon; category: TaskCategory }> = {
  grammar: { icon: TextProofingTools20Regular, category: "write" },
  rewrite: { icon: TextEditStyle20Regular, category: "write" },
  paraphrase: { icon: ArrowRepeatAll20Regular, category: "write" },
  simplify: { icon: TextGrammarWand20Regular, category: "write" },
  translate: { icon: Translate20Regular, category: "write" },
  summarize: { icon: TextBulletListSquare20Regular, category: "write" },
  "expand-article": { icon: DocumentEdit20Regular, category: "write" },
  "sales-copy": { icon: Megaphone20Regular, category: "write" },
  outline: { icon: TextBulletListTree20Regular, category: "present" },
  "text-to-slides": { icon: SlideAdd20Regular, category: "present" },
  diagram: { icon: Flowchart20Regular, category: "present" },
  "explain-formula": { icon: MathFormula20Regular, category: "analyze" },
  "generate-formula": { icon: CalculatorMultiple20Regular, category: "analyze" },
  "analyze-data": { icon: DataBarVertical20Regular, category: "analyze" },
  agent: { icon: BotSparkle20Regular, category: "agent" },
};

export function taskMeta(task: Pick<TaskView, "id" | "custom">): { icon: FluentIcon; category: TaskCategory } {
  if (task.custom) return { icon: PuzzlePiece20Regular, category: "custom" };
  return META[task.id] ?? { icon: Sparkle20Regular, category: "write" };
}

/** Rækkefølgen, kategorierne vises i. */
export const CATEGORY_ORDER: TaskCategory[] = ["agent", "write", "present", "analyze", "custom"];
