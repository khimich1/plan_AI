import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Textarea } from "@/shared/ui/Field";

export type AiInstructionBlockProps = {
  hint?: string;
  placeholder?: string;
  instruction: string;
  onInstructionChange: (value: string) => void;
  onApply: () => void;
  disabled?: boolean;
  isProcessing?: boolean;
};

export const AiInstructionBlock = ({
  hint,
  placeholder,
  instruction,
  onInstructionChange,
  onApply,
  disabled = false,
  isProcessing = false,
}: AiInstructionBlockProps) => (
  <FieldWrapper label="Инструкция для помощника" hint={hint}>
    <Textarea
      value={instruction}
      onChange={(event) => onInstructionChange(event.target.value)}
      placeholder={placeholder}
    />
    <div style={{ marginTop: "0.75rem" }}>
      <Button
        type="button"
        variant="ghost"
        onClick={onApply}
        disabled={disabled || isProcessing || !instruction.trim()}
      >
        {isProcessing ? "Обработка..." : "Применить инструкцию"}
      </Button>
    </div>
  </FieldWrapper>
);
