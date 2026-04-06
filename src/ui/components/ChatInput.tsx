import * as React from "react";
import { useRef, useEffect } from "react";
import { Textarea, Button, Combobox, Option, Tooltip, makeStyles } from "@fluentui/react-components";
import { Send24Regular, Dismiss24Regular, Stop24Regular } from "@fluentui/react-icons";
import { z } from "zod";
import { filterModels } from "../lib/model-search";
import type { ModelInfo } from "../lib/opencode-client";
import { modelInfoSchema } from "../lib/opencode-schemas";

const ImageAttachmentSchema = z.object({
  id: z.string().min(1),
  dataUrl: z.string().min(1),
  name: z.string().min(1),
});

export type ImageAttachment = z.infer<typeof ImageAttachmentSchema>;

interface ChatInputProps {
  value: string;
  onChange: (value: string) => void;
  onSend: () => void;
  onStop?: () => void;
  disabled?: boolean;
  isRunning?: boolean;
  images?: ImageAttachment[];
  onImagesChange?: (images: ImageAttachment[]) => void;
  selectedModel: string;
  onModelChange: (model: string) => void;
  models: ModelInfo[];
  selectedVariant?: string;
  onVariantChange: (variant: string | undefined) => void;
}

const useStyles = makeStyles({
  dock: {
    width: "min(calc(100% - 24px), 760px)",
    margin: "0 auto 12px",
    display: "flex",
    flexDirection: "column",
    alignItems: "stretch",
    boxSizing: "border-box",
  },
  tray: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "6px 10px",
    borderRadius: "12px 12px 0 0",
    border: "1px solid var(--oc-border)",
    borderBottom: "none",
    background: "var(--oc-bg-soft)",
    marginBottom: "-1px",
    position: "relative",
    zIndex: 1,
  },
  trayField: {
    flex: "1 1 0",
    minWidth: 0,
  },
  variantField: {
    flex: "0 0 100px",
    minWidth: "80px",
  },
  control: {
    minWidth: 0,
    width: "100%",
    height: "28px",
    fontSize: "11px",
    borderRadius: "8px",
    background: "var(--oc-bg)",
    border: "1px solid var(--oc-border) !important",
    padding: "0 8px",
    boxSizing: "border-box",
    ":hover": {
      background: "var(--oc-bg-soft-hover)",
    },
  },
  shell: {
    display: "flex",
    flexDirection: "column",
    borderRadius: "12px",
    backgroundColor: "var(--oc-bg)",
    border: "1px solid var(--oc-border)",
    boxShadow: "var(--oc-shadow)",
    overflow: "hidden",
    position: "relative",
    zIndex: 2,
  },
  shellHasTray: {
    borderTopLeftRadius: "0",
    borderTopRightRadius: "0",
  },
  body: {
    display: "flex",
    flexDirection: "column",
    padding: "10px 12px 8px",
  },
  input: {
    flex: 1,
    width: "100%",
    maxWidth: "100%",
    boxSizing: "border-box",
    minHeight: "52px",
    padding: "0",
    borderRadius: "0",
    border: "none !important",
    backgroundColor: "transparent !important",
    outline: "none !important",
    boxShadow: "none !important",
    color: "var(--text-strong)",
    fontSize: "14px",
    lineHeight: "1.5",
    "::after": {
      display: "none !important",
    },
    "& textarea": {
      padding: "0",
      minHeight: "52px",
    },
  },
  inputWrap: {
    flex: 1,
    width: "100%",
    minWidth: 0,
    padding: "0 0 4px",
  },
  footer: {
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-end",
    padding: "0 12px 10px",
    gap: "6px",
  },
  sendButton: {
    width: "34px",
    height: "34px",
    minWidth: "34px",
    padding: "0",
    backgroundColor: "var(--oc-accent)",
    border: "none",
    borderRadius: "10px",
    color: "var(--text-on-interactive-base, white)",
    ":hover": {
      backgroundColor: "var(--oc-accent-strong)",
    },
  },
  stopButton: {
    backgroundColor: "var(--oc-bg-soft)",
    color: "var(--text-strong)",
    border: "1px solid var(--oc-border)",
    ":hover": {
      backgroundColor: "var(--oc-bg-soft-hover)",
    },
  },
  attachmentStrip: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
    padding: "8px 12px 0",
  },
  thumbnailFrame: {
    position: "relative",
    width: "56px",
    height: "56px",
    borderRadius: "8px",
    overflow: "hidden",
    border: "1px solid var(--oc-border)",
    background: "var(--oc-bg-soft)",
  },
  thumbnailCover: {
    width: "100%",
    height: "100%",
    objectFit: "cover",
  },
  thumbnailDismiss: {
    position: "absolute",
    top: "3px",
    right: "3px",
    minWidth: "18px",
    width: "18px",
    height: "18px",
    padding: "0",
    backgroundColor: "var(--oc-bg)",
    border: "1px solid var(--oc-border)",
    borderRadius: "50%",
    cursor: "pointer",
    ":hover": {
      backgroundColor: "var(--oc-bg-soft-hover)",
    },
  },
});

export const ChatInput: React.FC<ChatInputProps> = ({
  value,
  onChange,
  onSend,
  onStop,
  disabled = false,
  isRunning = false,
  images = [],
  onImagesChange,
  selectedModel,
  onModelChange,
  models,
  selectedVariant,
  onVariantChange,
}) => {
  const styles = useStyles();
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const safeImages = React.useMemo(() => z.array(ImageAttachmentSchema).catch([]).parse(images), [images]);
  const safeModels = React.useMemo(() => z.array(modelInfoSchema).catch([]).parse(models), [models]);
  const [modelValue, setModelValue] = React.useState("");
  const [modelOpen, setModelOpen] = React.useState(false);
  const selectedLabel = React.useMemo(
    () => safeModels.find((item) => item.key === selectedModel)?.label || selectedModel,
    [safeModels, selectedModel],
  );
  const modelItems = React.useMemo(
    () => filterModels(safeModels, modelOpen ? modelValue : ""),
    [safeModels, modelOpen, modelValue],
  );
  const modelVariants = React.useMemo(() => {
    const current = safeModels.find((item) => item.key === selectedModel);
    return current?.variants ?? [];
  }, [safeModels, selectedModel]);

  useEffect(() => {
    // Restore focus after the input is cleared on send
    if (value === "") {
      inputRef.current?.focus();
    }
  }, [value]);

  useEffect(() => {
    if (!modelOpen) {
      setModelValue(selectedLabel);
    }
  }, [modelOpen, selectedLabel]);

  const onInputKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      onSend();
    }
  };

  const onClipboardPaste = async (e: React.ClipboardEvent) => {
    const clipboardItems = e.clipboardData?.items;
    if (!clipboardItems || !onImagesChange) return;

    // Collect image files from the clipboard
    const files = Array.from(clipboardItems)
      .filter((entry) => entry.type.startsWith("image/"))
      .map((entry) => entry.getAsFile())
      .filter((file): file is File => file !== null);

    if (files.length === 0) return;
    e.preventDefault();

    const pending = files.map(
      (file, idx) =>
        new Promise<ImageAttachment | null>((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => {
            const result = typeof reader.result === "string" ? reader.result : "";
            const validated = ImageAttachmentSchema.safeParse({
              id: crypto.randomUUID(),
              dataUrl: result,
              name: `clipboard-${Date.now()}-${idx}.png`,
            });
            resolve(validated.success ? validated.data : null);
          };
          reader.onerror = () => resolve(null);
          reader.readAsDataURL(file);
        }),
    );

    const resolved = (await Promise.all(pending)).filter(
      (attachment): attachment is ImageAttachment => attachment !== null,
    );

    if (resolved.length > 0) {
      onImagesChange([...safeImages, ...resolved]);
    }
  };

  const discardAttachment = (id: string) => {
    if (onImagesChange) {
      onImagesChange(safeImages.filter((attachment) => attachment.id !== id));
    }
  };

  const hasTray = safeModels.length > 0;

  return (
    <div className={styles.dock}>
      {hasTray && (
        <div className={styles.tray}>
          <div className={styles.trayField}>
            <Combobox
              className={styles.control}
              appearance="filled-darker"
              freeform
              placeholder="Search models"
              aria-label="Model"
              value={modelValue}
              onChange={(event) => setModelValue((event.target as HTMLInputElement).value)}
              onOpenChange={(_, data) => {
                setModelOpen(data.open);
                setModelValue(data.open ? "" : selectedLabel);
              }}
              onOptionSelect={(_, data) => {
                const nextModel = data.optionValue;
                if (nextModel && nextModel !== selectedModel) {
                  onModelChange(nextModel);
                }
                setModelOpen(false);
                setModelValue(data.optionText || selectedLabel);
              }}
            >
              {modelItems.map((model) => (
                <Option key={model.key} value={model.key} text={model.label}>
                  {model.label}
                </Option>
              ))}
            </Combobox>
          </div>
          <div className={styles.variantField}>
            <Combobox
              className={styles.control}
              appearance="filled-darker"
              placeholder="Effort"
              aria-label="Model effort"
              value={selectedVariant ?? "default"}
              disabled={modelVariants.length === 0}
              onOptionSelect={(_, data) => onVariantChange(data.optionValue || undefined)}
            >
              <Option value="" text="default">
                default
              </Option>
              {modelVariants.map((variant) => (
                <Option key={variant} value={variant} text={variant}>
                  {variant}
                </Option>
              ))}
            </Combobox>
          </div>
        </div>
      )}
      <div className={`${styles.shell} ${hasTray ? styles.shellHasTray : ""}`.trim()}>
        {safeImages.length > 0 && (
          <div className={styles.attachmentStrip}>
            {safeImages.map((image) => (
              <div key={image.id} className={styles.thumbnailFrame}>
                <img src={image.dataUrl} alt="Preview" className={styles.thumbnailCover} />
                <button
                  type="button"
                  className={styles.thumbnailDismiss}
                  onClick={() => discardAttachment(image.id)}
                  title="Remove image"
                  aria-label={`Remove image ${image.name}`}
                >
                  <Dismiss24Regular style={{ fontSize: '12px' }} />
                </button>
              </div>
            ))}
          </div>
        )}
        <div className={styles.body}>
          <div className={styles.inputWrap}>
            <Textarea
              ref={inputRef}
              className={styles.input}
              value={value}
              onChange={(e, data) => onChange(data.value)}
              onKeyDown={onInputKeyDown}
              onPaste={onClipboardPaste}
              placeholder="Ask OpenCode to work on the current document..."
              rows={2}
              disabled={disabled}
            />
          </div>
        </div>
        <div className={styles.footer}>
          {isRunning && (
            <Tooltip content="Stop response" relationship="label">
              <Button
                appearance="secondary"
                icon={<Stop24Regular />}
                onClick={onStop}
                disabled={disabled || !onStop}
                aria-label="Stop response"
                className={`${styles.sendButton} ${styles.stopButton}`.trim()}
              />
            </Tooltip>
          )}
          <Tooltip content="Send message" relationship="label">
            <Button
              appearance="primary"
              icon={<Send24Regular />}
              onClick={onSend}
              disabled={disabled || (!value.trim() && safeImages.length === 0)}
              aria-label="Send message"
              className={styles.sendButton}
            />
          </Tooltip>
        </div>
      </div>
    </div>
  );
};
