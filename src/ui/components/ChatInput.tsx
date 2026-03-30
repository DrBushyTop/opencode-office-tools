import * as React from "react";
import { useRef, useEffect } from "react";
import { Textarea, Button, Tooltip, makeStyles } from "@fluentui/react-components";
import { Send24Regular, Dismiss24Regular, Stop24Regular } from "@fluentui/react-icons";
import { z } from "zod";

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
}

const useStyles = makeStyles({
  inputContainer: {
    width: "min(calc(100% - 28px), 760px)",
    margin: "0 auto 12px",
    padding: "0",
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    borderRadius: "16px",
    backgroundColor: "var(--oc-bg)",
    border: "1px solid var(--oc-border)",
    boxShadow: "var(--oc-shadow)",
    overflow: "hidden",
    boxSizing: "border-box",
  },
  body: {
    display: "flex",
    alignItems: "stretch",
    gap: "12px",
    padding: "12px 16px",
  },
  input: {
    flex: 1,
    width: "100%",
    maxWidth: "100%",
    boxSizing: "border-box",
    minHeight: "72px",
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
  },
  inputWrap: {
    flex: 1,
    width: "100%",
    minWidth: 0,
    padding: "2px 4px",
  },
  sendButton: {
    width: "42px",
    height: "42px",
    minWidth: "42px",
    padding: "0",
    alignSelf: "flex-end",
    backgroundColor: "var(--oc-accent)",
    border: "none",
    borderRadius: "12px",
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
  imagePreviewContainer: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
    padding: "12px 12px 0",
  },
  imagePreview: {
    position: "relative",
    width: "64px",
    height: "64px",
    borderRadius: "10px",
    overflow: "hidden",
    border: "1px solid var(--oc-border)",
    background: "var(--oc-bg-soft)",
  },
  imagePreviewImg: {
    width: "100%",
    height: "100%",
    objectFit: "cover",
  },
  imageRemoveButton: {
    position: "absolute",
    top: "4px",
    right: "4px",
    minWidth: "20px",
    width: "20px",
    height: "20px",
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
}) => {
  const styles = useStyles();
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const safeImages = React.useMemo(() => z.array(ImageAttachmentSchema).catch([]).parse(images), [images]);

  useEffect(() => {
    // Refocus when value becomes empty (after sending)
    if (value === "") {
      inputRef.current?.focus();
    }
  }, [value]);

  const handleKeyPress = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      onSend();
    }
  };

  const handlePaste = async (e: React.ClipboardEvent) => {
    const items = e.clipboardData?.items;
    if (!items || !onImagesChange) return;

    const imageItems = Array.from(items).filter((item) => item.type.includes("image"));
    if (imageItems.length === 0) return;
    e.preventDefault();

    const nextImages = (await Promise.all(imageItems.map((item, index) => {
      const file = item.getAsFile();
      if (!file) return Promise.resolve(null);

      return new Promise<ImageAttachment | null>((resolve) => {
        const reader = new FileReader();
        reader.onload = (event) => {
          const dataUrl = typeof event.target?.result === "string" ? event.target.result : "";
          const parsed = ImageAttachmentSchema.safeParse({
            id: crypto.randomUUID(),
            dataUrl,
            name: `pasted-image-${Date.now()}-${index}.png`,
          });
          resolve(parsed.success ? parsed.data : null);
        };
        reader.onerror = () => resolve(null);
        reader.readAsDataURL(file);
      });
    }))).filter((image): image is ImageAttachment => image !== null);

    if (nextImages.length > 0) {
      onImagesChange([...safeImages, ...nextImages]);
    }
  };

  const handleRemoveImage = (id: string) => {
    if (onImagesChange) {
      onImagesChange(safeImages.filter((img) => img.id !== id));
    }
  };

  return (
    <div className={styles.inputContainer}>
      {safeImages.length > 0 && (
        <div className={styles.imagePreviewContainer}>
          {safeImages.map((image) => (
            <div key={image.id} className={styles.imagePreview}>
              <img src={image.dataUrl} alt="Preview" className={styles.imagePreviewImg} />
              <button
                type="button"
                className={styles.imageRemoveButton}
                onClick={() => handleRemoveImage(image.id)}
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
            onKeyDown={handleKeyPress}
            onPaste={handlePaste}
            placeholder="Ask OpenCode to work on the current document..."
            rows={3}
            disabled={disabled || isRunning}
          />
        </div>
        <Tooltip content={isRunning ? "Stop response" : "Send message"} relationship="label">
          <Button
            appearance={isRunning ? "secondary" : "primary"}
            icon={isRunning ? <Stop24Regular /> : <Send24Regular />}
            onClick={isRunning ? onStop : onSend}
            disabled={disabled || (isRunning ? !onStop : (!value.trim() && safeImages.length === 0))}
            aria-label={isRunning ? "Stop response" : "Send message"}
            className={`${styles.sendButton} ${isRunning ? styles.stopButton : ""}`.trim()}
          />
        </Tooltip>
      </div>
    </div>
  );
};
