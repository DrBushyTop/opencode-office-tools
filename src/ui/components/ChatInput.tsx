import * as React from "react";
import { useRef, useEffect } from "react";
import { Textarea, Button, Tooltip, makeStyles } from "@fluentui/react-components";
import { Send24Regular, Dismiss24Regular, Stop24Regular } from "@fluentui/react-icons";

export interface ImageAttachment {
  id: string;
  dataUrl: string;
  name: string;
}

interface ChatInputProps {
  value: string;
  onChange: (value: string) => void;
  onSend: () => void;
  onStop?: () => void;
  onSent?: () => void;
  disabled?: boolean;
  isRunning?: boolean;
  images?: ImageAttachment[];
  onImagesChange?: (images: ImageAttachment[]) => void;
}

const useStyles = makeStyles({
  inputContainer: {
    margin: "0 12px 12px",
    padding: "0",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    borderRadius: "12px",
    backgroundColor: "var(--oc-bg-strong)",
    border: "1px solid var(--oc-border)",
    boxShadow: "var(--oc-shadow)",
    overflow: "hidden",
  },
  body: {
    display: "flex",
    alignItems: "flex-end",
    gap: "8px",
    padding: "8px",
  },
  input: {
    flex: 1,
    width: "100%",
    maxWidth: "100%",
    boxSizing: "border-box",
    padding: "8px 10px",
    borderRadius: "0",
    border: "none !important",
    backgroundColor: "transparent !important",
    outline: "none !important",
    boxShadow: "none !important",
    color: "var(--oc-text)",
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
    borderRadius: "10px",
    background: "var(--oc-bg)",
    border: "1px solid var(--oc-border)",
  },
  sendButton: {
    width: "36px",
    height: "36px",
    minWidth: "36px",
    padding: "0",
    alignSelf: "flex-end",
    backgroundColor: "var(--oc-accent)",
    border: "none",
    borderRadius: "9px",
    color: "white",
    ":hover": {
      backgroundColor: "var(--oc-accent-strong)",
    },
  },
  imagePreviewContainer: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
    padding: "8px 8px 0",
  },
  imagePreview: {
    position: "relative",
    width: "80px",
    height: "80px",
    borderRadius: "8px",
    overflow: "hidden",
    border: "1px solid var(--oc-border)",
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
  isRunning = false,
  images = [],
  onImagesChange,
}) => {
  const styles = useStyles();
  const inputRef = useRef<HTMLTextAreaElement>(null);

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

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.type.indexOf('image') !== -1) {
        e.preventDefault();
        const file = item.getAsFile();
        if (file) {
          const reader = new FileReader();
          reader.onload = (event) => {
            const dataUrl = event.target?.result as string;
            const newImage: ImageAttachment = {
              id: crypto.randomUUID(),
              dataUrl,
              name: `pasted-image-${Date.now()}.png`,
            };
            onImagesChange([...images, newImage]);
          };
          reader.readAsDataURL(file);
        }
      }
    }
  };

  const handleRemoveImage = (id: string) => {
    if (onImagesChange) {
      onImagesChange(images.filter(img => img.id !== id));
    }
  };

  return (
    <div className={styles.inputContainer}>
      {images.length > 0 && (
        <div className={styles.imagePreviewContainer}>
          {images.map(image => (
            <div key={image.id} className={styles.imagePreview}>
              <img src={image.dataUrl} alt="Preview" className={styles.imagePreviewImg} />
              <button
                className={styles.imageRemoveButton}
                onClick={() => handleRemoveImage(image.id)}
                title="Remove image"
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
            placeholder="Type a message... (paste images with Ctrl+V)"
            rows={2}
          />
        </div>
        <Tooltip content={isRunning ? "Stop response" : "Send message"} relationship="label">
          <Button
            appearance={isRunning ? "secondary" : "primary"}
            icon={isRunning ? <Stop24Regular /> : <Send24Regular />}
            onClick={isRunning ? onStop : onSend}
            disabled={isRunning ? !onStop : (!value.trim() && images.length === 0)}
            aria-label={isRunning ? "Stop response" : "Send message"}
            className={styles.sendButton}
          />
        </Tooltip>
      </div>
    </div>
  );
};
