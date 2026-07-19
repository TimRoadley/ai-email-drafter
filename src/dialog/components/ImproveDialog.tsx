import * as React from "react";
import { useCallback, useEffect, useState } from "react";
import {
  Button,
  Dropdown,
  Option,
  Spinner,
  Textarea,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { rewriteTextForClarity } from "../../taskpane/ai-service";

/* global Office */

interface InitPayload {
  selectedText: string;
  baseUrl: string;
  apiKey?: string;
  models: string[];
  selectedModel: string;
}

interface OutgoingMessage {
  type: "ready" | "accept" | "cancel";
  text?: string;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    padding: "16px",
    gap: "12px",
    minHeight: "100vh",
    boxSizing: "border-box",
    fontFamily: tokens.fontFamilyBase,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
  },
  dropdown: {
    width: "100%",
  },
  textarea: {
    width: "100%",
  },
  actions: {
    display: "flex",
    gap: "8px",
    marginTop: "8px",
  },
  spinnerRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  errorText: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase200,
  },
});

function sendToParent(message: OutgoingMessage): void {
  Office.context.ui.messageParent(JSON.stringify(message));
}

const ImproveDialog: React.FC = () => {
  const styles = useStyles();
  const [init, setInit] = useState<InitPayload | null>(null);
  const [selectedModel, setSelectedModel] = useState<string>("");
  const [suggestion, setSuggestion] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");

  // Handshake: tell the parent we're ready, then wait for it to send back
  // the extracted text + settings via Dialog.messageChild.
  useEffect(() => {
    Office.onReady(() => {
      Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        (arg: any) => {
          try {
            const payload: InitPayload = JSON.parse(arg.message);
            setInit(payload);
            setSelectedModel(payload.selectedModel || payload.models[0] || "");
          } catch (e) {
            setError("Failed to read data from the add-in.");
          }
        },
        () => {
          sendToParent({ type: "ready" });
        },
      );
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const runRewrite = useCallback(
    async (payload: InitPayload, model: string) => {
      setIsLoading(true);
      setError("");
      try {
        const rewritten = await rewriteTextForClarity(
          payload.baseUrl,
          model,
          payload.selectedText,
          payload.apiKey,
        );
        setSuggestion(rewritten);
      } catch (e: any) {
        setError(e.message || "Failed to rewrite text.");
      } finally {
        setIsLoading(false);
      }
    },
    [],
  );

  useEffect(() => {
    if (init && selectedModel) {
      runRewrite(init, selectedModel);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [init]);

  const handleAccept = () => {
    sendToParent({ type: "accept", text: suggestion });
  };

  const handleCancel = () => {
    sendToParent({ type: "cancel" });
  };

  const handleRegenerate = () => {
    if (init && selectedModel) {
      runRewrite(init, selectedModel);
    }
  };

  if (!init) {
    return (
      <div className={styles.root}>
        <div className={styles.spinnerRow}>
          <Spinner size="tiny" />
          <span>Loading selected text...</span>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <div className={styles.section}>
        <span className={styles.label}>Selected text</span>
        <Textarea
          className={styles.textarea}
          value={init.selectedText}
          readOnly
          rows={4}
        />
      </div>

      {init.models.length > 1 && (
        <div className={styles.section}>
          <span className={styles.label}>Model</span>
          <Dropdown
            className={styles.dropdown}
            value={selectedModel}
            selectedOptions={[selectedModel]}
            onOptionSelect={(_, data) => {
              const nextModel = data.optionValue ?? "";
              setSelectedModel(nextModel);
              if (init && nextModel) {
                runRewrite(init, nextModel);
              }
            }}
          >
            {init.models.map((m) => (
              <Option key={m} value={m}>
                {m}
              </Option>
            ))}
          </Dropdown>
        </div>
      )}

      <div className={styles.section}>
        <span className={styles.label}>Suggested rewrite</span>
        <Textarea
          className={styles.textarea}
          value={suggestion}
          onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) =>
            setSuggestion(e.target.value)
          }
          rows={6}
          disabled={isLoading}
        />
      </div>

      {isLoading && (
        <div className={styles.spinnerRow}>
          <Spinner size="tiny" />
          <span>Rewriting...</span>
        </div>
      )}

      {error && <span className={styles.errorText}>{error}</span>}

      <div className={styles.actions}>
        <Button
          appearance="primary"
          onClick={handleAccept}
          disabled={isLoading || !suggestion.trim()}
        >
          Accept
        </Button>
        <Button appearance="secondary" onClick={handleCancel}>
          Cancel
        </Button>
        <Button
          appearance="subtle"
          onClick={handleRegenerate}
          disabled={isLoading}
        >
          Regenerate
        </Button>
      </div>
    </div>
  );
};

export default ImproveDialog;
