import * as React from "react";
import { useState, useEffect, useCallback, useRef } from "react";
import {
  Button,
  Checkbox,
  Dropdown,
  Field,
  Input,
  Option,
  Spinner,
  Textarea,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  ArrowClockwiseRegular,
  DocumentEditRegular,
  SettingsRegular,
  EyeRegular,
  EyeOffRegular,
} from "@fluentui/react-icons";
import { insertText, getEmailBody } from "../taskpane";
import {
  checkAiStatus,
  generateReply,
  AiStatus,
  DEFAULT_BASE_URL,
} from "../ai-service";
import { DEFAULT_PROMPT_TEMPLATE } from "../prompt-template";
import {
  getBaseUrl,
  saveBaseUrl,
  getSignature,
  saveSignature,
  saveSelectedModel,
  getThinkingEnabled,
  saveThinkingEnabled,
  getCustomPromptTemplate,
  saveCustomPromptTemplate,
  getApiKey,
  saveApiKey,
} from "../settings-store";

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
  topRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  statusRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    flex: 1,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  statusDot: {
    width: "10px",
    height: "10px",
    borderRadius: "50%",
    display: "inline-block",
    flexShrink: 0,
  },
  statusDotOnline: {
    backgroundColor: tokens.colorPaletteGreenForeground1,
  },
  statusDotOffline: {
    backgroundColor: tokens.colorPaletteRedForeground1,
  },
  statusDotChecking: {
    backgroundColor: tokens.colorNeutralForeground3,
  },
  iconButton: {
    minWidth: "32px",
    padding: "0 8px",
  },
  settingsPanel: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  settingsRow: {
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
  draftButton: {
    marginTop: "8px",
  },
  spinner: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginTop: "4px",
  },
  errorText: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  successText: {
    color: tokens.colorPaletteGreenForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  infoText: {
    color: tokens.colorNeutralForeground2BrandHover,
    fontSize: tokens.fontSizeBase200,
  },
  hintText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  panelActions: {
    display: "flex",
    gap: "8px",
    marginTop: "8px",
  },
  inputRow: {
    display: "flex",
    gap: "4px",
    alignItems: "stretch",
  },
  inputRowInput: {
    flex: 1,
    minWidth: 0,
  },
  revealButton: {
    flexShrink: 0,
    minWidth: "32px",
    padding: "0 8px",
  },
  errorTextInline: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteRedForeground1,
    marginTop: "2px",
  },
});

function migrateUrl(saved: string | null): string {
  // Legacy relative/local URLs no longer work now that the add-in is hosted
  // remotely and talks to the LLM endpoint directly. Reset them to the new
  // corporate default so existing users upgrade transparently.
  if (
    !saved ||
    saved === "/lm-api/v1" ||
    saved === "http://localhost:1234/v1" ||
    saved === "http://localhost:1233/lm-api/v1" ||
    saved === "https://localhost:1233/lm-api/v1"
  ) {
    return DEFAULT_BASE_URL;
  }
  return saved;
}

function isValidBaseUrl(url: string): boolean {
  return /^https?:\/\/[a-zA-Z0-9.-]+(:[0-9]+)?(\/.*)?$/.test(url.trim());
}

const App: React.FC = () => {
  const styles = useStyles();
  const [baseUrl, setBaseUrl] = useState<string>(() =>
    migrateUrl(getBaseUrl()),
  );
  const [selectedModel, setSelectedModel] = useState<string>("");
  const [status, setStatus] = useState<AiStatus | null>(null);
  const [isCheckingStatus, setIsCheckingStatus] = useState<boolean>(false);
  const [showSettings, setShowSettings] = useState<boolean>(false);
  const [showPromptEditor, setShowPromptEditor] = useState<boolean>(false);
  const [specificInstructions, setSpecificInstructions] = useState<string>("");
  const [isDrafting, setIsDrafting] = useState<boolean>(false);
  const [signature, setSignature] = useState<string>(() => getSignature());
  const [message, setMessage] = useState<{
    type: "error" | "success" | "info";
    text: string;
  } | null>(null);
  const [thinkingEnabled, setThinkingEnabled] = useState<boolean>(() =>
    getThinkingEnabled(),
  );
  const [customPromptTemplate, setCustomPromptTemplate] = useState<
    string | null
  >(() => getCustomPromptTemplate());
  const [promptEditorDraft, setPromptEditorDraft] = useState<string>("");
  const [apiKey, setApiKey] = useState<string>(() => getApiKey());
  const [showApiKey, setShowApiKey] = useState<boolean>(false);
  const [baseUrlError, setBaseUrlError] = useState<string>("");
  const [viewportHeight, setViewportHeight] = useState<number>(() =>
    typeof window !== "undefined" ? window.innerHeight : 800,
  );
  const abortControllerRef = React.useRef<AbortController | null>(null);

  useEffect(() => {
    const onResize = () => setViewportHeight(window.innerHeight);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  useEffect(() => {
    if (baseUrl && isValidBaseUrl(baseUrl)) {
      setBaseUrlError("");
      saveBaseUrl(baseUrl);
    } else if (!baseUrl) {
      setBaseUrlError("");
      // Don't persist empty; the default is supplied at use-time by callers.
    } else {
      setBaseUrlError("Base URL must start with http:// or https://");
    }
  }, [baseUrl]);

  useEffect(() => {
    saveSignature(signature);
  }, [signature]);

  useEffect(() => {
    if (selectedModel) saveSelectedModel(selectedModel);
  }, [selectedModel]);

  useEffect(() => {
    saveThinkingEnabled(thinkingEnabled);
  }, [thinkingEnabled]);

  useEffect(() => {
    saveCustomPromptTemplate(customPromptTemplate);
  }, [customPromptTemplate]);

  useEffect(() => {
    saveApiKey(apiKey);
  }, [apiKey]);

  const checkStatus = useCallback(async (): Promise<AiStatus> => {
    setIsCheckingStatus(true);
    setMessage(null);
    try {
      const result = await checkAiStatus(baseUrl, apiKey);
      setStatus(result);
      return result;
    } catch (e: any) {
      const errorStatus = { online: false, models: [], error: e.message };
      setStatus(errorStatus);
      return errorStatus;
    } finally {
      setIsCheckingStatus(false);
    }
  }, [baseUrl, apiKey]);

  useEffect(() => {
    checkStatus();
  }, [checkStatus]);

  useEffect(() => {
    if (status?.models && status.models.length > 0) {
      setSelectedModel((prev) =>
        status.models.includes(prev) ? prev : status.models[0],
      );
    } else {
      setSelectedModel("");
    }
  }, [status?.models]);

  const openPromptEditor = useCallback(() => {
    setPromptEditorDraft(customPromptTemplate || DEFAULT_PROMPT_TEMPLATE);
    setShowSettings(false);
    setShowPromptEditor((v) => !v);
  }, [customPromptTemplate]);

  const isOnline = status?.online && status.models.length > 0;

  const statusText = isCheckingStatus
    ? "Checking..."
    : !status
      ? "Checking..."
      : isOnline
        ? "Ready"
        : status.error
          ? `AI Offline: ${status.error}`
          : "AI Offline";

  const handleDraft = async () => {
    if (!specificInstructions.trim()) {
      setMessage({ type: "error", text: "Please enter instructions." });
      return;
    }

    setIsDrafting(true);
    setMessage(null);

    const currentStatus = await checkStatus();

    if (!currentStatus.online || currentStatus.models.length === 0) {
      setIsDrafting(false);
      setMessage({
        type: "error",
        text: "AI is offline. Check that LM Studio is running.",
      });
      return;
    }

    const model = selectedModel || currentStatus.models[0];
    const controller = new AbortController();
    abortControllerRef.current = controller;

    try {
      const emailBody = await getEmailBody();
      const reply = await generateReply(
        baseUrl,
        model,
        specificInstructions,
        emailBody,
        thinkingEnabled,
        customPromptTemplate || undefined,
        controller.signal,
        apiKey,
      );
      const withSignature = signature.trim()
        ? `${reply}\n\n${signature}`
        : reply;
      await insertText(withSignature);
      setMessage({
        type: "success",
        text: "Draft reply inserted at cursor position.",
      });
    } catch (error: any) {
      if (
        error.name === "AbortError" ||
        error.message?.toLowerCase().includes("aborted")
      ) {
        setMessage({ type: "info", text: "Draft cancelled" });
      } else {
        setMessage({
          type: "error",
          text: error.message || "Failed to generate reply.",
        });
      }
    } finally {
      abortControllerRef.current = null;
      setIsDrafting(false);
    }
  };

  const handleCancel = () => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.topRow}>
        <div className={styles.statusRow}>
          <span
            className={`${styles.statusDot} ${
              isCheckingStatus
                ? styles.statusDotChecking
                : isOnline
                  ? styles.statusDotOnline
                  : styles.statusDotOffline
            }`}
          />
          <span>{statusText}</span>
          <Button
            className={styles.iconButton}
            icon={<ArrowClockwiseRegular />}
            appearance="subtle"
            size="small"
            onClick={checkStatus}
            disabled={isCheckingStatus}
            title="Refresh"
          />
        </div>
        <Button
          className={styles.iconButton}
          icon={<DocumentEditRegular />}
          appearance="subtle"
          size="small"
          onClick={openPromptEditor}
          title="Edit Prompt"
        />
        <Button
          className={styles.iconButton}
          icon={<SettingsRegular />}
          appearance="subtle"
          size="small"
          onClick={() => {
            setShowPromptEditor(false);
            setShowSettings((v) => !v);
          }}
          title="Settings"
        />
      </div>

      {showSettings ? (
        <div className={styles.settingsPanel}>
          <div className={styles.settingsRow}>
            <span className={styles.label}>LLM Base URL</span>
            <Input
              placeholder="https://api.example.com"
              value={baseUrl}
              onChange={(e: React.ChangeEvent<HTMLInputElement>) =>
                setBaseUrl(e.target.value)
              }
              onBlur={checkStatus}
            />
            {baseUrlError && (
              <span className={styles.errorTextInline}>{baseUrlError}</span>
            )}
          </div>

          <div className={styles.settingsRow}>
            <span className={styles.label}>
              API Key (optional — sent as Bearer token)
            </span>
            <div className={styles.inputRow}>
              <Input
                className={styles.inputRowInput}
                type={showApiKey ? "text" : "password"}
                placeholder="Leave blank if your endpoint doesn't require auth"
                value={apiKey}
                onChange={(e: React.ChangeEvent<HTMLInputElement>) =>
                  setApiKey(e.target.value)
                }
              />
              <Button
                className={styles.revealButton}
                appearance="subtle"
                size="small"
                onClick={() => setShowApiKey((v) => !v)}
                icon={showApiKey ? <EyeOffRegular /> : <EyeRegular />}
                title={showApiKey ? "Hide API Key" : "Show API Key"}
              />
            </div>
            <span className={styles.hintText}>
              Stored locally on this device. Leave blank if your LLM endpoint
              doesn't require authentication.
            </span>
          </div>

          <div className={styles.settingsRow}>
            <span className={styles.label}>Reply Signature</span>
            <Textarea
              placeholder={"Auto signature to include in every reply"}
              value={signature}
              onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) =>
                setSignature(e.target.value)
              }
              rows={3}
            />
          </div>
        </div>
      ) : showPromptEditor ? (
        <div className={styles.settingsPanel}>
          <div className={styles.settingsRow}>
            <span className={styles.label}>Custom Prompt Template</span>
            <span className={styles.hintText}>
              Use {"{{emailBody}}"} for the current email body and{" "}
              {"{{specificInstructions}}"} for the instructions you type above.
            </span>
            <Textarea
              value={promptEditorDraft}
              onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) =>
                setPromptEditorDraft(e.target.value)
              }
              style={{ height: Math.max(120, viewportHeight * 0.5) }}
            />
          </div>
          <div className={styles.panelActions}>
            <Button
              onClick={() => {
                setCustomPromptTemplate(promptEditorDraft);
                setShowPromptEditor(false);
              }}
            >
              Save
            </Button>
            <Button
              appearance="secondary"
              onClick={() => setShowPromptEditor(false)}
            >
              Cancel
            </Button>
            <Button
              appearance="subtle"
              onClick={() => setPromptEditorDraft(DEFAULT_PROMPT_TEMPLATE)}
            >
              Reset to Default
            </Button>
          </div>
        </div>
      ) : (
        <>
          <Field label="Specific Instructions">
            <Textarea
              placeholder="How should AI craft the message?"
              value={specificInstructions}
              onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) =>
                setSpecificInstructions(e.target.value)
              }
              style={{ height: Math.max(120, viewportHeight * 0.5) }}
            />
          </Field>

          {isOnline && status?.models && status.models.length > 1 && (
            <div className={styles.settingsRow}>
              <span className={styles.label}>Model</span>
              <Dropdown
                className={styles.dropdown}
                value={selectedModel}
                selectedOptions={[selectedModel]}
                onOptionSelect={(_, data) =>
                  setSelectedModel(data.optionValue ?? "")
                }
              >
                {status.models.map((m) => (
                  <Option key={m} value={m}>
                    {m}
                  </Option>
                ))}
              </Dropdown>
            </div>
          )}

          <Checkbox
            label="Thinking (better quality, slower)"
            checked={thinkingEnabled}
            onChange={(_ev, data) => setThinkingEnabled(!!data.checked)}
          />

          <Button
            className={styles.draftButton}
            appearance="primary"
            size="large"
            onClick={isDrafting ? handleCancel : handleDraft}
            disabled={!isDrafting && (!isOnline || isCheckingStatus)}
          >
            {isDrafting ? "Cancel" : "Draft"}
          </Button>

          {isDrafting && (
            <div className={styles.spinner}>
              <Spinner size="tiny" />
              <span>Drafting...</span>
            </div>
          )}

          {message && (
            <div
              className={
                message.type === "error"
                  ? styles.errorText
                  : message.type === "success"
                    ? styles.successText
                    : styles.infoText
              }
            >
              {message.text}
            </div>
          )}
        </>
      )}
    </div>
  );
};

export default App;
