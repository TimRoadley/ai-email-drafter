/* global Office console URL window */

import { checkAiStatus, DEFAULT_BASE_URL } from "../taskpane/ai-service";
import {
  getBaseUrl,
  getSelectedModel,
  getApiKey,
} from "../taskpane/settings-store";

Office.onReady(() => {
  Office.actions.associate("rewriteForClarity", rewriteForClarity);
});

interface DialogInitPayload {
  selectedText: string;
  baseUrl: string;
  apiKey?: string;
  models: string[];
  selectedModel: string;
}

interface DialogOutgoingMessage {
  type: "ready" | "accept" | "cancel";
  text?: string;
}

async function rewriteForClarity(
  event: Office.AddinCommands.Event,
): Promise<void> {
  const item = Office.context.mailbox.item;
  if (!item) {
    event.completed();
    return;
  }

  try {
    const selectedText = await new Promise<string>((resolve, reject) => {
      (item as Office.MessageCompose).getSelectedDataAsync(
        Office.CoercionType.Text,
        (
          result: Office.AsyncResult<{ data: string; sourceProperty: string }>,
        ) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(result.error.message));
          } else {
            resolve(result.value?.data ?? "");
          }
        },
      );
    });

    if (!selectedText?.trim()) {
      item.notificationMessages.addAsync("rewrite-info", {
        type: Office.MailboxEnums.ItemNotificationMessageType
          .InformationalMessage,
        message:
          "Select text in the email body first, then click Improve Writing.",
        icon: "Icon.80x80",
        persistent: false,
      });
      event.completed();
      return;
    }

    const baseUrl = getBaseUrl() ?? DEFAULT_BASE_URL;
    const apiKey = getApiKey();

    if (!baseUrl) {
      item.notificationMessages.addAsync("rewrite-error", {
        type: Office.MailboxEnums.ItemNotificationMessageType
          .InformationalMessage,
        message:
          "Configure the LLM Base URL in the Draft with AI add-in settings first.",
        icon: "Icon.80x80",
        persistent: false,
      });
      event.completed();
      return;
    }

    let model = getSelectedModel();
    let models: string[] = model ? [model] : [];

    const status = await checkAiStatus(baseUrl, apiKey);
    if (!status.online || status.models.length === 0) {
      item.notificationMessages.addAsync("rewrite-error", {
        type: Office.MailboxEnums.ItemNotificationMessageType
          .InformationalMessage,
        message:
          "AI is offline. Check the Base URL and API key in the add-in settings.",
        icon: "Icon.80x80",
        persistent: false,
      });
      event.completed();
      return;
    }
    models = status.models;
    if (!model || !models.includes(model)) {
      model = models[0];
    }

    const initPayload: DialogInitPayload = {
      selectedText,
      baseUrl,
      apiKey,
      models,
      selectedModel: model,
    };

    const dialogUrl = new URL(
      "improve-dialog.html",
      window.location.href,
    ).toString();

    await new Promise<void>((resolve) => {
      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 60, width: 40, displayInIframe: true },
        (asyncResult: Office.AsyncResult<Office.Dialog>) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            item.notificationMessages.addAsync("rewrite-error", {
              type: Office.MailboxEnums.ItemNotificationMessageType
                .InformationalMessage,
              message: `Can't open rewrite dialog: ${asyncResult.error.message}`,
              icon: "Icon.80x80",
              persistent: false,
            });
            resolve();
            return;
          }

          const dialog = asyncResult.value;

          const closeDialog = () => {
            try {
              dialog.close();
            } catch {
              // dialog may already be closed
            }
            resolve();
          };

          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (arg: any) => {
              let message: DialogOutgoingMessage;
              try {
                message = JSON.parse(arg.message);
              } catch {
                return;
              }

              if (message.type === "ready") {
                dialog.messageChild(JSON.stringify(initPayload));
                return;
              }

              if (message.type === "accept") {
                const rewritten = message.text ?? "";
                item.body.setSelectedDataAsync(
                  rewritten,
                  { coercionType: Office.CoercionType.Text },
                  (result: Office.AsyncResult<void>) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                      console.error(
                        "Failed to insert rewritten text:",
                        result.error.message,
                      );
                      item.notificationMessages.addAsync("rewrite-error", {
                        type: Office.MailboxEnums.ItemNotificationMessageType
                          .InformationalMessage,
                        message: `Can't rewrite: ${result.error.message}`,
                        icon: "Icon.80x80",
                        persistent: false,
                      });
                    }
                    closeDialog();
                  },
                );
                return;
              }

              if (message.type === "cancel") {
                closeDialog();
              }
            },
          );

          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            // Dialog was closed by the user (e.g. clicked the X).
            resolve();
          });
        },
      );
    });
  } catch (error: any) {
    console.error("Rewrite for clarity failed:", error);
    item.notificationMessages.addAsync("rewrite-error", {
      type: Office.MailboxEnums.ItemNotificationMessageType
        .InformationalMessage,
      message: `Can't rewrite: ${error.message ?? "Unknown error"}`,
      icon: "Icon.80x80",
      persistent: false,
    });
  } finally {
    event.completed();
  }
}
