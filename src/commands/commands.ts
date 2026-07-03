/* global Office console */

import {
  checkAiStatus,
  rewriteTextForClarity,
  DEFAULT_BASE_URL,
} from "../taskpane/ai-service";
import {
  getBaseUrl,
  getSelectedModel,
  getApiKey,
} from "../taskpane/settings-store";

Office.onReady(() => {
  Office.actions.associate("rewriteForClarity", rewriteForClarity);
});

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
          "Select text in the email body first, then click Rewrite for Clarity.",
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

    if (!model) {
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
      model = status.models[0];
    }

    const rewritten = await rewriteTextForClarity(
      baseUrl,
      model,
      selectedText,
      apiKey,
    );

    await new Promise<void>((resolve, reject) => {
      item.body.setSelectedDataAsync(
        rewritten,
        { coercionType: Office.CoercionType.Text },
        (result: Office.AsyncResult<void>) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(result.error.message));
          } else {
            resolve();
          }
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
