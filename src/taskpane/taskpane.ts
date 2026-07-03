/* global Office console */

function textToHtml(text: string): string {
  return (
    text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\r\n|\r/g, "\n")
      .replace(/\n/g, "<br>") + "<br><br>"
  );
}

export async function insertText(text: string) {
  return new Promise<void>((resolve, reject) => {
    Office.context.mailbox.item!.body.setSelectedDataAsync(
      textToHtml(text),
      { coercionType: Office.CoercionType.Html },
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Error inserting text:", asyncResult.error.message);
          reject(asyncResult.error.message);
        } else {
          resolve();
        }
      },
    );
  });
}

export async function getEmailBody(): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    Office.context.mailbox.item!.body.getAsync(
      Office.CoercionType.Text,
      { asyncContext: "Read email body" },
      (asyncResult: Office.AsyncResult<string>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Error reading body:", asyncResult.error.message);
          reject(asyncResult.error.message);
        } else {
          resolve(asyncResult.value);
        }
      },
    );
  });
}
