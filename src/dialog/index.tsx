import * as React from "react";
import { createRoot } from "react-dom/client";
import ImproveDialog from "./components/ImproveDialog";
import { FluentProvider, webDarkTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render dialog UI after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webDarkTheme}>
      <ImproveDialog />
    </FluentProvider>,
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/ImproveDialog", () => {
    const NextImproveDialog = require("./components/ImproveDialog").default;
    root?.render(NextImproveDialog);
  });
}
