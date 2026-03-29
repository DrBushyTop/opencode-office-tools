/// <reference types="@types/office-js" />

import { z } from "zod";

const NavigationTargetSchema = z.enum([
  "edit-selection",
  "insert-timeline",
  "insert-estimate-table",
]);

type NavigationTarget = z.infer<typeof NavigationTargetSchema>;

function completeEvent(event?: Office.AddinCommands.Event) {
  event?.completed();
}

async function openTaskpane(target: NavigationTarget, event?: Office.AddinCommands.Event) {
  try {
    localStorage.setItem("navigationTarget", target);
    await Office.addin.showAsTaskpane();
  } finally {
    completeEvent(event);
  }
}

Office.onReady(() => {
  Office.actions.associate("editSelection", (event: Office.AddinCommands.Event) => {
    void openTaskpane("edit-selection", event);
  });
  Office.actions.associate("insertTimeline", (event: Office.AddinCommands.Event) => {
    void openTaskpane("insert-timeline", event);
  });
  Office.actions.associate("insertEstimateTable", (event: Office.AddinCommands.Event) => {
    void openTaskpane("insert-estimate-table", event);
  });
});
