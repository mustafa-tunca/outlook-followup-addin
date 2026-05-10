/* global Office, console */

/**
 * Ribbon action handler entry point.
 *
 * The unified manifest routes the "Create Follow-up" button to the
 * TaskPaneRuntime (openPage action), so this function file is the
 * fallback/programmatic path — useful when add-in code needs to open
 * the pane imperatively (e.g., from a notification action or a future
 * contextual trigger).
 */
Office.onReady(() => {
  // Register named functions that the manifest's executeFunction actions can call.
  // The function name must match the "id" in the manifest action.
  if (Office.actions) {
    Office.actions.associate("CreateFollowUpCommand", createFollowUpCommand);
  }
});

/**
 * Imperatively shows the pre-flight taskpane.
 * Called when the manifest routes the ribbon button to executeFunction
 * instead of openPage (e.g., in Classic Outlook with add-in commands v1).
 */
function createFollowUpCommand(event: Office.AddinCommands.Event): void {
  try {
    // showAsTaskpane() is available in Mailbox 1.9 / New Outlook.
    // In older hosts it falls back gracefully — the openPage runtime
    // in the manifest already handles the common case.
    if (Office.addin && typeof Office.addin.showAsTaskpane === "function") {
      Office.addin.showAsTaskpane().then(() => {
        event.completed();
      }).catch((err: unknown) => {
        console.error("showAsTaskpane failed:", err);
        event.completed();
      });
    } else {
      // Host doesn't support programmatic taskpane show — signal done.
      // The openPage runtime binding in the manifest handles this host.
      event.completed();
    }
  } catch (err) {
    console.error("createFollowUpCommand error:", err);
    event.completed();
  }
}
