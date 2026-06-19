// Maps the manifest FunctionName to this handler so Outlook can invoke it.
// This must be called at the top level before any event fires.
Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);

// ---------------------------------------------------------------------------
// Entry point — fires via OnNewAppointmentOrganizer event-based activation.
// Delegates immediately to the idempotency-guarded insertion function.
// ---------------------------------------------------------------------------
function onNewAppointmentComposeHandler(event) {
  insertTemplateIfNotAlreadyInserted(event);
}

// ---------------------------------------------------------------------------
// Inserts the meeting template into the appointment body exactly once.
//
// Uses a custom item property ("templateInserted") as a persistent flag so
// that reopening a draft, forwarding, or editing a sent meeting does NOT
// re-insert the template and create duplicates.
//
// Custom properties are stored with the Exchange item itself, so they
// survive across sessions and devices.
//
// event.completed() MUST be called in every code path — if it is omitted
// the add-in hangs and Outlook may mark it as unresponsive.
// ---------------------------------------------------------------------------
function insertTemplateIfNotAlreadyInserted(event) {
  var item = Office.context.mailbox.item;

  item.loadCustomPropertiesAsync(function (loadResult) {

    // If the custom properties bag itself failed to load, fail open:
    // attempt insertion anyway so the user still gets the template.
    // Worst case on a repeated open is a duplicate, which is preferable
    // to silently swallowing the error and hanging.
    if (loadResult.status === Office.AsyncResultStatus.Failed) {
      console.error("loadCustomPropertiesAsync failed: " + loadResult.error.message +
        " — proceeding with insertion as a fallback.");
      prependTemplate(event, null);
      return;
    }

    var customProps = loadResult.value;

    // Check if this item has already had the template inserted.
    // get() returns undefined for properties that have never been set.
    var alreadyInserted = customProps.get("templateInserted") === true;

    if (alreadyInserted) {
      // Template was inserted in a previous session (e.g. this is a reopened
      // draft). Skip insertion to prevent duplicates.
      console.log("templateInserted flag is set — skipping insertion.");
      event.completed();
      return;
    }

    // First time this item has been opened in compose — insert the template.
    prependTemplate(event, customProps);
  });
}

// ---------------------------------------------------------------------------
// Calls prependAsync with the HTML meeting template.
//
// On success, persists the "templateInserted" flag to the item's custom
// properties so future opens are no-ops. event.completed() is called inside
// the saveAsync callback to guarantee it fires after the save round-trip.
//
// customProps may be null when called as a fallback from a failed
// loadCustomPropertiesAsync — in that case we skip the save step.
// ---------------------------------------------------------------------------
function prependTemplate(event, customProps) {
  var item = Office.context.mailbox.item;

  var htmlBody =
    "<p><b>Statement of Achievement:</b></p>" +
    "<p><i>(What has to result from this meeting for the organizer to walk away elated with what was accomplished)</i></p>" +
    "<p><b>Meeting Type (Informational or Decision):</b></p>" +
    "<p><span style='color:#ff0000;'>NOTE: If you are remote and joining via TEAMs, turn on your video unless you are driving.</span></p>" +
    "<p><span style='color:#ff0000;'>NOTE: If you are declining – email the organizer your answers to the questions posed prior to the date/time of the meeting – and/or forward the invite to a proxy to take your place.</span></p>" +
    "<p><b>Agenda:</b></p>" +
    "<ul>" +
    "  <li>(5 minutes) Meeting Opening" +
    "    <ul>" +
    "      <li>Identify/confirm a facilitator and note taker</li>" +
    "      <li>Safety moment</li>" +
    "      <li>Mindful moment</li>" +
    "    </ul>" +
    "  </li>" +
    "  <li>Review Statement of Achievement and agenda</li>" +
    "  <li>(5–10 minutes) Review of Previous Meeting Clean Agreements</li>" +
    "  <li>Main Discussion" +
    "    <ul>" +
    "      <li>(__ minutes) Question #1</li>" +
    "      <li>(__ minutes) Question #2</li>" +
    "      <li>(__ minutes) Question #3</li>" +
    "      <li>(__ minutes) Question #4</li>" +
    "      <li>(5 minutes) What Clean Agreements do we have (including next mtg date and time)?</li>" +
    "      <li>(1 minute) Did this meeting fulfill the statement of achievement?</li>" +
    "    </ul>" +
    "  </li>" +
    "</ul>" +
    "<p>Is there a way this meeting or meeting topic could benefit from the use of AI?</p>";

  item.body.prependAsync(htmlBody, { coercionType: Office.CoercionType.Html }, function (prependResult) {

    if (prependResult.status === Office.AsyncResultStatus.Failed) {
      // Insertion failed (e.g. body not yet ready). Log the error and
      // complete the event so Outlook does not hang.
      console.error("prependAsync failed: " + prependResult.error.message);
      event.completed();
      return;
    }

    console.log("Meeting template inserted successfully.");

    // If we have no custom properties bag (fallback path), skip the save
    // and complete the event immediately.
    if (!customProps) {
      event.completed();
      return;
    }

    // Persist the flag so reopening this item in the future skips insertion.
    customProps.set("templateInserted", true);

    customProps.saveAsync(function (saveResult) {
      if (saveResult.status === Office.AsyncResultStatus.Failed) {
        // The flag did not persist. The template is already in the body,
        // so the only consequence is a possible duplicate on next open.
        // Still complete the event — do not hang.
        console.error("saveAsync failed: " + saveResult.error.message +
          " — templateInserted flag was not persisted.");
      } else {
        console.log("templateInserted flag saved to item custom properties.");
      }

      // Complete the event after the save attempt regardless of outcome.
      event.completed();
    });
  });
}
