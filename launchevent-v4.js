

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.

Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
console.log("new updated agenda");

/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/



function onNewAppointmentComposeHandler(event) {
  setMessage(event);
}


function setMessage(event) {
  var item = Office.context.mailbox.item;
  
  var htmlBody =
  "<p><b>Statement of Achievement:</b></p>" + 
  "<p><i>(What has to result from this meeting for the organizer to walk away elated with what was accomplished)</i></p>" + 
  "<p><b>Meeting Type (Informational or Decision):</b></p>" + 
  "<p><b>NOTE:</b> If you are remote and joining via TEAMs, turn on your video unless you are driving.</p>" + 
  "<p><b>NOTE:</b> If you are declining – email the organizer your answers to the questions posed prior to the date/time of the meeting – and/or forward the invite to a proxy to take your place.</p>" + 
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
  "</ul>";
  

  item.body.prependAsync(htmlBody, { coercionType: Office.CoercionType.Html }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Inserted content to body");
    } else {
      console.log("Error: " + asyncResult.error.message);
    }
    event.completed();
  });
}

