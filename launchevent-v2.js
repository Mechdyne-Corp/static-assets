

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.

Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
console.log("first line");

/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/


const htmlBody = "<p> <b> Statement of Achievement </b> <br/> " 
                  + " <b> Meeting Type (informational or decision): </b> <br/> "
                  +" <b> Agenda: </b> <br/> <b>Facilitator: </b> <br/> <b> Note Taker: </b> </p>";

function onNewAppointmentComposeHandler(event) {
  setMessage(event);
}

function setMessage(event) {
  const item = Office.context.mailbox.item;
 // item.body.setAsync(htmlBody, {coercionType : Office.CoercionType.Html},
 item.body.prependAsync(htmlBody, {coercionType : Office.CoercionType.Html},
    (asyncResult) => {
      /*
      if(asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        console.log("inserted html to body");
      } else {
        console.log("error" +asyncResult.error.message);
      }
        */
      event.completed();
    }
);

}
