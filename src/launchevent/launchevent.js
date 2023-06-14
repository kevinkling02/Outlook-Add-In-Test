/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});


function onNewAppointmentComposeHandler(event) {


    // Get a reference to the current message
    const item = Office.context.mailbox.item;

    // Write message property value to the task pane
      // Get a reference to the current item
  
      // Construct the HTML content
      const htmlContent = `
    <div style="background-color: #F0F0F0; padding: 10px;">
      <table style="width: 100%; border-collapse: collapse;">
        <thead>
          <tr>
            <th style="border: 1px solid #000; padding: 8px;">Topic</th>
            <th style="border: 1px solid #000; padding: 8px;">Goal</th>
            <th style="border: 1px solid #000; padding: 8px;">Additional Documents</th>
            <th style="border: 1px solid #000; padding: 8px;">Responsible for Presentation</th>
            <th style="border: 1px solid #000; padding: 8px;">Time Slot</th>
          </tr>
        </thead>
        <tbody id="agendaTableBody">
          <tr>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
          </tr>
          <tr>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
          </tr>
          <tr>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
          </tr>
          <tr>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
            <td style="border: 1px solid #000; padding: 8px;" contenteditable="true"></td>
          </tr>
        </tbody>
      </table>
    </div>
      `;
   
      // Set the HTML content as the body
      item.body.prependAsync(htmlContent, {coercionType: Office.CoercionType.Html,
        asyncContext: {var3: 1, var4: 2} }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('HTML inserted successfully');
        } else {
          console.error('Failed to insert HTML');
        }
      });

      // Notification
      const message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Agenda wurde automatisch hinzugefügt!",
        icon: "Icon.80x80",
        persistent: true,
      };
    
      // Show a notification message
      Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

}



// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
}

