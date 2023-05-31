/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/


function onNewAppointmentComposeHandler(event) {


    // Get a reference to the current message
    const item = Office.context.mailbox.item;

    // Write message property value to the task pane
      // Get a reference to the current item
  
      // Construct the HTML content
      const htmlContent = `
        <div style="background-color: #F0F0F0; padding: 10px;">
          <h1>Agenda</h1>
          <ul>
            <li>Agenda Item 1</li>
            <li>Agenda Item 2</li>
            <li>Agenda Item 3</li>
            <li>Agenda Item 4</li>
          </ul>
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
}



// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
}