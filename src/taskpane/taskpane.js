/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});




  export async function run() {
    /**
     * Insert your Outlook code here
     */
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
  