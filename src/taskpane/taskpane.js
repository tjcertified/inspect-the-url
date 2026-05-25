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
  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Domain: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  let domain = item.sender.emailAddress.split("@")[1];
  //const vt_url = "https://www.virustotal.com/api/v3/domains/" + domain;
  const vt_url = `/scan/${domain}`;
  const options = {method: 'GET'};
  fetch(vt_url, options)
  .then(response => response.json())
  .then(data => {
    if (!data.data || !data.data.attributes || !data.data.attributes.last_analysis_stats) {
      insertAt.appendChild(document.createTextNode("Error fetching data from VirusTotal"));
    } else {
      var result = domain;
      let is_malicious = data.data.attributes.reputation - data.data.attributes.last_analysis_stats.malicious;
      is_malicious < 0 ? result += ": malicious" : result += ": not malicious";
      insertAt.appendChild(document.createTextNode(result));
    }
    insertAt.appendChild(document.createElement("br"));
  })
  .catch(err => {
    insertAt.appendChild(document.createTextNode("Error fetching data from VirusTotal" + err));
    insertAt.appendChild(document.createElement("br"));
  });
}
