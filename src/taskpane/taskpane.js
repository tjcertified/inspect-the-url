/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const SECURITY_LEVEL = Object.freeze({
  SAFE: "safe",
  SUSPICIOUS: "suspicious",
  MALICIOUS: "malicious",
  UNKNOWN: "unknown"
});

/* global document, Office */

Office.onReady((info) => {

  if (info.host === Office.HostType.Outlook) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("close").addEventListener("click", close);
    inspect();
  }
});

function close() {
  Office.context.ui.closeContainer();
}

function show(elementId) {
  document.getElementById(elementId).style.display = "block";
}

function hide(elementId) {
  document.getElementById(elementId).style.display = "none";
}

export async function inspect() {
  show("progress-container");  // show "anaylyzing"
  const item = Office.context.mailbox.item;

  // Check reply-to mismatch
  const senderEmail = item.sender.emailAddress;
  const replyToEmail = item.replyTo ? item.replyTo[0].emailAddress : item.sender.emailAddress;
  const addressMatch = senderEmail.toLowerCase() === replyToEmail.toLowerCase();
  document.getElementById("address-match").textContent = addressMatch ? "Yes" : "No";
  document.getElementById("address-match").classList.add(addressMatch ? "safe-text" : "unsafe-text");
  
  // Check sender domain and reputation

  const domain = item.sender.emailAddress.split("@")[1];
  document.getElementById("domain").textContent = domain;

  const [domain_result, auth_result, url_result, attachment_result] = await Promise.all([
    inspectDomain(domain),
    inspectAuthentication(),
    inspectURLs(),
    inspectAttachments()
  ]);

  // Compute overall trust rating and set text and color accordingly.
  document.getElementById("trust-rating").textContent = [domain_result, auth_result, url_result, attachment_result].every(result => result === SECURITY_LEVEL.SAFE) ? "High" :
    [domain_result, auth_result, url_result, attachment_result].some(result => result === SECURITY_LEVEL.MALICIOUS) ? "Low" :
    [domain_result, auth_result, url_result, attachment_result].some(result => result === SECURITY_LEVEL.SUSPICIOUS) ? "Medium" :
    "Unknown";
  document.getElementById("trust-rating").classList.add([domain_result, auth_result, url_result, attachment_result].every(result => result === SECURITY_LEVEL.SAFE) ? "safe" :
    [domain_result, auth_result, url_result, attachment_result].some(result => result === SECURITY_LEVEL.MALICIOUS) ? "unsafe" :
    [domain_result, auth_result, url_result, attachment_result].some(result => result === SECURITY_LEVEL.SUSPICIOUS) ? "caution" :
    "caution");

  console.log(`Domain result: ${domain_result}, Authentication result: ${auth_result}, URL result: ${url_result}, Attachment result: ${attachment_result}`);



  hide("progress-container");
  show("results");
}

// Check email authentication results (SPF, DKIM, DMARC) from email headers
async function inspectAuthentication() {
  const status = await new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAllInternetHeadersAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const headers = result.value;
        var auth = document.getElementById("auth");
        if (headers.includes("spf=pass") && headers.includes("dkim=pass") && headers.includes("dmarc=pass")) {
          auth.textContent = "Authenticated";
          auth.classList.add("safe-text");
          resolve(SECURITY_LEVEL.SAFE);
        } else if (headers.includes("spf=fail") || headers.includes("dkim=fail") || headers.includes("dmarc=fail")) {
          auth.textContent = "Not Authenticated";
          auth.classList.add("unsafe-text");
          resolve(SECURITY_LEVEL.UNSAFE);
        } else {
          auth.textContent = "Unable to verify";
          auth.classList.add("caution-text");
          resolve(SECURITY_LEVEL.SUSPICIOUS);
      }
    } else {
      console.error("Failed to get email headers:", result.error);
      resolve(SECURITY_LEVEL.UNKNOWN);
    }
  });
});

  return status;
}

// Check domain reputation using VirusTotal API
async function inspectDomain(domain) {
  const vt_url = `/scan/${domain}`;

  const response = await fetch(vt_url);
  const data = await response.json();
  var result_container = document.getElementById("reputation");
  if (data.error) {
    if (data.error.code === "RateLimitExceeded") {
      // handle gracefully somehow later
      result_container.textContent = "API rate limit exceeded. Please try again later.";
      return SECURITY_LEVEL.UNKNOWN;
    } else if (data.error.code === "NotFound") {
      result_container.textContent = "Unreported Domain";
      result_container.classList.add("caution-text");
      return SECURITY_LEVEL.SUSPICIOUS;
    } 
  } else {
      const is_malicious = data.data.attributes.reputation - data.data.attributes.last_analysis_stats.malicious;
      result_container.textContent = is_malicious < 0 ? "Reported Malicious" : "No Malicious Reports";
      result_container.classList.add(is_malicious < 0 ? "unsafe-text" : "safe-text");
      return is_malicious < 0 ? SECURITY_LEVEL.MALICIOUS : SECURITY_LEVEL.SAFE;
  }
}

// Wrap Office's getAsync in a Promise for easier async/await usage
function getEmailBodyAsync(coercionType) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(coercionType, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(new Error(asyncResult.error.message));
            } else {
                resolve(asyncResult.value);
            }
        });
    });
}

// Analyze URLs in the email body to find mismatches between displayed text and actual href
async function inspectURLs() {
  try {
    const asyncResult = await getEmailBodyAsync(Office.CoercionType.Html);
    // lifted and modified from https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/
    var htmlParser = new DOMParser().parseFromString(asyncResult, "text/html");
    var links = htmlParser.getElementsByTagName("a");
    var phishyLinkCount = 0;
    var normalLinkCount = 0;
    Array.from(links).forEach(
      function (v, i) {
          var regExp = new RegExp('/+$');
          var vInnerText = v.innerText.toLowerCase().trim().replace(regExp, "");
          var hrefText = v.href.toLowerCase().trim().replace(regExp, "");
          var linkIsPhishy = ((vInnerText.search("http") == 0) && vInnerText != hrefText);

          if (linkIsPhishy) {
              phishyLinkCount++;
          }
          else {
              normalLinkCount++;
          }
      }
    );
  } catch (error) {
    console.error("Error inspecting URLs:", error);
    document.getElementById("links").textContent = "Error analyzing links";
    document.getElementById("links").classList.add("caution-text");
    return SECURITY_LEVEL.UNKNOWN;
  }

  document.getElementById("links").textContent = `${phishyLinkCount} of ${normalLinkCount}`;
  if (phishyLinkCount > 0) {
    document.getElementById("links").classList.add("unsafe-text");
    // return suspicious instead of malicious to avoid potential false positives.
    return SECURITY_LEVEL.SUSPICIOUS;
  } else {
    document.getElementById("links").classList.add("safe-text");
    return SECURITY_LEVEL.SAFE;
  }
}

// Compute SHA-256 hash of a file using the Web Crypto API
async function getFileHash(file) {
  const reader = new FileReader();

  const arrayBuffer = await new Promise((resolve, reject) => {
    reader.onload = async () => {
      try {
        const arrayBuffer = reader.result;
        const hashBuffer = await crypto.subtle.digest("SHA-256", arrayBuffer);
        resolve(hashBuffer);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => {
      reject(reader.error);
    };
    reader.readAsArrayBuffer(file);
  });

  const hashArray = Array.from(new Uint8Array(arrayBuffer));
  const hashHex = hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
  return hashHex;
}

function getAttachmentContentAsync(attachmentId) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

// Analyze email attachments by hashing them and checking against VirusTotal's file API
async function inspectAttachments() {
  const attachments = Office.context.mailbox.item.attachments;
  let maliciousAttachmentCount = 0;

  if (attachments && attachments.length > 0) {
    let promise_status = await Promise.all(attachments.map(async (attachment) => {
      try {
        const contentResult = await getAttachmentContentAsync(attachment.id);
        const content = contentResult.content;
        const hash = await getFileHash(content);
        const vt_url = `/file/${hash}`;
        const response = await fetch(vt_url);
        const data = await response.json();

        if (data.error) {
          if (data.error.code === "RateLimitExceeded") {
            console.warn("API rate limit exceeded. Please try again later.");
            return SECURITY_LEVEL.UNKNOWN;
          } else if (data.error.code === "NotFound") {
            return SECURITY_LEVEL.SUSPICIOUS;
          }
        } else {
          const is_malicious = data.data.attributes.reputation - data.data.attributes.last_analysis_stats.malicious;
          if (is_malicious < 0) {
            return SECURITY_LEVEL.MALICIOUS;
          } else {
            return SECURITY_LEVEL.SUSPICIOUS;  // returning suspicious for unreported files to err on the side of caution, since many benign files may simply not be in VT's database.
          }
        }
      } catch (error) {
        document.getElementById("attachments").textContent = "Failed to inspect attachment.";
        document.getElementById("attachments").classList.add("caution-text");
        console.error("Error inspecting attachment:", error);
        return SECURITY_LEVEL.SUSPICIOUS;
      }
    }));
    if (promise_status) {
      promise_status.forEach(status => {
        if (status === SECURITY_LEVEL.MALICIOUS) {
          maliciousAttachmentCount++;
        }
      });
    }
    if (maliciousAttachmentCount > 0) {
      document.getElementById("attachments").textContent = `${maliciousAttachmentCount} of ${attachments.length}`;
      document.getElementById("attachments").classList.add("unsafe-text");
      return SECURITY_LEVEL.MALICIOUS;
    }
  } else {
    document.getElementById("attachments").textContent = "No attachments";
    document.getElementById("attachments").classList.add("safe-text");
    return SECURITY_LEVEL.SAFE;
  }

  return maliciousAttachmentCount > 0 ? SECURITY_LEVEL.MALICIOUS : SECURITY_LEVEL.SUSPICIOUS;
}
