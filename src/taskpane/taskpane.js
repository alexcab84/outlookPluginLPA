let dialog;
let alert;
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    run();
  }
});

function processMessage(arg) {
  dialog.close();
}

function processMessageAlert(arg) {
  alert.close();
}

import "core-js";
import templateJson from "./config.json";
const ip_client = templateJson.ip_client;
const port_client = templateJson.port_client;

export async function run() {
  const item = Office.context.mailbox.item;
  const toAddresses = item.to ? item.to.map((t) => t.emailAddress).join(";") : "";
  const subject = item.subject ? item.subject.replace(/(["'<>'#])/g, "") : "N/A";

  const insertAt = document.getElementById("item-subject");
  insertAt.appendChild(createLabel("From: ", item.from.emailAddress));
  insertAt.appendChild(createLabel("To: ", toAddresses));
  insertAt.appendChild(createLabel("Subject: ", subject));

  let bodyObject = {};
  let textBody = {};

  Office.context.mailbox.item.body.getAsync("html", (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      bodyObject.body = result.value;

      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result2) => {
        if (result2.status === Office.AsyncResultStatus.Succeeded) {
          textBody.text = result2.value.replace(/(["'<>'#])/g, "");
          processAttachments(toAddresses, item.from.emailAddress, bodyObject.body, subject, textBody.text);
        }
      });
    }
  });
}

function createLabel(labelText, value) {
  const wrapper = document.createElement("div");
  wrapper.innerHTML = `<b>${labelText}</b><br>${value}<br>`;
  return wrapper;
}

async function processAttachments(to, from, body, subject, textBody) {
  const attachments = Office.context.mailbox.item.attachments || [];
  const attachmentContents = [];
  let attachNames = "";

  if (attachments.length === 0) {
    openDialog(to, from, ["N/A"], subject, "", textBody, body, attachmentContents);
    return;
  }

  let pendingCalls = attachments.length;

  attachments.forEach((attachment) => {
    if (attachment.attachmentType !== Office.MailboxEnums.AttachmentType.File) {
      pendingCalls--;
      return;
    }

    Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
      pendingCalls--;
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        attachmentContents.push({
          id: attachment.id,
          name: attachment.name,
          size: attachment.size,
          format: result.value.format,
          content: result.value.content,
        });
        attachNames += attachNames ? ";" + attachment.name : attachment.name;
      }
      if (pendingCalls === 0) {
        openDialog(
          to,
          from,
          Office.context.mailbox.item.cc.map((c) => c.emailAddress) || ["N/A"],
          subject,
          attachNames,
          textBody,
          body,
          attachmentContents
        );
      }
    });
  });
}

function constructDialogUrl(to, from, ccEmails, subject, attachNames, textBody) {
  const createdDate = new Date(Office.context.mailbox.item.dateTimeCreated).toISOString();
  return `https://${ip_client}:${port_client}/dialog.html?to=${to}&from=${from}&cc=${ccEmails.join(";")}&subject=${subject}&textBody=${textBody}&sentOn=${createdDate}&attachNames=${attachNames}`;
}

function openDialog(to, from, ccEmails, subject, attachNames, textBody, body, attachmentContents) {
  const url = constructDialogUrl(to, from, ccEmails, subject, attachNames, textBody);
  Office.context.ui.displayDialogAsync(url, { height: 75, width: 75, displayInIframe: true }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      dialog = asyncResult.value;
      console.log("Diálogo abierto correctamente:", dialog);

      if (dialog) {
        console.log("Agregando event handler...");
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (event) => {
          console.log("Evento recibido:", event.message);

          if (event.message === "ready") {
            console.log("Enviando datos al diálogo...");
            dialog.messageChild(JSON.stringify({ htmlBody: body, content: attachmentContents, textBody }));
          } else if (event.message === "closeDialog") {
            console.log("Cerrando diálogo...");
            dialog.close();
            Office.context.ui.closeContainer();
          } else if (event.message !== "ready" && event.message !== "closeDialog") {
            var respuesta = null;
            var jsonResponse = null;
            try {
              var respuesta = event.message;
              var jsonResponse = JSON.parse(respuesta);
              let statusValue = jsonResponse['status'];
              let dataValue = jsonResponse['data'];

              console.log('Status:', statusValue);
              console.log('Data:', dataValue);
              console.log(jsonResponse);
              //jsonResponse = jsonResponse.replace(/'/g, "\\'");
              //jsonData = JSON.parse(jsonResponse);
              if (statusValue === "error") {
                console.log("Respuesta erronea del diálogo:", dataValue);
                if (!alert || alert.closed) {
                  const windowWidth = window.innerWidth * 2;
                  const windowHeight = window.innerHeight / 4;
                  const screenWidth = window.screen.width;
                  const screenHeight = window.screen.height;

                  const left = (screenWidth - windowWidth) / 2;
                  const top = (screenHeight - windowHeight) / 2;

                  const windowFeatures = `width=${windowWidth},height=${windowHeight},top=${top},left=${left},resizable=no`;
                  alert = window.open("alert.html?param=" + dataValue, "alert", windowFeatures);
                } else {
                  alert.focus();
                }
              }
            } catch (e) {
              console.log(e);
            }
          }
        });
      } else {
        console.error("El diálogo no se inicializó correctamente.");
      }
    } else {
      console.error("Error al abrir el diálogo:", asyncResult.error);
    }
  });
}
