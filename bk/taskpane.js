let dialog;
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    run();
  }
});

function processMessage(arg) {
  dialog.close();
}

import "core-js";

export async function run() {
  let attachArrayId = "";
  let attachArrayName = "";
  let attachArrayContent = "";
  const item = Office.context.mailbox.item;
  const toLength = item.to ? item.to.length : 0;
  const attachmentsLength = item.attachments ? item.attachments.length : 0;
  let tos = "";

  // Concatenar todos los destinatarios
  for (let i = 0; i < toLength; i++) {
    tos += item.to[i].emailAddress + ";";
  }
  tos = tos.slice(0, -1); // Eliminar el último punto y coma
  const subject = Office.context.mailbox.item.subject;
  const insertAt = document.getElementById("item-subject");
  insertAt.appendChild(createLabel("From: ", item.from.emailAddress));
  insertAt.appendChild(createLabel("To: ", tos));
  insertAt.appendChild(createLabel("Subject: ", item.subject));

  let bodyObject = {};
  let textBody = {};
  Office.context.mailbox.item.body.getAsync("html", function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      bodyObject.body = result.value;

      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (result2) {
        if (result2.status === Office.AsyncResultStatus.Succeeded) {
          const body = result2.status;
          if (body && body.trim() !== "") {
            textBody.text = result2.value;
          }
          for (let i = 0; i < attachmentsLength; i++) {
            const nombre = item.attachments[i].name;
            attachArrayId += item.attachments[i].id + (i === attachmentsLength - 1 ? "" : ";");
            attachArrayName += item.attachments[i].name + (i === attachmentsLength - 1 ? "" : ";");

            //insertAt.appendChild(createLabel("Attachments: ", nombre));

            item.getAttachmentContentAsync(item.attachments[i].id, function (result) {
              if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
                attachArrayContent += result.value.content.toString() + (i === attachmentsLength - 1 ? "" : ";");
                //insertAt.appendChild(createLabel("Attachments_Content: ", result.value.content.toString()));
              }
            });
          }

          getAllAttachmentContents(tos, item.from.emailAddress, bodyObject.body, subject, textBody.text);
        }
      });

      // Procesar los archivos adjuntos
    }
  });
}

function createLabel(labelText, value) {
  const label = document.createElement("b");
  label.appendChild(document.createTextNode(labelText));
  const lineBreak = document.createElement("br");
  const textNode = document.createTextNode(value);
  const wrapper = document.createElement("div");

  wrapper.appendChild(label);
  wrapper.appendChild(lineBreak);
  wrapper.appendChild(textNode);
  wrapper.appendChild(lineBreak);
  return wrapper;
}

async function getAllAttachmentContents(to, from, body, subject, textBody) {
  const attachments = Office.context.mailbox.item.attachments;
  const attachmentContents = [];
  if (attachments.length === 0) {
    const ccRecipients = Office.context.mailbox.item.cc;
    const ccEmails = ccRecipients ? ccRecipients.map((recipient) => recipient.emailAddress) : ["N/A"];
    const characterCount = 1;
    if (characterCount <= 1) {
      const ccEmails2 = "N/A";
      openDialogAndSendData(constructDialogUrl2(to, from, ccEmails2, subject, textBody), attachmentContents, body);
    } else {
      openDialogAndSendData(constructDialogUrl(to, from, ccEmails, subject, textBody), attachmentContents, body);
    }
  }

  let pendingCalls = attachments.length;
  let attachNames = "";

  attachments.forEach((attachment) => {
    if (
      attachment.attachmentType === Office.MailboxEnums.AttachmentType.File ||
      attachment.attachmentType === Office.MailboxEnums.AttachmentType.Item
    ) {
      Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
        pendingCalls--;

        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const content = result.value;
          attachNames += ";" + attachment.name;
          attachmentContents.push({
            id: attachment.id,
            name: attachment.name,
            size: attachment.size,
            format: content.format,
            content: content.content,
          });
        } else {
          console.error(`Failed to get content for attachment ${attachment.name}:`, result.error);
        }

        // Cuando todas las llamadas asíncronas se completan
        if (pendingCalls === 0) {
          const ccRecipients = Office.context.mailbox.item.cc;
          const ccEmails =
            ccRecipients && ccRecipients.length > 0 ? ccRecipients.map((recipient) => recipient.emailAddress) : ["N/A"];
          openDialogAndSendData(
            constructDialogUrl(to, from, ccEmails, subject, attachNames, textBody),
            attachmentContents,
            body
          );
        }
      });
    } else {
      pendingCalls--;
    }
  });
}

function constructDialogUrl(to, from, ccEmails, subject, attachNames = "", textBody) {
  /*const createdDate = Office.context.mailbox.item.dateTimeCreated;
  const receivedDate = Office.context.mailbox.item.dateTimeModified;
  const localCreatedDate = getLocalDateString(new Date(createdDate));
  const localReceivedDate = getLocalDateString(new Date(receivedDate));
  const formattedCreatedDate = createdDate.toISOString();
  const formattedReceivedDate = receivedDate.toISOString();*/
  function toLocalISOString(date) {
    // Obtener la diferencia en minutos entre la hora local y UTC
    const tzOffsetMinutes = date.getTimezoneOffset();
    // Ajustar la fecha a la zona horaria local
    const localDate = new Date(date.getTime() - tzOffsetMinutes * 60000);

    // Construir manualmente el formato ISO 8601 completo
    return localDate.toISOString().replace('Z', '');
  }

  const createdDate = new Date(Office.context.mailbox.item.dateTimeCreated);
  const receivedDate = new Date(Office.context.mailbox.item.dateTimeCreated);

  // Obtener fechas en formato local ISO 8601
  const localCreatedDateISO = toLocalISOString(createdDate);
  const localReceivedDateISO = toLocalISOString(receivedDate);
  if(subject.length<1) {
    subject = "N/A";
  }
  return `https://localhost:3000/dialog.html?to=${to}&from=${from}&cc=${ccEmails.join(";")}&subject=${subject}&textBody=${textBody}&sentOn=${localCreatedDateISO}&receivedOn=${localReceivedDateISO}&attachNames=${attachNames.slice(1)}`;
}

function constructDialogUrl2(to, from, ccEmails, subject, attachNames = "", textBody) {
  /*const createdDate = new Date(Office.context.mailbox.item.dateTimeCreated);
  const receivedDate = new Date(Office.context.mailbox.item.dateTimeModified);

  const localCreatedDate = createdDate.toLocaleString(); // Fecha y hora locales
  const localReceivedDate = receivedDate.toLocaleString();

  const formattedCreatedDate = createdDate.toISOString();
  const formattedReceivedDate = receivedDate.toISOString();*/

  function toLocalISOString(date) {
    // Obtener la diferencia en minutos entre la hora local y UTC
    const tzOffsetMinutes = date.getTimezoneOffset();
    // Ajustar la fecha a la zona horaria local
    const localDate = new Date(date.getTime() - tzOffsetMinutes * 60000);

    // Construir manualmente el formato ISO 8601 completo
    return localDate.toISOString().replace('Z', '');
  }

  const createdDate = new Date(Office.context.mailbox.item.dateTimeCreated);
  const receivedDate = new Date(Office.context.mailbox.item.dateTimeCreated);

  // Obtener fechas en formato local ISO 8601
  const localCreatedDateISO = toLocalISOString(createdDate);
  const localReceivedDateISO = toLocalISOString(receivedDate);
  if(subject.length<1) {
    subject = "N/A";
  }
  return `https://localhost:3000/dialog.html?to=${to}&from=${from}&cc=${ccEmails}&textBody=${textBody}&subject=${subject}&sentOn=${localCreatedDateISO}&receivedOn=${localReceivedDateISO}&attachNames=${attachNames.slice(1)}`;
}

function getLocalDateString(dateString) {
  const date = new Date(dateString); // Convertir la cadena a un objeto Date
  return date.toLocaleString(); // Obtener la fecha en la zona horaria local
}

function openDialogAndSendData(url, attachmentContents, body, textBody) {
  Office.context.ui.displayDialogAsync(url, { height: 75, width: 50, displayInIframe: true }, function (asyncResult) {
    dialog = asyncResult.value;
    const mailData = { htmlBody: body, content: attachmentContents, textBody: textBody };

    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (event) {
      if (event.message === "ready") {
        dialog.messageChild(JSON.stringify(mailData));
      }
      if (event.message === "closeDialog") {
        dialog.close(); // Cerrar el diálogo
        Office.context.ui.closeContainer();
      }
    });
  });
}
