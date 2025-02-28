/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */
import "core-js";
import templateJson from './config.json';
var propertiesAnterior = null;
var cont = 0;
var tId = null;
Office.onReady(async (info) => {
  try {
	  var uri = `${templateJson.url_server}/filenet/searchtemplate`;
    var body = new URLSearchParams({ select: "Id, DocumentTitle, DateCreated, VersionSeries",
                                     from: "EntryTemplate",
                                     where: "IsCurrentVersion = true",
                                     maxrows: 1000,
                                     enable_content_search: false
                                    });

    var respuesta = await fetch(uri, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body,
    });

    if (!respuesta.ok) throw new Error(`HTTP status ${respuesta.status}`);
    var dataTemplate = await respuesta.json();
    
    dataTemplate.forEach(opcion => {
      CreateOptionList("selectTemplate", opcion.title, opcion.template);
    });
    var selectElement = document.getElementById("miSelect");
    var retorno = null;
    var properties = null;
    var className = null;
    var attachmentContents = null;
    var folder = null;
    var attachmentBody = null;
    var templateId = null;

    /*const button1 = document.getElementById("CargarButton1");
    if (!button1.hasAttribute("data-listener")) {
        button1.setAttribute("data-listener", "true");
        button1.addEventListener("click", miFuncion);
    }*/

    document.getElementById("CancelarButton").onclick = () => {Office.context.ui.messageParent("closeDialog"); };
    document.getElementById("selectTemplate").onchange = async function () { 

      retorno = await handleTemplateChange(this.value, info);
      properties = retorno.properties;
      className = retorno.className;
      attachmentContents = retorno.attachmentContents;
      folder = retorno.folder;
      attachmentBody = retorno.attachmentBody;
      templateId = retorno.templateId;
      tId = this.value;
    };
    
    var urlParams = new URLSearchParams(window.location.search);
    console.log(retorno);
    document.getElementById("CargarButton1").onclick = function (event) {
      event.stopPropagation();
      console.log(retorno);
        Office.context.ui.messageParent("ready");
        const mapValues = new Map();
        const mapTypes = new Map();
        
        var valor;
        for (var i = 0; i < properties.length; i++) {
          if (document.getElementById(properties[i].id+"-input") === null)
          {
            valor = "";
          }
          else {
            valor = document.getElementById(properties[i].id+"-input").value;
          }
          if (!properties[i].hidden==true) {
            if (valor == null) {
              valor="";
            }
            mapValues.set(properties[i].id, valor);
            if (properties[i].cardinality==="LIST" || properties[i].cardinality==="List") {
              mapTypes.set(properties[i].id, properties[i].dataType+"List");
            }
            else {
              mapTypes.set(properties[i].id, properties[i].dataType);
            }
          }
        }
        const mapArrayValues = Array.from(mapValues);
        const mapArrayTypes = Array.from(mapTypes);
        const jsonProperties = Object.fromEntries(mapArrayValues);
        const jsonTypes = Object.fromEntries(mapArrayTypes);
        SendDataToAPI(urlParams.get("attachNames"), className, attachmentContents, jsonProperties, templateJson.url_server+"/filenet/upload", folder, true, jsonTypes, attachmentBody, templateId);
      
    };    
    document.getElementById("CargarButton2").onclick = function (event) {
        event.stopPropagation();
        Office.context.ui.messageParent("ready");
        const mapValues = new Map();
        const mapTypes = new Map();
        var valor;
        for (var i = 0; i < properties.length; i++) {
          if (document.getElementById(properties[i].id+"-input") === null)
          {
            valor = "";
          }
          else {
            valor = document.getElementById(properties[i].id+"-input").value;
          }
          if (!properties[i].hidden==true) {
            if (valor == null) {
              valor="";
            }
            mapValues.set(properties[i].id, valor);
            if (properties[i].cardinality==="LIST" || properties[i].cardinality==="List") {
              mapTypes.set(properties[i].id, properties[i].dataType+"List");
            }
            else {
              mapTypes.set(properties[i].id, properties[i].dataType);
            }
          }
        }
        const mapArrayValues = Array.from(mapValues);
        const mapArrayTypes = Array.from(mapTypes);
        const jsonProperties = Object.fromEntries(mapArrayValues);
        const jsonTypes = Object.fromEntries(mapArrayTypes);
        if (urlParams.get("attachNames")) {
          SendDataToAPI(urlParams.get("attachNames"), className, attachmentContents, jsonProperties, templateJson.url_server+"/filenet/upload", folder, false, jsonTypes, attachmentBody, templateId);
        } else {
          SendDataToAPI("", "className", "", "", "/filenet/upload", "", false, "", attachmentBody, templateId);
        }
      
    };
  } catch (error) {
    console.error("Mensaje de error: ", error);
  }
});

async function handleTemplateChange(templateId, info) {
  try {
    // Solicitud para obtener el template desde el servidor
    var url2 = `${templateJson.url_server}/filenet/download`;
    var databody = new URLSearchParams({ docId: templateId.slice(1, -1) });
    var response = await fetch(url2, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: databody,
    });

    if (!response.ok) throw new Error(`HTTP status ${response.status}`);
    
    var data = await response.json();
    var jsonResponse = JSON.parse(atob(data[0].content));
    var arrayFolder = jsonResponse.folder.split(",");
    var folderId = arrayFolder[2];

    // Solicitud para obtener los detalles de la carpeta
    var url = `${templateJson.url_server}/filenet/getfolder`;
    var dataBody = new URLSearchParams({ folderId });
    var response2 = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: dataBody,
    });

    if (!response2.ok) throw new Error(`HTTP status ${response2.status}`);
    
    var data2 = await response2.json();
    var jsonResponse2 = data2[0].folder;

    // Definir las variables de clase y propiedades
    var className = jsonResponse.addClassName ?? null;
    var properties = jsonResponse.propertiesOptions ?? null;
    var saveInElement = jsonResponse.allowUserSelectFolder ?? null;

    // Si estamos en el contexto de Outlook
    if (info.host === Office.HostType.Outlook) {
      // Configuración del manejador para el evento de recibir mensaje del padre (Outlook)
      Office.onReady(() => {
        Office.context.ui.addHandlerAsync(
          Office.EventType.DialogParentMessageReceived,
          onMessageFromParent
        );
      });

      Office.context.ui.messageParent("ready");

      var urlParams = new URLSearchParams(window.location.search);
      
      return new Promise((resolve, reject) => {
        // Recibir mensaje del padre (Outlook)
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (event) => {
          try {
            var mailData = JSON.parse(event.message);
            var attachmentContents = mailData.content;
            var attachmentBody = mailData.htmlBody;

            // Preparar los valores para establecer en los elementos
            SetValuesToElements("classNameElem", className);
            SetValuesToElements("saveInFolder", jsonResponse2);
            SetValuesToElements("fileNames", urlParams.get("attachNames") || "");

            // Habilitar o deshabilitar el botón de cargar
            var cargarButton2 = document.getElementById("CargarButton2");
            cargarButton2.disabled = !attachmentContents.length;

            // Eliminar los elementos previos de propiedades, si es necesario
            if (propertiesAnterior !== null && propertiesAnterior !== undefined) {
              propertiesAnterior.forEach(prop => RemoveElementByProperties(prop.name, prop.id, prop.readOnly, urlParams, prop));
            }

            propertiesAnterior = properties;
            
            // Crear los nuevos elementos de propiedades
            properties.forEach(prop => CreateElementByProperties(prop.name, prop.id, prop.readOnly, urlParams, prop, templateId));

            // Preparar el objeto retorno con los valores procesados
            let retorno = {
              properties: properties,
              className: className,
              attachmentContents: attachmentContents,
              folder: jsonResponse.folder,
              attachmentBody: attachmentBody, 
              templateId: templateId
            };
            Office.context.ui.messageParent("done");
            // Resolver la promesa con los datos
            resolve(retorno);

          } catch (error) {
            console.error("Mensaje de error: ", error);
            reject(error); // Rechazar en caso de error
          }
        });
      });
    }

    return null; // Si no estamos en Outlook, retornar null
  } catch (error) {
    console.error("Error en handleTemplateChange:", error);
    return null; // En caso de error, retornar null
  }
}

async function SendDataToAPI(name, classname, content, properties, url, folder, oneFile, jsonTypes, attachmentBody, templateId) {
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (event) => {
    const mailData = JSON.parse(event.message);
    var attachmentContents = mailData.content;
    var htmlBody = mailData.htmlBody;
    var textBody = mailData.textBody;
    var attachArrayContents = "";
    var attachArrayNames = "";
    const arrayFolder = folder.split(",");
    var folderId = arrayFolder[2];
    // Display or process attachment contents here
    if (Array.isArray(attachmentContents)) {
      attachmentContents.forEach(attachment => {
        //console.log(`Attachment Name: ${attachment.name}`);
        //console.log(`Content Format: ${attachment.format}`);
        //console.log(`Content:`, attachment.content);
        attachArrayContents = attachArrayContents  + attachment.content + ";";
        attachArrayNames = attachArrayNames + attachment.name + ";";

    });
    }

    const payload = {
      body: String(htmlBody),
      properties: JSON.stringify(properties),
      name: String(attachArrayNames),
      classname: String(classname),
      content: String(attachArrayContents),
      folderId: String(folderId),
      oneFile: oneFile,
      types: JSON.stringify(jsonTypes),
      templateId: String(templateId)
  };
    const response = fetch(url, {
      method: "POST", // or 'POST', 'PUT', etc.
      //agent: agent, // Attach the agent here
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    })
      .then((response) => {
        if (!response.ok) {         
          return response.text().then(text => { Office.context.ui.messageParent(JSON.stringify({ status: "error", data: text })); });
        }
        else if (response.ok){
          console.log('Success:', response);
          ejecutarConRetraso();
          return response.json();
        }
      })
      .catch((error) => {
        console.error("Fetch error:", error);
      });
  });
  
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function ejecutarConRetraso() {
  await delay(2000); // 3 segundos de retraso
  Office.context.ui.messageParent("closeDialog");
}

function SetValuesToElements(name, value) {
  var elem = document.getElementById(name);
  if (elem) {
    elem.readOnly = true;
    elem.value = value;
  }
}

function CreateOptionList(element, display, value) {
  let insertAt = document.getElementById(element);
  var nuevaOpcion = document.createElement("option");
  nuevaOpcion.value = value;
  nuevaOpcion.textContent = display;
  insertAt.appendChild(nuevaOpcion);
}

function RemoveElementByProperties(name, id, readOnly, urlParams, properties) {
  if (id !== "DocumentTitle" && !properties.hidden) {
    var insertAtPropiedadesPadre = document.getElementById("table-propiedades");
    var insertAtPropiedades = document.getElementById(`${id}-tr`);
    insertAtPropiedadesPadre.innerHTML = "";
    if (insertAtPropiedades) { 
      insertAtPropiedades.remove();
      console.log(`Eliminado: ${id}-tr`);
    } else { 
      console.warn(`No se encontró el elemento para eliminar: ${id}-tr`); 
    }
  }
}

function CreateElementByProperties(name, id, readOnly, urlParams, properties, templateId) {
  if (id === "DocumentTitle" || properties.hidden === true) return;
  
  let tablePropiedades = document.getElementById("table-propiedades");
  if (!tablePropiedades) return;
  
  let tr = document.createElement("tr");
  tr.className = "border border-1";
  tr.id = `${id}-tr`;
  tablePropiedades.appendChild(tr);
  
  let tdName = document.createElement("td");
  tdName.style.width = "30%";
  tdName.textContent = `${name}:`;
  tdName.id = `${id}-td1`;
  tr.appendChild(tdName);
  
  let tdValue = document.createElement("td");
  tdValue.style.width = "70%";
  tdValue.className = "bg-secondary-subtle border-opacity-10";
  tdValue.id = `${id}-td2`;
  tr.appendChild(tdValue);
  
  let divCuadro = document.createElement("div");
  divCuadro.className = "row-cols-1 d-flex justify-content-center";
  divCuadro.id = `${id}-div`;
  tdValue.appendChild(divCuadro);
  
  let inputCuadro = document.createElement("input");
  inputCuadro.id = `${id}-input`;
  inputCuadro.className = "text";
  inputCuadro = InsertValuePerProperty(inputCuadro, id, urlParams, readOnly, properties);
  divCuadro.appendChild(inputCuadro);
}

function InsertValuePerProperty(inputCuadro, id, urlParams, readOnly, properties) {
  if (id === templateJson.from) {
    inputCuadro.value = urlParams.get("from");
  } else if (id === templateJson.subject) {
    inputCuadro.value = urlParams.get("subject");
  } else if (id === templateJson.body) {
    if(urlParams.get("textBody")==="undefined"){
      inputCuadro.value = urlParams.get("attachNames");
    }
    else{
      inputCuadro.value = urlParams.get("textBody");
    }
    
  } else if (id === templateJson.to) {
    inputCuadro.value = urlParams.get("to");
  } else if (id === templateJson.cc) {
    inputCuadro.value = urlParams.get("cc");
  } else if (id === templateJson.sentOn) {
    inputCuadro.value = urlParams.get("sentOn");
  } else if (id === templateJson.receivedOn) {
    inputCuadro.value = urlParams.get("sentOn");
  }
  else if ("defaultValue" in properties) {
    inputCuadro.value = properties.defaultValue;
  }
  if (readOnly === true) {
    inputCuadro.readOnly = true;
  }
  return inputCuadro;
}

function onMessageFromParent(arg) {
  let insertAt2 = document.getElementById("item-subject");
}