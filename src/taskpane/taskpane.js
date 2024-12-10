/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, localStorage, $, setTimeout, window, console */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    getWordLanguage();
    document.getElementById("auth-token").style.display = "none";
    document.getElementById("bg-color").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("tabs").style.display = "none";
    document.getElementById("tabs-stage").style.display = "none";
    document.getElementById("submitLogin").onclick = validateAndSubmitLogin;
    document.getElementById("eliminar-token").onclick = handleDeletToken;
    // document.getElementById("cancelAuthToken").onclick = hideAuthTokenDialog;

    document.getElementById("classify-document-container").onclick = toggleDocumentBox;
    document.getElementById("close-document-container").onclick = toggleDocumentBox;

    document.getElementById("get-classify-auto").onclick = handleSendDoc;

    document.getElementById("reload-plugin").onclick = restartPlugin;

    document.getElementById("openDialogButton").onclick = openOfficeDialog;
    document.getElementById("openDialogButton2").onclick = openPopup;

    // document.getElementById("btnHiddenContent").onclick = toggleContent;
    // document.getElementById("btnHiddenContent2").onclick = toggleContent2;
    // document.getElementById("toggleSubContent").onclick = toggleSubContent2;

    //Obtengo el idioma del documento y el país del documento desde el documento.
    // document.getElementById("idiomaDoc").onclick = getDocWord;
    // Revisa
    document.getElementById("get-match-groups").onclick = getMatchGroups;
    document.getElementById("get-match-groups").onclick = getMatchGroups;

    // document.getElementById("enviar-texto").onclick = handleSendText;
    document.getElementById("clean-document").onclick = newChat;
    // document.getElementById("revisa-btn").onclick = getWordText;
    document.getElementById("document-forward").addEventListener("click", handleSendAllText);
    document.getElementById("send-btn").onclick = sendInputText;
    document.getElementById("chat-input").onkeypress = (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        sendInputText();
      }
    };
    setupTabs();
    // assignDynamicIds();
  }
  run();
});

//función para ver el lenguaje del office
function getWordLanguage() {
  const lang = Office.context.contentLanguage;
  console.log("Idioma actual del documento: ", lang);

  if (lang === "es-AR") {
    console.log("El idioma es Español");
  } else if (lang === "en-US") {
    console.log("El idioma es Inglés");
  } else {
    console.log("Idioma no detectado");
  }
}

export async function run() {
  return Word.run(async (context) => {
    const authToken = localStorage.getItem("authTokenMia");

    setupMenuEvents();

    if (!authToken) {
      hideLoadingIndex();
      showAuthToken();
    } else {
      newChat();
      sessionActiveToken();
      //validAuthToken(authToken);
    }
    await context.sync();
  });
}

///////////////////////////////////////////////////////////////////////////////////////////////
function setupMenuEvents() {
  // Delegar evento para los interruptores
  $(document).on("click", ".switch-container", function (e) {
    e.stopPropagation(); // Detener propagación para que no afecte al colapsador
  });

  // Delegar evento para el icono de refrescar
  $(document).on("click", ".icon-refresh-revisa", function (e) {
    e.stopPropagation(); // Detener propagación para que no afecte al colapsador
  });

  // Delegar eventos para los menús y submenús
  $(document).on("click", ".collapser", function (e) {
    const $currentCollapser = $(this);
    const $parentMenu = $currentCollapser.closest(".collapse"); // Contenedor del submenú
    const $menuGroup = $currentCollapser.closest(".menu-group"); // Contenedor principal del grupo

    // Si es un submenú
    if ($parentMenu.length > 0) {
      // Cierra otros submenús dentro del mismo menú principal
      $parentMenu
        .find(".collapser")
        .not($currentCollapser)
        .removeClass("open")
        .addClass("collapsed")
        .next(".collapse")
        .collapse("hide");
    } else {
      // Si es un menú principal, cierra otros menús principales dentro del grupo
      $menuGroup
        .find(".collapser")
        .not($currentCollapser)
        .removeClass("open")
        .addClass("collapsed")
        .next(".collapse")
        .collapse("hide");
      console.log("entro en el else");
    }

    // Alterna el estado del menú actual o submenú
    $currentCollapser.toggleClass("collapsed open");
    $currentCollapser.next(".collapse").collapse("toggle");

    // Evita que el evento afecte a niveles superiores
    e.stopPropagation();
  });
  // // Detener propagación en switch-container
  // $(".switch-container").on("click", function (e) {
  //   e.stopPropagation(); // Detener el evento para que no afecte al colapsador
  // });

  // // Detener propagación en icon-refresh-revisa
  // $(".icon-refresh-revisa").on("click", function (e) {
  //   e.stopPropagation(); // Detener el evento para que no afecte al colapsador
  // });

  // // Evento para el menú
  // $(".menu-group").each(function () {
  //   const $menuGroup = $(this);

  //   $menuGroup.find(".collapser").on("click", function (e) {
  //     const $currentCollapser = $(this);
  //     const $parentMenu = $currentCollapser.closest(".collapse"); // Contenedor del submenú

  //     // Si es un submenú
  //     if ($parentMenu.length > 0) {
  //       // Cierra otros submenús dentro del mismo menú principal
  //       $parentMenu
  //         .find(".collapser")
  //         .not($currentCollapser)
  //         .removeClass("open")
  //         .addClass("collapsed")
  //         .next(".collapse")
  //         .collapse("hide");
  //     } else {
  //       // Si es un menú principal, cierra otros menús principales dentro del grupo
  //       $menuGroup
  //         .find(".collapser")
  //         .not($currentCollapser)
  //         .removeClass("open")
  //         .addClass("collapsed")
  //         .next(".collapse")
  //         .collapse("hide");
  //     }

  //     // Alterna el estado del menú actual o submenú
  //     $currentCollapser.toggleClass("collapsed open");
  //     $currentCollapser.next(".collapse").collapse("toggle");

  //     // Evita que el evento afecte a niveles superiores
  //     e.stopPropagation();
  //   });
  // });
}

// function openPopup() {
//   let contentData = "Este es un mensaje simple enviado desde el complemento.";

//   // Abrir una nueva ventana y pasar el texto como parámetro en la URL
//   let popup = window.open(
//     `https://victorwake.github.io/MiaWord/swal.html?message=${encodeURIComponent(contentData)}`,
//     "PopupWindow",
//     "width=400,height=300"
//   );

//   // Opcionalmente, puedes esperar a que se cargue la ventana para hacer algo más
//   popup.onload = function () {
//     console.log("Ventana emergente cargada");
//   };
// }

// function assignDynamicIds() {
//   // Inicializamos contadores para los IDs
//   let boxCounter = 1;
//   let contentCounter = 1;

//   // Asignar IDs a los elementos con la clase "grupo3-box"
//   document.querySelectorAll(".grupo3-box").forEach((box) => {
//     // Verifica si el div ya tiene un ID. Si no tiene, le asignamos uno nuevo.
//     if (!box.id) {
//       box.id = `sub-menu${boxCounter}`;
//       boxCounter++; // Incrementamos el contador solo para los elementos que reciben un nuevo ID
//     }
//   });

//   // Asignar IDs a los elementos con la clase "sub-hidden-content"
//   document.querySelectorAll(".sub-hidden-content").forEach((content) => {
//     // Verifica si el div ya tiene un ID. Si no tiene, le asignamos uno nuevo.
//     if (!content.id) {
//       content.id = `subHiddenContent${contentCounter}`;
//       contentCounter++; // Incrementamos el contador solo para los elementos que reciben un nuevo ID
//     }
//   });
// }

// function assignDynamicIds() {
//   let boxCounter = 1;
//   let contentCounter = 1;

//   document.querySelectorAll(".grupo3-box").forEach((box) => {
//     if (!box.id) {
//       box.id = `sub-menu${boxCounter}`;
//       boxCounter++;
//     }
//     box.addEventListener("click", () => toggleSubContent(box.id));
//   });

//   document.querySelectorAll(".sub-hidden-content").forEach((content) => {
//     if (!content.id) {
//       content.id = `subHiddenContent${contentCounter}`;
//       contentCounter++;
//     }
//   });
// }

// Función para expandir o contraer el submenú correspondiente
// function toggleSubContent(menuId) {
//   // Obtiene el número del submenú a partir del ID
//   const menuNumber = menuId.replace("sub-menu", "");
//   const subContent = document.getElementById(`subHiddenContent${menuNumber}`);

//   // Identifica a qué grupo de menú pertenece (menu1 o menu2)
//   const parentMenuClass = subContent.closest(".menu1") ? "menu1" : "menu2";

//   // Cierra cualquier submenú abierto dentro del mismo grupo
//   document.querySelectorAll(`.${parentMenuClass} .sub-hidden-content.expanded`).forEach((openSubContent) => {
//     if (openSubContent !== subContent) {
//       openSubContent.classList.remove("expanded");
//       openSubContent.style.height = "0";
//     }
//   });

//   // Alterna el submenú actual
//   subContent.classList.toggle("expanded");
//   if (subContent.classList.contains("expanded")) {
//     subContent.style.height = subContent.scrollHeight + "px";
//   } else {
//     subContent.style.height = "0";
//   }
// }

function openPopup(dataCode) {
  // Verificar si el contenido de dataCode es JSON válido
  let alternatives;
  try {
    alternatives = JSON.parse(dataCode.replace(/&quot;/g, '"'));
  } catch (error) {
    // Si no es JSON válido, asumir que es un mensaje personalizado
    alternatives = dataCode;
  }

  // Define el tamaño de la ventana emergente
  const width = 600;
  const height = 500;

  // Calcula la posición para centrar la ventana
  const left = window.screen.width / 2 - width / 2;
  const top = window.screen.height / 2 - height / 2;

  // Abre la ventana emergente centrada
  let popupWindow = window.open(
    "https://victorwake.github.io/MiaWord/postMessage.html",
    "PopupWindow",
    `width=${width},height=${height},left=${left},top=${top},toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no`
  );

  // Escuchar el mensaje del popup para saber cuándo está listo
  window.addEventListener("message", (event) => {
    if (event.data === "ready" && popupWindow) {
      // Enviar el contenido del data-code al popup
      popupWindow.postMessage({ alternatives }, "*");
    }
  });
}

function openOfficeDialog(dataCode) {
  // URL del diálogo
  const dialogUrl = "https://victorwake.github.io/MiaWord/tuDialogo.html";

  // Verificar si el contenido de dataCode es JSON válido
  let alternatives;
  try {
    alternatives = JSON.parse(dataCode.replace(/&quot;/g, '"'));
  } catch (error) {
    // Si no es JSON válido, asumir que es un mensaje personalizado
    alternatives = dataCode;
  }

  // Abrir el diálogo de Office
  Office.context.ui.displayDialogAsync(dialogUrl, { height: 50, width: 50, displayInIframe: true }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Error al abrir el diálogo:", asyncResult.error.message);
    } else {
      const dialog = asyncResult.value;

      console.log("Diálogo abierto exitosamente.");

      // Agregar manejador de evento para recibir mensajes del diálogo
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
        console.log("Mensaje recibido del diálogo:", args.message);
        dialog.close();
      });

      // Enviar el mensaje una vez abierto el diálogo
      setTimeout(function () {
        if (dialog) {
          // Enviar el contenido como JSON
          dialog.messageChild({ alternatives });
          console.log("Mensaje enviado al diálogo:", alternatives);
        } else {
          console.error("Error: El diálogo no se ha abierto correctamente.");
        }
      }, 2000);
    }
  });
}

// function openPopup() {
//   let contentData = "Este es un mensaje simple enviado desde el complemento.";

//   // Abrir una nueva ventana y pasar el texto como parámetro en la URL
//   let popup = window.open(
//     `https://victorwake.github.io/MiaWord/popup.html?message=${encodeURIComponent(contentData)}`,
//     "PopupWindow",
//     "width=400,height=300"
//   );

//   // Opcionalmente, puedes esperar a que se cargue la ventana para hacer algo más
//   popup.onload = function () {
//     console.log("Ventana emergente cargada");
//   };
// }

function isAuthTokenAvailable(authToken) {
  showLoadingMessage();
  if (authToken) {
    hideResponseSession();
    hideLoadingMessage();
    hideLoadingIndex();
    showTabs();
    // showClassifyDocument();
    document.getElementById("classify-document").style.display = "block";
    return localStorage.getItem("authTokenMia") !== null;
  } else {
    hideLoadingMessage();
    hideLoadingIndex();
    console.log("Token is not available.");
  }
}

// ##################################################################################################

// ##################################################################################################
function sessionActiveToken() {
  const authToken = getAuthToken();

  const data = {
    documentPath: null,
    documentText: "Doc",
    onlyCountries: null,
  };

  $.ajax({
    url: "https://miadev.miaintelligence.com:444/api/getDoclangDoccountryFromDocument",
    type: "POST",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${authToken}`,
    },
    data: JSON.stringify(data),
    success: function (response) {
      if (response) {
        isAuthTokenAvailable(authToken);
        getGenerativeClassifierConfig();
        console.log("sessionActiveToken: OK");
      }
    },
    error: function (xhr, status, error) {
      console.log("Error:", error);
      console.log("Status:", status);
      console.log("Response Text:", xhr.responseText);
      console.log("Ready State:", xhr.readyState);
      console.log("HTTP Status Code:", xhr.status);
      try {
        // Intenta parsear la respuesta JSON
        if (xhr.status === 401) {
          console.log("entro sessionActiveToken 401");
          showPopupEndOfSession();
        }
      } catch (e) {
        console.log("Error al parsear JSON de la respuesta:", e);
      }
    },
  });
}

// #####################################################################################
//                           Inicio idioma de documento y pais
// #####################################################################################
function getDoclangDoccountryFromDocument(normative_refs, criteria, buttonId, doc) {
  // addSpinAnimationRevisa();
  const authToken = getAuthToken();

  const data = {
    documentPath: null,
    documentText: doc,
    onlyCountries: [],
  };
  console.log("lo que le envio idioma del doc y pais: ", JSON.stringify(data));

  $.ajax({
    url: "https://miadev.miaintelligence.com:444/api/getDoclangDoccountryFromDocument",
    type: "POST",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${authToken}`,
    },
    data: JSON.stringify(data),
    success: function (response) {
      // console.log("idioma del doc y pais: ", JSON.stringify(response));
      getDocEvaluation(normative_refs, criteria, buttonId, doc, response);
      // clearContainers();
      // renderRevisa(response);
    },
    error: function (xhr, status, error) {
      // hideLoadingMessageEnv();
      // hideTypingIndicator();
      let errorMessage;
      console.log("Estado del error: ", status);
      console.log("Estado del error: ", error);
      if (xhr.status === 401) {
        errorMessage = xhr.responseJSON.message;
        showPopupEndOfSession();
      } else {
        errorMessage = "Error del servidor.";
      }
      // displayResponseText({ message: errorMessage });
      console.log({ message: errorMessage });
    },
    complete: function () {
      // hideLoadingMessageEnv();
    },
  });
}

function getDocWord(normative_refs, criteria, buttonId) {
  Word.run(async (context) => {
    const document = context.document;
    const body = document.body;
    body.load("text");

    await context.sync();

    const doc = body.text.trim() || "Documento vacio";

    getDoclangDoccountryFromDocument(normative_refs, criteria, buttonId, doc);
  });
}

// #####################################################################################
//                            Fin idioma de documento y pais
// #####################################################################################

// #####################################################################################
//                           Inicio evaluación del Doc.
// #####################################################################################
function spinCriterio(buttonId, action) {
  // Encontrar el contenedor más cercano con el ID del botón
  const cardBody = document.getElementById(buttonId)?.closest(".card-body.sub.px-3");

  if (cardBody) {
    // Buscar el ícono con las clases específicas dentro del contenedor
    const refreshIcon = cardBody.querySelector(".icon.icon-refresh.icon-white");

    if (refreshIcon) {
      if (action === "add") {
        // Agregar la clase si la acción es "add"
        refreshIcon.classList.add("rotate-icon-refresh");
      } else if (action === "remove") {
        // Quitar la clase si la acción es "remove"
        refreshIcon.classList.remove("rotate-icon-refresh");
      } else {
        console.log("Acción no válida:", action);
      }
    } else {
      console.log("No se encontró un ícono con la clase 'icon icon-refresh icon-white' dentro del contenedor.");
    }
  } else {
    console.log("No se encontró el contenedor para el ID:", buttonId);
  }
}

// Función que busca el contenido basado en el ID del botón
function findContentByButtonId(buttonId) {
  // Usamos el ID para encontrar el contenedor más cercano
  const cardBody = document.getElementById(buttonId)?.closest(".card-body.sub.px-3");

  if (cardBody) {
    spinCriterio(buttonId, "add");
    const textElements = cardBody.querySelectorAll("p.text-xs");
    const normative_refs = Array.from(textElements).map((p) => p.textContent.trim());

    const criteria = {
      dataCode: cardBody.dataset.code, // data-code del cardBody

      weight: parseInt(cardBody.querySelector(".badge-text-ligth-blue.weight")?.dataset.code || "0", 10),

      integrityHash: cardBody.querySelector(".integrity-hash")?.dataset.code || null,

      exclude: cardBody.querySelector(".exclude-class")?.dataset.code === "true",

      description: cardBody.querySelector(".span-sub-content")?.dataset.code || null,
    };
    getDocWord(normative_refs, criteria, buttonId);
  } else {
    console.log("No se encontró el contenedor para el ID:", buttonId);
  }
}

function getDocEvaluation(normative_refs, criteria, buttonId, doc, response) {
  const authToken = getAuthToken();

  // Construir el arreglo de objetos para `normative_refs`
  const parsedNormativeRefs = [];
  for (let i = 0; i < normative_refs.length; i += 2) {
    parsedNormativeRefs.push({
      normative_reference: normative_refs[i] || "", // Tomar el primer elemento como normative_reference
      threat: normative_refs[i + 1] || "", // Tomar el segundo elemento como threat
    });
  }
  // console.log("Parámetros recibidos:");
  // console.log("normative_refs:", normative_refs);
  // console.log("doc:", doc);
  // console.log("response:", response);
  // console.log("Arreglo completo de objetos construido:", parsedNormativeRefs);
  const doclang = (response && response.es) || "es";
  const country = (response && response.country) || "España";

  const doctypeElement = document.getElementById("tipo-doc-select");
  const subjectElement = document.getElementById("materia-select");

  const doctypeCode = doctypeElement?.value || null;
  const subject = subjectElement?.value || null;

  const data = {
    doc: doc, //le paso el documento
    scope: "nacional",
    doclang: doclang,
    doctype: doctypeCode,
    subject: subject,
    criteria: [
      {
        code: criteria.dataCode,
        normative_refs: parsedNormativeRefs, // Usar los valores procesados,
        integrity_hash: criteria.integrityHash,
        weight: criteria.weight,
        description: criteria.description,
        exclude: criteria.exclude,
        title: doc,
        content: doc,
      },
    ],
    documentPath: null,
    country: country,
  };
  console.log("lo que le envio a qualifyGenerative: ", JSON.stringify(data));

  $.ajax({
    url: "https://miadev.miaintelligence.com:444/api/qualifyGenerative",
    type: "POST",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${authToken}`,
    },
    data: JSON.stringify(data),
    success: function (response) {
      console.log("idioma del doc y pais: ", JSON.stringify(response));
      spinCriterio(buttonId, "remove");
      addCriterias(response, buttonId);
      // clearContainers();
      // renderRevisa(response);
    },
    error: function (xhr, status, error) {
      // hideLoadingMessageEnv();
      // hideTypingIndicator();
      let errorMessage;
      console.log("Estado del error: ", status);
      console.log("Estado del error: ", error);
      if (xhr.status === 401) {
        errorMessage = xhr.responseJSON.message;
        showPopupEndOfSession();
      } else {
        errorMessage = "Error del servidor.";
      }
      // displayResponseText({ message: errorMessage });
      console.log({ message: errorMessage });
    },
    complete: function () {
      // hideLoadingMessageEnv();
    },
  });
}

function addCriterias(response, buttonId) {
  const evaluation = response.data.data[0]?.evaluation;
  const explanation = response.data.data[0]?.explanation;
  const alternatives = response.data.data[0]?.alternatives || [];
  let resultText = "";

  // Determina el texto basado en la evaluación
  if (evaluation === "ko") {
    resultText = "NO CUMPLE";
  } else if (evaluation === "ok") {
    resultText = "CUMPLE";
  } else {
    resultText = "EVALUACIÓN DESCONOCIDA"; // Caso por defecto si no es ni "ok" ni "ko"
  }

  // Construye el contenido HTML
  let criteria = ` 
    <div class="!pt-4 ng-star-inserted mt-3">
      <span class="!font-semibold leading-6"> 
        RESULTADO:
        <span id="color-inserted" class="${evaluation === "ko" ? "text-red" : "text-green"}">${resultText}</span>
      </span>
      <br>
      <span class="span-sub-content mt-3">
      ${explanation}
      </span>
    </div>`;

  const button = document.getElementById(buttonId);
  if (button) {
    const correctButton = button.parentElement.querySelector(".btn-revisa.correct");
    if (correctButton) {
      // Agregar clases adicionales al elemento encontrado
      correctButton.classList.add("correct-act", "cursor-pointer");
      correctButton.setAttribute("data-code", JSON.stringify(alternatives));
      console.log("Clases y data-code agregados:", correctButton);
    } else {
      console.error("No se encontró el elemento con la clase 'correct'");
    }
    // Encuentra el div padre del botón
    const parentDiv = button.closest(".card"); // Ahora solo busca dentro del card
    if (parentDiv) {
      // Encuentra el div que contiene el botón y coloca el contenido antes de él
      const buttonParentDiv = button.closest("div"); // Este es el div donde está el botón
      if (buttonParentDiv) {
        // Inserta el contenido antes del div que contiene el botón
        buttonParentDiv.insertAdjacentHTML("beforebegin", criteria);
      } else {
        console.error("No se encontró un div padre del botón.");
      }

      // Busca el contenedor con la clase "rec-box" dentro de este card específico
      const recBox = parentDiv.querySelector(".rec-box");
      if (recBox) {
        // Limpia las clases existentes en recBox y asigna nuevas basadas en la evaluación
        recBox.classList.remove("rec-red", "rec-green", "rec-grey"); // Elimina clases previas
        recBox.classList.add(evaluation === "ko" ? "rec-red" : evaluation === "ok" ? "rec-green" : "rec-grey");

        // Buscar el <span> con la clase 'icon' dentro de recBox
        const iconSpan = recBox.querySelector(".icon");
        if (iconSpan) {
          // Limpia las clases del <span>, excepto 'icon'
          iconSpan.className = "icon"; // Resetea a solo 'icon'

          // Agrega la clase correspondiente según la evaluación
          if (evaluation === "ok") {
            iconSpan.classList.add("icon-check");
          } else if (evaluation === "ko") {
            iconSpan.classList.add("icon-x");
          }
        } else {
          console.error("No se encontró un <span> con la clase 'icon' dentro de rec-box.");
        }
      } else {
        console.error("No se encontró un elemento con la clase 'rec-box' dentro del card.");
      }
    } else {
      console.error("No se encontró el div padre con clase 'card' para el botón con ID:", buttonId);
    }
  } else {
    console.error("No se encontró un botón con ID:", buttonId);
  }
}

// #####################################################################################
//                            Fin evaluación del Doc.
// #####################################################################################

// #####################################################################################
//                         Inicio Clasificacion automatica
// #####################################################################################
function classifyGenerative(doc) {
  // console.log("Texto que le llega a sendDocToServer: " + doc);
  const authToken = getAuthToken();

  const data = {
    doc: doc,
  };

  // console.log("Payload enviado:", JSON.stringify(data));

  $.ajax({
    url: "https://miadev.miaintelligence.com:444/api/classifyGenerative", // El nuevo endpoint
    type: "POST",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${authToken}`,
    },
    data: JSON.stringify(data),
    success: function (response) {
      console.log("Clasificacion automatica: ", JSON.stringify(response));
      displayResponseDocType(response.data);
    },
    error: function (xhr, status, error) {
      hideLoadingMessageEnv();
      hideTypingIndicator();
      let errorMessage;
      console.log("Estado del error: ", status);
      console.log("Estado del error: ", error);
      if (xhr.status === 401) {
        errorMessage = xhr.responseJSON.message;
        showPopupEndOfSession();
        // console.log("entro");
        // const responseDocType = document.getElementById("response-session");
        // unauthorizedToken();
        // hideClassifyDocument();
        // responseDocType.innerText = "La sesión ha caducado.";
        // showResponseSession();
      } else {
        errorMessage = "Error del servidor.";
      }
      displayResponseText({ message: errorMessage });
    },
    complete: function () {
      // hideLoadingMessageEnv();
    },
  });
}

function handleSendDoc() {
  Word.run(async (context) => {
    removeDocTypeText();
    addSpinAnimation();
    const document = context.document;
    const body = document.body;
    body.load("text");

    await context.sync();

    const doc = body.text.trim() || "doc vacio";

    classifyGenerative(doc);
  });
}

function displayResponseDocType(response) {
  const responseDocType = document.getElementById("doctype");
  if (response && response.message) {
    // Si el objeto contiene 'message', lo mostramos
    removeSpinAnimation();
    responseDocType.innerText = response.message;
  } else if (response && response.doctype) {
    // Si el objeto contiene 'doctype', mostramos el tipo de documento
    removeSpinAnimation();
    responseDocType.innerText = response.doctype;
    initializeSelectsFromData(response);
  } else {
    // Si no hay respuesta válida
    removeSpinAnimation();
    responseDocType.innerText = "No response from server.";
  }
}

function addSpinAnimation() {
  const element = document.getElementById("spin");
  const spingRefresh = document.getElementById("sping-refresh");
  if (element) {
    element.classList.add("icon-refresh-2", "w-600", "spin-animation");
  }
  if (spingRefresh) {
    spingRefresh.classList.add("spin-animation");
  }
}

function removeSpinAnimation() {
  const element = document.getElementById("spin");
  const spingRefresh = document.getElementById("sping-refresh");
  if (element) {
    element.classList.remove("icon-refresh-2", "w-600", "spin-animation");
  }
  if (spingRefresh) {
    spingRefresh.classList.remove("spin-animation");
  }
}
function removeDocTypeText() {
  const responseDocType = document.getElementById("doctype");
  if (responseDocType) {
    responseDocType.innerText = "";
  }
}

// #####################################################################################
//                       Inicio Campos opcion clasificacion automatica
// #####################################################################################
function getGenerativeClassifierConfig() {
  const authToken = getAuthToken();

  const data = {
    lang: "es",
  };

  $.ajax({
    url: "https://miadev.miaintelligence.com:444/api/getGenerativeClassifierConfig",
    type: "POST",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${authToken}`,
    },
    data: JSON.stringify(data),
    success: function (response) {
      if (response && response.data) {
        // Pasa el objeto configData a la función para manejar el select
        handleSelects(response.data);
        handleSendDoc();

        setTimeout(function () {
          typeDocument();
        }, 6000);
      }
    },
    error: function (xhr, status, error) {
      console.log("Error:", error);
      console.log("Status:", status);
      console.log("Response Text:", xhr.responseText); // Ver el mensaje detallado del servidor
      console.log("Ready State:", xhr.readyState); // Ver el estado del objeto XMLHttpRequest
      console.log("HTTP Status Code:", xhr.status); // Ver el código de estado HTTP
      try {
        // Intenta parsear la respuesta JSON
        if (xhr.status === 401) {
          showPopupEndOfSession();
          console.log("entro unauthorizedToken");
          // const responseDocType = document.getElementById("response-session");
          // unauthorizedToken();
          // hideClassifyDocument();
          // responseDocType.innerText = "La sesión ha caducado.";
          // showResponseSession();
        }
      } catch (e) {
        console.log("Error al parsear JSON de la respuesta:", e);
      }
    },
  });
}

function handleSelects(configData) {
  // Llenar el primer select (Materia)
  populateMateriaSelect(configData.subject_codes);

  // Llenar el segundo select (Tipo de documento) con la selección por defecto
  const firstMateria = document.getElementById("materia-select").value;
  if (firstMateria) {
    populateTipoDocSelect(configData.default_config[firstMateria]);
  }

  // Mapeo entre los valores del select y las claves del JSON
  const selectToJsonKeyMap = {
    penal: "penal",
    privacidad: "PRIVACIDAD",
    sin_clasificar: "sin clasificar prueba",
  };

  // Escuchar los cambios en el primer select
  document.getElementById("materia-select").addEventListener("change", function () {
    const selectedMateria = this.value;

    // Buscar la clave correspondiente en el mapeo
    const jsonKey = selectToJsonKeyMap[selectedMateria];

    if (jsonKey) {
      // Ahora que tenemos la clave correcta, buscamos en el JSON
      if (configData.default_config[jsonKey]) {
        populateTipoDocSelect(configData.default_config[jsonKey]);
      } else {
        console.log("La clave no existe en el JSON:", jsonKey);
      }
    } else {
      console.log("No se encontró mapeo para el valor seleccionado:", selectedMateria);
    }
  });
}

// Llena el select de Materia
function populateMateriaSelect(subjectCodes) {
  const materiaSelect = document.getElementById("materia-select");
  materiaSelect.innerHTML = ""; // Limpiar opciones previas

  for (const [name, value] of Object.entries(subjectCodes)) {
    const option = document.createElement("option");
    option.value = value; // Valor que será usado después
    option.textContent = name; // Nombre visible en el select
    materiaSelect.appendChild(option);
  }
}

// Llena el select de Tipo de Documento basado en la Materia seleccionada
function populateTipoDocSelect(documentTypes) {
  const tipoDocSelect = document.getElementById("tipo-doc-select");
  const doctypeSpan = document.getElementById("doctype");

  tipoDocSelect.innerHTML = "";

  if (documentTypes) {
    documentTypes.forEach((doc) => {
      const option = document.createElement("option");
      option.value = doc.code;
      option.textContent = doc.doctype;
      option.dataset.doctype = doc.doctype;
      tipoDocSelect.appendChild(option);
    });
  }

  // actualiza el contenido de doctypeSpan cuando cambia la selección
  tipoDocSelect.addEventListener("change", function () {
    const selectedOption = tipoDocSelect.options[tipoDocSelect.selectedIndex];
    doctypeSpan.textContent = selectedOption.dataset.doctype || "";
  });

  // Establecer el valor inicial del span si se cambia el primer select
  if (tipoDocSelect.options.length > 0) {
    setTimeout(function () {
      doctypeSpan.textContent = tipoDocSelect.options[tipoDocSelect.selectedIndex].dataset.doctype;
    }, 2000);
  }
  //inicializo Revisa
  getMatchGroups();
}

function initializeSelectsFromData(data) {
  const materiaSelect = document.getElementById("materia-select");
  const tipoDocSelect = document.getElementById("tipo-doc-select");

  // Recorrer todas las opciones del primer select
  for (let i = 0; i < materiaSelect.options.length; i++) {
    // Si el valor de la opción coincide con el valor pasado a la función
    if (materiaSelect.options[i].value === data.subject_code) {
      // Seleccionar esa opción
      materiaSelect.selectedIndex = i;
      break;
    }
  }
  const event = document.createEvent("HTMLEvents");
  event.initEvent("change", true, false);
  materiaSelect.dispatchEvent(event);
  // Recorrer todas las opciones del segundo select
  let optionFound = false;

  for (let i = 0; i < tipoDocSelect.options.length; i++) {
    if (tipoDocSelect.options[i].value === data.code) {
      tipoDocSelect.selectedIndex = i;
      optionFound = true;
      break;
    }
  }

  if (!optionFound) {
    const newOption = document.createElement("option");
    newOption.value = "sin_clasificar";
    newOption.text = "sin clasificar";
    tipoDocSelect.appendChild(newOption);
    tipoDocSelect.selectedIndex = tipoDocSelect.options.length - 1;
  }
  typeDocument();
}

//Funcion para despleagr y contraer clasificacion automatica
function toggleDocumentBox() {
  const icon = document.querySelector(".icon-chevron-right");
  const content = document.getElementById("classify-document-body");

  // Alterna la rotación del icono
  icon.classList.toggle("rotated");

  // Expande o colapsa el contenido
  if (content.style.height === "350px") {
    content.style.height = "0"; // Colapsa si ya está expandido
  } else {
    content.style.height = "350px"; // Expande si está colapsado
  }
}
// #####################################################################################
//                           Fin Campos opcion clasificacion automatica
// #####################################################################################

// #####################################################################################
//                             Inicio pestaña del Chat
// #####################################################################################
// Función para enviar el texto desde el input
function sendInputText() {
  const input = document.getElementById("chat-input");
  // const authToken = localStorage.getItem("authTokenMia");

  const userMessage = input.value.trim();

  if (userMessage) {
    // Mostrar el mensaje del usuario en el chat
    const chatBox = document.getElementById("chat-box");

    // Formatear el mensaje del usuario y agregarlo al chat
    const userMessageElement = document.createElement("div");
    userMessageElement.classList.add("message", "user-message");
    userMessageElement.textContent = userMessage;
    chatBox.appendChild(userMessageElement);

    input.value = "";

    scrollToBottom();

    typeDocument().then((docData) => {
      sendTextToServer(userMessage, docData);
    });
  }
}

// Función para enviar el texto del chat a la API.
function sendTextToServer(text, docData) {
  console.log("texto que le llega a sendTextToServer: " + text);

  getDocumentText()
    .then((documentText) => {
      const authToken = getAuthToken();
      hideResponseText();
      showLoadingMessageEnv();
      showTypingIndicator();
      const conversationId = localStorage.getItem("conversation_id") || null;

      const data = {
        question: text,
        conversation_id: conversationId,
        actions: false,
        lang: "es",
        data: null,
        raw_text: documentText,
        meta: {
          doclang: "es",
          generativeClassification: {
            doctype: docData.generativeClassification.doctype,
            subject_code: docData.generativeClassification.subject_code,
            code: docData.generativeClassification.code,
            subject: docData.generativeClassification.subject,
          },
          doccountry: "España",
        },
        country: "España",
        expertMode: false,
      };

      console.log("Payload enviado:", JSON.stringify(data));

      $.ajax({
        url: "https://miadev.miaintelligence.com:444/api/callDocChat",
        type: "POST",
        contentType: "application/json",
        headers: {
          Authorization: `Bearer ${authToken}`,
        },
        data: JSON.stringify(data),
        success: function (response) {
          hideLoadingMessageEnv();
          hideTypingIndicator();
          console.log("Payload recibido:", JSON.stringify(response));
          displayResponseChat(response);
        },
        error: function (xhr, status, error) {
          hideLoadingMessageEnv();
          hideTypingIndicator();
          let errorMessage;
          console.log("Estado del error: ", status); //Lo pongo para qeu no me salte el error de ESLint
          console.log("Estado del error: ", error); //Lo pongo para qeu no me salte el error de ESLint
          if (xhr.status === 401) {
            displayResponseChat(
              JSON.stringify({
                data: {
                  messages: [{ role: "assistant", content: "La sesión ha caducado." }],
                  conversation_id: "session_expired",
                },
              })
            );
            setTimeout(function () {
              showPopupEndOfSession();
            }, 3000);
          } else {
            errorMessage = "Error del servidor.";
          }
          displayResponseText({ message: errorMessage });
        },
        complete: function () {
          hideLoadingMessageEnv();
        },
      });
    })
    .catch((error) => {
      // En caso de que ocurra un error al obtener el texto del documento
      console.error(error);
      displayResponseText({ message: error });
    });
}

function typeDocument() {
  return new Promise((resolve) => {
    // Obtener los select
    const materiaSelect = document.getElementById("materia-select");
    const tipoDocSelect = document.getElementById("tipo-doc-select");
    const doctype = document.getElementById("doctype");

    // Obtener valores seleccionados
    const subject_code = materiaSelect.value;
    let subject = subject_code;
    const tipoDoc = tipoDocSelect.value;

    if (subject) {
      if (subject === "privacidad") {
        subject = "PRIVACIDAD";
      }
      if (subject === "sin_clasificar") {
        subject = "sin clasificar prueba";
      }
    }

    const data = {
      generativeClassification: {
        doctype: doctype.textContent,
        subject_code: subject_code,
        code: tipoDoc,
        subject: subject,
      },
    };

    console.log(JSON.stringify(data));
    console.log("materia: " + subject_code + " Tipo Doc.: " + tipoDoc + " span: " + doctype.textContent);

    // Retornar el JSON como una promesa resuelta
    resolve(data);
  });
}

//Obtiene el documento de las paginas del Word
function getDocumentText() {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      try {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        resolve(body.text);
      } catch (error) {
        reject("Error al obtener el texto del documento: " + error);
      }
    });
  });
}

//Funcion paar mostrar las respuestas en el Chat
function displayResponseChat(response) {
  const chatBox = document.getElementById("chat-box");
  hideTypingIndicator();

  if (typeof response === "string") {
    try {
      response = JSON.parse(response);
    } catch {
      mostrarMensajeError("Respuesta del servidor no es un JSON válido.");
      return;
    }
  }

  if (response && response.data && response.data.messages) {
    if (response.data.conversation_id) {
      localStorage.setItem("conversation_id", response.data.conversation_id);
    }
    const assistantMessage = response.data.messages.reverse().find((message) => message.role === "assistant");

    if (assistantMessage && assistantMessage.content) {
      // Crear y mostrar el mensaje del asistente en el chat
      const botMessageElement = document.createElement("div");
      botMessageElement.classList.add("message", "bot-message");
      botMessageElement.textContent = assistantMessage.content;
      chatBox.appendChild(botMessageElement);
      scrollToBottom();
    } else {
      mostrarMensajeError("No hay contenido en el mensaje del asistente.");
    }
  } else {
    mostrarMensajeError("No response from server.");
  }
}

// Función para mostrar el indicador de "escribiendo"
function showTypingIndicator() {
  const chatBox = document.getElementById("chat-box");

  // Crear el loading
  const typingIndicator = document.createElement("div");
  typingIndicator.classList.add("message", "bot-message", "typing-indicator");
  typingIndicator.innerHTML = `
    <div class="typing typing-1"></div>
    <div class="typing typing-2"></div>
    <div class="typing typing-3"></div>
  `;

  chatBox.appendChild(typingIndicator);

  scrollToBottom();
}

// Función para ocultar el indicador de "escribiendo"
function hideTypingIndicator() {
  const typingIndicator = document.querySelector(".typing-indicator");
  if (typingIndicator) {
    typingIndicator.remove();
  }
}

//Funcion para mostrar en el chat si hay una respuesta de error
function mostrarMensajeError(texto) {
  const chatBox = document.getElementById("chat-box");
  const botMessageElement = document.createElement("div");
  botMessageElement.classList.add("message", "bot-message");
  botMessageElement.textContent = texto;
  chatBox.appendChild(botMessageElement);
  scrollToBottom();
}

// Función para desplazar el chat hacia el último mensaje
function scrollToBottom() {
  const chatBox = document.getElementById("chat-box");
  chatBox.scrollTop = chatBox.scrollHeight;
}

//Función para limpiar el, id de conversación e iniciar una nueva
function newChat() {
  const chatBox = document.getElementById("chat-box");
  localStorage.removeItem("conversation_id");
  chatBox.innerHTML = "";
}

function handleSendAllText() {
  Word.run(async (context) => {
    const document = context.document;
    const selection = document.getSelection();
    selection.load("text");

    const body = context.document.body;
    body.load("text");

    await context.sync();

    let text = selection.text.trim();
    let isFullDocument = false;

    // Si no hay texto seleccionado, se usará el cuerpo del documento
    if (text.length === 0) {
      text = body.text.trim();
      isFullDocument = true;
    }

    if (text.length === 0) {
      displayResponseText({ message: "El documento está vacío o no hay selección." });
      return;
    }

    typeDocument()
      .then((docData) => {
        if (!docData || !docData.generativeClassification) {
          displayResponseText({
            message: "No se encontraron datos adicionales para enviar junto con el texto.",
          });
          return;
        }

        handleSendToChat(text, docData, isFullDocument);
      })
      .catch((error) => {
        console.error("Error al obtener docData:", error);
        displayResponseText({ message: "Error al procesar los datos del documento." });
      });
  });
}

function handleSendToChat(text, docData, isFullDocument) {
  const chatBox = document.getElementById("chat-box");

  const userMessageElement = document.createElement("div");
  userMessageElement.classList.add("message", "user-message");

  if (isFullDocument) {
    userMessageElement.textContent = "Doc."; // Mostrar "Doc." para el documento completo
  } else {
    userMessageElement.textContent = text; // Mostrar texto seleccionado
  }

  chatBox.appendChild(userMessageElement);
  scrollToBottom();

  sendTextToServer(text, docData);
}

// #####################################################################################
//                           Fin pestaña del Chat
// #####################################################################################

// #####################################################################################
//                             Inicio pestaña de Revisa
// #####################################################################################
//capturo el id dinamico asignado al boton de revisa
document.addEventListener("click", function (event) {
  const button = event.target.closest(".btn-revisa.btn-green");

  if (button) {
    const buttonId = button.id;
    console.log("ID del botón clicado:", buttonId);
    findContentByButtonId(buttonId);
  }
});

document.addEventListener("click", function (event) {
  const button = event.target.closest(".correct-act"); // Detectar clic en elemento con clase 'correct-act'

  if (button) {
    const dataCode = button.getAttribute("data-code"); // Obtener el atributo 'data-code'

    if (dataCode && dataCode !== "[]") {
      // Si 'data-code' tiene contenido válido, enviarlo a openPopup
      openOfficeDialog(dataCode);
    } else {
      // Si 'data-code' está vacío o es '[]', enviar mensaje personalizado
      openOfficeDialog("No hay nada para corregir, el documento cumple las normas.");
    }
  }
});

// --------------------------------------------------------------------------------------//

function getMatchGroups() {
  addSpinAnimationRevisa();
  const authToken = getAuthToken();
  const doctypeElement = document.getElementById("tipo-doc-select");
  const doctypeCode = doctypeElement.value;
  console.log("Texto del select: " + doctypeCode);

  const data = {
    doctype_code: doctypeCode,
    country: "España",
    lang: "es",
  };

  console.log("Payload enviado:", JSON.stringify(data));

  $.ajax({
    url: "https://miadev.miaintelligence.com:444/api/getMatchGroups",
    type: "POST",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${authToken}`,
    },
    data: JSON.stringify(data),
    success: function (response) {
      console.log("Match Groups: ", JSON.stringify(response));
      // clearContainers();
      renderRevisa(response);
    },
    error: function (xhr, status, error) {
      // hideLoadingMessageEnv();
      // hideTypingIndicator();
      let errorMessage;
      console.log("Estado del error: ", status);
      console.log("Estado del error: ", error);
      if (xhr.status === 401) {
        errorMessage = xhr.responseJSON.message;
        showPopupEndOfSession();
      } else {
        errorMessage = "Error del servidor.";
      }
      // displayResponseText({ message: errorMessage });
      console.log({ message: errorMessage });
    },
    complete: function () {
      // hideLoadingMessageEnv();
    },
  });
}

function clearContainers() {
  const containerGroups = document.getElementById("qualifier-groups");
  const containerCustomGroups = document.getElementById("qualifier_custom_groups");

  // Vacía los contenedores
  if (containerGroups) {
    containerGroups.innerHTML = "";
  }
  if (containerCustomGroups) {
    containerCustomGroups.innerHTML = "";
  }
}

function displayDocTypeRevisa() {
  const doctypeElement = document.getElementById("doctype");

  if (doctypeElement) {
    removeSpinAnimationRevisa(); // Llama a la función removeSpinAnimation
    const doctypeContent = doctypeElement.textContent || doctypeElement.innerText;

    const responseDocType = document.getElementById("criteria-title");
    const doctypeElement1 = document.getElementById("own-criteria-title");

    if (responseDocType) {
      responseDocType.textContent = doctypeContent;
    } else {
      console.warn("El elemento con ID 'criteria-title' no existe.");
    }

    if (doctypeElement1) {
      doctypeElement1.textContent = doctypeContent;
    } else {
      console.warn("El elemento con ID 'own-criteria-title' no existe.");
    }
  } else {
    console.warn("El elemento con ID 'doctype' no existe.");
  }
}

function toggleSpinAnimationRevisa(action) {
  const elements = [
    { id: "spin-revisa", classes: ["icon-refresh-2", "w-600", "spin-animation"] },
    { id: "refresh-spin-revisa", classes: ["icon-refresh-2", "w-800", "spin-animation"] },
    { id: "refresh-spin-revisa-2", classes: ["icon-refresh-2", "w-800", "spin-animation"] },
    { id: "spin-revisa-1", classes: ["icon-refresh-2", "w-600", "spin-animation"] },
    { id: "refresh-groups", classes: ["spin-animation"] },
    { id: "refresh-groups-1", classes: ["spin-animation"] },
    { id: "sping-refresh", classes: ["spin-animation"] },
  ];

  elements.forEach(({ id, classes }) => {
    const element = document.getElementById(id);
    if (element) {
      element.classList[action](...classes);
    }
  });
}

function addSpinAnimationRevisa() {
  clearContainers();
  toggleSpinAnimationRevisa("add");
}

function removeSpinAnimationRevisa() {
  toggleSpinAnimationRevisa("remove");
}

function renderRevisa(data) {
  console.log("Datos recibidos:", data);
  displayDocTypeRevisa();
  // Selecciona el contenedor principal donde quieres insertar las tarjetas
  const containerGroups = document.getElementById("qualifier-groups");
  const containerCustomGroups = document.getElementById("qualifier_custom_groups");

  containerCustomGroups.innerHTML = "";

  const customGroupsEmpty = `<div class="generative-qualifier-container h-full">
  <div class="accordion">
    <div style="width: 100%; display: flex; justify-content: center;">
      <div class="info-alert alert-icon-flex my-1 text-xs py-2 !p-2 mx-4 ng-star-inserted">
        <span class="icon icon-info-circle1 info-alert-icon"></span>
        <span class="info-alert-text">
          No se han configurado grupos de criterios para: Tipo de criterio + Tipo
          de documento + País.
        </span>
      </div>
    </div>
  </div>
</div>
`;

  const renderGroups = (groups, container, prefix) => {
    if (
      !groups ||
      groups.length === 0 ||
      groups.every((group) => !group.associated_criterias || group.associated_criterias.length === 0)
    ) {
      container.innerHTML = customGroupsEmpty;
      return; // Finaliza la función
    }

    let idCard = 1;

    groups.forEach((group) => {
      const card = document.createElement("div");
      card.className = "card";

      const header = `
       <a class="card-header collapsed collapser">
         <div class="group2">
           <div class="left-content">
             <mat-icon class="mat-icon text-xl text-grey-medium material-icons mat-ligature-font mat-icon-no-color icono-izquierda">filter_list</mat-icon>
             <span class="text-sm !font-semibold text-left bW w-100">${group.title}</span>
           </div>
           <div class="right-content">
             <span class="icon icon-chevron-down"></span>
             <div class="switch-container">
               <input type="checkbox" id="switch2" class="switch-checkbox" checked />
               <label for="switch2" class="switch-label">Toggle 2</label>
             </div>
             <div class="icon-refresh-revisa"><span class="icon icon-refresh"></span></div>
           </div>
         </div>
       </a>
     `;
      card.innerHTML = header;

      const collapse = document.createElement("div");
      collapse.className = "collapse";

      group.associated_criterias.forEach((criteria, index) => {
        let recBoxClass = "rec-grey";
        let iconClass = "icon-minus";
        if (criteria.evaluation === "ko") {
          recBoxClass = "rec-red";
          iconClass = "icon-x";
        } else if (criteria.evaluation === "ok") {
          recBoxClass = "rec-green";
          iconClass = "icon-check";
        }

        const getClassByWeight = (weight, prefix) => {
          const classMap = {
            1: `${prefix}-green`,
            2: `${prefix}-yellow`,
            3: `${prefix}-orange`,
            4: `${prefix}-red`,
          };
          return classMap[weight] || `${prefix}-default`;
        };

        const getTextByWeight = (weight) => {
          const textMap = {
            1: "Bajo",
            2: "Medio",
            3: "Alto",
            4: "Muy Alto",
          };
          return textMap[weight] || "";
        };

        const roundIconClass = getClassByWeight(criteria.weight, "round-ico");
        const badgeBgClass = getClassByWeight(criteria.weight, "badge-bg");
        const badgeTextClass = getClassByWeight(criteria.weight, "badge-text");
        const importanceText = getTextByWeight(criteria.weight);

        const excludeYes = criteria.exclude
          ? `<div class="badge badge-bg-ligth-blue no-select">
             <span class="badge-text-ligth-blue">Excluyente: Sí</span>
           </div>`
          : "";

        const excludeIcon = criteria.exclude
          ? `<mat-icon _ngcontent-ng-c279190276="" class="mat-icon exclude-icon material-icons">warning</mat-icon>`
          : "";

        const uniqueId = `${prefix}submenu-${index}`;

        const esRefsCount = criteria.normative_refs?.es?.refs?.length || 0;
        const esRefs = criteria.normative_refs?.es?.refs || [];

        const normativasCards = esRefs
          .map(
            (ref) => `
             <div class="">
               <div class="alert-norm alert-icon-flex !mt-4 ng-star-inserted">
                 <span class="norm-margin">
                   <p class="font-bold">Normativas</p>
                   <p class="text-xs">${ref.normative_reference}</p>
                 </span>
               </div>
               <div class="danger-alert alert-icon-flex !mt-4 ng-star-inserted">
                 <span class="icon icon-danger danger-alert-icon"></span>
                 <span class="norm-margin">
                   <p class="font-bold">Incumplimiento</p>
                   <p class="text-xs">${ref.threat}</p>
                 </span>
               </div>
             </div>
           `
          )
          .join("");
        const dinamicId = `${prefix}normativas-btn-${idCard}`;
        const dynamicCardId = `${prefix}exclude-card-${idCard}`;
        const dynamicRevisadId = `${prefix}-revisa-${idCard}`;
        const subCard = `
         <div class="card">
           <a class="card-header collapser" data-toggle="collapse" data-target="#${uniqueId}">
               <div class="grupo3-box pb-3">
                   <div class="container-grup3 ms-2">
                       <span class="rec-box ${recBoxClass}">
                           <span class="icon ${iconClass}"></span>
                       </span>
                       <span class="text-left">
                       ${criteria.title}
                       </span>
                   </div>
                   <div class="right-icons">
                       ${excludeIcon}
                       <div class="round-ico ${roundIconClass}"></div>
                       <span class="icon icon-chevron-down ms-1"></span>
                   </div>
               </div>
           </a>
           <div id="${uniqueId}" class="collapse">
             <div class="card-body sub px-3" data-code="${criteria.code}">
              <div class="integrity-hash" data-code="${criteria.integrity_hash}"></div>
              <div class="exclude-class" data-code="${criteria.exclude}"></div>
               <span class="span-sub-content" data-code="${criteria.description}">${criteria.description}</span>

               <div class="container-badge">
                 <div class="badge ${badgeBgClass} no-select">
                   <div class="round-ico-badge ${roundIconClass}"></div>
                   <span id="criteria-importance" class="${badgeTextClass}">Importancia: ${importanceText}</span>
                 </div>
                 <div id="${dinamicId}" class="badge badge-bg-ligth-blue no-select cursor-pointer" onclick="toggleCard(${idCard}, '${prefix}')">
                   <span class="badge-text-ligth-blue weight" data-code="${criteria.weight}">NORMATIVAS: ${esRefsCount}</span>
                 </div>
                 ${excludeYes}
               </div>

                <div id="${dynamicCardId}" class="hidden">
                 ${normativasCards}
                </div>
               <div>
                   <div class="container-btn">
                       <span id="${dynamicRevisadId}" class="btn-revisa btn-green no-select no-close cursor-pointer">
                       <span class="icon icon-refresh icon-white"></span>Revisar</span>
                       <span class="btn-revisa btn-grey no-select no-close correct">
                       <span class="icon icon-task-square icon-white"></span>Corregir</span>
                   </div>
               </div>
             </div>
           </div>
         </div>
       `;
        collapse.innerHTML += subCard;
        idCard++;
      });

      card.appendChild(collapse);
      container.appendChild(card);
    });
  };

  // Renderiza ambos grupos
  renderGroups(data.data.qualifier_groups, containerGroups, "group");
  renderGroups(data.data.qualifier_custom_groups, containerCustomGroups, "customGroup");
}

// Función para enviar la solicitud y procesar la respuesta
// function sendReviewRequest(authToken, text) {
//   // Mostrar el indicador de carga mientras se espera la respuesta
//   authToken = "12345";
//   showLoadingRevisa();
//   $.ajax({
//     url: "https://servidor-complemento.onrender.com/api/revisa",
//     type: "POST",
//     contentType: "application/json",
//     data: JSON.stringify({ authToken: authToken, text: text }),
//     success: function (response) {
//       hideLoadingRevisa(); // Ocultar el indicador de carga
//       // Pintar la respuesta en el div res-revisar
//       jsonResponseTextRevisa(response);
//     },
//     error: function (xhr, status, error) {
//       hideLoadingRevisa(); // Ocultar el indicador de carga en caso de error
//       let errorMessage = "Error del servidor.";
//       console.log("Estado del error: ", status);
//       console.log("Estado del error: ", error);
//       if (xhr.status === 401) {
//         errorMessage = xhr.responseJSON.message;
//         showPopupEndOfSession();
//       }
//       displayResponseText({ message: errorMessage });
//     },
//     complete: function () {
//       hideLoadingRevisa();
//     },
//   });
// }

// Mueve la declaración de la función fuera del cuerpo principal de la función jsonResponseTextRevisa
// function updateCriterionDetails(selectedCriterion) {
//   const criteriaDescription = document.getElementById("criteria-description");
//   const criteriaImportance = document.getElementById("criteria-importance");
//   const criteriaNorms = document.getElementById("criteria-norms");
//   const boxBadge = document.getElementById("box-badge");

//   // Actualizamos los detalles del criterio seleccionado
//   criteriaDescription.innerText = selectedCriterion.description;
//   criteriaImportance.innerText = `Importancia: ${selectedCriterion.importance}`;
//   criteriaNorms.innerText = `NORMATIVAS: ${selectedCriterion.associatedNorms.length}`;

//   const importance = selectedCriterion.importance.toLowerCase();

//   // Remover las clases anteriores antes de aplicar las nuevas
//   criteriaImportance.classList.remove("badge-text-red", "badge-text-orange", "badge-text-green");
//   boxBadge.classList.remove("badge-bg-red", "badge-bg-orange", "badge-bg-green");

//   // Cambiar las clases del badge dependiendo de la importancia
//   if (importance === "alta") {
//     criteriaImportance.classList.add("badge-text-red");
//     boxBadge.classList.add("badge-bg-red");
//   } else if (importance === "media") {
//     criteriaImportance.classList.add("badge-text-orange");
//     boxBadge.classList.add("badge-bg-orange");
//   } else if (importance === "baja") {
//     criteriaImportance.classList.add("badge-text-green");
//     boxBadge.classList.add("badge-bg-green");
//   }
// }

// function jsonResponseTextRevisa(response) {
//   const responseContainer = document.getElementById("response-text-revisa");
//   const groupNameContainer = document.getElementById("group-name");
//   const criteriaSelect = document.getElementById("criteria-select");

//   // Limpiamos el select de criterios antes de agregar nuevos
//   criteriaSelect.innerHTML = "";

//   if (response && response.reviewedCriteria) {
//     const criteriaMap = new Map(); // Usar Map para búsquedas rápidas de criterios

//     // Crear un fragmento de documento para agregar todas las opciones de una vez
//     const fragment = document.createDocumentFragment();

//     // Recorrer cada "group" dentro de los criterios legales o técnicos
//     response.reviewedCriteria.forEach((criteriaType) => {
//       criteriaType.groups.forEach((group) => {
//         // Mostrar el nombre del grupo en el h3
//         responseContainer.innerText = response.documentType; // Tipo de documento
//         groupNameContainer.innerText = group.group; // Nombre del grupo

//         group.criteria.forEach((criterion) => {
//           // Guardar cada criterio en el mapa usando el título como clave
//           criteriaMap.set(criterion.title, criterion);

//           // Crear una nueva opción para cada criterio
//           const option = document.createElement("option");
//           option.value = criterion.title;
//           option.text = criterion.title;
//           fragment.appendChild(option);
//         });
//       });
//     });

//     // Agregar el fragmento completo al DOM de una sola vez
//     criteriaSelect.appendChild(fragment);

//     // Inicialmente, mostrar los detalles del primer criterio del select
//     if (criteriaMap.size > 0) {
//       const firstCriterion = criteriaMap.values().next().value;
//       criteriaSelect.value = firstCriterion.title; // Selecciona el primer criterio por defecto
//       updateCriterionDetails(firstCriterion); // Muestra los detalles del primer criterio
//     }

//     // Listener para cuando cambia la selección en el select
//     criteriaSelect.addEventListener("change", function () {
//       const selectedTitle = this.value;
//       const selectedCriterion = criteriaMap.get(selectedTitle);

//       if (selectedCriterion) {
//         // Actualizar la UI con los detalles del criterio seleccionado
//         updateCriterionDetails(selectedCriterion);
//       } else {
//         console.error("Criterio seleccionado no encontrado.");
//       }
//     });
//   } else {
//     responseContainer.innerText = "No response from server.";
//   }
// }

// function toggleContent() {
//   const content = document.getElementById("hiddenContent");
//   content.classList.toggle("expanded");
//   if (content.classList.contains("expanded")) {
//     content.style.height = content.scrollHeight + "px";
//   } else {
//     // toggleSubContent();
//     content.style.height = "0";
//     // Cierra todos los submenús si el contenedor principal se colapsa
//     document.querySelectorAll(".menu1.expanded").forEach((subContent) => {
//       subContent.classList.remove("expanded");
//       subContent.style.height = "0";
//     });
//   }
// }

// function toggleContent2() {
//   const content = document.getElementById("hiddenContent2");
//   content.classList.toggle("expanded");

//   if (content.classList.contains("expanded")) {
//     content.style.height = content.scrollHeight + "px";
//   } else {
//     content.style.height = "0";

//     // Cierra todos los submenús si el contenedor principal se colapsa
//     document.querySelectorAll(".menu2.expanded").forEach((subContent) => {
//       subContent.classList.remove("expanded");
//       subContent.style.height = "0";
//     });
//   }
// }

// function toggleSubContent2() {
//   const subContent = document.getElementById("subHiddenContent");
//   subContent.classList.toggle("expanded");

//   if (subContent.classList.contains("expanded")) {
//     subContent.style.height = subContent.scrollHeight + "px";
//   } else {
//     subContent.style.height = "0";
//   }
// }

// #####################################################################################
//                           Fin pestaña de Revisa
// #####################################################################################

// #####################################################################################
//                             Inicio Manejo de pestañas
// #####################################################################################
//Funcion de manejo de pestañas
function setupTabs() {
  const tabs = document.querySelectorAll(".tabs-nav li a");
  const sections = document.querySelectorAll(".tabs-stage > div");

  // Mostrar el contenido del primer tab por defecto
  sections.forEach((section) => (section.style.display = "none")); // Ocultamos todas las secciones
  sections[0].style.display = "block"; // Mostramos el contenido del primer tab

  // Añadir un evento de clic a cada pestaña
  tabs.forEach((tab) => {
    tab.addEventListener("click", function (event) {
      event.preventDefault();

      // Desactivar la pestaña activa actual
      document.querySelector(".tabs-nav li.tab-active").classList.remove("tab-active");

      // Activar la pestaña clicada
      this.parentElement.classList.add("tab-active");

      // Ocultar todas las secciones
      sections.forEach((section) => (section.style.display = "none"));

      // Mostrar la sección correspondiente a la pestaña clicada
      const target = this.getAttribute("href");
      document.querySelector(target).style.display = "block";
    });
  });
}

// #####################################################################################
//                             Fin Manejo de pestañas
// #####################################################################################

// #####################################################################################
//                             Inicio de login
// #####################################################################################

function validateAndSubmitLogin() {
  const emailUser = document.getElementById("emailUser").value.trim();
  const passwordUser = document.getElementById("passwordUser").value.trim();
  const errorContainer = document.getElementById("error-message-email");
  const errorPassword = document.getElementById("error-message-password");

  errorContainer.style.display = "none";
  errorPassword.style.display = "none";

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  if (emailUser.length === 0 || !emailRegex.test(emailUser)) {
    errorContainer.innerText = "El campo email requiere un formato válido (por ejemplo, xxx@xxx.xxx)";
    errorContainer.style.display = "block";
    return;
  }

  if (passwordUser.length === 0) {
    errorPassword.innerText = "Por favor, ingrese una contraseña.";
    errorPassword.style.display = "block";
    return;
  }

  loginProcess(emailUser, passwordUser);
  // hideAuthTokenDialog();
}

function loginProcess(emailUser, passwordUser) {
  hideResponseSession();
  showLoadingMessage();
  $.ajax({
    // url: "http://localhost:3001/api/authtoken",
    // url: "https://servidor-complemento.onrender.com/api/authtoken",
    url: "https://miadev.miaintelligence.com:444/api/auth/logoutFromOtherSessionsAndLogin",
    type: "POST",
    contentType: "application/json",
    data: JSON.stringify({
      email: emailUser,
      password: passwordUser,
      remember_me: false,
    }),
    success: function (response) {
      hideLoadingMessage();

      if (response.data && response.data.original && response.data.original.message) {
        // Si hay un mensaje de error, lo mostramos
        const errorMessage = response.data.original.message;
        const errorContainer = document.getElementById("error-message-server");
        errorContainer.innerText = errorMessage;
        errorContainer.style.display = "block";
      } else {
        // Si no hay mensaje de error, intentamos obtener el access_token
        const accessToken = response.data?.access_token;

        if (accessToken) {
          localStorage.setItem("authTokenMia", accessToken);
          newChat();
          displayAuthResponse();
          getGenerativeClassifierConfig();
          setTimeout(function () {
            document.getElementById("classify-document").style.display = "block";
          }, 2100);
        } else {
          // Mensaje de error genérico si no hay token (aunque no debería ocurrir con credenciales válidas)
          console.error("Token de acceso no encontrado en la respuesta.");
          const errorContainer = document.getElementById("error-message-server");
          errorContainer.innerText = "Error: No se pudo obtener el token de acceso.";
          errorContainer.style.display = "block";
        }
      }
    },
    error: function (xhr) {
      hideLoadingMessage();
      console.log("Estado del xhr ", xhr);
      const errorContainer = document.getElementById("error-message-server");
      errorContainer.innerText = "Error del servidor.";
      errorContainer.style.display = "block";
    },
    complete: function () {
      hideLoadingMessage();
    },
  });
}

function displayAuthResponse() {
  hideLoadingMessage();
  const responseContainer = document.getElementById("response-token");
  showResponseToken();
  responseContainer.innerText = "Autenticado con éxito";
  setTimeout(function () {
    hideAuthToken();
    showTabs();
  }, 2000);
}

// #####################################################################################
//                             Fin Login
// #####################################################################################

// #####################################################################################
//                             Inicio de Util
// #####################################################################################
// Función para mostrar la respuesta en el panel
function displayResponseText(response) {
  const responseContainer = document.getElementById("response-text");
  if (response && response.message) {
    responseContainer.innerText = response.message;
  } else {
    responseContainer.innerText = "No response from server.";
  }
}
//funcion para obtener el token del localStore
function getAuthToken() {
  return localStorage.getItem("authTokenMia");
}

// funcion para mostrar la respuesta en la pestaña Extrae
function hideResponseText() {
  document.getElementById("response-text").innerText = "";
}

//funcion para limpiar el token y reiniciar el complemento
function handleDeletToken() {
  localStorage.removeItem("authTokenMia");
  document.getElementById("response-delete-token").innerText = "Token eliminado";

  setTimeout(function () {
    window.location.reload();
  }, 2000);
}

//funcion para limpiar el token y reiniciar el complemento
function restartPlugin() {
  localStorage.removeItem("authTokenMia");

  setTimeout(function () {
    window.location.reload();
  }, 2000);
}

// Función para obtener el texto
// function getWordText() {
//   Word.run(async (context) => {
//     const document = context.document;
//     const selection = document.getSelection();
//     selection.load("text");

//     const body = context.document.body;
//     body.load("text");

//     await context.sync();

//     const authToken = localStorage.getItem("authTokenMia");
//     let textToSend = selection.text.trim();

//     if (textToSend.length === 0) {
//       textToSend = body.text.trim();
//     }

//     if (textToSend.length === 0) {
//       displayResponseTextRevisa({ message: "El documento está vacío.<br> o no hay selección." });
//     } else {
//       displayResponseCleanTextRevisa();
//       sendReviewRequest(authToken, textToSend);
//     }
//   });
// }

//Show section

function showPopupEndOfSession() {
  document.getElementById("popup").style.display = "flex";
}

function showLoadingMessage() {
  document.getElementById("loadingMessage").style.display = "block";
}

function showLoadingMessageEnv() {
  document.getElementById("loadingMessageEnv").style.display = "block";
}

// function showLoadingRevisa() {
//   document.getElementById("loadingRevisa").style.display = "block";
// }

function showAuthToken() {
  document.getElementById("auth-token").style.display = "block";
  document.getElementById("bg-color").style.display = "flex";
}

function showTabs() {
  document.getElementById("tabs").style.display = "block";
  document.getElementById("tabs-stage").style.display = "block";
}

function showResponseToken() {
  document.getElementById("response-token").style.display = "block";
}

// function showResponseSession() {
//   document.getElementById("response-session").style.display = "block";
// }

// Hiden section

// function hidePopupEndOfSession() {
//   document.getElementById("popup").style.display = "none";
// }

function hideLoadingIndex() {
  document.getElementById("loading-index").style.display = "none";
}

function hideLoadingMessage() {
  document.getElementById("loadingMessage").style.display = "none";
}

function hideLoadingMessageEnv() {
  document.getElementById("loadingMessageEnv").style.display = "none";
}

// function hideLoadingRevisa() {
//   document.getElementById("loadingRevisa").style.display = "none";
// }

function hideAuthToken() {
  document.getElementById("auth-token").style.display = "none";
  document.getElementById("bg-color").style.display = "none";
}

function hideResponseSession() {
  document.getElementById("response-session").style.display = "none";
}

// Viejos

// function handleSendText() {
//   Word.run(async (context) => {
//     const selection = context.document.getSelection();
//     selection.load("text");

//     await context.sync();

//     const authToken = localStorage.getItem("authTokenMia");
//     const selectedText = selection.text.trim();

//     if (selectedText.length === 0) {
//       displayResponseText({ message: "No hay texto seleccionado." });
//     } else {
//       sendTextToServer(authToken, selectedText);
//     }
//   });
// }

// function errorMessage() {
//   document.getElementById("error-message").style.display = "none";
// }

//End Hiden section
// #############################################################################################
// unauthorizedToken
// function showUnauthorizedToken() {
//   document.getElementById("auth-token").style.display = "block";
//   document.getElementById("bg-color").style.display = "flex";
//   // showClassifyDocument();
//   // document.getElementById("classify-document").style.display = "block";
// }

// function hideUnauthorizedToken() {
//   document.getElementById("auth-token").style.display = "none";
//   document.getElementById("bg-color").style.display = "none";
//   document.getElementById("tabs").style.display = "none";
//   document.getElementById("tabs-stage").style.display = "none";
// }

//End Show section
// function showClassifyDocument() {
//   const element = document.getElementById("classify-document");
//   element.classList.remove("display-none");
// }

// function hideClassifyDocument() {
//   const element = document.getElementById("classify-document");
//   if (element) {
//     element.classList.add("display-none");
//   }
// }

// function validAuthToken(authToken) {
//   showLoadingMessage();
//   $.ajax({
//     // url: "http://localhost:3001/api/validtoken",
//     url: "https://servidor-complemento.onrender.com/api/authtoken",
//     type: "POST",
//     contentType: "application/json",
//     data: JSON.stringify({ authToken: authToken }),
//     success: function (response) {
//       hideLoadingMessage();
//       displayResponseVilidToken(response);

//       hideLoadingIndex();
//       showTabs();
//     },
//     error: function (xhr, status, error) {
//       hideLoadingMessage();
//       const errorContainer = document.getElementById("error-message");
//       let errorMessage;

//       if (xhr.status === 401) {
//         errorMessage = xhr.responseJSON.message;
//         // showAuthTokenDialog();
//       } else {
//         errorMessage = "Error del servidor.";
//       }

//       errorContainer.innerText = errorMessage;
//       errorContainer.style.display = "block";
//     },
//     complete: function () {
//       hideLoadingMessage();
//     },
//   });
//   hideLoadingIndex();
//   showAuthToken();
// }

// function displayResponseVilidToken(response) {
//   const responseContainer = document.getElementById("response-token");
//   if (response && response.message) {
//     responseContainer.innerText = response.message;
//     errorMessage();
//     showResponseToken();
//     if (response.message === "token autenticado con exito") {
//       hideAuthToken();
//       showTabs();
//     }
//   } else {
//     responseContainer.innerText = "No response from server.";
//   }
// }

// function displayResponse(response) {
//   const responseContainer = document.getElementById("response-token");
//   if (response && response.message) {
//     responseContainer.innerText = response.message;
//     errorMessage();
//     showResponseToken();
//     if (response.message === "token autenticado con exito") {
//       setTimeout(function () {
//         hideAuthToken();
//         showTabs();
//       }, 2000);
//     }
//   } else {
//     responseContainer.innerText = "No response from server.";
//   }
// }

// function hideAuthTokenDialog() {
//   document.getElementById("response-token").innerText = " ";
//   document.getElementById("authTokenInput").value = "";
//   document.getElementById("error-message").innerText = " ";
// }

// function unauthorizedToken() {
//   localStorage.removeItem("authTokenMia");
//   hideUnauthorizedToken();
//   showUnauthorizedToken();
// }

// function displayResponseChat(response) {
//   const chatBox = document.getElementById("chat-box");

//   hideTypingIndicator();
//   if (response && response.message) {
//     // Crear un nuevo elemento para la respuesta del servidor
//     const botMessageElement = document.createElement("div");
//     botMessageElement.classList.add("message", "bot-message");
//     botMessageElement.textContent = response.message;

//     chatBox.appendChild(botMessageElement);

//     scrollToBottom();
//   } else {
//     // En caso de que no haya una respuesta válida del servidor
//     const botMessageElement = document.createElement("div");
//     botMessageElement.classList.add("message", "bot-message");
//     botMessageElement.textContent = "No response from server.";

//     chatBox.appendChild(botMessageElement);
//     scrollToBottom();
//   }
// }

// Función para procesar y renderizar la respuesta
// function renderReviewResponse(data) {
//   const container = $("#res-revisar");
//   container.empty(); // Limpiar el contenido anterior

//   // Crear estructura HTML de la respuesta
//   let html = `<h2>${data.documentType}</h2>`;

//   // Iterar por cada criterio revisado
//   data.reviewedCriteria.forEach(criteriaType => {
//     html += `<h3>${criteriaType.type}</h3>`; // Tipo de revisión (Legal, Técnico, etc.)

//     criteriaType.groups.forEach(group => {
//       html += `<h4>${group.group}</h4>`; // Grupo (Responsabilidades, Confidencialidad, etc.)

//       group.criteria.forEach(criterion => {
//         html += `
//           <div class="criterion">
//             <h5>${criterion.title}</h5>
//             <p><strong>Descripción:</strong> ${criterion.description}</p>
//             <p><strong>Importancia:</strong> ${criterion.importance}</p>
//             <p><strong>Resultado del análisis:</strong> ${criterion.analysisResult}</p>
//             <p><strong>Normas asociadas:</strong> ${criterion.associatedNorms.join(', ')}</p>
//             <p><strong>Propuestas de corrección:</strong> ${criterion.correctionProposals.length > 0 ? criterion.correctionProposals.join(', ') : 'Ninguna'}</p>
//           </div>
//         `;
//       });
//     });
//   });

//   // Agregar el resumen global al final
//   html += `
//     <div class="global-review-result">
//       <h3>Resultado Global</h3>
//       <p><strong>Visa Jurídica:</strong> ${data.globalReviewResult.legalVisa}</p>
//       <p><strong>Porcentaje de cumplimiento:</strong> ${data.globalReviewResult.compliancePercentage}%</p>
//     </div>
//   `;

//   // Añadir la clase container-shadow
//   container.html(html).addClass("container-shadow");
// }
// Función para limpiar la respuesta en el panel
// function displayResponseCleanTextRevisa() {
//   const responseContainer = document.getElementById("response-text-revisa");
//   responseContainer.innerHTML = "";
// }

// // Función para mostrar la respuesta en el panel
// function displayResponseTextRevisa(response) {
//   const responseContainer = document.getElementById("response-text-revisa");
//   if (response && response.message) {
//     responseContainer.innerHTML = response.message;
//   } else {
//     responseContainer.innerText = "No response from server.";
//   }
// }
