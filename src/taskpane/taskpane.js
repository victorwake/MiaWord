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

    // document.getElementById("enviar-texto").onclick = handleSendText;
    document.getElementById("clean-document").onclick = newChat;
    document.getElementById("revisa-btn").onclick = getWordText;
    document.getElementById("document-forward").addEventListener("click", handleSendAllText);
    document.getElementById("send-btn").onclick = sendInputText;
    document.getElementById("chat-input").onkeypress = (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        sendInputText();
      }
    };
    getGenerativeClassifierConfig();
    setupTabs();
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

    if (!authToken) {
      hideLoadingIndex();
      showAuthToken();
    } else {
      //validAuthToken(authToken);
      isAuthTokenAvailable(authToken);
    }
    await context.sync();
  });
}

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

// Hiden section
function hideLoadingIndex() {
  document.getElementById("loading-index").style.display = "none";
}

function hideLoadingMessage() {
  document.getElementById("loadingMessage").style.display = "none";
}

function hideLoadingMessageEnv() {
  document.getElementById("loadingMessageEnv").style.display = "none";
}

function hideLoadingRevisa() {
  document.getElementById("loadingRevisa").style.display = "none";
}

function hideAuthToken() {
  document.getElementById("auth-token").style.display = "none";
  document.getElementById("bg-color").style.display = "none";
}

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

//Show section
function showLoadingMessage() {
  document.getElementById("loadingMessage").style.display = "block";
}

function showLoadingMessageEnv() {
  document.getElementById("loadingMessageEnv").style.display = "block";
}

function showLoadingRevisa() {
  document.getElementById("loadingRevisa").style.display = "block";
}

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

function hideResponseSession() {
  document.getElementById("response-session").style.display = "none";
}
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

// #############################################################################################

//Funcion para las pestañas
// #############################################################################################
function setupTabs() {
  // Seleccionamos todos los enlaces dentro de las pestañas
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

//Funcion de llamado a la API
// #############################################################################################
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

function hideResponseText() {
  document.getElementById("response-text").innerText = "";
}

function handleDeletToken() {
  localStorage.removeItem("authTokenMia");
  document.getElementById("response-delete-token").innerText = "Token eliminado";

  setTimeout(function () {
    window.location.reload();
  }, 2000);
}

// function unauthorizedToken() {
//   localStorage.removeItem("authTokenMia");
//   hideUnauthorizedToken();
//   showUnauthorizedToken();
// }

function getAuthToken() {
  return localStorage.getItem("authTokenMia");
}

function newChat() {
  const chatBox = document.getElementById("chat-box");
  localStorage.removeItem("conversation_id");
  chatBox.innerHTML = "";
}

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

function handleSendAllText() {
  Word.run(async (context) => {
    const document = context.document;
    const selection = document.getSelection();
    selection.load("text");

    const body = context.document.body;
    body.load("text");

    await context.sync();

    let text = selection.text.trim();
    // console.log(text);

    // Si no hay texto seleccionado, se enviará todo el documento
    if (text.length === 0) {
      text = body.text.trim();
    }

    if (text.length === 0) {
      displayResponseText({ message: "El documento está vacío o no hay selección." });
    } else {
      sendTextToServer(text);
    }
  });
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

// Mueve la declaración de la función fuera del cuerpo principal de la función jsonResponseTextRevisa
function updateCriterionDetails(selectedCriterion) {
  const criteriaDescription = document.getElementById("criteria-description");
  const criteriaImportance = document.getElementById("criteria-importance");
  const criteriaNorms = document.getElementById("criteria-norms");
  const boxBadge = document.getElementById("box-badge");

  // Actualizamos los detalles del criterio seleccionado
  criteriaDescription.innerText = selectedCriterion.description;
  criteriaImportance.innerText = `Importancia: ${selectedCriterion.importance}`;
  criteriaNorms.innerText = `NORMATIVAS: ${selectedCriterion.associatedNorms.length}`;

  const importance = selectedCriterion.importance.toLowerCase();

  // Remover las clases anteriores antes de aplicar las nuevas
  criteriaImportance.classList.remove("badge-text-red", "badge-text-orange", "badge-text-green");
  boxBadge.classList.remove("badge-bg-red", "badge-bg-orange", "badge-bg-green");

  // Cambiar las clases del badge dependiendo de la importancia
  if (importance === "alta") {
    criteriaImportance.classList.add("badge-text-red");
    boxBadge.classList.add("badge-bg-red");
  } else if (importance === "media") {
    criteriaImportance.classList.add("badge-text-orange");
    boxBadge.classList.add("badge-bg-orange");
  } else if (importance === "baja") {
    criteriaImportance.classList.add("badge-text-green");
    boxBadge.classList.add("badge-bg-green");
  }
}

function jsonResponseTextRevisa(response) {
  const responseContainer = document.getElementById("response-text-revisa");
  const groupNameContainer = document.getElementById("group-name");
  const criteriaSelect = document.getElementById("criteria-select");

  // Limpiamos el select de criterios antes de agregar nuevos
  criteriaSelect.innerHTML = "";

  if (response && response.reviewedCriteria) {
    const criteriaMap = new Map(); // Usar Map para búsquedas rápidas de criterios

    // Crear un fragmento de documento para agregar todas las opciones de una vez
    const fragment = document.createDocumentFragment();

    // Recorrer cada "group" dentro de los criterios legales o técnicos
    response.reviewedCriteria.forEach((criteriaType) => {
      criteriaType.groups.forEach((group) => {
        // Mostrar el nombre del grupo en el h3
        responseContainer.innerText = response.documentType; // Tipo de documento
        groupNameContainer.innerText = group.group; // Nombre del grupo

        group.criteria.forEach((criterion) => {
          // Guardar cada criterio en el mapa usando el título como clave
          criteriaMap.set(criterion.title, criterion);

          // Crear una nueva opción para cada criterio
          const option = document.createElement("option");
          option.value = criterion.title;
          option.text = criterion.title;
          fragment.appendChild(option);
        });
      });
    });

    // Agregar el fragmento completo al DOM de una sola vez
    criteriaSelect.appendChild(fragment);

    // Inicialmente, mostrar los detalles del primer criterio del select
    if (criteriaMap.size > 0) {
      const firstCriterion = criteriaMap.values().next().value;
      criteriaSelect.value = firstCriterion.title; // Selecciona el primer criterio por defecto
      updateCriterionDetails(firstCriterion); // Muestra los detalles del primer criterio
    }

    // Listener para cuando cambia la selección en el select
    criteriaSelect.addEventListener("change", function () {
      const selectedTitle = this.value;
      const selectedCriterion = criteriaMap.get(selectedTitle);

      if (selectedCriterion) {
        // Actualizar la UI con los detalles del criterio seleccionado
        updateCriterionDetails(selectedCriterion);
      } else {
        console.error("Criterio seleccionado no encontrado.");
      }
    });
  } else {
    responseContainer.innerText = "No response from server.";
  }
}

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
// ##################################################################################################
// Función para obtener el texto
function getWordText() {
  Word.run(async (context) => {
    const document = context.document;
    const selection = document.getSelection();
    selection.load("text");

    const body = context.document.body;
    body.load("text");

    await context.sync();

    const authToken = localStorage.getItem("authTokenMia");
    let textToSend = selection.text.trim();

    if (textToSend.length === 0) {
      textToSend = body.text.trim();
    }

    if (textToSend.length === 0) {
      displayResponseTextRevisa({ message: "El documento está vacío.<br> o no hay selección." });
    } else {
      displayResponseCleanTextRevisa();
      sendReviewRequest(authToken, textToSend);
    }
  });
}

// Función para limpiar la respuesta en el panel
function displayResponseCleanTextRevisa() {
  const responseContainer = document.getElementById("response-text-revisa");
  responseContainer.innerHTML = "";
}

// Función para mostrar la respuesta en el panel
function displayResponseTextRevisa(response) {
  const responseContainer = document.getElementById("response-text-revisa");
  if (response && response.message) {
    responseContainer.innerHTML = response.message;
  } else {
    responseContainer.innerText = "No response from server.";
  }
}
// ##################################################################################################

// #####################################################################################
//                             Clasificacion automatica
// #####################################################################################
function classifyGenerative(doc) {
  console.log("Texto que le llega a sendDocToServer: " + doc);
  const authToken = getAuthToken();

  const data = {
    doc: doc,
  };

  console.log("Payload enviado:", JSON.stringify(data));

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

    let doc = body.text.trim();

    if (doc.length === 0) {
      removeSpinAnimation();
      displayResponseDocType({ message: "Documento vacío." });
    } else {
      classifyGenerative(doc);
    }
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
  tipoDocSelect.innerHTML = ""; // Limpiar opciones previas

  if (documentTypes) {
    documentTypes.forEach((doc) => {
      const option = document.createElement("option");
      option.value = doc.code; // Código del documento
      option.textContent = doc.doctype; // Nombre del tipo de documento
      tipoDocSelect.appendChild(option);
    });
  }
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
}
// #####################################################################################
//                           Fin Campos opcion clasificacion automatica
// #####################################################################################

// #####################################################################################
//                             Inicio Prestaña del Chat
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

    sendTextToServer(userMessage);
  }
}

// Función para enviar el texto del chat a la API.
function sendTextToServer(text) {
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
            doctype: "Contrato encargo tratamiento (ES)",
            subject_code: "privacidad",
            code: "cont_encargo",
            subject: "PRIVACIDAD",
          },
          doccountry: "España",
          kind: "conversation",
          path: "",
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
              window.location.reload();
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
// #####################################################################################
//                           Fin Prestaña del Chat
// #####################################################################################

// #####################################################################################
//                             Inicio Prestaña de Revisa
// #####################################################################################

// Función para enviar la solicitud y procesar la respuesta
function sendReviewRequest(authToken, text) {
  // Mostrar el indicador de carga mientras se espera la respuesta
  showLoadingRevisa();
  $.ajax({
    url: "https://servidor-complemento.onrender.com/api/revisa",
    type: "POST",
    contentType: "application/json",
    data: JSON.stringify({ authToken: authToken, text: text }),
    success: function (response) {
      hideLoadingRevisa(); // Ocultar el indicador de carga
      // Pintar la respuesta en el div res-revisar
      jsonResponseTextRevisa(response);
    },
    error: function (xhr, status, error) {
      hideLoadingRevisa(); // Ocultar el indicador de carga en caso de error
      let errorMessage = "Error del servidor.";
      console.log("Estado del error: ", status);
      console.log("Estado del error: ", error);
      if (xhr.status === 401) {
        errorMessage = xhr.responseJSON.message;
      }
      displayResponseText({ message: errorMessage });
    },
    complete: function () {
      hideLoadingRevisa();
    },
  });
}

// Función para mostrar la respuesta en el panel
function displayResponseText(response) {
  const responseContainer = document.getElementById("response-text");
  if (response && response.message) {
    responseContainer.innerText = response.message;
  } else {
    responseContainer.innerText = "No response from server.";
  }
}
