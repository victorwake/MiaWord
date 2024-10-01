/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
/* global localStorage console $ */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");

    await context.sync();

    const selectedText = selection.text;
    const authToken = localStorage.getItem("authToken");
    // const authToken = "1234"

    if (selectedText) {
      sendTextToServer(authToken, selectedText);
    } else {
      document.getElementById("response-container").innerText = "No text selected.";
    }
  });
}

function sendTextToServer(authToken, text) {
  console.log("Enviando authToken:", authToken);
  showLoadingMessage();
  $.ajax({
    url: "https://servidor-complemento.onrender.com/api/text",
    type: "POST",
    contentType: "application/json",
    data: JSON.stringify({ text: text, authToken: authToken }),
    success: function (response) {
      hideLoadingMessage();
      displayResponse(response);
    },
    error: function (xhr, status, error) {
      hideLoadingMessage();
      let errorMessage;
      if (xhr.status === 401) {
        errorMessage = xhr.responseJSON.message;
        showAuthTokenDialog();
      } else {
        errorMessage = "Error del servidor.";
      }
      displayResponse({ message: errorMessage });
    },
    complete: function () {
      hideLoadingMessage();
    },
  });
}

function sendAuthToken(authToken) {
  showLoadingMessage();
  $.ajax({
    url: "https://servidor-complemento.onrender.com/api/authtoken",
    type: "POST",
    contentType: "application/json",
    data: JSON.stringify({ authToken: authToken }),
    success: function (response) {
      hideLoadingMessage();
      displayResponse(response);
    },
    error: function (xhr, status, error) {
      hideLoadingMessage();
      const errorContainer = document.getElementById("error-message");
      let errorMessage;

      if (xhr.status === 401) {
        errorMessage = xhr.responseJSON.message;
        showAuthTokenDialog();
      } else {
        errorMessage = "Error del servidor.";
      }

      errorContainer.innerText = errorMessage;
      errorContainer.style.display = "block";
    },
    complete: function () {
      hideLoadingMessage();
    },
  });
}
function hideLoadingMessage() {
  document.getElementById("loadingMessage").style.display = "none";
}

function showLoadingMessage() {
  document.getElementById("loadingMessage").style.display = "block";
}

function showAuthTokenDialog() {
  document.getElementById("auth-dialog").style.display = "block";
}
document.getElementById("submitAuthToken").onclick = handleSubmitAuthToken;

function hideAuthTokenDialog() {
  document.getElementById("auth-dialog").style.display = "none";
  document.getElementById("response-container").innerText = "";
}
document.getElementById("cancelAuthToken").onclick = hideAuthTokenDialog;

function handleSubmitAuthToken() {
  const authToken = document.getElementById("authTokenInput").value.trim();
  const errorContainer = document.getElementById("error-message");

  if (authToken.length === 0) {
    errorContainer.innerText = "Por favor, ingresa un token válido.";
    errorContainer.style.display = "block";
    return;
  }

  localStorage.setItem("authToken", authToken);
  sendAuthToken(authToken);
  hideAuthTokenDialog();
}

function displayResponse(response) {
  const responseContainer = document.getElementById("response-container");
  if (response && response.message) {
    responseContainer.innerText = response.message;
  } else {
    responseContainer.innerText = "No response from server.";
  }
}

