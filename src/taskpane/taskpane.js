/*
 * Copyright (c) Syngraphus LLC. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


const xlsxParser = require("xlsx");
let wb = xlsxParser.WorkBook;

let openAIResponse = "";
let accountId = "";
const formattedResponse = new Map();
  

let fileContent = "";
let assistantId = "";

let currentFontName = "";
let currentFontSize = "";

let isLicensed = false
let MAX_UNLICENSED_ANNOTATIONS = 10;
/* #region Login */
const config = {
  auth: {
    clientId: "baf61594-46cb-47a0-9be8-7939e1487286", // This is the ONLY mandatory field; everything else is optional.
    authority: "https://syngraphusb2corganization.b2clogin.com/syngraphusb2corganization.onmicrosoft.com/B2C_1_signupsignin1", // Choose sign-up/sign-in user-flow as your default.
    knownAuthorities: ["syngraphusb2corganization.b2clogin.com"], // You must identify your tenant's domain as a known authority.
    redirectUri: REDIRECT_URL, 
    
  },
  cache: {
    cacheLocation: "sessionStorage", // Configures cache location. "sessionStorage" is more secure, but "localStorage" gives you SSO between tabs.
    storeAuthStateInCookie: false, // If you wish to store cache items in cookies as well as browser cache, set this to "true".
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        switch (level) {
          case msal.LogLevel.Error:
            console.error(message);
            return;
          case msal.LogLevel.Info:
            console.info(message);
            return;
          case msal.LogLevel.Verbose:
            console.debug(message);
            return;
          case msal.LogLevel.Warning:
            console.warn(message);
            return;
        }
      }
    }
  }
};

async function start() {

  await login();

  $("#login-section").hide();

  $("#excelfile-section").show();
}


async function login() {

    const loginRequest = {
      scopes: ["User.ReadWrite"],
    };
    const msalInstance = new msal.PublicClientApplication(config);

      await msalInstance
      .loginPopup()
      .then(function (loginResponse) {
        accountId = loginResponse.account.homeAccountId;

        //var email = loginResponse.account.idTokenClaims.emails[0];

        // Display signed-in user content, call API, etc.
        $("#dialog_welcome").text("Welcome "+loginResponse.account.name+"!");
        
        const licenseExpirationDate = new Date(loginResponse.account.idTokenClaims.extension_LicenseExpirationDate);
        const today = new Date();
        if (loginResponse.account.idTokenClaims.extension_LicenseNumber && licenseExpirationDate > today) {
          isLicensed = true;
        }
        
        if (accountId) {
          showNavBar(true,loginResponse.account.idTokenClaims.extension_LicenseExpirationDate);}
        else {
          showNavBar(false,"");
        }

        //loginResponse.account.idTokenClaims.extension_LicenseNumber 
      })
      .catch(function (error) {
        //login failure
        console.log(error);
        showNavBar(false,"");
      });  

 }

async function logout() {
    const msalInstance = new msal.PublicClientApplication(config);
      
    const logoutRequest = {
      account: msalInstance.getAccountByHomeId(accountId),
      mainWindowRedirectUri: REDIRECT_URL,
    };

    msalInstance["browserStorage"].clear();
    await msalInstance.logoutPopup(logoutRequest);

    showNavBar(false,"");
}

function showNavBar(isLoggedIn,licenseExpirationDate)
{
  if(isLoggedIn)
  {
    $("#btnLogin").hide();
    $("#btnLogout").show();

    if (isLicensed) {
      $("#dialog_license").text("Licensed (Exp: " + licenseExpirationDate + ")");
    } else {
      $("#dialog_license").text("Unlicensed ("+ MAX_UNLICENSED_ANNOTATIONS +" annotations limit)");
    }
  }
  else
  {
    $("#btnLogin").show();
    $("#btnLogout").hide();

    isLicensed = false;

    $("#dialog_welcome").text("");
    $("#dialog_license").text("Unlicensed (10 annotations limit)");
  }
}
/* #endregion */

/* #region File Processing */

// Gets the contents of the selected file.
async function getXlsxFileContents() {
async function getXlsxFileContents() {
  const reader = new FileReader();
  const myXLSXFile = document.getElementById("xlsx-file");
  const use_utf8 = true;

  const myXLSXFile = document.getElementById("xlsx-file");
  const use_utf8 = true;

  reader.onloadend = function () {
    var data = new Uint8Array(reader.result);
    wb = xlsxParser.read(data, { type: "array", codepage: use_utf8 ? 65001 : void 0 });

    populateAnnotationsFromFile();
    showReferencesSection();
  };
  reader.readAsArrayBuffer(myXLSXFile.files[0]);
}

function to_json(workbook) {
  var result = {};
  workbook.SheetNames.forEach(function (sheetName) {
    var roa = xlsxParser.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
    if (roa.length > 0) {
      result[sheetName] = roa;
    }
  });
  return result;
}

// Populates the radio buttons with the citations from the file.
async function populateAnnotationsFromFile() {
  //var result=to_csv(wb);
  let annotationsFromFile = to_json(wb);
  console.log(annotationsFromFile);

  let $populateRadio = $("#populate-radio");
  let $radioButtons = $("#radio-buttons");
  $radioButtons.empty();

  var newArticle = "";

  let annotationHtmlTotal = "";

  //TODO Create map
  const map = new Map();

  for (let i = 0; i < annotationsFromFile.Sheet1.length; ++i) {
    var article = String(annotationsFromFile.Sheet1[i].Article);

    var note = String(annotationsFromFile.Sheet1[i].Annotation);

    if (newArticle != article) {
      map.set(article, [note]);
    } else {
      map.get(article).push(note);
    }
    newArticle = article;
  }

  var i = 0;
  var j = 0;
  for (const [key, value] of map) {
    if(!isLicensed && j >= MAX_UNLICENSED_ANNOTATIONS) //limit to 10 articles for unlicensed users
    {
      break;
    }

    var sectionKey = "sectionKey" + String(i);

    var sectionHeaderHtml = `<section>
      <div class="container">
      <p><span><a id="${sectionKey}" name='aheader' href='#'>${key}</a></span></p>`;

    for (note of value) {
      var rowKey = "rowkey" + String(j);
      let annotationHtml = `<input type="checkbox" id='checkbox_${sectionKey}_${rowKey}' name="annotation" value='${key.replace("'", "&#039;")}|${note}' checked disabled>
        <label for="${rowKey}">${note}</label><br>
        <br>`;
      annotationHtmlTotal += annotationHtml;
      j += 1;
    }
    var sectionFooterHtml = `<div>
    <div class="taria" id="openAIPanel_${sectionKey}" style="display: none;"></div>
       <div id="buttonsection_${sectionKey}" style="display: none;">
         <span>
           <button class="ms-Button ms-Button--primary" id="querybutton_${sectionKey}" name="querybutton" didabled>
               <div class="ms-Button-label">Suggest changes for Article</div>
            </button>
            <button class="ms-Button ms-Button--primary" id="approvebutton_${sectionKey}" name="approvebutton" disabled>
               <div class="ms-Button-label">Approve and replace</div>
             </button>
          </span>
        </div>
        <div class="error" id="error_${sectionKey}" style="display: none;"></div>
      </section>`;

    $radioButtons.append(sectionHeaderHtml + annotationHtmlTotal + sectionFooterHtml);
    i += 1;
    annotationHtmlTotal = "";
  }
  $radioButtons.appendTo($populateRadio);

  // Click on article link
  $("a[name=aheader]").on("click", function (event) {
    event.preventDefault();
    tryCatch(findArticleWhenClickOnSection(event.target.innerText, event.target.id));
  });

  //click on dynamic buttons
  $(":button").on("click", async function () {

    let id = String($(this).prop("id"));

    if(id === "btnProcessAll") //static button has its own handler
    {
      return;
    }

    let values = id.split("_");
    let sectionId = values[1];

    //find all checkboxes in this section
    var allSectionCheckboxes = $("input[type='checkbox'][id^=" + "checkbox_" + String(sectionId) + "]");

    //find article for this section
    var article = allSectionCheckboxes[0].value.split("|")[0];

    switch (values[0]) {
      case "querybutton":
        let annotations = [];
        allSectionCheckboxes.each(function () {
          let value = $(this).val().split("|")[1];
          annotations.push(value);
        });

        openAIPanel("openAIPanel_" + sectionId, article, annotations, "article", "error_" + sectionId);

        //enable approve and insert buttons
        $("#approvebutton_" + sectionId).removeAttr("disabled");
        break;
      case "approvebutton":

         //Remove clor attributes from openAI response
        let $panel = $("#openAIPanel_" + sectionId);

        var clone = $panel.clone();
        var elements = $(clone);
        elements.find('*').removeAttr('style'); //remove highlighted text `style` attribute
      
        //set font to current paragraph
        elements.find('*').css('font-family', currentFontName);
        elements.find('*').css('font-size', (96 / 72 * currentFontSize) + "px"); ////Font size returned by the doc is in pt -> convert to px
        openAIResponseNoColor = elements.html(); 
    
        formattedResponse.set(article, openAIResponseNoColor);

        await replaceArticleWithSuggestion(article);
        break;
      /*       case "insertbutton":
        await insertSuggestionUnderArticle();
        break; */
      default:
        console.log("Unknown button action");
    }
  });

  //Click on annotation checkbox
  $("input[name='annotation'][type='checkbox']").on("click", function () {
    let id = String($(this).prop("id"));
    let values = id.split("_");
    let sectionId = values[1];

    enableDisableSectionElements(sectionId);
  });
}

/* #endregion */

/* #region  User Interface */

Office.onReady((info) => {
  $(document).ready(function () {
    if (info.host === Office.HostType.Word) {
      document.getElementById("app-body").style.display = "flex";
      $("#xlsx-file").on("change", () => tryCatch(getXlsxFileContents));

      $(window).off('beforeunload'); //prevent "Leave site?" dialog

      //to create async call
      document.getElementById("btnLogin").onclick=async() => {
      await login();
     
      };

      //to create async call
      document.getElementById("btnLogout").onclick=async() => {
        await logout();
        
        };

      //to create async call
      document.getElementById("btnStart").onclick=async() => {
        await start();
        
        };

      //to create async call
      document.getElementById("btnProcessAll").onclick=async() => {
        await processAllAnotations();
      };


      search();
    }
  });
});

// When the user closes the Word window, delete assistant on the server.
window.onbeforeunload = async function() {

 await deleteAssistant();

};




// Searches the references list for the search text.
async function search() {
  let $search = $("#search");
  let $radioButtons = $("#radio-buttons");
  $search.on("search keyup", function () {
    let searchText = $(this).val();
    if (searchText) {
      $radioButtons.children().each(function () {
        let $this = $(this);
        if ($this.text().search(new RegExp(searchText, "i")) < 0) {
          $this.hide();
        } else {
          $this.show();
        }
      });
    } else {
      $("#radio-buttons")
        .children()
        .each(function () {
          $(this).show();
        });
    }
    $radioButtons.change();
  });
}

// Shows the reference section.
async function showReferencesSection() {
  //hide file dialog
  let $excelSection = $("#excelfile-section");
  $excelSection.hide();
  $excelSection.change();

  let $referenceSection = $("#references-section");
  $referenceSection.show();
  $referenceSection.change();
}

function enableDisableSectionElements(sectionId) {

  let panel = $("#openAIPanel_" + sectionId);
  let sectionQueryButton = $("#querybutton_" + sectionId);
  let sectionApproveButton = $("#approvebutton_" + sectionId);
  
  //let processAllChangesButton = $("#btnProcessAll");


 let buttonsection = $("#buttonsection_" + sectionId);
 buttonsection.show();

  //let allButtons = $(":button").not(sectionQueryButton).not(processAllChangesButton).not(sectionApproveButton);
  //llButtons.attr("disabled", "disabled");
  //allButtons.attr("hidden", "true");
  
  //find all other checkboxes and disable them
  var allCheckboxes = $("input[type='checkbox'][id^='checkbox_']");
  allCheckboxes.prop("disabled", true);

  //find all checkboxes in this section and enable them
  var allSectionCheckboxes = $("input[type='checkbox'][id^=" + "checkbox_" + String(sectionId) + "]");
  allSectionCheckboxes.prop("disabled", false);

  if (allSectionCheckboxes.filter(":checked").length == 0) {
    //if all unchecked disable Suggestion button
    sectionQueryButton.attr("disabled", "disabled");
  } else {
    //at least one checked - enable Suggestion and apply button
    sectionQueryButton.removeAttr("disabled");

    // Only enable Approve button if panel is not empty
    if ($("#openAIPanel_" + sectionId).html().trim() !== "") {
      sectionApproveButton.removeAttr("disabled");
    }
  }
}

// Open AI panel on <Suggest Changes> click. Close all other AI panels.
async function processAllAnotations() {
 try {
  $("#please_wait").show();

  //find all checked checkboxes
  var allCheckboxes = $("input[type='checkbox'][id^='checkbox_']");
  var allCheckedCheckboxes = allCheckboxes.filter(":checked");

  var articleAnnotations = [];

  allCheckedCheckboxes.each(function () {
    var article = this.value.split("|")[0];
    var annotation = this.value.split("|")[1];
    var existingArticle = articleAnnotations.find(a => a.article === article);

    if (existingArticle) {
      existingArticle.annotations.push(annotation);
    } else {
      articleAnnotations.push({
        article: article,
        annotations: [annotation]
      });
    }
  });

  
  // textarea is empty
  await getOpenAIResponseFromAssistant(articleAnnotations, "document", "errorProcessAll");

        // Wait until the article processing is complete
        await new Promise((resolve) => {
          document.addEventListener("articleProcessingComplete", function handler(event) {
            document.removeEventListener("articleProcessingComplete", handler);
            resolve();
          });
        });
  
  // Create a map from the openAIResponse where key is a header outlined by <h1> and </h1> tags and value is the HTML text under the header
  const parser = new DOMParser();
  const doc = parser.parseFromString(openAIResponse, 'text/html');
  const headers = doc.querySelectorAll('h1');
  const map = new Map();
  
  headers.forEach(header => {
    const key = header.textContent;
    let value = '';
    let sibling = header.nextElementSibling;

    while (sibling && sibling.tagName !== 'H1') {
      value += sibling.outerHTML;
      sibling = sibling.nextElementSibling;
    }

    map.set(key, value);
    
  });

  // Iterate through the map and log the key-value pairs
  for (const [key, value] of map) {

      // Find the element with name "aheader" and value equal to the key of the map
      let element = $(`a[name='aheader']`).filter(function() {
        return $(this).text() === key;
      });

    
      if (element.length) {

        let $openAISection = $("#openAIPanel_" + element[0].id);
        $openAISection.html(value); //populate panel with response from openAI
        $openAISection.show();
    
   
      }
  
  }

  //Hide Wait, re-enable controls
  $("#please_wait").hide();

  
  $("#color-annotation-map").show();

}
catch (error) {
  $("#please_wait").hide();
  $errorSection.text(error.message);
  $errorSection.show();
  console.log(error);
}
}




// Open AI panel on <Suggest Changes> click. Close all other AI panels.
async function openAIPanel(panelId, article, annotations, scope, errorsectionId) {
  let $openAISection = $("#" + panelId);

  $("#please_wait").show();


  var articleAnnotations = [];
  articleAnnotations.push(
    {
        article: article,
        annotations: annotations
      }
    );
  
  // textarea is empty
  await getOpenAIResponseFromAssistant(articleAnnotations, scope, errorsectionId);

        // Wait until the article processing is complete
        await new Promise((resolve) => {
          document.addEventListener("articleProcessingComplete", function handler(event) {
            document.removeEventListener("articleProcessingComplete", handler);
            resolve();
          });
        });
  

  $openAISection.html(openAIResponse); //populate panel with response from openAI


  //Hide Wait, re-enable controls
  $("#please_wait").hide();

  $openAISection.show();
  $("#color-annotation-map").show();
}


async function getOpenAIResponseFromAssistant(articleAnnotations, scope, errorsectionid) {
  try {
    let $errorSection = $("#" + errorsectionid);
    $errorSection.hide();
    let $generalerrorSection = $("#generalError");
    $generalerrorSection.hide();
    
    fileContent="";

      //When assistantId empty, current file will be uploaded and service will create a new assistant 
      //Otherwise it will re-use existing assistant based on assistantId
      if (assistantId === "") 
      {
        //Prepare file for upload
        prepFile();

        // Wait until the file preparation is complete
        await new Promise((resolve) => {
          document.addEventListener("filePreparationComplete", function handler(event) {
            document.removeEventListener("filePreparationComplete", handler);
            resolve();
          });
        });

      }


      let data = {
        articles: articleAnnotations,
        assistantId: assistantId,
        fileContent: fileContent,
        scope: scope
      };

      await processArticle(data,errorsectionid);
  }
  catch (error) {
    $("#please_wait").hide();
    $errorSection.text(error.message);
    $errorSection.show();
    console.log(error);
  }
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.log(error);
  }
}

/* #endregion */

/* #region  Word Processing */
// Inserts the suggestion after selected text in the document.
//Remove all paragraphs for this article
async function replaceArticleWithSuggestion(article) {
  await Word.run(async (context) => {
    try {
      //TODO Need better algorithm to identify header string
      //var heading1Regex = "^[0-9]+."; //String begins with number followed by period
      var heading1Regex = "^[0-9]+[. ]"; // String begins with number followed by period or space
      let paragraphs = context.document.body.paragraphs;
      context.load(paragraphs, ["text"], ["items"], ["font"]);
      // Synchronize the document state by executing the queued commands
      await context.sync();

      for (let i = 0; i < paragraphs.items.length; ++i) {
        let item = paragraphs.items[i];

        if (item.text.trim().replace("’", "'").toUpperCase() === article.trim().replace("’", "'").toUpperCase()) {
          let paragraphsToDelete = [];

          var firstParagraph = paragraphs.items[i + 1];

          //first paragraph after header must NOT BE another header, because it could be Table of Content
          if (firstParagraph.text.match(heading1Regex)) {
            continue; // it is table of content - look for the next one
          }

          var nextParagraph;
          var j = i + 2;

          while (true) {
            nextParagraph = paragraphs.items[j];
            if (nextParagraph.text.match(heading1Regex)) {
              break;
            } else {
              paragraphsToDelete.push(nextParagraph);
            }
            j = j + 1;
          }

          for (const p of paragraphsToDelete) {
            p.delete();
          }

          await context.sync();

              // Insert the suggestion after the selected text in the document.
          if (formattedResponse.has(article)) {
            firstParagraph.insertHtml(formattedResponse.get(article), "Replace");
            console.log(`Key: ${key}, Value: ${formattedResponse.get(key)}`);
          }


          await context.sync();
         

          break;
        }
      }
      await context.sync();
    } catch (error) {
      //console.log(error);
    }
  });
}



async function findArticleWhenClickOnSection(articleText, sectionId) {
  await Word.run(async (context) => {
    try {
      var errorsectionid = "error_" + String(sectionId);

      enableDisableSectionElements(sectionId);

      // Queue a command to search the document and ignore punctuation.
      const searchResults = context.document.body.search(String(String(articleText)), { ignorePunct: true });
      searchResults.load("items");
      await context.sync();

      var found;

      //getLast
      for (const p of searchResults.items) {
        found = p;
      }

      found.select();

      let nextRange = found.getNextTextRangeOrNullObject(["."],true);
      nextRange.load("font");
      await context.sync();

      currentFontName = nextRange.font.name;
      currentFontSize = nextRange.font.size;
      
      await context.sync();

     
    } catch (error) {
      let $errorSection = $("#" + errorsectionid);
      $errorSection.text("Cannot find the article...");
      $errorSection.show();
      //console.log(error);
    }
  });
}

/* #endregion */

/* #region  API */
async function processArticle(data,errorsectionid) {
  var jsonData = JSON.stringify(data);

  $.ajax({
    type: "POST",
    url: API_URL + "ProcessArticle",
    data: jsonData,
    dataType: "json",
    contentType: "application/json; charset=utf-8",
  })
    .done(function (result) {
      assistantId = result.assistantId;
      openAIResponse = String(result.htmlString);

    // Trigger a custom event to notify that the article processed
    const event = new CustomEvent("articleProcessingComplete");
    document.dispatchEvent(event);
    })
    .fail(function (jqXHR, textStatus) {
      let $errorSection = $("#" + errorsectionid);
      $errorSection.text(jqXHR.responseText);
      $errorSection.show();
    });
}

async function deleteAssistant() {
  
  $.ajax({
    type: "POST",
    url: API_URL + "DeleteAssistant",
    data: JSON.stringify(assistantId),
    dataType: "json",
    contentType: "application/json; charset=utf-8",
  })
    .done(function (result) {
    })
    .fail(function (jqXHR, textStatus) {
      let $errorSection = $("#generalError");
      $errorSection.text(jqXHR.responseText);
      $errorSection.show();
    });
}
/* #endregion */

/* #region  Word file processor */
/* // Get a slice from the file and then call sendSlice.
function getSlice(state) {
  state.file.getSliceAsync(state.counter, function (result) {
    if (result.status == Office.AsyncResultStatus.Succeeded) {
      updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
      sendSlice(result.value, state);
    } else {
      updateStatus(result.status);
    }
  });
}

function sendSlice(slice, state) {
  var data = slice.data;

  // If the slice contains data, create an HTTP request.
  if (data) {
    // Encode the slice data, a byte array, as a Base64 string.
    // NOTE: The implementation of myEncodeBase64(input) function isn't
    // included with this example. For information about Base64 encoding with
    // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
    var fileData = myEncodeBase64(data);

    // Create a new HTTP request. You need to send the request
    // to a webpage that can receive a post.
    var request = new XMLHttpRequest();

    // Create a handler function to update the status
    // when the request has been sent.
    request.onreadystatechange = function () {
      if (request.readyState == 4) {
        updateStatus("Sent " + slice.size + " bytes.");
        state.counter++;

        if (state.counter < state.sliceCount) {
          getSlice(state);
        } else {
          closeFile(state);
        }
      }
    };

    request.open("POST", "https://localhost:44314/api/LegalContract/PostFileContent");
    request.setRequestHeader("Slice-Number", slice.index);

    // Send the file as the body of an HTTP POST
    // request to the web server.
    request.send(fileData);
  }
}

function closeFile(state) {
  // Close the file when you're done with it.
  state.file.closeAsync(function (result) {
    // If the result returns as a success, the
    // file has been successfully closed.
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      updateStatus("File closed.");
    } else {
      updateStatus("File couldn't be closed.");
    }
  });
} */

// Usually we encode the data in base64 format before sending it to server.
function encodeBase64(docData) {
  var s = "";
  for (var i = 0; i < docData.length; i++) s += String.fromCharCode(docData[i]);
  return window.btoa(s);
}

// Call getFileAsync() to start the retrieving file process.
function prepFile() {
  Office.context.document.getFileAsync("compressed", { sliceSize: 10240 }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      document.getElementById("log").textContent = JSON.stringify(asyncResult);
    } else {
      getAllSlices(asyncResult.value);
    }
  });
}

// Get all the slices of file from the host after "getFileAsync" is done.
function getAllSlices(file) {
  var sliceCount = file.sliceCount;
  var sliceIndex = 0;
  var docdata = [];
  var getSlice = function () {
    file.getSliceAsync(sliceIndex, function (asyncResult) {
      if (asyncResult.status == "succeeded") {
        docdata = docdata.concat(asyncResult.value.data);
        sliceIndex++;
        if (sliceIndex == sliceCount) {
          file.closeAsync();
          onGetAllSlicesSucceeded(docdata);
        } else {
          getSlice();
        }
      } else {
        file.closeAsync();
        document.getElementById("log").textContent = JSON.stringify(asyncResult);
      }
    });
  };
  getSlice();
}

/* // Upload the docx file to server after obtaining all the bits from host.
function onGetAllSlicesSucceeded(docxData) {
  $.ajax({
      type: "POST",
      url: "https://localhost:44314/api/LegalContract/PostFileContent",
      //data: encodeBase64(docxData),
      data: JSON.stringify(encodeBase64(docxData)),
      dataType:"json",
     contentType: "application/json; charset=utf-8",
  }).done(function (data) {
     // document.getElementById("documentXmlContent").textContent = data;
  }).fail(function (jqXHR, textStatus) {
  });
}  */

// Upload the docx file to server after obtaining all the bits from host.
function onGetAllSlicesSucceeded(docxData) {
 
  fileContent = encodeBase64(docxData);
  // Trigger a custom event to notify that the file preparation is complete
  const event = new CustomEvent("filePreparationComplete");
  document.dispatchEvent(event);
}


/* #endregion */
