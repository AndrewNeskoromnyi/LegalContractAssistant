/*
 * Copyright (c) Syngraphus LLC. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
//import OpenAI from "openai";
//let baseAPIUrl = "https://localhost:44314/api/LegalContract/";

let baseAPIUrl = "https://legalcontractwebapi.azurewebsites.net/api/LegalContract/";

const xlsxParser = require("xlsx");
let wb = xlsxParser.WorkBook;

let openAIResponse = "";
let openAIResponseNoColor = "";

let fileContent = "";
let assistantId = "";

let currentFontName = "";
let currentFontSize = "";

/* #region File Processing */

// Gets the contents of the selected file.
async function getXlsxFileContents() {
  const reader = new FileReader();
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
       <div>
         <span>
           <button class="ms-Button ms-Button--primary" id="querybutton_${sectionKey}" name="querybutton" disabled>
               <span class="ms-Button-label">Suggest changes</span>
            </button>
            <button class="ms-Button ms-Button--primary" id="approvebutton_${sectionKey}" name="approvebutton" disabled>
               <span class="ms-Button-label">Approve and replace</span>
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

  $("a[name=aheader]").on("click", function (event) {
    event.preventDefault();
    tryCatch(findArticleWhenClickOnSection(event.target.innerText, event.target.id));
  });

  $(":button").on("click", async function () {
    //TODO
    //identify which button by name

    let id = String($(this).prop("id"));
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

        openAIPanel("openAIPanel_" + sectionId, article, annotations, "error_" + sectionId);

        //enable approve and insert buttons
        $("#approvebutton_" + sectionId).removeAttr("disabled");
        $("#insertbutton_" + sectionId).removeAttr("disabled");
        break;
      case "approvebutton":
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
      $("#btnLogin").on("click", () => tryCatch(login));

      search();
    }
  });
});

// When the user closes the Word window, delete assistant on the server.
window.onbeforeunload = function() {

 deleteAssistant();

};

function login() {
  //TODO login
  $("#login-section").hide();

  $("#header-section").show();
}

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
  let $referenceSection = $("#references-section");
  $referenceSection.show();
  $referenceSection.change();
}

function enableDisableSectionElements(sectionId) {
  let sectionQueryButton = $("#querybutton_" + sectionId);
  let allButtons = $(":button").not(sectionQueryButton);
  allButtons.attr("disabled", "disabled");

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
    //at least one checked - enable Suggestion button
    sectionQueryButton.removeAttr("disabled");
  }
}

// Open AI panel on <Suggest Changes> click. Close all other AI panels.
async function openAIPanel(panelId, article, annotations, errorsectionId) {
  let $openAISection = $("#" + panelId);

  $("#please_wait").show();

  // textarea is empty
  await getOpenAIResponseFromAssistant(article, annotations, errorsectionId);

        // Wait until the article processing is complete
        await new Promise((resolve) => {
          document.addEventListener("articleProcessingComplete", function handler(event) {
            document.removeEventListener("articleProcessingComplete", handler);
            resolve();
          });
        });
  

  $openAISection.html(openAIResponse); //populate panel with response from openAI

  var clone = $openAISection.clone();

  var elements = $(clone);
  elements.find('*').removeAttr('style'); //remove highlighted text `style` attribute

  //set font to current paragraph
  elements.find('*').css('font-family', currentFontName);
  elements.find('*').css('font-size', (96 / 72 * currentFontSize) + "px"); ////Font size returned by the doc is in pt -> convert to px
  

  openAIResponseNoColor = elements.html();

  //Hide Wait, re-enable controls
  $("#please_wait").hide();

  $openAISection.show();
  $("#color-annotation-map").show();
}

function showWait() {
  document.getElementById("overlay").style.display = "flex";
}

function hideWait() {
  document.getElementById("overlay").style.display = "none";
}

async function getOpenAIResponseFromAssistant(article, annotations, errorsectionid) {
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
        article: article,
        annotations: annotations,
        assistantId: assistantId,
        fileContent: fileContent
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

/* // Listen for the custom event to notify that the file preparation is complete
document.addEventListener("filePreparationComplete", function(event) {
  // Access the file content from the event detail
  fileContent = event.detail.fileContent;
  console.log("File preparation complete. File content is ready.");
  // You can now proceed with further processing using the fileContent
}); */


/* // Clears the selected radio button.
async function clearSelection() {
  $("input[name='citation'][type='radio']:checked").prop("checked", false);
  $(".input").hide(); // hides all openAI panels with class input
  $(".error").hide(); // hides all openAI panels with class error
  clearSelected();
  disableButtons();
}

// Sets the selected item.
async function setSelected(text) {
  $("#selected").text(text);
} */

/* // Clears the selected item.
async function clearSelected() {
  $("#selected").text("");
}

// Enables the buttons.
async function enableButtons() {
  $(".ms-Button").removeAttr("disabled");
}

// Disables the buttons.
async function disableButtons() {
  $(".ms-Button").attr("disabled", "disabled");
} */

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

          firstParagraph.insertHtml(openAIResponseNoColor, "Replace");
          await context.sync();
         

          break;
        }
      }
      await context.sync();
    } catch (error) {
      console.log(error);
    }
  });
}

/* // Inserts the suggestion after selected text in the document.
async function insertSuggestion() {
  await Word.run(async (context) => {
    const radioId = $("input[name='citation'][type='radio']:checked").attr("id");
    let $openAISection = $("#openAIPanel" + radioId);
    let suggestion = $openAISection.text();

    const doc = context.document;
    const originalRange = doc.getSelection();

    originalRange.insertHtml(suggestion, "After");

    await context.sync();
    console.log(`Inserted suggestion: ${citationsuggestion}`);
  });
} */

/* // Replaces selected text in the document with suggestion.
async function insertSuggestionUnderArticle() {
  await Word.run(async (context) => {
    const radioId = $("input[name='citation'][type='radio']:checked").attr("id");
    let $openAISection = $("#openAIPanel" + radioId);
    let suggestion = $openAISection.text();

    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(suggestion, Word.InsertLocation.replace);

    await context.sync();

    console.log(`Replacement suggestion: ${suggestion}`);
  });
} */

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
      $errorSection.text(error.message);
      $errorSection.show();
      console.log(error);
    }
  });
}
/* #endregion */

/* #region  API */
async function processArticle(data,errorsectionid) {
  var jsonData = JSON.stringify(data);

  $.ajax({
    type: "POST",
    url: baseAPIUrl + "ProcessArticle",
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
    url: baseAPIUrl + "DeleteAssistant",
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
