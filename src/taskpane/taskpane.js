/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import OpenAI from "openai";

const bibtexParser = require("@orcid/bibtex-parse-js");
let bibFileContent;
let openAIResponse = "";

Office.onReady((info) => {
  $(document).ready(function () {
    if (info.host === Office.HostType.Word) {
      document.getElementById("app-body").style.display = "flex";
      $("#bib-file").on("change", () => tryCatch(getFileContents));
      search();
      $("#replace-suggestion").on("click", () => tryCatch(replaceSuggestion));
      $("#insert-suggestion").on("click", () => tryCatch(insertSuggestion));
      $("#clear").on("click", () => tryCatch(clearSelection));
    }
  });
});

// test openAI
async function getOpenAIResponse(question) {

  //todo do not exposu API key in browser - dangerouslyAllowBrowser: true
    const openai = new OpenAI({
      apiKey: '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',
      dangerouslyAllowBrowser: true
  });
  
    const chatCompletion = await openai.chat.completions.create({
      messages: [{ role: "user", content: "Say "+ question }],
      model: "gpt-4o-mini",
  });
  openAIResponse = chatCompletion.choices[0]?.message?.content;

  }

// test openAI
async function getOpenAIResponseFromAssistant(article,annotation,errorsectionid) {
try{
  //todo do not exposu API key in browser - dangerouslyAllowBrowser: true
    const openai = new OpenAI({
      apiKey: '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',
      dangerouslyAllowBrowser: true
  });
  
let assistantId = "asst_BnExFeePPpmWiBzyNfkWrAaP";
let question = 'Change Article ' + article + ' in a way that "' + annotation + '". Show only revised version of '+ article +'. Highlight the changes using Italic font.';

const thread = await openai.beta.threads.create({
  messages: [
    {
      role: 'user',
      content: String(question),
    },
  ],
});

let threadId = thread.id;
  console.log('Created thread with Id: ' + threadId);

//await setSelected(String(threadId));

 const run = openai.beta.threads.runs
  .stream(threadId, {
    assistant_id: assistantId,
  });

const result = await run.finalRun();

if (result.status == 'completed') {
  const messages = await openai.beta.threads.messages.list(threadId);
  
  openAIResponse = messages.data[0].content[0].text.value;
  openAIResponse = String(openAIResponse).replace("\n",String.fromCharCode(11));
  
}


}

catch (error) {
  let $errorSection = $("#"+errorsectionid);
  $errorSection.text(error.message);
  $errorSection.show();
  console.log(error);
}


  }



// Gets the contents of the selected file.
async function getFileContents() {
  const myBibFile = document.getElementById("bib-file");
  const reader = new FileReader();
  reader.onloadend = function () {
    bibFileContent = reader.result;
    populateCitationsFromFile();
    showReferencesSection();
  };
  reader.readAsBinaryString(myBibFile.files[0]);
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
      $("#radio-buttons").children().each(function () {
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

// Populates the radio buttons with the citations from the file.
async function populateCitationsFromFile() {
    let citationsFromFile = bibtexParser.toJSON(bibFileContent);
    console.log(citationsFromFile);

    let $populateRadio = $("#populate-radio");
    let $radioButtons = $("#radio-buttons");
    $radioButtons.empty();
    for (let citation in citationsFromFile) {
      let citationHtml = `<section><input type="radio" id="${citationsFromFile[citation].citationKey}" name="citation" value='${citationsFromFile[citation].entryTags.article}|${citationsFromFile[citation].entryTags.annotation}'>
      <label for="${citationsFromFile[citation].citationKey}"><b>${citationsFromFile[citation].entryTags.annotation}</b><br>${citationsFromFile[citation].entryTags.article}</label><br><br>
      <div>
      <textarea class="input" id="openAIPanel${citationsFromFile[citation].citationKey}" rows="8" cols="80"  style="display: none;"></textarea>
      <\div>
      <div class="error" id="error${citationsFromFile[citation].citationKey}" style="display: none;">
      <\div>
      </section>`;
      $radioButtons.append(citationHtml);
    }
    $radioButtons.appendTo($populateRadio);//test

    $("input[name='citation'][type='radio']").on("click", function () {
      if ($(this).prop("checked")) {
        //replace quatation
        let value = String($(this).prop("value"));
        let values = value.split("|");
        let article = values[0];
        let annotation = values[1];

        setSelected(article);
        enableButtons();

        let errorSectionId=`error${$(this).prop("id")}`;
        findArticle(article.replace('Article','').trim(),errorSectionId)

        
        openAIPanel(`openAIPanel${$(this).prop("id")}`,article,annotation,errorSectionId);

      } else {
        clearSelected();
        disableButtons();
      }
    });
    $populateRadio.change();
}


// Open AI panel for radio button. Close all other AI panels.
async function openAIPanel(panelId,article,annotation,errorsectionId) {

$('.input').hide(); // hides all openAI panels with class input
let $openAISection = $("#"+panelId);  

//only access openAI if not yet populated
if (String($openAISection.text()).trim().length == 0)
{
  //disable all radio-buttons and show Please wait
  let $radioButtons = $("#radio-buttons");
  $radioButtons.children('input[type=radio]').prop('disabled', true);
  $('#please_wait').show();

  // textarea is empty
  await getOpenAIResponseFromAssistant(article,annotation,errorsectionId);
  $openAISection.text(openAIResponse); //populate panel with response from openAI

  //Hide Wait, re-enable controls
  $radioButtons.children('input[type=radio]').prop('disabled', false);
  $('#please_wait').hide();
}

$openAISection.show();

}


// Inserts the suggestion after selected text in the document.
async function insertSuggestion() {
  await Word.run(async (context) => {
    const radioId = $("input[name='citation'][type='radio']:checked").attr('id');
    let $openAISection = $("#openAIPanel"+radioId);
    let suggestion = $openAISection.text();
    
    const doc = context.document;
    const originalRange = doc.getSelection();
    
    originalRange.insertText(suggestion, Word.InsertLocation.end);

    await context.sync();
    console.log(`Inserted suggestion: ${citationsuggestion}`);
  });
}

// Replaces selected text in the document with suggestion.
async function replaceSuggestion() {
  await Word.run(async (context) => {
    const radioId = $("input[name='citation'][type='radio']:checked").attr('id');
    let $openAISection = $("#openAIPanel"+radioId);
    let suggestion = $openAISection.text();

    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(suggestion, Word.InsertLocation.replace);

    await context.sync();

    console.log(`Replacement suggestion: ${suggestion}`);
  });
}

// Clears the selected radio button.
async function clearSelection() {
  $("input[name='citation'][type='radio']:checked").prop("checked", false);
  $('.input').hide(); // hides all openAI panels with class input
  $('.error').hide(); // hides all openAI panels with class error
  clearSelected();
  disableButtons();
}

// Sets the selected item.
async function setSelected(text) {
  $("#selected").text(text);
}

// Scroll to selected article.
async function findArticle(articletext,errorsectionid) {
  await Word.run(async (context) => {
    try {
    // Queue a command to search the document and ignore punctuation.
    const searchResults = context.document.body.search(articletext, {ignorePunct: true});
    //const searchResults = context.document.body.search(articletext);
   
    searchResults.getFirst().select();
    await context.sync();
    }
    catch (error) {
      let $errorSection = $("#"+errorsectionid);
      $errorSection.text(error.message);
      $errorSection.show();
      console.log(error);
    }
  });
}


// Clears the selected item.
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
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.log(error);
  }
}
