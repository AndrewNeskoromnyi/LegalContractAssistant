/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import OpenAI from "openai";

const xlsxParser = require('xlsx');
let wb = xlsxParser.WorkBook;

let openAIResponse = "";
let selectedArticle="";

Office.onReady((info) => {
  $(document).ready(function () {
    if (info.host === Office.HostType.Word) {
      document.getElementById("app-body").style.display = "flex";
      $("#xlsx-file").on("change", () => tryCatch(getXlsxFileContents));
      search();
     // $("#replace-suggestion").on("click", () => tryCatch(replaceSuggestion));
     // $("#insert-suggestion").on("click", () => tryCatch(insertSuggestionAfterArticleHeader));
     // $("#clear").on("click", () => tryCatch(clearSelection));
    }
  });
});

// test openAI
async function getOpenAIResponse(question) {

  //todo do not exposu API key in browser - dangerouslyAllowBrowser: true
    const openai = new OpenAI({
      apiKey: '',
      dangerouslyAllowBrowser: true
  });
  
    const chatCompletion = await openai.chat.completions.create({
      messages: [{ role: "user", content: "Say "+ question }],
      model: "gpt-4o-mini",
  });
  openAIResponse = chatCompletion.choices[0]?.message?.content;

  }

// test openAI
async function getOpenAIResponseFromAssistantTEST(article,annotation,errorsectionid) {
 
    openAIResponse = "<html><p>First paragraph</p><p>Second paragraph</p></html>";
    
  }

// test openAI
async function getOpenAIResponseFromAssistant(article,annotation,errorsectionid) {
try{
  //todo do not exposu API key in browser - dangerouslyAllowBrowser: true
    const openai = new OpenAI({
      apiKey: '',
      dangerouslyAllowBrowser: true
  });
  
let assistantId = "asst_BnExFeePPpmWiBzyNfkWrAaP";
let question = 'Change Article ' + article + ' in a way that "' + annotation + '". Show only revised version of '+ article +'.Convert text to HTML. Use <html> tag for beginning of the text. Use </html> tag for the end of the text. Highlight changed text in yellow. Highlight added text in red. Do not show article name. Convert original links to html.';

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


 const run = openai.beta.threads.runs
  .stream(threadId, {
    assistant_id: assistantId,
  });

const result = await run.finalRun();

if (result.status == 'completed') {
  const messages = await openai.beta.threads.messages.list(threadId);
  
  openAIResponse = messages.data[0].content[0].text.value;
  let htmlBegin=String(openAIResponse).indexOf("<html>");
  let htmlEnd=String(openAIResponse).indexOf("</html>");
  openAIResponse = String(openAIResponse).substring(htmlBegin,htmlEnd+7).trim();
  
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
async function getXlsxFileContents() {
  const reader = new FileReader();
  const myXLSXFile = document.getElementById("xlsx-file");
  const use_utf8 = true;

  reader.onloadend = function () {
    
    var data = new Uint8Array(reader.result);
    wb = xlsxParser.read(data, {type: 'array', codepage: use_utf8 ? 65001 : void 0});
    populateAnnotationsFromFile();
    showReferencesSection();
  };
  reader.readAsArrayBuffer(myXLSXFile.files[0]);
}

function to_csv(workbook) {
	var result = [];
	workbook.SheetNames.forEach(function(sheetName) {
		var csv = xlsxParser.utils.sheet_to_csv(workbook.Sheets[sheetName]);
		if(csv.length > 0){
			result.push("SHEET: " + sheetName);
			result.push("");
			result.push(csv);
		}
	});
	return result.join("\n");
}

function to_json(workbook) {
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var roa = xlsxParser.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
		if(roa.length > 0){
			result[sheetName] = roa;
		}
	});
	return result;
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
async function populateAnnotationsFromFile() {

  //var result=to_csv(wb);
  let annotationsFromFile = to_json(wb);
  console.log(annotationsFromFile);

  let $populateRadio = $("#populate-radio");
  let $radioButtons = $("#radio-buttons");
  $radioButtons.empty();

  var sectionKeyId = 0;
  var newArticle = "";
  
  let annotationHtmlTotal = "";

//TODO Create map
const map = new Map();


for (let i = 0; i < annotationsFromFile.Sheet1.length; ++i) {
  var article = String(annotationsFromFile.Sheet1[i].Article);

  //var sectionKey = 'sectionKey'+String(sectionKeyId);
  //var rowKey = 'rowkey'+String(annotationsFromFile.Sheet1[i].__rowNum__);
  var note = String(annotationsFromFile.Sheet1[i].Annotation);

  if(newArticle != article)
    {
      map.set(article,[note]);
    }
    else
    {
      map.get(article).push(note);
    }
    newArticle = article;
}

var i=0;
var j=0;
  for (const [key, value] of map) {
      
   // var article = annotationsFromFile.Sheet1[i][Object.keys(annotationsFromFile.Sheet1[i])[0]];
        
    var sectionKey = 'sectionKey'+String(i);
   
      var sectionHeaderHtml = `<section>
      <div class="container">
      <p><span><b>${key}</b></span></p>`

      for(note of value){
        var rowKey = 'rowkey'+String(j); 
        let annotationHtml =`<input type="checkbox" id='checkbox_${sectionKey}_${rowKey}' name="annotation" value='${key.replace("'","&#039;")}|${note}' checked>
        <label for="${rowKey}">${note}</label><br>
        <br>`
        annotationHtmlTotal += annotationHtml;
        j+=1;
    }
      var sectionFooterHtml =`<div>
            <div class="taria" id="openAIPanel_${sectionKey}" style="display: none;"></div>
            <\div>
                    <div>
                        <span>
                            <button class="ms-Button" id="querybutton_${sectionKey}" name="querybutton">
                                <span class="ms-Button-label">Suggest changes</span>
                            </button>
                            <button class="ms-Button" id="approvebutton_${sectionKey}" name="approvebutton" disabled>
                                <span class="ms-Button-label">Approve and replace</span>
                            </button>
                            <button class="ms-Button" id="insertbutton_${sectionKey}" name="insertbutton" disabled>
                                <span class="ms-Button-label">Insert under</span>
                            </button>
                        </span>
                    </div>
            <div class="error" id="error_${sectionKey}" style="display: none;"><\div>
            <\div>
            </section>`;

            $radioButtons.append(sectionHeaderHtml + annotationHtmlTotal + sectionFooterHtml);
    i+=1;
    annotationHtmlTotal = "";

  }
  $radioButtons.appendTo($populateRadio);


  $(":button").on("click", function () {

//TODO 
//identify which button by name

      let id = String($(this).prop("id"));
      let values = id.split("_");
      let sectionId = values[1];

     
    //TODO combine all annotations to pass to generateOpenAI response
    
    //find all checkboxes in this section
   var allSectionCheckboxes = $("input[type='checkbox'][id^="+"checkbox_"+String(sectionId)+"]");

    //find article for this section
    var article=  allSectionCheckboxes[0].value.split("_")[0];


    openAIPanel(("openAIPanel_"+ sectionId),article,"",("error_"+sectionId));

    });

  $("input[name='annotation'][type='checkbox']").on("click", function () {

    let id = String($(this).prop("id"));
    let values = id.split("_");
    let sectionId = values[1];
    //let rowId = values[2];

    let queryButtonId="querybutton_"+sectionId;
    let queryButton=$('#'+queryButtonId);

    //find all checkboxes in this section
   var allSectionCheckboxes = $("input[type='checkbox'][id^="+"checkbox_"+String(sectionId)+"]");
   
    if(allSectionCheckboxes.filter(":checked").length == 0)
    {
      //if all unchecked disable Suggestion button
      queryButton.attr("disabled", "disabled");
    }
    else
    { //at least one checked - enable Suggestion button
      queryButton.removeAttr("disabled");
    }

/*     if ($(this).prop("checked")) {
      queryButton.removeAttr("disabled");
    } 
    else {
      queryButton.attr("disabled", "disabled");
    } */
  });





  $populateRadio.change();
}

// Open AI panel for radio button. Close all other AI panels.
async function openAIPanel(panelId,article,annotation,errorsectionId) {
  let $openAISection = $("#"+panelId);  
    $('#please_wait').show();
  
    // textarea is empty
    await getOpenAIResponseFromAssistantTEST(article,annotation,errorsectionId);
  
    $openAISection.html(openAIResponse); //populate panel with response from openAI
  
    //Hide Wait, re-enable controls
    $('#please_wait').hide();
  
  $openAISection.show();
  
  }


/* // Open AI panel for radio button. Close all other AI panels.
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
  await getOpenAIResponseFromAssistantTEST(article,annotation,errorsectionId);


  $openAISection.html(openAIResponse); //populate panel with response from openAI

  //Hide Wait, re-enable controls
  $radioButtons.children('input[type=radio]').prop('disabled', false);
  $('#please_wait').hide();
}

$openAISection.show();

} */

// Inserts the suggestion after selected text in the document.
//Remove all paragraphs for this article
async function insertSuggestionAfterArticleHeader() {
  await Word.run(async (context) => {

    try {
      //TODO Need better algorithm to identify header string
    var heading1Regex = "^[0-9]+\."; //String begins with number followed by period
    
    let paragraphs = context.document.body.paragraphs;
    context.load(paragraphs, ['text'],['items']);
    // Synchronize the document state by executing the queued commands
    await context.sync();

    for (let i = 0; i < paragraphs.items.length; ++i) {
        let item = paragraphs.items[i];

        if (item.text.trim().replace("’","'").toUpperCase() === selectedArticle.trim().replace("’","'").toUpperCase()) {

              let paragraphsToDelete = [];
              
              var firstParagraph= paragraphs.items[i+1];
              
              //first paragraph after header must NOT BE another header, because it could be Table of Content
              if (firstParagraph.text.match(heading1Regex))
              {
                  continue; // it is table of content - look for the next one 
              }

              var nextParagraph;
              var j=i+2;

              while (true) {
                nextParagraph = paragraphs.items[j];
                if (nextParagraph.text.match(heading1Regex))
                {
                  break;
                }
                else
                {
                  paragraphsToDelete.push(nextParagraph);
                }
                j=j+1;
            }
             
            for (const p of paragraphsToDelete) {
              p.delete();
          }
          
                await context.sync();

                firstParagraph.insertHtml(openAIResponse, 'Replace');
                await context.sync();

                break;

              }

    }
    await context.sync()
  }
 catch (error) {
    console.log(error);
  }
  });
}

// Inserts the suggestion after selected text in the document.
async function insertSuggestion() {
  await Word.run(async (context) => {
    const radioId = $("input[name='citation'][type='radio']:checked").attr('id');
    let $openAISection = $("#openAIPanel"+radioId);
    let suggestion = $openAISection.text();
    
    const doc = context.document;
    const originalRange = doc.getSelection();

    
    originalRange.insertHtml(suggestion, 'After');

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
    searchResults.load('items');
    await context.sync()

    var found;

    //getLast
    for (const p of searchResults.items) {
      found = p;
  }

    found.select();
   
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
