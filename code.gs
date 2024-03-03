const API_KEY = "API_KEY";
const MODEL_TYPE = "gpt-3.5-turbo"; //chatGPT model
const points="3";
var hex_color = '#890eed';
var hex_white="#ffffff"
function onOpen() {
    SlidesApp.getUi().createMenu("Generate")
     .addItem("Generate Presentation", "generatePresentation")
    .addItem("Generate Table of Content", "generateToc")
        .addItem("Generate Slide", "generatePrompt")
        .addItem("Generate Image", "generateImage")
        .addToUi();
SlidesApp.getUi().showDialog("index");


  }
  






function generatePresentation(){
 const presentation = SlidesApp.getActivePresentation();
    const selectedText = presentation.getSelection().getCurrentPage().getShapes()[0].getText().asString();
    const slide = presentation.getSelection().getCurrentPage();
    var prompt = "Generate Table of contents each separated by comma for presentation with 7 slides on topic" + selectedText;
    const temperature = 0.2;
    const maxTokens = 2060;
 // showProgressBar();
    const requestBody = {
      model: MODEL_TYPE,
      messages: [{role: "user", content: prompt}],
      temperature,
      max_tokens: maxTokens,
    };
  
    var requestOptions = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + API_KEY,
      },
      payload: JSON.stringify(requestBody),
    };
  
  

    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
    var responseText = response.getContentText();
    var json = JSON.parse(responseText);
    var generatedText = json['choices'][0]['message']['content'];
    var toc=generatedText.split("\n")
    // slide.insertTextBox(toc.toString);
    Logger.log(generatedText);
     // Logger.log(toc[0]);

 var slideToc=presentation.insertSlide(0);

    slideToc.insertTextBox(generatedText.toString(),0,60,700,400);


      
  
  for(var i=0;i<7;i++){
    var slidePpt=presentation.insertSlide(i+1);
    var titleSlide=slidePpt.insertTextBox(toc[i]);
    




var prompt="Generate very short "+points+" points or short paragraph for presentation on "+toc[i];


 const requestBody = {
      model: MODEL_TYPE,
      messages: [{role: "user", content: prompt}],
      temperature,
      max_tokens: maxTokens,
    };

 var requestOptions = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + API_KEY,
      },
      payload: JSON.stringify(requestBody),
    };

 var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
    var responseText = response.getContentText();
    var json = JSON.parse(responseText);
    var generatedText = json['choices'][0]['message']['content'];
    var bg = slideToc.getBackground();



 // bg.setSolidFill(hex_color);
var detailsSlide=slidePpt.insertTextBox(generatedText.toString(),0,50,700,400);


  // Set the foreground color of the text to red

Utilities.sleep(20000);
  
  }
  }

function generateToc() {
    const presentation = SlidesApp.getActivePresentation();
    const selectedText = presentation.getSelection().getCurrentPage().getShapes()[0].getText().asString();
    const slide = presentation.getSelection().getCurrentPage();
    const prompt = "Generate short Table of contents for presentation with 7 slides on " + selectedText;
  
    const temperature = 0.5;
    const maxTokens = 2060;
     // var hex_color = '#890eed'; //The background color set in dark gray
  var selection = SlidesApp.getActivePresentation().getSelection();
   var currentPage = selection.getCurrentPage();

    const requestBody = {
      model: MODEL_TYPE,
      messages: [{role: "user", content: prompt}],
      temperature,
      max_tokens: maxTokens,
    };
  
    const requestOptions = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + API_KEY,
      },
      payload: JSON.stringify(requestBody),
    };
  

  
  
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
    const responseText = response.getContentText();
    const json = JSON.parse(responseText);
    const generatedText = json['choices'][0]['message']['content'];
    Logger.log(generatedText);
    const slideToc =presentation.insertSlide(1)

      const bg = slideToc.getBackground();
 // bg.setSolidFill(hex_color);
    const titleToc=slideToc.insertTextBox("Table of Contents")
    titleToc.setTitle("Table of Content")
slideToc.insertTextBox(generatedText.toString(),0,50,700,400);

   // slide.insertTextBox(generatedText.toString(),0,60,700,400);
  }

  function generateTweets() {
    const presentation = SlidesApp.getActivePresentation();
    const selectedText = presentation.getSelection().getCurrentPage().getShapes()[0].getText().asString();
    const slide = presentation.getSelection().getCurrentPage();
    const prompt = "Generate 3 Tweets on " + selectedText;
    const temperature = 0;
    const maxTokens = 20;
  
    const requestBody = {
      model: MODEL_TYPE,
      messages: [{role: "user", content: prompt}],
      temperature,
      max_tokens: maxTokens,
    };
  
    const requestOptions = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + API_KEY,
      },
      payload: JSON.stringify(requestBody),
    };
  
  

    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
    const responseText = response.getContentText();
    const json = JSON.parse(responseText);
    const generatedText = json['choices'][0]['message']['content'];
    Logger.log(generatedText);
    slide.insertTextBox(generatedText.toString());

}
function generateImage(){
   const presentation = SlidesApp.getActivePresentation();
    const selectedText = presentation.getSelection().getCurrentPage().getShapes()[0].getText().asString();
    const slide = presentation.getSelection().getCurrentPage();
    const prompt = selectedText;

 temperature= 0
  maxTokens = 2500
  var imageLinks=[]

  //  var prompt = data[i][0]
    const requestBody2 = {
      "prompt": prompt,
      "n": 1,
      "size": "512x512"
    };

    const requestOptions2 = {
      "method": "POST",
      "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer "+API_KEY
      },
      "payload": JSON.stringify(requestBody2)
    }
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/images/generations", requestOptions2);
    // Parse the response and get the generated text
    var responseText = response.getContentText();
    var json = JSON.parse(responseText);
    var url1=json['data'][0]['url']
   

    var im1=UrlFetchApp.fetch(url1).getContent()
    var blob1 = Utilities.newBlob(im1, 'image/png', data[i][0]+'-Image1');
    //slide.insertImage(blob1);
    var save1 = DriveApp.createFile(blob1);
    var imgUrl1 = save1.getUrl();
     slide.insertTextBox(imgUrl1);



}

 function generatePrompt() {
    const presentation = SlidesApp.getActivePresentation();
    const selectedText = presentation.getSelection().getCurrentPage().getShapes()[0].getText().asString();
    const slide = presentation.getSelection().getCurrentPage();
    const prompt = "Generate very short bullet points for presentation on " + selectedText;
    const temperature = 0.5;
    const maxTokens = 2060;
  
    const requestBody = {
      model: MODEL_TYPE,
      messages: [{role: "user", content: prompt}],
      temperature,
      max_tokens: maxTokens,
    };
  
    const requestOptions = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + API_KEY,
      },
      payload: JSON.stringify(requestBody),
    };
  

  
  
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
    const responseText = response.getContentText();
    const json = JSON.parse(responseText);
    const generatedText = json['choices'][0]['message']['content'];
    Logger.log(generatedText);
    slide.insertTextBox(generatedText.toString(),0,60,700,400);
  }

function showProgressBar() {
  var html = HtmlService.createHtmlOutput('<progress value="0" max="100"></progress>');
  var dialog = SlidesApp.getUi().showModalDialog(html, 'Progress');

  // Update the progress bar every second
  var progress = 0;
  var intervalId = setInterval(function() {
    progress += 10;
    html.setContent('<progress value="' + progress + '" max="100"></progress>');
    if (progress >= 100) {
      clearInterval(intervalId);
      dialog.close();
    }
  }, 1000);
}


function generateToc() {
    const presentation = SlidesApp.getActivePresentation();
    const selectedText = presentation.getSelection().getCurrentPage().getShapes()[0].getText().asString();
    const slide = presentation.getSelection().getCurrentPage();
    const prompt = "Generate short Table of contents for presentation with 7 slides on " + selectedText;
    const temperature = 0.5;
    const maxTokens = 2060;
      var hex_color = '#890eed'; //The background color set in dark gray
  var selection = SlidesApp.getActivePresentation().getSelection();
   var currentPage = selection.getCurrentPage();

    const requestBody = {
      model: MODEL_TYPE,
      messages: [{role: "user", content: prompt}],
      temperature,
      max_tokens: maxTokens,
    };
  
    const requestOptions = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + API_KEY,
      },
      payload: JSON.stringify(requestBody),
    };
  

  
  
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
    const responseText = response.getContentText();
    const json = JSON.parse(responseText);
    const generatedText = json['choices'][0]['message']['content'];
    Logger.log(generatedText);
    const slideToc =presentation.insertSlide(1)

      const bg = slideToc.getBackground();
  bg.setSolidFill(hex_color);
    const titleToc=slideToc.insertTextBox("Table of Contents")
    titleToc.setTitle("Table of Content")
slideToc.insertTextBox(generatedText.toString(),0,50,700,400);

   // slide.insertTextBox(generatedText.toString(),0,60,700,400);
  }

  function generateTweets() {
    const presentation = SlidesApp.getActivePresentation();
    const selectedText = presentation.getSelection().getCurrentPage().getShapes()[0].getText().asString();
    const slide = presentation.getSelection().getCurrentPage();
    const prompt = "Generate 3 Tweets on " + selectedText;
    const temperature = 0;
    const maxTokens = 20;
  
    const requestBody = {
      model: MODEL_TYPE,
      messages: [{role: "user", content: prompt}],
      temperature,
      max_tokens: maxTokens,
    };
  
    const requestOptions = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + API_KEY,
      },
      payload: JSON.stringify(requestBody),
    };
  
  
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
    const responseText = response.getContentText();
    const json = JSON.parse(responseText);
    const generatedText = json['choices'][0]['message']['content'];
    Logger.log(generatedText);
    slide.insertTextBox(generatedText.toString());
  }
