/** SlideTemplate
 */

// [START SlideTemplate]

var TEMPLATE_PREFIX = '${';
var TEMPLATE_SUFFIX = '}';
var TEMPLATE_IMAGE = 'IMAGE';

/**
 * @OnlyCurrentDoc Limits the script to only accessing the current presentation.
 */

/**
 * Create a open  menu item.
 * @param {Event} event The open event.
 */
function onOpen(event) {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Open','showSidebar')
      .addItem('picker','showPicker')
      .addToUi();
}

/**
 * Open the Add-on upon install.
 * @param {Event} event The install event.
 */
function onInstall(event) {
  onOpen(event);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var ui = HtmlService
      .createHtmlOutputFromFile('sidebar')
      .setTitle('SlideTemplate');
  SlidesApp.getUi().showSidebar(ui);
}

/**
 * Recursively gets child text elements a list of elements.
 * @param {PageElement[]} elements The elements to get text from.
 * @return {Text[]} An array of text elements.
 */
function getElementTexts(elements) {
  var texts = [];
  elements.forEach(function(element) {
    switch (element.getPageElementType()) {
// This is not working, variables in GROUP elements are not discovered now
//      case SlidesApp.PageElementType.GROUP:
//        element.asGroup().getChildren().forEach(function(child) {
//          texts = texts.concat(getElementTexts(child));
//        });
//        break;
      case SlidesApp.PageElementType.TABLE:
        var table = element.asTable();
        for (var y = 0; y < table.getNumColumns(); ++y) {
          for (var x = 0; x < table.getNumRows(); ++x) {
            texts.push(table.getCell(x, y).getText());
          }
        }
        break;
      case SlidesApp.PageElementType.SHAPE:
        texts.push(element.asShape().getText());
        break;
    }
  });
  return texts;
}

function findAll(regex, sourceString, aggregator) {
  const arr = regex.exec(sourceString);
  if (arr === null) return aggregator;  
  const newString = sourceString.slice(arr.index + arr[0].length);
  return findAll(regex, newString, aggregator.concat([arr[1].slice(2, -1)]));
}

function removeDups(names) {
  var unique = {};
  names.forEach(function(i) {
    if(!unique[i]) {
      unique[i] = true;
    }
  });
  return Object.keys(unique);
}

//function copyAndTemplate(folderId,varList) {
function copyAndTemplate(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  Logger.log(folder.getName());          
  //var file = DriveApp.getFileById(SlidesApp.getActivePresentation(),getId());
  //var newFile = file.makeCopy(file.getName(),folder);
  //folder.addFile(newFile);
  //SlidesApp.openById(newFile.getId());
  //template(varList);
}

function template(varList) {
  Logger.log('template');
  templateSmart(varList);
  Logger.log(varList);
  var presentation = SlidesApp.getActivePresentation();
  for (key in varList) {
    Logger.log(key  + '=' + varList[key]);
    if (varList[key] !== null) presentation.replaceAllText(TEMPLATE_PREFIX + key + TEMPLATE_SUFFIX, varList[key], true);
  }
}

function collectVars() {
  var presentation = SlidesApp.getActivePresentation();
  //TODO: collect vars from masters and layouts too
  var slides = presentation.getSlides();
  Logger.log("Number of slide" + slides.length);
  //TODO TEMPLATE_PREFIX
  var re = /(\${[A-Za-z0-9]+})/;
  var templateVars = [];
  for (var i = 0; i < slides.length; i++) {
    var slide = presentation.getSlides()[i];    
    var texts = getElementTexts(slide.getPageElements()).forEach(function(text) {
        //Logger.log(typeof text);
        var ptv = [];
        ptv = findAll(re, text.asRenderedString(),ptv);
        for (item in ptv) {
          if (templateVars.indexOf(ptv[item]) === -1) {
            templateVars.push(ptv[item]);
          }
        }
    });
    Logger.log("Slide " + i);
    //Logger.log(templateVars);
  }
  Logger.log(templateVars);
  return templateVars;
}

/**
 * Apply all "smart" templates like IMAGE to each slide
 * @param {string} varList The varList
 */

function templateSmart(varList) {
  Logger.log('templateSmart');
  var presentation = SlidesApp.getActivePresentation();
  
  var masters = presentation.getMasters();
  for (var i = 0; i < masters.length; i++) {
    var replacedElements=[];
    var elements = masters[i].getPageElements().forEach(function(element) {   
     if (element.getPageElementType() ===  SlidesApp.PageElementType.SHAPE) {
       element = element.asShape()
       Logger.log("replace IMAGE master " +element);
       text = element.getText();
       for (key in varList) {
         if ((varList[key] !== null) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX + TEMPLATE_IMAGE)) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX+key))) {
           Logger.log("replace IMAGE master " + i + " " + key  + '=' + varList[key]);
           var image = masters[i].insertImage( varList[key]);
           resizeImage(image,element);
           replacedElements.push(element);
         }
       }
     }
    });
    Logger.log(replacedElements);
    replacedElements.forEach(function(element) {
      element.remove();
    });
  }

  var layouts = presentation.getLayouts();
  for (var i = 0; i < layouts.length; i++) {
    var replacedElements=[];
    var elements = layouts[i].getPageElements().forEach(function(element) {   
     if (element.getPageElementType() ===  SlidesApp.PageElementType.SHAPE) {
       element = element.asShape();
       text = element.getText();
       for (key in varList) {
         if ((varList[key] !== null) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX + TEMPLATE_IMAGE)) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX+key))) {
           Logger.log("replace IMAGE layout " + i + " " + key  + '=' + varList[key]);
           var image = layouts[i].insertImage( varList[key]);
           resizeImage(image,element);
           replacedElements.push(element);
         }
       }
     }
    });
    Logger.log(replacedElements);
    replacedElements.forEach(function(element) {
      element.remove();
    });
  }
  
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; i++) {
    var replacedElements=[];
    var elements = slides[i].getPageElements().forEach(function(element) {   
     if (element.getPageElementType() ===  SlidesApp.PageElementType.SHAPE) {
       element = element.asShape()
       text = element.getText();
       for (key in varList) {
         if ((varList[key] !== null) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX + TEMPLATE_IMAGE)) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX+key))) {
           Logger.log("replace IMAGE slide " + i + " " + key  + '=' + varList[key]);
           var image = slides[i].insertImage( varList[key]);
           resizeImage(image,element);
           replacedElements.push(element);
         }
       }
     }
    });
    Logger.log(replacedElements);
    replacedElements.forEach(function(element) {
      element.remove();
    });
  }
  return;
}

/**
 * Resizes an image to the same size as the element it is replacing
 * @param {Image} image An Image object to resize
 * @param {element} the element (Shape) to size to
 */
function resizeImage(image, element) {
    image.setWidth(element.getWidth());
    image.setHeight(element.getHeight());
    image.setLeft(element.getLeft());
    image.setTop(element.getTop());
}

/**
 * Displays an HTML-service dialog that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('picker.html')
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SlidesApp.getUi().showModalDialog(html, 'Select Folder');
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}
// [END SlideTemplate]