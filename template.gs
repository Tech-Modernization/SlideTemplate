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
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; i++) {
    var replacedElements=[];
    var slide = presentation.getSlides()[i];
    var elements = slide.getPageElements().forEach(function(element) {   
     if (element.getPageElementType() ===  SlidesApp.PageElementType.SHAPE) {
       element = element.asShape()
       text = element.getText();
       for (key in varList) {
         if ((varList[key] !== null) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX + TEMPLATE_IMAGE)) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX+key))) {
           Logger.log("replace IMAGE slide " + i + " " + key  + '=' + varList[key]);
           replaceImageTemplate(slide, varList[key], element);
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
 * Creates a image on the current slide from the given link, replacing the template text box
 * @param {string} imageUrl A String object representing an image URL
 * @param {text} the text box to replace
 */
function replaceImageTemplate(slide, imageUrl, element) {
    var image = slide.insertImage(imageUrl);
    image.setWidth(element.getWidth());
    image.setHeight(element.getHeight());
    image.setLeft(element.getLeft());
    image.setTop(element.getTop());
}
// [END SlideTemplate]