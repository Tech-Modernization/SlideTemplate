/** SlideTemplate
 */

// [START SlideTemplate]
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
      .addItem('Help','showHelp')
      .addToUi();
  //TODO:  call updateVars() script
  //SlidesApp.getUi().Button.('#run-reload').click(updateVars);
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
      .setTitle('Template');
  SlidesApp.getUi().showSidebar(ui);
}

/**
 * Opens a dialogbox with help.
 */
function showHelp() {
  var ui = SlidesApp.getUi();
  var result = ui.alert(
    'Provides a way to templatize slides using variables.',
    'Variables like ${XXX} are globally replaced.',
    ui.ButtonSet.OK);
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
  Logger.log(varList);
  var presentation = SlidesApp.getActivePresentation();
  for (key in varList) {
    Logger.log(key  + '=' + varList[key]);
    if (varList[key] !== null) presentation.replaceAllText('${' + key + '}', varList[key], true);
  }
}

function collectVars() {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  Logger.log("Number of slide" + slides.length);
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
// [END SlideTemplate]
