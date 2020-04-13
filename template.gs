/** SlideTemplate
 */

// [START SlideTemplate]

//var TEMPLATE_PREFIX = '{{';
//var TEMPLATE_SUFFIX = '}}';
var TEMPLATE_PREFIX = '${';
var TEMPLATE_SUFFIX = '}';
var TEMPLATE_IMAGE = 'IMAGE';

//TODO: not used here yet, in future save/recall all variable values under description and for searching
var TAG_PREFIX = '#';
var TAG_SUFFIX = "_";

/**
 * (AT) OnlyCurrentDoc Limits the script to only accessing the current presentation.
 */

/**
 * Create a open  menu item.
 * @param {Event} event The open event.
 */
function onOpen(event) {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Open','showSidebar')
      // no picker
      //.addItem('picker','showPicker')
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
       case SlidesApp.PageElementType.GROUP:
        var grp = element.asGroup();
        for (var child in grp.getChildren()) {
          try {
            texts = texts.concat(getElementTexts(child));
          } catch (err) {
            //This is not working, variables in GROUP elements are not discovered in this case
            Logger.log(err);
          }
        }
        break;
      case SlidesApp.PageElementType.TABLE:
        var table = element.asTable();
        for (var y = 0; y < table.getNumColumns(); ++y) {
          for (var x = 0; x < table.getNumRows(); ++x) {
            try {
              texts.push(table.getCell(x, y).getText());
            } catch (err) {
                //This operation is only allowed on the head (upper left) cell of the merged cells
                Logger.log(err);
            }
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
  var re = new RegExp(regex.replace("$", "\\$"));
  const arr = re.exec(sourceString);
  if (arr === null) return aggregator;  
  const newString = sourceString.slice(arr.index + arr[0].length);
  return findAll(regex, newString, aggregator.concat([arr[1].slice(TEMPLATE_PREFIX.length, -(TEMPLATE_SUFFIX.length)).trim()]));
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

function copyAndOpen(folderId) {
  if (folderId.length > 0) {
      var folder = DriveApp.getFolderById(folderId);
      Logger.log(folder.getName());    
      var presentation = SlidesApp.getActivePresentation();
      var id = presentation.getUrl().split("id=").pop();
      Logger.log(id);
      var file = DriveApp.getFileById(id);
      Logger.log(file.getName());          
      var newFile = file.makeCopy(file.getName(),folder);
      folder.addFile(newFile);
      Logger.log(newFile.getUrl());          
      SlidesApp.openByUrl(newFile.getUrl());
  }
}

function getCurrentFile() {
  var presentation = SlidesApp.getActivePresentation();
  var id = presentation.getUrl().split("id=").pop();
  Logger.log(id);
  try {
    var file = DriveApp.getFileById(id);
  } catch (err) {
    Logger.log(err);
  }
  return file;
}

function templateDescriptionVars(varList) {
  var file = getCurrentFile();
  var desc = file.getDescription();
  for (key in varList) {
    k = key.toString();
    if (!k.startsWith(TEMPLATE_IMAGE)) {
      if (varList[key] !== null) desc.replaceAll 
      desc = desc.replace(TEMPLATE_PREFIX + key + TEMPLATE_SUFFIX, varList[key]);
    }
  }
  file.setDescription(desc);
}

function template(varList) {  
  Logger.log(varList);
  templateDescriptionVars(varList);
  templateSmart(varList);
  var presentation = SlidesApp.getActivePresentation();
  for (key in varList) {
    k = key.toString();
    if (!k.startsWith(TEMPLATE_IMAGE)) {
      //Logger.log(k  + '=' + varList[key]);
      if (varList[key] !== null) presentation.replaceAllText(TEMPLATE_PREFIX + key + TEMPLATE_SUFFIX, varList[key], true);
    }
  }
}

function collectVars() {
  var re = "(" + TEMPLATE_PREFIX + "[A-Za-z0-9_ ]+" + TEMPLATE_SUFFIX + ")";
  var templateVars = [];
  templateVars = findAll(re, getCurrentFile().getDescription(),templateVars);  
  Logger.log(templateVars);
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  Logger.log("Number of slide" + slides.length);
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
    //Logger.log("Slide " + i);
  }
  //TODO: collect vars from masters and layouts too
  var masters = presentation.getMasters();
  for (var i = 0; i < masters.length; i++) {
    var texts = getElementTexts(masters[i].getPageElements()).forEach(function(text) {
        //Logger.log(typeof text);
        var ptv = [];
        ptv = findAll(re, text.asRenderedString(),ptv);
        for (item in ptv) {
          if (templateVars.indexOf(ptv[item]) === -1) {
            templateVars.push(ptv[item]);
          }
        }
    });
    //Logger.log("master " + i);
  }
  var layouts = presentation.getLayouts();
  for (var i = 0; i < layouts.length; i++) {
    var texts = getElementTexts(layouts[i].getPageElements()).forEach(function(text) {
        //Logger.log(typeof text);
        var ptv = [];
        ptv = findAll(re, text.asRenderedString(),ptv);
        for (item in ptv) {
          if (templateVars.indexOf(ptv[item]) === -1) {
            templateVars.push(ptv[item]);
          }
        }
    });
    //Logger.log("layout " + i);
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
       try {
         text = element.getText();
         for (key in varList) {
           if ((key.startsWith(TEMPLATE_IMAGE)) && (varList[key] != null) && (varList[key].length > 0) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX + TEMPLATE_IMAGE)) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX+key))) {
             Logger.log("replace IMAGE master " + i + " " + key  + '=' + varList[key]);
             var image = masters[i].insertImage(varList[key]);
             resizeImage(image,element);
             replacedElements.push(element);
           }
         }
       } catch (err) {
         Logger.log(err);
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
       try {
         text = element.getText();
         Logger.log(text.asRenderedString());
         for (key in varList) {
           if ((key.startsWith(TEMPLATE_IMAGE)) && (varList[key] !== null) && (varList[key].length > 0) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX + TEMPLATE_IMAGE)) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX+key))) {
             Logger.log("replace IMAGE layout " + i + " " + key  + '=' + varList[key]);
             var image = layouts[i].insertImage( varList[key]);
             resizeImage(image,element);
             replacedElements.push(element);
           }
         }
       } catch (err) {
         Logger.log(err);
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
       element = element.asShape();
       text = element.getText();
       try {
         for (key in varList) {
           if ((key.startsWith(TEMPLATE_IMAGE)) && (varList[key] !== null) && (varList[key].length > 0) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX + TEMPLATE_IMAGE)) && (text.asRenderedString().startsWith(TEMPLATE_PREFIX+key))) {
             Logger.log("replace IMAGE slide " + i + " " + key  + '=' + varList[key]);
             var image = slides[i].insertImage( varList[key]);
             resizeImage(image,element);
             replacedElements.push(element);
           }
         }
       } catch (err) {
         Logger.log(err);
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
