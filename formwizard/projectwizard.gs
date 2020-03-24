//This will post the Google Form as a json object to a google cloud function

// Replace with the URL to your deployed Cloud Function
var url = "https://us-central1-slidetemplate.cloudfunctions.net/form-trigger";

// This function will be called when the form is submitted
function onSubmit(event) {

  // The event is a FormResponse object:
  // https://developers.google.com/apps-script/reference/forms/form-response
  var formResponse = event.response;

  // Gets all ItemResponses contained in the form response
  // https://developers.google.com/apps-script/reference/forms/form-response#getItemResponses()
  
  var itemResponses = formResponse.getItemResponses();

  // Gets the actual response strings from the array of ItemResponses, just the values as a string[]
  //var responses = itemResponses.map(function getResponse(e) { return e.getItem().getTitle().replace(/\s+/g, '')+e.getResponse(); });

  var responses = new Object();
  itemResponses.forEach(function(e) { responses[e.getItem().getTitle().replace(/\s+/g, '')] = e.getResponse();  });
    
  // Post the payload as JSON to our Cloud Function  
  UrlFetchApp.fetch(
    url,
    {
      "method": "POST",
      "payload": JSON.stringify({
        "responses": responses
      })
    }
  );

}