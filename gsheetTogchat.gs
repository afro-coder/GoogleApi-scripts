// This demonstrates a Google On edit function which then posts to a Google Webhook, no authentication needed its all within the org
// https://dev.to/afrocoder/how-google-sheets-and-webhooks-made-my-life-easier-a05


// Function was taken from
// https://joeybronner.fr/blog/google-apps-script-get-current-user-email-from-a-spreadsheet-add-on/
function getCurrentUserEmail() {
    var userEmail = Session.getActiveUser().getEmail();
    if (userEmail === '' || !userEmail || userEmail === undefined) {
        userEmail = PropertiesService.getUserProperties().getProperty('userEmail');
        if (!userEmail) {
            var protection = SpreadsheetApp.getActive().getRange('A1').protect();
            protection.removeEditors(protection.getEditors());
            var editors = protection.getEditors();
            if (editors.length === 2) {
                var owner = SpreadsheetApp.getActive().getOwner();
                editors.splice(editors.indexOf(owner), 1);
            }
            userEmail = editors[0];
            protection.remove();
            PropertiesService.getUserProperties().setProperty('userEmail', userEmail);
        }
    }
    Logger.log(userEmail);
    return userEmail;
}

function myEditNew(e) {
// The e is an event Object
//https://developers.google.com/apps-script/guides/triggers/events

if (e.value)
{
Logger.log(e.range);

Logger.log("Data: "+e.value);

var val=e.value;
// This function is to get the current user that created the task its documented in the Gist at the end of the post
var cur_user=getCurrentUserEmail();

var text = Utilities.formatString('Incident created by *%s*: ```%s``` <users/all>',cur_user,val);
}

//1. Bot Test room
// [Args[Webhook URL, Space Name, Thread Name]]

// To get the Thread ID read this Stackoverflow answer

// https://webapps.stackexchange.com/questions/117392/get-link-to-specific-conversation-thread-and-or-message-in-a-chat-room-in-google

var urls = [
  [
   "https://chat.googleapis.com/v1/spaces/room_id/messages?key=Webhook_ID",
   "spaces/room_id/messages/space_name.space_name",
   "spaces/Thread_ID/threads/Thread_ID"
   ],
  
];

// Lots of Logging to see if Things are working properly.
Logger.log(e.range.getA1Notation());
Logger.log(e.range.getColumn());
Logger.log(val);


if (e.range.getA1Notation().startsWith('C') && val != "")
{
// Synchronously Post it to the Rooms
for(i=0;i<urls.length;i++)
{
var payload={
  "name":urls[i][1],
  "thread":{
    "name":urls[i][2]
  },
  "text":text
};
var options = {
  'method' : 'post',
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(payload)
};
Logger.log(options);

// UrlFetchApp Documentation
// https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app
var response = UrlFetchApp.fetch(urls[i][0], options);
Logger.log(response);

}

}
}
