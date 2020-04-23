//// <reference path="../node_modules/officejs.dialogs/dialogs.js" />
/// <reference path="../node_modules/easyews/easyews.js" />
var dialogs = require('officejs.dialogs')

const restrictedDomains = ["outlook.com", "gmail.com"];
var sendEvent;
var addinName = "";
var item;
var isRestricted = false;
var thisUserDomainByName = "";


Office.initialize = function (reason) {
  // init here
  item = Office.context.mailbox.item;
  domain = getDomain(thisUser);
  addinName = "Bam Alert";
};


function onSendEvent(event) {
  sendEvent = event; // grab this so it does not get cleaned up
  var thisUserDomainByEmail;
  var thisUserDomainByName;
  isRestricted = false;

  // Get recipients
  Office.context.mailbox.item.to.getAsync({
    asyncContext: event
  }, function (asyncResult) {
    asyncResult.value.forEach(function (recip, index) {
      thisUserDomainByEmail = getDomain(recip.emailAddress)
      thisUserDomainByName = getDomain(recip.displayName)

      // Check if sending to restricted domains
      if (restrictedDomains.indexOf(thisUserDomainByEmail) > -1 || restrictedDomains.indexOf(thisUserDomainByName) > -1) {
        isRestricted = true;
      }

      // If sending to restricted domain...
      if (isRestricted) {

        // Pop an alert
        MessageBox.Show("Are you sure you want to send this mail?", "You're sending mail to restricted domain", MessageBoxButtons.YesNo, MessageBoxIcons.Stop, false, null, function (checkedButton) {

          // Check if send message approved by user
          if (checkedButton == "Yes") {

            // Get the subject
            item.subject.getAsync({
              asyncContext: sendEvent
            }, function (asyncResult) {
              if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                console.log(asyncResult.error.message);
              } else {
                //var val = asyncResult.value;
                console.log(asyncResult.value, "line 83", asyncResult.value)
                console.log("bitchhhh")
                // Check if word to add is already in the subject
                // var checkSubject = (new RegExp(/\[אושר\]/)).test(asyncResult.value);
                // console.log(/\[אושר\]/.test(asyncResult.value));
                // if (!checkSubject) {
                //   // Add [אושר]: to subject line.
                //   var subject = '[אושר]: ' + asyncResult.value;
                //   console.log("no keyword in subject")

                //   mailboxItem.subject.setAsync(subject, {
                //     asyncContext: sendEvent
                //   }, function (asyncResult) {
                //     if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                //       console.log("couldnt cahgne the subject");
                //     } else {
                //       console.log("subject changed")
                //     }

                //   });
                // }

              }
            });

            // Allow sending message
            sendEvent.completed({ allowEvent: true });
          } else {
            console.log("dont send");

            // Notification Message (error)
            Office.context.mailbox.item.notificationMessages.addAsync("error", { type: "errorMessage", message: "ההודעה לא נשלחה" });

            // Prevent email from being sent
            sendEvent.completed({ allowEvent: false });
          }
        });
      }
    });
  });
}

function changeSubject(event) {
  item.subject.getAsync({
    asyncContext: event
  }, function (asyncResult) {
    // Match string.
    var checkSubject = (new RegExp(/\[אושר\]/)).test(asyncResult.value);

    // Add [אושר]: to subject line.
    subject = '[אושר]: ' + asyncResult.value;

    if (!checkSubject) {
      subjectOnSendChange(subject, asyncResult.asyncContext);
    }
  })
}


function subjectOnSendChange(subject, event) {
  mailboxItem.subject.setAsync(subject, {
    asyncContext: event
  }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      mailboxItem.notificationMessages.addAsync('NoSend', {
        type: 'errorMessage',
        message: 'Unable to set the subject.'
      });

      // Block send.
      //asyncResult.asyncContext.completed({ allowEvent: false });
    } else {
      console.log("subject changed")
      // Allow send.
      //asyncResult.asyncContext.completed({ allowEvent: true });
    }

  });
}




/**
 * Displays an error to the user
 * @param {string} error 
 */
function showError(error, callback) {
  // uses the OfficeJS.dialogs Alert. See:
  // https://github.com/davecra/OfficeJS.dialogs
  // an error occurred trying to get all the emails on To/CC/BCC
  Alert.Show("Unable to process TO/CC/BCC: " + error, function () {
    // Notification Message (error)
    Office.context.mailbox.item.notificationMessages.addAsync("error", {
      type: "errorMessage",
      message: "The Outlook Demo add-in failed to process this message."
    });
  }, callback); // Alert.Show
}


function showInformation(msg) {
  item.notificationMessages.addAsync("information", {
    type: "informationalMessage",
    message: msg,
    icon: "icon16",
    persistent: false
  });
}

/**
 * Gets the domain portion of an email address. For example:
 *  - user@exchange.contoso.com = contoso.com
 *  - user@constoso.com = contoso.com
 * @param {string} user The email address of the user
 * @returns {string} domain name returned
 */
function getDomain(user) {
  /** @type {string} */
  var fullDomain = user.split("@")[1];
  /** @type {string[]} */
  var parts = fullDomain.split(".");
  /** @type {string} */
  var domain = parts[0] + "." + parts[1];
  if (parts.length > 2) {
    domain = parts[parts.length - 2] + "." + parts[parts.length - 1];
  }
  return domain;
}