(function () {
  "use strict";
  
  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
      $(document).ready(setInitialDisplay());
  };

})();

function getTaskUrl() {
  return Office.context.mailbox.item.getRegExMatchesByName("TaskUrl")[0];
}

function getDomainWithSuite(taskUrl) {
  taskUrl = taskUrl || getTaskUrl()
  return taskUrl.substr(0, taskUrl.indexOf("/tempo") + 1);
}

function setInitialDisplay() {

    /* Get task url and extract domain & task ID */
    var taskUrl = getTaskUrl();

    var tasksTextIndex = taskUrl.indexOf("tasks");
    var urlSplitBySlash = taskUrl.split("/");
    var taskId = urlSplitBySlash[urlSplitBySlash.length - 1];
        
    /* Add appian JS script to head */
    var script = document.createElement("script");
    script.type = "text/javascript";
    script.src = taskUrl.substr(0, tasksTextIndex) + "ui/sail-client/embeddedBootstrap.nocache.js";
    script.setAttribute("id", "appianEmbedded");
    script.innerHTML = null;
    document.head.appendChild(script);

    var newTask = generateAndInsertTaskTag(taskId);
    waitForSignIn($(newTask));
}

function generateAndInsertTaskTag(taskId) {
  var newTask = document.createElement('appian-task');
  newTask.setAttribute("id", "new-task");
  newTask.setAttribute("taskId", taskId);
  document.body.appendChild(newTask);
  
  return newTask;
}

function waitForSignIn(wrappedNewTask) {
  function waitForEmbeddedSignIn() {
    $.ajax({
      type: 'GET',
      url: getDomainWithSuite() + "auth?appian_environment=tempo",
      contentType: 'text/plain',
      xhrFields: {
        withCredentials: true
      },
      statusCode: {
        403: function(){
          $("#sign-out-link").show();
          wrappedNewTask.off('DOMSubtreeModified', waitForEmbeddedSignIn);
        }
      }
    });
  }
  wrappedNewTask.on('DOMSubtreeModified', waitForEmbeddedSignIn);
}

function signOut() {
    var domainWithSuite = getDomainWithSuite();
    var logoutUrl = domainWithSuite + "logout";
    
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.open("GET", logoutUrl, true); // false for synchronous request
    xmlHttp.withCredentials = true;
    xmlHttp.send(null);

    var ua = window.navigator.userAgent;
    var msie = ua.indexOf("MSIE ");
    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))  // If Internet Explorer, return version number
    {
        // Internet explorer breaks iframes if location.reload is used, so for IE use a hidden link to go back to the login screen
        document.getElementById('logoutlink').click();
    }
    else  // If another browser, return 0
    {
        window.location.reload(true);
    }
}
