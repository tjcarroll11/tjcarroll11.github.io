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
    script.src = taskUrl.substr(0, tasksTextIndex) + "tempo.nocache.js";
    script.innerHTML = null;
    document.head.appendChild(script);

    var newTask = generateAndInsertTaskTag(taskId);
    waitForLoad(newTask);
}

function generateAndInsertTaskTag(taskId) {
  var newTask = document.createElement('appian-task');
  newTask.setAttribute("id", "new-task");
  newTask.setAttribute("taskId", taskId);
  newTask.style.display = "none";
  newTask.addEventListener("submit", handleSubmit, false);
  document.body.appendChild(newTask);
  
  return newTask;
}

function waitForLoad(newTask) {
  var wrappedNewTask = $(newTask);
  
  function waitForEmbeddedContent(){
    var links = wrappedNewTask.find('a');
    if (wrappedNewTask.find('form').length > 0 || links.length > 0) {
      viewTaskOrSignIn(wrappedNewTask, links);
      newTask.removeEventListener('DOMSubtreeModified', waitForEmbeddedContent);
    }
  }
  
  newTask.addEventListener('DOMSubtreeModified', waitForEmbeddedContent);
}

function viewTaskOrSignIn(wrappedNewTask, links) {
    if (links.length == 1) {
        var link = links[0];
        if (link.innerHTML == "Sign In") {
            function waitForEmbeddedSignIn() {
              var xmlHttp = new XMLHttpRequest();
              xmlHttp.withCredentials = true;
              xmlHttp.open("GET", getDomainWithSuite() + "auth?appian_environment=tempo"); // false for synchronous request
              xmlHttp.onreadystatechange = function(){
                if (xmlHttp.readyState == 4){
                  if (xmlHttp.status == 403) {
                    showSignOutLink();
                    wrappedNewTask.off('DOMSubtreeModified', waitForEmbeddedSignIn);
                  }
                };
              };
              xmlHttp.send(null);
            }
            wrappedNewTask.on('DOMSubtreeModified', waitForEmbeddedSignIn);
            link.click();
            wrappedNewTask.show();
        } else {
            viewTask(wrappedNewTask);
        }
    } else {
        viewTask(wrappedNewTask);
    }
}

function showSignOutLink() {
  $("#sign-out-link").show();
}

function viewTask(wrappedNewTask) {
    wrappedNewTask.show();
    showSignOutLink();
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

/* This function is called by the submit event listener */
function handleSubmit() {
    alert("The task has been submitted!");
}
