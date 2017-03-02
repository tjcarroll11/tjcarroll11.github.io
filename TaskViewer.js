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
}

function generateAndInsertTaskTag(taskId) {
  var newTask = document.createElement('appian-task');
  newTask.setAttribute("id", "new-task");
  newTask.setAttribute("taskId", taskId);
  document.body.appendChild(newTask);
  
  return newTask;
}
