(function () {
  "use strict";

  // Call the initialize API first
  microsoftTeams.app.initialize().then(function () {
    microsoftTeams.app.getContext().then(function (context) {
      if (context?.app?.host?.name) {
        updateHubState(context.app.host.name);
      }
    });
  });

  function updateHubState(hubName) {
    if (hubName) {
      document.getElementById("hubState").innerHTML = "in " + hubName;
    }
  }
})();

function doAuth() {
  microsoftTeams.authentication.authenticate({
    url: window.location.origin + "/static/auth-start.html",
    width: 600,
    height: 535,
    isExternal: true,
  }).then ((result) => {
    document.getElementById("authResult").innerHTML = "Success: " + result;
  }).catch((reason) => {
    console.log("Error: " + reason);
  });
}
