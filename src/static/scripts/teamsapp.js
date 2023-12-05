(function () {
  "use strict";

  // Call the initialize API first
  microsoftTeams.app.initialize().then(function () {
    microsoftTeams.app.getContext().then(function (context) {
      if (context?.app?.host?.name) {
        updateHubState(context.app.host.name);
      }
      microsoftTeams.authentication.authenticate({url:'https://app.klyck.io/auth/realms/mettler/protocol/openid-connect/auth?response_type=code&client_id=whut-frontend&redirect_uri=https://klyckio.azurewebsites.net/klyckIoMessageExtension/token.html'});
    });
  });

  function updateHubState(hubName) {
    if (hubName) {
      document.getElementById("hubState").innerHTML = "in " + hubName;
    }
  }
})();
