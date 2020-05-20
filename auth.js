
const msalClient = new Msal.UserAgentApplication(msalConfig);

if (msalClient.getAccount() && !msalClient.isCallback(window.location.hash)) {
   updatePage(msalClient.getAccount(), Views.home);
}

async function signIn() {
    try {
      await msalClient.loginPopup(loginRequest);
      console.log('id_token acquired at: ' + new Date().toString());
      if (msalClient.getAccount()) {
        updatePage(msalClient.getAccount(), Views.home);
      }
    } catch (error) {
      console.log(error);
      updatePage(null, Views.error, {
        message: 'Error logging in',
        debug: error
      });
    }
  }
  
  function signOut() {
    msalClient.logout();
  }