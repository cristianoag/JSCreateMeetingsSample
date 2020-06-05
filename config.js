const msalConfig = {
    auth: {
      clientId: 'XXXXXX',
      redirectUri: 'http://localhost:8080'
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
      forceRefresh: false
    }
  };
  
  const loginRequest = {
    scopes: [
      'openid',
      'profile',
      'user.read',
      'onlineMeetings.ReadWrite'
    ]
  }