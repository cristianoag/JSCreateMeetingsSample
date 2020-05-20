const msalConfig = {
    auth: {
      clientId: '165c52b9-32cc-48f7-a5c1-676158bbbb0e',
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