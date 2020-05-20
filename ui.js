const authenticatedNav = document.getElementById('authenticated-nav');
const accountNav = document.getElementById('account-nav');
const mainContainer = document.getElementById('main-container');

const Views = { error: 1, home: 2, meeting: 3};

function createElement(type, className, text) {
  var element = document.createElement(type);
  element.className = className;

  if (text) {
    var textNode = document.createTextNode(text);
    element.appendChild(textNode);
  }

  return element;
}

function showAuthenticatedNav(account, view) {
  authenticatedNav.innerHTML = '';

  if (account) {
    var MeetingNav = createElement('li', 'nav-item');

    var MeetingLink = createElement('button',
      `btn btn-secondary nav-link${view === Views.meeting ? ' active' : '' }`,
      'Create Meeting');
    
    MeetingLink.setAttribute('onclick', 'createMeeting();');
    MeetingNav.appendChild(MeetingLink);

    authenticatedNav.appendChild(MeetingNav);
  }
}

function showAccountNav(account) {
  accountNav.innerHTML = '';

  if (account) {
    // Show the "signed-in" nav
    accountNav.className = 'nav-item dropdown';

    var dropdown = createElement('a', 'nav-link dropdown-toggle');
    dropdown.setAttribute('data-toggle', 'dropdown');
    dropdown.setAttribute('role', 'button');
    accountNav.appendChild(dropdown);

    var userIcon = createElement('i',
      'far fa-user-circle fa-lg rounded-circle align-self-center');
    userIcon.style.width = '32px';
    dropdown.appendChild(userIcon);

    var menu = createElement('div', 'dropdown-menu dropdown-menu-right');
    dropdown.appendChild(menu);

    var userName = createElement('h5', 'dropdown-item-text mb-0', account.name);
    menu.appendChild(userName);

    var userEmail = createElement('p', 'dropdown-item-text text-muted mb-0', account.userName);
    menu.appendChild(userEmail);

    var divider = createElement('div', 'dropdown-divider');
    menu.appendChild(divider);

    var signOutButton = createElement('button', 'dropdown-item', 'Sign out');
    signOutButton.setAttribute('onclick', 'signOut();');
    menu.appendChild(signOutButton);
  } else {
    // Show a "sign in" button
    accountNav.className = 'nav-item';

    var signInButton = createElement('button', 'btn btn-link nav-link', 'Sign in');
    signInButton.setAttribute('onclick', 'signIn();');
    accountNav.appendChild(signInButton);
  }
}

function showWelcomeMessage(account) {
  var jumbotron = createElement('div', 'jumbotron');

  var heading = createElement('h1', null, 'Create Teams Meetings via Graph');
  jumbotron.appendChild(heading);

  var lead = createElement('p', 'lead',
    'This sample app shows how to create Teams Meetings using Microsoft Graph API ' +
    ' from JavaScript.');
  jumbotron.appendChild(lead);

  if (account) {
    var welcomeMessage = createElement('h4', null, `Welcome ${account.name}!`);
    jumbotron.appendChild(welcomeMessage);

    var callToAction = createElement('p', null,
      'Use the Create Meeting button at the top of the page to create meetings.');
    jumbotron.appendChild(callToAction);
  } else {
    var signInButton = createElement('button', 'btn btn-primary btn-large',
      'Click here to sign in');
    signInButton.setAttribute('onclick', 'signIn();')
    jumbotron.appendChild(signInButton);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(jumbotron);
}

function showError(error) {
  var alert = createElement('div', 'alert alert-danger');

  var message = createElement('p', 'mb-3', error.message);
  alert.appendChild(message);

  if (error.debug)
  {
    var pre = createElement('pre', 'alert-pre border bg-light p-2');
    alert.appendChild(pre);

    var code = createElement('code', 'text-break text-wrap',
      JSON.stringify(error.debug, null, 2));
    pre.appendChild(code);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(alert);
}

function showMeeting(meetinginfo) {
    var MeetingDiv = createElement('div', 'meetingdiv');
    var MeetingLine = createElement('div', 'divTableRow', null);

    var td = createElement('div', 'divTableCell', meetinginfo.subject);
    MeetingLine.appendChild(td);

    td = createElement('div', 'divTableCellTime', meetinginfo.startDateTime);
    MeetingLine.appendChild(td);

    td = createElement('div', 'divTableCellTime', meetinginfo.endDateTime);
    MeetingLine.appendChild(td);

    td = createElement('div', 'divTableCell', null);
    var link = createElement('a', null, 'Teams Link');
    link.title = 'Teams Link';
    link.href = meetinginfo.joinWebUrl;
    td.appendChild(link);
    MeetingLine.appendChild(td);

    td = createElement('div', 'divTableCell', null);
    link = createElement('a', null, 'Web Link');
    link.title = 'Web Link';
    link.href = meetinginfo.joinWebUrl + '&webjoin=true';
    td.appendChild(link);
    MeetingLine.appendChild(td);

    MeetingDiv.appendChild(MeetingLine);
    mainContainer.appendChild(MeetingDiv);

    console.log(MeetingDiv);
}

function updatePage(account, view, data) {
  if (!view || !account) {
    view = Views.home;
  }

  showAccountNav(account);
  showAuthenticatedNav(account, view);

  switch (view) {
    case Views.error:
        showError(data);
      break;
    case Views.home:
        showWelcomeMessage(account);
      break;
    case Views.meeting:
        showMeeting(data);
        break;
  }
}

updatePage(null, Views.home);