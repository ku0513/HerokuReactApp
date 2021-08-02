import logo from './logo.svg';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";

function App() {

  microsoftTeams.initialize();
  // const baseUrl = `https://${window.location.hostname}:${window.location.port}`;
  const baseUrl = `https://${window.location.hostname}`;
  let userPrincipalName;
  let userObjectId;
  console.log(baseUrl);

  microsoftTeams.settings.registerOnSaveHandler((saveEvent) => {

    microsoftTeams.settings.setSettings({
      suggestedDisplayName: "Heroku ReactApp",
      entityId: "ReactApp",
      contentUrl: baseUrl,
      websiteUrl: baseUrl,
      removeUrl: null
    });
    saveEvent.notifySuccess();
  });

  microsoftTeams.getContext((context) => {
    userPrincipalName = context.userPrincipalName;
    userObjectId = context.userObjectId;
    console.log(userPrincipalName);
    console.log(userObjectId);
  });

  microsoftTeams.settings.setValidityState(true);
  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          Edit <code>src/App.js</code> and save to reload.
        </p>
        <a
          className="App-link"
          href="https://reactjs.org"
          target="_blank"
          rel="noopener noreferrer"
        >
          Learn React with Heroku
          <br></br>
          userObjectID {userObjectId}
          <br></br>
          userPrincipalName {userPrincipalName}
        </a>
      </header>
    </div>
  );
}

export default App;
