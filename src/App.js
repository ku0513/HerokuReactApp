import logo from './logo.svg';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";
import {useState} from "react"


function App() {

  microsoftTeams.initialize();
  // const baseUrl = `https://${window.location.hostname}:${window.location.port}`;
  const baseUrl = `https://${window.location.hostname}`;
  const [userPrincipalName, setUserPrincipalName] = useState(`userName`);
  const [userObjectId, setUserObjectId] = useState(`userId`);
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
    setUserObjectId(context.userObjectId);
    setUserPrincipalName(context.userPrincipalName);
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
          userObjectID: {userObjectId}
          <br></br>
          userPrincipalName: {userPrincipalName}
        </a>
      </header>
    </div>
  );
}

export default App;
