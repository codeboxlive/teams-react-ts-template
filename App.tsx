import { useEffect, useState, useRef } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { CODEBOX_LIVE_ORIGINS } from "@codeboxlive/extensions-core";
// Create new components and import them like this
import Header from "./Header";

export default function App() {
  const initRef = useRef<boolean>(false);
  const [contextValue, setContextValue] = useState<string>("loading");

  useEffect(() => {
    if (initRef.current) {
      return;
    }
    initRef.current = true;
    microsoftTeams.app
      .initialize(CODEBOX_LIVE_ORIGINS)
      .then(() => {
        microsoftTeams.app
          .getContext()
          .then((context: microsoftTeams.app.Context) => {
            setContextValue(JSON.stringify(context, null, 2));
          })
          .catch((error) => setContextValue(error.message));
      })
      .catch((error) => setContextValue(error.message));
  });
  return (
    <>
      <Header />
      <h3>{"Teams app context:"}</h3>
      <p>{contextValue}</p>
    </>
  );
}
