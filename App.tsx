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
        setContextValue("SDK initialized");
      })
      .catch((error) => setContextValue(error.message));
  }, []);

  return (
    <>
      <Header />
      <div>
        <button
          onClick={() => {
            microsoftTeams.app
              .getContext()
              .then((context: microsoftTeams.app.Context) => {
                setContextValue(JSON.stringify(context, null, 4));
              })
              .catch((error) => setContextValue(error.message));
          }}
        >
          {"Get app context"}
        </button>
        <button
          onClick={() => {
            microsoftTeams.pages
              .getConfig()
              .then((config: microsoftTeams.pages.InstanceConfig) => {
                setContextValue(JSON.stringify(config, null, 4));
              })
              .catch((error) => setContextValue(error.message));
          }}
        >
          {"Get page config"}
        </button>
      </div>
      <h3>{"Response:"}</h3>
      <div
        style={{
          whiteSpace: "pre",
          padding: "8px",
          backgroundColor: "black",
          color: "white",
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "520px",
          lineHeight: "200%",
        }}
      >
        {contextValue}
      </div>
    </>
  );
}
