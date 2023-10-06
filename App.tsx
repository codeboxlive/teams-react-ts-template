import { useEffect, useState, useRef } from "react";
import * as teamsJs from "@microsoft/teams-js";
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
    teamsJs.app
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
            teamsJs.app
              .getContext()
              .then((context: teamsJs.app.Context) => {
                setContextValue(JSON.stringify(context, null, 4));
              })
              .catch((error) => setContextValue(error.message));
          }}
        >
          {"Get app context"}
        </button>
        <button
          onClick={() => {
            teamsJs.pages
              .getConfig()
              .then((config: teamsJs.pages.InstanceConfig) => {
                setContextValue(JSON.stringify(config, null, 4));
              })
              .catch((error) => setContextValue(error.message));
          }}
        >
          {"Get page config"}
        </button>
        {teamsJs.pages.currentApp.isSupported() && (
          <>
            <button
              onClick={() => {
                teamsJs.pages.currentApp
                  .navigateToDefaultPage()
                  .then(() => {
                    setContextValue("navigateToDefaultPage succeeded");
                  })
                  .catch((error) => setContextValue(error.message));
              }}
            >
              {"Nav to default"}
            </button>
            <button
              onClick={() => {
                teamsJs.pages.currentApp
                  .navigateTo({
                    pageId: "test",
                    subPageId: "optional",
                  })
                  .then(() => {
                    setContextValue("navigateTo succeeded");
                  })
                  .catch((error) => setContextValue(error.message));
              }}
            >
              {"Nav to test"}
            </button>
          </>
        )}
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
