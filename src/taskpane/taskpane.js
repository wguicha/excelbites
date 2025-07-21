/* global document, Office, module */

import App from "./components/App";
import { initializeIcons } from "@fluentui/react";
import React from "react";
import ReactDOM from "react-dom/client";
import "./i18n";

initializeIcons();

let isOfficeInitialized = false;

const root = ReactDOM.createRoot(document.getElementById("container"));

const render = (Component) => {
  root.render(<Component />);
};

/* Initial render when the page is loaded */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});
