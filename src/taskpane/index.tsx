import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { getCaseId } from "../helpers/officeHelpers";
/* global AppContainer, Component, document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Pacts Word Add In";

const render = Component => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

Office.onReady(function() {
  Word.run(function(context) {
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    console.log("Titleeee >>>>> " + getCaseId());
    // Queue a command to load the id property for all of content controls.
    context.load(contentControls, "id");

    return context.sync().then(function() {
      if (contentControls.items.length === 0) {
        console.log("No content control found.");
      } else {
        contentControls.items[0].insertHtml(
          "<strong>HTML content inserted into the content control.</strong>",
          "Start"
        );
      }
    });
  }).catch(function(error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
