import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
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
  console.log(">>>>>Intialized");
  render(App);
};

Office.onReady(function(info) {
  console.log("<<<<<<<<< in onready");
  console.log(info);
  return Word.run(async context => {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("$none"); // Don't need any properties; just wrap each paragraph with a content control.

    await context.sync();

    for (let i = 0; i < paragraphs.items.length; i++) {
      let contentControl = paragraphs.items[i].insertContentControl();
      // For even, tag "even".
      if (i % 2 === 0) {
        contentControl.tag = "even";
      } else {
        contentControl.tag = "odd";
      }
    }
    console.log("Content controls inserted: " + paragraphs.items.length);

    await context.sync();
  });
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
