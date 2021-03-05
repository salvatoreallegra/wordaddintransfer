import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
//import { getCaseId } from "../helpers/officeHelpers";
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
  Word.run(async function(context) {
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;

    var document = context.document;

    document.properties.load("title");

    // Queue a command to load the id property for all of content controls.
    context.load(contentControls, "id");
    // let extract = new CaseIdExtractor();
    // extract.getCaseId();

    await context.sync();
    let titleOfDoc = document.properties.title;
    getCaseId(titleOfDoc);
    createCaseIdXmlPart();
  });
});

function getCaseId(strDocTitle) {
  let str = strDocTitle;
  let txtCaseId = str.match(/(\d+)/);
  let caseId = parseInt(txtCaseId[0]);
  console.log("STR Doc ", caseId);
}

function createCaseIdXmlPart() {
  const xmlPartId = "caseId";
  const xmlString = "<AddIn xmlns='http://schemas.pacts.com/caseId'><caseId name= caseId> </caseId></AddIn>";

  Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
    Office.context.document.settings.set(xmlPartId, asyncResult.value.id);

    Office.context.document.settings.saveAsync(function(asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log("Settings save failed. Error: " + asyncResult.error.message);
      } else {
        console.log("Saved new XML Part");
      }
    });
  });
}

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
