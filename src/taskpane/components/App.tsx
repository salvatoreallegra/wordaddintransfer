import * as React from "react";
import Header from "./Header";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { GroupedComponent } from "./GroupedComponent";
import { FetchXMLHelper } from "../../helpers/fetchXMLParser";

/* global Button Header, HeroList, HeroListItem, Progress, Word */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

Office.onReady(function() {
  Word.run(async function(context) {
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;

    var document = context.document;

    document.properties.load("title");

    // Queue a command to load the id property for all of content controls.
    context.load(contentControls, "id");

    await context.sync();
    let titleOfDoc = document.properties.title;
    let caseId = getCaseIdFromDocTitle(titleOfDoc);
    createCaseIdXmlPart(caseId);
    insertCaseIdIntoXMLPart(caseId);
  });
});

function getCaseIdFromDocTitle(strDocTitle) {
  let str = strDocTitle;
  let txtCaseId = str.match(/(\d+)/);
  let caseId = parseInt(txtCaseId[0]);

  return caseId;
}

function createCaseIdXmlPart(caseId) {
  const xmlPartId = "caseId";
  const xmlString = '<AddIn xmlns="http://schemas.pacts.com/caseId"><caseId name="' + caseId + '"> </caseId></AddIn>';

  //Find out if the caseId XML Part exists, if it does we don't make another one.
  Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/caseId", function(eventArgs) {
    //If there are no XML Parts in this namespace we create it.
    if (eventArgs.value.length === 0) {
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
  });
}
function insertCaseIdIntoXMLPart(caseId) {
  Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/case", function(eventArgs) {
    eventArgs.value.forEach(function(item) {
      Office.context.document.customXmlParts.getByIdAsync(item.id, function(result) {
        var xmlPart = result.value;
        xmlPart.getXmlAsync(function(eventArgs) {
          const idInserter = new FetchXMLHelper(eventArgs.value);
          idInserter.insertFilterWithCaseId(caseId);
        });
      });
    });
  });
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
  }

  click = async () => {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="PACTS Word Add-In" />
        {/* <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}> */}
        {/* <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p> */}
        <GroupedComponent />
        {/* <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button> */}

        {/* </HeroList> */}
      </div>
    );
  }
}
