import * as React from "react";
import Header from "./Header";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { GroupedComponent } from "./GroupedComponent";
import { FetchXMLHelper } from "../../helpers/fetchXMLParser";
import { parseString, Builder } from "xml2js";
// import { getFocusOutlineStyle } from "@uifabric/styling";
// import { getScrollbarWidth, shallowCompare } from "@uifabric/utilities";

/* global Button Header, HeroList, HeroListItem, Progress, Word */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  xmlWithCase: [];
  xmlPartResponse: any;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      xmlWithCase: [],
      xmlPartResponse: {}
    };
    this.addCaseIdToPart = this.addCaseIdToPart.bind(this);
    this.mockResponseToState = this.mockResponseToState.bind(this);
  }

  componentDidMount() {
    let currentComponent = this;
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
      ],
      xmlWithCase: [],
      xmlPartResponse: {}
    });

    this.addCaseIdToPart(currentComponent);
    this.mockResponseToState(currentComponent);
  }

  printToken() {
    Word.run(async function(context) {
      // Create a proxy object for the content controls collection.

      await context.sync();

      //console.log("XML with case ", xmlWithCase);
    });
  }

  componentDidUpdate() {
    // Typical usage (don't forget to compare props):
    //if (this.props.userID !== prevProps.userID) {
    console.log("Lifecycle update ", this.state.xmlPartResponse);
    //}
  }
  //This is a general look of what a fetchxml query might return
  //We will store the results in state for use in the content controls
  mockResponseToState(currentComponent) {
    // Office.getAccessToken(function(result) {
    //   if (result.status === "succeeded") {
    //     var token = result.value;
    //     console.log(token);
    //     // ...
    //   } else {
    //     console.log("Error obtaining token", result.error);
    //   }
    // });

    // Office.auth.getAccessToken(function(result) {
    //   if (result.status === "succeeded") {
    //     var token = result.value;
    //     console.log(token);
    //     // ...
    //   } else {
    //     console.log("Error obtaining token", result.error);
    //   }
    // });
    //JOHN Pseudocode
    // foreach(content in ContentControl){
    //   const controlEntityName = //gets entity name from control
    //   const controlFieldName = //gets field name from control.
    //   const dataset = //this is the dataset in state that has all the responses of data

    //   dataset[controlEntityName] [0][controlFieldName] //grabs the first value for the given entity and field
    // }

    // so the dataset should be stored like

    // const dataset = {
    //   entityName1: [], //array of values returned from response, for each value, we want to access the field value ideally like value.fieldName
    //   entityName2: [valuesHere],
    //   entityName3: [valuues],
    //   etc...
    // }

    Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/", eventArgs => {
      eventArgs.value.forEach(item => {
        Office.context.document.customXmlParts.getByIdAsync(item.id, result => {
          var xmlPart = result.value;
          xmlPart.getXmlAsync(function(eventArgs) {
            const xmlHelper = new FetchXMLHelper(eventArgs.value);
            xmlHelper.parseFetchXMLNoParams();
            const groupsArray = xmlHelper.getStrippedGroups();
            const tableName = groupsArray[0].name;
            const valueResponseOne = [
              {
                pacts_name: "Active Duty - Airforce Jake Skellington",
                pacts_retired: false
              },
              {
                pacts_name: "Active Duty - Jake Skellington",
                pacts_retired: true
              }
            ];

            //this represents one http responsse for one xml part
            // const valueResponseTwo = [
            //   {
            //     pacts_name: "Active Duty - Airforce Jake Skellington",
            //     pacts_dateentered: "2020-12-01T00:00:00Z"
            //   },
            //   {
            //     pacts_name: "Active Duty - Jake Skellington",
            //     pacts_dateentered: "2020-12-01T00:00:00Z"
            //   }
            // ];

            // const dataset = {
            //   entityName1: [], //array of values returned from response, for each value, we want to access the field value ideally like value.fieldName
            //   entityName2: [valuesHere],
            //   entityName3: [valuues],
            //   etc...
            // }
            const dataset = {};
            dataset[tableName] = [valueResponseOne];

            console.log("data set, ", dataset);

            currentComponent.setState({
              xmlPartResponse: [...currentComponent.state.xmlPartResponse, dataset]
            });
          });
        });
      });
    });
  }

  //iterate throgh controls and grab the control info so i can build the state above
  grabContentControlData = async () => {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */
      // Create a proxy object for the content controls collection.
      var contentControls = context.document.contentControls;

      // Queue a command to load the id property for all of content controls.
      context.load(contentControls, "title, id");
      if (contentControls.items.length === 0) {
        console.log("No content control found.");
      } else {
        contentControls.items[0].insertHtml(
          "<strong>HTML content inserted into the content control.</strong><table><tr><td>Hello</td><td>World</td></tr></table>",
          "Start"
        );

        for (let i = 0; i < contentControls.items.length; i++) {
          console.log("Content Control Titles " + contentControls.items[i].title);
        }
      }

      await context.sync();
    });
  };

  addCaseIdToPart(currentComponent) {
    Office.onReady(function() {
      Word.run(async function(context) {
        // Create a proxy object for the content controls collection.
        var contentControls = context.document.contentControls;

        var document = context.document;

        document.properties.load("url");

        // Queue a command to load the id property for all of content controls.
        context.load(contentControls, "id");

        await context.sync();
        //let titleOfDoc = document.properties.title;
        let url = Office.context.document.url;
        console.log("The url is ", url);
        let caseId = getCaseIdFromDocTitle(url);
        createCaseIdXmlPart(caseId);
        insertCaseIdIntoXMLPart(caseId);
        //console.log("XML with case ", xmlWithCase);
      });
    });
    function getCaseIdFromDocTitle(strDocTitle) {
      let str = strDocTitle;
      // let txtCaseId = str.match(/(\d+)/);
      // let caseId = parseInt(txtCaseId[0]);
      var caseIdArr = str.toString().match(/.*\/(.+?)\./);
      let caseId = caseIdArr[1];
      caseId = caseId.split("-");
      const caseIdSplit = caseId[1];
      console.log("Split ", caseIdSplit);

      console.log("Now ....", caseIdSplit);
      return caseIdSplit;
    }

    function createCaseIdXmlPart(caseId) {
      const xmlPartId = "caseId";
      const xmlString =
        '<AddIn xmlns="http://schemas.pacts.com/caseId"><caseId name="' + caseId + '"> </caseId></AddIn>';

      //Find out if the caseId XML Part exists, if it does we don't make another one.
      Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/caseId", function(
        eventArgs
      ) {
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
      Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/", eventArgs => {
        eventArgs.value.forEach(item => {
          Office.context.document.customXmlParts.getByIdAsync(item.id, result => {
            var xmlPart = result.value;
            xmlPart.getXmlAsync(function(eventArgs) {
              parseString(eventArgs.value, (err, result) => {
                if (result) {
                  console.log("result now ...", result);
                  for (const [key, value] of Object.entries(result)) {
                    console.log("Keys and values ", `${key}: ${value}`);
                  }
                  if (result.AddIn.fetch[0] !== null || result.AddIn.fetch[0] !== undefined) {
                    result.AddIn.fetch[0].entity[0].filter[0].condition[0].$.value = caseId;

                    const xmlBuilder = new Builder({ headless: true });

                    let newXml = xmlBuilder.buildObject(result);
                    console.log("New XML ", newXml);
                    currentComponent.setState({
                      xmlWithCase: [...currentComponent.state.xmlWithCase, newXml]
                    });

                    console.log("Inside async ....", currentComponent.state.xmlWithCase);
                  } else {
                    console.log("Fetch xml is null or undefined");
                  }
                } else if (err) {
                  console.log(err);
                }
              });
            });
          });
        });
      });
    }
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

        <GroupedComponent />
      </div>
    );
  }
}
