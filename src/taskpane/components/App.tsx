import * as React from "react";
import Header from "./Header";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { GroupedComponent } from "./GroupedComponent";
//import { FetchXMLHelper } from "../../helpers/fetchXMLParser";
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
}

// Office.onReady(function() {
//   Word.run(async function(context) {
//     // Create a proxy object for the content controls collection.
//     var contentControls = context.document.contentControls;

//     var document = context.document;

//     document.properties.load("url");

//     // Queue a command to load the id property for all of content controls.
//     context.load(contentControls, "id");

//     await context.sync();
//     //let titleOfDoc = document.properties.title;
//     let url = Office.context.document.url;
//     console.log("The url is ", url);
//     let caseId = getCaseIdFromDocTitle(url);
//     createCaseIdXmlPart(caseId);
//     insertCaseIdIntoXMLPart(caseId);
//     //console.log("XML with case ", xmlWithCase);
//   });
// });

// function getCaseIdFromDocTitle(strDocTitle) {
//   let str = strDocTitle;
//   // let txtCaseId = str.match(/(\d+)/);
//   // let caseId = parseInt(txtCaseId[0]);
//   var caseIdArr = str.toString().match(/.*\/(.+?)\./);
//   let caseId = caseIdArr[1];
//   caseId = caseId.split("-");
//   const caseIdSplit = caseId[1];
//   console.log("Split ", caseIdSplit);

//   console.log("Now ....", caseIdSplit);
//   return caseIdSplit;
// }

// function createCaseIdXmlPart(caseId) {
//   const xmlPartId = "caseId";
//   const xmlString = '<AddIn xmlns="http://schemas.pacts.com/caseId"><caseId name="' + caseId + '"> </caseId></AddIn>';

//   //Find out if the caseId XML Part exists, if it does we don't make another one.
//   Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/caseId", function(eventArgs) {
//     //If there are no XML Parts in this namespace we create it.
//     if (eventArgs.value.length === 0) {
//       Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
//         Office.context.document.settings.set(xmlPartId, asyncResult.value.id);

//         Office.context.document.settings.saveAsync(function(asyncResult) {
//           if (asyncResult.status == Office.AsyncResultStatus.Failed) {
//             console.log("Settings save failed. Error: " + asyncResult.error.message);
//           } else {
//             console.log("Saved new XML Part");
//           }
//         });
//       });
//     }
//   });
// }
// function insertCaseIdIntoXMLPart(caseId) {
//   Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/case", function(eventArgs) {
//     eventArgs.value.forEach(function(item) {
//       Office.context.document.customXmlParts.getByIdAsync(item.id, function(result) {
//         var xmlPart = result.value;
//         xmlPart.getXmlAsync(function(eventArgs) {
//           // const idInserter = new FetchXMLHelper(eventArgs.value);
//           // idInserter.insertFilterWithCaseId(caseId);

//           parseString(eventArgs.value, function(err, result) {
//             if (result) {
//               console.log("result now ...", result);
//               if (result.AddIn.fetch[0] !== null || result.AddIn.fetch[0] !== undefined) {
//                 result.AddIn.fetch[0].entity[0].filter[0].condition[0].$.value = caseId;
//                 // result.AddIn.fetch[0].entity[0].filter = [
//                 //   { condition: { $: { attribute: "incidentid", operator: "eq", value: caseId } } }
//                 // ];
//                 const xmlBuilder = new Builder();
//                 let newXml = xmlBuilder.buildObject(result);
//                 // this.setState({
//                 //   xmlWithCase: this.state.xmlWithCase.push();
//                 // })
//                 this.setState({ xmlWithCase: [...this.state.xmlWithCase, newXml] });
//                 console.log("New XML ", newXml);
//               } else {
//                 console.log("Fetch xml is null or undefined");
//               }
//             } else if (err) {
//               console.log(err);
//             }
//           });
//         });
//       });
//     });
//   });
// }

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      xmlWithCase: []
    };
    this.addCaseIdToPart = this.addCaseIdToPart.bind(this);
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
      xmlWithCase: []
    });
    this.addCaseIdToPart(currentComponent);
    console.log("Please work state   ", this.state.xmlWithCase);
  }
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
      Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/case", eventArgs => {
        eventArgs.value.forEach(item => {
          Office.context.document.customXmlParts.getByIdAsync(item.id, result => {
            var xmlPart = result.value;
            xmlPart.getXmlAsync(function(eventArgs) {
              parseString(eventArgs.value, (err, result) => {
                if (result) {
                  console.log("result now ...", result);
                  if (result.AddIn.fetch[0] !== null || result.AddIn.fetch[0] !== undefined) {
                    result.AddIn.fetch[0].entity[0].filter[0].condition[0].$.value = caseId;
                    // result.AddIn.fetch[0].entity[0].filter = [
                    //   { condition: { $: { attribute: "incidentid", operator: "eq", value: caseId } } }
                    // ];
                    const xmlBuilder = new Builder();

                    let newXml = xmlBuilder.buildObject(result);
                    console.log("New XML ", newXml);
                    // console.log("New state ", this.state.xmlWithCase);
                    currentComponent.setState({
                      xmlWithCase: [...currentComponent.state.xmlWithCase, newXml]
                    });
                    //this.setState({ xmlWithCase: [...this.state.xmlWithCase, newXml] });
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

  // createCaseIdXmlPart(caseId) {
  //   const xmlPartId = "caseId";
  //   const xmlString = '<AddIn xmlns="http://schemas.pacts.com/caseId"><caseId name="' + caseId + '"> </caseId></AddIn>';

  //   //Find out if the caseId XML Part exists, if it does we don't make another one.
  //   Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/caseId", function(eventArgs) {
  //     //If there are no XML Parts in this namespace we create it.
  //     if (eventArgs.value.length === 0) {
  //       Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
  //         Office.context.document.settings.set(xmlPartId, asyncResult.value.id);

  //         Office.context.document.settings.saveAsync(function(asyncResult) {
  //           if (asyncResult.status == Office.AsyncResultStatus.Failed) {
  //             console.log("Settings save failed. Error: " + asyncResult.error.message);
  //           } else {
  //             console.log("Saved new XML Part");
  //           }
  //         });
  //       });
  //     }
  //   });
  // }
  // this.insertCaseIdIntoXMLPart(caseId) {
  //   Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/case", function(eventArgs) {
  //     eventArgs.value.forEach(function(item) {
  //       Office.context.document.customXmlParts.getByIdAsync(item.id, function(result) {
  //         var xmlPart = result.value;
  //         xmlPart.getXmlAsync(function(eventArgs) {
  //           // const idInserter = new FetchXMLHelper(eventArgs.value);
  //           // idInserter.insertFilterWithCaseId(caseId);

  //           parseString(eventArgs.value, function(err, result) {
  //             if (result) {
  //               console.log("result now ...", result);
  //               if (result.AddIn.fetch[0] !== null || result.AddIn.fetch[0] !== undefined) {
  //                 result.AddIn.fetch[0].entity[0].filter[0].condition[0].$.value = caseId;
  //                 // result.AddIn.fetch[0].entity[0].filter = [
  //                 //   { condition: { $: { attribute: "incidentid", operator: "eq", value: caseId } } }
  //                 // ];
  //                 const xmlBuilder = new Builder();
  //                 let newXml = xmlBuilder.buildObject(result);
  //                 // this.setState({
  //                 //   xmlWithCase: this.state.xmlWithCase.push();
  //                 // })
  //                 this.setState({ xmlWithCase: [...this.state.xmlWithCase, newXml] });
  //                 console.log("New XML ", newXml);
  //               } else {
  //                 console.log("Fetch xml is null or undefined");
  //               }
  //             } else if (err) {
  //               console.log(err);
  //             }
  //           });
  //         });
  //       });
  //     });
  //   });
  // }

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
