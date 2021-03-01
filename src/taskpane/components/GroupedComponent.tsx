import * as React from "react";
import axios from "axios";
import { TextField } from "office-ui-fabric-react/lib/TextField";

import { Button, ButtonType } from "office-ui-fabric-react";
import { FetchXMLHelper } from "../../helpers/fetchXMLParser";
import { MultiLineTextBox } from "./MultiLineTextBox";
import {
  DefaultButton,
  DetailsHeader,
  DetailsList,
  IColumn,
  IDetailsHeaderProps,
  IDetailsList,
  IGroup,
  IRenderFunction,
  IToggleStyles,
  mergeStyles,
  Toggle,
  IButtonStyles,
  SelectionMode
} from "office-ui-fabric-react";

const margin = "0 20px 20px 0";
const controlWrapperClass = mergeStyles({
  display: "flex",
  flexWrap: "wrap"
});
const toggleStyles: Partial<IToggleStyles> = {
  root: { margin: margin },
  label: { marginLeft: 10 }
};
const addItemButtonStyles: Partial<IButtonStyles> = { root: { margin: margin } };

export interface IDetailsListGroupedExampleItem {
  key: string;
  name: string;
  // color: string;
}

export interface IDetailsListGroupedExampleState {
  items: IDetailsListGroupedExampleItem[];
  previousItems: IDetailsListGroupedExampleItem[];
  groups: any; //IGroup[];
  showItemIndexInView: boolean;
  isCompactMode: boolean;
  textBoxText: string;
  value: string;
}

const _blueGroupIndex = 2;

//<pacts xmlns='http://pacts/entity name here'>
let xmlDoc2 =
  "<AddIn xmlns='http://schemas.pacts.com/'>" +
  '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">' +
  '<entity name="case">' +
  '<attribute name="caseId" />' +
  '<attribute name="iolation" />' +
  '<attribute name="createdon" />' +
  '<attribute name="whatever" />' +
  '<attribute name="caseworkername" />' +
  '<order attribute="title" descending="false" />' +
  '<filter type="and">' +
  '<condition attribute="statecode" operator="eq" value="0" />' +
  "</filter>" +
  "</entity>" +
  "</fetch>" +
  "</AddIn>";

let groupItemsMap = new Map();

export class GroupedComponent extends React.Component<{}, IDetailsListGroupedExampleState> {
  private _root = React.createRef<IDetailsList>();
  private _columns: IColumn[];

  constructor(props: {}) {
    super(props);

    this.state = {
      items: [],
      previousItems: [],
      groups: [],
      showItemIndexInView: false,
      isCompactMode: false,
      textBoxText: "",
      value: ""
    };

    this._columns = [
      { key: "name", name: "Tables and Fields", fieldName: "name", minWidth: 100, maxWidth: 200, isResizable: true }
    ];

    this.handleChange = this.handleChange.bind(this);
    this.populateGridFromXmlOnAdd = this.populateGridFromXmlOnAdd.bind(this);
  }

  //Next sprint fetch the xml from a tables called report datasets, will make api
  //will grab everything in the reports table  it iwll retrieve in this format one entity per in table
  // '<entity name="incident">' +
  //   '<attribute name="title" />' +
  //   '<attribute name="ticketnumber" />' +
  //   '<attribute name="createdon" />' +
  //   '<attribute name="incidentid" />' +
  //   '<attribute name="caseorigincode" />'
  //  after getting this will store each result as a xml part stored in the doc
  //  this is what will be displayed in the grid
  // when doc is reopened, all the xml parts will be displayed in the grid
  // set the name of each xml part to the name of the enity in the fetchxml

  runOnMount = async () => {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("ComponentDidMount", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };

  populateGridFromXmlPartOnMount = async () => {
    let component = this;
    return Word.run(async context => {
      //Get all the xml parts for the namespace in the doc, then populate grid
      Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/case", function(eventArgs) {
        console.log("Found " + eventArgs.value.length + " parts with this namespace");
        console.log("Event args", eventArgs);
        console.log("Event Args value ", eventArgs.value);
        eventArgs.value.forEach(function(id) {
          const pactsXmlId = Office.context.document.settings.get("case");
          console.log("Checking id ", pactsXmlId);
          console.log("Id of xml part ", id.id);
          // const pactsXmlId = Office.context.document.settings.get("case");

          Office.context.document.customXmlParts.getByIdAsync(id.id, asyncResult => {
            asyncResult.value.getXmlAsync(asyncResult => {
              console.log("Value Based on ID  ", asyncResult.value);
              console.log("Office settings ", Office.context.document.settings);
              const fetchXMLHelper = new FetchXMLHelper(asyncResult.value);
              fetchXMLHelper.parseFetchXML();

              //we will use this items variable in our initial state below
              const items = fetchXMLHelper.getStrippedItems();
              const groups = fetchXMLHelper.getStrippedGroups();

              // FetchXMLHelper.xmlPartIds.push(groups);
              // groupItemsMap.set(groups[0].name, items);
              // console.log("Map ", groupItemsMap);
              // console.log("One item per group ", groupItemsMap.get("case"));

              console.log("Items on Mount>>>>>>>> ", items);
              console.log("Groups on Mount>>>>>>>>>>", groups);

              // component.setState({ items: items });
              // component.setState({ groups: groups });
              component.setState({ items: [...component.state.items, ...items] });
              component.setState({ groups: [...component.state.groups, ...groups] });
            });
          });
        });
      });

      // const pactsXmlId = Office.context.document.settings.get("case");

      // Office.context.document.customXmlParts.getByIdAsync(pactsXmlId, asyncResult => {
      //   asyncResult.value.getXmlAsync(asyncResult => {
      //     console.log("Value Based on ID  ", asyncResult.value);
      //     console.log("Office settings ", Office.context.document.settings);
      //     const fetchXMLHelper = new FetchXMLHelper(asyncResult.value);
      //     fetchXMLHelper.parseFetchXML();

      //     //we will use this items variable in our initial state below
      //     const items = fetchXMLHelper.getStrippedItems();
      //     const groups = fetchXMLHelper.getStrippedGroups();

      //     FetchXMLHelper.xmlPartIds.push(groups);
      //     groupItemsMap.set(groups[0].name, items);
      //     console.log("Map ", groupItemsMap);
      //     console.log("One item per group ", groupItemsMap.get("case"));

      //     console.log("Items on Mount>>>>>>>> ", items);
      //     console.log("Groups on Mount>>>>>>>>>>", groups);

      //     this.setState({ items: items });
      //     this.setState({ groups: groups });
      //     this.showXMLPartsInNamespace();
      //   });
      // });
      await context.sync();
    });
  };

  getAllXmlParts() {}

  populateGridFromXmlOnAdd = async xmlPartId => {
    return Word.run(async context => {
      console.log("From pop grid on click ...", xmlPartId);

      const pactsXmlId = Office.context.document.settings.get(xmlPartId);
      Office.context.document.customXmlParts.getByIdAsync(pactsXmlId, asyncResult => {
        asyncResult.value.getXmlAsync(asyncResult => {
          const fetchXMLHelper = new FetchXMLHelper(asyncResult.value);
          fetchXMLHelper.parseFetchXML();

          //we will use this items variable in our initial state below
          const items = fetchXMLHelper.getStrippedItems();
          const groups = fetchXMLHelper.getStrippedGroups();

          if (FetchXMLHelper.xmlPartIds.includes(groups)) {
          } else {
            FetchXMLHelper.xmlPartIds.push(groups);
          }

          console.log("Items on Add Click >>>>>>>> ", items);
          console.log("Groups on Add Click >>>>>>>>>>", groups);

          console.log("State groups ", this.state.groups);
          console.log("State items ", this.state.items);
          this.setState({ items: [...this.state.items, ...items] });
          this.setState({ groups: [...this.state.groups, ...groups] });
          console.log("Current state of items ", this.state.items);
          console.log("Current state of groups ", this.state.groups);
        });
      });
      await context.sync();
    });
  };

  public componentDidMount() {
    this.runOnMount();
    this.populateGridFromXmlPartOnMount();
  }

  public componentWillUnmount() {
    if (this.state.showItemIndexInView) {
      const itemIndexInView = this._root.current!.getStartItemIndexInView();
      alert("first item index that was in view: " + itemIndexInView);
    }
  }

  runOnChange = (item: IDetailsListGroupedExampleItem) => {
    return Word.run(async context => {
      var serviceNameRange = context.document.getSelection();
      var serviceNameContentControl = serviceNameRange.insertContentControl();

      // serviceNameContentControl.set({
      //   color: "red",
      //   title: "Odd ContentControl #" + (i + 1),
      //   appearance: "Tags"
      // });

      //serviceNameContentControl.subtype; //gets content control type
      serviceNameContentControl.title = "Service Name";
      serviceNameContentControl.title = item.name;
      serviceNameContentControl.tag = "serviceName";
      serviceNameContentControl.appearance = "Tags";
      serviceNameContentControl.color = "blue";
      await context.sync();
    });
  };

  // updateContentControls = () => {
  //   return Word.run(async context => {
  //     var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
  //     serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
  //     await context.sync();
  //   });
  // };

  add = () => {
    return Word.run(async context => {
      var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
      serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
      await context.sync();
    });
  };

  setGetXMLPart = functionAsParam => {
    //we probably need to validate the xml entered into the multiline textbox
    let enteredXmlString = this.state.value;
    console.log("Entered xml string ", enteredXmlString);

    //Parse the table name out of the xml, this will be the Key or Id for the xml part saved in the doc
    const fetchXMLHelperTextBox = new FetchXMLHelper(enteredXmlString);
    fetchXMLHelperTextBox.parseFetchXML();
    let strippedGroups = fetchXMLHelperTextBox.getStrippedGroups();
    let xmlPartId = strippedGroups[0].name;
    console.log("xml part id ", xmlPartId);

    const pactsXmlId = Office.context.document.settings.get(xmlPartId);
    console.log("Pacts xml id ", pactsXmlId); //if null

    if (pactsXmlId === null) {
      //create new xml part with xmlpartid as key
      const xmlString = enteredXmlString; //this.state.value;

      //Office.context.document.settings.
      Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
        console.log("New XML Part Created");
        Office.context.document.settings.set(xmlPartId, asyncResult.value.id);
        console.log("Async id", asyncResult.value.id);
        Office.context.document.settings.saveAsync(function() {
          if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.log("Settings save failed. Error: " + asyncResult.error.message);
          } else {
            console.log("Settings saved.");
            functionAsParam(xmlPartId);
          }
        });
      });
    } else {
      //delete the existing xml part with the same key name

      const pactsXmlId = Office.context.document.settings.get(xmlPartId);
      console.log(pactsXmlId);
      Office.context.document.customXmlParts.getByIdAsync(pactsXmlId, function(result) {
        //create new xml part with table name as key

        var xmlPart = result.value;
        console.log("XML Part", xmlPart);

        xmlPart.deleteAsync(function(eventArgs) {
          //write("The XML Part has been deleted.");
          console.log("xml part deleted");
          const xmlString = enteredXmlString; //this.state.value;
          console.log(xmlString);
          //Office.context.document.settings.
          Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
            Office.context.document.settings.set(xmlPartId, asyncResult.value.id);
            console.log("Async id When ", asyncResult.value.id);

            Office.context.document.settings.saveAsync(function(asyncResult) {
              if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                console.log("Settings save failed. Error: " + asyncResult.error.message);
              } else {
                console.log("Saved new XML Part");
                functionAsParam(xmlPartId);
                //this.populateGridFromXmlOnAdd(xmlPartId);
              }
            });
          });
        });
      });
    }
    //this.populateGridFromXmlOnAdd(xmlPartId);
    // Office.context.document.customXmlParts.getByIdAsync(pactsXmlId, asyncResult => {
    //   asyncResult.value.getXmlAsync(asyncResult => {
    //     console.log("Value Based on ID Check This Now ", asyncResult.value);
    //     console.log("Office settings ", Office.context.document.settings);
    //   });
    // });

    //Deletes xml part, good version

    // const pactsXmlId = Office.context.document.settings.get("Rogue");
    // console.log(pactsXmlId);
    // Office.context.document.customXmlParts.getByIdAsync(pactsXmlId, function(result) {
    //   var xmlPart = result.value;
    //   xmlPart.deleteAsync(function(eventArgs) {
    //     //write("The XML Part has been deleted.");
    //     console.log("xml part deleted");
    //   });
    // });

    //Creates an xmlPart and associates id with it

    // const xmlString = xmlDoc2; //this.state.value;
    // console.log(xmlString);
    // //Office.context.document.settings.
    // Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
    //   Office.context.document.settings.set("Dungeon", asyncResult.value.id);
    //   console.log("Async id", asyncResult.value.id);
    //   Office.context.document.settings.saveAsync();
    // });

    this.setState({ value: "" });

    // const pactsXmlId = Office.context.document.settings.get("PactsXml");
    // Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId, asyncResult => {
    //   asyncResult.value.getXmlAsync(asyncResult => {
    //     console.log("Value Based on ID  ", asyncResult.value);
    //     console.log("Office settings ", Office.context.document.settings);
    //   });
    // });
  };

  handleChange(event) {
    this.setState({ value: event.target.value });
  }

  public render() {
    const { items, groups, isCompactMode } = this.state;

    return (
      <div>
        {/* <div className={controlWrapperClass}>
          <DefaultButton onClick={this._addItem} text="Add an item" styles={addItemButtonStyles} />
          <Toggle
            label="Compact mode"
            inlineLabel
            checked={isCompactMode}
            onChange={this._onChangeCompactMode}
            styles={toggleStyles}
          />
          <Toggle
            label="Show index of first item in view when unmounting"
            inlineLabel
            checked={this.state.showItemIndexInView}
            onChange={this._onShowItemIndexInViewChanged}
            styles={toggleStyles}
          />
        </div> */}
        <DetailsList
          componentRef={this._root}
          items={items}
          groups={groups}
          columns={this._columns}
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          ariaLabelForSelectionColumn="Toggle selection"
          checkButtonAriaLabel="Row checkbox"
          onRenderDetailsHeader={this._onRenderDetailsHeader}
          //this is basically the onClick
          //see what arrows do and may need to disable based on ev(event)
          //next step is get the content control
          onActiveItemChanged={this.runOnChange}
          groupProps={{
            showEmptyGroups: true
          }}
          selectionMode={SelectionMode.single}
          onRenderItemColumn={this._onRenderColumn}
          compact={isCompactMode}
        />
        {/* <Button
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          onClick={this.updateContentControls}
        >
          Update Content Controls
        </Button> */}

        <MultiLineTextBox />
        {/* Might need to use the value field to clear Textfield, look at react form docs https://reactjs.org/docs/forms.html */}
        <TextField label="Enter FetchXML" multiline rows={3} value={this.state.value} onChange={this.handleChange} />
        <Button
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          onClick={() => this.setGetXMLPart(this.populateGridFromXmlOnAdd)}
        >
          Add
        </Button>
      </div>
    );
  }

  // private _addItem = (): void => {
  //   const items = this.state.items;
  //   const groups = [...this.state.groups];
  //   groups[_blueGroupIndex].count++;

  //   this.setState(
  //     {
  //       items: items.concat([
  //         {
  //           key: "item-" + items.length,
  //           name: "New item " + items.length,
  //           color: "blue"
  //         }
  //       ]),
  //       groups
  //     },
  //     () => {
  //       if (this._root.current) {
  //         this._root.current.focusIndex(items.length, true);
  //       }
  //     }
  //   );
  // };

  private _onRenderDetailsHeader(props: IDetailsHeaderProps, _defaultRender?: IRenderFunction<IDetailsHeaderProps>) {
    return <DetailsHeader {...props} ariaLabelForToggleAllGroupsButton={"Expand collapse groups"} />;
  }

  private _onRenderColumn(item: IDetailsListGroupedExampleItem, _index: number, column: IColumn) {
    const value =
      item && column && column.fieldName ? item[column.fieldName as keyof IDetailsListGroupedExampleItem] || "" : "";

    return <div data-is-focusable={true}>{value}</div>;
  }

  // private _onShowItemIndexInViewChanged = (_event: React.MouseEvent<HTMLInputElement>, checked: boolean): void => {
  //   this.setState({ showItemIndexInView: checked });
  // };

  // private _onChangeCompactMode = (_ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
  //   this.setState({ isCompactMode: checked });
  // };
}
