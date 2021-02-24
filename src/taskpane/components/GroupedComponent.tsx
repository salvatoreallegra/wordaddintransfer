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
  color: string;
}

export interface IDetailsListGroupedExampleState {
  items: IDetailsListGroupedExampleItem[];
  groups: IGroup[];
  showItemIndexInView: boolean;
  isCompactMode: boolean;
  textBoxText: string;
  value: string;
}

const _blueGroupIndex = 2;

// let xmlDoc =
//   "<AddIn xmlns='http://schemas.pacts.com/datas/1.0'>" +
//   '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">' +
//   '<entity name="incident">' +
//   '<attribute name="title" />' +
//   '<attribute name="ticketnumber" />' +
//   '<attribute name="createdon" />' +
//   '<attribute name="incidentid" />' +
//   '<attribute name="caseorigincode" />' +
//   '<order attribute="title" descending="false" />' +
//   '<filter type="and">' +
//   '<condition attribute="statecode" operator="eq" value="0" />' +
//   "</filter>" +
//   "</entity>" +
//   '<entity name="case">' +
//   '<attribute name="caseId" />' +
//   '<attribute name="description" />' +
//   '<attribute name="createdon" />' +
//   '<attribute name="incidentid" />' +
//   '<attribute name="caseorigincode" />' +
//   '<order attribute="title" descending="false" />' +
//   '<filter type="and">' +
//   '<condition attribute="statecode" operator="eq" value="0" />' +
//   "</filter>" +
//   "</entity>" +
//   "</fetch>" +
//   "</AddIn>";

//<pacts xmlns='http://pacts/entity name here'>
let xmlDoc2 =
  "<AddIn xmlns='http://schemas.skynet.com/dataschematest/1.0'>" +
  '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">' +
  '<entity name="incident">' +
  '<attribute name="title" />' +
  '<attribute name="ticketnumber" />' +
  '<attribute name="createdon" />' +
  '<attribute name="incidentid" />' +
  '<attribute name="caseorigincode" />' +
  '<order attribute="title" descending="false" />' +
  '<filter type="and">' +
  '<condition attribute="statecode" operator="eq" value="0" />' +
  "</filter>" +
  "</entity>" +
  "</fetch>" +
  "</AddIn>";

// //must pass fetchxml string when creating object
// const fetchXMLHelper = new FetchXMLHelper(xmlDoc);
// fetchXMLHelper.parseFetchXML();
// console.log("Fetchxmlhelper   ", fetchXMLHelper);

// //we will use this items variable in our initial state below
// const items = fetchXMLHelper.getStrippedItems();
// const groups = fetchXMLHelper.getStrippedGroups();
// console.log("Items>>>>>>>> ", items);

export class GroupedComponent extends React.Component<{}, IDetailsListGroupedExampleState> {
  private _root = React.createRef<IDetailsList>();
  private _columns: IColumn[];

  constructor(props: {}) {
    super(props);

    this.state = {
      items: [],
      groups: [],
      showItemIndexInView: false,
      isCompactMode: false,
      textBoxText: "",
      value: ""
    };

    this._columns = [
      { key: "name", name: "Tables and Fields", fieldName: "name", minWidth: 100, maxWidth: 200, isResizable: true }
      //  { key: "color", name: "Color", fieldName: "color", minWidth: 100, maxWidth: 200 }
    ];

    this.handleChange = this.handleChange.bind(this);
  }

  //Get the names or something so I can operate on content controls

  //Fetch xml is embedded in the document somewhere, grab it so I can make an axios request to CDS

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

  populateGridFromXmlPart = async () => {
    return Word.run(async context => {
      const pactsXmlId = Office.context.document.settings.get("PactsXml");

      Office.context.document.customXmlParts.getByIdAsync(pactsXmlId, asyncResult => {
        asyncResult.value.getXmlAsync(asyncResult => {
          console.log("Value Based on ID  ", asyncResult.value);
          console.log("Office settings ", Office.context.document.settings);
          const fetchXMLHelper = new FetchXMLHelper(asyncResult.value);
          fetchXMLHelper.parseFetchXML();

          //we will use this items variable in our initial state below
          const items = fetchXMLHelper.getStrippedItems();
          const groups = fetchXMLHelper.getStrippedGroups();
          console.log("Items on Mount>>>>>>>> ", items);
          console.log("Groups on Mount>>>>>>>>>>", groups);
          this.setState({ items: items });
          this.setState({ groups: groups });
        });
      });
      console.log(">>>>>>>>>>>>> Jubby");
      await context.sync();
    });
  };
  public componentDidMount() {
    this.runOnMount();

    this.populateGridFromXmlPart();
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

  setGetXMLPart = () => {
    console.log("Entered xml string ", this.state.value);
    // console.log("onclick worked");
    // const xmlString =
    //   "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    // Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
    //   console.log(asyncResult.value.id);
    //   asyncResult.value.getXmlAsync(asyncResult => {
    //     console.log(asyncResult.value);
    //   });
    // });

    //Creates an xmlPart and associates id with it
    // const xmlString = xmlDoc2; //this.state.value;
    // //Office.context.document.settings.
    // Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
    //   Office.context.document.settings.set("Rogue", asyncResult.value.id);
    //   console.log("Async id", asyncResult.value.id);
    //   Office.context.document.settings.saveAsync();
    // });

    //Deletes an xml part easy version, this isn't working
    // Office.context.document.settings.remove("TestSecond");
    // Office.context.document.settings.saveAsync();

    //Deletes xml part, good version, checking to see if it works yet

    const pactsXmlId = Office.context.document.settings.get("Rogue");
    console.log(pactsXmlId);
    Office.context.document.customXmlParts.getByIdAsync(pactsXmlId, function(result) {
      var xmlPart = result.value;
      xmlPart.deleteAsync(function(eventArgs) {
        //write("The XML Part has been deleted.");
        console.log("xml part deleted");
      });
    });

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
        <TextField label="Enter FetchXML" multiline rows={3} onChange={this.handleChange} />
        <Button
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          onClick={this.setGetXMLPart}
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
