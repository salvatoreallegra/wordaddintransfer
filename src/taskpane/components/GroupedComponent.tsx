import * as React from "react";
import axios from "axios";
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
}

const _blueGroupIndex = 2;

let xmlDoc =
  "<AddIn xmlns='http://schemas.contoso.com/review/1.0'>" +
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
  '<entity name="case">' +
  '<attribute name="caseId" />' +
  '<attribute name="description" />' +
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

//must pass fetchxml string when creating object
const fetchXMLHelper = new FetchXMLHelper(xmlDoc);
fetchXMLHelper.parseFetchXML();
console.log("Fetchxmlhelper   ", fetchXMLHelper);

//we will use this items variable in our initial state below
const items = fetchXMLHelper.getStrippedItems();
const groups = fetchXMLHelper.getStrippedGroups();
console.log("Items>>>>>>>> ", items);

export class GroupedComponent extends React.Component<{}, IDetailsListGroupedExampleState> {
  private _root = React.createRef<IDetailsList>();
  private _columns: IColumn[];

  constructor(props: {}) {
    super(props);

    this.state = {
      items: items,
      //[
      //   { key: "a", name: "ContactId", color: "red" },
      //   { key: "b", name: "b", color: "red" },
      //   { key: "x", name: "xyz", color: "gold" },
      //   { key: "c", name: "c", color: "blue" },
      //   { key: "d", name: "d", color: "blue" },
      //   { key: "e", name: "e", color: "blue" }
      //]
      // This is based on the definition of items
      // groups: [
      //   { key: "groupred0", name: 'Offender: "Blue"', startIndex: 0, count: 2, level: 0 },
      //   { key: "groupgreen2", name: 'Crimes: "green"', startIndex: 2, count: 0, level: 0 },
      //   { key: "groupblue2", name: 'Drug Use: "blue"', startIndex: 2, count: 3, level: 0 }
      // ]

      groups: groups,
      showItemIndexInView: false,
      isCompactMode: false,
      textBoxText: ""
    };

    this._columns = [
      { key: "name", name: "Tables and Fields", fieldName: "name", minWidth: 100, maxWidth: 200, isResizable: true }
      //  { key: "color", name: "Color", fieldName: "color", minWidth: 100, maxWidth: 200 }
    ];
  }

  //Get the names or something so I can operate on content controls

  //Fetch xml is embedded in the document somewhere, grab it so I can make an axios request to CDS

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
  public componentDidMount() {
    this.runOnMount();
    axios.get(`https://jsonplaceholder.typicode.com/users`).then(res => {
      const persons = res.data;
      // this.setState({ apiData: persons });
      console.log(persons);
      // console.log("Data from state ", this.state.apiData);
    });
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
    // console.log("onclick worked");
    // const xmlString =
    //   "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    // Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
    //   console.log(asyncResult.value.id);
    //   asyncResult.value.getXmlAsync(asyncResult => {
    //     console.log(asyncResult.value);
    //   });
    // });
    //debugger;
    // const xmlString =
    //   "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    // Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
    //   Office.context.document.settings.set("ReviewersID", asyncResult.value.id);
    //   console.log("Async id", asyncResult.value.id);
    //   Office.context.document.settings.saveAsync();
    // });
    const reviewersXmlId = Office.context.document.settings.get("ReviewersID");
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId, asyncResult => {
      asyncResult.value.getXmlAsync(asyncResult => {
        console.log("Value Based on ID  ", asyncResult.value);
        console.log("Office settings ", Office.context.document.settings);
      });
    });
  };

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
