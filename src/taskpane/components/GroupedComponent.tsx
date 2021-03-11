import * as React from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import axios from "axios";
import { Button, ButtonType } from "office-ui-fabric-react";
import { FetchXMLHelper } from "../../helpers/fetchXMLParser";
import {
  DetailsHeader,
  DetailsList,
  IColumn,
  IDetailsHeaderProps,
  IDetailsList,
  IRenderFunction,
  SelectionMode
} from "office-ui-fabric-react";

//John: export { tableFields };

export interface IDetailsListGroupedExampleItem {
  key: string;
  name: string;
}

export interface IFieldTemplate {
  groups: any;
  items: IDetailsListGroupedExampleItem[];
}
export interface IDetailsListGroupedExampleState {
  tableFields: IFieldTemplate;
  items: IDetailsListGroupedExampleItem[];
  previousItems: IDetailsListGroupedExampleItem[];
  groups: any; //IGroup[];
  showItemIndexInView: boolean;
  isCompactMode: boolean;
  textBoxText: string;
  value: string;
}

export class GroupedComponent extends React.Component<{}, IDetailsListGroupedExampleState> {
  private _root = React.createRef<IDetailsList>();
  private _columns: IColumn[];

  constructor(props: {}) {
    super(props);

    this.state = {
      tableFields: null,
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
    //this.populateGridFromXmlOnAdd = this.populateGridFromXmlOnAdd.bind(this); John
  }

  runOnMount = async () => {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */
      let config = {
        headers: {
          "Content-Type": "application-json"
        }
      };
      axios
        .post(
          "https://prod-08.usgovvirginia.logic.azure.us:443/workflows/8e0e7d7919b2426ca76b41385cb1d4f4/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BAqRIP9ud3MDh92IzSlskmMSK7hEkn2rWUxSpSy7Sp8",
          {
            Dataset: "Finn"
          },
          config
        )
        .then(
          response => {
            console.log("***********************", response);
          },
          error => {
            console.log("****************************", error);
          }
        );

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("ComponentDidMount", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };
  grabContentControl = () => {
    Word.run(function(context) {
      // Create a proxy object for the content controls collection.
      var contentControls = context.document.contentControls;

      // Queue a command to load the id property for all of content controls.
      context.load(contentControls, "title, id");

      return context.sync().then(function() {
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
      });
    }).catch(function(error) {
      console.log("Error: " + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  };

  setDisplay = (asyncResult, component, contentXmlPart, contentXmlParts) => {
    const fetchXMLHelper = new FetchXMLHelper(asyncResult.value);
    fetchXMLHelper.parseFetchXML(component.state.tableFields); //John

    //we will use this items variable in our initial state below
    const items = fetchXMLHelper.getStrippedItems();
    const groups = fetchXMLHelper.getStrippedGroups();

    component.setState({
      //come back to properly changing state
      tableFields: {
        groups,
        items
      }
    });

    //John :
    // tableFields.push({
    //   name: groups,
    //   fields: items
    // });

    component.setState({ items: [...component.state.items, ...items] }); //John: do we need this? It was commented out.
    //for each group, check my corresponding items from the other array and get the count
    //set the startIndex to the count of my corresponding item array
    component.setState({ groups: [...component.state.groups, ...groups] });

    contentXmlPart = contentXmlParts && contentXmlParts.shift();
    if (contentXmlPart) {
      contentXmlPart.getXmlAsync(asyncResult => {
        component.setDisplay(asyncResult, component, contentXmlPart, contentXmlParts);
      });
    }
  };

  populateGridFromXmlPartOnMount = async () => {
    let component = this;
    return Word.run(async context => {
      //Get all the xml parts for the namespace in the doc, then populate grid

      Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.pacts.com/", function(eventArgs) {
        console.log("Found " + eventArgs.value.length + " parts with this namespace");
        console.log("Event args", eventArgs);
        console.log("Event Args value ", eventArgs.value);

        //eventArgs.value.forEach(function(contentXmlPart) {John
        const contentXmlPart = eventArgs.value && eventArgs.value.shift();
        const pactsXmlId = Office.context.document.settings.get("case");
        console.log("Checking id ", pactsXmlId);
        //John: await Office.context.document.customXmlParts.getByIdAsync(id.id, asyncResult => {

        contentXmlPart.getXmlAsync(asyncResult => {
          component.setDisplay(asyncResult, component, contentXmlPart, eventArgs.value);
        });
      });

      await context.sync();
    });
  };

  populateGridFromXmlOnAdd = async xmlPartId => {
    return Word.run(async context => {
      const pactsXmlId = Office.context.document.settings.get(xmlPartId);
      Office.context.document.customXmlParts.getByIdAsync(pactsXmlId, asyncResult => {
        asyncResult.value.getXmlAsync(asyncResult => {
          const fetchXMLHelper = new FetchXMLHelper(asyncResult.value);
          fetchXMLHelper.parseFetchXML(this.state.tableFields);

          //we will use this items variable in our initial state below
          const items = fetchXMLHelper.getStrippedItems();
          const groups = fetchXMLHelper.getStrippedGroups();
          this.setState({
            //come back to properly changing state
            tableFields: {
              groups,
              items
            }
          });

          this.setState({ items: [...this.state.items, ...items] });
          this.setState({ groups: [...this.state.groups, ...groups] });
        });
      });
      await context.sync();
    });
  };

  public componentDidMount() {
    this.runOnMount();
    this.populateGridFromXmlPartOnMount();
    this.grabContentControl();
  }

  public componentWillUnmount() {
    if (this.state.showItemIndexInView) {
      const itemIndexInView = this._root.current!.getStartItemIndexInView();
      console.log(itemIndexInView);
    }
  }

  runOnChange = (item: IDetailsListGroupedExampleItem) => {
    return Word.run(async context => {
      var serviceNameRange = context.document.getSelection();
      var serviceNameContentControl = serviceNameRange.insertContentControl();

      serviceNameContentControl.title = "Service Name";
      serviceNameContentControl.title = item.name;
      serviceNameContentControl.tag = "serviceName";
      serviceNameContentControl.appearance = "Tags";
      serviceNameContentControl.color = "blue";
      await context.sync();
    });
  };

  add = () => {
    return Word.run(async context => {
      var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
      serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
      await context.sync();
    });
  };

  //If there is no xml part in the doc we will get an undefined error
  setGetXMLPart = functionAsParam => {
    //we probably need to validate the xml entered into the multiline textbox
    let enteredXmlString = this.state.value;

    //Parse the table name out of the xml, this will be the Key or Id for the xml part saved in the doc
    const fetchXMLHelperTextBox = new FetchXMLHelper(enteredXmlString);
    fetchXMLHelperTextBox.parseFetchXML(this.state.tableFields);
    let strippedGroups = fetchXMLHelperTextBox.getStrippedGroups();
    let xmlPartId = strippedGroups[0].name;

    const pactsXmlId = Office.context.document.settings.get(xmlPartId);

    if (pactsXmlId === null) {
      //create new xml part with xmlpartid as key
      const xmlString = enteredXmlString; //this.state.value;

      //Office.context.document.settings.
      Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
        Office.context.document.settings.set(xmlPartId, asyncResult.value.id);

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

        xmlPart.deleteAsync(function() {
          //write("The XML Part has been deleted.");

          const xmlString = enteredXmlString; //this.state.value;

          //After deleting the existing custom xml Part, we now re-create it
          Office.context.document.customXmlParts.addAsync(xmlString, asyncResult => {
            Office.context.document.settings.set(xmlPartId, asyncResult.value.id);

            Office.context.document.settings.saveAsync(function(asyncResult) {
              if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                console.log("Settings save failed. Error: " + asyncResult.error.message);
              } else {
                console.log("Saved new XML Part");
                functionAsParam(xmlPartId);
              }
            });
          });
        });
      });
    }

    this.setState({ value: "" });
  };

  handleChange(event) {
    this.setState({ value: event.target.value });
  }

  public render() {
    const { items, groups, isCompactMode } = this.state;

    return (
      <div>
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

  private _onRenderDetailsHeader(props: IDetailsHeaderProps, _defaultRender?: IRenderFunction<IDetailsHeaderProps>) {
    return <DetailsHeader {...props} ariaLabelForToggleAllGroupsButton={"Expand collapse groups"} />;
  }

  private _onRenderColumn(item: IDetailsListGroupedExampleItem, _index: number, column: IColumn) {
    const value =
      item && column && column.fieldName ? item[column.fieldName as keyof IDetailsListGroupedExampleItem] || "" : "";

    return <div data-is-focusable={true}>{value}</div>;
  }
}
