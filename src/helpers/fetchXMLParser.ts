import { uuid } from "uuidv4";

export class FetchXMLHelper {
  fetchXML;
  strippedItems = [];
  strippedGroups = [];
  // groupsObj = {
  //   count:0

  // };
   groupsObj = {
    key: "",
    name: "",
    startIndex: 0,
    count: 0,
    level: 0
  };

  static xmlPartIds = [];

  constructor(fetchXML) {
    this.fetchXML = fetchXML;
  }

  parseFetchXML() {
    var node = new DOMParser().parseFromString(this.fetchXML, "text/xml").documentElement;

    var nodes = node.querySelectorAll("*");
    var nodeName = null;
    let nodeValue = null;
    
    //get the number of items under a group

    //get rid of the for loop and hardcode the object
    //need to use attribute as counter
    //for (let i = 0; i < nodes.length; i++) {}

    let itemCounter = 0;
    for (var i = 0; i < nodes.length; i++) {

      
     
      nodeName = nodes[i].nodeName; //get text of the node
      nodeValue = nodes[i].getAttribute("name");
      if (nodeName === "entity") {
         this.groupsObj = {
          key: uuid(),
          name: nodeValue,
          startIndex: 0,
          count: itemCounter,
          level: 0
        };
        //
        //this.strippedGroups.push(groupsObj);
      }
      if (nodeName === "attribute") {
        let stateObj = {
          key: uuid(),
          name: nodeValue
        };
        itemCounter++;
        this.strippedItems.push(stateObj);
      }
    }
    this.groupsObj["count"] = itemCounter;
    this.strippedGroups.push(this.groupsObj);
    console.log("Item Counter ", itemCounter )
    console.log("Inside fetchxml Module ", this.strippedGroups);
    console.log("Inside fetchxml module....", this.strippedItems);
  }
  getStrippedItems() {
    return this.strippedItems;
  }
  getStrippedGroups() {
    return this.strippedGroups;
  }
}
