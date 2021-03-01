import { uuid } from "uuidv4";

export class FetchXMLHelper {
  fetchXML;
  strippedItems = [];
  strippedGroups = [];

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
    for (let i = 0; i < nodes.length; i++) {}
    for (var i = 0; i < nodes.length; i++) {
     
      nodeName = nodes[i].nodeName; //get text of the node
      nodeValue = nodes[i].getAttribute("name");
      if (nodeName === "entity") {
        let groupsObj = {
          key: uuid(),
          name: nodeValue,
          startIndex: 0,
          count: 5,
          level: 0
        };
        //
        this.strippedGroups.push(groupsObj);
      }
      if (nodeName === "attribute") {
        let stateObj = {
          key: uuid(),
          name: nodeValue
        };
        this.strippedItems.push(stateObj);
      }
    }
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
