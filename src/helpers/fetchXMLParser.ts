import { uuid } from "uuidv4";

//export default
// const _xmlString =
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
//   "</fetch>";

// export default
export class FetchXMLHelper {
  fetchXML;
  strippedItems = [];
  strippedGroups = [];

  constructor(fetchXML) {
    this.fetchXML = fetchXML;
  }

  parseFetchXML() {
    var node = new DOMParser().parseFromString(this.fetchXML, "text/xml").documentElement;

    var nodes = node.querySelectorAll("*");
    var nodeName = null;
    let nodeValue = null;
    //let itemsLength = 0;
    // let arrGroupsItems = [];
    // let objGroupItem = {};

    //get the number of items under a group
    for (let i = 0; i < nodes.length; i++) {}
    for (var i = 0; i < nodes.length; i++) {
      //var text = null;
      //if (nodes[i].childNodes.length == 1 && nodes[i].childNodes[0].nodeType == 3)
      //if nodeType == text node
      nodeName = nodes[i].nodeName; //get text of the node
      nodeValue = nodes[i].getAttribute("name");
      if (nodeName === "entity") {
        // this.strippedGroups.push(nodeValue);
        let groupsObj = {
          key: uuid(),
          name: nodeValue,
          startIndex: 0,
          count: 5,
          level: 0
        };
        this.strippedGroups.push(groupsObj);
      }
      if (nodeName === "attribute") {
        let stateObj = {
          key: uuid(),
          name: nodeValue
        };
        this.strippedItems.push(stateObj);
      }
      // this.strippedItems.push(nodeName);
    }
    console.log("Inside fetchsml Module ", this.strippedGroups);
    console.log("Inside fetchxml module....", this.strippedItems);
  }
  getStrippedItems() {
    return this.strippedItems;
  }
  getStrippedGroups() {
    return this.strippedGroups;
  }
}
