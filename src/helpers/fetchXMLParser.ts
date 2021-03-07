import { uuid } from "uuidv4";
//John: import {tableFields} from "../taskpane/components/GroupedComponent"

export class FetchXMLHelper {
  fetchXML;
  thisTableFields = [];
  strippedItems = [];
  strippedGroups = [];
  
   groupsObj = {
    key: "",
    name: "",
    startIndex: 0,
    count: 0,
    level: 0
  };

  constructor(fetchXML) {
    this.fetchXML = fetchXML;
   // this.endIndex = endIndex;
   //John: this.thisTableFields = tableFields;
   
  } 

  parseFetchXML(tableFields) {
    var node = new DOMParser().parseFromString(this.fetchXML, "text/xml").documentElement;

    var nodes = node.querySelectorAll("*");
    var nodeName = null;
    let nodeValue = null;    

    let itemCounter = 0;
    
    for (var i = 0; i < nodes.length; i++) {      
     
      nodeName = nodes[i].nodeName; //get text value or the name of the node
      nodeValue = nodes[i].getAttribute("name");
      if (nodeName === "entity") {
         this.groupsObj = {
          key: uuid(),
          name: nodeValue,
          startIndex: 0,
          count: itemCounter,
          level: 0
        };
        
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
      if(nodeName === "filter"){
        console.log("filter found");
      }
    }
    
    this.groupsObj["count"] = itemCounter;
 
    this.strippedGroups.push(this.groupsObj);
    //John: let tablesFields = this.getTablesFields();
    for(let i = 0; i < tableFields && tableFields.length; i++){ //John: 

      console.log("Loop....",tableFields[i]);
    }
    
    

    return tableFields;
  }

  insertFilterWithCaseId(){
    var node = new DOMParser().parseFromString(this.fetchXML, "text/xml").documentElement;

    var nodes = node.querySelectorAll("*");
    var nodeName = null;
    //let nodeValue = null;       
    
    for (var i = 0; i < nodes.length; i++) {      
     
      nodeName = nodes[i].nodeName; //get text value or the name of the node
      //nodeValue = nodes[i].getAttribute("name");
      if(nodeName === "filter"){
        console.log("filter found", nodeName);
      }
    }
  }
  getTablesFields(){
    return this.thisTableFields;
  }
  getStrippedItems() {
    return this.strippedItems;
  }
  getStrippedGroups() {
    return this.strippedGroups;
  }
 
}
