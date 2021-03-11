import { uuid } from "uuidv4";
//John: import {tableFields} from "../taskpane/components/GroupedComponent"
import {parseString,Builder} from 'xml2js';

export class FetchXMLHelper {
  fetchXML = "";
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
    //John: let tablesFields = this.getTablesFields();
    for(let i = 0; i < tableFields && tableFields.length; i++){ //John: 

      console.log("Loop....",tableFields[i]);
    }
      return tableFields;
  }

  parseFetchXMLNoParams() {
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
    //John: let tablesFields = this.getTablesFields();
    
  }


   insertFilterWithCaseId(caseId){    
    
    parseString(this.fetchXML,  function(err,result){
     if(result){
       console.log("result now ...",result)
      if(result.AddIn.fetch[0] !== null || result.AddIn.fetch[0] !== undefined){
        result.AddIn.fetch[0].entity[0].filter = [{condition: {$: {attribute:"incidentid",operator:"eq",value:caseId}}}];
        const xmlBuilder = new Builder();
        let newXml = xmlBuilder.buildObject(result);      
        console.log("New XML ",newXml);
        
      }
      else{
        console.log("Fetch xml is null or undefined");
      }
     }
     else if (err){
       console.log(err);
     }
    
     });     
     
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
