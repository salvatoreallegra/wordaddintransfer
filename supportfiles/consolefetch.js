function test() {
  var fetchXmlQuery = `<fetch attribute="schemaname" operator="eq" >
    <entity name="pacts_militaryrecord" >
    <link-entity name="contact" from="contactid" to="pacts_clientid" >
    <link-entity name="incident" from="customerid" to="contactid" >
    <filter>
    <condition attribute="incidentid" operator="eq" value="c545391a-74f8-4708-9d6e-5f0a9f9be0e9" />
    </filter>
    </link-entity>
    </link-entity>
    </entity>
    </fetch>
    `;

  const globalContext = Xrm.Utility.getGlobalContext();
  const crmUrl = globalContext.getClientUrl();

  var req = new XMLHttpRequest();
  req.open("GET", crmUrl + "/api/data/v9.0/pacts_militaryrecords?fetchXml=" + encodeURI(fetchXmlQuery), true);
  req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
  req.onreadystatechange = function() {
    if (this.readyState === 4) {
      req.onreadystatechange = null;
      if (this.status === 200) {
        var results = JSON.parse(this.response);
        console.log("Fetch in console ", results);
      } else {
        alert(this.statusText);
      }
    }
  };
  req.send();
}

function test() {
  var fetchXmlQuery = `<fetch>
  <entity name="pacts_presentenceinvestigation" >
    <attribute name="pacts_typeid" />
    <attribute name="pacts_convictionbycode" />
    <attribute name="pacts_caseid" />
    <attribute name="pacts_isdefenseattorneypresent" />
    <link-entity name="incident" from="incidentid" to="pacts_caseid" >
      <filter>
        <condition attribute="incidentid" operator="eq" value="c545391a-74f8-4708-9d6e-5f0a9f9be0e9" />
      </filter>
    </link-entity>
  </entity>
</fetch>`;

  const globalContext = Xrm.Utility.getGlobalContext();
  const crmUrl = globalContext.getClientUrl();

  var req = new XMLHttpRequest();
  req.open("GET", crmUrl + "/api/data/v9.0/pacts_presentenceinvestigations?fetchXml=" + encodeURI(fetchXmlQuery), true);
  req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
  req.onreadystatechange = function() {
    if (this.readyState === 4) {
      req.onreadystatechange = null;
      if (this.status === 200) {
        var results = JSON.parse(this.response);
        console.log("Fetch in console ", results);
      } else {
        alert(this.statusText);
      }
    }
  };
  req.send();
}

// ****************another one with attributes

function test() {
  var fetchXmlQuery = `<fetch attribute="schemaname" operator="eq" >
    <entity name="pacts_militaryrecord" >
    <attribute name="pacts_dateentered" />  
    <attribute name="pacts_name" />  
    <link-entity name="contact" from="contactid" to="pacts_clientid" >
    <link-entity name="incident" from="customerid" to="contactid" >
    <filter>
    <condition attribute="incidentid" operator="eq" value="c545391a-74f8-4708-9d6e-5f0a9f9be0e9" />
    </filter>
    </link-entity>
    </link-entity>
    </entity>
    </fetch>
    `;

  const globalContext = Xrm.Utility.getGlobalContext();
  const crmUrl = globalContext.getClientUrl();

  var req = new XMLHttpRequest();
  req.open("GET", crmUrl + "/api/data/v9.0/pacts_militaryrecords?fetchXml=" + encodeURI(fetchXmlQuery), true);
  req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
  req.onreadystatechange = function() {
    if (this.readyState === 4) {
      req.onreadystatechange = null;
      if (this.status === 200) {
        var results = JSON.parse(this.response);
        console.log("Fetch in console ", results);
      } else {
        alert(this.statusText);
      }
    }
  };
  req.send();
}
