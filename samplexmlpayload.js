{
  "FetchXml" : "<AddIn xmlns='http://schemas.pacts.com/case'>
  <fetch mapping='blah'>  
     <entity name='system'>   
        <attribute name='processor'/>
        <attribute name='alu'/>
        <attribute name='bus'/>
              
        <link-entity name='systemuser' to='owninguser'>   
           <filter type='and'>   
              <condition attribute='lastname' operator='ne' value='Cannon' />   
            </filter>   
        </link-entity>   
     </entity>   
  </fetch>
  </AddIn>"
}
