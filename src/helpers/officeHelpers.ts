
export function extractCaseId(){
Word.run(function (context) {
    var document = context.document;
    document.properties.load("author, title");
    
    return context.sync()
      .then(function () {
        let title = document.properties.title;
        getCaseId(title);
      });
  });
}

export function getCaseId(title){
    return title;
}

//This function adds a new XMLPart to the document
