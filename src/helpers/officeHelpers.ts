
export function getCaseId(){
Word.run(function (context) {
    var document = context.document;
    document.properties.load("author, title");
    
    return context.sync()
      .then(function () {
        console.log("The author of this document is " + document.properties.author + " and the title is '" + document.properties.title + "'");
        return document.properties.title.toString();
      });
  });
}