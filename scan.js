var fs = require('fs')
var strFile = '';
//get the command line arg for the file
process.argv.forEach(function (val, index, arrArguments) {
  if(arrArguments.length === 3){ 
  	//test the file
  	//if(rgxTest.test(arrArguments[2])){ 
  		strFile=arrArguments[2]; 
  	//}else{ console.log('invalid file'); }
  	console.log( /^[a-z]:((\\|\/)[a-z0-9\s_@\-^!#$%&+={}\[\]]+)+\.[doc|docx]$/i.test(strFile) );
  }
});

/*
//pass the file argument to be run
if(strFile !== ''){
	fs.readFile(strFile, 'utf8', function (err,data) {
	  if (err) {
	    return console.log(err);
	  }
	  console.log(data);
	});
}
*/


/*  this is for docxespecifically, not doc
var dxe = require('docx-extractor');
 
dxe.lastModified(strFile, function(data){
    console.log(data)
});

*/

var WordExtractor = require("word-extractor");
var extractor = new WordExtractor();
var extracted = extractor.extract(strFile);
extracted.then(function(doc) {
  console.log(doc);
  //console.log(doc.getBody());
});
