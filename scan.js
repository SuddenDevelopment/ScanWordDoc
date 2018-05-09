var fs = require('fs')
var strFile = '';
var CFB = require('cfb');
//get the command line arg for the file

fnTest = function(strFile){
  console.log('go');
  var cfb = CFB.read(strFile, {type: 'file'});
  //console.log( CFB.find(cfb, 'Module1') );
  //build the return object that has all the answers
  var objResponse={
    "fileType":'doc',
    "hasMacro":false,
    "macroTxt":'',
    "isValid":false,
    "hasContent":false
  };
  //var data = workbook.content;

  var arrFiles = cfb.FullPaths;

  for(var i=0; i<arrFiles.length;i++){
    if(arrFiles[i] === 'Root Entry/Macros/'){
      objResponse.hasMacro=true;
    }
  }

  //find Content, no content is no bueno
  var objDoc = CFB.find(cfb, 'WordDocument');
  var strData = objDoc.content;
  if(strData.length > 1){
    objResponse.hasContent=true;
  }

  console.log(objResponse);
  return objResponse;
}

process.argv.forEach(function (val, index, arrArguments) {
  if(arrArguments.length === 3){ 
  	//test the file
  		if( /^.[doc|docx]$/i.test(arrArguments[2]) === false ){
        strFile=arrArguments[2];
        fnTest(strFile);
      }else{}
  }
});
