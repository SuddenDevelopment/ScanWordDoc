var fs = require('fs')
var strFile = '';
var CFB = require('cfb');
//get the command line arg for the file

fnTest = function(strFile){
  console.log('go');
  var cfb = CFB.read(strFile, {type: 'file'});
  var workbook = CFB.find(cfb, 'Workbook');
  //build the return object that has all the answers
  var objResponse={
    "hasMacro":false,
    "macroTxt":'',
    "isValid":false
  };
  //var data = workbook.content;

  var arrFiles = cfb.FullPaths;

  for(var i=0; i<arrFiles.length;i++){
    if(arrFiles[i] === 'Root Entry/Macros/'){
      objResponse.hasMacro=true;
    }
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
