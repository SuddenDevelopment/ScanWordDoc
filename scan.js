var libFs = require('fs');
var libCfb = require('cfb');
var libPath = require('path');

//these are the only allowed extensions:
var arrAllowed=['.doc','.xls'];

fnTest = function(strFile){
  console.log('go');
  var objFileInfo = libFs.lstatSync(strFile);
  if(objFileInfo.isDirectory() === true){
    console.log(strFile + ' is a directory, processing.')
    fnProcessDirectory(strFile);
  }else if(objFileInfo.isFile() === true){
    var strExtension = libPath.extname(strFile);
    if(arrAllowed.indexOf(strExtension) > -1){ fnTestFile(strFile); }
    else{ console.log(strExtension + ' extension not supported on file: ' + strFile); }
  }
}

fnProcessDirectory=function(strDirectory){
  //add a trailing slash, makes it more consistent
  if( strDirectory.charAt(strDirectory.length-1) !== '/' ){ strDirectory=strDirectory+'/'; }
  libFs.readdir(strDirectory, (err, arrFiles) => {
    for(var i=0;i<arrFiles.length;i++){
      //console.log(arrFiles[i]);
      fnTest(strDirectory+arrFiles[i]);
    }
  })
}

fnTestFile=function(strFile){
  var objCfb = libCfb.read(strFile, {type: 'file'});
  var objResponse={
    "fileType":'doc',
    "hasMacro":false,
    "macroTxt":'',
    "isValid":false,
    "hasContent":false
  };
  //var data = workbook.content;

  var arrFiles = objCfb.FullPaths;

  for(var i=0; i<arrFiles.length;i++){
    if(arrFiles[i] === 'Root Entry/Macros/'){
      objResponse.hasMacro=true;
    }
  }

  //find Content, no content is no bueno
  var objDoc = libCfb.find(objCfb, 'WordDocument');
  var strData = objDoc.content;
  if(strData.length > 1){
    objResponse.hasContent=true;
  }

  console.log(objResponse);
  return objResponse;
}

//grab the 3rd argument and run it, tests for what it is are inside the functions
fnTest(process.argv[2]);
