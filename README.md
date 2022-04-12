## Usage

Run cmd.exe as Administrator

Navigate to your Office folder:
```
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
```

Convert your Office license to volume one if possible:
If your Office is got from Microsoft, this step is required. On the contrary, if you install Office from a Volume ISO file, this is optional so just skip it if you want.
```
for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
```

Use KMS client key to activate your Office.
Make sure your PC is connected to the internet, then run the following command:
```
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /unpkey:BTDRB >nul
cscript ospp.vbs /unpkey:KHGM9 >nul
cscript ospp.vbs /unpkey:CPQVG >nul
cscript ospp.vbs /sethst:s8.uk.to
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act
```

If you see the error 0xC004F074, it means that your internet connection is unstable or the server is busy. Please make sure your device is online and try the command “act” again until you succeed.

**More details on**: https://msguides.com/
