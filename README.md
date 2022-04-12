## Microsoft Office Activation using Deployment Tool
Run `officedeploymenttool_14931-20120.exe`<br><br>
Extract files to `C:\Office` folder<br><br>
Modify the file according to the desired office suite<br>
`configuration-Office2021Enterprise.xml`:
```xml
<!-- Office 2021 enterprise client configuration file sample. To be used for Office 2021 
     enterprise volume licensed products only, including Office 2021 Professional Plus,
     Visio 2021, and Project 2021.

     Do not use this sample to install Office 365 products.

     For detailed information regarding configuration options visit: http://aka.ms/ODT. 
     To use the configuration file be sure to remove the comments

     The following sample allows you to download and install Office 2021 Professional Plus,
     Visio 2021 Professional, and Project 2021 Professional directly from the Office CDN.

     This configuration file will remove all other Click-to-Run products in order to avoid
     product conflicts and ensure successful setup.
 -->



<Configuration>

  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume">
      <Language ID="en-us" />
      <ExcludeApp ID="Lync" />
    </Product>
    <Product ID="VisioPro2021Volume">
      <Language ID="en-us" />
    </Product>
    <Product ID="ProjectPro2021Volume">
      <Language ID="en-us" />
    </Product>
  </Add>

  <Remove All="True" />

  <!--  <RemoveMSI All="True" /> -->

  <!--  <Display Level="None" AcceptEULA="TRUE" />  -->

  <!--  <Property Name="AUTOACTIVATE" Value="1" />  -->

</Configuration>

```
Don't forget to modify <Product ID="ProPlus2021Volume"> variables according to the desired products.<br>
#### Here you can find all supported ID for Office Deployment:<br>
https://docs.microsoft.com/fr-fr/office365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run

<hr>
  
Run `cmd.exe` as Administrator and execute the following command:<br>
`.\setup.exe /configure .\configuration-Office2021Enterprise.xml`
<br>
  
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
