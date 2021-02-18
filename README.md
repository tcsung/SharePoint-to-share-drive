# Use SharePoint as network drive on Windows computer.


## Preface
If you like SharePoint, then you probably seeking to use it like a network drive, this tool can accomplish this on your Windows computer. 

## Requirement
Windows 10 computer, if computer still using Windows 7 then try to install Powershell 5.1  and also .Net4 before use it.  Never test for Windows 8 but should be fine if you have Powershell 5.1 installed.

## How to
The idea for this script is to use Internet Explorer to complete the SharePoint authentication, then user just provide their share folder path to let script help you to map the network drive. It is simple enough for you to run the script to enable the feature. If this is first time you run this script, it may not work because the registry update to the IE is not yet functioning, please try close the everything and run the script again.

If you found the network drive was malfunction, try to use the same script again, probably because your SharePoint login session was expired.
