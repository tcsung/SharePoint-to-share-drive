# Use SharePoint as network drive on Windows computer.


## Preface
If you like SharePoint, then you probably seeking to use it like a network drive, this tool can accomplish this on your Windows computer. 

## Requirement
Windows 10 computer, if computer still using Windows 7 then try to install Powershell 5.1  and also .Net 4 before use it.  Never test for Windows 8 but should be fine if you have Powershell 5.1 installed.

## How to
The idea for this script is to use Internet Explorer to complete the SharePoint authentication.  First, in the script GUI you need to provide your SharePoint address and then click 'Run Internet Explorer', after you login your SharePoint then go back to this tools, provide the share folder path to let script help you to map the network drive.

## Known issues and solutions
1. If this is first time you run this script, it may not work because the registry update to the IE is not yet functioning, please try close the everything and run the script again.
2. If you found the network drive was malfunction after a period of time, try to use the same script again, probably because your SharePoint login session was expired.
3. If you cannot map the drive but you ensure that already login the SharePoint in your IE, then probably the folder path you input wrongly.  Try double check:
- Whether the folder name contain spare, if so change the space character (' ') to %20
- Whether the folder name contain dot sign, if so change the dot character ('.') to %2E
