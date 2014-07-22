'Determine OS bitness in VBS
'Each Function is a viable option to check for 64 or 32 bit OS
'TODO This is really dirty script and the methods probably dont work correctly, although the content is good.

FUNCTION DetermineAddressWidth()
'**************************************************************************************
	Bits = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
'**************************************************************************************
END FUNCTION

FUNCTION Is32BitOS()
'**************************************************************************************
    Const Path = "winmgmts:root\cimv2:Win32_Processor='cpu0'"
    Is32BitOS = (GetObject(Path).AddressWidth = 32)
'**************************************************************************************
END FUNCTION

FUNCTION Is64BitOS()
'**************************************************************************************
    Const Path = "winmgmts:root\cimv2:Win32_Processor='cpu0'"
    Is64BitOS = (GetObject(Path).AddressWidth = 64)
'**************************************************************************************
END FUNCTION


SUB OSBitnessReg()
'**************************************************************************************
	'Often we come across a situation where we need to see the Operating system bit information and then install application or do any customization. Also many times we have different MSI for different OS type and we are in a confusion how to deploy them through 1 package in a deployment tool. Or in many cases how to explain it to helpdesk to install which MSI for which machine.
	'It is easier to give them a script which will automatically determine the OS type and will install the corresponding application.
	'Here is a script which I use for this purpose:
	'Lines to get the computer Name
	 
	Set wshShell = WScript.CreateObject( "WScript.Shell" )
	 strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
	 
	  
	 '===============================================================
	 'To check whether the OS is 32 bit or 64 bit of Windows 7
	 '===============================================================
	 
	'Lines to detect whether the OS is 32 bit or 64 bit of Windows 7 

	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputerName & "\root\default:StdRegProv") 
	 
	   strKeyPath = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
	 
	   strValueName = "Identifier"
	 
	oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
	 '===============================================================
	 
	'Checking Condition whether the build is 64bit or 32 bit
	 
	   if (instr(strValue,"64")) then
	 
	'Perform functions for 64-bit OS.
	 End If
	 
	elseif (instr(strValue,"x86")) then
	 
	'Perform functions for 32-bit OS
	 'End If
	 '===============================================================
'**************************************************************************************
END SUB
