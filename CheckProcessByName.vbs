FUNCTION CheckProcessByName(intInstances, strProcName)
'**************************************************************************************
  'Check to see if a certain number of processes are running
  'Returns Boolean TRUE if the criteria is met
  '
  'SAMPLE USAGE:
  'Dim boolMyProcess
  'boolMyProcess = CheckProcessByName (2 "serversalive.exe")
  'If boolMyProcess Then wscript.echo("All processes running.")
  
  Dim objService
  Dim i
  i = 0
  Set objService = GetObject ("winmgmts:")
  
  For Each objProcess In objService.InstancesOf ("Win32_Process")
  	If objProcess.Name = strProcName Then	
  		i = i + 1
  	End If
  Next
  
  If i = intInstances Then
  	CheckProcessByName = TRUE
  Else
  	CheckProcessByName = FALSE
  End If
'**************************************************************************************
END FUNCTION
