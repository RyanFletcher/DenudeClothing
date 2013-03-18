<% 
	Public Function OpenRecordSet(InstanceVar,ConnectionVar,SourceVar)
		Set InstanceVar = CreateObject("ADODB.Recordset")
		InstanceVar.ActiveConnection = ControlPanel(ConnectionVar)
		InstanceVar.Source = SourceVar
		InstanceVar.CursorType = 3
		InstanceVar.CursorLocation = 2
		InstanceVar.Open()
	End Function
	
	Public Function SQLTextString(StringVar)
	
		If StringVar <> "" Then
			StringVar = Replace(StringVar,"'","''")
			StringVar = Replace(StringVar,chr(34),"''")
		End If
		
		SQLTextString = StringVar
		
	End Function
	
	Public Function SQLIntegerString(StringVar)
	
		StringPermittedVar = "0123456789-."
		NewStringVar = ""
		
		For CharNoVar = 1 to Len(StringVar)
			CharVar = Mid(StringVar,CharNoVar,1)
			If InStr(StringPermittedVar,CharVar) > 0 Then NewStringVar = NewStringVar & CharVar
		Next
		
		If NewStringVar = "" Then NewStringVar = 0
		
		SQLIntegerString = NewStringVar
		
	End Function
	
	Public Function SQLCommand(InstanceVar,ConnectionVar,CommandVar)
	
		Set InstanceVar = CreateObject("ADODB.Command")
		InstanceVar.ActiveConnection = IND_Conn(ConnectionVar)
		InstanceVar.CommandText = CommandVar
		InstanceVar.Execute
		InstanceVar.ActiveConnection.Close()
		
	End Function

%>
