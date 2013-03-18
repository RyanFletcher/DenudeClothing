<% 
ReDim DEN_Conn(0)
DEN_Conn(0) = "Provider=SQLOLEDB; Data Source=209.17.116.5,1433; Network Library=DBMSSOCN; Initial Catalog=denude_clothing; User ID=ryan_fletcher; Password=Macbook3;"

	Public Function OpenRecordSet(InstanceVar,ConnectionVar,SourceVar)
		Set InstanceVar = CreateObject("ADODB.Recordset")
		InstanceVar.ActiveConnection = DEN_Conn(0)
		InstanceVar.Source = SourceVar
		InstanceVar.CursorType = 3
		InstanceVar.CursorLocation = 2
		InstanceVar.Open()
	End Function

%>
