<% 
'Copyright of Denude Clothing Ltd - Ryan Fletcher (CEO)

'1. OpenRecordSet
	Public Function OpenRecordSet(InstanceVar,ConnectionVar,SourceVar)
		Set InstanceVar = CreateObject("ADODB.Recordset")
		InstanceVar.ActiveConnection = DEN_Conn(0)
		InstanceVar.Source = SourceVar
		InstanceVar.CursorType = 3
		InstanceVar.CursorLocation = 2
		InstanceVar.Open()
	End Function

'2. GetBasket: Function to retrieve the users or guests basket. Arguments: UserIDVar = Session("UserID") GuestIDVar = Session("GuestIDVar"). Depending on the values of the two arguments will determine the SourceVar to use. 	
	Public Function GetBasket(UserIDVar,GuestIDVar)
		'Determine whether the customer is a user or a guest.
	  	 If UserIDVar <> "" Then 
	        SourceVar = "SELECT Basket.ItemID, Basket.UserID, Basket.Quantity, Items.ItemName, Items.UnitPrice, Items.ImageSource FROM Basket INNER JOIN Items ON Basket.ItemID = Items.ItemID WHERE UserID =" & UserIDVar 
	       ElseIf GuestIDVar <> "" Then
	        SourceVar = "SELECT Basket.ItemID, Basket.UserID, Basket.Quantity, Items.ItemName, Items.UnitPrice, Items.ImageSource FROM Basket INNER JOIN Items ON Basket.ItemID = Items.ItemID WHERE UserID =" & GuestIDVar 
		  Else
		 End If
		Call OpenRecordSet(GetBasket,0,SourceVar)
		  If GetBasket.EOF Then %>
           <tr>
            <td colspan="6" align="center" valign="top">
             <p id="FormLabel">You do not have any items in your basket, please select 'Add to basket' to add items to your basket</p>
            </td>
           </tr>
		<% Else %>
        <form method="post" action="Basket-Checkout-OrderSummary.asp" name="ProcessOrder" id="ProcessOrder">
		<%  While NOT GetBasket.EOF
		   				ItemIDVar = GetBasket.Fields("ItemID")
						ItemNameVar = GetBasket.Fields("ItemName")
						UnitPriceVar = GetBasket.Fields("UnitPrice")
						ImageSourceVar = GetBasket.Fields("ImageSource")
						QuantityVar = GetBasket.Fields("Quantity") %>
                <tr>
                 <td></td>
                 <td>Item Name</td>
                 <td width="11%">Price</td>
                 <td width="30%">Quantity</td>
                </tr>
                <tr>
                 <td id="Footer" width="23%" height="101" align="center">
                  <img width="120" height="120" src="../../Images/<%= ImageSourceVar %>" name="<%= ItemNameVar %>" /></td>
                 <td id="Footer" width="15%">
                  <p><%= ItemNameVar %></p>
                 </td>
                 <td id="Footer">
                  <input type="text" width="1" value="<%= QuantityVar %>" name="Quantity<%=ItemIDVar %>">
                 </td>
                 <td id="Footer">
                 <select name="ItemSizes">
                 	<option selected="selected" value="0">Please select a size</option>
                    <% SourceVar = "SELECT ItemSizes.ItemSizeID, ItemSizes.ItemID,  ItemSizes.ItemSize, ItemSizes.Available FROM ItemSizes WHERE ItemID = " & ItemIDVar & " AND Available = 1"
						Call OpenRecordSet(ItemSizes,0,SourceVar)
							While NOT ItemSizes.EOF
                            	ItemSizeIDVar = ItemSizes.Fields("ItemSizeID")
								ItemSizeVar = ItemSizes.Fields("ItemSize") %>
                            <option value="<%= ItemSizeIDVar %>"><%= ItemSizeVar %></option>
                   <%         ItemSizes.MoveNext()
				  			Wend
						ItemSizes.Close()	 %> 
                 </td>
                 <td id="Footer">
                  <p>£<%= UnitPriceVar %></p>
                 </td>
                 <td id="Footer" width="21%">
                  <p><a href="Basket-Remove-Process.asp?ItemID=<%= ItemIDVar %>&amp;Redirect=Basket">X</a></p>
                 </td>
                </tr> 
               <% GetBasket.MoveNext()
		       Wend
		      GetBasket.Close()
				If UserIDVar <> "" Then 
				   SourceVar = "SELECT SUM(Items.UnitPrice * Basket.Quantity) AS Total,Basket.UserID FROM Basket INNER JOIN Items ON Basket.ItemID = Items.ItemID WHERE UserID =" & UserIDVar & " GROUP BY Basket.UserID"
				ElseIf GuestIDVar <> "" Then
				    SourceVar = "SELECT SUM(Items.UnitPrice * Basket.Quantity) AS Total,Basket.GuestID FROM Basket INNER JOIN Items ON Basket.ItemID = Items.ItemID WHERE GuestID =" & GuestIDVar & " GROUP BY Basket.GuestID"
				End If
			   	Call OpenRecordSet(TotalOrderValue,0,SourceVar)
					While NOT TotalOrderValue.EOF
			     TotalOrderValueVar = TotalOrderValue.Fields("Total")
			TotalOrderValue.MoveNext()
			Wend
		TotalOrderValue.Close()
        If TotalOrderValueVar <> "" Then %>
           <tr>
            <td colspan="6" id="Footer"><p>Total: £<%= TotalOrderValueVar %></p></td>
           </tr>
           <tr>
            <td><input type="submit" value="Proceed to Checkout" /></td></tr>
			<% End If %>
            </form>
	<%
	    End If 
	 End Function

'	 Private Function GetRecordSet(InstanceVar,TableVar,Split(SelectVar,"'"),ArgumentVar,ArgumentValueVar,ConditionVar)
'	 			SourceVar = "SELECT" & SelectVar & "FROM" & TableVar
'				If ArgumentValueVar <> "None" Then
'				SourceVar = SourceVar & "WHERE" & ArgumentVar & ConditionVar & ArgumentValueVar
'				End If
'					Call OpenRecordSet(InstanceVar,1,SourceVar)
'					 ReDim ColumnArrayVar()
'					 	For Each ColumnArrayVar
'							ColumnArrayVar(SelectVar) = InstanceVar.Fields(SelectVar)
'						Next
'					InstanceVar.Close()		
'	 End Function
'	 
'     Private Function GetSQLColumns(RecordSetVar,TableVar,ColumnVar,ArgumentVar,ArgumentValueVar,WhileNOTVar, HTMLVar,IfEOFVar,IfEOFValueVar)
'         If InStr(RecordSetVar,1) = "D" Then
'                SourceVar = "SELECT DiscussionID, ClientID, DiscussionTitle, DiscussionSubject, CreationDate FROM " & TableVar                
'                    If Len(ColumnVar) <> 0 Then
'                           SourceVar = SourceVar & "WHERE " & ColumnVar & " " & ArgumentVar & " " ArgumentValueVar
'                    Else
'                    End If
'                    Call OpenRecordSet(RecordSetVar,0,SourceVar)
'                    Call GetHTML()
'                        If IfEOFVar <> 0 Then
'                            If RecordSetVar.EOF Then
'                                Response.Write(IfEOFValueVar)
'                            End If
'                        End If        
'                        If WhileVar <> 0 Then
'                            While NOT RecordSetVar.EOF
'                                      DiscussionIDVar = RecordSetVar.Fields("DiscussionID")
'                                      DiscussionTitleVar = RecordSetVar.Fields("DiscussionTitle")
'                                      DiscussionSubjectVar = RecordSetVar.Fields("DiscussionSubject")
'                                       ClientIDVar = RecordSetVar.Fields("ClientID")
'                                            Response.Write(HTMLVar)
'                                        RecordSetVar.MoveNext()
'                                    Wend        
'                            RecordSetVar.Close()
'                        Else
'                            Response.Write(HTMLVar)
'                        End If                  
%>