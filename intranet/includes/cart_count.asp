<%
intOrderID = Request("intOrderID")
Set Conn2 = Server.CreateObject("ADODB.Connection")
Conn2.Open Application("cstring")
Total = 0
If intOrderID<>"" then
	Set rs2 = Conn2.Execute("Select count(orderID) AS tot from itemsOrdered WHERE orderID=" & intOrderID & ";")
	If not rs2.EOF then
		Total = rs2("tot")
	End If
End If

Set rs2=Nothing
Conn2.Close
Set Conn2=Nothing
If Total > 1 then
	plurial = "s"
Else
	plurial = ""
End IF
If Total > 0 then
	Response.Write "Your quote currently has " & Total & " item" & plurial & "."
Else
	Response.Write " "
End If
%>
