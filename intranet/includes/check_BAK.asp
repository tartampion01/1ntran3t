<%
'Force SSL
If Request.ServerVariables("SERVER_PORT")<>443 AND Request.ServerVariables("SERVER_NAME")<>"test.reseaudynamique.com" Then
  Dim strSSLURL
  strSSLURL= "https://"
  strSSLURL= strSSLURL & "www.reseaudynamique.com" 'Request.ServerVariables("SERVER_NAME")
  strSSLURL= strSSLURL & Request.ServerVariables("URL")
  If LEN(Request.QueryString)>2 then strSSLURL = strSSLURL & "?" & Request.QueryString
  Response.Redirect strSSLURL
End If

Dim Dealers(8,1)
Dealers(0,0)="Master"
Dealers(1,0)="CCB"
Dealers(2,0)="CIA"
Dealers(3,0)="CIE"
Dealers(4,0)="CIWI"
Dealers(5,0)="GR"
Dealers(6,0)="RDL"
Dealers(7,0)="CCA"
Dealers(8,0)="CI"

Dealers(0,1)="Centre du camion Beaudoin"
Dealers(1,1)="Centre du camion Beaudoin"
Dealers(2,1)="Camions Inter-Anjou"
Dealers(3,1)="Camions Inter-Élite"
Dealers(4,1)="Camions Inter-West-Island"
Dealers(5,1)="Garage Redmond"
Dealers(6,1)="Rivière-du-loup"
Dealers(7,1)="Centre du camion Amiante"
Dealers(8,1)="Charest International"

Dim UserLevel(2)
	UserLevel(0)="Utilisateur"
	UserLevel(1)="Administrateur"

Function DealerCombo(dealer_id)
	If Session("dealer_id")=0 then	
		Dim i
		For i = 0 to Ubound(dealers)
			IF i=dealer_id then
				strHTM = "selected"
			Else
				strHTM = ""
			End If
			dealerCombo = dealerCombo & "<option value='" & i & "' " & strHTM & ">" & dealers(i,0) & "</option>" & VbCrLf
		Next
	Else
		if dealer_id <> "" then
			dealerCombo = "<option value='" & dealer_id & "' selected>" & dealers(dealer_id,0) & "</option>"
		Else
			dealerCombo = "<option value='" & session("dealer_id") & "' selected>" & dealers(session("dealer_id"),0) & "</option>"
		End If
	End If
End Function

Function DeleteFile(vfilename)
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	If FSO.FileExists(Server.MapPath(vfilename)) then
		FSO.DeleteFile Server.MapPath(vfilename)
	End If
End Function

Function ResizeImage(FileName,MaxWidth,MaxHeight,Overwrite)
	'Reduce and compress large image
	
	on error resume next
	
	Dim Image
	
	Filename = Server.MapPath(Filename)

	Set Image = Server.CreateObject("GflAx.GflAx")
	
	FileName=lcase(FileName)

	Image.LoadBitmap FileName
	Image.SaveJPEGQuality = 60
	
	'ignore  X or Y modes
	IF MaxWidth=0 then MaxWidth=Image.Width
	IF MaxHeight=0 then MaxHeight=Image.Height 
	
	If Image.Height > MaxWidth  or Image.Height > MaxHeight then
		If Image.Width > MaxWidth then
			Image.Resize MaxWidth,Int(Image.Height*MaxWidth/Image.Width) 
		End IF
		If Image.Height > MaxHeight then
			Image.Resize Int(Image.Width*MaxHeight/Image.Height),MaxHeight
		End If

		If Overwrite then
			Image.SaveBitmap FileName
		Else
			LastSlashPos = InStrRev(Filename,"\")
			Image.SaveBitmap Left(Filename,LastSlashPos) & "small\" & Right(Filename,Len(Filename)-LastSlashPos)
		End If
	
	End if

	if err.number <> 0 then
		response.write "An error occured creating Image:" & err.description & "<br>"
	End If

	Set Image = Nothing
	
	'response.end 

	
End Function


Function MakeLocalDate(dteDate)
	'** Converts a Date to yyyy/mm/dd String
	If dteDate <> "" then
		yyyy = Year(dteDate)
		mm = Right(FormatNumber(Month(dteDate)/100,2),2)
		dd= Right(FormatNumber(Day(dteDate)/100,2),2)
		MakeLocalDate = dd & "/" & mm & "/" & yyyy
	End IF
End Function

Function MySQLDate(strDate)
	'** Converts a Date from dd/mm/yyyy to yyyy/mm/dd String
	MySQLDate=""
	on error resume next
	If strDate <> "" then
		Dim dte
		dte = split(strDate,"/")
		yyyy = left(dte(2),4)
		mm = dte(1)
		dd= dte(0)
		MySQLDate = yyyy & "/" & mm & "/" & dd
	End IF
End Function

Function GetUserName(ID)
'*** Return a UserID First & Last Name
	Dim Conn,rs
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("cstring")
	
	If ID <> "" then
		Set rs=Conn.Execute("Select FirstName, LastName from intranet_users WHERE ID=" & ID)
		If not rs.EOF then
			GetUserName = Server.HTMLEncode(rs("FirstName") & " " & rs("LastName"))
		End If
		Set rs=Nothing	
	End If
	
	Conn.Close
	Set Conn=Nothing
End Function

Function GetUserEmail(ID)
'*** Return a UserID email Address
	Dim Conn,rs
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("cstring")
	
	If ID <> "" then
		Set rs=Conn.Execute("Select email from intranet_users WHERE ID=" & ID)
		If not rs.EOF then
			GetUserEmail = rs("email")
		End If
		Set rs=Nothing	
	End If
	
	Conn.Close
	Set Conn=Nothing
End Function

Function MakeUserCombo(SelectID)
'**************************************************************
'* Returns an html combo list of all users
'* with <option value=ID>FullName</option>
'* Where the SelectID is the selected value
'* Send 0 to get no selected
'**************************************************************
	Dim Conn,rs
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("cstring")
	
	If SelectID=0 then alluserCombo = "<option value='0'>S&eacute;lectionnez -></option>"
	
	Set rs=Conn.Execute("SELECT ID, CONCAT(FirstName,' ',LastName) As FullName from intranet_users ORDER BY FullName;")
	Do While Not rs.EOF
		IF rs("ID")=SelectID then
			strHTM = "selected"
		Else
			strHTM = ""
		End If
		alluserCombo = alluserCombo & "<option " & strHTM & " value='" & rs("ID") & "'>" & Server.HTMLEncode(rs("FullName")) & "</option>" & VbCrLf
		rs.MoveNext
	Loop
	MakeUserCombo = alluserCombo

	Conn.Close
	Set Conn=Nothing
End Function

Function DeletePictures(id,vpath)
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	If FSO.FileExists(Server.MapPath(vpath & "pic"&ID&".jpg")) then
		FSO.DeleteFile Server.MapPath(vpath & "pic"&ID&".jpg")
	End If
	If FSO.FileExists(Server.MapPath(vpath & "small/pic"&ID&".gif")) then
		FSO.DeleteFile Server.MapPath(vpath & "small/pic"&ID&".gif")
	End If
End Function

Function CopyPictures(old_id,new_id,vpath)
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	If FSO.FileExists(Server.MapPath(vpath & "pic"&old_id&".jpg")) then
		FSO.CopyFile Server.MapPath(vpath & "pic"&old_id&".jpg"),Server.MapPath(vpath & "pic"&new_id&".jpg")
	End If
	If FSO.FileExists(Server.MapPath(vpath & "small/pic"&old_id&".gif")) then
		FSO.CopyFile Server.MapPath(vpath & "small/pic"&old_id&".gif"),Server.MapPath(vpath & "small/pic"&new_id&".gif")
	End If
End Function

fullPath = Request.ServerVariables("PATH_INFO")
IF Request.QueryString <> "" then
	Query = "?" & Request.QueryString
End If
URL = Server.URLEncode(fullPath&Query)

'Path = lcase(Left(fullPath,(InstrRev(Request.ServerVariables("PATH_INFO"),"/"))))

If Session("userID") = "" AND Request.Cookies("userID") = "" then
	Response.Redirect("/intranet/login/?URL=" & URL)
End If

If Session("userID")="" then
	Session("userID")=Request.Cookies("userID")
	Session("username")=Request.Cookies("username")
	Session("fullname")=Request.Cookies("fullname")
End If



'Check for administrative rights
If Session("Level")="" then
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("cstring")

	Set rs=Conn.Execute("SELECT Level,Dealer_id from intranet_users WHERE ID="& Session("userID")& ";")
	If not rs.EOF then
		Session("Level")=rs("Level")
		Session("Dealer_id")=rs("Dealer_id")
	End If
	
	Set rs=Nothing
	Conn.Close
	Set Conn=Nothing
End If	


'Check for rights
IF Cint(Session("Level")) < Cint(PageLevel) then
	Response.Redirect("/intranet/login/?err=3&URL=" & Path)
End IF

%>