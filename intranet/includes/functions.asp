
<%

Function Decode(sIn)
    dim x, y, abfrom, abto
    Decode="": ABFrom = "!@#$%^&*()+=-|\/{}|[]?><"

    For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next 
    For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next 
    For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next 

    abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
    For x=1 to Len(sin): y=InStr(abto, Mid(sin, x, 1))
        If y = 0 then
            Decode = Decode & Mid(sin, x, 1)
        Else
            Decode = Decode & Mid(abfrom, y, 1)
        End If
    Next
End Function

Function Encode(sIn)
    dim x, y, abfrom, abto
    Encode="": ABFrom = "!@#$%^&*()+=-|\/{}|[]?><"

    For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next 
    For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next 
    For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next 


    abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
    For x=1 to Len(sin): y = InStr(abfrom, Mid(sin, x, 1))
        If y = 0 Then
             Encode = Encode & Mid(sin, x, 1)
        Else
             Encode = Encode & Mid(abto, y, 1)
        End If
    Next
End Function 

Function ReadFile(vpath)
	dim fso,ObjFile
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Set ObjFile = fso.OpenTextFile(Server.MapPath(vpath),1)
	ReadFile = ObjFile.ReadAll
	ObjFile.Close
	Set fso = Nothing
End Function

Function SendEmail(email,subject,message,senderEmail,senderName,priority)
	on Error Resume Next
	Dim objCDOSYSMail,objCDOSYSCon

	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60 
	objCDOSYSCon.Fields.Update
	
	Set objCDOSYSMail.Configuration = objCDOSYSCon

	objCDOSYSMail.Sender = wea_senderName & " <" & wea_senderEmail & ">"

	if senderEmail = "" then
		senderEmail = wea_senderEmail
		senderName = wea_senderName
	else
		if senderName = "" then senderName = senderEmail
	end if


	objCDOSYSMail.From = senderName & " <" & senderEmail & ">"
	
	objCDOSYSMail.To = email
	
	objCDOSYSMail.Subject = subject
	objCDOSYSMail.HTMLBody = message
	
	if priority <> "" then
		if FSOExists(priority) then objCDOSYSMail.AddAttachment server.MapPath(priority)
	end if

	'if priority > 5 or priority < 1 then priority = -1 ' -1 = no priority ****TO DO
'    With cdoMessage.Fields 
' 
'        ' for Outlook: 
'        .Item(cdoImportance) = cdoHigh  
'        .Item(cdoPriority) = cdoPriorityUrgent  
' 
'        ' for Outlook Express: 
'        .Item("urn:schemas:mailheader:X-Priority") = 1 
' 
'        .Update 
'    End With 	


	objCDOSYSMail.Send

	Set objCDOSYSMail = Nothing
	Set objCDOSYSCon = Nothing
	
	If err<>0 then
		SendEmail = false
	else
		SendEmail = true
	End If

End Function

Function IsValidEmail(sEmail)
  IsValidEmail = false
  Dim regEx, retVal
  Set regEx = New RegExp

  regEx.Pattern ="^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,6}$" 
  regEx.IgnoreCase = true
  retVal = regEx.Test(sEmail)
  If not retVal Then
    exit function
  End If
  IsValidEmail = true
End Function


Function SetUserSession(user_ID,user_PW)
' populates all user sessions variables
Dim rs
'	Set Conn = Server.CreateObject("ADODB.Connection")
'	Conn.Open Application("cstring")
	Set rs=Conn.Execute("SELECT *, ifnull(DATEDIFF(CURDATE(), passupdate),'100') AS passupdate from users WHERE ID='" & user_ID & "' AND password='" & nEncrypt(user_PW,"") & "';")
	If not rs.EOF then
		Session("userID")= rs("ID")
		Session("userPW")= user_PW
		Session("username")= rs("username")
		Session("fullname")= rs("FirstName") & " " & rs("LastName")
		Session("master") = rs("ismaster")
		Session("group_id") = rs("group_id")
		Session("areacode") = rs("areacode")

	else
		session.Abandon()
		response.Redirect("/_mib/login/")
	End If
	
	
	Set rs=Nothing
	
'	Conn.Close
'	Set Conn=Nothing
End Function





%>