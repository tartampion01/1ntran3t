<!-- #include virtual="/intranet/includes/functions.asp" -->
<%


'if request.QueryString("k")<>"" then
'
'	Set Conn = Server.CreateObject("ADODB.Connection")
'	Conn.Open Application("cstring")
'
'	Set rs=Conn.Execute("SELECT id,password,username from intranet_users;")
'	
'	Do while not rs.EOF
'		response.write rs("username")
'		Conn.Execute("update intranet_users set password='"& encode(rs("password")) &"' WHERE id='"& rs("ID") &"';")
'		rs.MoveNext
'	loop
'	
'end if

UserName = Request.Form("username")
Password = Request.Form("password")
URL = Request.QueryString("url")

IF URL="" then URL = "/intranet/"

If UserName <>"" then

	'Revoke Admin rights
	Session("Level")=""
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("cstring")

	Set rs=Conn.Execute("SELECT * from intranet_users WHERE username='" & UserName & "';")

	If NOT rs.EOF then
		Session("userID")= rs("ID")
		Session("username")= rs("username")
		Session("fullname")= Server.HTMLEncode(rs("FirstName") & " " & rs("LastName"))	
		If Password = rs("password") then
			If Request.Form("cookie")="ON" then
				Response.Cookies("userID") = Session("userID")
				Response.Cookies("userID").Expires = Date + 30
				Response.Cookies("username") = Session("username")
				Response.Cookies("username").Expires = Date + 30
				Response.Cookies("fullname") = Session("fullname")
				Response.Cookies("fullname").Expires = Date + 30
			Else
				Response.Cookies("userID").Expires = Date - 1000
				Response.Cookies("username").Expires = Date - 1000
				Response.Cookies("fullname").Expires = Date - 1000
			End If
			Response.Redirect(URL)
		Else
			Response.Redirect("/intranet/login/?err=1&URL=" & Server.URLEncode(URL))
		End If
	Else
		Response.Redirect("/intranet/login/?err=2&URL=" & Server.URLEncode(URL))
	End If	

	Set rs=Nothing
	Conn.Close
	Set Conn=Nothing
End If

If Request.Cookies("userID")<>"" then
	cookieCheck = "checked"
End If

If Request.QueryString("err")=3 then
	ErrMsg = "Vous ne pouvez accéder à cette section."
End If
If Request.QueryString("err")=2 then
	ErrMsg = "Nom d'utilisateur invalide"
End If
If Request.QueryString("err")=1 then
	ErrMsg = "Mot de passe invalide"
End If


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<META NAME="ROBOTS" CONTENT="NOINDEX, NOFOLLOW">
<title>Intranet - Acc&egrave;s utilisateur</title>
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
</head>

<body>
	<div id="bg_image_accueil">
		<div id="image_accueil"></div>
	</div>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr> 
    <td align="center"> 
            <form method="POST" action="default_nms.asp?URL=<%=Server.URLEncode(URL)%>">
              <table width="250" border="0" cellpadding="0" cellspacing="5">
          <tr>
            <td align="left">
              <h2 align="center">Bienvenue!</h2>
            <span class="err"><%=ErrMsg%></span></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0" id="login">
              <tr>
                <th><strong>Acc&egrave;s utilisateur </strong></th>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td align="left"><strong>Utilisateur:</strong></td>
                      <td align="right"><input type="text" name="username" value="<%=session("username")%>" style="width:150px" />                      </td>
                    </tr>
                    <tr>
                      <td align="left" nowrap="nowrap"><strong>Mot de passe :</strong></td>
                      <td align="right"><input type="password" name="password" style="width:150px" /></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td><table border="0" width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><input type="checkbox" name="cookie" value="ON" />                      </td>
                      <td align="left">M&eacute;moriser </td>
                      <td align="right"><input name="submit" type="submit" class="button" value="Login" />                      </td>
                    </tr>
                </table></td>
              </tr>
            </table>            </td>
          </tr>
          <tr>
            <td><a href="retreive.asp">Mot de passe oublié?</a></td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table> 

</body>
</html>
