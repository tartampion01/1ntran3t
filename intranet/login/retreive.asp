<!-- #include virtual="/intranet/includes/functions.asp" -->
<%

UserName = trim(Request.Form("username"))
Email = trim(Request.Form("email"))

If UserName & email <>"" then

	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("cstring")


	Set rs=Conn.Execute("SELECT password,CONCAT(firstname,' ',lastname) as fullname,username,email from intranet_users WHERE username='" & UserName & "' or  email='" & email & "' AND email <> '';")
	If NOT rs.EOF then
	
			message = ReadFile("password.html")
			message = replace(message,"[fullname]",rs("fullname"))
			message = replace(message,"[username]",rs("username"))
			message = replace(message,"[password]",decode(rs("password")))

			
			if SendEmail(rs("email"),"Vos code d'accès au l'intranet",message,"sallard@servicesinfo.info","ReseauDynamique",priority)then
				msg="<span class=ok>Vos acces vous ont été envoyé par courriel à l'adress: " & rs("email") & "</span>"
			else
				ErrMsg="Une erreur est survenue lors de l'envoi à '"& rs("email") &"'. Veuillez communiquer avec l'administrateur"
			end if
			
	Else
		ErrMsg="Désolé, nous n'avons pas ce nom d'usager ou courriel dans notre système.<br/>Veuillez communiquer avec un administrateur."
	End If

	Set rs=Nothing
	Conn.Close
	Set Conn=Nothing
Else
	msgerr=""
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<META NAME="ROBOTS" CONTENT="NOINDEX, NOFOLLOW">
<title>Intranet - R&eacute;cup&eacute;ration du mot de passe</title>
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
</head>

<body>
	<div id="bg_image_accueil">
		<div id="image_accueil"></div>
	</div>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr> 
    <td align="center">
    
    
    		<%=msg%>
     		<span class="err"><%=ErrMsg%></span>
            <form method="POST" action="retreive.asp">
              <table width="260" border="0" cellpadding="0" cellspacing="5">
          <tr>
            <td align="left">
              <h2 align="center">R&eacute;cup&eacute;ration de vos acc&egrave;s</h2>
            </td>
          </tr>
          <tr>
            <td align="left">Veuillez entrer votre nom d'usager ou votre courriel pour recevoir vos codes d'acc&egrave;s par courriel.</td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0" id="login">
              <tr>
                <th>&nbsp;</th>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td align="center" nowrap="nowrap"><strong>Utilisateur :</strong></td>
                      <td align="right"><input type="text" name="username" value="<%=username%>" style="width:150px" />                      </td>
                    </tr>
                    <tr>
                      <td align="center" nowrap="nowrap">ou</td>
                      <td align="center" nowrap="nowrap">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="center" nowrap="nowrap"><strong>Courriel :</strong></td>
                      <td align="right"><input name="email" type="text" id="email" value="<%=email%>" style="width:150px" /></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td><table border="0" width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                      <td align="right"><input name="submit" type="submit" class="button" value="R&eacute;cup&eacute;rer!" />                      </td>
                    </tr>
                </table></td>
              </tr>
            </table>            </td>
          </tr>
          <tr>
            <td><a href=".">Je m'en rappelle!</a></td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table> 

</body>
</html>
