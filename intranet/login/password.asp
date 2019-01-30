<!-- #include virtual="/intranet/includes/check.asp" -->
<!-- #include virtual="/intranet/includes/functions.asp" -->
<%
' Modify Password
If Request.Form("modUser")<>"" then

	If 	Request.Form("password")=Request.Form("password1") and Request.Form("password")<>"" then
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open Application("cstring")
			strSQL = strSQL & "Update intranet_users SET password='" & encode(Replace(Request.Form("password"),"'","''")) & "'"
			strSQL = strSQL & " WHERE ID=" & Session("UserID") & ";"
			Conn.Execute(strSQL)
			Conn.Close
			Set Conn=Nothing
			errMsg = errMsg & "<span class=ok>La mot de passe a été modifié avec succès.</span>"
	Else
			errMsg = errMsg & "<span class=err>La vérification du mot de passe a échouée.</span>"
	End If
End If

%>
<!DOCTYPE html>
<!-- InstanceBegin template="/Templates/backend.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Intranet - <%=dealers(session("dealer_id"),1)%></title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" -->

<link rel="stylesheet" type="text/css" media="screen" href="/intranet/includes/jquery.validate.password.css" />

<script type="text/javascript" src="/intranet/includes/jquery.js"></script>
<script type="text/javascript" src="/intranet/includes/jquery.validate.js"></script>
<script type="text/javascript" src="/intranet/includes/jquery.validate.password.js"></script>

<script type="text/javascript">
$(document).ready(function() {

	// validate signup form on keyup and submit
	var validator = $("#register").validate({
		rules: {
			password1: {
				required: true,
				equalTo: "#password"
			}
		},
		messages: {
			password1: {
				required: "Répétez le mot de passe",
				minlength: jQuery.format("Entrez au minimum {0} caractères"),
				equalTo: "Entrez le même mot de passe que plus haut"
			}
		},
		// the errorPlacement has to take the table layout into account
		errorPlacement: function(error, element) {
			error.appendTo( element.parent() );
		},
		// specifying a submitHandler prevents the default submit, good for the demo
//		submitHandler: function() {
//			alert("submitted!");
//		},
		// set this class to error-labels to indicate valid fields
		success: function(label) {
			// set &nbsp; as text for IE
			label.html("&nbsp;").addClass("checked");
		}
	});
	
	
});

	

</script>

<style>
#password {float:left;} 
.password-meter {float:left; margin-left:10px;}
label.error { color:red; padding-left:10px; line-height:20px; }
</style>


<!-- InstanceEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- InstanceParam name="id" type="text" value="" -->
</head>
<body id="">
<div id="container">
<div id="header">
	<ul id="nav">
		<!-- UTILISATEUR -->
		<%IF Session("Level")=0 then%>
		<li id="invent_tab"><a href="/intranet/inventory/">Inventaire</a></li>
		<%End If%>
		<!-- ADMINISTRATEUR -->
		<%IF Session("Level")=1 then%>
		<li id="invent_tab"><a href="/intranet/inventory/">Inventaire</a></li>
		<li id="users_tab"><a href="/intranet/users/">Utilisateurs</a></li>
		<li id="emplois_tab"><a href="/intranet/emplois/">Emplois</a></li>
		<%End If%>
		<!-- RH -->
		<%IF Session("Level")=2 then%>
		<li id="emplois_tab"><a href="/intranet/emplois/">Emplois</a></li>
		<%End If%>
		<!-- ReadOnly -->
		<%IF Session("Level")=3 then%>
		<li id="invent_tab"><a href="/intranet/inventory/">Inventaire</a></li>
		<%End If%>
	</ul>
	<ul id="second_nav">
		<li><em><%="<strong>" & Session("FullName") & "</strong> (" & UserLevel(session("level")) & ")"%></em></li>
		<li><a href="/intranet/login/password.asp">Mot de passe</a></li>
		<li><a href="/intranet/login/logout.asp">Quitter</a></li>
	</ul>
</div>
<div id="content">
<!-- InstanceBeginEditable name="content" --> 
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr> 
    <td valign="top"><h1>Modification du mot de passe</h1></td>
  </tr>
  <%If errMsg <> "" then%>
  <tr> 
    <td valign="top"><%=ErrMsg%></td>
  </tr>
  <%End IF%>
</table>

<form method="post" action="password.asp" id="register">
<table border="0" width="100%" cellspacing="0" cellpadding="3">
    <tr>
      <th colspan="2" align="left" valign="top" nowrap="nowrap"><a href="/intranet/"><img src="/intranet/images/ico_fup.gif" width="16" height="16" border="0" alt="Back to the list" /></a></th>
      </tr>
    <tr>
      <th valign="middle" align="left" nowrap="nowrap">Nom:</th>
      <td valign="middle" bgcolor="#eeeeee" style="border-bottom:1px solid #CCC;"><%=Session("FullName")%></td>
    </tr>
    <tr> 
      <th valign="middle" align="left" nowrap="nowrap"><b>Utilisateur:</b></th>
      <td valign="middle" width="100%" bgcolor="#eeeeee" style="border-bottom:1px solid #CCC;"><%=Session("username")%></td>
    </tr>
    <tr>
      <th valign="top" align="left" nowrap="nowrap">&nbsp;</th>
      <td valign="top" bgcolor="#eeeeee">Votre mot de pass doit contenir au minimum 6 caract&egrave;res, une lettre minuscule, une lettre majuscule et un chiffre.<br>
        Un caract&egrave;re sp&eacute;cial est aussi recommand&eacute;.</td>
    </tr>
    <tr> 
      <th valign="top" align="left" nowrap="nowrap"><b>Mot de passe:</b></th>
      <td valign="top" width="100%" bgcolor="#eeeeee"> 
        <input type="password" class="password" name="password" id="password" size="15" /> 	
              				<div class="password-meter">
	  					<div class="password-meter-message">&nbsp;</div>
	  					<div class="password-meter-bg">
		  					<div class="password-meter-bar"></div>
	  					</div>
	  				</div>        </td>
    </tr>
    <tr>
      <th valign="top" align="left" nowrap="nowrap">R&eacute;p&eacute;tez</th>
      <td valign="top" bgcolor="#eeeeee"><input type="password" name="password1" id="password1" size="15" /></td>
    </tr>
    <tr> 
      <th valign="top" align="left" nowrap="nowrap">&nbsp;</th>
      <td valign="top" width="100%" bgcolor="#eeeeee"> 
        <input name="modUser" type="submit" class="button" value="Modifier" />      </td>
    </tr>
</table>
  </form>

<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>