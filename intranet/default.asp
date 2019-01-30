<%Response.Expires=-1441%>
<%
'Any Users
PageLevel=0
%>
<!-- #include virtual="/intranet/includes/check.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Intranet - <%=dealers(session("dealer_id"),1)%></title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" -->
 

<script language="JavaScript" type="text/javascript">
<!--
function Delete(ID) {
	if (confirm("WARNING! This operation will remove this user definitively.")){
		string="details.asp?action=del&ID="+ID;
		document.location.href=string;
	}
}
// -->
</script>

<link rel="stylesheet" href="/intranet/includes/styles_app.css" type="text/css" />
<!-- InstanceEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- InstanceParam name="id" type="text" value="home" -->
</head>
<body id="home">
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
     
      <div id="image_accueil"></div>
      <h3 align="center">Bienvenue 
        <br />
        <%=Session("FullName")%>    !<br />
        <br />  
        <br />
      Vous &ecirc;tes maintenant branch&eacute; en mode s&eacute;curis&eacute;!    </h3>
      <p align="center">&nbsp;</p>
    <p align="center"><span id="siteseal"><script type="text/javascript" src="https://seal.godaddy.com/getSeal?sealID=c03bqggyuYwXeL19KAtAZxe2oXJVehFDFKI79ZtzOdbciJ0zZ5wf7wdek"></script></span></p>
	<div id="bg_image_accueil">
		
	</div>
	
<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>