<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" LCID="3084"%>
<%Response.Expires=-1441%>
<%
'Any Users
PageLevel=0
%>
<!--#include virtual="/intranet/includes/check.asp" -->
<!--#include virtual="/intranet/includes/fck/fckeditor.asp" -->
<%

on error resume next

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("cstring")

Set RequestForm = Server.CreateObject("ABCUpload4.XForm")
RequestForm.Overwrite = True
RequestForm.AbsolutePath = False
RequestForm.MaxUploadSize = 999999999

ID = Request.QueryString("ID")
IF ID="" then ID=RequestForm("emploi_id")

' New Record
If (RequestForm<>"" And ID="") OR RequestForm("_clone")<>"" then

	Conn.Execute("INSERT offresEmplois (emploi_id) VALUES (NULL);")
	Set rs = Conn.Execute("SELECT LAST_INSERT_ID() AS ID;")
	ID = rs("ID")
	Set rs=Nothing

End IF

' Delete Doc
IF request.querystring("action")="del" then
	Conn.Execute("Delete from offresEmplois Where emploi_id=" & ID & ";")
	Response.Redirect("default.asp")
End IF


' Modify Doc
If RequestForm<>"" then
	'On Error Resume next	
	strSQL= "Update offresEmplois SET "
	Separator=""
	 For each item in RequestForm
		If item <> "emploi_id" and Left(item,1)<>"_" then
			If InStr(item,"date")<>0 then
				If RequestForm(item)<>"" then
				 	strSQl = strSQL & Separator & item & "='" & MySQLDate(RequestForm(item)) & "'"
				Else
					strSQl = strSQL & Separator & item & "=NULL"
				End IF
			Else
				If Left(item,3)="int" then
					IF IsNumeric(RequestForm(item)) Then
						strSQl = strSQL & Separator & item & "=" & RequestForm(item)
					Else
						strSQl = strSQL & Separator & item & "=0"
					End IF
				Else
					strSQl = strSQL & Separator & item & "='" & Replace(Replace(RequestForm(item),"""","&quot;"),"'","''") & "'"
				End If
			End IF
			Separator = ", "
		End If
	 Next
	 
	 strSQL = strSQL & " Where emploi_id = " & ID & ";"

	 Conn.Execute(strSQL)

	 If RequestForm("_clone")<>"" then
	 	Conn.Execute("UPDATE offresEmplois SET referenceInterne = CONCAT(referenceInterne,' COPIE') Where emploi_id = " & ID & ";")
	 End IF

	 if err.number <> 0 then
	 	response.write strsql & err.description
		response.end
	 End If
	 
	 Response.Redirect("details.asp?ID="&ID)

End If

HaveSaveRights = Session("Level")

If ID <> "" then

	strSQL = "Select * from offresEmplois WHERE emploi_id=" & ID & ";"
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open strSQL,Conn
		
	For each item in rs.Fields
		Execute(item.Name & "=rs(""" & item.Name & """)")
	Next
	rs.Close

	'retailprice = retailprice + specialprice1+specialprice2+specialprice3+specialprice4+specialprice5

	if dealer_id <> session("dealer_id") AND session("dealer_id")<>0 then HaveSaveRights = False
	if Session("Level")=0 then HaveSaveRights = False
	
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml">
<!-- InstanceBegin template="/Templates/backend.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Intranet - <%=dealers(session("dealer_id"),1)%></title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" -->

<link rel="stylesheet" type="text/css" media="all" href="/intranet/includes/jscalendar-1.0/skins/aqua/theme.css" title="Aqua" />

<script type="text/javascript" src="/intranet/includes/jscalendar-1.0/calendar.js"></script>
<script type="text/javascript" src="/intranet/includes/jscalendar-1.0/lang/calendar-fr.js"></script>
<script type="text/javascript" src="/intranet/includes/jscalendar-1.0/calendar-setup.js"></script>


<script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
<script src="//ajax.googleapis.com/ajax/libs/jqueryui/1.10.3/jquery-ui.min.js"></script>
<link rel="stylesheet" type="text/css" href="../includes/jquery/lightbox/lightbox.css" />
<script type="text/javascript" src="../includes/jquery/lightbox/lightbox.min.js"></script>
<style>
 .lb-nav {
   display: none !important;
 }
</style>
<script language="JavaScript" type="text/javascript">
<!--
function Delete(ID) {
	if (confirm("ATTENTION!\nCette action supprimera définitivement cette fiche.\nVoulez-vous continuer?")){
		string="details.asp?action=del&ID="+ID;
		document.location.href=string;
	}
}
// -->


function DeletePDF(ID,type) {
	if (confirm("Supprimer le document PDF?")){
		string="details.asp?action=delpdf&type="+type+"&ID="+ID;
		document.location.href=string;
	}
}

function DeletePic(id,picid) {
	if (confirm("Voulez-vous supprimer la photo?")){
		window.location = "details.asp?id="+id+"&picid="+picid;
	}
}

function ParseNumber(num) {

	if (num == '') { return 0; }
	num = num.replace(',','.');
	return parseFloat(num);
}

function UpdatePrice() {

	var specialprice1 = ParseNumber($("#specialprice1").val());
	var specialprice2 = ParseNumber($("#specialprice2").val());
	var specialprice3 = ParseNumber($("#specialprice3").val());
	var specialprice4 = ParseNumber($("#specialprice4").val());
	var specialprice5 = ParseNumber($("#specialprice5").val());
	var total = specialprice1 + specialprice2 + specialprice3 + specialprice4 + specialprice5;

    $('#sum_special').html('<b>'+ total +' $</b>');
    $('.inventory_equip_total').html('Total: <b>'+ total +' $</b>');


<%if HaveSaveRights then%>	
	
	var cost = ParseNumber($("#dealernet").val());
	var profit = ParseNumber($("#profit").val());
	
	if ($("#profit_type").val()==1 && profit > 0) {
		profit =  (profit/100)*cost ; // %
	} 

	//alert (profit);
	var retail = Math.round(cost + profit);

    $('#sum_special_total').html('<b>'+ (total + retail) +' $</b>');
	
<%end if%>
}


$(function(){
	$("#addpic").click(function(){
		$("#browsetd").prepend('<input type="file" name="_img_file"/><br>');
	});
	
	$("#allpics").sortable();


	$("#reserve").click(function(){
		var msg;
		if($("#resclient").val().length < 2 || $("#resuserid").val()==0 || $("#_resdatetime").val()=="") {
		alert("Veuillez choisir un représentant, une date, et saisir un nom de client");
		return false;		
		} else {
			$("#frm").submit();
		}
	});

	$(".special_price_input,#profit,#dealernet").keyup(function(){
		UpdatePrice();
	});

	$("#profit_type").change(function(){
		UpdatePrice();
	});
	
	// Force update on inital load.
	UpdatePrice();

});




</script>
<SCRIPT Language="Javascript">
<!--

function printit(){  
if (window.print) {
    window.print() ;  
} else {
    var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
    WebBrowser1.ExecWB(6, 2);//Use a 1 vs. a 2 for a prompting dialog box
	WebBrowser1.outerHTML = "";  
}
}

//printit();
//-->
</script>


<script language="javascript">
<!--

var nn4 = (document.layers) ? true : false
var ie = (document.all) ? true : false
var dom = (document.getElementById && !document.all) ? true : false

function browser(id){
  if(nn4) {
  path = document.layers[id]
  }
  else if(ie) {
  path = document.all[id]
  } 
  else {
  path = document.getElementById(id)
  }
return path  //return the path to the css layer depending on which browser is looking at the page
}	

function findPosX(obj)
{
	var curleft = 0;
	if (obj.offsetParent)
	{
		while (obj.offsetParent)
		{
			curleft += obj.offsetLeft
			obj = obj.offsetParent;
		}
	}
	else if (obj.x)
		curleft += obj.x;
	return curleft;
}

function findPosY(obj)
{
	var curtop = 0;
	if (obj.offsetParent)
	{
		while (obj.offsetParent)
		{
			curtop += obj.offsetTop
			obj = obj.offsetParent;
		}
	}
	else if (obj.y)
		curtop += obj.y;
	return curtop;
}


function ShowUpload(element,div) {
	if (browser(div).style.display=="") {
		browser(div).style.display="none";
		}
	else {
		browser(div).style.display="";
		browser(div).style.left=findPosX(element)+ 17 + 'px';
		browser(div).style.top=findPosY(element) + 'px';
	}
}

//-->
</script>
<style type="text/css">
<!--
a.thickbox img { max-width:90px; max-height:90px}
a.thickbox:hover img { max-width:none; max-height:none;}

-->
</style>
<!-- InstanceEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- InstanceParam name="id" type="text" value="inventory" -->
</head>
<body id="emplois">
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
	</ul>
	<ul id="second_nav">
		<li><em><%="<strong>" & Session("FullName") & "</strong> (" & UserLevel(session("level")) & ")"%></em></li>
		<li><a href="/intranet/login/password.asp">Mot de passe</a></li>
		<li><a href="/intranet/login/logout.asp">Quitter</a></li>
	</ul>
</div>
<div id="content">
<!-- InstanceBeginEditable name="content" -->
<h1>Emploi <%=referenceInterne%></h1>
<form action="details.asp" method="post" name="frm" id="frm" enctype="multipart/form-data">
    <input type="hidden" name="emploi_id" value="<%=emploi_id%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="2" >
  <tr> 
    <th valign="top" align="left" nowrap="nowrap" colspan="4" class="titlebar"><a href="."><img src="/intranet/images/ico_fup.gif" width="16" height="16" border="0" alt="Back to the list" /></a></th>
  </tr>
  <tr> 
    <td colspan="4" align="left">&nbsp;</td>
    </tr>
  <tr>
    <td height="100%" class="cell_label">Titre :</td>
    <td width="50%" class="cell_content">
        <input class="" maxlength="100" name="titre" type="text" id="titre" value="<%=titre%>" size="50"/></td>
    <td class="cell_label">Fonctions :</td>
    <td width="50%" class="cell_content">
		<textarea class="textarea" maxlength="500" name="fonctions" cols="50" rows="3" id="fonctions"><%=fonctions%></textarea class="textarea" ><br/>
	</td>
  </tr>
    
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Référence interne :</td>
      <td class="cell_content"><input maxlength="50" size="50" name="referenceInterne" type="text" id="referenceInterne" value="<%=referenceInterne%>" maxlength="20"/>      </td>
      <td height="100%" align="left" nowrap="nowrap" class="cell_label" >Date de début : </td>
      <td align="left" class="cell_content" ><input maxlength="20" type="text" name="dateDebut" id="dateDebut" value="<%=MakeLocalDate(dateDebut)%>" />
        <a href="#" id="triggerint"><img src="/intranet/includes/jscalendar-1.0/img.gif" alt="Calendrier" width="20" height="16" border="0" align="absmiddle" /></a>
		<script type="text/javascript">
		  Calendar.setup(
			{
			  inputField  : "dateDebut",   // ID of the input field
			  ifFormat    : "%d/%m/%Y",    // the date format
			  align		  : "Bl",
			  button      : "triggerint"   // ID of the button
			}
		   );
		</script>
		</td>
    </tr>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Succursale :</td>
      <td class="cell_content">
	  <select name="succursale" id="succursale" style="width:200px;">
		  <option <%if succursale="-" then response.write "selected"%>>-</option>
		  <option <%if succursale="Anjou" then response.write "selected"%>>Anjou</option>
          <option <%if succursale="Boucherville" then response.write "selected"%>>Boucherville</option>
		  <option <%if succursale="Drummondville" then response.write "selected"%>>Drummondville</option>
		  <option <%if succursale="Joliette" then response.write "Joliette"%>>Joliette</option>
		  <option <%if succursale="Québec" then response.write "Québec"%>>Québec</option>
		  <option <%if succursale="Rivière-du-Loup" then response.write "selected"%>>Rivière-du-Loup</option>
		  <option <%if succursale="Saint-Georges" then response.write "selected"%>>Saint-Georges</option>
		  <option <%if succursale="Saint-Hyacinthe" then response.write "selected"%>>Saint-Hyacinthe</option>
		  <option <%if succursale="Shawinigan" then response.write "selected"%>>Shawinigan</option>
		  <option <%if succursale="Thetford Mines" then response.write "selected"%>>Thetford Mines</option>
		  <option <%if succursale="Trois-Rivières" then response.write "selected"%>>Trois-Rivières</option>
		  <option <%if succursale="Victoriaville" then response.write "selected"%>>Victoriaville</option>
        </select>
		</td>
	  <td height="100%" nowrap="nowrap" class="cell_label"></td>
	  <td height="100%" nowrap="nowrap" class="cell_label"></td>
    </tr>
	
	<tr> 
	  <td class="cell_label" colspan="4" align="left"><font color="black"><b>Exigences et conditions</b></font></td>
    </tr>
    <tr>
      <td class="cell_label">Niveau d'études:</td>
      <td class="cell_content">
      	<textarea class="textarea" name="niveauEtudes" maxlength="200" cols="50" rows="2" id="niveauEtudes"><%=niveauEtudes%></textarea class="textarea" ></td>
	  <td class="cell_label">Années d'expérience:</td>
      <td class="cell_content">
      	<textarea class="textarea" name="anneesExperience" maxlength="200" cols="50" rows="2" id="anneesExperience"><%=anneesExperience%></textarea class="textarea" ></td>
    </tr>
	
	<tr>
      <td class="cell_label">Compétences:</td>
      <td class="cell_content">
      	<textarea class="textarea" name="descCompetences" maxlength="1000" cols="50" rows="2" id="descCompetences"><%=descCompetences%></textarea class="textarea" ></td>
	  <td class="cell_label">Langues:</td>
      <td class="cell_content">
      	<textarea class="textarea" name="langues" maxlength="200" cols="50" rows="2" id="langues"><%=langues%></textarea class="textarea" ></td>
    </tr>
	
    <tr>
      <td class="cell_label">Statut d'emploi:</td>
      <td class="cell_content">
      	<textarea class="textarea" name="statutEmploi" maxlength="200" cols="50" rows="2" id="statutEmploi"><%=statutEmploi%></textarea class="textarea" ></td>
	  <td class="cell_label">Salaire:</td>
      <td class="cell_content">
      	<textarea class="textarea" name="salaire" maxlength="50" cols="50" rows="2" id="salaire"><%=salaire%></textarea class="textarea" ></td>
    </tr>
	
	<tr>
      <td class="cell_label">Heures/semaine:</td>
      <td class="cell_content">
      	<textarea class="textarea" name="heuresSemaine" maxlength="200" cols="50" rows="2" id="heuresSemaine"><%=heuresSemaine%></textarea class="textarea" ></td>
	  <td class="cell_label">Autres:</td>
      <td class="cell_content">
      	<textarea class="textarea" name="autres" maxlength="200" cols="50" rows="2" id="autres"><%=autres%></textarea class="textarea" ></td>
    </tr>
	<tr> 
		<td class="cell_label" colspan="4" align="left"><font color="black"><b>Information de contact:</b></font></td>
    </tr>
	
   <tr>
	<td height="100%" nowrap="nowrap" class="cell_label">Nom:</td>
      <td class="cell_content"><input maxlength="100" size="50" name="contactNom" type="text" id="contactNom" value="<%=contactNom%>" maxlength="20"/>      </td>
      <td height="100%" align="left" nowrap="nowrap" class="cell_label" >Titre: </td>
      <td align="left" class="cell_content" ><input maxlength="100" size="50" type="text" name="contactTitre" id="contactTitre" value="<%=contactTitre%>" /></td>
	<tr> 
	  <td height="100%" nowrap="nowrap" class="cell_label">Telephone:</td>
      <td class="cell_content"><input maxlength="100" size="50" name="contactTel" type="text" id="contactTel" value="<%=contactTel%>" maxlength="20"/>      </td>
      <td height="100%" align="left" nowrap="nowrap" class="cell_label" >Courriel: </td>
      <td align="left" class="cell_content" ><input maxlength="100" size="50" type="text" name="contactCourriel" id="contactCourriel" value="<%=contactCourriel%>" /></td>
    </tr>
	<tr> 
      <td height="100%" colspan="4" align="left">&nbsp;</td>
      </tr>

<%
If HaveSaveRights then
	buttonstate = ""
Else
	buttonstate = "disabled"
End If
%>
	<tr>
	<td height="100%" nowrap="nowrap" class="cell_label">Afficher sur le site web :</td>
      <td align="left" class="cell_content" >
		<input <%if displayOnWeb=1 then response.write "checked"%> name="displayOnWeb" type="radio" value="1" id="soui" />
        <label for="soui">oui</label>
        <input <%if displayOnWeb=0 then response.write "checked"%> name="displayOnWeb" type="radio" value="0" id="snon" />
        <label for="snon">non</label>
	  </td>
	  <td height="100%" nowrap="nowrap" class="cell_label">Poste comblé :</td>
	  <td align="left" class="cell_content" >
		<input <%if filled=1 then response.write "checked"%> name="filled" type="radio" value="1" id="soui" />
        <label for="soui">oui</label>
        <input <%if filled=0 then response.write "checked"%> name="filled" type="radio" value="0" id="snon" />
        <label for="snon">non</label>
	  </td>
	</tr>
    <tr>
      <td height="100%" align="left" nowrap="nowrap" class="cell_label" >Actions :</td>
      <td colspan="3" align="left" class="cell_content" > <input type="submit" value="Sauvegarder" class="button" <%=buttonstate%> />
        <% If ID <> "" then%> <input name="_clone" type="submit" class="button" id="_clone" value="Dupliquer" <%=buttonstate%> />
        <input onClick="Delete(<%=ID%>)" type="button" value="Supprimer" name="del" class="button" <%=buttonstate%> />
        <%End If%> <input onClick="printit()" type="button" value="Imprimer" name="_print" class="button"/></td>
      </tr>
</table>  
</form>
<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>
<%

if err.number <> 0 then
	response.write "An error occured:" & err.description & "<br>"
	response.end 
End If
		
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>