<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" LCID="3084"%>
<%Response.Expires=-1441%>
<%
'Any Users
PageLevel=0
%>
<!--#include virtual="/intranet/includes/check.asp" -->
<!--#include virtual="/intranet/includes/fck/fckeditor.asp" -->
<% 

	Function ConvertFileToBase64( file )

		' This script reads jpg picture converts it to base64
		' code using encoding abilities of MSXml2.DOMDocument object and saves

		Const fsDoOverwrite     = true  ' Overwrite file with base64 code
		Const fsAsASCII         = false ' Create base64 code file as ASCII file
		Const adTypeBinary      = 1     ' Binary file is encoded

		' Variables for writing base64 code to file
		Dim objFSO
		Dim objFileOut

		' Variables for encoding
		Dim objXML
		Dim objDocElem

		' Variable for reading binary picture
		Dim objStream

		' Open data stream from picture
		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Type = adTypeBinary
		objStream.Open()

		objStream.LoadFromFile(file)

		' Create XML Document object and root node
		' that will contain the data
		Set objXML = CreateObject("MSXml2.DOMDocument")
		Set objDocElem = objXML.createElement("Base64Data")
		objDocElem.dataType = "bin.base64"

		' Set binary value
		objDocElem.nodeTypedValue = objStream.Read()

		ConvertFileToBase64 = objDocElem.text

		' Clean all
		Set objFSO = Nothing
		Set objFileOut = Nothing
		Set objXML = Nothing
		Set objDocElem = Nothing
		Set objStream = Nothing
	End Function
	
	Function ResizeImage(srcFileName,destFileName,MaxWidth,MaxHeight,Crop)
		'Reduce and compress large image
		
		on error resume next
		
		Dim Image
				
		srcFileName = Server.MapPath(srcFileName)
		
		Set Image = Server.CreateObject("GflAx.GflAx")
		
		Image.LoadBitmap srcFileName
		Image.SaveJPEGQuality = 60
		
		'ignore  X or Y modes
		IF MaxWidth=0 then MaxWidth=Image.Width
		IF MaxHeight=0 then MaxHeight=Image.Height 
		
'		if crop then
'			
'			
'							
'			Response.write  (Image.Width - MaxWidth)/2 & "," & (Image.Height - MaxHeight)/2 & "," & MaxWidth & "," & MaxHeight
'			response.end
'			Image.Crop (Image.Width - MaxWidth)/2,(Image.Height - MaxHeight)/2,MaxWidth,MaxHeight
'			
'		else
	
			If Image.Width > MaxWidth then
				Image.Resize MaxWidth,Int(Image.Height*MaxWidth/Image.Width) 
			End IF
			If Image.Height > MaxHeight then
				Image.Resize Int(Image.Width*MaxHeight/Image.Height),MaxHeight
			End If			
			
		'end if

		Image.SaveBitmap Server.MapPath(destFileName)
		
		if err.number <> 0 then
			response.write "An error occured creating Image:" & err.description & "<br>"
			response.end 
		End If
	
		Set Image = Nothing
		
		'response.end 
		
	End Function
function FSODelete(vpath) 

	Dim FSO	, fn

	On Error Resume Next

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	if FSO.FileExists(Server.MapPath(vPath)) then
		FSO.DeleteFile Server.MapPath(vPath),True
	End IF
	
	Set FSO = Nothing
		
	if err<>0 then
		response.write vpath
		response.write err.description
		response.end
	End IF

end function

function FSOExists(vpath) 
	Dim FSO	, fn
	FSOExists = false
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	if FSO.FileExists(Server.MapPath(vPath)) then
		FSOExists = true
	End IF
	Set FSO = Nothing
end function



Set Conn = Server.CreateObject("ADODB.Connection")
'Commenter ligne suivante avec apostrophe si problème avec connection DB
Conn.Mode = adModeReadWrite
Conn.Open Application("cstring")

Set RequestForm = Server.CreateObject("ABCUpload4.XForm")
RequestForm.Overwrite = True
RequestForm.AbsolutePath = False
RequestForm.MaxUploadSize = 999999999

ID = Request.QueryString("ID")
IF ID="" then ID=RequestForm("ID")

' New Record
If (RequestForm<>"" And ID="") OR RequestForm("_clone")<>"" then

	Conn.Execute("INSERT inventory (ID) VALUES (NULL);")
	Set rs = Conn.Execute("SELECT LAST_INSERT_ID() AS ID;")
	ID = rs("ID")
	Set rs=Nothing

End IF

' Delete Doc
IF request.querystring("action")="del" then
	set rs = conn.execute("SELECT name,id FROM inv_pictures WHERE product_id='"&id&"'")
	do while not rs.EOF
		FSODelete("/intranet/pics/img"&rs("id")&".jpg")
		FSODelete("/intranet/pics/small/img"&rs("id")&".jpg")
		Conn.execute("Delete from inv_pictures WHERE id='"& rs("id") & "';")
		rs.movenext
	loop
	Call DeleteFile("/intranet/pdf/prop"& ID & ".pdf")
	Call DeleteFile("/intranet/pdf/inv"& ID & ".pdf")
	Conn.Execute("Delete from inventory Where ID=" & ID & ";")
	Response.Redirect("default.asp")
End IF


' Delete PDF
IF Request.QueryString("action")="delpdf" then
	Call DeleteFile("/intranet/pdf/"& request.QueryString("type") & ID & ".pdf")
	Response.Redirect("details.asp?ID="&ID)
End IF



'delete image
if request.querystring("picid") <> "" then
	Conn.execute("Delete from inv_pictures WHERE id="&request.querystring("picid"))
		FSODelete("/intranet/pics/img"&request.querystring("picid")&".jpg")
		FSODelete("/intranet/pics/small/img"&request.querystring("picid")&".jpg")
	set rs = nothing
	response.Redirect("details.asp?id=" & request.querystring("id"))
end if

' Modify Doc
If RequestForm<>"" then
	'On Error Resume next	
	strSQL= "Update inventory SET "
	Separator=""
	 For each item in RequestForm
		If item <> "ID" and Left(item,1)<>"_" then
			If InStr(item,"Date")<>0 then
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
	 
	'Active checkbox
	 'IF RequestForm("intStatus")="" then strSQL = strSQL & ",intStatus=0"
	 'IF RequestForm("intNew")="" then strSQL = strSQL & ",intNew=0"
	 'IF RequestForm("intHot")="" then strSQL = strSQL & ",intHot=0"
	 
	 ' Reserver
	 If IsDate(RequestForm("_resdatetime")) and RequestForm("_resclient")<>"" then
		If RequestForm("_resuserid") > 0 then
			strSQL = strSQL & ",resdatetime='" & MySQLDate(RequestForm("_resdatetime")) & " " & TimeValue(RequestForm("_resdatetime")) & "'"
			strSQL = strSQL & ",resuserid=" & RequestForm("_resuserid")
			strSQL = strSQL & ",resclient='" & Replace(Replace(RequestForm("_resclient"),"""","&quot;"),"'","''") & "'"
		end if	
	 End If
	 
	 'if session("level")=1 then strSQL = strSQL & ",resclient='" & Replace(Replace(RequestForm("_resclient"),"""","&quot;"),"'","''") & "'"
	 
	 'response.write strSQL
	 'response.end
	 
	 If RequestForm("_cancel") <> "" then
	  	strSQL = strSQL & ",resdatetime=Null"
		strSQL = strSQL & ",resuserid=0"
		strSQL = strSQL & ",resclient=''"	
	 End If 	 
	 
	 strSQL = strSQL & " Where ID = " & ID & ";"
	'Response.Write 
	

	' if err.number <> 0 then
	 '	response.write strsql & err.description
	'	response.end
	' End If

	 Conn.Execute(strSQL)


	 
	 If RequestForm("_clone")<>"" then
	 	Conn.Execute("UPDATE inventory SET stock = CONCAT(stock,' COPIE') Where ID = " & ID & ";")
		'Call CopyPictures(RequestForm("ID"),ID,"/product_images/")
	 End IF


	'set fs=Server.CreateObject("Scripting.FileSystemObject")
			
	Set theField = RequestForm("_prop")(1)
	If theField.FileExists Then theField.Save("/intranet/pdf/prop" & ID & ".pdf")
	
	Set theField = RequestForm("_inv")(1)
	If theField.FileExists Then theField.Save("/intranet/pdf/inv" & ID & ".pdf")

	Set theField = RequestForm("_special1")(1)
	If theField.FileExists Then theField.Save("/intranet/pdf/special-1-" & ID & ".pdf")
	Set theField = RequestForm("_special2")(1)
	If theField.FileExists Then theField.Save("/intranet/pdf/special-2-" & ID & ".pdf")
	Set theField = RequestForm("_special3")(1)
	If theField.FileExists Then theField.Save("/intranet/pdf/special-3-" & ID & ".pdf")
	Set theField = RequestForm("_special4")(1)
	If theField.FileExists Then theField.Save("/intranet/pdf/special-4-" & ID & ".pdf")
	Set theField = RequestForm("_special5")(1)
	If theField.FileExists Then theField.Save("/intranet/pdf/special-5-" & ID & ".pdf")

	'set fs=nothing

	'Calculate retail
	Conn.Execute("UPDATE inventory SET retailprice = ROUND(specialprice1+specialprice2+specialprice3+specialprice4+specialprice5+if(profit_type=0,dealernet+profit,dealernet+(profit/100*dealernet))) Where ID = " & ID & ";")
 

	'image order
	pic_order = 0
	for each pic_id in RequestForm("_picture_id")
		 Conn.Execute("Update inv_pictures SET intorder='" & pic_order & "' WHERE id='"& pic_id &"';")	
		 pic_order=pic_order + 1
	next	 
	 
	'image upload
	for each items in RequestForm("_img_file")
		Set theField = items
		If theField.FileExists and theField.ImageType <> 0 Then
		
			Conn.Execute("INSERT inv_pictures (ID,product_id) VALUES (NULL,'"&ID&"');")
			Set rs = Conn.Execute("SELECT LAST_INSERT_ID() AS ID;") 
			picID = rs("ID")
			picfilename =  "img" & picID & ".jpg"
			theField.Save ("/intranet/pics/tmp/" & picfilename)
			Call ResizeImage("/intranet/pics/tmp/"&picfilename,"/intranet/pics/"&picfilename,800,800,false)
			Call ResizeImage("/intranet/pics/tmp/"&picfilename,"/intranet/pics/small/"&picfilename,100,100,false)
			
			'To insert image as base64 into DB
			Dim b64
			b64 = "data:image/jpeg;base64," & ConvertFileToBase64(Server.MapPath("/intranet/pics/tmp/"&picfilename))
						
			Conn.Execute("UPDATE inv_pictures SET intorder='" & pic_order & "', name='"&picfilename&"', base64_picture='" & b64 & "' WHERE ID ='"&picID&"';")
			
			pic_order=pic_order + 1		
			FSODelete("/intranet/pics/tmp/"&picfilename) 
			
		end if
		set rs = nothing
	next


	 if err.number <> 0 then
	 	response.write strsql & err.description
		response.end
	 End If
	 
	 Response.Redirect("details.asp?ID="&ID)


End If

HaveSaveRights = Session("Level")

'HH MM combos
'For i = 0 to 23
'	hhCombo = hhCombo & "<option value=" & i & ">" & Right(formatNumber(i/100,2),2) & "</option>" & VbCrLf
'Next
'For i = 0 to 59
'	mmCombo = mmCombo & "<option value=" & i & ">" & Right(formatNumber(i/100,2),2) & "</option>" & VbCrLf
'Next

If ID <> "" then

	strSQL = "Select *, cast(ifnull(DATEDIFF(CURDATE(),invoicedate),0) as SIGNED) as daysin,cast(ifnull(DATEDIFF(CURDATE(),receivedate),0) as SIGNED) as daysrec from inventory WHERE ID=" & ID & ";"
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open strSQL,Conn
		
	For each item in rs.Fields
		Execute(item.Name & "=rs(""" & item.Name & """)")
	Next
	rs.Close

	'retailprice = retailprice + specialprice1+specialprice2+specialprice3+specialprice4+specialprice5

	if dealer_id <> session("dealer_id") AND session("dealer_id")<>0 then HaveSaveRights = False
	'Utilisateur
	if Session("Level")=0 then HaveSaveRights = False
	'Readonly
	if Session("Level")=3 then HaveSaveRights = False
	
	
Else
	hh = Hour(Now())
	mm = Minute(Now())

End If




hhCombo = Replace(hhCombo,"="& hh &">","="& hh &" selected>")
mmCombo = Replace(mmCombo,"="& mm &">","="& mm &" selected>")



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
<body id="inventory">
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
</div><div id="content">
<!-- InstanceBeginEditable name="content" -->
<h1>Stock #<%=stock%></h1>
<form action="details.asp" method="post" name="frm" id="frm" enctype="multipart/form-data">
    <input type="hidden" name="ID" value="<%=ID%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="2" >
  <tr> 
    <th valign="top" align="left" nowrap="nowrap" colspan="4" class="titlebar"><a href="."><img src="/intranet/images/ico_fup.gif" width="16" height="16" border="0" alt="Back to the list" /></a></th>
  </tr>
  <tr> 
    <td colspan="4" align="left">&nbsp;</td>
    </tr>
  <tr>
    <td height="100%" class="cell_label">Succursale :</td>
    <td width="50%" class="cell_content"><select name="dealer_id">
        <%=DealerCombo(dealer_id)%>
    </select></td>
    <td class="cell_label">Boni :</td>
    <td width="50%" class="cell_content"><input class="err" name="bonus" type="text" id="bonus" value="<%=bonus%>" size="6"/>
      $</td>
  </tr>
    
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label"># Stock :</td>
      <td class="cell_content"><input name="stock" type="text" id="stock" value="<%=stock%>" maxlength="8"/>      </td>
      <td height="100%" align="left" nowrap="nowrap" class="cell_label" >Date de production : </td>
      <td align="left" class="cell_content" ><input type="text" name="nointerestDate" id="nointerestDate" value="<%=MakeLocalDate(nointerestDate)%>" />
        <a href="#" id="triggerint"><img src="/intranet/includes/jscalendar-1.0/img.gif" alt="Calendrier" width="20" height="16" border="0" align="absmiddle" /></a></td>
    </tr>
  
  	  	<script type="text/javascript">
		  Calendar.setup(
			{
			  inputField  : "nointerestDate",         // ID of the input field
			  ifFormat    : "%d/%m/%Y",    // the date format
			  align		  : "Bl",
			  button      : "triggerint"       // ID of the button
			}
		   );
		</script>
    
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label"># Commande :</td>
      <td class="cell_content"><input name="ordernumber" type="text" id="ordernumber" value="<%=ordernumber%>"/>
        <%
	  Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	  if FSO.FileExists(Server.MapPath("/intranet/pdf/prop" & ID & ".pdf")) then
		%>
        <a href="/intranet/pdf/prop<%=ID%>.pdf" target="_blank"><img src="../images/ico_pdf.gif" alt="PDF" width="16" height="16" border="0" align="absmiddle" /></a>
        <%
		FileExists=True
	  End If
	  Set FSO = Nothing
	  
	  
	  IF HaveSaveRights then
	  
		%>
        <a href="#" onClick="ShowUpload(this,'prop');"><img src="../images/ico_upload.gif" alt="T&eacute;l&eacute;charger" width="16" height="16" border="0" align="absmiddle" /></a>
        <%IF FileExists Then%>
        <a href="#" onClick="DeletePDF(<%=ID%>,'prop');"><img src="../images/ico_del.gif" alt="Supprimer" width="16" height="16" border="0" align="absmiddle" /></a>
        <%End If
		Else
		%>
		<img src="../images/ico_upload_off.gif" alt="Veuillez sauvegarder la fiche" width="16" height="16" border="0" align="absmiddle" />
		<%
		End IF		  
		%> 
		Proposition 
		<div id="prop" style="display:none;" class="upload">
		  <input type="file" name="_prop" /><br />
		  <%if id<>"" then %><input type="submit" name="_upload" value="Upload!" /><%end if%>
		</div></td>
      <td height="100%" nowrap="nowrap" class="cell_label">V&eacute;rifier prix  :</td>
      <td class="cell_content"><input <%if checkprice=1 then response.write "checked"%> name="checkprice" type="radio" value="1" id="oui" /><label for="oui">oui</label>
         <input <%if checkprice=0 then response.write "checked"%> name="checkprice" type="radio" value="0" id="non" /><label for="non">non</label>         </td>
    </tr>
    <% IF HaveSaveRights then %>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label"># Commande VENDU :</td>
      <td class="cell_content"><input name="ordersold" type="text" id="ordersold" value="<%=ordersold%>" maxlength="8"/></td>
      <td height="100%" nowrap="nowrap" class="cell_label">&nbsp;</td>
      <td class="cell_content">&nbsp;</td>
    </tr>
    <% End if %>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Nom du client:</td>
      <td class="cell_content"><input name="clientsold" type="text" id="clientsold" value="<%=clientsold%>"/></td>
      <td height="100%" nowrap="nowrap" class="cell_label">Demander SPA   :</td>
      <td class="cell_content"><input <%if reqspa=1 then response.write "checked"%> name="reqspa" type="radio" value="1" id="soui" />
        <label for="soui">oui</label>
        <input <%if reqspa=0 then response.write "checked"%> name="reqspa" type="radio" value="0" id="snon" />
        <label for="snone">non</label></td>
    </tr>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Initiales du vendeur:</td>
      <td class="cell_content"><input name="initialsold" type="text" id="initialsold" value="<%=initialsold%>" maxlength="3" style="width:40px;"/></td>
      <td height="100%" nowrap="nowrap" class="cell_label">Programme vente  : </td>
      <td class="cell_content"><input style="width:100px;" name="salesprogram" type="text" id="salesprogram" value="<%=salesprogram%>"/>
Terme :
  <input style="width:35px;" name="salesterm" type="text" id="salesterm" value="<%=salesterm%>"/>
jours </td>
    </tr>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Marque / Model :</td>
      <td class="cell_content">
		<select name="marque" id="marque">
          <option>-</option>
		  <option <%if marque="Asetrail" then response.write "selected"%>>Asetrail</option>
		  <option <%if marque="BWS" then response.write "selected"%>>BWS</option>
          <option <%if marque="Kalmar" then response.write "selected"%>>Kalmar</option>
          <option <%if marque="International" then response.write "selected"%>>International</option>
          <option <%if marque="Doepker" then response.write "selected"%>>Doepker</option>
		  <option <%if marque="Di-Mond" then response.write "selected"%>>Di-Mond</option>
          <option <%if marque="Isuzu" then response.write "selected"%>>Isuzu</option>
		  <option <%if marque="Yanmar" then response.write "selected"%>>Yanmar</option>
        </select>
		<select name="model" id="model">
          <option>-</option>
          <option <%if model="CF 500" then response.write "selected"%>>CF 500</option>
		  <option <%if model="CV" then response.write "selected"%>>CV</option>
		  <option <%if model="CV515" then response.write "selected"%>>CV515</option>
		  <option <%if model="CV 4x4" then response.write "selected"%>>CV 4x4</option>
          <option <%if model="4300" then response.write "selected"%>>4300</option>
          <option <%if model="4400" then response.write "selected"%>>4400</option>
          <option <%if model="7400" then response.write "selected"%>>7400</option>
          <option <%if model="7500" then response.write "selected"%>>7500</option>
          <option <%if model="7600" then response.write "selected"%>>7600</option>
          <option <%if model="7700" then response.write "selected"%>>7700</option>
          <option <%if model="5500" then response.write "selected"%>>5500</option>
          <option <%if model="5600" then response.write "selected"%>>5600</option>
          <option <%if model="5900" then response.write "selected"%>>5900</option>
          <option <%if model="8600" then response.write "selected"%>>8600</option>
          <option <%if model="9200" then response.write "selected"%>>9200</option>
          <option <%if model="9400" then response.write "selected"%>>9400</option>
          <option <%if model="9900" then response.write "selected"%>>9900</option>
		  <option <%if model="Dry Van Tandem" then response.write "selected"%>>Dry Van Tandem</option>
		  <option <%if model="Dry Van Tridem" then response.write "selected"%>>Dry Van Tridem</option>
		  <option <%if model="FTR" then response.write "selected"%>>FTR</option>
          <option <%if model="TerraStar" then response.write "selected"%>>TerraStar</option>
          <option <%if model="Pro Star 113" then response.write "selected"%>>Pro Star 113</option>
          <option <%if model="Pro Star 122" then response.write "selected"%>>Pro Star 122</option>
          <option <%if model="Pro Star 125" then response.write "selected"%>>Pro Star 125</option>
          <option <%if model="Lone Star" then response.write "selected"%>>Lone Star</option>
		  <option <%if model="LT" then response.write "selected"%>>LT</option>
		  <option <%if model="HX515" then response.write "selected"%>>HX515</option>
		  <option <%if model="HX520" then response.write "selected"%>>HX520</option>
		  <option <%if model="HX615" then response.write "selected"%>>HX615</option>
		  <option <%if model="HX620" then response.write "selected"%>>HX620</option>
		  <option <%if model="HV507" then response.write "selected"%>>HV507</option>
		  <option <%if model="HV513 SFA" then response.write "selected"%>>HV513 SFA</option>
		  <option <%if model="HV607" then response.write "selected"%>>HV607</option>
		  <option <%if model="HV613 SBA" then response.write "selected"%>>HV613 SBA</option>
		  <option <%if model="Galv Super B" then response.write "selected"%>>Galv Super B</option>
		  <option <%if model="Galv Quad" then response.write "selected"%>>Galv Quad</option>
		  <option <%if model="Galv drop deck" then response.write "selected"%>>Galv drop deck</option>
		  <option <%if model="Galv drop deck beaver tail" then response.write "selected"%>>Galv drop deck beaver tail</option>
		  <option <%if model="Galv Flat" then response.write "selected"%>>Galv Flat</option>
		  <option <%if model="Impact Dump" then response.write "selected"%>>Impact Dump</option>
		  <option <%if model="MV607 SBA" then response.write "selected"%>>MV607 SBA</option>
		  <option <%if model="MV607 SBA LP" then response.write "selected"%>>MV607 SBA LP</option>
		  <option <%if model="NPR" then response.write "selected"%>>NPR</option>
		  <option <%if model="NPR écomax" then response.write "selected"%>>NPR écomax</option>
		  <option <%if model="NPR HD Diesel" then response.write "selected"%>>NPR HD Diesel</option>
		  <option <%if model="NPR HD Essence" then response.write "selected"%>>NPR HD Essence</option>
		  <option <%if model="NPRXD" then response.write "selected"%>>NPRXD</option>
		  <option <%if model="NPREFI" then response.write "selected"%>>NPREFI</option>
		  <option <%if model="NQR" then response.write "selected"%>>NQR</option>
		  <option <%if model="NQR Diesel" then response.write "selected"%>>NQR Diesel</option>
		  <option <%if model="NRR" then response.write "selected"%>>NRR</option>
          <option <%if model="MXT" then response.write "selected"%>>MXT</option>
          <option <%if model="RXT" then response.write "selected"%>>RXT</option>
          <option <%if model="CXT" then response.write "selected"%>>CXT</option>
		  <option <%if model="Semi-remorque" then response.write "selected"%>>Semi-remorque</option>
		  <option <%if model="Remorque 25T" then response.write "selected"%>>Remorque 25T</option>
		  <option <%if model="Remorque 20T" then response.write "selected"%>>Remorque 20T</option>
		  <option <%if model="RH" then response.write "selected"%>>RH</option>
		  <option <%if model="RH613" then response.write "selected"%>>RH613</option>
          <option <%if model="Ottawa C30" then response.write "selected"%>>Ottawa C30</option>
          <option <%if model="Ottawa C50" then response.write "selected"%>>Ottawa C50</option>
          <option <%if model="Ottawa C60" then response.write "selected"%>>Ottawa C60</option>
          <option <%if model="Bo&icirc;tes" then response.write "selected"%>>Bo&icirc;tes</option>
		  <option <%if model="V2" then response.write "selected"%>>V2</option>
		  <option <%if model="VI017" then response.write "selected"%>>VI017</option>
		  <option <%if model="VI025" then response.write "selected"%>>VI025</option>
		  <option <%if model="VI035" then response.write "selected"%>>VI035</option>
		  <option <%if model="VI055" then response.write "selected"%>>VI055</option>
        </select>
          <select name="config" id="config">
            <option>-</option>
            <option <%if config="4 x 4" then response.write "selected"%>>4 x 4</option>
            <option <%if config="4 x 2" then response.write "selected"%>>4 x 2</option>
            <option <%if config="6 x 4" then response.write "selected"%>>6 x 4</option>
			<option <%if config="10 x 6" then response.write "selected"%>>10 x 6</option>
        </select>
		&nbsp;Ann&eacute;e:&nbsp;<input type="text" name="strAnnee" size="5" value="<%=strAnnee%>" class="field" /></td>
      <td height="100%" nowrap="nowrap" class="cell_label">Prix coutant :</td>
      <td class="cell_content">
      <% IF HaveSaveRights then %>
      $ <input name="dealernet" type="text" id="dealernet" value="<%=dealernet%>" size="15" />
        <%
	  end if
	  FileExists=False
	  Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	  if FSO.FileExists(Server.MapPath("/intranet/pdf/inv" & ID & ".pdf")) then
		%>
        <a href="/intranet/pdf/inv<%=ID%>.pdf" target="_blank"><img src="../images/ico_pdf.gif" alt="PDF" width="16" height="16" border="0" align="absmiddle" /></a>
        <%
		FileExists=True
	  End If
	  Set FSO = Nothing
	  
	  
	  IF HaveSaveRights then
	  
		%>
        <a href="#" onClick="ShowUpload(this,'inv');"><img src="../images/ico_upload.gif" alt="T&eacute;l&eacute;charger" width="16" height="16" border="0" align="absmiddle" /></a>
        <%IF FileExists Then%>
        <a href="#" onClick="DeletePDF(<%=ID%>,'inv');"><img src="../images/ico_del.gif" alt="Supprimer" width="16" height="16" border="0" align="absmiddle" /></a>
        <%End If
		Else
		%>
        <img src="../images/ico_upload_off.gif" alt="Veuillez sauvegarder la fiche" width="16" height="16" border="0" align="absmiddle" />
        <%
		End IF		  
		%>
Facture
<div id="inv" style="display:none;" class="upload">
  <input type="file" name="_inv" />
  <%if id<>"" then %>
  <input type="submit" name="_upload2" value="Upload!" />
  <%end if%>
</div></td>
    </tr>
	<tr>
      <td height="100%" nowrap="nowrap" class="cell_label"># S&eacute;rie / Ann&eacute;e :</td>
      <td class="cell_content"><input name="serial" type="text" id="serial" value="<%=serial%>" maxlength="9"/></td>
        
        
      
      <% IF HaveSaveRights then %>
      
      <td height="100%" nowrap="nowrap" class="cell_label">Profit :</td>
      <td class="cell_content">
      
      <input name="profit" type="text" id="profit" value="<%=profit%>" size="10" />
        <select name="profit_type" id="profit_type">
          <option value="1">%</option>
          <option <%if profit_type=0 then response.write "selected"%> value="0">$</option>
        </select>
        </td>
        
        <%else
        	response.write " <td height='100%' class='cell_label'></td><td class='cell_content'></td>"
        end if%>
        
        
    </tr>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Date de facture :</td>
      <td class="cell_content"><input onFocus="this.blur();" type="text" name="invoiceDate" id="invoiceDate" value="<%=MakeLocalDate(invoicedate)%>">
          <a href="#" id="trigger"><img src="/intranet/includes/jscalendar-1.0/img.gif" alt="Calendrier" width="20" height="16" border="0" align="absmiddle" /></a>
          <%if daysin <> "0" then%>
          <span class="err">En inventaire depuis <%=daysin%> jours</span>
          <%end if%></td>
          
          
             <script type="text/javascript">
		  Calendar.setup(
			{
			  inputField  : "invoiceDate",         // ID of the input field
			  ifFormat    : "%d/%m/%Y",    // the date format
			  align		  : "Br",
			  button      : "trigger"       // ID of the button
			}
		   );
		</script>
   
      
      <td height="100%" nowrap="nowrap" class="cell_label">Total (&eacute;quip. sp&eacute;cial) :</td>
      <td class="cell_content"><div id="sum_special"></div></td>
    </tr>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Date de r&eacute;ception :</td>
      <td class="cell_content"><input onFocus="this.blur();" type="text" name="receiveDate" id="receiveDate" value="<%=MakeLocalDate(receivedate)%>">
          <a href="#" id="received"><img src="/intranet/includes/jscalendar-1.0/img.gif" alt="Calendrier" width="20" height="16" border="0" align="absmiddle" /></a>
          <%if daysrec <> "0" then%>
          <span class="err">Re&ccedil;u depuis <%=daysrec%> jours</span>
          <%end if%></td>
    
          <td height="100%" nowrap="nowrap" class="cell_label">Prix vendant :</td>
          <td class="cell_content"><div id="sum_special_total"><b><%=retailprice%> $</b></div></td>
          <script type="text/javascript">
		  Calendar.setup(
			{
			  inputField  : "receiveDate",         // ID of the input field
			  ifFormat    : "%d/%m/%Y",    // the date format
			  align		  : "Br",
			  button      : "received"       // ID of the button
			}
		   );
		</script>      
      </tr>
    
    <tr>
      <td height="100%" colspan="4" align="left" nowrap="nowrap" >&nbsp;</td>
      </tr>
    
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">WB:</td>
      <td class="cell_content"><input name="wb" type="text" id="wb" value="<%=wb%>" size="10" /></td>
      <td height="100%" nowrap="nowrap" class="cell_label">D&eacute;mo  :</td>
      <td class="cell_content">
          <select name="demo" id="demo">
        <option>-</option>
        <option <%if demo="Dompeur 6 roues" then response.write "selected"%>>Dompeur 6 roues</option>
        <option <%if demo="Dompeur 10 roues" then response.write "selected"%>>Dompeur 10 roues</option>
        <option <%if demo="Dompeur 12 roues" then response.write "selected"%>>Dompeur 12 roues</option>
        <option <%if demo="Dry box 16'" then response.write "selected"%>>Dry box 16‘</option>
        <option <%if demo="Dry box 20’" then response.write "selected"%>>Dry box  20’</option>
		<option <%if demo="Dry box 22’" then response.write "selected"%>>Dry box 22’</option>
        <option <%if demo="Dry box 24’" then response.write "selected"%>>Dry box 24’</option>
        <option <%if demo="Dry box 26’" then response.write "selected"%>>Dry box 26’</option>
        <option <%if demo="Dry box 28’" then response.write "selected"%>>Dry box 28’</option>
        <option <%if demo="Dry box 30’" then response.write "selected"%>>Dry box 30’</option>
        <option <%if demo="Roll off" then response.write "selected"%>>Roll off</option>
        <option <%if demo="Plate-forme (Quincaillerie ou autre application)" then response.write "selected"%>>Plate-forme (Quincaillerie ou autre application)</option>
		<option <%if demo="Plate-forme de remorquage" then response.write "selected"%>>Plate-forme de remorquage</option>
        <option <%if demo="Towing" then response.write "selected"%>>Towing</option>
        <option <%if demo="VR" then response.write "selected"%>>VR</option>
        <option <%if demo="Bétonnière" then response.write "selected"%>>Bétonnière</option>
        <option <%if demo="Autres" then response.write "selected"%>>Autres</option>
          </select>      </td>
    </tr>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Moteur : </td>
      <td class="cell_content"><select name="engine" id="engine">
        <option>-</option>
		<option <%if engine="A26" then response.write "selected"%>>A26</option>
		<option <%if engine="B6.7" then response.write "selected"%>>B6.7</option>
        <option <%if engine="ISB" then response.write "selected"%>>ISB</option>
        <option <%if engine="ISL" then response.write "selected"%>>ISL</option>
        <option <%if engine="ISM" then response.write "selected"%>>ISM</option>
        <option <%if engine="ISX" then response.write "selected"%>>ISX</option>
        <option <%if engine="C11" then response.write "selected"%>>C11</option>
        <option <%if engine="C13" then response.write "selected"%>>C13</option>
        <option <%if engine="C15" then response.write "selected"%>>C15</option>
        <option <%if engine="VT275" then response.write "selected"%>>VT275</option>
        <option <%if engine="VT365" then response.write "selected"%>>VT365</option>
        <option <%if engine="DT466" then response.write "selected"%>>DT466</option>
        <option <%if engine="DT570" then response.write "selected"%>>DT570</option>
		<option <%if engine="Inter 6.6" then response.write "selected"%>>Inter 6.6</option>		
        <option <%if engine="HT530" then response.write "selected"%>>HT530</option>
        <option <%if engine="HT570" then response.write "selected"%>>HT570</option>
		<option <%if engine="L9" then response.write "selected"%>>L9</option>
		<option <%if engine="N9" then response.write "selected"%>>N9</option>
		<option <%if engine="N10" then response.write "selected"%>>N10</option>
        <option <%if engine="N13" then response.write "selected"%>>N13</option>
        <option <%if engine="MAXX FORCE DT" then response.write "selected"%>>MAXX FORCE DT</option>
        <option <%if engine="MAXX FORCE 5" then response.write "selected"%>>MAXX FORCE 5</option>
        <option <%if engine="MAXX FORCE 7" then response.write "selected"%>>MAXX FORCE 7</option>
        <option <%if engine="MAXX FORCE 8" then response.write "selected"%>>MAXX FORCE 8</option>
        <option <%if engine="MAXX FORCE 9" then response.write "selected"%>>MAXX FORCE 9</option>
        <option <%if engine="MAXX FORCE 10" then response.write "selected"%>>MAXX FORCE 10</option>
        <option <%if engine="MAXX FORCE 11" then response.write "selected"%>>MAXX FORCE 11</option>
        <option <%if engine="MAXX FORCE 13" then response.write "selected"%>>MAXX FORCE 13</option>
        <option <%if engine="MAXX FORCE 15" then response.write "selected"%>>MAXX FORCE 15</option>
		<option <%if engine="4JJ1-TC" then response.write "selected"%>>4JJ1-TC</option>
		<option <%if engine="4HK1-TC" then response.write "selected"%>>4HK1-TC</option>
		<option <%if engine="GMPT-V8" then response.write "selected"%>>GMPT-V8</option>
		<option <%if engine="X15" then response.write "selected"%>>X15</option>
       </select>
       
       <strong>HP:</strong>   <input name="hp" type="text" id="hp" value="<%=hp%>" size="3"/>
	   <strong>Carburant alt:</strong> 
	   <input <%if intCarburantAlt=1 then response.write "checked"%> name="intCarburantAlt" type="radio" value="1" id="soui" />
        <label for="soui">oui</label>
        <input <%if intCarburantAlt=0 then response.write "checked"%> name="intCarburantAlt" type="radio" value="0" id="snon" />
        <label for="snone">non</label>
	   </td>
      <td height="100%" nowrap="nowrap" class="cell_label">Suspension Arri&egrave;re : </td>
      <td class="cell_content"><select name="rearsuspension" id="rearsuspension">
          <option>-</option>
		    <option <%if rearsuspension="9 880" then response.write "selected"%>>9 880</option>
			<option <%if rearsuspension="11 000" then response.write "selected"%>>11 000</option>
			<option <%if rearsuspension="12 000" then response.write "selected"%>>12 000</option>
			<option <%if rearsuspension="12 900" then response.write "selected"%>>12 900</option>
			<option <%if rearsuspension="13 000" then response.write "selected"%>>13 000</option>
			<option <%if rearsuspension="13 500" then response.write "selected"%>>13 500</option>
			<option <%if rearsuspension="14 308" then response.write "selected"%>>14 308</option>
			<option <%if rearsuspension="14 550" then response.write "selected"%>>14 550</option>
			<option <%if rearsuspension="15 500" then response.write "selected"%>>15 500</option>
			<option <%if rearsuspension="17 500" then response.write "selected"%>>17 500</option>
			<option <%if rearsuspension="18 500" then response.write "selected"%>>18 500</option>
			<option <%if rearsuspension="19 800" then response.write "selected"%>>19 800</option>
			<option <%if rearsuspension="20 000" then response.write "selected"%>>20 000</option>
			<option <%if rearsuspension="21 000" then response.write "selected"%>>21 000</option>
			<option <%if rearsuspension="23 000" then response.write "selected"%>>23 000</option>
			<option <%if rearsuspension="23 500" then response.write "selected"%>>23 500</option>
			<option <%if rearsuspension="30 000" then response.write "selected"%>>30 000</option>
			<option <%if rearsuspension="31 000" then response.write "selected"%>>31 000</option>
			<option <%if rearsuspension="34 000" then response.write "selected"%>>34 000</option>
			<option <%if rearsuspension="36 000" then response.write "selected"%>>36 000</option>
			<option <%if rearsuspension="40 000" then response.write "selected"%>>40 000</option>
			<option <%if rearsuspension="46 000" then response.write "selected"%>>46 000</option>
			<option <%if rearsuspension="52 000" then response.write "selected"%>>52 000</option>
			<option <%if rearsuspension="65 000" then response.write "selected"%>>65 000</option>
			<option <%if rearsuspension="80 000" then response.write "selected"%>>80 000</option>
        </select></td>
    </tr>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Transmission : </td>
      <td class="cell_content"><select name="transtype" id="transtype">
        <option value="">-</option>
        <option <%if transtype="Allison" then response.write "selected"%>>Allison</option>
		<option <%if transtype="Eaton Procision Auto" then response.write "selected"%>>Eaton Procision Auto</option>		
		<option <%if transtype="EATON ULTRASHIFT+" then response.write "selected"%>>EATON ULTRASHIFT+</option>
		<option <%if transtype="Eaton Endurant" then response.write "selected"%>>EATON Endurant</option>
        <option <%if transtype="Fuller" then response.write "selected"%>>Fuller</option>
		<option <%if transtype="AISIN A460" then response.write "selected"%>>AISIN A460</option>
		<option <%if transtype="AISIN A465" then response.write "selected"%>>AISIN A465</option>
		<option <%if transtype="6L90 GMPT" then response.write "selected"%>>6L90 GMPT</option>
              </select> <input name="transmission" type="text" id="transmission" value="<%=transmission%>"/>      </td>
      <td height="100%" nowrap="nowrap" class="cell_label">Grandeur des pneus :</td>
      <td class="cell_content"><input name="tiresize" type="text" id="tiresize" value="<%=tiresize%>" size="10" /></td>
    </tr>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Essieu Avant : </td>
      <td class="cell_content"><select name="frontaxle" id="frontaxle">
          <option>-</option>
			<option <%if frontaxle="6 000" then response.write "selected"%>>6 000</option>
			<option <%if frontaxle="6 830" then response.write "selected"%>>6 830</option>
			<option <%if frontaxle="7 000" then response.write "selected"%>>7 000</option>
			<option <%if frontaxle="7 275" then response.write "selected"%>>7 275</option>
			<option <%if frontaxle="8 000" then response.write "selected"%>>8 000</option>
			<option <%if frontaxle="8 500" then response.write "selected"%>>8 500</option>
			<option <%if frontaxle="9 000" then response.write "selected"%>>9 000</option>
			<option <%if frontaxle="10 000" then response.write "selected"%>>10 000</option>
			<option <%if frontaxle="11 000" then response.write "selected"%>>11 000</option>
			<option <%if frontaxle="12 000" then response.write "selected"%>>12 000</option>
            <option <%if frontaxle="12 350" then response.write "selected"%>>12 350</option>
			<option <%if frontaxle="13 000" then response.write "selected"%>>13 000</option>
			<option <%if frontaxle="13 200" then response.write "selected"%>>13 200</option>
			<option <%if frontaxle="14 000" then response.write "selected"%>>14 000</option>
			<option <%if frontaxle="14 600" then response.write "selected"%>>14 600</option>
			<option <%if frontaxle="16 000" then response.write "selected"%>>16 000</option>
			<option <%if frontaxle="18 000" then response.write "selected"%>>18 000</option>
			<option <%if frontaxle="20 000" then response.write "selected"%>>20 000</option>
			<option <%if frontaxle="21 000" then response.write "selected"%>>21 000</option>
			<option <%if frontaxle="22 000" then response.write "selected"%>>22 000</option>
			<option <%if frontaxle="23 000" then response.write "selected"%>>23 000</option>
			<option <%if frontaxle="50 000" then response.write "selected"%>>50 000</option>
        </select></td>
      <td height="100%" nowrap="nowrap" class="cell_label">Roues : </td>
      <td class="cell_content"><select name="wheel" id="wheel">
        <option>-</option>
        <option <%if wheel="acier" then response.write "selected"%>>acier</option>
        <option <%if wheel="aluminium" then response.write "selected"%>>aluminium</option>
        <option <%if wheel="AR aluminium ext Acier int" then response.write "selected"%>>AR aluminium ext Acier int</option>
		<option <%if wheel="av-alu/ar-acier" then response.write "selected"%>>av-alu/ar-acier</option>
              </select></td>
    </tr>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Essieu Arri&egrave;re : </td>
      <td class="cell_content"><select name="rearaxle" id="rearaxle">
          <option>-</option>
			<option <%if rearaxle="10 500" then response.write "selected"%>>10 500</option>
			<option <%if rearaxle="11 000" then response.write "selected"%>>11 000</option>
			<option <%if rearaxle="11 020" then response.write "selected"%>>11 020</option>
			<option <%if rearaxle="12 200" then response.write "selected"%>>12 200</option>
			<option <%if rearaxle="13 500" then response.write "selected"%>>13 500</option>
			<option <%if rearaxle="14 550" then response.write "selected"%>>14 550</option>
			<option <%if rearaxle="15 500" then response.write "selected"%>>15 500</option>
			<option <%if rearaxle="17 000" then response.write "selected"%>>17 000</option>
			<option <%if rearaxle="19 000" then response.write "selected"%>>19 000</option>
			<option <%if rearaxle="19 800" then response.write "selected"%>>19 800</option>
			<option <%if rearaxle="20 000" then response.write "selected"%>>20 000</option>
			<option <%if rearaxle="21 000" then response.write "selected"%>>21 000</option>
			<option <%if rearaxle="22 000" then response.write "selected"%>>22 000</option>
			<option <%if rearaxle="23 000" then response.write "selected"%>>23 000</option>
			<option <%if rearaxle="24 000" then response.write "selected"%>>24 000</option>
			<option <%if rearaxle="26 000" then response.write "selected"%>>26 000</option>
			<option <%if rearaxle="30 000" then response.write "selected"%>>30 000</option>
			<option <%if rearaxle="38 000" then response.write "selected"%>>38 000</option>
			<option <%if rearaxle="34 000" then response.write "selected"%>>34 000</option>
			<option <%if rearaxle="40 000" then response.write "selected"%>>40 000</option>
			<option <%if rearaxle="46 000" then response.write "selected"%>>46 000</option>
			<option <%if rearaxle="52 000" then response.write "selected"%>>52 000</option>
			<option <%if rearaxle="58 000" then response.write "selected"%>>58 000</option>
			<option <%if rearaxle="70 000" then response.write "selected"%>>70 000</option>	  
			<option <%if rearaxle="134 000" then response.write "selected"%>>134 000</option>
        </select></td>
      <td height="100%" nowrap="nowrap" class="cell_label">Ratio : </td>
      <td class="cell_content"><input name="ratio" type="text" id="ratio" value="<%=ratio%>"/>      </td>
    </tr>
    <tr>
      <td height="100%" colspan="4" nowrap="nowrap">&nbsp;</td>
      </tr>
    <tr>
      <td height="100%" nowrap="nowrap" class="cell_label">Couchette : </td>
      <td class="cell_content"><select name="sleeper" id="sleeper">
          <option>-</option>
			<option <%if sleeper="LR 51" then response.write "selected"%>>LR 51</option>
			<option <%if sleeper="HR 51" then response.write "selected"%>>HR 51</option>
			<option <%if sleeper="HR 72" then response.write "selected"%>>HR 72</option>
			<option <%if sleeper="Sky Rise" then response.write "selected"%>>Sky Rise</option>
			<option <%if sleeper="Day Cab" then response.write "selected"%>>Day Cab</option>
            <option <%if sleeper="Crew Cab" then response.write "selected"%>>Crew Cab</option>
			<option <%if sleeper="Extended Cab" then response.write "selected"%>>Extended Cab</option>
			<option <%if sleeper="LR 56" then response.write "selected"%>>LR 56</option>
			<option <%if sleeper="HR 56" then response.write "selected"%>>HR 56</option>
			<option <%if sleeper="HR 73" then response.write "selected"%>>HR 73</option>
        </select></td>
      <td height="100%" nowrap="nowrap" class="cell_label">Couleur : </td>
      <td class="cell_content"><input name="color" type="text" id="color" value="<%=color%>"/>      </td>
    </tr>
    <tr>
      <td rowspan="2" valign="top" nowrap="nowrap" class="cell_label">&Eacute;quipement sp&eacute;cial  : </td>
      <td rowspan="2" valign="top" class="cell_content">
      	<textarea name="specialequipment" cols="50" rows="5" id="specialequipment"><%=specialequipment%></textarea><br/>
      	<input type="text" name="special1" size="60" value="<%=special1%>" /> 
      		<input type="text" name="specialprice1" size="8" id="specialprice1" class="special_price_input" value="<%=specialprice1%>" /><span class="invemntory_equip_margin">$</span>

<!-- Special Upload 1 -->
<% 
FileExists=False
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(Server.MapPath("/intranet/pdf/special-1-" & ID & ".pdf")) then
%>
	<a href="/intranet/pdf/special-1-<%=ID%>.pdf" target="_blank"><img src="../images/ico_pdf.gif" alt="PDF" width="16" height="16" border="0" align="absmiddle" /></a>
<%
	FileExists=True
End If

Set FSO = Nothing
IF HaveSaveRights then
%>
	<a href="#" onClick="ShowUpload(this,'special_upload1');"><img src="../images/ico_upload.gif" alt="T&eacute;l&eacute;charger" width="16" height="16" border="0" align="absmiddle" /></a>
	<%IF FileExists Then%>
		<a href="#" onClick="DeletePDF(<%=ID%>,'special-1-');"><img src="../images/ico_del.gif" alt="Supprimer" width="16" height="16" border="0" align="absmiddle" /></a>
	<%End If
Else
%>
	<img src="../images/ico_upload_off.gif" alt="Veuillez sauvegarder la fiche" width="16" height="16" border="0" align="absmiddle" />
<%
End IF		  
%>

<div id="special_upload1" style="display:none;" class="upload">
	<input type="file" name="_special1" /><br />
	<%if id<>"" then %><input type="submit" name="_upload3" value="Upload!" /><%end if%>
</div>
<!-- /Special Upload 1 -->

      	<br/>
      	<input type="text" name="special2" size="60" value="<%=special2%>" /> 
      		<input type="text" name="specialprice2" size="8" id="specialprice2" class="special_price_input" value="<%=specialprice2%>" /><span class="invemntory_equip_margin">$</span>

<!-- Special Upload 2 -->
<% 
FileExists=False
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(Server.MapPath("/intranet/pdf/special-2-" & ID & ".pdf")) then
%>
	<a href="/intranet/pdf/special-2-<%=ID%>.pdf" target="_blank"><img src="../images/ico_pdf.gif" alt="PDF" width="16" height="16" border="0" align="absmiddle" /></a>
<%
	FileExists=True
End If

Set FSO = Nothing
IF HaveSaveRights then
%>
	<a href="#" onClick="ShowUpload(this,'special_upload2');"><img src="../images/ico_upload.gif" alt="T&eacute;l&eacute;charger" width="16" height="16" border="0" align="absmiddle" /></a>
	<%IF FileExists Then%>
		<a href="#" onClick="DeletePDF(<%=ID%>,'special-2-');"><img src="../images/ico_del.gif" alt="Supprimer" width="16" height="16" border="0" align="absmiddle" /></a>
	<%End If
Else
%>
	<img src="../images/ico_upload_off.gif" alt="Veuillez sauvegarder la fiche" width="16" height="16" border="0" align="absmiddle" />
<%
End IF		  
%>

<div id="special_upload2" style="display:none;" class="upload">
	<input type="file" name="_special2" /><br />
	<%if id<>"" then %><input type="submit" name="_upload4" value="Upload!" /><%end if%>
</div>
<!-- /Special Upload 2 -->

      	<br/>
      	<input type="text" name="special3" size="60" value="<%=special3%>" /> 
      		<input type="text" name="specialprice3" size="8" id="specialprice3" class="special_price_input" value="<%=specialprice3%>" /><span class="invemntory_equip_margin">$</span>

<!-- Special Upload 3 -->
<% 
FileExists=False
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(Server.MapPath("/intranet/pdf/special-3-" & ID & ".pdf")) then
%>
	<a href="/intranet/pdf/special-3-<%=ID%>.pdf" target="_blank"><img src="../images/ico_pdf.gif" alt="PDF" width="16" height="16" border="0" align="absmiddle" /></a>
<%
	FileExists=True
End If

Set FSO = Nothing
IF HaveSaveRights then
%>
	<a href="#" onClick="ShowUpload(this,'special_upload3');"><img src="../images/ico_upload.gif" alt="T&eacute;l&eacute;charger" width="16" height="16" border="0" align="absmiddle" /></a>
	<%IF FileExists Then%>
		<a href="#" onClick="DeletePDF(<%=ID%>,'special-3-');"><img src="../images/ico_del.gif" alt="Supprimer" width="16" height="16" border="0" align="absmiddle" /></a>
	<%End If
Else
%>
	<img src="../images/ico_upload_off.gif" alt="Veuillez sauvegarder la fiche" width="16" height="16" border="0" align="absmiddle" />
<%
End IF		  
%>

<div id="special_upload3" style="display:none;" class="upload">
	<input type="file" name="_special3" /><br />
	<%if id<>"" then %><input type="submit" name="_upload5" value="Upload!" /><%end if%>
</div>
<!-- /Special Upload 3 -->

      	<br/>
      	<input type="text" name="special4" size="60" value="<%=special4%>" /> 
      		<input type="text" name="specialprice4" size="8" id="specialprice4" class="special_price_input" value="<%=specialprice4%>" /><span class="invemntory_equip_margin">$</span>

<!-- Special Upload 4 -->
<% 
FileExists=False
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(Server.MapPath("/intranet/pdf/special-4-" & ID & ".pdf")) then
%>
	<a href="/intranet/pdf/special-4-<%=ID%>.pdf" target="_blank"><img src="../images/ico_pdf.gif" alt="PDF" width="16" height="16" border="0" align="absmiddle" /></a>
<%
	FileExists=True
End If

Set FSO = Nothing
IF HaveSaveRights then
%>
	<a href="#" onClick="ShowUpload(this,'special_upload4');"><img src="../images/ico_upload.gif" alt="T&eacute;l&eacute;charger" width="16" height="16" border="0" align="absmiddle" /></a>
	<%IF FileExists Then%>
		<a href="#" onClick="DeletePDF(<%=ID%>,'special-4-');"><img src="../images/ico_del.gif" alt="Supprimer" width="16" height="16" border="0" align="absmiddle" /></a>
	<%End If
Else
%>
	<img src="../images/ico_upload_off.gif" alt="Veuillez sauvegarder la fiche" width="16" height="16" border="0" align="absmiddle" />
<%
End IF		  
%>

<div id="special_upload4" style="display:none;" class="upload">
	<input type="file" name="_special4" /><br />
	<%if id<>"" then %><input type="submit" name="_upload6" value="Upload!" /><%end if%>
</div>
<!-- /Special Upload 4 -->

      	<br/>
      	<input type="text" name="special5" size="60" value="<%=special5%>" /> 
      		<input type="text" name="specialprice5" size="8" id="specialprice5" class="special_price_input" value="<%=specialprice5%>" /><span class="invemntory_equip_margin">$</span>

<!-- Special Upload 5 -->
<% 
FileExists=False
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(Server.MapPath("/intranet/pdf/special-5-" & ID & ".pdf")) then
%>
	<a href="/intranet/pdf/special-5-<%=ID%>.pdf" target="_blank"><img src="../images/ico_pdf.gif" alt="PDF" width="16" height="16" border="0" align="absmiddle" /></a>
<%
	FileExists=True
End If

Set FSO = Nothing
IF HaveSaveRights then
%>
	<a href="#" onClick="ShowUpload(this,'special_upload5');"><img src="../images/ico_upload.gif" alt="T&eacute;l&eacute;charger" width="16" height="16" border="0" align="absmiddle" /></a>
	<%IF FileExists Then%>
		<a href="#" onClick="DeletePDF(<%=ID%>,'special-5-');"><img src="../images/ico_del.gif" alt="Supprimer" width="16" height="16" border="0" align="absmiddle" /></a>
	<%End If
Else
%>
	<img src="../images/ico_upload_off.gif" alt="Veuillez sauvegarder la fiche" width="16" height="16" border="0" align="absmiddle" />
<%
End IF		  
%>

<div id="special_upload5" style="display:none;" class="upload">
	<input type="file" name="_special5" /><br />
	<%if id<>"" then %><input type="submit" name="_upload7" value="Upload!" /><%end if%>
</div>
<!-- /Special Upload 5 -->

      	<br/>

      	<div class="inventory_equip_total">Total: <b>$ 0</b></div>      </td>
      <td height="100%" valign="top" nowrap="nowrap" class="cell_label">Lieu physique   : </td>
      <td valign="top" class="cell_content"><select name="location" id="location">
          <option>-</option>
                <option <%if location="Anjou" then response.write "selected"%>>Anjou</option>
                <option <%if location="Boucherville" then response.write "selected"%>>Boucherville</option>
                <option <%if location="Chicoutimi" then response.write "selected"%>>Chicoutimi</option>
                <option <%if location="Clermont" then response.write "selected"%>>Clermont</option>
                <option <%if location="Drummondville" then response.write "selected"%>>Drummondville</option>
                <option <%if location="Forestville" then response.write "selected"%>>Forestville</option>
                <option <%if location="Joliette" then response.write "selected"%>>Joliette</option>
				<option <%if location="Granby" then response.write "selected"%>>Granby</option>
				<option <%if location="Matane" then response.write "selected"%>>Matane</option>
				<option <%if location="Montréal" then response.write "selected"%>>Montréal</option>
                <option <%if location="Québec" then response.write "selected"%>>Québec</option>
				<option <%if location="Rimouski" then response.write "selected"%>>Rimouski</option>
                <option <%if location="Rivière-du-Loup" then response.write "selected"%>>Rivière-du-Loup</option>
                <option <%if location="Sept-Îles" then response.write "selected"%>>Sept-&Icirc;les</option>  
                <option <%if location="St-Georges" then response.write "selected"%>>St-Georges</option>
                <option <%if location="St-Hyacinthe" then response.write "selected"%>>St-Hyacinthe</option>
				<option <%if location="St-Marie de Beauce" then response.write "selected"%>>St-Marie de Beauce</option>
                <option <%if location="Shawinigan" then response.write "selected"%>>Shawinigan</option>
                <option <%if location="Thetford Mines" then response.write "selected"%>>Thetford Mines</option>
                <option <%if location="Trois-Rivières" then response.write "selected"%>>Trois-Rivi&egrave;res</option>
                <option <%if location="Victoriaville" then response.write "selected"%>>Victoriaville</option>
                <option <%if location="Sous-traitant" then response.write "selected"%>>Sous-traitant</option>
        </select></td>
    </tr>
    <tr>
      <td height="100%" valign="top" nowrap="nowrap" class="cell_label">&nbsp;</td>
      <td valign="top" class="cell_content">&nbsp;</td>
    </tr>
    <tr>
      <td height="100%" colspan="4" nowrap="nowrap">&nbsp;</td>
    </tr>
    
    <tr>
      <td valign="top" class="cell_label"><label>Images :</label>
      <p style="font-weight:normal;margin:0;"><em>Utiliser votre souris pour modifier l'ordre d'affichage des images</em></p></td>
      <td class="cell_content" id="allpics" colspan="3"><%
                       Set rs=Conn.execute("Select * from inv_pictures WHERE product_id='" & ID & "' ORDER BY intorder")
                       Do While Not rs.EOF
                       %>
                  <table border="0" cellspacing="0" cellpadding="0" style="float:left;margin:0 5px 5px 0;border:1px solid #666; background-color:#EEE;padding:1px;">
                    <tr>
                      <td colspan="2" align="center"><a data-lightbox="<%=rs("name")%>"  href="/intranet/pics/<%=rs("name")%>" target="_blank"><img src="/intranet/pics/small/<%=rs("name")%>" border="1" height="60" alt="Click for a larger image" /></a></td>
                    </tr>
					<%IF Session("Level")=1 then%>
                    <tr>
                      <td class='pic_sort'><input name="_picture_id" type="hidden" value="<%=rs("id")%>" /></td>
                      <td align="right"><input type="button" onClick="DeletePic(<%=ID%>,<%=rs("id")%>);" value="X" alt="Remove" class="smbutton" style="width:18px;height:18px;" /></td>
                    </tr>
                    <%end IF%>
                  </table>
          <%
                            rs.MoveNext
                        Loop
                        %></td>
    </tr>
    <%IF Session("Level")=1 then%>
    <tr>
      <td valign="top" class="cell_label">&nbsp;</td>
      <td class="cell_content" id="browsetd" colspan="3"><input type="file" name="_img_file"/>
          <input type="button" id="addpic" value="+" class="smbutton" style="width:18px;height:18px;" />      
	</td>
    </tr>
    <%end if%>
    
    <tr>
      <td height="100%" colspan="4" nowrap="nowrap">&nbsp;</td>
    </tr>
    
    <tr>
      <td nowrap="nowrap" class="cell_label">&Eacute;tat:</td>
      <td colspan="3" class="cell_content"><%
	  If resuserid=0 then
	  %>Pr&eacute;sentement disponible<%
	  Else
	  %>
      Pr&eacute;sentement r&eacute;serv&eacute; par <a href="mailto:<%=GetUserEmail(resuserid)%>"><%=GetUserName(resuserid)%></a> 
      <%
	  IF HaveSaveRights=1 then
	  %>
	  pour <strong><%=resclient%></strong> 
	  <%	  
	  end if
	  %>      
      jusqu'au <%=MakeLocalDate(resdatetime) & " " & Right(formatNumber(Hour(resdatetime)/100,2),2) & ":" & Right(formatNumber(Minute(resdatetime)/100,2),2)%><%
	  End If
	  %></td>
      </tr>
	<%
	IF HaveSaveRights=1 then
	  If resuserid=0 then
	%>
    <tr>
      <td align="left" nowrap="nowrap" class="cell_label">R&eacute;serv&eacute;:</td>
      <td colspan="3" class="cell_content">pour: 
        <select name="_resuserid" id="resuserid" class="field">
          <%=MakeUserCombo(0)%>
      </select> 
      	&nbsp;
        client: <input type="text" id="resclient" name="_resclient" value="<%=resclient%>" />&nbsp;
        jusqu'au 
        <input onFocus="this.blur();" type="text" name="_resdatetime" id="_resdatetime" value="<%=resdatetime%>" />
        <a href="#" id="restrigger"><img src="/intranet/includes/jscalendar-1.0/img.gif" alt="Calendrier" width="20" height="16" border="0" align="absmiddle" /></a></td>
      </tr>
	  
	   <script type="text/javascript">
		  Calendar.setup(
			{
			  inputField  : "_resdatetime",         // ID of the input field
			  ifFormat    : "%d/%m/%Y %H:%M",    // the date format
			  showsTime	  : true,
			  align		  : "Tr",
			  singleClick : false,
			  button      : "restrigger"       // ID of the button
			}
		   );
		</script>
	  
	  <%
	  End IF
	  %>
    <tr>
      <td align="left" nowrap="nowrap" class="cell_label">Action:</td>
      <td colspan="3" class="cell_content">
	  	<%
		If resuserid=0 then
		%>
        <input id="reserve" name="_reserve" type="button" class="field" value="R&eacute;server!" />
      	<%
		Else
		%>
        <input name="_cancel" type="submit" class="field" value="Lib&eacute;rer!" />
		<%End IF%></td>
      </tr>
	<%
	End IF
	%>
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
      <td colspan="3" align="left" class="cell_content" ><input <%if DisplayOnWebSite=1 then response.write "checked"%> name="DisplayOnWebSite" type="radio" value="1" id="soui" />
        <label for="soui">oui</label>
        <input <%if DisplayOnWebSite=0 then response.write "checked"%> name="DisplayOnWebSite" type="radio" value="0" id="snon" />
        <label for="snone">non</label></td>
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
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>