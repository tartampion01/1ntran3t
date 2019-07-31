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
<%

If Request.QueryString("page") = "" then
   Page = 1 'We're on the first page
Else
   Page = Request.QueryString("page")
End If

If uniqueListName="" then
	uniqueListName = Replace(Request.ServerVariables("SCRIPT_NAME"),"/","")
	uniqueListName = Replace(uniqueListName,".","")
End If

If Session(uniqueListName&"OrderBy")="" AND defaultOrderBy<>"" then
	Session(uniqueListName&"OrderBy")= defaultOrderBy
End If

Page=Cint(Page)

currentPage = Request.ServerVariables("SCRIPT_NAME")

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("cstring")

'Explicitly Create a recordset object
Set rs = Server.CreateObject("ADODB.Recordset")

'Set the cursor location property
rs.CursorLocation = 3 'adUseClient

If Session("ItemPerPage") = "" then	Session("ItemPerPage")=999
If Request.QueryString("ItemPerPage")<>"" then Session("ItemPerPage")=Request.QueryString("ItemPerPage")
itemperpage = Session("ItemPerPage")
if itemperpage = "-" then itemperpage = 999

'For Print
'If request.QueryString("print") <> "" then
'	itemperpage = 999
'	page = 1
'End If


'Set the cache size = to the # of records/page
rs.CacheSize = itemperpage

If Request("OrderBy") <> "" then
	If Session(uniqueListName & "Sort")="ASC" then
		Session(uniqueListName & "Sort")="DESC"
	Else
		Session(uniqueListName & "Sort")="ASC"
	End If
	Session(uniqueListName & "Orderby")=Request("Orderby")
End If

'Readonly user
If Session("Level") = 3 then
	strSQL = strSQL & "WHERE dealer_id='" & Session("dealer_id") & "' "
End If

If DefaultWHERE = "" then
	strSQl = strSQL & "WHERE "
Else
	strSQl = strSQL & "WHERE " & DefaultWHERE & "   AND "
End If

For i = 0 to Ubound(Headers,1)
	If Request.QueryString(Headers(i,0))<>"" then
		If Request.QueryString(Headers(i,0))="-1" then
			Session(uniqueListName&(Headers(i,0)))=""
		Else
			Session(uniqueListName&(Headers(i,0)))=Request.QueryString(Headers(i,0))
		End IF
	End If
	IF Session(uniqueListName&(Headers(i,0))) <> "" and Headers(i,2) then
		theField = tableName & "." & Headers(i,0)
		If Ubound(Headers,2)=6 then
			IF Headers(i,6)<> "" then
				theField = Headers(i,6)
			End IF
		End If
		strSQl = strSQL & theField & "='" & Session(uniqueListName&(Headers(i,0))) & "'   AND "
	End If
Next

'Search Feature
IF Request.form("Search")<>"" then
	Session(uniqueListName & "Search")= Trim(Request.form("Search"))
End If
'Reset
If Request.QueryString = "reset" then Session(uniqueListName & "Search")=""

SearchArray = Split(Session(uniqueListName & "Search")," ")
SearchItemsArray = Split(SearchItems,",")

If Session(uniqueListName & "Search") <> "" then
	strSQL = strSQL & "(" 
	For i = 0 to Ubound(SearchArray,1)
		For j = 0 to Ubound(SearchItemsArray,1)
			If InStr(SearchItemsArray(j),"ID")<>0 then
				If IsNumeric(SearchArray(i)) then
					strSQL = strSQL &  SearchItemsArray(j) & " = " & SearchArray(i) & "      OR "
				End If
			Else
				strSQL = strSQL &  SearchItemsArray(j) & " Like '%" & SearchArray(i) & "%'    OR "
			End If
		Next
	Next
	strSQL = Left(strSQL,Len(strSQL)-6)&")"
Else
	strSQL = Left(strSQL,Len(strSQL)-6)
End If


IF Session(uniqueListName & "OrderBy") <> "" then
	if  Session(uniqueListName & "Sort") = "" then  Session(uniqueListName & "Sort") = "asc"
	strSQl = strSQL & " Order by " & Session(uniqueListName & "OrderBy") & " " & Session(uniqueListName & "Sort")
End If

strSQl = strSQL & ";"

'On Error resume next
Response.Write(strSQL)
'Response.End

rs.Open strSQL, Conn

count = rs.RecordCount

If count <> 0 then
	rs.MoveFirst
	rs.PageSize = itemperpage

	'Get the max number of pages
	LastPage = rs.PageCount

	'Set the absolute page
	rs.AbsolutePage = Page
End If

%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css">
<body>

<table width="100%" border="0" cellspacing="0" cellpadding="3">
  <tr>
    <td><h1><%=Title%></h1></td>
<%If SearchItems<>"" then%>
    <form method="post" action="<%=currentPage%>">
	  <td width="33%" valign="bottom">
        <table border="0" cellspacing="0" cellpadding="0" class="noprint">
          <tr>
            <td nowrap>Recherche rapide :</td>
            <td>
              <input type="text" name="Search" value="<%=Session(uniqueListName & "Search")%>">
            </td>
            <td nowrap>
              <input type="image" border="0" name="Submit" src="/intranet/images/ok.gif" alt="Rechercher">
              <a href="./?reset"><img src="/intranet/images/ico_cancel.gif" width="18" height="18" border="0" alt="Annuler"></a>            </td>
          </tr>
        </table>
      </td>
	</form>
<%End If%>

<%
if print_action = "" then print_action = "href=""#"" onClick=""printit();"""
%>

    <td width="33%" align="right" valign="middle"><a <%=print_action%> class="bt prn noprint">imprimer</a><%if session("level")=1 then %><a href="details.asp" class="bt add noprint">ajouter</a><%End If%></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
        
  <tr class="titlebar noprint">
    <%
If Count > 0 then
%>
       
    <td width="100%"><%=Count%> items trouv&eacute;s - <b>page <%=Page%> de <%=LastPage%></b></td>
    <%
Else
%>
    <td>Aucun item trouv&eacute; </td>
    <%
End If
%>
  
      <td><form method="get">Afficher
        <select name="ItemPerPage" OnChange="this.form.submit()">
          <option selected><%=Session("ItemPerPage")%></option>
		  <option>10</option>
          <option>20</option>
          <option>30</option>
          <option>40</option>
		  <option>50</option>
		  <option>99</option>
		  <option>-</option>
        </select>
            items par page
      </form></td>
    
    <%
If Page <> 1 And LastPage > 1 then
%>
    <td width="20"><a href="<%=currentPage%>?page=<%=Page-1%>" class="bt_arrowleft"></a></td>
    <%
Else
%>
    <td width="20"><a class="bt_arrowleft off"></a></td>
    <%
End If
%>
    <td><strong><%=Page%></strong></td>
<%
If Page <> LastPage And LastPage > 1 then
%>
    <td width="20"><a href="<%=currentPage%>?page=<%=Page+1%>" class="bt_arrowright"></a></td>
    <%
Else
%>
    <td><a class="bt_arrowright off"></a></td>
    <%
End If
%>
  </tr>
      
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <%
For i = 0 to Ubound(Headers,1)
If Session(uniqueListName & "OrderBy")=Headers(i,0) then
	orderImage="&nbsp;<img src='/intranet/images/" & Session(uniqueListName & "Sort") & ".gif' vspace='2' border=0 width=9 height=5>"	
Else
	orderImage=""
End IF
%>
	<th nowrap="nowrap" title="Trier par <%=Headers(i,1)%>" onMouseOver="this.className='over';" onMouseOut="this.className='';" onClick="document.location.href='<%=currentPage%>?OrderBy=<%=Headers(i,0)%>';">
	<%=Headers(i,1)&orderImage%></th>
    <%
Next
%>
     
  </tr>
  <form name=go>
    <tr class="filterbar">
      <%
For i = 0 to Ubound(Headers,1)
	If Headers(i,2) then
		newCombo="<Select name='" & Headers(i,0) & "' OnChange=""document.location.href='" & currentPage & "?" & Headers(i,0) & "=' + document.go." & Headers(i,0) & ".options[document.go." & Headers(i,0) & ".selectedIndex].value;"">" & VbCrLf
		
		'ReadOnly
		If Headers(i,0) = "dealer_id" AND Session("Level") = 3 then
			
			newCombo = ""
			
		Else
			newCombo = newCombo & "<option value=-1>tout</option>" & VbCrLf
			Set rsCombo = Conn.Execute(Headers(i,4))
			Do While Not rsCombo.EOF
				IF Session(uniqueListName&(Headers(i,0)))=CStr(rsCombo(Headers(i,0))) then
					strHTM = "selected"
				Else
					strHTM = ""
				End If
				Execute("cbvalue=" & Headers(i,5))
				newCombo = newCombo & "<option " & strHTM & " value='" & Server.HtmlEncode(rsCombo(Headers(i,0))) & "'>" & cbvalue & "</option>" & VbCrLf
				rsCombo.MoveNext
			Loop
			newCombo = newCombo & "</select>"
		End IF
	Else
		newCombo="&nbsp;"
	End If
	
%>
	  <td><%=newCombo%></td>
      <%
	
Next
%>
       
  </tr>
  </form>
  <%
i = 0
Do While not rs.EOF And rs.AbsolutePage < Page+1
ct = ct + 1
If ct/2 = Cint(ct/2) then
	TRStyle = "normal"
Else
	TRStyle = "high"
End If
Execute("overdue=Cbool(" & overdue_condition & ")")
If overdue then
	TRStyle = "overdue"
End If
'numID = Right(FormatNumber(rs("ID")/100000,5),5)

' Si l'utilisateur est = 3 qui équivaut < utilisateur ReadOnly on lui enlève le click du lien sur l'affichage de la liste des camions
stringForReadOnlyUser = "document.location.href='details.asp?ID=" & rs("ID") & "&r=" & Server.URLEncode(currentpage &"?page="&page) & "'"
If Session("Level") = 3 then stringForReadOnlyUser = "return false;" End If

%>
  <!--<tr OnClick="document.location.href='details.asp?ID=<%=rs("ID")%>&r=<%=Server.URLEncode(currentpage&"?page="&page)%>'" class='<%=TRStyle%>' OnMouseOver="this.className='over';" OnMouseOut="this.className='<%=TRStyle%>';">-->
  <tr OnClick="<%= Response.Write (stringForReadOnlyUser) %>" class='<%=TRStyle%>' OnMouseOver="this.className='over';" OnMouseOut="this.className='<%=TRStyle%>';">
  
    <%
For i = 0 to Ubound(Headers,1)
%>
	<td valign="top" nowrap="nowrap">
      <%
	If Headers(i,3)<>"" then
		Execute("Response.Write " & Headers(i,3))
	Else
		Response.Write rs(Headers(i,0))
	End If
	%>
      &nbsp;</td>
    <%
Next
	rs.MoveNext
Loop

%>
  </tr>
</table>
<%
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <tr class="titlebar noprint">
    <%
If Count > 0 then
%>
       
    <td width="100%"><%=Count%> items trouv&eacute;s - <b>page <%=Page%> de <%=LastPage%></b></td>
    <%
Else
%>
    <td>Aucun item trouv&eacute; </td>
    <%
End If
%>
  
      <td><form method="get">Afficher
        <select name="ItemPerPage" OnChange="this.form.submit()">
          <option selected><%=Session("ItemPerPage")%></option>
		  <option>10</option>
          <option>20</option>
          <option>30</option>
          <option>40</option>
		  <option>50</option>
		  <option>99</option>
		  <option>-</option>
        </select>
            items par page
      </form></td>
    
    <%
If Page <> 1 And LastPage > 1 then
%>
    <td width="20"><a href="<%=currentPage%>?page=<%=Page-1%>" class="bt_arrowleft"></a></td>
    <%
Else
%>
    <td width="20"><a class="bt_arrowleft off"></a></td>
    <%
End If
%>
    <td><strong><%=Page%></strong></td>
<%
If Page <> LastPage And LastPage > 1 then
%>
    <td width="20"><a href="<%=currentPage%>?page=<%=Page+1%>" class="bt_arrowright"></a></td>
    <%
Else
%>
    <td><a class="bt_arrowright off"></a></td>
    <%
End If
%>
  </tr>
</table>

