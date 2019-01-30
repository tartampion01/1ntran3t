
<%


'0 is field name
'1 is display name
'2 is true or flase to show a combo list
'3 is a string of the asp code to show value in the list (optional)
'4 is query string to retrive combo data. Use AS to make ID same as FieldName id needed
'5 is a string of the asp code to show value in the combo (using rsCombo)

If Request.QueryString("page") = "" then
   Page = 1 'We're on the first page
Else
   Page = Request.QueryString("page")
End If

Page=Cint(Page)

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("cstring")

'Explicitly Create a recordset object
Set rs = Server.CreateObject("ADODB.Recordset")

'Set the cursor location property
rs.CursorLocation = 3 'adUseClient

If Request("ItemPerPage") = "" then
	Response.Cookies("ItemPerPage")=10
Else
	Response.Cookies("ItemPerPage")=Request("ItemPerPage")
	Response.Cookies("ItemPerPage").Expires = Date + 30	
End If

'Set the cache size = to the # of records/page
rs.CacheSize = Request("ItemPerPage")

If Request("OrderBy") <> "" then
	If Session(tableName & "Sort")="ASC" then
		Session(tableName & "Sort")="DESC"
	Else
		Session(tableName & "Sort")="ASC"
	End If
	Session(tableName & "Orderby")=Request("Orderby")
End If

strSQl = strSQL & "WHERE "
For i = 0 to Ubound(Headers,1)
	If Request.QueryString(Headers(i,0))<>"" then
		If Request.QueryString(Headers(i,0))="-1" then
			Session(Headers(i,0))=""
		Else
			Session(Headers(i,0))=Request.QueryString(Headers(i,0))
		End IF
	End If
	IF Session(Headers(i,0)) <> "" and Headers(i,2) then
		strSQl = strSQL & Headers(i,0) & "=" & Session(Headers(i,0)) & "   AND "
	End If
Next

'Search Feature
IF Request.form("Search")<>"" then
	Session(tableName & "Search")= Trim(Request.form("Search"))
End If
'Reset
If Request.QueryString = "reset" then Session(tableName & "Search")=""

SearchArray = Split(Session(tableName & "Search")," ")
SearchItemsArray = Split(SearchItems,",")

If Session(tableName & "Search") <> "" then
	strSQL = strSQL & "(" 
	For i = 0 to Ubound(SearchArray,1)
		For j = 0 to Ubound(SearchItemsArray,1)
			strSQL = strSQL &  SearchItemsArray(j) & " Like '%" & SearchArray(i) & "%'    OR "
		Next
	Next
	strSQL = Left(strSQL,Len(strSQL)-6)&")"
Else
	strSQL = Left(strSQL,Len(strSQL)-6)
End If


IF Session(tableName & "OrderBy") <> "" then
	strSQl = strSQL & " Order by " & Session(tableName & "OrderBy") & " " & Session(tableName & "Sort")
End If

strSQl = strSQL & ";"

rs.Open strSQL, Conn

count = rs.RecordCount

If count <> 0 then
	rs.MoveFirst
	rs.PageSize = Request("ItemPerPage")

	'Get the max number of pages
	LastPage = rs.PageCount

	'Set the absolute page
	rs.AbsolutePage = Page
End If

%>
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td class="title"><%=Title%></td>
<%If SearchItems<>"" then%>
    <form method="post" action="default.asp">
	  <td valign="bottom">
        <table border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>Quick Search:</td>
            <td>
              <input type="text" name="Search" class="field" value="<%=Session(tableName & "Search")%>">
            </td>
            <td>
              <input type="image" border="0" name="Submit" src="/manager/images/ok.gif" width="19" height="18" alt="Effectuer la recherche">
              <a href="./?reset"><img src="/manager/images/cancel.gif" width="18" height="18" border="0" alt="Annuler la recherche"></a>
            </td>
          </tr>
        </table>
        
      </td>
	</form>
<%End If%>
    <td valign="middle" align="right"><a href="details.asp">+ Add New</a></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
        
  <tr>
    <%
If Count > 0 then
%>
       
    <td class="bargrey" width="100%">&nbsp;<%=Count%> items found - <b>page <%=Page%> of <%=LastPage%></b></td>
    <%
Else
%>
    <td class="bargrey" width="100%">&nbsp;No item found</td>
    <%
End If
%>
    <form method="get">
      <td nowrap class="bargrey">&nbsp;Show
<select name="ItemPerPage" class="field" OnChange="this.form.submit()">
          <option selected><%=Request("ItemPerPage")%></option>
		  <option>10</option>
          <option>20</option>
          <option>30</option>
          <option>40</option>
		  <option>50</option>
		  <option>99</option>
        </select>
            items per page&nbsp;
          </td>
    </form>
    <%
If Page <> 1 And LastPage > 1 then
%>
    <td onMouseOver="this.className='barblue';" onMouseOut="this.className='barblack';" class="barblack" width="20"><a href="default.asp?page=<%=Page-1%>"><img src="/manager/images/left.gif" width="20" height="20" border="0" alt="Previous"></a></td>
    <%
Else
%>
    <td class="bargrey" width="20"><img src="/manager/images/left.gif" width="20" height="20" border="0"></td>
    <%
End If
%>
    <td class="bargrey"><b>&nbsp;<%=Page%>&nbsp;</b></td>
<%
If Page <> LastPage And LastPage > 1 then
%>
    <td onMouseOver="this.className='barblue';" onMouseOut="this.className='barblack';" class="barblack" width="20"><a href="default.asp?page=<%=Page+1%>"><img src="/manager/images/right.gif" width="20" height="20" border="0" alt="Next"></a></td>
    <%
Else
%>
    <td class="bargrey" width="20"><img src="/manager/images/right.gif" width="20" height="20" border="0"></td>
    <%
End If
%>
      </tr>
      
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="2">
  <tr>
    <%
For i = 0 to Ubound(Headers,1)
If Session(tableName & "OrderBy")=Headers(i,0) then
	orderImage="&nbsp;<img src='/manager/images/" & Session(tableName & "Sort") & ".gif' border=0 width=8 height=7>"	
Else
	orderImage=""
End IF
%>
	<th onMouseOver="this.className='over';" onMouseOut="this.className='';" onClick="this.children.tags('A')[0].click();" class="thnormal" nowrap>
	<a href="default.asp?OrderBy=<%=Headers(i,0)%>"><%=Headers(i,1)&orderImage%></a></th>
    <%
Next
%>
     
  </tr>
  <form name=go>
    <tr>
      <%
For i = 0 to Ubound(Headers,1)
	If Headers(i,2) then
		newCombo="<Select name='" & Headers(i,0) & "' class='smallfield' OnChange=""document.location.href='default.asp?" & Headers(i,0) & "=' + document.go." & Headers(i,0) & ".options[document.go." & Headers(i,0) & ".selectedIndex].value;"">" & VbCrLf
		newCombo = newCombo & "<option value=-1>Show All</option>" & VbCrLf
		Set rsCombo = Conn.Execute(Headers(i,4))
		Do While Not rsCombo.EOF
			IF Session(Headers(i,0))=CStr(rsCombo(Headers(i,0))) then
				strHTM = "selected"
			Else
				strHTM = ""
			End If
			Execute("cbvalue=" & Headers(i,5))
			newCombo = newCombo & "<option " & strHTM & " value='" & Server.HtmlEncode(rsCombo(Headers(i,0))) & "'>" & cbvalue & "</option>" & VbCrLf
			rsCombo.MoveNext
		Loop
		newCombo = newCombo & "</select>"
	Else
		newCombo="&nbsp;"
	End If
	
%>
	  <td class='bargrey'><%=newCombo%></td>
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
numID = Right(FormatNumber(rs("ID")/100000,5),5)

%>
  <tr OnClick="document.location.href='details.asp?ID=<%=rs("ID")%>'" class='<%=TRStyle%>' OnMouseOver="this.className='over';" OnMouseOut="this.className='<%=TRStyle%>';">
    <%
For i = 0 to Ubound(Headers,1)
%>
	<td valign="top">
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
        
  <tr>
    <%
If Count > 0 then
%>
       
    <td class="bargrey" width="100%">&nbsp;<%=Count%> items found - <b>page <%=Page%> of <%=LastPage%></b></td>
    <%
Else
%>
    <td class="bargrey" width="100%">&nbsp;No item found</td>
    <%
End If
%>
    <form method="get">
      <td nowrap class="bargrey">&nbsp;Show
<select name="ItemPerPage" class="field" onChange="this.form.submit()">
          <option selected><%=Request("ItemPerPage")%></option>
		  <option>10</option>
          <option>20</option>
          <option>30</option>
          <option>40</option>
		  <option>50</option>
		  <option>99</option>
        </select>
            items per page&nbsp;
          </td>
    </form>
    <%
If Page <> 1 And LastPage > 1 then
%>
    <td onMouseOver="this.className='barblue';" onMouseOut="this.className='barblack';" class="barblack" width="20"><a href="default.asp?page=<%=Page-1%>"><img src="/manager/images/left.gif" width="20" height="20" border="0"></a></td>
    <%
Else
%>
    <td class="bargrey" width="20"><img src="/manager/images/left.gif" width="20" height="20" border="0"></td>
    <%
End If
%>
    <td class="bargrey"><b>&nbsp;<%=Page%>&nbsp;</b></td>
<%
If Page <> LastPage And LastPage > 1 then
%>
    <td onMouseOver="this.className='barblue';" onMouseOut="this.className='barblack';" class="barblack" width="20"><a href="default.asp?page=<%=Page+1%>"><img src="/manager/images/right.gif" width="20" height="20" border="0"></a></td>
    <%
Else
%>
    <td class="bargrey" width="20"><img src="/manager/images/right.gif" width="20" height="20" border="0"></td>
    <%
End If
%>
      </tr>
      
</table>
