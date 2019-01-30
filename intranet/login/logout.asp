<%
Response.Cookies("userID").Expires = Date - 1000
Response.Cookies("username").Expires = Date - 1000
Response.Cookies("fullname").Expires = Date - 1000
Session.Abandon
Response.Redirect("/intranet/")
%>