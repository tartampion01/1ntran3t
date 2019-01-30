<%@EnableSessionState=False%>
<!--#include virtual="/intranet/includes/upload/_upload.asp" -->

<%
Response.Expires = -10000
Server.ScriptTimeOut = 1000


Function ResizeImage(inpath,outpath,MaxWidth,MaxHeight,params,saveparams)
	'Reduce and compress large image
	Dim Image,OrigImage,rect

	Set Image = Server.CreateObject("ImageGlue5.Canvas")
	Set OrigImage = Server.CreateObject("ImageGlue5.Graphic")
	Set rect = Server.CreateObject("ImageGlue5.XRect") 
	
	OrigImage.SetFile inpath
	rect.String = OrigImage(1).Rectangle 
	
	If MaxWidth then Image.Width = MaxWidth
	If MaxHeight then Image.Height = MaxHeight

	Image.DrawFile inpath,params
	
	Image.SaveAs outpath,saveparams
	
	Set Image = Nothing
End Function


on error resume next


Set Form = New ASPForm
Form.SizeLimit = &HA00000
Form.UploadID = Request.QueryString("UploadID")
If Form.State = fsCompletted Then 'Completted
  'was the Form successfully received?
  if Form.State = 0 then
	  For Each File In Form.Files.Items
		  File.SaveAs Server.Mappath(Form("vpath") &  form("filename") & ".pdf" )
	  Next
  End If

ElseIf Form.State > 10 then
  Const fsSizeLimit = &HD
  Select case Form.State
		case fsSizeLimit: response.write  "<br><Font Color=red>Source form size (" & Form.TotalBytes & "B) exceeds form limit (" & Form.SizeLimit & "B)</Font><br>"
		case else response.write "<br><Font Color=red>Some form error.</Font><br>"
  end Select
  
  response.end
  
End If'Form.State = 0 then


returnaddress = Form("return") & "?ID=" & Form("ID")

Response.Redirect(returnaddress)

%>