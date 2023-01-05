<!--#include file="secrets.NoGit"-->
<% 
Set cnn = Server.CreateObject("ADODB.Connection") 
cnn.Open ConnectKristiLog() 
response.write("<table><tr><td>ID</td><td>IP</td><td>URL</td><td>Hit</td><td>Cookie</td></tr>")
strSQL = "SELECT " & TopOrParam() & "CONCAT('<tr><td>',ID,'</td><td>',IP,'</td><td>',URL,'</td><td>',DATEADD(hh,-6,HitDateTime),'</td><td>',CookieID,'</td></tr>') FROM LogKristi ORDER BY ID DESC" 
set xResult=cnn.Execute(strSQL)

'Cycle through record set and display the results and increment record position with MoveNext method. 
 Do Until xResult.EOF 
    Response.Write xResult(0) '& "<br/>" ' objFirstName & " " & objLastName & "<BR>" 
    xResult.MoveNext
Loop 
response.write("</table>")
cnn.Close
   
'Lifted from here: https://stackoverflow.com/questions/21131587/how-to-define-the-ip-address-using-classic-asp  
Function GetIP()
    Dim strIP 
	strIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
    If strIP = "" Then strIP = Request.ServerVariables("REMOTE_ADDR")
    GetIP = strIP
End Function

'Lifted from here: https://stackoverflow.com/questions/968756/how-to-generate-a-guid-in-vbscript
Function CreateGUID
  Dim TypeLib
  Set TypeLib = CreateObject("Scriptlet.TypeLib")
  CreateGUID = Mid(TypeLib.Guid, 2, 36)
End Function
  
Function UrlOrGivenParam  
  resultValue = Request.QueryString("u")
  If resultValue = "" Then resultValue = Request.ServerVariables("URL")
  UrlOrGivenParam = resultValue
End Function

Function TopOrParam 
  resultValue = Request.QueryString("t")
  If resultValue = "" Then resultValue = 50 '" TOP 50" 'Request.ServerVariables("URL")
  TopOrParam = " TOP " & resultValue * 1 & " " 'resultValue
End Function
%> 