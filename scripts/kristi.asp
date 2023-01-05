<!--#include file="secrets.NoGit"-->
<% 
'inspired by: 
'https://stackoverflow.com/questions/10437309/classic-asp-set-cookies-in-addition-to-session
'https://learn.microsoft.com/en-us/previous-versions/iis/6.0-sdk/ms524771(v=vs.90)

'Cookie business:
CookieName = "KristiLakeNatureTrail"
CookieValue = Request.Cookies(CookieName)
' set a simple cookie if we don't already have one
If CookieValue = "" Then 
   Response.Cookies(CookieName) = CreateGUID()
   Response.Cookies(CookieName).expires = DateAdd( "yyyy", 5, Date ) '5 years seems good enough
   CookieValue = Request.Cookies(CookieName)
End If   
							   
UrlValue = UrlOrGivenParam() 

Set cnn = Server.CreateObject("ADODB.Connection") 

'DEBUG ErrorHandling
' Turn off error Handling
'On Error Resume Next
cnn.Open ConnectKristiLog() 
'DEBUG: Code here that you want to catch errors from Error Handler
'If Err.Number <> 0 Then
'   Response.Write("hmm")
'   Response.Write("Error:[" & Err.Number & "]")
'   Response.Write("Description:[" & Err.Description & "]")
'   ' Error Occurred - Trap it
'   On Error Goto 0 ' Turn error handling back on for errors in your handling block
'   ' Code to cope with the error here
'End If
'On Error Goto 0 ' Reset error handling.

'Open a recordset using the Open method and use the connection established by the Connection object. 
strSQL = "INSERT INTO [dbo].[LogKristi] ([IP] ,[URL] ,[HitDateTime] ,[CookieID]) SELECT '" & GetIP() & "', '" & UrlValue & "',GETUTCDATE(),'" & CookieValue & "'" 
set xResult=cnn.Execute(strSQL)
							
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
  resultValue = SanitizeSQL(Request.QueryString("u"))
  If resultValue = "" Then resultValue = Request.ServerVariables("URL")
  UrlOrGivenParam = resultValue
End Function

'very crude - better ways to do likey
Function SanitizeSQL(givenText)
result = ""
'only allow alphas numbers and slash
allowed = "abcdefghijklmnopqrstuvwxyz/0123456789"
For TheGivenIndex = 1 to Len(givenText)
 thisAllowed = false
 For TheAllowedIndex = 1 to Len(allowed)
    If lcase(mid(givenText,TheGivenIndex,1)) = mid(allowed, TheAllowedIndex, 1) Then
	   thisAllowed = true
	   exit For
	End If
 Next
 If thisAllowed Then
    result = result & mid(givenText,TheGivenIndex,1)
 End If
 SanitizeSQL = result
Next

End Function

Function TopOrParam 
  resultValue = Request.QueryString("t")
  If resultValue = "" Then resultValue = 50 '" TOP 50" 'Request.ServerVariables("URL")
  TopOrParam = " TOP " & resultValue * 1 & " " 'resultValue
End Function
%> 
<script type="text/javascript">
// Your cookie is <%=CookieValue%> - feel free to delete it - it is just for this site to track unique hits...
</script>