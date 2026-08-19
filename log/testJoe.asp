<%
Response.Write "LOGON_USER: " & Request.ServerVariables("LOGON_USER") & "<br>"
Response.Write "REMOTE_USER: " & Request.ServerVariables("REMOTE_USER") & "<br>"
Response.Write "AUTH_USER: " & Request.ServerVariables("AUTH_USER") & "<br>"
Response.Write "<br>"
'Show all server variables
For Each Item In Request.ServerVariables
Response.Write Item & " = " & Request.ServerVariables(Item) & "<br>"
Next
%>