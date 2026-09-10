<%      

   '--------------------------------------------------------
   ' utilities for opening and closing connections to the
   ' expedition database
   '--------------------------------------------------------

   '--------------------------------------------------------
  
   Sub open_db_conn

    On Error Resume Next
   
    Set wptConn = Server.CreateObject("ADODB.Connection")
    wptConn.CommandTimeout = 30
    wptConn.ConnectionTimeout = 15
    'Joe Gomez - Added explicit DB Connection for Future
	'Proofing when upgrading server to 2012R2
	'3/28/2017
	'wptConn.Provider = Application("providerStr")
    'wptConn.Open (Application("DSN"))
    wptConn.Open "PROVIDER=SQLOLEDB;DATA SOURCE=perseus;UID=expddba;PWD=password;DATABASE=expd"

    
    If (IsNull(wptConn)) Then
       Response.Write ("<b>Unable to establish a connection to the Waypoint database.</b><br>")
       Session.Abandon
    Else
       ' check for connection errors, but ignore these warnings:
       '  [Microsoft][ODBC SQL Server Driver][SQL Server]Changed database context to 'expd_test'.
       '  [Microsoft][ODBC SQL Server Driver][SQL Server]Changed language setting to us_english
    
       If wptConn.Errors.Count > 0 Then
          For intLoop = 0 To wptConn.Errors.Count - 1
             If Not ((InStr (wptConn.Errors(intLoop).Description,"Changed database context") > 0) or _
                     (InStr (wptConn.Errors(intLoop).Description,"Changed language setting") > 0)) Then
                Response.Write "Error Number: " & wptConn.Errors(intLoop).Number
                Response.Write " - " & wptConn.Errors(intLoop).Description & "<P>"
                Session.Abandon
             End If
          Next          
       End If
       wptConn.Errors.Clear
    End If   

   End Sub 'open_db_conn

   '--------------------------------------------------------
  
   Sub close_db_conn

    wptConn.Close
    Set wptConn = Nothing

   End Sub 'close_db_conn

   '--------------------------------------------------------

%>