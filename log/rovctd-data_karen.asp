<!--#include file="ado1v_karen.inc"-->
<%Response.Buffer = False%>
<%Server.ScriptTimeout = 300%>

<%

'*** TO FIX ***'

' very large queries (start&end date) cause users browser to crash (and slow down fog)
' limit number of records to 2000 (or some other safe amount)

' since each row of data is related to the next, and there are an arbitrary number of
' rows in dive, it makes no sense to have the user page through the records a fixed number
' at a time.

' when the user queries accross multiple days, they should get the results one day (or
' possibly dive) at a time.  Actually dive would make the most sense.

'*** ****** ***'

'Author: Dan Wilkin

'Date: 9/24/1999
'MBARI ROVCTD Database Migration Project
'
'This active server page does the following:
'	1.  Runs a query against a SQL Server database to get rovctd data
'	2.  Writes the data as a comma delimited text file, an html table,
'       or as an Excel file.
'
'Languages Used:
'	VBscript
'
'This asp script is called from default.htm in the queries section of the expd web.
'It receives the following input:
'
' Note: example values for each input are included.
'
' http://fog.shore.mbari.org/expd/queries/rovctd-data_karen.asp?
' year=2000&
' yday=234&
' dive=1024&
' date_type=srt_end&
' srt_mm=03&
' srt_dd=04&
' srt_yyyy=1998&
' srt_hh=18&
' srt_min=34&
' srt_ss=02&
' end_mm=11&
' end_dd=17&
' end_yyyy=2000&
' end_hh=02&
' end_min=08&
' end_ss=57&
' pltfrm=tibr&
' return_type=comma&
' B3=Submit

'There are three radio button sets:
'1. date_type (yyyyddd,dive,srt_end) - The value indicates which date-time
'   inputs should be used in the database query.
'2. pltfrm (vnta,tibr,docr,mini) - The value indicates which ROV was used, Ventana,
'   Tiburon Dock Ricketts and MiniROV.
'3. return_type (comma,table,excel) - The value indicates whether the user
'   wants the data returned in comma deliminated, html table,
'   or MS Excel format.

' System Design:
'
' 1. Read the value of the date_type input.
' 2. If the date_type is dive, query the dive table to find out the StartDTG
'    and EndDTG.  If the date_type is srt_end, construct StartDTG and EndDTG
'    strings from the inputs.  If the date_type is yyyyddd, use the inputs
'    "year" and "yday" from the QueryString.
' 3. Call the appropriate stored procedure.
' 4. Display the results in comma delimited, html table, or MS Excel format.

Dim varTitle			' variant that holds the title of the webpage
Dim varPltfrm			' holds the platform (tibr/vnta/docr/mini)
Dim varROVName			' holds the rov name (Tiburon/Ventana/Doc Ricketts/MiniROV)
Dim varStartDTG			' start date-time group
Dim varEndDTG			' end date-time group
Dim cmdRovctd			' ADO command object
Dim conRovctd			' ADO connection object
Dim rstRovctd			' ADO recordset object
Dim rstDive				' ADO recordset object for accessing the dive table
Dim varDiveSQL			' SQL string for accessing the dive table
Dim longFldCount		' count of the number of fields in the recordset
Dim n					' counting variable for for loops
Dim FieldArray()		' array of field names
Dim quitnow				' 1 - indicates that further processing should cease
Dim divewarning			' 1 - indicates that the dive times must be entered
Dim varDelimiter		' vbtab or comma
Dim varDatabase			' database name
Dim varDSN				' data source name
Dim varConnectionString	' database connection string

' error handling
Dim varErrNum			' holds the number of an error that occurs
Dim varErrDes			' holds the error description

' tell VBScript that we are handling errors in our code
On Error Resume Next

' set the ScriptTimeout to a large number, the user might try to retrieve a large
' amount of data.  The default timeout is 90 seconds.
Server.ScriptTimeout = 600 'seconds (10 minutes)

Application("providerStr") = "sqloledb"
Application("connectionStr") = "Server=fog;Database=expd_ks;UID=expddba;PWD=password;"
'Application("connectionStr") = "Server=perseus;Database=expd;UID=expddba;PWD=password;"

' create a connection object
Set conRovctd = Server.CreateObject("ADODB.Connection")

' set the provider
conRovctd.Provider = Application("providerStr")

' open a connection to the expedition database
conRovctd.Open Application("connectionStr")

' set quitnow and divewarning to false
quitnow = 0
divewarning = 0

varErrNum = Err.number
varErrDes = Err.description

if varErrNum <> 0 then
    Call WriteHTMLHeader ()
    Response.Write ("<center>A connection to database could not be established.</center>")
    Call WriteHTMLFooter ()
    quitnow = 1
end if

' if a connection to the database was established, continue
if (quitnow = 0) then

    ' create a command object and set its connection to the expedition database connection
    Set cmdRovctd = Server.CreateObject("ADODB.Command")
    Set cmdRovctd.ActiveConnection = conRovctd
    ' set the command type to stored procedure
    cmdRovctd.CommandType = adCmdStoredProc

' Response.Write ("<br>FOO_1<br>")  

    ' chose which stored procedure to call based on the date_type input value
    Select Case (Request.QueryString("date_type"))

        Case "yyyyddd"
            ' use the getRovCtdDY stored procedure to query the database
            cmdRovctd.CommandText = "getRovCtdDY"

            ' set the values of the parameters
            cmdRovctd.Parameters.Refresh
            cmdRovctd.Parameters(1).Value = Request.QueryString("year")
            cmdRovctd.Parameters(2).Value = Request.QueryString("yday")
            cmdRovctd.Parameters(3).Value = Request.QueryString("pltfrm")
        
        Case "dive"
            ' first we query the dive table to get the start and end of the dive        
            Set rstDive = Server.CreateObject("ADODB.Recordset")
            varDiveSQL = "SELECT DiveStartDtg, DiveEndDtg FROM dive WHERE DiveNumber = " & _
                         Request.QueryString("dive") & " AND RovName = '" & _
                         Request.QueryString("pltfrm") & "'"

            ' open the recordset
            rstDive.Open varDiveSQL, conRovctd, adOpenForwardOnly, adLockOptimistic, adCmdText

            ' give the user feedback if the requested dive was not found in the dive table        
            If (rstDive.BOF and rstDive.EOF) then
                Call WriteHTMLHeader ()

                Response.Write ("<center><b>Sorry, there is no dive with Dive Number: " & _
                                Request.QueryString("dive") & " and Rov Name: " & _
                                Request.QueryString("pltfrm") & " in the dive table.</b></center>")
                Call WriteHTMLFooter ()

                ' set the quitnow flag, so we don't do any more processing
                quitnow = 1
            Else
                ' the dive table query was successful, so get the start and end datetimes
                varStartDTG = rstDive("DiveStartDtg")
                varEndDTG = rstDive("DiveEndDtg")
                
                if (IsNull(varStartDTG) or isnull(varEndDTG)) then
                    Call WriteHTMLHeader ()
                end if
                
                if IsNull(varStartDTG) then
                    Response.Write ("<center><b>DiveStartDtg has not been entered for this dive yet.</b></center>")
                    quitnow = 1
                    divewarning = 1
                end if

                if isnull(varEndDTG) then
                    Response.Write ("<center><b>DiveEndDtg has not been entered for this dive yet.</b></center>")
                    quitnow = 1
                    divewarning = 1
                end if

                if divewarning = 1 then
                    Response.Write ("<br><center><b>The dive start and end times must be entered before " & _
                                    "you can retrieve ROVCTD data by dive number.</b><br><br>")
                    Response.Write ("<A HREF=" & Chr(34) & "../log/dive_karen.asp?RovName=" & Request.QueryString("pltfrm") & _
                                    "&DiveNumber=" & Request.QueryString("dive") & Chr(34) & _
                                    ">Edit Dive " & Request.QueryString("dive") & "</A></center>")
                    Call WriteHTMLFooter ()
                end if

                ' we're done getting the dive info, so close the recordset
                rstDive.Close

                cmdRovctd.CommandText = "getRovCtd"

                cmdRovctd.Parameters.Refresh
                cmdRovctd.Parameters(1).Value = varStartDTG
                cmdRovctd.Parameters(2).Value = varEndDTG
                cmdRovctd.Parameters(3).Value = Request.QueryString("pltfrm")
            End if '(rstDive.BOF and rstDive.EOF)

        Case "srt_end"  
            cmdRovctd.CommandText = "getRovCtd"

            ' create the varStartDTG from the srt inputs
            varStartDTG = ""

            if Request.QueryString("srt_mm") = "" then
                varStartDTG = varStartDTG & "01" & "/"
            else
                varStartDTG = varStartDTG & Request.QueryString("srt_mm") & "/"
            end if
        
            if Request.QueryString("srt_dd") = "" then
                varStartDTG = varStartDTG & "01" & "/"
            else
                varStartDTG = varStartDTG & Request.QueryString("srt_dd") & "/"
            end if
        
            varStartDTG = varStartDTG & Request.QueryString("srt_yyyy") & " "

            if Request.QueryString("srt_hh") = "" then
                varStartDTG = varStartDTG & "00" & ":"
            else
                varStartDTG = varStartDTG & Request.QueryString("srt_hh") & ":"
            end if
        
            if Request.QueryString("srt_min") = "" then
                varStartDTG = varStartDTG & "00" & ":"
            else
                varStartDTG = varStartDTG & Request.QueryString("srt_min") & ":"
            end if
        
            if Request.QueryString("srt_ss") = "" then
                varStartDTG = varStartDTG & "00"
            else
                varStartDTG = varStartDTG & Request.QueryString("srt_ss")
            end if

            ' create the varEndDTG from the end inputs
            varEndDTG = ""

            if Request.QueryString("end_mm") = "" then
                varEndDTG = varEndDTG & "12" & "/"
            else
                varEndDTG = varEndDTG & Request.QueryString("end_mm") & "/"
            end if
        
            if Request.QueryString("end_dd") = "" then
                varEndDTG = varEndDTG & "31" & "/"
            else
                varEndDTG = varEndDTG & Request.QueryString("end_dd") & "/"
            end if
            
            varEndDTG = varEndDTG & Request.QueryString("end_yyyy") & " "

            if Request.QueryString("end_hh") = "" then
                varEndDTG = varEndDTG & "23" & ":"
            else
                varEndDTG = varEndDTG & Request.QueryString("end_hh") & ":"
            end if
        
            if Request.QueryString("end_min") = "" then
                varEndDTG = varEndDTG & "59" & ":"
            else
                varEndDTG = varEndDTG & Request.QueryString("end_min") & ":"
            end if
        
            if Request.QueryString("end_ss") = "" then
                varEndDTG = varEndDTG & "59"
            else
                varEndDTG = varEndDTG & Request.QueryString("end_ss")
            end if

            if ((not (IsDate(varStartDTG))) or (not (IsDate(varEndDTG)))) then
                Call WriteHTMLHeader()
        
                if (not (IsDate(varStartDTG))) then
                    Response.Write ("<center><b>Invalid date: " & varStartDTG & "</b></center><br>")
                end if
        
                if (not (IsDate(varEndDTG))) then
                    Response.Write ("<center><b>Invalid date: " & varEndDTG & "</b></center><br>")
                end if
                
                Call WriteHTMLFooter()
                quitnow = 1       
            else
                cmdRovctd.Parameters.Refresh
                cmdRovctd.Parameters(1).Value = varStartDTG
                cmdRovctd.Parameters(2).Value = varEndDTG
                cmdRovctd.Parameters(3).Value = Request.QueryString("pltfrm")  

            end if

    
        Case Else
            ' the user failed to make a date_type selection (this is not possible on the form, but
            ' this ASP might be called in other ways
            Call WriteHTMLHeader ()
            Response.Write ("<center><b>You must choose a date_type of either 'yyyyddd', 'dive', or 'srt_end.'</b></center>")
            Call WriteHTMLFooter ()
            quitnow = 1        
    End Select

        Response.Write (cmdRovctd)


end if ' (quitnow = 0)

' make sure the above steps succeeded before proceeding
if (quitnow = 0) then
' Response.Write ("<br>FOO_2<br>")
    Response.Write (cmdRovctd)

    ' create the recordset by executing the stored procedure
    Set rstRovctd = cmdRovctd.Execute

    ' make sure that some data was returned, give the use a message if none was returned.
    If (rstRovctd.BOF and rstRovctd.EOF) then
        Call WriteHTMLHeader ()
        Response.Write ("<center><b>Sorry, no data was found in the RovCtd table for your query.</b></center>")
        Call WriteHTMLFooter ()
    Else
        '***** This is the first content written to the user's browser (except for errors) *****'
        ' choose which type of content and write the appropriate response
        
        if Request.QueryString("return_type") = "table" then
            Response.Write ("<html>")
            Response.Write ("<head>")

            varPltfrm = Request.QueryString("pltfrm")
            Response.Write ("varROVName")

            ' create the title from the parameters
            If (varPltfrm = "vnta") then
                varROVName = "Ventana"
            else
                if (varPltfrm = "tibr") then
                    varROVName = "Tiburon"
                else
                    if (varPltfrm = "docr") then
                        varROVName = "Doc Ricketts"
                    else
                        if (varPltfrm = "mini") then
                            varROVName = "MiniROV"
                        else
                            varROVName = varPltfrm
                        end if
                    end if
                end if
            end if


            Response.Write ("varROVName")

            varTitle = "ROV CTD Data Archive for " & varROVName

            ' write the title
            Response.Write ("<title>" & varTitle & "</title>")
            
            Response.Write ("</head>")
            Response.Write ("<body bgcolor=#ffffff>")
        else
            if Request.QueryString("return_type") = "comma" then
                Response.ContentType = "text/plain"
            else    
                if Request.QueryString("return_type") = "excel" then
                    Response.ContentType = "application/vnd.ms-excel"
                else
                    Call WriteHTMLHeader ()
                    Response.Write ("<center><b>You must choose a return_type of either 'comma', 'table', or 'excel.'</b></center>")
                    Call WriteHTMLFooter ()
                end if 'Request.QueryString("return_type") = "excel"
            end if 'Request.QueryString("return_type") = "comma"
        end if 'Request.QueryString("return_type") = "table"
    
        'Now output the data.
        if (Request.QueryString("return_type") = "table") then

            %>
            <table align="left" border="3" width="100%" cellspacing="1" cellpadding="3" bordercolorlight="#008080" bordercolordark="#C0C0C0">
            <%
        
            ' Get the field names from the recordset
            longFldCount = rstRovctd.Fields.Count
            ReDim FieldArray(longFldCount)
    
            Response.Write ("<tr nowrap><b>")
            For n=0 to longFldCount-1
                FieldArray(n) = rstRovctd.Fields(n).Name
                Response.Write ("<th nowrap>")
                Response.Write (rstRovctd.Fields(n).Name)
                Response.Write ("</th>")
            Next
            Response.Write ("</b></tr>")

            Do While Not rstRovctd.EOF
                Response.Write ("<tr nowrap>")
                for n=0 to longFldCount-1
                    Response.Write ("<td nowrap>")
                    if (IsNull(rstRovCtd(FieldArray(n)))) then
                        Response.Write ("&nbsp;")
                    else
                        Response.Write (rstRovCtd(FieldArray(n)))
                    end if
                    Response.Write ("</td>")
                next
                Response.Write("</tr>")      
                rstRovctd.MoveNext
            Loop
   
            rstRovctd.Close      
            Response.Write ("</table>")
        
            Call WriteHTMLFooter ()
        else
            if (Request.QueryString("return_type") = "comma") or (Request.QueryString("return_type") = "excel") then
            
                ' Set the delimiter (vbTab or comma)
                if (Request.QueryString("return_type") = "comma") then
                    varDelimiter = ","
                else
                ' the return_type is excel, so use tab as the delimiter
                    varDelimiter = vbTab
                end if
            
                ' Get the number of field names from the recordset
                longFldCount = rstRovctd.Fields.Count
                     
                ReDim FieldArray(longFldCount)
                
                ' put the field names into an array and also output them
                For n=0 to longFldCount-2
                    FieldArray(n) = rstRovctd.Fields(n).Name
                    Response.Write (rstRovctd.Fields(n).Name & varDelimiter)
                Next
                FieldArray(longFldCount-1) = rstRovctd.Fields(longFldCount-1).Name
                Response.Write (rstRovctd.Fields(longFldCount-1).Name)
                Response.Write (vbCrLf)                

                ' now loop through the records and write out the data
                Do While Not rstRovctd.EOF
                    for n=0 to longFldCount-2 
                        Response.Write (rstRovCtd(FieldArray(n)) & varDelimiter)
                    next
                    Response.Write (rstRovCtd(FieldArray(longFldCount-1)))
                    Response.Write(vbCrLf)
                    rstRovctd.MoveNext
                Loop
   
                rstRovctd.Close
                    
            end if 'Request.QueryString("return_type") = "comma"
        end if '(Request.QueryString("return_type") = "table") or (Request.QueryString("return_type") = "excel")
    end if '(rstRovctd.BOF and rstRovctd.EOF)
end if '(quitnow = 0)

' close the connection to the database	
conRovctd.Close

' subroutine that writes an html header
Sub WriteHTMLHeader ()
    Response.Write ("<html>")
    Response.Write ("<head>")
    Response.Write ("<title>ROV CTD Data Archive</title>")
    Response.Write ("</head>")
    Response.Write ("<body bgcolor=#ffffff>")
    Response.Write ("<img src=" & Chr(34) & "../images/mbarilogo-120_sh.gif" & Chr(34) & _
                    "border=" & Chr(34) & "0" & Chr(34) & "WIDTH=" & Chr(34) & "120" & _
                    Chr(34) & "HEIGHT=" & Chr(34) & "70" & Chr(34) & "><br><br>")
End Sub

' subroutine that writes an html footer
Sub WriteHTMLFooter ()
    Response.Write ("</body>")
    Response.Write ("</html>")
End Sub
%>