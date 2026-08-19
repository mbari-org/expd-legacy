<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>

<title>VARS video annotations</title>
</head>

<body bgcolor="#FFFFFF">

<!--rewritten for VARS on the Western Flyer August 2005 by Jenny Paduan-->

  <%  
  '''''''''''''''''''''''''''''''''''''
  ' Variable and Object Declarations '
  ''''''''''''''''''''''''''''''''''''

  'stores the ship name from the http request and convert to VARS version
  ' Dim strShipName
  ' Dim strVarsShipName
  
  'stores ROV name from the http request
  ' Dim strRovName
  
  'stores begin and end DTG from the http request
  Dim startCollDTG
  Dim endCollDTG
  
  'stores frame code from the http request
  Dim strFrameCode
  
  'stores converted annotation format
  Dim strFormat
  Dim strTempFormatCode
  Dim strTempObs
  
  'stores converted still image URL
  'Dim strStillImage

  'stores SQL statements
  Dim strSQL
  Dim strSQL2

    
  ''''''''''''''''''''''
  ' Query the database '
  ''''''''''''''''''''''
     
  ' establish a connection to the VARS Database
  Set oConn = Server.CreateObject("ADODB.Connection")
  oConn.Provider = "sqloledb"
  oConn.Open "Server=alaskanwind;DATABASE=VARS;UID=***;******;"

  ' HTTP request should look like http://fww.wf.mbari.org/expd/queries/annotations.asp?qDiveName=879

  ' get the values from the HTTP Request Object    
  ' strRovName = Request.QueryString("qRovName")
  ' strShipName = Request.QueryString("qShipName")
  strDiveName = Request.QueryString("qDiveName")
  
  ' strDiveNameNum = Request.QueryString("qDiveNameNum")
  ' strDiveNum = mid(strDiveNameNum,5)
  ' startCollDTG = Request.QueryString("qSCollDTG")
  ' endCollDTG = Request.QueryString("qECollDTG")
  ' strFrameCode = Request.QueryString("qFrame")
           'Response.Write(startCollDTG)
           'Response.Write(endCollDTG)
           
 ' Convert ship name and write URL to Expedition database if an MBARI ship 
 ' (add new ship names if others are ever in VARS database)
 '          if strShipName ="ptlo" then
 '            strVarsShipName = "Point Lobos"
 '            strVarsROVName = "Ventana"
 '          else
 '            if strShipName="wfly" then
'             	strVarsShipName = "Western Flyer"
  '           	strVarsROVName = "Tiburon"
  '           else 
  '             strVarsShipName = null
  '           end if
  '         end if
  strVarsROVName = "Tiburon"
  strVideoArchiveName = "T0" & strDiveName
    
  %>



 
 
  <!--Open record set and get all annotations-->

  <%

  strSQL = " SELECT DISTINCT "
  strSQL = strSQL & "  CONVERT(CHAR(8), RecordedDate,1)  + ' ' + SUBSTRING(CONVERT(CHAR(19), RecordedDate, 120), 12, 8) AS DateTime, "
  strSQL = strSQL & "  VideoArchiveName AS Tape#, "
  strSQL = strSQL & "  TapeTimeCode, "
  strSQL = strSQL & "  ConceptName, "
  strSQL = strSQL & "  Image AS StillImageURL, "
  strSQL = strSQL & "  Depth, "
  strSQL = strSQL & "  Latitude, "
  strSQL = strSQL & "  Longitude, "
  strSQL = strSQL & "  CameraDirection AS Direction, "
  strSQL = strSQL & "  isNull(Observer, '') AS Observer, "
  strSQL = strSQL & "  FormatCode, "
  strSQL = strSQL & "  ObservationID_FK AS ObservationID "
  strSQL = strSQL & " FROM Annotations "
  strSQL = strSQL & " WHERE "
  strSQL = strSQL & "  (ShipName = '" & strVarsShipName & "' OR RovName = '" & strVarsROVName & "') AND "
  strSQL = strSQL & " VideoArchiveName like '" & strVideoArchiveName & "%' "
  ' strSQL = strSQL & "  RecordedDate BETWEEN '" & startCollDTG & "' And '" & endCollDTG & "' "
  strSQL = strSQL & " Order By TapeTimeCode "
  ' Response.Write(strSQL)

 
  
  %>
  <br>
  <!-- Execting SQL:
  <%=strSQL%>
  -->
  <%
  Set oRS = oConn.Execute (strSQL)

  If Not (oRS.BOF and oRS.EOF) Then
  'If Not (oRS Is Nothing) Then
     'Return table of annotations
     oRS.MoveFirst
  '# e.g. qSCollDTG=3/24/2003+7:13:15+PM&qECollDTG=3/24/2003+9:48:15+PM
  %>
  <table border="0" cellspacing="0" cellpadding="2" valign= "top" >
   <tr>
    <td 
  <a href = "http://fww.wf.mbari.org" border = "0"><img src ="images/wf1-120.jpg"></a>&nbsp;&nbsp;&nbsp;&nbsp;
    </td>
    <td>
  <strong>Video annotations from VARS database for <em>Tiburon</em> dive <%=strDiveName%> <!--<%=strDiveNameNum%>--></strong>
  <!--<b>&#183;</b> 
  <a href="VARS2staging.asp?RovName=<%=strVarsROVName%>&qSCollDTG=<%=startCollDTG%>&qECollDTG=<%=endCollDTG%>">Select samples for loading into the Samples database</a>
  -->
  </td>
   </tr>
  </table>
  
  <br>
  
<table border="1" cellspacing="0" cellpadding="2">
  <tr>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Date and Time</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Tape#</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>TapeTimeCode</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Concept</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Association</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Framegrab</strong></td>
    <!--<td bgcolor="#92C9C9" valign="top" align="left"><strong>Depth(m)</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Latitude</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Longitude</strong></td>
    -->
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Direction</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Observer</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Format</strong></td>
  </tr>  
  
   <%
     If Not (oRS.BOF and oRS.EOF) Then
        oRS.MoveFirst
             
        Do While Not oRS.EOF  
        'make a new row   
            %>
            <tr>
            <%
            recCount = recCount + 1
            
            'convert format code
            strTempFormatCode = Trim(oRS("FormatCode"))
            
            Select case strTempFormatCode
           
             ' for outline
           case "o"
               strFormat = "outline" 
           case "O"
               strFormat = "outline"
              
            'for detail  
            case "d" 
               strFormat = "detail"
            case "D" 
               strFormat = "detail"
               
            'for transect
            case "t" 
               strFormat = "transect"
               
            'for shipboard PV1,PV2,WT1,WT2
            case "1" 
               strTempObs = Trim(oRS("Observer"))
               if strTempObs = "" or strTempObs = "conn" or strTempObs = "jana" or strTempObs = "schlin" or strTempObs = "rodgers" then
                   strFormat = "detailed"
               else strFormat = "shipboard"
               end if
               
             case "2"
               strTempObs = Trim(oRS("Observer"))
               if strTempObs = "" or strTempObs = "conn" or strTempObs = "jana" or strTempObs = "schlin" or strTempObs = "rodgers" then
                   strFormat = "detailed"
               else strFormat = "shipboard"
               end if
             
             case "g"
               strFormat = "geology"
                             
             case "r"
               strFormat = "revised"
                           
            'for any other format code
            case else
               strFormat = strTempFormatCode

            end select
            
        
         ' Create a string for all association values
          strAssociation = ""
          strSQL2 = "SELECT LinkName, ToConcept, LinkValue FROM Association " & _
                      "WHERE (ObservationID_FK = " & oRS("ObservationID") & ")"
            
            Set oRS2 = oConn.Execute (strSQL2)
           
            If Not (oRS2.BOF and oRS2.EOF) Then
               oRS2.MoveFirst
             
               Do While Not oRS2.EOF  
                  strAssociation = strAssociation & oRS2("LinkName")
                  If ((oRS2("ToConcept") <> oRS("ConceptName")) and _
                      (oRS2("ToConcept") <> "nil")) Then
                     strAssociation = strAssociation & " " & oRS2("ToConcept")
                  end if
                  If ((oRS2("LinkValue") <> "nil")) Then
                     strAssociation = strAssociation & " " & oRS2("LinkValue")
                  end if
                  oRS2.MoveNext
                  If Not oRS2.EOF then
                     strAssociation = strAssociation & ", "
                  end if  
               Loop
            End if 
            'Close oRS2 result set
            oRS2.Close
            Set oRS2 = Nothing 

            %>
              <td valign="top" align="left"><%=oRS("DateTime")%>&nbsp;</td>
              <td valign="top" align="center" nowrap><%=oRS("Tape#")%>&nbsp;</td>
              <td valign="top" align="left"><%=oRS("TapeTimeCode")%>&nbsp;</td>
              <td valign="top" align="left"><%=oRS("ConceptName")%>&nbsp;</td>
              <td valign="top" align="left"><%=strAssociation%>&nbsp;</td>
              <%
              if (IsNull(oRS("StillImageURL"))) then
                Response.Write("<td valign='top' align='left'>&nbsp;</td>")
              else
           ' file:/C:/Documents and Settings/tiburon/VARS/data/Tiburon/images/0879/01_33_30_12.jpg   --in VARS
           ' http://fww.wf.mbari.org/Tiburon/images/0879/   --needs to be for web
           ' strDiveNum = mid(strDiveNameNum,5)
              strStillImageURL = right(oRS("StillImageURL"),36) 
              strStillImageURL = "http://fww.wf.mbari.org" & strStillImageURL
             
              
                Response.Write("<td valign='top' align='left'><a href='" & strStillImageURL & "' target=fg>image</a></td>")
                ' Response.Write("<td valign='top' align='left'><a href='" & oRS("StillImageURL") & "' target=fg>image</a></td>")
                'View thumbnails of framegrabs in the table...takes a long time to load
                'Response.Write("<td valign='top' align='left'><a href='" & oRS("StillImageURL") & "' target=_blank><img src=" & oRS("StillImageURL") & " width=90></a></td>")
              end if
              %>
           <!--   <td valign="top" align="left"><%=oRS("Depth")%>&nbsp;</td>
              <td valign="top" align="left"><%=oRS("Latitude")%>&nbsp;</td>
              <td valign="top" align="left"><%=oRS("Longitude")%>&nbsp;</td>
            -->
              <td valign="top" align="left"><%=oRS("Direction")%>&nbsp;</td>
              <td valign="top" align="left"><%=oRS("Observer")%>&nbsp;</td>
              <td valign="top" align="left"><%=strFormat%>&nbsp;</td>
            </tr>
            <%
            
            oRS.MoveNext 
        Loop
        
     End if 
     else
     'Close result set
     oRS.Close
     Set oRS = Nothing  
     Response.Write("<strong>Sorry, no annotations are in the VARS database for the dive specified.</strong>")
  End If
  %>
  
</table>
<p>
<%=recCount%> annotations returned
</p>
  
  
  
  <%
  ' Close the connection to the VARS Database
  
  oConn.Close
  Set oConn = Nothing
  %>
    <p><font face="Arial, Helvetica, sans ser'if"><strong><a href="javascript:window.close();">
    <small>&lt;&lt; Close this Window &gt;&gt;</small></a></strong></font></p>
    <p><em><small>Query returned directly from the VARS Database on the Western Flyer</small></em> <br>
    <em><small>Copyright © 2005 MBARI</small></em> </td>

</body>
</html>
















