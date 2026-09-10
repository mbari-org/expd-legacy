<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Video annotations</title>
</head>

<body bgcolor="#FFFFFF">

  <%  
  '''''''''''''''''''''''''''''''''''''
  ' Variable and Object Declarations '
  ''''''''''''''''''''''''''''''''''''

  'stores the ship name from the http request
  Dim strShipName
  
  'stores begin and end DTG from the http request
  Dim startCollDTG
  Dim endCollDTG
  
  'stores frame code from the http request
  Dim strFrameCode
  
  'stores background color for table cells
  Dim bgcol
  
  'stores converted annotation format
  Dim strFormat
  Dim strTempFormatCode
  Dim strTempObs
  
  'stores converted still image URL
  Dim strStillImage

  'stores SQL statements
  Dim strSQL
  Dim strSQL2

  'used for dividing up the Still Image URL
  Dim tempArrayURLParts
  Dim strPlatform
  Dim strYearAndYDay
  Dim strImageName
  Dim strYear
  Dim strYDay
  
  ''''''''''''''''''''''
  ' Query the database '
  ''''''''''''''''''''''
     
  ' establish a connection to the VIMS Database
  Set oConn = Server.CreateObject("ADODB.Connection")
  oConn.Provider = "sqloledb"
  oConn.Open "Server=godzilla;DATABASE=Vims;UID=everyone;Password=guest;"

  ' get the values from the HTTP Request Object    
  strShipName = Request.QueryString("qShipName")
  strDiveNameNum = Request.QueryString("qDiveNameNum")
  startCollDTG = Request.QueryString("qSCollDTG")
  endCollDTG = Request.QueryString("qECollDTG")
  strFrameCode = Request.QueryString("qFrame")
           'Response.Write(startCollDTG)
           'Response.Write(endCollDTG)
    
  %>

  <!--Open record set-->

  <%
  
  strSQL = " SELECT DISTINCT CONVERT(CHAR(8), RecordedDtg,1) "
  strSQL = strSQL & " + ' ' + RecordedTime AS DateTime, Annotation_TapeNumber AS Tape#," 
  strSQL = strSQL & " TapeTimeCode, ConceptName, StillImageURL, Depth, Direction," 
  strSQL = strSQL & " isNull(Observer, '') as Observer, FormatCode, ObservationID "
  strSQL = strSQL & " FROM BasicAnnotations JOIN Camera ON (Annotation_CameraID = CameraID) "
  strSQL = strSQL & " JOIN AncillaryData ON (Annotation_AncillaryDataID = AncillaryDataID)"
  strSQL = strSQL & " WHERE ShipName = '" & strShipName & "' AND"
  strSQL = strSQL & " RecordedDTG BETWEEN '" & startCollDTG 
  strSQL = strSQL & "' And '" & endCollDTG & "'"
  strSQL = strSQL & " Order By DateTime"
  
  
  Set oRS = oConn.Execute (strSQL)

  If Not (oRS.BOF and oRS.EOF) Then
  'If Not (oRS Is Nothing) Then
     'Return table of annotations
     oRS.MoveFirst

  %>
  

  <strong>Video annotations for dive <%=strDiveNameNum%>.</strong><br>
  
<table border="1" cellspacing="0" cellpadding="2">
  <tr>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Date and Time</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Tape#</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>TapeTimeCode</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Concept</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Association</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Framegrab</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Depth(m)</strong></td>
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
            'make the cell backgrounds yellow if the annotation matches the sample's
            If oRS("TapeTimeCode") = strFrameCode then
               bgcol = "yellow"
            else 
                bgcol = "white"
            end if
            
            'convert format code
            strTempFormatCode = Trim(oRS("FormatCode"))
            
            Select case strTempFormatCode
           
             ' for outline
            case "o"
               strFormat = "outline"
              
            'for detail  
            case "d" 
               strFormat = "detail"
               
            'for transect
            case "t" 
               strFormat = "transect"
               
            'for shipboard
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
                             
             case "r"
               strFormat = "revised"
                           
            'for any other format code
            case else
               strFormat = strTempFormatCode

            end select
            
             
        'Convert the still image URL
         if Not (IsNull(oRS("StillImageURL"))) then 
            strStillImage = Trim(oRS("StillImageURL"))
            'for DEBUGGING, force image filename (comment out if/then)
            'strStillImage = "Ventana/1997335/04:21:14:24.rgb"
            '
            tempArrayURLParts = Split (strStillImage,"/")
            strPlatform = tempArrayURLParts(0)
            strYearAndYDay = tempArrayURLParts(1)
            strImageName = tempArrayURLParts(2)
           
            strYear = Trim(Left(strYearAndYDay,4))
            strYDay = Trim(Right(strYearAndYDay,3))
            strImageName = Replace (strImageName, ".rgb", ".jpg")
            strImageName = Server.URLEncode(strImageName)  

            strStillImage = "http://mww.mbari.org/ARCHIVE/frameGrabs/" & strPlatform
            strStillImage = strStillImage & "/stills/" & strYear & "/" & strYDay & "/"
            strStillImage = strStillImage & strImageName
         
            'Response.Write("Debug " & strStillImage & "<br>")
         end if
  
         ' Create a string for all association values
          strAssociation = ""
          strSQL2 = "SELECT LinkName, ToConcept, LinkValue FROM Association " & _
                      "WHERE (Association_ObservationID = " & oRS("ObservationID") & ")"
            
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
              <td bgcolor="<%=bgcol%>" valign="top" align="left"><%=oRS("DateTime")%>&nbsp;</td>
              <td bgcolor="<%=bgcol%>" valign="top" align="center"><%=oRS("Tape#")%>&nbsp;</td>
              <td bgcolor="<%=bgcol%>" valign="top" align="left"><%=oRS("TapeTimeCode")%>&nbsp;</td>
              <td bgcolor="<%=bgcol%>" valign="top" align="left"><%=oRS("ConceptName")%>&nbsp;</td>
              <td bgcolor="<%=bgcol%>" valign="top" align="left"><%=strAssociation%>&nbsp;</td>
              <%
              if (IsNull(oRS("StillImageURL"))) then
                Response.Write("<td bgcolor=" & bgcol & " valign='top' align='left'>&nbsp;</td>")
              else
                Response.Write("<td bgcolor=" & bgcol & " valign='top' align='left'><a href='" & strStillImage & "' target=_blank>image</a></td>")
                'View thumbnails of framegrabs in the table...takes a long time to load
                'Response.Write("<td bgcolor=" & bgcol & " valign='top' align='left'><a href='" & strStillImage & "' target=_blank><img src=" & strStillImage & " width=90></a></td>")
              end if
              %>
              <td bgcolor="<%=bgcol%>" valign="top" align="left"><%=oRS("Depth")%>&nbsp;</td>
              <td bgcolor="<%=bgcol%>" valign="top" align="left"><%=oRS("Direction")%>&nbsp;</td>
              <td bgcolor="<%=bgcol%>" valign="top" align="left"><%=oRS("Observer")%>&nbsp;</td>
              <td bgcolor="<%=bgcol%>" valign="top" align="left"><%=strFormat%>&nbsp;</td>
            </tr>
            <%
            
            oRS.MoveNext 
        Loop
        
     End if 
     else
     'Close result set
     oRS.Close
     Set oRS = Nothing  
     Response.Write("<strong>Sorry, no annotations are in the VIMS database for the time limits specified.</strong>")
  End If
  %>
  
</table>
   <p>Note that not all interpretation files created aboard ship have been 
   rectified yet with those created in the video lab or added to the annotation database.</p>
   <p>To view the full richness of the video annotation database, including position, ROV CTDO, 
   and camera setting data associated with the annotations, or to download ASCII versions of the data, 
   query using the <a href="/itd/video/VIMS/vimsquery.htm" target="blank">
   Vims Query java interface</a>. 
   </p>
  
  
  <%
  ' Close the connection to the VIMS Database
  
  oConn.Close
  Set oConn = Nothing
  %>
    <p><font face="Arial, Helvetica, sans ser'if"><strong><a href="javascript:window.close();">
    <small>&lt;&lt; Close this Window &gt;&gt;</small></a></strong></font></p>
    <p><em><small>Query returned directly from the Video Annotations Database</small></em> <br>
    <em><small>Copyright © 2000 MBARI</small></em> </td>

</body>
</html>
















