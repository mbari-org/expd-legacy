<%@ Language=VBScript %>
<%Response.Buffer = False%>
<%Server.ScriptTimeout = 300%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>VARS video annotations</title>
</head>

<body bgcolor="#FFFFFF">

  <%  
  '''''''''''''''''''''''''''''''''''''
  ' Variable and Object Declarations '
  ''''''''''''''''''''''''''''''''''''

  'stores the ship name from the http request and convert to VARS version
  Dim strShipName
  Dim strVarsShipName
  
  'stores begin and end DTG from the http request
  Dim startCollDTG
  Dim endCollDTG
  
  'stores frame code from the http request
  Dim strFrameCode
  
  'stores converted annotation format
  Dim strFormat
  Dim strTempAnnotationMode
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
  oConn.ConnectionTimeout=180
  'oConn.Open "Server=perseus,1433;DATABASE=VARS;UID=***;******;"
  oConn.Open "Server=venus,1433;DATABASE=M3_ANNOTATIONS;UID=***;******;"

  ' get the values from the HTTP Request Object    
  strShipName = Request.QueryString("qShipName")
  strDiveNameNum = Request.QueryString("qDiveNameNum")
  strRovInitial = mid(strDiveNameNum,1,1)
  strDiveNum = mid(strDiveNameNum,5)
  startCollDTG = Request.QueryString("qSCollDTG")
  endCollDTG = Request.QueryString("qECollDTG")
  strFrameCode = Request.QueryString("qFrame")
           'Response.Write(startCollDTG)
           'Response.Write(endCollDTG)
           

 Response.Write("<!-- strRovInitial = " + strRovInitial + "-->")
           
 ' Convert ship name and write URL to Expedition database if an MBARI ship 
 ' (add new ship names if others are ever in VARS database)
		if strShipName ="ptlo" then
			strVarsShipName = "Point Lobos"
			strVarsROVName = "Ventana"
         elseif strShipName="wfly" then
            strVarsShipName = "Western Flyer"
            if strRovInitial="d" then
            	strVarsROVName = "Doc Ricketts"
            else
            	 strVarsROVName = "Tiburon"
            end if
		elseif strShipName ="rcsn" then
			strVarsShipName = "Rachel Carson"
			strVarsROVName = "Ventana"
		else 
		  strVarsShipName = null
		end if    
  %>



 
 
  <!--Open record set and get all annotations-->

  <%
  ' Old BasicAnnotations View query
  'strSQL = " SELECT DISTINCT CONVERT(CHAR(8), RecordedDtg,1) "
  'strSQL = strSQL & " + ' ' + RecordedTime AS DateTime, Annotation_TapeNumber AS Tape#," 
  'strSQL = strSQL & " TapeTimeCode, ConceptName, StillImageURL, Depth, Latitude, Longitude, Direction," 
  'strSQL = strSQL & " isNull(Observer, '') as Observer, AnnotationMode, ObservationID "
  'strSQL = strSQL & " FROM BasicAnnotations JOIN CameraData ON (Annotation_CameraID = id) "
  'strSQL = strSQL & " JOIN AncillaryData ON (Annotation_AncillaryDataID = AncillaryDataID)"
  'strSQL = strSQL & " WHERE (ShipName = '" & strVarsShipName & "' OR PlatformName = '" & strVarsROVName & "') AND"
  'strSQL = strSQL & " RecordedDTG BETWEEN '" & startCollDTG 
  'strSQL = strSQL & "' And '" & endCollDTG & "'"
  'strSQL = strSQL & " Order By DateTime"
  
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
  strSQL = strSQL & "  AnnotationMode, "
  strSQL = strSQL & "  ObservationID_FK AS ObservationID "
  'strSQL = strSQL & " FROM Annotations "
  strSQL = strSQL & " FROM Annotations_legacy "
  strSQL = strSQL & " WHERE "
  strSQL = strSQL & "  (ShipName = '" & strVarsShipName & "' OR RovName = '" & strVarsROVName & "') AND "
  strSQL = strSQL & "  RecordedDate BETWEEN '" & startCollDTG & "' And '" & endCollDTG & "' "
  strSQL = strSQL & " Order By DateTime "

 
  
  %>
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
  
  
  
  <strong>Video annotations from M3_ANNOTATIONS database for dive <%=strDiveNameNum%></strong>
  <b>&#183;</b> 
  <a href="VARS2staging.asp?RovName=<%=strVarsROVName%>&qSCollDTG=<%=startCollDTG%>&qECollDTG=<%=endCollDTG%>&qDiveNameNum=<%=strDiveNameNum%>">Select samples for loading into the Samples database</a>
  <br>
  
<table border="1" cellspacing="0" cellpadding="2">
  <tr>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Date and Time</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Tape#</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>TapeTimeCode</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Concept</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Association</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Framegrab</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Depth(m)</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Latitude</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Longitude</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Direction</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>Observer</strong></td>
    <td bgcolor="#92C9C9" valign="top" align="left"><strong>AnnotationMode</strong></td>
  </tr>  
  
   <%
     If Not (oRS.BOF and oRS.EOF) Then
        oRS.MoveFirst
             
        Do While Not oRS.EOF  
			if InStr(oRS("StillImageURL"), ".png") > 0 Then
				oRS.MoveNext
                if oRS.EOF Then
                    Exit Do
                End If
			End If
			
			'make a new row   
            %>
            <tr>
            <%
            recCount = recCount + 1
            
            'convert format code
            strTempAnnotationMode = Trim(oRS("AnnotationMode"))
            
            Select case strTempAnnotationMode
           
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
               strFormat = strTempAnnotationMode

            end select
            
        
         ' Create a string for all association values
          strObservationID = oRS("ObservationID")
		  If Not IsNull(strObservationID) Then
			
			  strAssociation = ""
			  'strSQL2 = "SELECT LinkName, ToConcept, LinkValue FROM Association " & _
			  '            "WHERE (ObservationID_FK = " & strObservationID & ")"
			  strSQL2 = "SELECT LinkName, ToConcept, LinkValue FROM Associations_legacy " & _
						  "WHERE (ObservationID_FK = '" & strObservationID & "')"
				%>
				  <!-- Exectuing strSQL2:
				  <%=strSQL2%>
				  -->
				<%
						  
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
				  <td valign="top" align="center"><%=oRS("Tape#")%>&nbsp;</td>
				  <td valign="top" align="left"><%=oRS("TapeTimeCode")%>&nbsp;</td>
				  <td valign="top" align="left"><%=oRS("ConceptName")%>&nbsp;</td>
				  <td valign="top" align="left"><%=strAssociation%>&nbsp;</td>
				  <%
				  if (IsNull(oRS("StillImageURL"))) then
					Response.Write("<td valign='top' align='left'>&nbsp;</td>")
				  else
				  
					Response.Write("<td valign='top' align='left'><a href='" & oRS("StillImageURL") & "' target=fg>image</a></td>")
					'View thumbnails of framegrabs in the table...takes a long time to load
					'Response.Write("<td valign='top' align='left'><a href='" & oRS("StillImageURL") & "' target=_blank><img src=" & oRS("StillImageURL") & " width=90></a></td>")
				  end if
				  %>
				  <td valign="top" align="left"><%=oRS("Depth")%>&nbsp;</td>
				  <td valign="top" align="left"><%=oRS("Latitude")%>&nbsp;</td>
				  <td valign="top" align="left"><%=oRS("Longitude")%>&nbsp;</td>
				  <td valign="top" align="left"><%=oRS("Direction")%>&nbsp;</td>
				  <td valign="top" align="left"><%=oRS("Observer")%>&nbsp;</td>
				  <td valign="top" align="left"><%=strFormat%>&nbsp;</td>
				</tr>
				<%
				
			' uuid exists
			End If
            oRS.MoveNext 
        Loop
        
     End if 
     else
     'Close result set
     oRS.Close
     Set oRS = Nothing  
     Response.Write("<strong>Sorry, no annotations are in the VARS database for the time limits specified.</strong>")
  End If
  %>
  
</table>
<p>
<%=recCount%> annotations returned
</p>
   <p>To view the full richness of the video annotation database, including ROV CTDO, 
   camera setting data associated with the annotations, or to download ASCII versions of the data, 
   query using the <a href="http://tienhou/vimsproject/Developer/index.html" target="blank">
   Vars Query java interface</a> (download from link found under "Webstart"). 
   </p>
  
  
  <%
  ' Close the connection to the VARS Database
  
  oConn.Close
  Set oConn = Nothing
  %>
    <p><font face="Arial, Helvetica, sans ser'if"><strong><a href="javascript:window.close();">
    <small>&lt;&lt; Close this Window &gt;&gt;</small></a></strong></font></p>
    <p><em><small>Query returned directly from the VARS Database</small></em> <br>
    <em><small>Copyright © 2004 MBARI</small></em> </td>

</body>
</html>
















