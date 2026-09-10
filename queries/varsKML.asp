<%@ LANGUAGE = PerlScript %>

<% 
	#
	# See the bottom of http://code.google.com/apis/kml/documentation/kml_tut.html
	# for the Python version.
	#
	use OLE;
	use Win32::ASP;
	use Win32::OLE qw( in );		# Needed to iterate through list

    	$bbox = GetFormValue('BBOX');
    	
    	##print "bbox = $bbox\n";
    	
    	if ( $bbox ) {
    		($west, $south, $east, $north) = split(',', $bbox);
    	}
    	else {	# Some default bounding box - for testing
    		($west, $south, $east, $north) = (-124, 35.2, -122, 36);
    	}
    	
    	$center_lng = sprintf("%f.6", (($east - $west) / 2) + $west);
	$center_lat = sprintf("%f.6", (($north - $south) / 2) + $south);
	
	#
	# Open connection to VARS database
	#
	$Conn = CreateObject OLE "ADODB.Connection";
	$Conn->{'Provider'} = "sqloledb";
	$Conn->Open("Server=perseus,1433;Database=VARS;UID=everyone;PWD=guest;");
		
	#
	# Count the records in this view
	#
	my $sql = <<EOS;
SELECT     COUNT(*) AS Count
FROM         dbo.PhysicalData 
WHERE     (Latitude > $south) AND (Longitude > $west) AND (Longitude < $east) AND (Latitude < $north)
EOS
	##Win32::ASP::DebugPrint("\nExecuting sql: $sql");
	$RS = $Conn->Execute($sql);
	
	$Errors = $Conn->Errors();
	$errorText = '';
	foreach $error (keys %$Errors) {
            	##print $error->{Description}, "\n";
            	$errorText .= $error->{Description} . "\n";
    	}
    	
    	$count = $RS->Fields('Count')->Value;
    	$annontationsToDisplay = 500;
    	$stride = int($count / $annontationsToDisplay);
    	
#  
	$sql = <<EOS;	
SELECT rtrim(convert(char, RecordedDate,126)) + 'Z' as DateTime, RecordedDate, PhysicalDataID_FK,
       videoArchiveName AS Tape#, Temperature, Salinity, Oxygen, Light, RovName, DiveNumber,
       TapeTimeCode, ConceptName, Image AS StillImageURL, Depth, Latitude, Longitude, CameraDirection AS Direction, ISNULL(Observer, '') AS Observer, 
       AnnotationMode, ObservationID_FK AS ObservationID
FROM   dbo.Annotations
WHERE  (Latitude > $south) AND (Longitude > $west) AND (Longitude < $east) AND (Latitude < $north) AND PhysicalDataID_FK % $stride = 1
EOS
	
	##Win32::ASP::DebugPrint("\nExecuting sql: $sql");
	$RS = $Conn->Execute($sql);
	
	$Errors = $Conn->Errors();
	
	foreach $error (keys %$Errors) {
            	##print $error->{Description}, "\n";
            	$errorText .= $error->{Description} . "\n";
    	}

    	my $lat, $lon, $depth;
				
	my $placemarks = '';
	
	while ( !$RS->EOF ) {
	
		$lat = $RS->Fields('Latitude')->Value;
		$lon = $RS->Fields('Longitude')->Value;
		$depth = sprintf("%.1f", $RS->Fields('Depth')->Value);
		
		$isodate = $RS->Fields('DateTime')->Value;
		
		$fgURL = $RS->Fields('StillImageURL')->Value;
		$concept = $RS->Fields('ConceptName')->Value;
		
		$placemarks .= "<Placemark>\n<name>$concept</name>\n";
		$placemarks .= "<TimeStamp>\n<when>$isodate</when>\n</TimeStamp>\n";
		$placemarks .= "<description><![CDATA[";
		$placemarks .= "<div><b>$concept</b><br/><br/></div>";
		$placemarks .= "<div><b>Date: </b>" . $RS->Fields('DateTime')->Value . "\n";
		$placemarks .= "<div><b>Dive: </b>" . $RS->Fields('RovName')->Value . " " . $RS->Fields('DiveNumber')->Value . "\n";
		$placemarks .= "<div><b>Tape Timecode: </b>" . $RS->Fields('TapeTimeCode')->Value . "\n";
		$placemarks .= "<div><b>Depth: </b>$depth</div>\n";
		$placemarks .= "<div><b>Temperature: </b>" . sprintf("%.3f", $RS->Fields('Temperature')->Value) . " C\n";
		$placemarks .= "<div><b>Salinity: </b>" . sprintf("%.2f", $RS->Fields('Salinity')->Value) . "\n";
		$placemarks .= "<div><b>Oxygen: </b>" . sprintf("%.2f", $RS->Fields('Oxygen')->Value) . " ml/l\n";
		$placemarks .= "<div><b>Light: </b>" . sprintf("%.1f", $RS->Fields('Light')->Value) . " %\n";

        	$placemarks .= "<div><img src=\"$fgURL\"/></div>" if $fgURL;
        	$placemarks .= "]]></description>\n";
		$placemarks .= "<styleUrl>#varsPhotoStyleMap</styleUrl>\n";
		$placemarks .= "<Point>\n<altitudeMode>absolute</altitudeMode>\n";
		$placemarks .= "<coordinates>$lon,$lat,-${depth}</coordinates>\n</Point>\n";
		$placemarks .= "</Placemark>\n";
		
		$RS->MoveNext;
		
	} # End while ( !$RS->EOF )
	
	my $kml = <<EOK;
<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://earth.google.com/kml/2.2"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:schemaLocation="http://earth.google.com/kml/2.2
    http://code.google.com/apis/kml/schema/kml22beta.xsd"> 

<Folder>
<name>Annotations</name>
<Document>
<name>Individual Annotations</name>
	<Style id="varsPhotoPlacemark">
	    <IconStyle>
	    	<color>ff00ffff</color>
            	<scale>0.3</scale>
            	<Icon>
                    <href>http://maps.google.com/mapfiles/kml/shapes/square.png</href>
            	</Icon>
            </IconStyle>
	    <BalloonStyle>
		  <bgColor>ffffffff</bgColor> 
		  <textColor>ff000000</textColor> 
		  <text>\$[description]</text>           
		  <displayMode>default</displayMode> 
	    </BalloonStyle>
	</Style>
	<StyleMap id="varsPhotoStyleMap">
		<Pair>
			<key>normal</key>
			<styleUrl>#varsPlacemark</styleUrl>
		</Pair>
		<Pair>
			<key>highlight</key>
			<styleUrl>#varsPhotoPlacemark</styleUrl>
		</Pair>
	</StyleMap>	
	<Style id="varsPlacemark">
	    <IconStyle>
	    	<color>ff00ffff</color>
            	<scale>0.5</scale>
            	<Icon>
                    <href>http://maps.google.com/mapfiles/kml/shapes/square.png</href>
            	</Icon>
            </IconStyle>
	</Style>

	

<visibility>0</visibility>
<open>0</open>
EOK 

if ( $errorText ) {
	$kml .= "<description>.<![CDATA[Error:<br><br>$errorText<br><br>Resulting from query:<br><pre>$sql</pre>]]></description>\n";
}
else {
	$kml .= "<description>.<![CDATA[$count annotations in this view, showing every $stride<br><br>Resulting from query<br><pre>$sql</pre>]]></description>\n";
}

$kml .= <<EOK;
$placemarks
</Document>
</Folder>
</kml>
EOK

    	$Response->{'ContentType'} = 'application/vnd.google-earth.kml+xml';
    	$Response->AddHeader('content-disposition', 'attachment; filename=samplesDBquery.kml');
    	print $kml;

%>
