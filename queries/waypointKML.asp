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
		
	my $sql = <<EOS;
	
SELECT     WaypointName, Latitude, Longitude, Depth, 
           rtrim(convert(char, CreateDTG,126)) + 'Z' as CreateTime,
           rtrim(convert(char, ExpireDTG,126)) + 'Z' as ExpireTime,
           Altitude, Comment, dbo.Person.LastName, dbo.Person.FirstName
FROM       dbo.Waypoint INNER JOIN
              dbo.Person ON dbo.Waypoint.WaypointOwnerID = dbo.Person.PersonID	
WHERE     (Latitude > $south) AND (Latitude < $north) AND 
          (Longitude > $west) AND (Longitude < $east)
EOS
	
	
	##Win32::ASP::DebugPrint("\nExecuting sql: $sql");
	
	$Conn = CreateObject OLE "ADODB.Connection";
	$Conn->{'Provider'} = "sqloledb";
	$Conn->Open("Server=perseus;Database=Expd;UID=everyone;PWD=guest;");
	
	$RS = $Conn->Execute($sql);
	
	$Errors = $Conn->Errors();
	foreach $error (keys %$Errors) {
            	print $error->{Description}, "\n";
    	}

    	my $sampleID;
    	my $lat, $lon, $depth;
				
	my $placemarks = '';
	
	my $count = 0;
	while ( !$RS->EOF ) {
	
		$lat = $RS->Fields('Latitude')->Value;
		$lon = $RS->Fields('Longitude')->Value;
		$depth = $RS->Fields('Depth')->Value;
		
		$createTime = $RS->Fields('CreateTime')->Value;
		$expireTime = $RS->Fields('expireTime')->Value;

		$name = $RS->Fields('WaypointName')->Value;
		$owner = $RS->Fields('FirstName')->Value . " " . $RS->Fields('LastName')->Value;
		$comment = $RS->Fields('Comment')->Value;
		$altitude = $RS->Fields('Altitude')->Value;

		
		$placemarks .= "<Placemark>\n<name>$name</name>\n";
		$placemarks .= "<TimeSpan>\n<begin>$createTime</begin>\n<end>$expireTime</end>\n</TimeSpan>\n";
		$placemarks .= "<description><![CDATA[";
		$placemarks .= "<div><u><b>Waypoint</b></u></div><br>\n";
		$placemarks .= "<div><b>Name:</b> $name</div>\n";
		$placemarks .= "<div><b>Owner:</b> $owner</div>\n";
		$placemarks .= "<div><b>Create Time:</b> $createTime</div>\n";
		$placemarks .= "<div><b>Expire Time:</b> $expireTime</div>\n";
		$placemarks .= "<div><b>Comment:</b> $comment</div>\n";
        	$placemarks .= "]]></description>\n";
		$placemarks .= "<styleUrl>#msn_square_copy29</styleUrl>\n";
		$placemarks .= "<Point>\n<altitudeMode>absolute</altitudeMode>\n";
		$placemarks .= "<coordinates>$lon,$lat,-${depth}</coordinates>\n</Point>\n";
		$placemarks .= "</Placemark>\n";
		
		$RS->MoveNext;
		$count++;
		
	} # End while ( !$RS->EOF )
	
	my $kml = <<EOK;
<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://earth.google.com/kml/2.2"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:schemaLocation="http://earth.google.com/kml/2.2
    http://code.google.com/apis/kml/schema/kml22beta.xsd"> 


<Folder>
<name>Waypoints</name>
<Document>
<name>Individual Waypoints</name>
	<Style id="sn_square_copy29">
		<IconStyle>
			<color>ffff0000</color>
			<scale>0.5</scale>
			<Icon>
				<href>http://maps.google.com/mapfiles/kml/shapes/square.png</href>
			</Icon>
		</IconStyle>
		<LabelStyle>
			<scale>0.7</scale>
		</LabelStyle>		
	</Style>
	<Style id="sh_square_copy13">
		<IconStyle>
			<color>ffff0000</color>
			<scale>0.66</scale>
			<Icon>
				<href>http://maps.google.com/mapfiles/kml/shapes/square.png</href>
			</Icon>
		</IconStyle>
		<LabelStyle>
			<scale>0.77</scale>
		</LabelStyle>
		<BalloonStyle>
		  <bgColor>ffffffff</bgColor> 
		  <textColor>ff000000</textColor> 
		  <text>\$[description]</text>           
		  <displayMode>default</displayMode> 
		</BalloonStyle>
	</Style>
	<StyleMap id="msn_square_copy29">
		<Pair>
			<key>normal</key>
			<styleUrl>#sn_square_copy29</styleUrl>
		</Pair>
		<Pair>
			<key>highlight</key>
			<styleUrl>#sh_square_copy13</styleUrl>
		</Pair>
	</StyleMap>
<visibility>0</visibility>
<open>0</open>
<description><![CDATA[$count Waypoints in this view<br><br>Resulting from query<br><pre>$sql</pre>]]></description>
$placemarks
</Document>
</Folder>
</kml>
EOK

    	$Response->{'ContentType'} = 'application/vnd.google-earth.kml+xml';
    	print $kml;

%>
