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
    	$qSQL = GetFormValue('qSQL');
    	$clampToSeaFloor = GetFormValue('clampToSeaFloor');
    	
    	##print "bbox = $bbox\n";
    	
    	if ( $bbox ) {
    		($west, $south, $east, $north) = split(',', $bbox);
    	}
    	elsif ( $qSQL ) {
    		# Use this SQL for the query below
    	}
    	else {	# Some default bounding box - for testing
    		($west, $south, $east, $north) = (-124, 35.2, -122, 36);
    	}
    	
    	$center_lng = sprintf("%f.6", (($east - $west) / 2) + $west);
	$center_lat = sprintf("%f.6", (($north - $south) / 2) + $south);
		
	my $sql;
	if ( $qSQL ) {		# Add DateTime and put in some newlines for cleaner display in the description
		# Need to add ISO date time to SELECT
		##$qSQL =~ s/SELECT \*/SELECT rtrim(convert(char, CollectionEventDTG,126)) + 'Z' as DateTime, Longitude, Latitude, Depth, SampleID, FrameGrabImageURL, CollectionVIMSConcept, SampleRefName/;
		##if ( $qSQL =~ /SELECT distinct/i ) {
		##	$qSQL =~ s/SELECT distinct /SELECT distinct rtrim(convert(char, CollectionEventDTG,126)) + 'Z' as DateTime, /;
		##}
		##elsif ( qSQL =~ /SELECT / ) {
		##	$qSQL =~ s/SELECT /SELECT rtrim(convert(char, CollectionEventDTG,126)) + 'Z' as DateTime, /;
		##}
		$qSQL =~ s/FROM/, rtrim(convert(char, CollectionEventDTG,126)) + 'Z' as DateTime \nFROM/;
		$qSQL =~ s/WHERE/\nWHERE/;
		$qSQL =~ s/ORDER/\nORDER/;
		$sql = $qSQL;
	}
	else {
		$sql = <<EOS;	
SELECT     Sample.SampleID, CollectionEvent.Latitude, CollectionEvent.Longitude, CollectionEvent.Depth,
           rtrim(convert(char, CollectionEvent.CollectionEventDTG,126)) + 'Z' as DateTime,
           CollectionEvent.FrameGrabImageURL, CollectionEvent.CollectionVIMSConcept, Sample.SampleRefName
FROM       CollectionEvent INNER JOIN
              Sample ON CollectionEvent.CollectionEventID = Sample.CollectionEventID
WHERE     (CollectionEvent.Latitude > $south) AND (CollectionEvent.Latitude < $north) AND 
          (CollectionEvent.Longitude > $west) AND (CollectionEvent.Longitude < $east)
EOS
	}
	
	##Win32::ASP::DebugPrint("\nExecuting sql: $sql");
	
	
	$Conn = CreateObject OLE "ADODB.Connection";
	$Conn->{'Provider'} = "sqloledb";
	$Conn->Open("Server=perseus;Database=MBARI_Samples;UID=***;******;");
	
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
	

		$sampleID = $RS->Fields('SampleID')->Value;
		
		$lat = $RS->Fields('Latitude')->Value;
		$lon = $RS->Fields('Longitude')->Value;
		$depth = $RS->Fields('Depth')->Value;
		
		unless ($lat && $lon && $depth) {
			$RS->MoveNext;
			next;
		}
		
		$depth = -$depth;
		$isodate = $RS->Fields('DateTime')->Value;
		
		$fgURL = $RS->Fields('FrameGrabImageURL')->Value;
		$concept = $RS->Fields('CollectionVIMSConcept')->Value;
		$name = $RS->Fields('SampleRefName')->Value;
		
		$placemarks .= "<Placemark>\n<name>$name</name>\n";
		$placemarks .= "<TimeStamp>\n<when>$isodate</when>\n</TimeStamp>\n";
		$placemarks .= "<description><![CDATA[";
		$placemarks .= "<div>$concept</div>\n";
        	$placemarks .= "<div><a href=\"https://mww.mbari.org/SamplesDB/queries/selectSample.asp?ID=$sampleID\">$name</a></div>";
        	$placemarks .= "<div><img src=\"$fgURL\"/></div>\n" if $fgURL;
        	$placemarks .= "]]></description>\n";
		$placemarks .= "<styleUrl>#samplesPhotoStyleMap</styleUrl>\n";
		$placemarks .= "<Point>\n";
		if ($clampToSeaFloor) {
			$placemarks .= "<altitudeMode>clampToSeaFloor</altitudeMode>\n";
			$depth = 0;
		}
		else {
			$placemarks .= "<altitudeMode>absolute</altitudeMode>\n";
		}
		$placemarks .= "<coordinates>$lon,$lat,${depth}</coordinates>\n";
		$placemarks .= "</Point>\n";
		$placemarks .= "</Placemark>\n";
		
		$RS->MoveNext;
		$count++;
		
	} # End while ( !$RS->EOF )
	
	my $kml = <<EOK;
<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://earth.google.com/kml/2.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:schemaLocation="http://earth.google.com/kml/2.1
    http://code.google.com/apis/kml/schema/kml21.xsd">


<Folder>
EOK
	if ( $qSQL ) {
		$kml .= "<name>Samples DB Query</name>\n";
	}
	else {
		$kml .= "<name>Samples</name>\n";
	}
	
	$kml .= <<EOK;
<Document>
<name>Individual Samples</name>
	<Style id="samplePhotoPlacemark">
		<IconStyle>
			<Icon>
				<href>http://maps.google.com/mapfiles/kml/paddle/red-square.png</href>
			</Icon>
		</IconStyle>
		<BalloonStyle>
		  <bgColor>ffffffff</bgColor> 
		  <textColor>ff000000</textColor> 
		  <text>\$[description]</text>           
		  <displayMode>default</displayMode> 
		</BalloonStyle>
	</Style>
	<StyleMap id="samplesPhotoStyleMap">
		<Pair>
			<key>normal</key>
			<styleUrl>#samplePlacemark</styleUrl>
		</Pair>
		<Pair>
			<key>highlight</key>
			<styleUrl>#samplePhotoPlacemark</styleUrl>
		</Pair>
	</StyleMap>
	<Style id="samplePlacemark">
		<IconStyle>
			<Icon>
				<href>http://maps.google.com/mapfiles/kml/paddle/red-blank.png</href>
			</Icon>
		</IconStyle>
	</Style>
<visibility>0</visibility>
<open>0</open>
<description><![CDATA[$count Samples this view<br><br>Resulting from query<br><pre>$sql</pre>]]></description>
$placemarks
</Document>
</Folder>
</kml>
EOK

    	
    	
    	$Response->{'ContentType'} = 'application/vnd.google-earth.kml+xml';
    	$Response->AddHeader('content-disposition', 'attachment; filename=samplesDBquery.kml');
    	
    	print $kml;

%>
