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
SELECT     RovName, DiveNumber, DiveLatMid, DiveLonMid, DiveDepthMid,
           rtrim(convert(char, DiveStartDtg,126)) + 'Z' as StartDate,
           rtrim(convert(char, DiveEndDtg,126)) + 'Z' as EndDate
FROM       dbo.Dive
WHERE      (DiveLonMid > $west) AND (DiveLatMid > $south) AND
               (DiveLonMid < $east) AND (DiveLatMid < $north)
ORDER BY   RovName, DiveNumber
EOS
	
	##Win32::ASP::DebugPrint("\nExecuting sql: $sql");
	
	$Conn = CreateObject OLE "ADODB.Connection";
	$Conn->{'Provider'} = "sqloledb";
	$Conn->Open("Server=perseus;Database=expd;UID=everyone;PWD=guest;");
	
	$RS = $Conn->Execute($sql);
	
	$Errors = $Conn->Errors();
	foreach $error (keys %$Errors) {
            	print $error->{Description}, "\n";
    	}
    	my $r;	# Short single letter ROV name
    	my $diveNum;
    	my $lat, $lon;
				
	my $placemarks = '';
	
	my $count = 0;
	while ( !$RS->EOF ) {
	
		$rov = $RS->Fields('RovName')->Value;
	
		$r = 'V' if $RS->Fields('RovName')->Value eq 'vnta';
		$r = 'T' if $RS->Fields('RovName')->Value eq 'tibr';
		$r = 'D' if $RS->Fields('RovName')->Value eq 'docr';
		
		$diveNum = $RS->Fields('DiveNumber')->Value;
		$lat = $RS->Fields('DiveLatMid')->Value;
		$lon = $RS->Fields('DiveLonMid')->Value;
		$dep = $RS->Fields('DiveDepthMid')->Value;
		
		$beg = $RS->Fields('StartDate')->Value;
		$end = $RS->Fields('EndDate')->Value;
		
		# 2000-05-06T16:48:26Z
		$beg =~ /(\d\d\d\d)-/;
		$year = $1;
		
		$placemarks .= <<EOK;
<NetworkLink>
<name>${r}${diveNum}</name>
<TimeSpan>
<begin>$beg</begin>
<end>$end</end>
</TimeSpan>
<visibility>1</visibility>
<open>0</open>
<description></description>
<refreshVisibility>0</refreshVisibility>
<flyToView>0</flyToView>
<Link>
<href>https://mww.mbari.org/3DReplay/geo${year}/$rov/${rov}${diveNum}.kmz</href>
<refreshInterval>2</refreshInterval>
<viewRefreshMode>onRequest</viewRefreshMode>
<viewRefreshTime>1</viewRefreshTime>
</Link>
</NetworkLink>
EOK
		
		##$placemarks .= "<Placemark>\n<name>${r}${diveNum}</name>\n";
		##$placemarks .= "<TimeSpan>\n<begin>$beg</begin>\n<end>$end</end>\n</TimeSpan>\n";
		##$placemarks .= "<Point>\n<coordinates>$lon,$lat,$dep</coordinates>\n</Point>\n";
		##$placemarks .= "</Placemark>\n";
		
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
<name>Individual Dives</name>
<Document>
<name>Individual Dives</name>
<visibility>0</visibility>
<open>0</open>
<description><![CDATA[$count dives this view<br><br>Resulting from query<br><pre>$sql</pre>]]></description>
<!--
<Placemark>
    <name>${r}${diveNum}</name>
    <Point>
        <coordinates>$center_lng,$center_lat</coordinates>
    </Point>
</Placemark>
-->

$placemarks

</Document>
</Folder>
</kml>
EOK

    	
    	$Response->{'ContentType'} = 'application/vnd.google-earth.kml+xml';
    	print $kml;

%>
