<%@ LANGUAGE = PerlScript %>

<% 
	#
	# See the bottom of http://code.google.com/apis/kml/documentation/kml_tut.html
	# for the Python version.
	#
	use OLE;
	use Win32::ASP;
	use Win32::OLE qw( in );		# Needed to iterate through list
	use LWP::Simple;

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
SELECT     TerrainName, MinLat, MaxLat, MinLon, MaxLon, TopTileURL, TopImageURL, CenterResolution, NumLevels
FROM       dbo.Terrain      
WHERE     TerrainType = 'kml' 
EOS
##          and ( ($north BETWEEN MinLat AND MaxLat) AND ($west BETWEEN MinLon AND MaxLon) OR
##          ($north BETWEEN MinLat AND MaxLat) AND ($east BETWEEN MinLon AND MaxLon) OR
##          ($south BETWEEN MinLat AND MaxLat) AND ($west BETWEEN MinLon AND MaxLon) OR
##          ($south BETWEEN MinLat AND MaxLat) AND ($east BETWEEN MinLon AND MaxLon) OR  (MinLat BETWEEN $south AND $north) AND ((MinLon BETWEEN $west AND $east) OR (MaxLon BETWEEN $west AND $east)) OR
##          (MaxLat BETWEEN $south AND $north) AND ((MinLon BETWEEN $west AND $east) OR (MaxLon BETWEEN $west AND $east)) OR
##          (MinLon BETWEEN $west AND $east) AND ((MinLat BETWEEN $south AND $north) OR (MaxLat BETWEEN $south AND $north)) OR
##          (MaxLon BETWEEN $west AND $east) AND ((MinLat BETWEEN $south AND $north) OR (MaxLat BETWEEN $south AND $north)) )
##EOS
	
	
	##Win32::ASP::DebugPrint("\nExecuting sql: $sql");
	
	$Conn = CreateObject OLE "ADODB.Connection";
	$Conn->{'Provider'} = "sqloledb";
	$Conn->Open("Server=perseus;Database=Expd;UID=everyone;PWD=guest;");
	
	$RS = $Conn->Execute($sql);
	
	$Errors = $Conn->Errors();
	foreach $error (keys %$Errors) {
            	print $error->{Description}, "\n";
    	}
		
	my $networkLinks = '';
	
	my $count = 0;
	while ( !$RS->EOF ) {
	
		$name = $RS->Fields('TerrainName')->Value;
		
		#
		# Skip obvious coverages we don't want
		#
		if ($name =~ /globe/i || $name =~ /NEP_Hawaii/i) {
			$RS->MoveNext;
			next;
		}

		$minLat = $RS->Fields('MinLat')->Value;
		$maxLat = $RS->Fields('MaxLat')->Value;
		$minLon = $RS->Fields('MinLon')->Value;
		$maxLon = $RS->Fields('MaxLon')->Value;
		
		
		
		##$isodate = $RS->Fields('DateTime')->Value;
		
		$topTileURL = $RS->Fields('TopTileURL')->Value;
		$topImageURL = $RS->Fields('TopImageURL')->Value;
		$name = $RS->Fields('TerrainName')->Value;
		
		my $logFile = $topTileURL;
		$logFile =~ s/1\.kml/$name.out/;
		my $outputLog = get($logFile);
		$outputLog =~ s/\</\&#60;/g;
		$outputLog =~ s/\>/\&#62;/g;
		
		##die $outputLog; 
		
		$networkLinks .= <<EONL;
		
<NetworkLink>
	<name>$name</name>
	<Snippet maxLines="1" >
		Coverage details
	</Snippet>
	<description><![CDATA[<pre>$outputLog</pre>]]></description>

	<open>0</open>
	<Link>
		<href>$topTileURL</href>
	</Link>
</NetworkLink>
EONL
				
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
<name>Terrain coverages</name>
<Document>
<name>Individual Terrain coverages</name>
	
<visibility>0</visibility>
<open>0</open>
<description><![CDATA[$count coverages this view<br><br>Resulting from query<br><pre>$sql</pre>]]></description>
$networkLinks
</Document>
</Folder>
</kml>
EOK

    	$Response->{'ContentType'} = 'application/vnd.google-earth.kml+xml';
    	print $kml;

%>
