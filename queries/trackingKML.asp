<%@ LANGUAGE = PerlScript %>

<%

# Script to query MBARItracking position table for position fixes in certain time bounds
# and build KML on the fly for visualization in Google Earth.
#
# $ID:$

#
# Mike McCann  24 June 20009
# MBARI


use POSIX;
use OLE;
use Win32::ASP;
use Win32::OLE qw( in );		# Needed to iterate through list

#
# Debugging stuff:  All debugging print statements are commented out with 
# '##'.  Must turn on text/html to see the debugging output, should you
# need to do that.
#
print "Content-type: text/html\r\n\r\n";
print "Hello, World.";

##print "String = $ENV{'QUERY_STRING'}\n\n<p>";
@values = split(/&/, $ENV{'QUERY_STRING'});
my %qVars = ();
foreach $i (@values) {
	($varname, $mydata) = split(/=/, $i);
	$qVars{$varname} = $mydata;
	print "The value of $varname is $mydata\n\n<p>";
}

#
# Get parameters from Query String.  GE will provide bbox.  
#

$hoursBack = $qVars{'hoursBack'};
$withTimeStamp = $qVars{'withTimeStamp'};
$stride = $qVars{'stride'};

$dataSince = POSIX::strftime("%Y-%m-%d %H:%M:%S", POSIX::gmtime(time-(3600*$hoursBack)) );
##print  "dataSince = $dataSince\n";
    	
##print "bbox = $bbox\n";
    	


#
# Set up database connection
#
##print "\nConnecting to the Database<br>";
##my $dbh = DBI->connect("DBI:Sybase:server=equinox.shore.mbari.org:51001", "everyone", "guest", {PrintError => 1});
##$dbh->do("use MBARItracking");

#
# Open connection to MBARItracking database
#
$Conn = CreateObject OLE "ADODB.Connection";
$Conn->{'Provider'} = "SQLNCLI";
$Conn->Open("Server=EQUINOX\SQL2008;Database=MBARItracking;UID=everyone;PWD=guest;");

$Errors = $Conn->Errors();
foreach $error (keys %$Errors) {
    	print $error->{Description}, "\n";
   }

#
# Count the records for the query and set up the stride for subsampling
# - careful that we use the same join and where clause as we do below...
#
my $countSQL = <<EOS;

SELECT count(*) AS Count
  FROM [MBARItracking].[dbo].[position]
  WHERE datetime > '$dataSince'
EOS

##print "\nExecuting countSQL: $countSQL";
##my $sth = $dbh->prepare($countSQL);
##$sth->execute;
##while ( my $rrow = $sth->fetchrow_arrayref() ) {
##	$numRows = $$rrow[0];
##}

Win32::ASP::DebugPrint("\nExecuting sql: $countSQL");
$RS = $Conn->Execute($countSQL);

## !!! Fails to get a $RS here  -  Need to confirm connection and OLEDB stuff for SQL2008....   26 June 2009 !!!



$Errors = $RS->Errors();
foreach $error (keys %$Errors) {
    	print $error->{Description}, "\n";
}

$numRows = $RS->Fields('Count')->Value;


print "<br>numRows = $numRows\n";


die;


#
# Construct SQL statements with DateTime given in ISO8601 format.  If SQL passed in then
# construct that query with the additional ISO8601 date format.  Apply the stride from above.
#
my $sql;
$sql = <<EOS;	
SELECT     *, 
           rtrim(convert(char, datetime,126)) + 'Z' as DateTime
FROM       [MBARItracking].[dbo].[position]
WHERE      datetime > '$dataSince'
EOS

	
#
# Execute the SQL, and build the placemarks
#
##print "\nExecuting sql: $sql";
$sth = $dbh->prepare($sql);
$sth->execute;

my $placemarks = '';
my $count = 0;
my %mmsiHash = ();
while ( my $rrow = $sth->fetchrow_hashref() ) {
	##foreach my $k ( keys %$rrow ) {
		##print "k = $k, v = " . $$rrow{$k} . "\n";
	##}
	##print "\n<br>";

		
	$lat = $$rrow{'latitude'};
	$lon = $$rrow{'longitude'};

	$mmsi = $$rrow{'mmsi'};
	$mmsiHash{$mmsi}++ if $mmsi;
	##print "mmsi = $mmsi, mmsiHash = " . $mmsiHash{$mmsi} . "\n";

    $imei = $$row{'imei'};
    $imeiHash{$imei}++ if $imei;

	unless ($lat && $lon) {
			next;
	}

	$isodate = $$rrow{'DateTime'};

	$placemarks .= <<EOP;
    <Placemark>
		<name>$mmsi</name>
		<TimeStamp>
			<when>$isodate</when>
		</TimeStamp>
		<styleUrl>#P${mmsi}</styleUrl>
		<Point>
			<altitudeMode>absolute</altitudeMode>
			<coordinates>$lon,$lat,0</coordinates>
		</Point>
	</Placemark>

EOP

	$count++;

} # End while ( my $rrow = $sth->fetchrow_hashref() )
$dbh->disconnect;

#
# Create a placemark color for each mmsi.  Use the mmsi number divided by 256
# as a seed to create a unique color for each platform
#
my $styles = '';
foreach my $mmsi (keys %mmsiHash) {

	$color = 'ff' . sprintf("%06lx", $mmsi/256);
	$styles .= <<EOS;
	<Style id="P${mmsi}">
		<IconStyle>
			<color>$color</color>
			<scale>0.30</scale>
			<Icon>
				<href>http://maps.google.com/mapfiles/kml/shapes/triangle.png</href>
			</Icon>
		</IconStyle>
		<LabelStyle>
			<scale>0</scale>
		</LabelStyle>
		<BalloonStyle>
		  <bgColor>ffffffff</bgColor> 
		  <textColor>ff000000</textColor> 
		  <text>\$[description]</text>           
		</BalloonStyle>
	</Style>

EOS

}
	

#
# Build up KML and deliver it to the client
#	
my $kml = <<EOK;
<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://earth.google.com/kml/2.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:schemaLocation="http://earth.google.com/kml/2.1
    http://code.google.com/apis/kml/schema/kml21.xsd">


<Document>
  <name>MBARI Tracking data</name>
  <description><![CDATA[$count Fixes<br><br>Resulting from query<br><pre>$sql</pre>]]></description>
$styles
$placemarks
</Document>
</kml>
EOK


$Response->{'ContentType'} = 'application/vnd.google-earth.kml+xml';
print $kml;
%>


