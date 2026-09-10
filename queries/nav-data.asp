<%@ LANGUAGE = PerlScript %>
<%

=head1 NAME

camlog-data.asp - Return navigation log data from the Expedition database

=head1 SYNOPSIS

    > http://expd.mbari.org/expd/queries/nav-data.asp?ROVName=vnta&DiveNumber=2606
    > http://expd.mbari.org/expd/queries/nav-data.asp?ROVName=vnta&DiveNumber=2606&type=raw
    > http://expd.mbari.org/expd/queries/nav-data.asp?ROVName=vnta&Esecs=1203443481

=head1 DESCRIPTION

=cut

use Win32::ASP;
use OLE;

# Tried these 2 statements to fix problem with D78, but no joy.
##$Response->Buffer(0);
##$Server->ScriptTimeout(300);

##$_dsn = "Server=fog;Database=expd_ks;UID=expddba;PWD=password;";
$_dsn = "Server=perseus;Database=expd;UID=expddba;PWD=password;";

$ROVName = GetFormValue('ROVName');
$DiveNumber = GetFormValue('DiveNumber');
$rawFlag = GetFormValue('type');
$esecs = GetFormValue('Esecs');
$decim = GetFormValue(decim);

##print "$ROVName, $DiveNumber, $rawFlag, $esecs, $decim\n\n";



if ($ROVName && $DiveNumber){
	returnRecords($ROVName, $DiveNumber);
}
elsif ($ROVName && $esecs){
	returnRecords($ROVName, 0, $esecs);
}
else {
%>
<html>
<head>
<title>Nav data error<%= $ShipName%><%= $YYYYDDD%></title>
</head>
<body>
<h2>Could not return data for ROVName = <%=$ROVName%> DiveNumber = <%=$DiveNumber%></h2>
</body>
</html>

<%
}



#--------------------------------------------------------------------
#

=head3 returnRecords()

Return records that match the dive from the CleanNav or RawNav view

=cut

sub returnRecords {
	my ($rov, $number, $esecs) = @_;
		
	if ( $rov =~ /Ventana/i || $rov =~ /vnta/i ) {
		$view = "dbo.Ventana";
		$rovName = 'vnta';
	}
	elsif ( $rov =~ /Tiburon/i || $rov =~ /tibr/i ) {
		$view = "dbo.Tiburon";
		$rovName = 'tibr';
	}
	elsif ( $rov =~ /DocRicketts/i || $rov =~ /docr/i ) {
		$view = "dbo.DocRicketts";
		$rovName = 'docr';
	}
	elsif ( $rov =~ /MiniROV/i || $rov =~ /mini/i ) {
		$view = "dbo.Minirov";
		$rovName = 'mini';
	}
	else {
	%>
		<h2>'<%=$rov%> is not a valid ROVName.  Please choose 'vnta', 'docr', 'tibr' or 'mini'</h2>
	<%
		return;
	}
	
	if ($esecs) {
		$deltaSecs = 10;	# time span to find match for position
		$sEsecs = $esecs - $deltaSecs / 2;
		$eEsecs = $esecs + $deltaSecs / 2;
	}
		
	
	# Default is to return CleanNav, if 'type=raw' is specified then return from RawNav view
	if ( $rawFlag =~ /raw/i ) {
		$view .= "RawNavData";
	}
	else {
		$view .= "CleanNavData";
	}
		
	
	#
	# Open DB
	#
	$Conn = CreateObject OLE "ADODB.Connection";
	$Conn->{'Provider'} = "sqloledb";
	$Conn->Open($_dsn);
	
	
	Win32::ASP::AddDeathHook( sub { $Conn->Close } );	# In case asp dies do some cleanup
	
	if(!Conn) {
	        $Errors = $Conn->Errors();
	        print "<B>Errors:</B>";
	        foreach $error (keys %$Errors) {
	                print $error->{Description}, "\n";
	        }
	        die "<BR>Empty Connection object. Does some other app have $dsn open?";
	}

	#
	# Query for the dive requested from the view that was set above
	#
	$sql = "SELECT dbo.Dive.RovName, dbo.Dive.DiveNumber, CONVERT(varchar(22), $view.DateTimeGMT, 120) as DateTime24,  \n";
	$sql .= " $view.EpochSecs, $view.Latitude, $view.Longitude, $view.Pressure, $view.Depth, $view.Altitude, $view.Heading,  \n";
    $sql .= " $view.Pitch, $view.Roll, $view.ShipLatitude, $view.ShipLongitude,  \n";
    $sql .= " $view.ShipHeading, $view.QCFlag \n";
	$sql .= " FROM  $view INNER JOIN \n";
	$sql .= " dbo.Dive ON $view.DateTimeGMT > dbo.Dive.DiveStartDtg AND $view.DateTimeGMT < dbo.Dive.DiveEndDtg AND  \n";
	$sql .= " $view.DateTimeGMT > dbo.Dive.DiveStartDtg \n";
	if ($number) {
		$sql .= " WHERE (dbo.Dive.DiveNumber = $number) AND (dbo.Dive.RovName = '$rovName') \n";
	}
	elsif ($esecs) {
		$sql .= " WHERE ($view.EpochSecs > $sEsecs) AND ($view.EpochSecs < $eEsecs) AND (dbo.Dive.RovName = '$rovName') \n";
	}
		
	# If you uncomment this it will destory the text/plain output
	##Win32::ASP::DebugPrint("\nExecuting SQL: \n$sql");
	
	$RS = $Conn->Execute($sql);
	if(!$RS) {
	    $Errors = $Conn->Errors();
	    print "Errors:\n";
	    foreach $error (keys %$Errors) {
	   		print $error->{Description}, "\n";
	  	}	
		$RS->Close;
	 	die "<BR>Empty Return Set. Does some other app have $dsn open?";
	}
	
	@fields = qw(RovName DiveNumber DateTime24 EpochSecs Latitude Longitude Pressure Depth Altitude Heading Pitch Roll ShipLatitude ShipLongitude ShipHeading QCFlag);
	$Response->{'ContentType'} = 'text/plain';
	my $str = '';
	for my $f (@fields) {
		$str .= "${f},";
	}
	$str =~ s/,$/\n/;
	$Response->write($str);

	my $decimate = (defined $decim) ? $decim : 1;
	##Win32::ASP::DebugPrint("\ndecimate = $decimate \n");

	my $count = 0;
	while ( !$RS->EOF ) {
		$count++; 
		if ($count % $decimate) {
			$RS->MoveNext; 
			next;
		}
		my $record = '';
		foreach my $f ( @fields ) { 
			if ( $f =~ /Lat/i || $f =~ /Lon/i ) {
				$record .= sprintf("%.6f", $RS->Fields($f)->value) . ",";
			}
			elsif (  $f =~ /Pres/i || $f =~ /Dep/i || $f =~ /Alt/i ) {
				$record .= sprintf("%6.1f", $RS->Fields($f)->value) . ",";
			}
			elsif ( $f =~ /Head/i) {
				$record .= sprintf("%05.1f", $RS->Fields($f)->value) . ",";
			}
			elsif ( $f =~ /Pitch/i || $f =~ /Roll/i) {
				$record .= sprintf("% 6.2f", $RS->Fields($f)->value) . ",";
			}
			else {
				$record .= $RS->Fields($f)->value . ",";
			}
		}
		$record =~ s/,$/\n/;
		$Response->write($record);
		$RS->MoveNext; 
	} 

} # End returnRecords()
%>