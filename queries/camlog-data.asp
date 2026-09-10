<%@ LANGUAGE = PerlScript %>

<%

=head1 NAME

camlog-data.asp - Return camera log data from the Expedition database

=head1 SYNOPSIS

    > http://expd.mbari.org/expd/queries/camlog-data.asp?ROVName=vnta&DiveNumber=2606

=head1 DESCRIPTION

=cut

use Win32::ASP;
use OLE;

##$_dsn = "Server=fog;Database=expd_ks;UID=expddba;PWD=password;";
$_dsn = "Server=perseus;Database=expd;UID=expddba;PWD=password;";


$ROVName = GetFormValue('ROVName');
$DiveNumber = GetFormValue('DiveNumber');



if ($ROVName && $DiveNumber){
	returnRecords($ROVName, $DiveNumber);
}
else {
%>
<html>
<head>
<title>Camlog data error<%= $ShipName%><%= $YYYYDDD%></title>
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

Return records that match the ive from the camlog view

=cut

sub returnRecords {
	my ($rov, $number) = @_;
		
	if ( $rov =~ /Ventana/i || $rov =~ /vnta/i ) {
		$view = "dbo.VentanaCamlogData";
		$rovName = 'vnta';
	}
	elsif ( $rov =~ /Tiburon/i || $rov =~ /tibr/i ) {
		$view = "dbo.TiburonCamlogData";
		$rovName = 'tibr';
	}
	elsif ( $rov =~ /DockRickets/i || $rov =~ /docr/i ) {
		$view = "dbo.DocRickettsCamlogData";
		$rovName = 'docr';
	}
	else {
	%>
		<h2>'<%=$rov%> is not a valid ROVName.  Please choose 'vnta', 'docr' or 'tibr'</h2>
	<%
		return;
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
	
	if ( $rovName eq 'docr' ) {
		$focusZoomIris = "$view.focusVolts, $view.zoomVolts, $view.irisVolts";
		@fields = qw(RovName DiveNumber DateTime24 betaTimecode hdTimecode focusVolts zoomVolts irisVolts);
	}
	else {
		$focusZoomIris = "$view.mainFocus, $view.mainZoom";
		@fields = qw(RovName DiveNumber DateTime24 betaTimecode hdTimecode mainFocus mainZoom);
	}

	#
	# Query for the dive requested from the view that was set above
	#
	$sql = "SELECT dbo.Dive.RovName, dbo.Dive.DiveNumber, CONVERT(varchar(22), $view.DateTimeGMT, 120) as DateTime24, $view.betaTimecode,  \n";
	$sql .= " $view.hdTimecode, $focusZoomIris \n";
	$sql .= " FROM  $view INNER JOIN \n";
	$sql .= " dbo.Dive ON $view.DateTimeGMT > dbo.Dive.DiveStartDtg AND $view.DateTimeGMT < dbo.Dive.DiveEndDtg AND  \n";
	$sql .= " $view.DateTimeGMT > dbo.Dive.DiveStartDtg \n";
	$sql .= " WHERE (dbo.Dive.DiveNumber = $number) AND (dbo.Dive.RovName = '$rovName') \n";
		
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
	
	
	$Response->{'ContentType'} = 'text/plain';
	if ( $rovName eq 'docr' ) {
		$Response->write("RovName,DiveNumber,DateTimeGMT,betaTimecode,hdTimecode,focusVolts,zoomVolts,irisVolts\n");
	}
	else {
		$Response->write("RovName,DiveNumber,DateTimeGMT,betaTimecode,hdTimecode,mainFocus,mainZoom\n");
	}
	while ( !$RS->EOF ) { 
		my $record = '';
		foreach my $f ( @fields ) { 
			$record .= $RS->Fields($f)->value . ",";
		}
		$record =~ s/,$/\n/;
		$Response->write($record);
		$RS->MoveNext; 
	} 

} # End returnRecords()
%>