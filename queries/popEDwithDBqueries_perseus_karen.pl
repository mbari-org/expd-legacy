# Program to populate ExpeditionData table in Expd with links to
# queries from web-accessible databases.
#
# Mike McCann MBARI
# 1 Oct 1999
# Made to work with Solstice
# 5/14/03
# Neil Conner
# Use SSL certificate bundle following April 2017 typhoon upgrade - Mike McCann -26 September 2017
#
# Updates for SQL Server 2016:
# This perl script renamed to include the SQL Server name (popEDwithDBqueries_perseus.pl) to facilitate
# testing as well as future database server upgrades.
# Database connection string updated for perseus.
# Neil Conner 4/2/2018

# 12-Jul-2018 Changed splserver authentication to Integrated Security=SSPI  -rschramm


=head1 NAME

popEDwithDBqueries_perseus.pl - Populate ExpeditionData table with Database queries

=head1 SYNOPSIS

    > perl popEDwithDBqueries_perseus.pl <startExpeditionID> [endExpeditionID] | <last(ndays)>

=head1 DESCRIPTION

This script steps though all the expeditions defined by the calling arguments
picks of the start & end times then constructs a URL with these times to various
web accessible databases (e.g. SamplesDB, Expd, Rovctd, etc.).  This is done for both
the expedtion times and the dives associated with each expedition.


The process is basically as follows:

1. Loop through the expeditions, query DBs, collect URL, LinkText, and foreign keys

2. Loop through the dives of each expedition, query DBs, collect URL, LinkText, and foreign keys

3. Loop through each of the collected records and insert or update them into the ExpeditionData table using a stored procedure.


=cut

use LWP::Simple;
use LWP::UserAgent;
use OLE;
use Carp;
use SSDS;


##$_dsn = 'Server=perseus;DATABASE=expd;UID=;Password=;';
#$_dsn = 'Server=perseus;DATABASE=expd;Integrated Security=SSPI';
$_dsn = 'Server=fog;DATABASE=expd_ks;Integrated Security=SSPI';
$debug = 0;


#
# Deal with command line options, if none then set start & end year to this year
#
$pltfrm = $ARGV[0];
if ( $ARGV[0] =~ /^\d+/ ) {
	$sExpdID = $ARGV[0];
	if ( $ARGV[1] =~ /\d+/ ) {
		$eExpdID = $ARGV[1];
	}
}
elsif ( $ARGV[0] =~ /last(\d+)/ ) {
	$daysago = $+;	
}
else {
	die "\nUsage: $0 <startExpeditionID> [endExpeditionID] | <last(ndays)>\n\n";
}

$eExpdID = $sExpdID unless $eExpdID;

print "\n$0 run started on ", scalar localtime, "\nfor $ARGV[0] ";
print "to $ARGV[1] " if $ARGV[1];
print "\n-------------------------------------------------------------------\n";

#
# Open the Expd database, check for errors
#
$Conn = CreateObject OLE "ADODB.Connection";
$Conn->{'Provider'} = "sqloledb";
$Conn->Open($_dsn);
if(!Conn) {
   	$Errors = $Conn->Errors();
    	print "Error in opening connection to $dsn:\n";
    	foreach $error (keys %$Errors) {
            	print $error->{Description}, "\n";
    	}
    	croak "Empty Connection object.";
} # End if(!Conn)

if ( $daysago > 0 ) {
	#
	# Initial query on DB to get list of expeditionIDs to go through
	#
	$sql = "SELECT ExpeditionID FROM Expedition WHERE ";
	$sql .= "( DateDiff(dy, ScheduledStartDtg, getdate()) < $daysago ) AND ";
	$sql .= "( DateDiff(dy, ScheduledStartDtg, getdate()) >= 0 ) ";
	$sql .= "ORDER BY ScheduledStartDtg";

	print "Executing: ", $sql, "\n";
	my $RSlist = $Conn->Execute($sql);
	if(!$RSlist) {
		print "\n$0: Empty Return Set: RSlist from:\n$_dsn\nfor:\n$sql\n\n";
	}
	else {
		while ( ! $RSlist->EOF ) {
			push @eIDlist, $RSlist->Fields('ExpeditionID')->Value;
			$RSlist->MoveNext();
		}
	}
	$RSlist->close();
}
else {
	@eIDlist = ($sExpdID..$eExpdID);
}
	
#
# Get urls & check for expeditions & dives
#
undef @ExpeditionDataTable;

print "Going through Expeditions...\n";
foreach $eID (@eIDlist) {
	print "eID = $eID\n";
	
	$LinkText = '';
	$url = '';
	
	#
	# Query DB for Expedition scheduled and actual times and year and year-day
	#
	$sql = "SELECT StartDtg, EndDtg, ScheduledStartDtg, ScheduledEndDtg, ShipName, ";
	$sql .= " DATEPART(yy, ScheduledStartDtg) AS ssyyyy, DATEPART(dy, ScheduledStartDtg) AS ssddd, ";
	$sql .= " DATEPART(yy, StartDtg) AS asyyyy, DATEPART(dy, StartDtg) AS asddd, ";
	$sql .= " DATEPART(yy, ScheduledEndDtg) AS seyyyy, DATEPART(dy, ScheduledEndDtg) AS seddd, ";
	$sql .= " DATEPART(yy, EndDtg) AS aeyyyy, DATEPART(dy, EndDtg) AS aeddd ";
	$sql .= " FROM Expedition WHERE ExpeditionID = $eID";
	print "Executing: ", $sql, "\n" if $debug;
	my $RSexpd = $Conn->Execute($sql);
	if(!$RSexpd) {
 		print "\n$0: Empty Return Set: RSexpd from:\n$_dsn\nfor:\n$sql\n\n";
	}
	else {
		$estart = $RSexpd->Fields(StartDtg)->Value;
		$eend = $RSexpd->Fields(EndDtg)->Value;
		if ( ! $estart || ! $eend ) {
			$estart = $RSexpd->Fields(ScheduledStartDtg)->Value;
			$eend = $RSexpd->Fields(ScheduledEndDtg)->Value;
		}
		##print "estart = $estart, eend = $eend\n";
					
		#
		# Collect links that apply to the Expedition:
		#
		
		#
		# Loop over year-days of the Expedition and add links 
		#
		$syyyy = (defined $RSexpd->Fields(asyyyy)->Value) ? $RSexpd->Fields(asyyyy)->Value : $RSexpd->Fields(ssyyyy)->Value;
		$sddd = (defined $RSexpd->Fields(asddd)->Value) ? $RSexpd->Fields(asddd)->Value : $RSexpd->Fields(ssddd)->Value;
		$eyyyy = (defined $RSexpd->Fields(aeyyyy)->Value) ? $RSexpd->Fields(aeyyyy)->Value : $RSexpd->Fields(seyyyy)->Value;
		$eddd = (defined $RSexpd->Fields(aeddd)->Value) ? $RSexpd->Fields(aeddd)->Value : $RSexpd->Fields(seddd)->Value;
		if ( $syyyy == $eyyyy ) {
			$yyyy = $syyyy;
		}
		else {
			print "syyy ($syyyy) != eyyyy ($eyyyy).  Setting yyyy to $syyyy.\n";
			$yyyy = $syyyy;
		}
		for ( my $ddd = $sddd; $ddd <= $eddd; $ddd++ ) {
			#
			# - Collect rovctd data that has made it into the database from the rovctdlogr file (replaces the original static Der.txt link)
			#   Use acutal (a) times if they've been entered in the database, otherwise use scheduled (s) times
			#
			$DataType = 'rovctd';
			print "Calling get_rovctd_link(): yyyy = $yyyy, ddd = $ddd\n" if $debug;
			
			($url,$LinkText,$LinkComment,$Place) = get_rovctd_link(	
									ShipName => $RSexpd->Fields(ShipName)->Value, 
									DeviceID => $RSdive->Fields(DeviceID)->Value,
									ExpeditionID_FK => $RSdive->Fields(ExpeditionID_FK)->Value,
									yyyy => $yyyy,
									ddd  => $ddd ) ;
			$str = "$DataType,,";
			$str .= ",$eID,$url,$LinkText,$LinkComment,$UnixPath,$WindowsPath,$Place";
			$str .= ",$0";
			print "str = $str\n" if $debug;
			push @ExpeditionDataTable, $str if $LinkText && $url;
		}
		
		#
		# - Collect any Survey AUVCTD data links ($LinkInfo is "$url,$LinkText,$LinkComment")
		#
		$DataType = 'auvctdSurvey';
		($url,$LinkText,$LinkComment,$SurvNum,$Place,$prID) = get_auvctd_link(	
										ShipName => $RSexpd->Fields(ShipName)->Value, 
										StartDtg => $estart,
										EndDtg => $eend ) ;
		
		$str = "$DataType,drdo,$SurvNum";
		$str .= ",$eID,$url,$LinkText,$LinkComment,$UnixPath,$WindowsPath,$Place";
		$str .= ",$0";
		print "str = $str\n" if $debug;
		push @ExpeditionDataTable, $str if $LinkText && $url;
		
		# Add another link to the last data products
		$url = "viewSSDS.asp?ProcessRunId=" . $prID;
		$str = "$DataType,drdo,$SurvNum";
		$str .= ",$eID,$url,Data Products from SSDS,ProcessRuns, DataContainers, and Resources,$UnixPath,$WindowsPath,100";
		$str .= ",$0";
	
	}
	
	#
	# Find Dives for this expedition
	#my $ExpdID = $RS->Fields('ExpeditionID_FK')->value if $RS;

	# Check to see if Non-MBARI ship where ExpeditionID_FK = 0, otherwise look for MBARI ships ExpeditionIDs
	if ( $rovName = 'mini' ) && ( $ExpeditionID_FK == 0 ) {
		$sql = "SELECT * FROM Dive WHERE ExpeditionID_FK = 0 ORDER BY DiveNumber";
	}
	else {
	    $sql = "SELECT * FROM Dive WHERE ExpeditionID_FK = $eID ORDER BY DiveNumber";
    }
	print "Executing: ", $sql, "\n" if $debug;
	my $RSdive = $Conn->Execute($sql);
	if(!$RSdive) {
 		print "\n$0: Empty Return Set: RSdive from:\n$_dsn\nfor:\n$sql\n\n";
	}
	else {
		while ( ! $RSdive->EOF ) {
			$dstart = $RSdive->Fields(DiveStartDtg)->Value;
			$dend = $RSdive->Fields(DiveEndDtg)->Value;
			
			print "Dive: " . $RSdive->Fields(RovName)->Value . $RSdive->Fields(DiveNumber)->Value;
			print "\ndnumber = " . $RSdive->Fields(DiveNumber)->Value . ", dstart = $dstart, dend = $dend\n" if $debug;
			
			#
			# Collect ROVCTD Links ($LinkInfo is "$url,$LinkText,$LinkComment")
			#
			$DataType = 'rovctddb';
			($url,$LinkText,$LinkComment) = get_rovctddb_link(	
											RovName => $RSdive->Fields(RovName)->Value, 
											DiveNumber =>  $RSdive->Fields(DiveNumber)->Value );
			##print "\nmain(): rovctd Place = $Place\n";
			$str = "$DataType," . $RSdive->Fields(RovName)->Value . "," . $RSdive->Fields(DiveNumber)->Value;
			$str .= ",$eID,$url,$LinkText,$LinkComment,$UnixPath,$WindowsPath,$Place";
			$str .= ",$0";
			@tStr = split('\s+', $LinkText);
			push @ExpeditionDataTable, $str;
			print " ctd: " . $tStr[0];
			
			# 
			# For CTD Data add links for the time series (ts) and profile (pr) dynamic plots
			#
			##### disabled, 18Jun2015 rschramm
			####if ( $LinkText ) {
			####	push @ExpeditionDataTable, $str;
			####	
			####	$tsPlotUrl = "http://mww.mbari.org:8080/rovctd/pages/stackedtimeseries.jsp?platform=";
			####	$tsPlotUrl .= $RSdive->Fields(RovName)->Value . "&dive=";
			####	$tsPlotUrl .= $RSdive->Fields(DiveNumber)->Value . "&domain=epochsecs&noHTMLHeader=0&r1=p&r2=t&r3=s&r4=light&r5=o2";
			####	$tsPlotText = "Time Series Plot for " . $RSdive->Fields(RovName)->Value . $RSdive->Fields(DiveNumber)->Value;
			####	$Place = $Place + 100;		# Have all plot links follow data links
			####	$tsPlotStr = "$DataType," . $RSdive->Fields(RovName)->Value . "," . $RSdive->Fields(DiveNumber)->Value;
			####	$tsPlotStr .= ",$eID,$tsPlotUrl,$tsPlotText ,$LinkComment,$UnixPath,$WindowsPath,$Place";
			####	$tsPlotStr .= ",$0";
			####	##print "\nmain(): pushing $tsPlotStr\n";
			####	push @ExpeditionDataTable, $tsPlotStr;
			####	
			####	$prPlotUrl = "http://mww.mbari.org:8080/rovctd/pages/profile.jsp?platform=";
			####	$prPlotUrl .= $RSdive->Fields(RovName)->Value . "&dive=";
			####	$prPlotUrl .= $RSdive->Fields(DiveNumber)->Value . "&domain=p&noHTMLHeader=0&r1=t&r2=s&r3=o2&r4=light";
			####	$prPlotText = "Profile Plot for " . $RSdive->Fields(RovName)->Value . $RSdive->Fields(DiveNumber)->Value;
			####	$Place = $Place + 200;		# Have all plot links follow data links
			####	$prPlotStr = "$DataType," . $RSdive->Fields(RovName)->Value . "," . $RSdive->Fields(DiveNumber)->Value;
			####	$prPlotStr .= ",$eID,$prPlotUrl,$prPlotText ,$LinkComment,$UnixPath,$WindowsPath,$Place";
			####	$prPlotStr .= ",$0";
			####	##print "\nmain(): pushing $prPlotStr\n";
			####	push @ExpeditionDataTable, $prPlotStr;
			####}
			
			#
			# Collect SamplesDB Links ($LinkInfo is "$url,$LinkText,$LinkComment")
			#
			$DataType = 'samplesdb';
			($url,$LinkText,$LinkComment) = get_samples_link(	
											ShipName => $RSexpd->Fields(ShipName)->Value, 
											RovName => $RSdive->Fields(RovName)->Value, 
											DiveNumber =>  $RSdive->Fields(DiveNumber)->Value,
											StartDtg =>  $RSdive->Fields(DiveStartDtg)->Value,
											EndDtg =>  $RSdive->Fields(DiveEndDtg)->Value ) ;
			$str = "$DataType," . $RSdive->Fields(RovName)->Value . "," . $RSdive->Fields(DiveNumber)->Value;
			$str .= ",$eID,$url,$LinkText,$LinkComment,$UnixPath,$WindowsPath,$Place";
			$str .= ",$0";
			@tStr = split('\s+', $LinkText);
			print " samples: " . $tStr[0];
			push @ExpeditionDataTable, $str if $LinkText;
			
			#
			# Collect VIMS Annotations Links ($LinkInfo is "$url,$LinkText,$LinkComment")
			#
			$DataType = 'annotations';
			($url,$LinkText,$LinkComment) = get_annotations_link(	
											ShipName => $RSexpd->Fields(ShipName)->Value, 
											RovName => $RSdive->Fields(RovName)->Value, 
											DiveNumber =>  $RSdive->Fields(DiveNumber)->Value,
											StartDtg =>  $RSdive->Fields(DiveStartDtg)->Value,
											EndDtg =>  $RSdive->Fields(DiveEndDtg)->Value ) ;
			$str = "$DataType," . $RSdive->Fields(RovName)->Value . "," . $RSdive->Fields(DiveNumber)->Value;
			$str .= ",$eID,$url,$LinkText,$LinkComment,$UnixPath,$WindowsPath,$Place";
			$str .= ",$0";
			@tStr = split('\s+', $LinkText);
			print " anno: " . $tStr[0];
			push @ExpeditionDataTable, $str if $LinkText;
			
			
			
			#
			# Collect Camlog database Links ($LinkInfo is "$url,$LinkText,$LinkComment")
			#
			$DataType = 'camlogdb';
			($url,$LinkText,$LinkComment) = get_camlogdb_link(	
											ShipName => $RSexpd->Fields(ShipName)->Value, 
											RovName => $RSdive->Fields(RovName)->Value, 
											DiveNumber =>  $RSdive->Fields(DiveNumber)->Value,
											StartDtg =>  $RSdive->Fields(DiveStartDtg)->Value,
											EndDtg =>  $RSdive->Fields(DiveEndDtg)->Value ) ;
			$str = "$DataType," . $RSdive->Fields(RovName)->Value . "," . $RSdive->Fields(DiveNumber)->Value;
			$str .= ",$eID,$url,$LinkText,$LinkComment,$UnixPath,$WindowsPath,$Place";
			$str .= ",$0";
			@tStr = split('\s+', $LinkText);
			print " camlog: " . $tStr[0];
			push @ExpeditionDataTable, $str if $LinkText;	
			
			
			#
			# Collect CleanNav database Links ($LinkInfo is "$url,$LinkText,$LinkComment")
			#
			$DataType = 'CleanNavdb';
			($url,$LinkText,$LinkComment) = get_CleanNavdb_link(	
											ShipName => $RSexpd->Fields(ShipName)->Value, 
											RovName => $RSdive->Fields(RovName)->Value, 
											DiveNumber =>  $RSdive->Fields(DiveNumber)->Value,
											StartDtg =>  $RSdive->Fields(DiveStartDtg)->Value,
											EndDtg =>  $RSdive->Fields(DiveEndDtg)->Value ) ;
			$str = "$DataType," . $RSdive->Fields(RovName)->Value . "," . $RSdive->Fields(DiveNumber)->Value;
			$str .= ",$eID,$url,$LinkText,$LinkComment,$UnixPath,$WindowsPath,$Place";
			$str .= ",$0";
			@tStr = split('\s+', $LinkText);
			print " CleanNav: " . $tStr[0];
			push @ExpeditionDataTable, $str if $LinkText;	
			
			
			#
			# Collect RawNav database Links ($LinkInfo is "$url,$LinkText,$LinkComment")
			#
			$DataType = 'RawNavdb';
			($url,$LinkText,$LinkComment) = get_RawNavdb_link(	
											ShipName => $RSexpd->Fields(ShipName)->Value, 
											RovName => $RSdive->Fields(RovName)->Value, 
											DiveNumber =>  $RSdive->Fields(DiveNumber)->Value,
											StartDtg =>  $RSdive->Fields(DiveStartDtg)->Value,
											EndDtg =>  $RSdive->Fields(DiveEndDtg)->Value ) ;
			$str = "$DataType," . $RSdive->Fields(RovName)->Value . "," . $RSdive->Fields(DiveNumber)->Value;
			$str .= ",$eID,$url,$LinkText,$LinkComment,$UnixPath,$WindowsPath,$Place";
			$str .= ",$0";
			@tStr = split('\s+', $LinkText);
			print " RawNav: " . $tStr[0];
			push @ExpeditionDataTable, $str if $LinkText;	
			
			
			print "\n";
			
			$RSdive->MoveNext;
		} # End while
		
		$RSdive->Close;
		
	} # End if
	
	$RSexpd->Close;
	
} # End for ($eID = ...

	

#
# Go through constructed list and insert each record into the table
#
print "\nAttempting to insert ", $#ExpeditionDataTable+1, " records into ExpeditionData Table...\n";
print "(x: already in table, .: succesful insert, o:modified by same program, -:no return from SP, probably successful insert)\n";
print "-------------------------------------------\n";
foreach (@ExpeditionDataTable) {

	my $sql = "exec insertExpeditionData ";
	my $i = 0;
	foreach $field ( split(',', $_) ) {
		if ($i == 2 || $i == 3 || $i == 9) {	# These are ints
			$sql .= ($field > 0) ? "$field," : "0,";
		}
		else {
			$sql .= "'$field',";
		}
		$i++;
	} # End foreach $field (...
	
	$sql =~ s/,$//;
	
	print "Executing sql:\n$sql\n\n" if ($debug);
	$RS = $Conn->Execute($sql);
	if(!$RS) {
 		print "\n$0: Empty Return Set: RS from:\n$_dsn\nfor:\n$sql\n\n";
	}
	else {
		my $retval = 2;
		$retval = $RS->Fields(0)->Value if $RS->Fields(0);
		if ( $retval < 0 ) {
			print "x";
		}
		elsif ( $retval == 1 ) {
			print "o";
		}
		elsif ( $retval == 2 ) {
			print "-";
		}
		else {
			print ".";
		}

		$RS->Close;
	}
	
} # End foreach (@Expe...

$Conn->Close;

#--------------------------------------------------------------------
#

=head3 get_mww_records(url)

Uses LWP UserAgent to get Response from mww for given url, count the number of records and return that value


=cut

sub get_mww_records {
	my $url = $_[0];
	
	#
	# Use Certificate now - mpm 26 September 2017
	#
	$ua = LWP::UserAgent->new(
		##ssl_opts => {
		##	SSL_cert_file   => 'C:\\local\\ssl\\ca-bundle.crt',
		##	verify_hostname => 1,
		##}
	);
	$req = HTTP::Request->new(GET => "$url");
	$req->authorization_basic('SHORE\rovctd', '2020rov');
	
	my $resp = $ua->request($req);
	my $num_records = 0;
	if ($resp->is_success) {
		$content = $resp->content;
		@records = split('\n', $content);
		print "\n\nNUMBER OF RECORDS = $#records\n" if $debug;
	}
	else {
		print "Failed to GET $url\n";
		print "\n\nError: " . $resp->status_line . "\n";
        print $resp->as_string;
	}

	return \@records;
}


#--------------------------------------------------------------------
#

=head3 get_rovctddb_link()

Uses LWP to test urls to database queries and counts the number of records returned.


=cut

sub get_rovctddb_link {

	my %hash = @_;
	my $RovName = $hash{RovName};
	my $DiveNumber = $hash{DiveNumber};
	my $baseurl = "http://mww.mbari.org";
	
	my $url = "/expd/queries/rovctd-data_karen.asp?date_type=dive&return_type=comma";
	$url .= "&pltfrm=" . $RovName;
	$url .= "&dive=" . $DiveNumber;
	$Place = $DiveNumber - 1000000;		# Will we ever do more than a million dives?
	##print "\nget_rovctddb_link(): Place = $Place\n";
		
	#
	# Test the link
	#
	print "get_rovctddb_link(): testing link $baseurl$url\n" if $debug;

	my $records_ref = get_mww_records("$baseurl$url");
	
	# Subtract the heeader line
	my $num_records = $#$records_ref - 1;
	
	my $LinkText = '';
	my $LinkComment = '';
	
	if ($num_records > 0) {
		$LinkText = "$num_records records for dive " . $RovName . $DiveNumber;
		$LinkComment = '';
	}

	return ($url,$LinkText,$LinkComment);

} # End get_rovctddb_link()

#--------------------------------------------------------------------
#

=head3 get_samples_link()

Uses LWP to test urls to database queries and counts the number of records returned.


=cut

sub get_samples_link {

	my %hash = @_;
	my $ShipName = $hash{ShipName};
	my $StartDtg = $hash{StartDtg};
	my $EndDtg = $hash{EndDtg};
	my $RovName = $hash{RovName};
	my $DiveNumber = $hash{DiveNumber};
	
	my $url = "htt";			# God, I hate Interdev hyperlink full urls!
	##$url .= "p://godzilla.mbari.org/SamplesDB/queries/qryExpdDate.idc?";
	$url .= "p://mww.mbari.org/SamplesDB/queries/returnTable_karen.asp?";
	$url .= "date1=" . urlencode($StartDtg);
	$url .= "&date2=" . urlencode($EndDtg);
	$url .= "&ShipName=Point+Lobos" if $ShipName eq 'ptlo';
	$url .= "&ShipName=Western+Flyer" if $ShipName eq 'wfly';
	$url .= "&ShipName=Rachel+Carson" if $ShipName eq 'rcsn';
	
	$Place = $DiveNumber;
		
	#
	# Test the link
	#
	print "get_samples_link(): trying url...\n$url\n" if $debug;

	my $records_ref = get_mww_records("$url");
	
	$num_records = 0;
	foreach (@$records_ref) {
		$num_records++ if /selectSample_karen.asp\?ID/;
	}
	
	my $LinkText = '';
	my $LinkComment = '';
	if ( $nsamples > 0 ) {
		$LinkText = "$num_records sample";
		$LinkText .= "s" if $nsamples > 1;
		$LinkText .= " for dive " . $RovName . $DiveNumber;
		$LinkComment = '';
	}
	
	##print "\nget_samples_link(): returned string = $url,$LinkText,$LinkComment\n";
	return ($url,$LinkText,$LinkComment);

} # End get_samples_link()

#--------------------------------------------------------------------
#

=head3 get_annotations_link()

Uses LWP to test urls to database queries and counts the number of records returned.


=cut

sub get_annotations_link {

	my %hash = @_;
	my $ShipName = $hash{ShipName};
	my $StartDtg = $hash{StartDtg};
	my $EndDtg = $hash{EndDtg};
	my $RovName = $hash{RovName};
	my $DiveNumber = $hash{DiveNumber};
	
	my $url = "htt";			# God, I hate Interdev hyperlink full urls!
	$url .= "p://mww.mbari.org/expd/queries/annotations.asp?";
	$url .= "qSCollDTG=" . urlencode($StartDtg);
	$url .= "&qECollDTG=" . urlencode($EndDtg);
	$url .= "&qShipName=ptlo" if $ShipName eq 'ptlo';
	$url .= "&qShipName=rcsn" if $ShipName eq 'rcsn';
	$url .= "&qShipName=wfly" if $ShipName eq 'wfly';
	$url .= "&qDiveNameNum=" . $RovName . $DiveNumber;
	
	$Place = $DiveNumber;
		
	#
	# Test the link
	#
	print "get_annotations_link(): trying url...\n$url\n" if $debug;

	my $records_ref = get_mww_records("$url");
	
	my $nannotations = 0;
	foreach (@$records_ref) {
		$nannotations++ if /\<tr\>/;
	}
	print "\nnannotations = $nannotations\n" if $debug;
	$nannotations = $nannotations - 1;
	my $LinkText = '';
	my $LinkComment = '';
	if ( $nannotations > 0 ) {
		$LinkText = "$nannotations annotations";
		$LinkText .= " for dive " . $RovName . $DiveNumber;
		$LinkComment = '';
	}
	
	print "\nget_annotations_link(): returned string = $url\n$LinkText,$LinkComment\n" if $debug;
	return ($url,$LinkText,$LinkComment);

} # End get_annotations_link()



#--------------------------------------------------------------------
#

=head3 get_camlogdb_link()

Uses LWP to test urls to database queries and counts the number of records returned.


=cut

sub get_camlogdb_link {

	my %hash = @_;
	my $RovName = $hash{RovName};
	my $DiveNumber = $hash{DiveNumber};
	my $baseurl = "http://mww.mbari.org";
	
	my $url = "/expd/queries/camlog-data.asp?ROVName=$RovName&DiveNumber=$DiveNumber";
	$Place = $DiveNumber - 1000000;		# Will we ever do more than a million dives?
	##print "\nget_camlogdb_link(): Place = $Place\n";
		
	#
	# Test the link
	#
	print "get_camlogdb_link(): testing link $baseurl$url\n" if $debug;

	my $records_ref = get_mww_records("$baseurl$url");
	
	my $LinkText = '';
	my $LinkComment = '';
	if ( $#$records_ref > 15 ) {
		$LinkText = "$#records records for dive " . $RovName . $DiveNumber;
		$LinkComment = '';
	}
	##print "get_camlogdb_link(): Num records = $#records\n" if $debug;
	return ($url,$LinkText,$LinkComment);

} # End get_camlogdb_link()


#--------------------------------------------------------------------
#

=head3 get_CleanNavdb_link()

Uses LWP to test urls to database queries and counts the number of records returned.


=cut

sub get_CleanNavdb_link {

	my %hash = @_;
	my $RovName = $hash{RovName};
	my $DiveNumber = $hash{DiveNumber};
	my $baseurl = "http://mww.mbari.org";
	
	my $url = "/expd/queries/nav-data_karen.asp?ROVName=$RovName&DiveNumber=$DiveNumber";
	$Place = $DiveNumber - 1000000;		# Will we ever do more than a million dives?
	##print "\nget_CleanNavdb_link(): url = $url\n";
	##print "\nget_CleanNavdb_link(): Place = $Place\n";
		
	#
	# Test the link
	#
	print "get_CleanNavdb_link(): testing link $baseurl$url\n" if $debug;

	my $records_ref = get_mww_records("$baseurl$url");
	
	my $LinkText = '';
	my $LinkComment = '';
	if ( $#$records_ref > 15 ) {
		$num_records = $#$records_ref - 1;
		$LinkText = "$num_records Edited records for dive " . $RovName . $DiveNumber;
		$LinkComment = '';
	}
	##print "get_CleanNavdb_link(): Num records = $#records\n" if $debug;
	return ($url,$LinkText,$LinkComment);

} # End get_CleanNavdb_link()

#--------------------------------------------------------------------
#

=head3 get_RawNavdb_link()

Uses LWP to test urls to database queries and counts the number of records returned.


=cut

sub get_RawNavdb_link {

	my %hash = @_;
	my $RovName = $hash{RovName};
	my $DiveNumber = $hash{DiveNumber};
	my $baseurl = "http://mww.mbari.org";
	
	my $url = "/expd/queries/nav-data.asp?ROVName=$RovName&DiveNumber=$DiveNumber&type=raw";
	$Place = $DiveNumber - 1000000;		# Will we ever do more than a million dives?
	##print "\nget_RawNavdb_link(): Place = $Place\n";
		
	#
	# Test the link
	#
	print "get_RawNavdb_link(): testing link $baseurl$url\n" if $debug;
	
	my $records_ref = get_mww_records("$baseurl$url");

		
	my $LinkText = '';
	my $LinkComment = '';
	if ( $#$records_ref > 15 ) {
		my $num_records = $#$records_ref - 1;
		$LinkText = "$num_records Raw records for dive " . $RovName . $DiveNumber;
		$LinkComment = '';
	}
	print "get_RawNavdb_link(): Num records = $#records\n" if $debug;
	return ($url,$LinkText,$LinkComment);

} # End get_RawNavdb_link()



#--------------------------------------------------------------------
#

=head3 get_rovctd_link()

Query for rovctd data by year-day from database (serves same purpose as old Der.txt static file link)
This link is needed for assigning dive start and END times in the Expedition database.


=cut

sub get_rovctd_link {

	my %hash = @_;
	my $ShipName = $hash{ShipName};
	my $DeviceID = $hash{DeviceID};
	my $ExpeditionID_FK = $hash{ExpeditionID_FK};
	my $yyyy = $hash{yyyy};
	my $ddd = $hash{ddd};
	
	return if $ShipName eq 'zphr';		# The Zephyr (for now) does not have an ROV
	$rovName = 'vnta' if $ShipName eq 'ptlo';
	$rovName = 'tibr' if $ShipName eq 'wfly';
	$rovName = 'docr' if ($ShipName eq 'wfly' && $yyyy > 2008);
	$rovName = 'vnta' if $ShipName eq 'rcsn';
	$rovName = 'mini' if $DeviceID == 1859;
	
	#
	# Get records from the asp trhat looks like this:
	# http://mww.mbari.org/expd/queries/rovctd-data.asp?date_type=yyyyddd&return_type=comma&pltfrm=vnta&year=2008&yday=326
	#
	my $url = 'http://mww.mbari.org/expd/queries/rovctd-data_karen.asp?date_type=yyyyddd&return_type=comma&pltfrm=';
	$url .= $rovName . "&year=" . $yyyy . "&yday=" . $ddd;
	print "get_rovctd_link(): getting records from url = $url\n" if $debug;
	
	#
	# Log into Canyon Head to get return from data base 
	#
	$ua = LWP::UserAgent->new;
	$req = HTTP::Request->new(GET => $url);
	$req->authorization_basic('SHORE\rovctd', '2020rov');
	
	my $LinkText = '';
	my $LinkComment = '';
	my $Place = 5;			# Have this appear after rovctdlogr (1) and Der.txt (4) links
	
	#
	# Return nothing of "no data" found in response
	#
	my $response = $ua->request($req)->as_string;
	return ($url,$LinkText,$LinkComment) if $response =~ /no data was found/;	
	
	@records = split('\n', $response );
	
	if ( $#records > 0) {
		$LinkText = "$#records rovctd records for year-day " . $ddd;
	}
	print "get_rovctd_link(): LinkText = $LinkText\n" if $debug;
	
	return ($url, $LinkText, $LinkComment, $Place);
	
} # End get_rovctd_link()


#--------------------------------------------------------------------
#

=head3 get_auvctd_link()

Uses SSDS to get URLs to survey data products.


=cut

sub get_auvctd_link {

	my %hash = @_;
	my $ShipName = $hash{ShipName};
	my $StartDtg = $hash{StartDtg};
	my $EndDtg = $hash{EndDtg};
	
	return if $ShipName ne 'zphr';		# Only Zephyr in Expd has AUVCTD surveys (this may change)
	
	# Test for before when we correct DB for GMT times
	##$StartDtg =  '7/1/2003';
	##$EndDtg = '7/17/2003';
	
	if ($debug) {
		print "ShipName = $ShipName\n";
		print "StartDtg = $StartDtg\n";
		print "EndDtg = $EndDtg\n";
	}
		
	#
	# For now add 7 hours to SSDS times (until we fix the times in SSDS
	# If either start or end time of AUV run falls into time span of Zephyr cruise find it and look for resources
	#
	my $sql = "SELECT * \n";
	$sql .= " FROM  dbo.DataProducer INNER JOIN \n";
	$sql .= "       dbo.Device ON dbo.DataProducer.DeviceID_FK = dbo.Device.id INNER JOIN \n";
	$sql .= "       dbo.DeviceType ON dbo.Device.DeviceTypeID_FK = dbo.DeviceType.id \n";
	$sql .= " WHERE (dbo.DeviceType.Name = 'AUV') AND ((dbo.DataProducer.StartDateTime BETWEEN '$StartDtg' AND '$EndDtg') \n";
	$sql .= "       OR (dbo.DataProducer.EndDateTime BETWEEN '$StartDtg' AND '$EndDtg'))";
	
	##print "sql = \n$sql\n";

	my $ssds = new SSDS();
        $ssds->ssdsServer("http://polyp:8080/");
        my $deplAccess = new SSDS::DeploymentAccess();
        my $resAccess = new SSDS::ResourceAccess();
        
        ##$deplAccess->debug(3);
	$depls = $deplAccess->findBySQL($sql);
	my $survStr = '';
	if ( defined $depls ) {
		my @deplAttr = ${$depls}[0]->get_attribute_names();
		
		foreach my $depl (@{$depls}) {
			##print "\nAttributes of Deployment: " . $depl->name() . "\n";
			##foreach my $a (@deplAttr) {
			##	next unless $depl->$a;
			##	print "  $a = " . $depl->$a . "\n";
			##}
			##print "\n";
			$depl->name() =~ /(\d\d\d\d)\.(\d\d\d)\.(\d\d)/;
			$survStr = "${1}_${2}";
		}
		
		#
		# Find Resources for placing in Expd
		#
		if ($debug) {
			print "Looking for resources with names like $survStr...\n";
		}
		##$resAccess->debug(3);
		$ress = $resAccess->findByLikeName($survStr);
		if ( defined $ress ) {
			foreach my $res (@{$ress}) {
				##print "\nAttributes of Resource: " . $res->name() . "\n";
				##foreach my $a ($res->get_attribute_names()) {
				##	next unless $res->$a;
				##	print "  $a = " . $res->$a . "\n";
				##}
				##print "\n";
				if ($res->name() =~ /biolume\.png/ ) {
					$SurvNum = $survStr;
					$SurvNum =~ s/_//;
					$Place = 1;
					$url = $res->url();
					$LinkText = $res->name();
					$LinkComment = '';
					last;
				}
			}
		}
		
		#
		# Find Full resolution Matlab data file for the survey and find its creator's ID
		#
		$dfs = $dataFileAccess->findByLikeName($survStr);
		if ( defined $dfs ) {
			foreach my $res (@{$dfs}) {
			}
		}
			
	}
	else {
		print "No AUV dives found for $ShipName cruise starting on $StartDtg\n";
	}
	
	return ($url,$LinkText,$LinkComment,$SurvNum,$Place,$prID);

} # End get_auvctd_link()




#--------------------------------------------------------------------
#

=head3 urlencode()

Take a string (mostly DTG) and urlencode it.


=cut

sub urlencode {
	my $in = $_[0];
	$in =~ s/\s/+/g;
	$in =~ s#/#%2F#g;
	$in =~ s#:#%3A#g;
	
	return $in;
	
} # End urlencode()

__END__


=head1 SEE ALSO

	LWP

=head1 AUTHOR

Mike McCann F<E<lt>mccann@mbari.orgE<gt>>

This script was developed as part of the MBARI 


http://www.mbari.org/~mccann

