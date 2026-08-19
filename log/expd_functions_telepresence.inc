
<% 
=head1 NAME

expd_functions_telepresence.inc - Common included functions for Expedition Log apps

=head1 SYNOPSIS

    https://mww.mbari.org/expd/log/expd_functions_telepresence.inc

=head1 DESCRIPTION

Standard functions used by pre- and post- cruise asps.

Mike McCann MBARI

November 1998

=head1 FUNCTIONS

=cut

#--------------------------------------------------------------------
#

=head3 open_database()

Create the database connection object.  Receive the DSN.
Does a little error checking. $Conn is global.

Author: Mike McCann

Date Created: 11/3/8

=cut

sub open_database {

	my $dsn = $_[0];
	
	my $_debug = 0;
	Win32::ASP::DebugPrint("\nopen_database(): passed in dsn = '$dsn'\n") if $_debug;
	
	$Application->{'DSN'} = "Server=perseus;Database=expd;UID=***;******;" unless $dsn;
	#$Application->{'DSN'} = "Server=fog;Database=expd_ks;UID=***;******;" unless $dsn;	
	$dsn = $Application->{'DSN'} unless $dsn;
	Win32::ASP::DebugPrint("\nopen_database(): dsn = '$dsn'\n") if $_debug;
	
	$Conn = CreateObject OLE "ADODB.Connection";
	##$Conn->{'Provider'} = $Application->{'providerStr'};
	$Conn->{'Provider'} = "sqloledb";
	$Conn->Open($dsn);
	
	Win32::ASP::AddDeathHook( sub { $Conn->Close } );	# In case asp dies do some cleanup
	
	if(keys %{$Conn->Errors()}) {
	        $Errors = $Conn->Errors();
	        print "<B>Error in opening connection to $dsn:</B>";
	        foreach $error (keys %$Errors) {
	                print $error->{Description}, "\n";
	        }
	        die "<BR>Empty Connection object.";
	}

}	# End open_database()


#--------------------------------------------------------------------
#

=head3 open_vars_database()

Create the database connection object.  Receive the DSN.
Does a little error checking. $ConnVARS is global.

Author: Mike McCann

Date Created: 8/8/2005

=cut

sub open_vars_database {

	my $dsn = $_[0];
	
	Win32::ASP::DebugPrint("\nopen_vars_database(): passed in dsn = '$dsn'\n");
	
	$dsn = "Server=cassiopeia,1433;Database=VARS;UID=***;******;" unless $dsn;	
	Win32::ASP::DebugPrint("\nopen_vars_database(): using dsn = '$dsn'\n");
	
	# VARS still used for Tapes - Mike McCann 6 March 2019
	$ConnVARS = CreateObject OLE "ADODB.Connection";
	$ConnVARS->{'Provider'} = "sqloledb";
	$ConnVARS->Open($dsn);
	
	# Also open a connection to the M3_ANNOTATIONS database - Mike McCann 6 March 2019
	$dsn_m3 = "Server=perseus,1433;Database=M3_ANNOTATIONS;UID=***;******;";
	#$dsn_m3 = "Server=fog,1433;Database=M3_ANNOTATIONS;UID=***;******;";
	$ConnM3 = CreateObject OLE "ADODB.Connection";
	$ConnM3->{'Provider'} = "sqloledb";
	$ConnM3->Open($dsn_m3);
	
	Win32::ASP::AddDeathHook( sub { $ConnVARS->Close } );	# In case asp dies do some cleanup
	Win32::ASP::AddDeathHook( sub { $ConnM3->Close } );	# In case asp dies do some cleanup
	
	if(keys %{$ConnVARS->Errors()}) {
	        $Errors = $ConnVARS->Errors();
	        print "<B>Error in opening VARS database connection to $dsn:</B>";
	        foreach $error (keys %$Errors) {
	                print $error->{Description}, "\n";
	        }
	        die "<BR>Empty Connection object.";
	}
	if(keys %{$ConnM3->Errors()}) {
	        $Errors = $ConnM3->Errors();
	        print "<B>Error in opening VARS database connection to $dsn_m3:</B>";
	        foreach $error (keys %$Errors) {
	                print $error->{Description}, "\n";
	        }
	        die "<BR>Empty Connection object.";
	}

}	# End open_vars_database()


#--------------------------------------------------------------------
#

=head3 open_samples_database()

Create the database connection object.  Receive the DSN.
Does a little error checking. $ConnSamples is global.

Author: Mike McCann

Date Created: 8/8/2005

=cut

sub open_samples_database {

	my $dsn = $_[0];
	
	Win32::ASP::DebugPrint("\nopen_samples_database(): passed in dsn = '$dsn'\n");
	
	$dsn = "Server=perseus;Database=MBARI_Samples;UID=***;******;" unless $dsn;
	#$dsn = "Server=fog;Database=MBARI_Samples;UID=***;******;" unless $dsn;
	Win32::ASP::DebugPrint("\nopen_samples_database(): using dsn = '$dsn'\n");
	
	$ConnSamples = CreateObject OLE "ADODB.Connection";
	$ConnSamples->{'Provider'} = "sqloledb";
	$ConnSamples->Open($dsn);
	
	Win32::ASP::AddDeathHook( sub { $ConnSamples->Close } );	# In case asp dies do some cleanup
	
	if(keys %{$ConnSamples->Errors()}) {
	        $Errors = $ConnSamples->Errors();
	        print "<B>Error in opening VARS database connection to $dsn:</B>";
	        foreach $error (keys %$Errors) {
	                print $error->{Description}, "\n";
	        }
	        die "<BR>Empty Connection object.";
	}

}	# End open_samples_database()



#--------------------------------------------------------------------
#

=head3 FixString()

Put quotes around text values and nothing around numbers
escape embedded quotes.

Author: Mike McCann

Date Created: 10/28/98

=cut

sub FixString {
	
#
# String, type
#

local ($str, $type) = @_;

if ($type =~ /text/i ) {
	$str =~ s/\'/\'\'/g;
	$fixed_string = "'$str'";
}

if ($type =~ /num/i || $str eq 'NULL' ) {
	$fixed_string = "$str";
}

if ($type =~ /num/i && $str eq '' ) {
	$fixed_string = "NULL";
}


return $fixed_string;

}	# End FixString()

#--------------------------------------------------------------------
#

=head3 toGMT()

Convert local time array to a GMT string that is good
for a database DTG load.

Author: Mike McCann

Date Created: 10/28/98

=cut

sub toGMT {

	my @l = @_;

	return -1 unless @l;

	$secs = timelocal(@l);
	
	@g = gmtime($secs);
	
	my $mon = ($g[4] < 10) ? "0" . $g[4] + 1 : $g[4] + 1;
	my $day = ($g[3] < 10) ? "0" . $g[3] : $g[3];
	my $hr12 = (($g[2] > 12) ? $g[2] - 12 : $g[2]);
	my $hr = ($hr12 < 10 ) ? "0" . $hr12 : $hr12;
	my $min = ($g[1] < 10 ) ? "0" . $g[1] : $g[1];
	my $sec = ($g[0] < 10 ) ? "0" . $g[0] : $g[0];

	$GMTstr = $g[5] + 1900 . "-$mon-$day $hr:$min:$sec " .  (($g[2] > 11) ? "PM" : "AM");

	return $GMTstr;

}	# End toGMT()

#--------------------------------------------------------------------
#

=head3 toGMTfromES()

Convert local time array to a GMT string that is good
for a database DTG load.  Should have overloaded toGMT, but
I feel this is safer.

Author: Mike McCann

Date Created: 7/19/99

=cut

sub toGMTfromES {

	my $es = $_[0];

	return -1 unless $es;
	
	@g = gmtime($es);
	
	my $mon = ($g[4] < 10) ? "0" . $g[4] + 1 : $g[4] + 1;
	my $day = ($g[3] < 10) ? "0" . $g[3] : $g[3];
	my $hr12 = (($g[2] > 12) ? $g[2] - 12 : $g[2]);
	my $hr = ($hr12 < 10 ) ? "0" . $hr12 : $hr12;
	my $min = ($g[1] < 10 ) ? "0" . $g[1] : $g[1];
	my $sec = ($g[0] < 10 ) ? "0" . $g[0] : $g[0];

	$GMTstr = $g[5] + 1900 . "-$mon-$day $hr:$min:$sec " .  (($g[2] > 11) ? "PM" : "AM");

	return $GMTstr;

}	# End toGMTfromES()

#--------------------------------------------------------------------
#

=head3 toYDfromDate()

Compute Year and YearDay given a GMT m/d/y.

Author: Mike McCann

Date Created: 7/27/99

=cut

sub toYDfromDate {

	my ($mo, $da, $yr, $hr, $mn) = @_;
	##print "toYDfromDate(); yr = $yr\n";
	
	if ( $mo < 1 || $mo > 12 || $da < 1 || $da > 31 ||
		 $hr < 0 || $hr > 24 || $mn < 0 ||$mn > 60 ) {
		 return "$mo, $da, $yr, $hr, $mn";
	}
	$mo--;
	$hr = 0 if $hr=24;
	
	my $es = timegm(0,$mn,$hr,$da,$mo,$yr);
	
	my @g = gmtime($es);
	
	my $YYYYDDD = sprintf("%4d%03d", $g[5] + 1900, $g[7] + 1);
	
	return $YYYYDDD;

}	# End toYDfromDate()


#--------------------------------------------------------------------
#

=head3 URLEncode()

Make variables safe for a URL.

Author: http://eclectic.kluge.net/codesnippets/perl/URLEncode.html

Date Created: 10/28/98

=cut

sub URLEncode
{
    my($url)=@_;
    my(@characters)=split(/(\%[0-9a-fA-F]{2})/,$url);

    foreach(@characters)
    {
        if ( /\%[0-9a-fA-F]{2}/ ) # Escaped character set ...
        {
            # IF it is in the range of 0x00-0x20 or 0x7f-0xff
            #    or it is one of  "<", ">", """, "#", "%",
            #                     ";", "/", "?", ":", "@", "=" or "&"
            # THEN preserve its encoding
            #"
            unless ( /(20|7f|[0189a-fA-F][0-9a-fA-F])/i
                    || /2[2356fF]|3[a-fA-F]|40/i )
            {
                s/\%([2-7][0-9a-fA-F])/sprintf "%c",hex($1)/e;
            }
        }
        else # Other stuff
        {
            # 0x00-0x20, 0x7f-0xff, <, >, and " ... "
            s/([\000-\040\177-\377\074\076\042])
             /sprintf "%%%02x",unpack("C",$1)/egx;
        }
    }
    return join("",@characters);
}
%>

<%
#--------------------------------------------------------------------
#

=head3 dataLinks()

Do DB query against ExpeditionData table and output href links
a text.  Assumes $Conn object already opened.  This function is
to be called within an area where the DB is opened.

Author: Mike McCann

Date Created: 9/15/99

=cut

sub dataLinks {
	my (%hash) = @_;

	my $ExpdDeviceID_FK = $hash{'expdDeviceID_FK'};
	my $DiveNumber_FK = $hash{'DiveNumber_FK'};
	my $RovName_FK = $hash{'RovName_FK'} || $hash{'RovName'} ;
	my $ExpeditionID_FK = $hash{'ExpeditionID_FK'};
	my $DataType = $hash{'DataType'};
	my $_verbose = $hash{'verbose'};
	my $SearchDiveNumber = $hash{'SearchDiveNumber'};

	my $expdDevID =  $ExpdDeviceID_FK if $ExpdDeviceID_FK ne 0;

	use Benchmark;
	
	Win32::ASP::DebugPrint("\ndataLinks(): DiveNumber_FK = $DiveNumber_FK\nExpeditionID_FK = $ExpeditionID_FK\nDataType = $DataType\n");
	my $start = new Benchmark;


	#
	# Get a Dives RecordSet to loop over
	#
	if ( $DataType eq "samplesdb" || $DataType eq "annotations" ) {
		#
		# Query for dive numbers from this Expedition
		#
		# EXPEDITION LIST OF DIVES FOR SAMPLES AND ANNOTATIONS

		my $sql = "SELECT RovName, DiveNumber, DiveStartDtg, DiveEndDtg, ExpeditionID_FK \n";			
		$sql .= " FROM dbo.Dive \n";
		if ( $ExpeditionID_FK ) {
			if ( $ExpeditionID_FK ne 0 ) {
				$sql .= " WHERE ExpeditionID_FK = " . $ExpeditionID_FK . " \n";
				$sql .= " ORDER BY DiveNumber";
			}
			else {
				$sql .= " WHERE ExpeditionID_FK = 0\n";
				$sql .= " ORDER BY DiveNumber";
			}
		}
		# SAMPLES AND ANNOTATIONS LIST BASED ON DIVENUMBER
		elsif ( $DiveNumber_FK ) {
			$sql .= "WHERE DiveNumber = " . $DiveNumber_FK;
			$sql .= " AND RovName = '$RovName_FK'";
		}
		else {
			$sql = "";
			return;
		}

		Win32::ASP::DebugPrint("\ndataLinks():\nExecuting SQL: $sql");
		my $RS = $Conn->Execute($sql);
		if(!$RS) {
		 	 $Response->Write("<BR>Empty Return Set: RS. for <br>$sql ");
		}

		####################################
        ######  DataType = SAMPLES #########
		####################################
		if ( $DataType eq "samplesdb" ) {
			#
			# Loop through all dive numbers to build string to return
			#
			$str = "";
			while ( !$RS->EOF ) { 
				my $rovName = $RS->Fields('RovName')->Value;
				my $diveNumber = $RS->Fields('DiveNumber')->Value;
				my $diveStartDtg = $RS->Fields('DiveStartDtg')->Value;
				my $diveEndDtg = $RS->Fields('DiveEndDtg')->Value;

				if ( $rovName eq 'vnta' ) {
					$varsRovName = 'Ventana';
					$varsShipName = 'Point Lobos';
					$shipName = 'ptlo';
				}
				elsif ( $rovName eq 'tibr' ) {
					$varsRovName = 'Tiburon';
					$varsShipName = 'Western Flyer';
					$shipName = 'wfly';
				}
				elsif ( $rovName eq 'docr' && $expdDevID eq 991 ) {
					$varsRovName = 'Doc Ricketts';
					$varsShipName = 'Western Flyer';
					$shipName = 'wfly';
				}
				elsif ( $rovName eq 'mini' && $expdDevID eq 991 ) {
					$varsRovName = 'MiniROV';
					$varsShipName = 'Western Flyer';
					$shipName = 'wfly';
				}
				elsif ( $rovName eq 'mini' && $expdDevID eq 1799 ) {
					$varsRovName = 'MiniROV';
					$varsShipName = 'Rachel Carson';
					$shipName = 'rcsn';
				}
				elsif ( $rovName eq 'mini' && $expdDevID eq 1866 ) {
					$varsRovName = 'MiniROV';
					$varsShipName = 'Bold Horizon';
					$shipName = 'bhrzn';
				}
				elsif ( $rovName eq 'mini' && $ExpeditionID_FK eq 0 ) {
					$varsRovName = 'MiniROV';
					$varsShipName = 'Non-MBARI Ship';
					$shipName = 'othr';
				}
				elsif ( $rovName eq 'mini' && !$ExpeditionID_FK ) {
					$varsRovName = 'MiniROV';
					$varsShipName = 'Non-MBARI Ship';
					$shipName = 'othr';
				}
				
#print "<br>ROV_samplesDB_SAMPLES: RovName: $rovName, DiveNumber: $diveNumber, EXPDDeviceID: $expdDevID<br>";		
#print "VARS INFO_SAMPLES: varsRovName: $varsRovName, varsShipName: $varsShipName, shipName: $shipName<br>";

				#
				# Query MBARI_Samples to get count and return link directly
				#
				Win32::ASP::DebugPrint("\ndataLinks(): O.K. DataType should be samplesdb\n");
				my $sqlSamples = "SELECT COUNT(*) \n";
				$sqlSamples .= " FROM dbo.vwCollSamplePerson \n";
				$sqlSamples .= " WHERE (ShipName = '$varsShipName') AND (CollectionEventDTG BETWEEN '$diveStartDtg' AND '$diveEndDtg') \n";
				Win32::ASP::DebugPrint("\ndataLinks(): $rovName $diveNumber: samples query sqlSamples = \n$sqlSamples\n");
				$RSSamples = $ConnSamples->Execute($sqlSamples);
				if(!$RSSamples) {
				 	die "<BR>Dive: $rovName $diveNumber <BR>Empty Return Set: RSSamples for <BR>$sqlSamples ";
				}
				my $count = $RSSamples->Fields(0)->Value;
				
				if ( $count > 0 ) {
#print "HAVE $count SAMPLES!!!!!<br>";
					$url = "https://mww.mbari.org/SamplesDB/queries/returnTable_telepresence.asp?date1=${diveStartDtg}&date2=${diveEndDtg}&ShipName=$varsShipName";
					$LinkText = $count . " samples for dive ${rovName}${diveNumber}";
					if ( $SearchDiveNumber ) {
						$LinkText =~ s#($SearchDiveNumber)#<b>$1</b>#g;
					}
					$str .= " <a href=\"javascript:OpenData('" . $url . "','$DataWindow')\">$LinkText</a> ";			
					$str .= "<b>&middot;</b>";
				}
				$RS->MoveNext;
			}
			$RS->Close;
				
		} # End if ( $DataType eq "samplesdb_" )
		


		####################################
        ###### DataType = ANNOTATIONS ######
		####################################
		elsif ( $DataType eq "annotations" ) {
			#
			# Loop through all dive numbers to build string to return
			#
			$str = "";
			while ( !$RS->EOF ) { 

				my $rovName = $RS->Fields('RovName')->Value;
				my $diveNumber = $RS->Fields('DiveNumber')->Value;
				my $diveStartDtg = $RS->Fields('DiveStartDtg')->Value;
				my $diveEndDtg = $RS->Fields('DiveEndDtg')->Value;
				my $diveName = $rovName . $diveNumber ;

				if ( $rovName eq 'vnta' ) {
					$varsRovName = 'Ventana';
					$varsShipName = 'Point Lobos';
					$shipName = 'ptlo';
				}
				elsif ( $rovName eq 'tibr' ) {
					$varsRovName = 'Tiburon';
					$varsShipName = 'Western Flyer';
					$shipName = 'wfly';
				}
				elsif ( $rovName eq 'docr' ) {
					$varsRovName = 'Doc Ricketts';
					$varsShipName = 'Western Flyer';
					$shipName = 'wfly';
				}
				elsif ( $rovName eq 'mini' && $expdDevID eq 991 ) {
					$varsRovName = 'MiniROV';
					$varsShipName = 'Western Flyer';
					$shipName = 'wfly';
				}
				elsif ( $rovName eq 'mini' && $expdDevID eq 1799 ) {
					$varsRovName = 'MiniROV';
					$varsShipName = 'Rachel Carson';
					$shipName = 'rcsn';
				}
				elsif ( $rovName eq 'mini' && $expdDevID eq 1866 ) {
					$varsRovName = 'MiniROV';
					$varsShipName = 'Bold Horizon';
					$shipName = 'bhrzn';
				}
				elsif ( $rovName eq 'mini' && $ExpeditionID_FK eq 0 ) {
					$varsRovName = 'MiniROV';
					$varsShipName = 'Non-MBARI Ship';
					$shipName = 'othr';
				}
#print "<br>ROV_samplesDB_ANNOTATIONS: RovName: $rovName, DiveNumber: $diveNumber, EXPDDeviceID: $expdDevID<br>";		
#print "VARS INFO_ANNOTATIONS: varsRovName: $varsRovName, varsShipName: $varsShipName, shipName: $shipName<br>";	

				#
				# Query for annotations from VARS database which has already been opened
				# TapeTimeCode, ConceptName are the distinct values [COUNT DISTINCT works for only one column :-(]
				#
				##my $sqlVARS = "SELECT DISTINCT TapeTimeCode, ConceptName \n";
				##$sqlVARS .= " FROM  dbo.Annotations \n";
				##$sqlVARS .= " WHERE (ShipName = '$varsShipName') AND (RecordedDate BETWEEN '$diveStartDtg' AND '$diveEndDtg') OR \n";
		  		##$sqlVARS .= "       (RecordedDate BETWEEN '$diveStartDtg' AND '$diveEndDtg') AND (RovName = '$varsRovName') \n";
		  		
		  		# What Brian gave me
		  		##my $sqlVARS = "SELECT ";
				##$sqlVARS .= "	     count(obs.id) ";
				##$sqlVARS .= "	FROM ";
				##$sqlVARS .= "	     VideoArchiveSet AS vas LEFT OUTER JOIN ";
				##$sqlVARS .= "	     VideoArchive AS va ON va.VideoArchiveSetID_FK = vas.id LEFT OUTER ";  
				##$sqlVARS .= "	JOIN ";
				##$sqlVARS .= "	     VideoFrame AS vf ON vf.VideoArchiveID_FK = va.id LEFT OUTER JOIN ";
				##$sqlVARS .= "	     Observation AS obs ON obs.VideoFrameID_FK = vf.id ";
				##$sqlVARS .= "	WHERE ";
				##$sqlVARS .= "	     vas.PlatformName = '$varsRovName' AND ";
				##$sqlVARS .= "	     vf.RecordedDtg BETWEEN '$diveStartDtg' AND '$diveEndDtg' "; 
				
				# Rich's old school SQL - passed through enterprise manager
				##my $sqlVARS = "SELECT     COUNT(obs.id) AS Expr1 ";
				##$sqlVARS .= "FROM         dbo.VideoArchiveSet vas INNER JOIN ";
				##$sqlVARS .= "                      dbo.VideoArchive va ON vas.id = va.VideoArchiveSetID_FK INNER JOIN ";
				##$sqlVARS .= "                      dbo.VideoFrame vf ON va.id = vf.VideoArchiveID_FK INNER JOIN ";
				##$sqlVARS .= "                      dbo.Observation obs ON vf.id = obs.VideoFrameID_FK ";
				##$sqlVARS .= "WHERE     (vas.PlatformName = '$varsRovName') AND (vf.RecordedDtg BETWEEN '$diveStartDtg' AND '$diveEndDtg') ";

				# Simple count SQL from Annotations_legacy View in M3_ANNOTATIONS
				my $sqlM3 = "select count(DateTime) ";
				$sqlM3 .= " from (SELECT ";
				$sqlM3 .= "	DISTINCT CONVERT(CHAR(8), RecordedDate,1) + ' ' + SUBSTRING(CONVERT(CHAR(19) ";
				$sqlM3 .= "	, RecordedDate, 120), 12, 8) AS DateTime, VideoArchiveName AS Tape, ";
				$sqlM3 .= "	TapeTimeCode, ConceptName, Image AS StillImageURL, DEPTH, Latitude, ";
				$sqlM3 .= " Longitude, CameraDirection AS Direction, isNull(Observer, '') AS Observer, ";
				$sqlM3 .= "	AnnotationMode, ObservationID_FK AS ObservationID ";
				$sqlM3 .= " FROM ";
				$sqlM3 .= "	Annotations_legacy NOLOCK ";
				$sqlM3 .= " WHERE (RovName = '$varsRovName') AND (RecordedDate BETWEEN '$diveStartDtg' AND '$diveEndDtg') ) as subQuery";
				
		  		Win32::ASP::DebugPrint("\ndataLinks(): $rovName $diveNumber: annotations query sqlM3 = \n$sqlM3\n");
		  		
		  		my $stQ = new Benchmark;
		  		
				$RSM3 = $ConnM3->Execute($sqlM3);
				if(!$RSM3) {
				 	##die "<BR>Dive: $rovName $diveNumber <BR>Empty Return Set: RSM3 for <BR>$sqlM3 ";
				 	next;
				}
				my $count = $RSM3->Fields(0)->Value;
				$RSM3->Close;
				
				my $enQ = new Benchmark;
				my $diff = timediff($stQ, $enQ);
				Win32::ASP::DebugPrint("\ndataLinks(): Time taken for Query = ", timestr($diff, 'all'), "\n");

				if ( $count > 0 ) {	
					Win32::ASP::DebugPrint("\ndataLinks(): count = $count\n");
					$url = "https://mww.mbari.org/expd/queries/annotations_telepresence.asp?qSCollDTG=${diveStartDtg}&qECollDTG=${diveEndDtg}&qShipName=${shipName}&qDiveNameNum=${rovName}${diveNumber}";
					$LinkText = $count . " annotations for dive ${rovName}${diveNumber}";
					if ( $SearchDiveNumber ) {
						$LinkText =~ s#($SearchDiveNumber)#<b>$1</b>#g;
					}
					$str .= " <a href=\"javascript:OpenData('" . $url . "','$DataWindow')\">$LinkText</a> ";			
					$str .= "<b>&middot;</b>";
				}
				$RS->MoveNext;
			}
			$RS->Close;	
		} # End elsif ( $DataType eq "annotations" )	
	} # End $DataType eq "samplesdb" || $DataType eq "annotations" ) 	

	else { 
#print "<br><b>DataType = $DataType</b><br>";
		###########################################################################
		# Check for and provide links for rovctd data associated with this dive
		###########################################################################
		my $sql = "";
		$sql = "SELECT DISTINCT DataType, URL, LinkText, PlaceInList, DiveNumber_FK FROM ExpeditionData\n";
		$sql .= "WHERE ";
		if ( $ExpeditionID_FK ) {
			$sql .= "ExpeditionID_FK = " . $ExpeditionID_FK;
		}
		elsif ( $DiveNumber_FK ) {
			$sql .= "DiveNumber_FK = " . $DiveNumber_FK;
			$sql .= " AND RovName_FK = '$RovName_FK'";
		}
		else {
			$sql = "";
			return;
		}
		$sql .= " AND DataType = '$DataType' ";
		$sql .= "ORDER BY PlaceInList, DiveNumber_FK";


#print "$sql<br>";

		Win32::ASP::DebugPrint("\ndataLinks():\nExecuting SQL: $sql");
		##Win32::ASP::DebugPrint("\nSearchDiveNumber = $SearchDiveNumber\n");
		my $RS = $Conn->Execute($sql);

		if(!$RS) {
			$Response->Write("<br>Empty Return Set: RS. for <br>$sql <br>");
		}
		#
		# Construct link string
		#
		$str = "";
		my $expd3Dlink = "https://mww.mbari.org/3Dreplay/3D_telepresence.asp?";
		my $n3D = 0;
		if ( !$RS->EOF ) { 
			while ( !$RS->EOF ) { 
				my $DataWindow = $DataType;
				my $LinkText = $RS->Fields('LinkText')->Value;
				my $Link3D = $RS->Fields('LinkText')->Value;
				if ( $SearchDiveNumber ) {
					##Win32::ASP::DebugPrint("\nBefore LinkText = $LinkText\n");	
					$LinkText =~ s#($SearchDiveNumber)#<b>$1</b>#g;
					##Win32::ASP::DebugPrint("\nAfter  LinkText = $LinkText\n");
				}
				$DataWindow = 'GifPlot' if $LinkText =~ /gif/i;
				if ( $DataType eq '3Dreplay' ) {
					$str .= "<input type=\"checkbox\" name=\"dives\" value=\"";
					$str .= $LinkText . "\" checked>";
					$str .=" <a href=\"javascript:OpenData('" . $RS->Fields('URL')->Value . "','$DataWindow')\">";
					$str .= $LinkText . "</a> ";
					$str .= "<b>&middot;</b>";
					$expd3Dlink .= "dives=" . $Link3D . "&";
					$n3D++;
				}
				else {
					## Pretty much any other DataType
					$str .= " <a href=\"javascript:OpenData('" . $RS->Fields('URL')->Value . "','$DataWindow')\">";

					if ( $LinkText ) {
						$str .=$LinkText  . "</a> ";
						$str .= "<b>&middot;</b>";
					}
					else {
						$str .=$LinkText  . "</a> ";
					}
				#print "<br>STRING1: $str<br>";
				}
				$RS->MoveNext;
			} # End while ( !$RS->EOF )
		}
		$RS->Close;			
		} # End if ( $DataType eq "samplesdb" )
	
	if ( $DataType eq '3Dreplay' && $n3D > 1 ) {
		$str .= " <a href=\"$expd3Dlink\" target=\"3Dreplay\">View these $n3D dives</a>";
	}

	#
	# Add link to ROVCTD data table - now done with popEDwithDBqueries_perseus.pl on draco.shore.mbari.org under Expedition folder.
	#
	#if ( $RovName = 'mini' && $DataType eq 'rovctd' ) {
    #	$str .= " <a href=javascript:OpenData('/expd/queries/rovctd-data_telepresence.asp?date_type=dive&dive=";
    #	$str .= "$DiveNumber_FK&pltfrm=$RovName_FK&return_type=comma','$DataType')>Data for dive $RovName_FK$DiveNumber_FK (CSV)</a> ";
	#	$str .= "<b>&middot;</b>";
	#}
	
	
	$str =~ s/<b>&middot;<\/b>$//;		# Remove last bullet separator
	
	my $end = new Benchmark;
	my $diff = timediff($end, $start);
	Win32::ASP::DebugPrint("\ndataLinks(): Time taken for $DataType = ", timestr($diff, 'all'), " seconds\n");
	
	return $str;
	
}	# End dataLinks()


#--------------------------------------------------------------------
#

=head3 tapeLinks()

Do DB query against VARS to get list of tapes for a dive.
Assumes $ConnVARS object already opened.  This function is
to be called within an area where the DB is opened.

Author: Mike McCann

Date Created: 7/28/2011

=cut

sub tapeLinks {
	my (%hash) = @_;
	
	my $DiveNumber = $hash{'DiveNumber'};
	my $RovName = $hash{'RovName'};
	
	my $varsRovName;
	$varsRovName = 'Doc Ricketts' if $RovName eq 'DocRicketts' || $RovName eq 'docr';
	$varsRovName = 'Tiburon' if $RovName eq 'tibr';
	$varsRovName = 'Ventana' if $RovName eq 'vnta';
	$varsRovName = 'MiniROV' if $RovName eq 'mini';
	
	my $sql =<<EOS;
select
	[VideoArchive].[videoArchiveName]      as [TapeName],
	[VideoArchiveSet].[PlatformName],
	[VideoArchiveSet].[ShipName],
	[VideoArchiveSet].[FormatCode],
	[VideoArchiveSet].[LAST_UPDATED_TIME],
	[CameraPlatformDeployment].[SeqNumber] as [DiveNumber],
	[CameraPlatformDeployment].[ChiefScientist] 
from
	[dbo].[CameraPlatformDeployment] [CameraPlatformDeployment] 
		inner join [dbo].[VideoArchiveSet] [VideoArchiveSet] 
		on [CameraPlatformDeployment].[VideoArchiveSetID_FK] = [VideoArchiveSet]
		.[id] 
			inner join [dbo].[VideoArchive] [VideoArchive] 
			on [VideoArchive].[VideoArchiveSetID_FK] = [VideoArchiveSet].[id] 
where
	([VideoArchiveSet].[PlatformName] = '$varsRovName') and
	([CameraPlatformDeployment].[SeqNumber] = $DiveNumber) 
order by
	[CameraPlatformDeployment].[SeqNumber]
EOS

	Win32::ASP::DebugPrint("\ntapeLinks():\nExecuting SQL: $sql");
	
	my $RS = $ConnVARS->Execute($sql);
	if(!$RS) {
	 	 $Response->Write("<BR>Empty Return Set: RS. for <br>$sql ");
	}

	#
	# Construct link string
	#
	$str = "";
	my $tapeLink = "tape_telepresence.asp?";
	if ( !$RS->EOF ) { 
		while ( !$RS->EOF ) { 
			my $DataWindow = 'TapeEdit';
			my $LinkText = $RS->Fields('TapeName')->Value;

			if ( $SearchDiveNumber ) {
				##Win32::ASP::DebugPrint("\nBefore LinkText = $LinkText\n");	
				$LinkText =~ s#($SearchDiveNumber)#<b>$1</b>#g;
				##Win32::ASP::DebugPrint("\nAfter  LinkText = $LinkText\n");
			}
			
			$str .= " <a href=\"" . $tapeLink . "tapeName=" . $LinkText . "\" target=\"tapeedit\">";
			$str .= $LinkText . "</a> ";
			$str .= "<b>&middot;</b>";
			
			$RS->MoveNext;
		} # End while ( !$RS->EOF )
	}
	$RS->Close;
		
	$str =~ s/<b>&middot;<\/b>$//;
	return $str;
	
}	# End tapeLinks()


#--------------------------------------------------------------------
#

=head3 tapeSummaryLinks()

Do DB query against VARS to get summary list of tapes for an expedition.
Assumes $ConnVARS object already opened.  This function is
to be called within an area where the DB is opened.

Author: Mike McCann

Date Created: 8/1/2011

=cut

sub tapeSummaryLinks {
	my (%hash) = @_;
	
	my $ExpeditionID_FK = $hash{'ExpeditionID_FK'};
	my $DiveNumber_FK = $hash{'DiveNumber_FK'} || $hash{'DiveNumber'};
	my $RovName_FK = $hash{'RovName_FK'} || $hash{'RovName'} ;
	my $_verbose = $hash{'verbose'};
	my $SearchDiveNumber = $hash{'SearchDiveNumber'};
	
	my $varsRovName;
	$varsRovName = 'Doc Ricketts' if $RovName eq 'DocRicketts' || $RovName eq 'docr';
	$varsRovName = 'Tiburon' if $RovName eq 'tibr';
	$varsRovName = 'Ventana' if $RovName eq 'vnta';
	$varsRovName = 'MiniROV' if $RovName eq 'mini';
	
	#
	# Query for dive numbers from this Expedition
	#
	my $sql = "SELECT RovName, DiveNumber, DiveStartDtg, DiveEndDtg, ExpeditionID_FK \n";
	$sql .= " FROM dbo.Dive \n";
	if ( $ExpeditionID_FK ) {
		$sql .= " WHERE ExpeditionID_FK = " . $ExpeditionID_FK . " \n";
		$sql .= " ORDER BY DiveNumber";
	}
	elsif ( $DiveNumber_FK ) {
		$sql .= "WHERE DiveNumber = " . $DiveNumber_FK;
		$sql .= " AND RovName = '$RovName_FK'";
	}
	else {
		$sql = "";
		return;
	}
	
	Win32::ASP::DebugPrint("\ntapeSummaryLinks():\nExecuting SQL: $sql");
	my $RSdive = $Conn->Execute($sql);
	if(!$RSdive) {
	 	 $Response->Write("<BR>Empty Return Set: RSdive. for <br>$sql ");
	}
	
		
	#
	# Loop through all dive numbers to build string to return
	#
	$str = "";
	my $tapeSummaryLink = "tapeSummary_telepresence.asp?";

	while ( !$RSdive->EOF ) { 
		my $rovName = $RSdive->Fields('RovName')->Value;
		my $diveNumber = $RSdive->Fields('DiveNumber')->Value;
		my $diveStartDtg = $RSdive->Fields('DiveStartDtg')->Value;
		my $diveEndDtg = $RSdive->Fields('DiveEndDtg')->Value;
		if ( $rovName eq 'vnta' ) {
			$varsRovName = 'Ventana';
			$varsShipName = 'Point Lobos';
			$shipName = 'ptlo';
		}
		elsif ( $rovName eq 'tibr' ) {
			$varsRovName = 'Tiburon';
			$varsShipName = 'Western Flyer';
			$shipName = 'wfly';
		}
		elsif ( $rovName eq 'docr' ) {
			$varsRovName = 'Doc Ricketts';
			$varsShipName = 'Western Flyer';
			$shipName = 'wfly';
		}
		elsif ( $rovName eq 'mini' && $expdDevID eq 991 ) {
			$varsRovName = 'MiniROV';
			$varsShipName = 'Western Flyer';
			$shipName = 'wfly';
		}
		elsif ( $rovName eq 'mini' && $expdDevID eq 1799 ) {
			$varsRovName = 'MiniROV';
			$varsShipName = 'Rachel Carson';
			$shipName = 'rcsn';
		}
		elsif ( $rovName eq 'mini' && $expdDevID eq 1866 ) {
			$varsRovName = 'MiniROV';
			$varsShipName = 'Bold Horizon';
			$shipName = 'bhrzn';
		}
		elsif ( $rovName eq 'mini' && $ExpeditionID_FK eq 0 ) {
			$varsRovName = 'MiniROV';
			$varsShipName = 'Non-MBARI Ship';
			$shipName = 'othr';
		}				
	    #elsif ( $rovName eq 'mini' ) {
		#	$varsRovName = 'MiniROV';
		#	$varsShipName = 'Western Flyer' || 'Rachel Carson';
		#	$shipName = 'wfly' || 'rscn';
		#}
				
		#
		# Get number of tapes from VARS - per Lonny's communication on 9/21/2011 this is not
		# as accurte as the numbers that are in the TapeSummary table, which the Video Lab
		# can easily edit.			
		#
	
#		$sql =<<EOS;
#select
#	[VideoArchive].[videoArchiveName]      as [TapeName],
#	[VideoArchiveSet].[PlatformName],
#	[VideoArchiveSet].[ShipName],
#	[VideoArchiveSet].[FormatCode],
#	[VideoArchiveSet].[LAST_UPDATED_TIME],
#	[CameraPlatformDeployment].[SeqNumber] as [DiveNumber],
#	[CameraPlatformDeployment].[ChiefScientist] 
#from
#	[dbo].[CameraPlatformDeployment] [CameraPlatformDeployment] 
#		inner join [dbo].[VideoArchiveSet] [VideoArchiveSet] 
#		on [CameraPlatformDeployment].[VideoArchiveSetID_FK] = [VideoArchiveSet]
#		.[id] 
#			inner join [dbo].[VideoArchive] [VideoArchive] 
#			on [VideoArchive].[VideoArchiveSetID_FK] = [VideoArchiveSet].[id] 
#where
#	([VideoArchiveSet].[PlatformName] = '$varsRovName') and
#	([CameraPlatformDeployment].[SeqNumber] = $diveNumber) 
#order by
#	[CameraPlatformDeployment].[SeqNumber]
#EOS
#	
#		Win32::ASP::DebugPrint("\ntapeSummaryLinks():\nExecuting SQL: $sql");
#		
#		my $RStape = $ConnVARS->Execute($sql);
#		if(!$RStape) {
#		 	 $Response->Write("<BR>Empty Return Set: RStape. for <br>$sql ");
#		}
#	
#		#
#		# Construct link string
#		#
#		my $sdCount = 0;
#		my $hdCount = 0;
#		if ( !$RStape->EOF ) { 
#			while ( !$RStape->EOF ) { 
#				if ($RStape->Fields('TapeName')->Value =~ /HD$/) {
#					$hdCount++;
#				}
#				else {
#					$sdCount++;
#				}
#	
#				
#				$RStape->MoveNext;
#			} # End while ( !$RStape->EOF )
#		}
#		$RStape->Close;

		#
		# Get number of tapes from the TapeSummary table
		#
		$sql =<<EOS;
select
	[TapeSummary].[NumSDTapes],
	[TapeSummary].[NumHDTapes] 
from
	[dbo].[TapeSummary] [TapeSummary] 
where
	([TapeSummary].[DiveNumber] = $diveNumber) and
	([TapeSummary].[RovName] ='$rovName')
EOS

		$sdCount = 0;
		$hdCount = 0;
		
		Win32::ASP::DebugPrint("\ntapeSummaryLinks():\nExecuting SQL: $sql");
		
		my $RStape = $Conn->Execute($sql);
		if(!$RStape) {
		 	 $Response->Write("<BR>Empty Return Set: RStape. for <br>$sql ");
		}
		$sdCount = $RStape->Fields('NumSDTapes')->Value if $RStape->Fields('NumSDTapes')->Value;
		$hdCount = $RStape->Fields('NumHDTapes')->Value if $RStape->Fields('NumHDTapes')->Value;

		$LinkText = '';
		if ($sdCount && !$hdCount) {
		 	$LinkText .= "$sdCount SD tape";
		 	$LinkText .= "s" if $sdCount > 1;
		 	$LinkText .= " for ${rovName}${diveNumber}";
		}
		if ( $sdCount && $hdCount ) {
			$LinkText .= "$sdCount SD and $hdCount HD tapes";
			$LinkText .= " for ${rovName}${diveNumber}";
		}
		if (!$sdCount && $hdCount) {
		 	$LinkText .= "$hdCount HD tape";
		 	$LinkText .= "s" if $hdCount > 1;
		 	$LinkText .= " for ${rovName}${diveNumber}";
		}
		elsif ( !$sdCount && !$hdCount ){
			$LinkText = "Enter tape information for ${rovName}${diveNumber}";
		}
		
		if ( $SearchDiveNumber ) {
			##Win32::ASP::DebugPrint("\nBefore LinkText = $LinkText\n");	
			$LinkText =~ s#($SearchDiveNumber)#<b>$1</b>#g;
			##Win32::ASP::DebugPrint("\nAfter  LinkText = $LinkText\n");
		}
		
		$str .= " <a href=\"" . $tapeSummaryLink . "rovName=" . $rovName . "&diveNumber=" . $diveNumber . "\" target=\"tapeedit\">";
		$str .= $LinkText . "</a> ";
		$str .= "<b>&middot;</b>";

		$RSdive->MoveNext;
		
	} # End while ( !$RSdive->EOF ) { 
		
	
	$str =~ s/<b>&middot;<\/b>$//;
	return $str;
	
}	# End tapeSummaryLinks()

%>

<%

#--------------------------------------------------------------------
# Called at the beginning of each asp.  Need to do this because 
# ActiveState Perl 5.6.1 and higher (to 5.8) are broken and no longer read 
# the global.asa file like 5.22 did.     Mike McCann 22 June 2005

=head3 init_AppVars()
=cut

sub init_AppVars {
	
	my $_debug = 0;
	
	#
	# Define application variables
	#
	$Application->{'providerStr'} = "sqloledb";
	#$Application->{'DSN'} = "Server=fog;Database=expd_ks;UID=***;******;";
	$Application->{'DSN'} = "Server=perseus;Database=expd;UID=***;******;";
	
	$Application->{'AppDir'} = "/expd/log";
	$Application->{'BaseURL'} = "https://mww.mbari.org";
	$Application->{'BaseURL'} .= $Application->{'AppDir'};
	
	### WARNINGS... to test if this email step works correctly, enter an precruise
	### while logged in to windows as any mbari users who is NOT an administrator
	### (or registeredUser ?) as defined at ~ line 690 thru 720 below.
	### also, if an admin accidently manages to create a test expd then don't forget
	### to manually delete it from the database via SQL Enterprise manager...
	### r.schramm 8/15/07
	
	##$Application->{'logistics_email'} = 'salamy@mbari.org';
	$Application->{'logistics_email'} = 'mschultz@mbari.org';
	##$Application->{'logistics_email'} = 'sara@mbari.org';
	##$Application->{'logistics_email'} = 'fana@mbari.org';
	##$Application->{'logistics_email'} = 'cruiseadmin@mbari.org';
	
	#
	# Dive table fields
	#
	@DiveFields = (	'RovName', 'DiveNumber', 'ExpeditionID_FK', 'DiveStartDtg',
				    'DiveEndDtg', 'DiveChiefScientist', 'BriefAccomplishments' );
	
	
	#
	# List of fields in their order of appearance for the precruise web pages
	# If database design is changed then this table should be changed too.  We
	# do this because we can control the order in which they appear with this list.
	#
	@ExpeditionFields = ( 	'ExpeditionID',
	                        'ShipName',
	                        'ShipSeqNum', 
	                        'Purpose', 
	                        'StatCode', 
	                        'ExpdChiefScientist',
	                        'ExpdPrincipalInvestigator',
	                        'ScheduledStartDtg', 
	                        'ScheduledEndDtg',
	 		      			'EquipmentDesc', 
	 		      			'Participants', 
							'RegionDesc', 
							'PlannedTrackDesc',
							'StartDtg', 
							'EndDtg', 
							'Accomplishments', 
							'ScientistComments',
							'SciObjectivesMet', 
							'OperatorComments', 
							'AllEquipmentFunctioned', 
							'OtherComments',
							'UpdatedBy', 
							'isModified' );
		

	#
	# Only the fields to be updated in a precruise revision
	#
	@PreUpdateExpeditionFields = ( 'ExpdChiefScientist', 'ExpdPrincipalInvestigator',
					'ScheduledStartDtg', 'ScheduledEndDtg', 
					'Purpose', 'EquipmentDesc', 'Participants', 
					'PlannedTrackDesc',
					'ExpeditionID' );
	
	#
	# Only the fields to be updated in the post cruise
	#
	@PostUpdateExpeditionFields = ( 'StatCode', 'ExpdChiefScientist', 'ExpdPrincipalInvestigator',
						  'StartDtg', 'EndDtg', 'Accomplishments', 'SciObjectivesMet',
						  'ScientistComments', 'AllEquipmentFunctioned', 'OperatorComments',
						  'OtherComments', 'ExpeditionID' );
		

	$Application->{'DiveFields'} = \@DiveFields;
	$Application->{'ExpeditionFields'} = \@ExpeditionFields;
	$Application->{'PostUpdateExpeditionFields'} = \@PostUpdateExpeditionFields;
	$Application->{'PreUpdateExpeditionFields'} = \@PreUpdateExpeditionFields;
	
	
	#
	# Set DBAdministrator flag based on login names of the logistics coordinator(s)
	# Note: DBAdministator should also be RegisteredUser (enter IP address below as well).
	#
	Win32::ASP::DebugPrint("\ninit_AppVars(): LOGON_USER = " . $Request->ServerVariables('LOGON_USER')->{Item} . "\n") if $_debug;
	if (   $Request->ServerVariables('LOGON_USER')->{Item} =~ /mschultz/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /salamy/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /fana/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /sara/
	   ) {
		$Session->{'DBAdministrator'} = 1;
	}
	else {
		$Session->{'DBAdministrator'} = 0;
	}
	
	
	#
	# Set RegisteredUser flag based on login name
	#
	if ( ! defined($Session->{'RegisteredUser'}) ) {
		Win32::ASP::DebugPrint("\ninit_AppVars(): Session->{'RegisteredUser'} is not defined\n");
	}
	if (   $Request->ServerVariables('LOGON_USER')->{Item} =~ /mschultz/ 
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /sara/		
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /fana/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /paje/ 
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /robs/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /mccann/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /rich/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /lonny/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /svonthun/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /schlin/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /linda/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /mnakagawa/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /salamy/
		|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /mbasset/
	   ) {
		$Session->{'RegisteredUser'} = 1;
	}
	
	



} # End init_AppVars()


#--------------------------------------------------------------------
#

#=head3 num2month()

#Convert number month to name month.

#num2month() should be called whenever you need to convert a number month
#to a nameMonth

#Author: Karen Salamy

#Date Created: 10/25/2019

#=cut

#sub num2month {

#    my $num2month = qw( 1 Jan  2 Feb  3 Mar  4 Apr  5 May  6 Jun  7 Jul  8 Aug  9 Sep  10 Oct  11 Nov  12 Dec );

#);

#} # End getMonth()


#--------------------------------------------------------------------
#

=head3 close_session()

Close the session.  Has the effect of clearing out all the session
variables that are accumulated in preparation for the database load
or email sent to logistics coordinator.  If the user is a registered
user, we do need to remember that session variable so that if she
goes on to enter another postcruise it hasn't forgotten her identity.

close_session() should be called whenever a post or pre cruise edit session
has finished.

Author: Mike McCann

Date Created: 11/22/99

=cut

sub close_session {

	my $reg_user = $Session->{'RegisteredUser'};
	my $log_email = $Session->{'logistics_email'};
	$Session->Abandon;
	$Session->{'RegisteredUser'} = $reg_user;
	$Session->{'logistics_email'} = $log_email;

} # End close_session()
%>


