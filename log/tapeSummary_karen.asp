<%@ LANGUAGE = PerlScript %>

<!--#include file="expd_functions_karen.inc"-->

<%
if ( GetFormValue('resultsAs') ne 'excelTxt' && GetFormValue('resultsAs') ne 'tapeLabelPdf' ) {
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Tape Summary Information</title>
<link rel="stylesheet" type="text/css" href="/expd/styles/expd_karen.css">
<script language="JavaScript" type="text/javascript">
<!-- hide from old browsers

function OpenData(url,type) 
{
   
   if (navigator.appVersion.indexOf("(X11") != -1 || navigator.appVersion.indexOf("(Mac") != -1) {
	if (type == "frameGrabs") {
		var win = window.open(url, type, "menubar=yes,toolbar=yes,alwaysRaised=yes,location=yes,status=1,scrollbars=yes,resizable=yes");
	}
	else if (type == "logrparse") {
		var win = window.open(url, type, "menubar=yes,alwaysRaised=yes,location=yes,status=1,scrollbars=yes,resizable=yes,width=500,height=300");
	}
	else {
		var win = window.open(url, type, "menubar=yes,alwaysRaised=yes,location=yes,status=1,scrollbars=yes,resizable=yes");
	}
   }
   else {
    if (type == "frameGrabs" || type == "digitalImages") {
		var win = window.open(url, type, "menubar=yes,toolbar=yes,alwaysRaised=yes,location=yes,status=1,scrollbars=yes,resizable=yes");
	}
	else if (type == "logrparse") {
		var win = window.open(url, type, "menubar=yes,alwaysRaised=yes,location=yes,status=1,scrollbars=yes,resizable=yes,width=500,height=300");
	}
	else {
		//var win = window.open(url, type, "menubar=yes,alwaysRaised=yes,location=yes,status=1,scrollbars=yes,resizable=yes");
		var win = window.open(url, type);
	}
   }
}

// -->
</script>
</head>
<body bgcolor="#FFFFFF">
<table border="0" cellpadding="5" width="680">
<tr>
<td valign="top" width="20%">
<a href="http://mww.mbari.org">
<img src="/expd/log/images/mbarilogo-120_sh.gif" border="0" width="120" height="70"></a>
</td>
<td valign="top" width="80%">

<%
}
%>

<% 
=head1 NAME

tapeSummary_karen.asp - Application for displaying and editing information in Expd's tape table

=head1 SYNOPSIS

    http://expd.mbari.org/expd/log/tapeSummary_karen.asp

=head1 DESCRIPTION

Tape information editing Active Server Page.
To replace Video Lab's Excel/Access method of adding additional information about the Video tape archive.

Mike McCann MBARI

July 2011

=head1 FUNCTIONS

=cut


# Warning: Don't use CGI, it freezes IIS when a new object is created.
use Win32::ASP;
use OLE;
use Time::Local;
use POSIX;
##use PDF::API2;
use Win32::OLE::Variant;

$this_script = 'tapeSummary_karen.asp';

init_AppVars();

$dsn = $Application->{'DSN'};

#
# Check if Video Lab Staff in order to enable the submit buttons
#
$vlStaffFlag = 1;
if ( $Request->ServerVariables('LOGON_USER')->{Item} =~ /lonny/i
	|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /schlin/i
	|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /svonthun/i
	|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /mccann/i
	|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /linda/i
	|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /jana/i
	|| $Request->ServerVariables('LOGON_USER')->{Item} =~ /dana/i
	) {
	$vlStaffFlag = 1;
}


if ( GetFormValue('rovName') && GetFormValue('diveNumber') && ! GetFormValue('step') ) {
	%><h2>Tape Summary Information for <%=GetFormValue('rovName')%><%=GetFormValue('diveNumber')%></h2>
	<P align=left><FONT color=#408080>Edit information in Expd</FONT></P>
	</td></tr></table>
	 <% 
	edit(GetFormValue('rovName'), GetFormValue('diveNumber'), GetFormValue('freeFormFields'));
}

elsif ( GetFormValue('step') eq 'AddDiveToTapeSummary') {
	%><h2>Submitting changes to Expd</h2>
	</td></tr></table><font face="Arial,Helvetica"></font><% 
	addDiveToTapeSummary(GetFormValue('rovName'), GetFormValue('diveNumber'));
	%>
	<p><a href="<%=$this_script%>?rovName=<%=$rovName%>&diveNumber=<%=$diveNumber%>">Edit Tape information for <%=$rovName%><%=$diveNumber%></a></p>
	<%	
}
elsif ( GetFormValue('step') eq 'UpdateTapeSummary') {
	%><h2>Submitting changes to Expd</h2>
	</td></tr></table><font face="Arial,Helvetica"></font><% 
	updateTapeSummary(GetFormValue('rovName'), GetFormValue('diveNumber'));
	%>
	<p><a href="<%=$this_script%>?rovName=<%=$rovName%>&diveNumber=<%=$diveNumber%>">Edit Tape information for <%=$rovName%><%=$diveNumber%></a></p>
	<%	
}
elsif ( GetFormValue('step') eq 'QueryForm') {
	%><h2>Query TapeSummary information</h2>
	</td></tr></table><font face="Arial,Helvetica"></font><% 
	queryForm();	
}
elsif ( GetFormValue('step') eq 'QueryDB' && GetFormValue('resultsAs') eq 'diveList' ) {
	%><h2>Results from TapeSummary query</h2>
	</td></tr></table><font face="Arial,Helvetica"></font><% 
	queryDB();	
}
elsif ( GetFormValue('step') eq 'QueryDB' && GetFormValue('resultsAs') eq 'excelTxt' ) {
	# No HTML
	$Response->{'ContentType'} = 'text/tab-separated-values';	
	$Response->AddHeader('content-disposition', 'attachment; filename=tapeSummary.txt');
	queryDB();	
}
elsif ( GetFormValue('step') eq 'QueryDB' && GetFormValue('resultsAs') eq 'tapeLabelPdf' ) {
	# No HTML
	$Response->{'ContentType'} = 'application/pdf';	
	$Response->AddHeader('content-disposition', 'attachment; filename=tapeLabels.pdf');
	##$Response->{'ContentType'} = 'text/html';			# uncommnet this, and comment out above 2 lines if printing debug statements
	queryDB();
}


else {
	%><h2>Unknown step: <%= GetFormValue('step')%></h2>
	</td></tr></table><% 
}
%>

<%
if ( GetFormValue('resultsAs') ne 'excelTxt' && GetFormValue('resultsAs') ne 'tapeLabelPdf' ) {
%>
<!--<p class="smalltext" align="center">
&middot;
<br>
<a href=/expd/queries/viewExpd.kml>Query from Google Earth</a>
&middot;
<a href=waypoint_karen.asp>Waypoint management</a>
&middot;
<a href="http://mww2.shore.mbari.org/events/calendar.cgi">Web calendar</a>
&middot;
<a href="/3Dreplay/3Dterrain.asp">Dives over terrain</a>
&middot;
<a href="precruise_karen.asp">Precruise entry</a>
&middot;
<a href="postcruise_karen.asp">Keyword search</a>
&middot;
<a href="postcruise_karen.asp?search=advanced">Advanced search</a>
&middot;
<a href="default_karen.asp">Administrative page</a>
&middot;
<a href="http://www.mbari.org/dmo/cruise_planning/cruise.htm">Cruise planning &amp; scheduling</a>
&middot;
<a href="/itd/video/ShipProcedures/index.htm">Shipboard video procedures</a>
&middot;
<a href="/samplesDB/Queries/">Samples Database</a>
&middot;
<a href="http://tramontane2.shore.mbari.org/expd/log/tapeSummary.asp?step=QueryForm">Video Tapes</a>
</p>

<hr>
Expedition Database maintenance team: 
Mike McCann, Jenny Paduan, Rich Schramm, Teresa Cardoza, Lonny Lundsten.
-->
</body>
</html>
<%
}
%>
<!--#include file="postcruise_ftr_karen.inc"-->
<%
#--------------------------------------------------------------------
#

=head3 queryForm()

Present TapeSummary fields and return Expedition report - or other reports 

Author: Mike McCann

Date Created: 8/7/2011

=cut


sub queryForm {

	#
	# Open the database, this time to insert fields
	#
	open_database($dsn);		# Creates $Conn object as a global variable
	
	my $sql =<<EOS;
SELECT distinct Annotators
  FROM [EXPD].[dbo].[TapeSummary]
  ORDER BY annotators
EOS
	Win32::ASP::DebugPrint("\nedit():\nExecuting Expd SQL: $sql");
	
	$RSannotators = $Conn->Execute($sql);
	if(!$RSannotators) {
	 	 $Response->Write("<BR>Empty Return Set: RSannotators. for <br>$sql ");
	}	
	
	
	$sql =<<EOS;
SELECT distinct DiveChiefScientist
  FROM [EXPD].[dbo].[Dive]
  ORDER BY DiveChiefScientist
EOS
	Win32::ASP::DebugPrint("\nedit():\nExecuting Expd SQL: $sql");
	
	$RSdiveChfSci = $Conn->Execute($sql);
	if(!$RSdiveChfSci) {
	 	 $Response->Write("<BR>Empty Return Set: RSdiveChfSci. for <br>$sql ");
	}	

	$sql =<<EOS;
SELECT distinct Camera
  FROM [EXPD].[dbo].[TapeSummary]
  ORDER BY Camera
EOS
	Win32::ASP::DebugPrint("\nedit():\nExecuting Expd SQL: $sql");
	
	my $RScamera = $Conn->Execute($sql);
	if(!$RScamera) {
	 	 $Response->Write("<BR>Empty Return Set: RScamera. for <br>$sql ");
	}
	
	@month_names = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov",
                  "Dec");

	$lstr = localtime(time+86400*180);	# The year for 180 days from now
	@lstr2 = split('\s+', $lstr);
	my $year_10 = $lstr2[4];
	my $year_180 = $lstr2[4];

	for (my $day = 1; $day <= 31; $day++ ) {
		$qday_option_list .= "<option value=\"$day\">$day</option>";
		$qday_2digit_option_list .= "<option value=\"" . sprintf("%02d", $day) . "\">" . sprintf("%02d", $day) . "</option>";
	}
	$monNum = 1;
	foreach my $mon ( @month_names ) {
		$qmonth_option_list .= "<option value=\"$mon\">$mon</option>";
		$qmonth_num_option_list .= "<option value=\"" . sprintf("%02d", $monNum) . "\">$mon</option>";
		$monNum++;
	}
	for (my $y = $year_180; $y >= 1988; $y--) {
		$qyear_option_list .= "<option value=\"$y\">$y</option>";
	}
	for (my $y = $year_10; $y >= 1988; $y--) {
		$qrev_year_option_list .= "<option value=\"$y\">$y</option>";
	}

	%>
	
<form method="POST" action="<%= $this_script%>" id="form1" name="form1">

  <input type="hidden" name="step" value="QueryDB">
  <table border="0" cellpadding="5" cellspacing="0">
    <tr>  
	  <td colspan="4" align="left" valign="top" bgcolor="#F0F0F0">
        <b>Join fields below with:</b><input type="radio" name="Conjunction" value="AND" checked>AND 
	<input type="radio" name="Conjunction" value="OR">OR
	  </td> 
    </tr>

    <tr>
	  <td valign="top" colspan="2"><strong><font face="Arial,Helvetica" size="-1">Expedition:</font></strong></td>

      <td valign="top"><font face="Arial,Helvetica" size="-1"><b>ShipName</b></font></td>
      <td><select name="qShipName">
		<option value=""></option>		
		<option value="ptlo">Point Lobos</option>
		<option value="rcsn">Rachel Carson</option>
		<option value="wfly">Western Flyer</option>
		<option value="zphr">Zephyr</option>
		<option value="bhrzn">Non-MBARI Ship - Bold Horizon</option>
		</select>
      </td>
      
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>


    <tr>
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">&nbsp;&nbsp;Date Choices</font></strong></td>
      <td><select name="qDates">
		<option value=""></option>
		<option value="Future">Future</option>
		<option value="last7">Last 7 days</option>
		<option value="last30">Last 30 days</option>
		<option value="last60">Last 60 days</option>
		<option value="last180">Last 180 days</option>
		<%= $qrev_year_option_list %>
		</select> 
		<span class="smallbrownText">Based on scheduled Expedition start time</span>
	  </td>
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">YYYYDDD</font></strong></td>
      <td><input type="text" name="qYYYYDDD" size="7" value=""> 
      <span class="smallbrownText">e.g. 1999075</span></td>
    </tr>
    <tr>
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">&nbsp;&nbsp;Start Date</font></strong></td>
      <td>
        <font face="Arial,Helvetica" size="-1">First of </font><select name="qExpdSmon" size="1">
		<option value=""></option>
		<%= $qmonth_option_list %>
		</select>
		<font face="Arial,Helvetica" size="-1"> </font><select name="qExpdSyr">
		<option value=""></option>
		<%= $qyear_option_list %>
		</select> 
		<span class="smallbrownText">Specify End Date too ---&gt;</span>
	  </td>
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">End Date</font></strong></td>
      <td>
        <font face="Arial,Helvetica" size="-1">First of </font><select name="qExpdEmon" size="1">
		<option value=""></option>
		<%= $qmonth_option_list %>
		</select> 
		<font face="Arial,Helvetica" size="-1"> </font><select name="qExpdEyr">
		<option value=""></option>
		<%= $qyear_option_list %>
		</select> 
	  </td>
    </tr>

    <tr>
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">DiveNumber</font></strong></td>
      <td><input type="text" name="qDiveName" size="5">
      <span class="smallbrownText">May precede with letter, e.g. V3296, T1001, D1</span>
      
      </td>
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">Chief Scientist</font></strong></td>
      <td><select name="qDiveChfSci">
		<%while (!$RSdiveChfSci->EOF) {
            		$diveChfSciChoice = $RSdiveChfSci->Fields('DiveChiefScientist')->Value;
					my $diveChfSci_option = "<option value=\"$diveChfSciChoice\"";
					$diveChfSci_option .= ">$diveChfSciChoice</option>\n";
					$Response->write($diveChfSci_option);
					$RSdiveChfSci->MoveNext;
				 } %>
		</select>
	  </td>
    </tr>
	
	<tr>
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">Camera</font></strong></td>	  
	  <td><select name="qCamera">
		<option value=""></option>
		<%while (!$RScamera->EOF) {
            		$cameraChoice = $RScamera->Fields('Camera')->Value;
					my $camera_option = "<option value=\"$cameraChoice\"";
					$camera_option .= ">$cameraChoice</option>\n";
					$Response->write($camera_option);
					$RScamera->MoveNext;
				 } %>
		</select>
	  </td>
	   
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">Dive Date</font></strong></td>
      <td>
      		<select name="qDiveDay" size="1">
      		<option value=""></option>
		<%= $qday_2digit_option_list %>
		</select> 
		<select name="qDiveMon" size="1">
		<option value=""></option>
		<%= $qmonth_num_option_list %>
		</select> 
		<select name="qDiveYr" size="1">
		<option value=""></option>
		<%= $qyear_option_list %>
		</select> 
      
      </td>
    </tr>
    
    <tr>
      <td valign="top"><font face="Arial,Helvetica" size="-1"><b>Notes:</b></font></td>
      <td><input type="text" name="qNotes" size="20">
      </td>
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">Mission:</font></strong></td>
      <td><input type="text" name="qMission" size="20">
      </td> 
    </tr>
    
    <tr>
      <td valign="top"><font face="Arial,Helvetica" size="-1"><b>Annotators:</b></font></td>
	  
	  <td><select name="qAnnotators" size="1">
				<option value=""></option>
            	<%while (!$RSannotators->EOF) {
            		$annotatorChoice = $RSannotators->Fields('Annotators')->Value;
					my $annotator_option = "<option value=\"$annotatorChoice\"";
					$annotator_option .= ">$annotatorChoice</option>\n";
					$Response->write($annotator_option);
					$RSannotators->MoveNext;
				 } %>
            </select></td>
	  
	  
      <td valign="top">&nbsp;</td>
      <td>&nbsp;
      </td> 
    </tr>
	
    <tr>
	  <td valign="top" colspan="1"><strong><font face="Arial,Helvetica" size="-1">DateAnnotated:</font></strong></td>
	  <td><input type="checkbox" name="DateAnnotatedIsNull" value="1"></input>IS NULL</td>
    </tr>
	
	<tr>
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">&nbsp;&nbsp;Start</font></strong></td>
      <td>
		<select name="qSday" size="1">
		<option value=""></option>
		<%= $qday_option_list %>
		</select>
		<select name="qSmon" size="1">
		<option value=""></option>
		<%= $qmonth_option_list %>
		</select>
		<select name="qSyr">
		<option value=""></option>
		<%= $qyear_option_list %>
		</select> 
		<span class="smallbrownText">Specify End Date too ---&gt;</span>
	  </td>
      <td valign="top"><strong><font face="Arial,Helvetica" size="-1">End</font></strong></td>
      <td>
		<select name="qEday" size="1">
		<option value=""></option>
		<%= $qday_option_list %>
		</select>
		<select name="qEmon" size="1">
		<option value=""></option>
		<%= $qmonth_option_list %>
		</select> 
		<select name="qEyr">
		<option value=""></option>
		<%= $qyear_option_list %>
		</select> 
	  </td>
    </tr>
    
    
    </table>
    <br>
    <font face="Arial,Helvetica" size="-1">
  Return results as: <input type=radio name="resultsAs" value="diveList" checked>Dive list
  <input type=radio name="resultsAs" value="excelTxt">Tab-delimited .txt file
  <input type=radio name="resultsAs" value="tapeLabelPdf">Tape Label PDF file
  </font>
    <br>
    <br>
  <input type="submit" name="Continue" value="Search for tape information"> 
  <input type="reset" name="Reset form" value="Reset form fields">
  
</form>   
    <%
	
	$Conn->Close;
		
}	# End queryForm()


#--------------------------------------------------------------------
#

=head3 queryDB()

Present TapeSummary fields and return Expedition report - or other reports 

Author: Mike McCann

Date Created: 8/8/2011

=cut


sub queryDB {

	my $_debug = 0;
	print "Inside queryDB()" if $_debug;

	#
	# Open the database, this time to insert fields
	#
	open_database($dsn);		# Creates $Conn object as a global variable
	
	$conjunction = GetFormValue('Conjunction');
	$qMission = GetFormValue('qMission');
	$qNotes = GetFormValue('qNotes');
	$qAnnotators = GetFormValue('qAnnotators');
	$qDiveName = GetFormValue('qDiveName');
	$qDiveChfSci = GetFormValue('qDiveChfSci');
	$qCamera = GetFormValue('qCamera');
	$qShipName = GetFormValue('qShipName');
	$resultsAs = GetFormValue('resultsAs');
	$dateAnnotatedIsNull = GetFormValue('DateAnnotatedIsNull');
							
	
	$qSDate = "'" . GetFormValue('qSday') . ' ' . GetFormValue('qSmon') . ' ' . GetFormValue('qSyr') . "'" if (GetFormValue('qSday') && GetFormValue('qSmon') && GetFormValue('qSyr'));
	$qEDate = "'" . GetFormValue('qEday') . ' ' . GetFormValue('qEmon') . ' ' . GetFormValue('qEyr') . "'" if (GetFormValue('qEday') && GetFormValue('qEmon') && GetFormValue('qEyr'));

	# Query for dives and Expeditions from data in TapeSummary table
	$where_clause = "WHERE ";
	$where_clause .= "\n  Notes LIKE '%" . $qNotes . "%'" if $qNotes;
	$where_clause .= "\n  $conjunction " if ($qNotes && $qMission);
	$where_clause .= "\n  Mission LIKE '%" . $qMission . "%'" if $qMission;
	$where_clause .= "\n  $conjunction " if (($qNotes || $qMission) && $qAnnotators);
	$where_clause .= "\n  Annotators = '" . $qAnnotators . "'" if $qAnnotators;
	$where_clause .= "\n  $conjunction " if (($qNotes || $qMission || $qAnnotators) && $qDiveName);
	
	if ( $qDiveName ) {	
		$qDiveName =~ /(^[tvdTVD]*)(\d+)$/;
		$qDiveNumber = $2;	# Got to pick off 2 first for some reason (?)
		if ( $1 =~ /T/i || GetFormValue('qRovName') eq 'tibr' ) {
			$qRovName = 'tibr';
		}
		elsif ( $1=~ /V/i || GetFormValue('qRovName') eq 'vnta' ) {
			$qRovName = 'vnta';
		}
		elsif ( $1=~ /D/i || GetFormValue('qRovName') eq 'docr' ) {
			$qRovName = 'docr';
		}
	
		$where_clause .= "\n  [Dive].[RovName] = '" . $qRovName . "' AND " if $qRovName;
		$where_clause .= " [Dive].[DiveNumber] = $qDiveNumber";
	}
	$where_clause .= "\n  $conjunction " if (($qNotes || $qMission || $qAnnotators || $qDiveName) && $qDiveChfSci);
	$where_clause .= " [Dive].[DiveChiefScientist] = '$qDiveChfSci'" if $qDiveChfSci;
	$where_clause .= "\n  $conjunction " if (($qNotes || $qMission || $qAnnotators || $qDiveName || $qDiveChfSci) && $qCamera);
	$where_clause .= " [TapeSummary].[Camera] = '$qCamera'" if $qCamera;
	
	
	$date_annotated_clause = "DateAnnotated between $qSDate and $qEDate" if ($qSDate && $qEDate);
	if ($where_clause eq "WHERE " && $date_annotated_clause) {
		$where_clause .= $date_annotated_clause;
	}
	elsif ($where_clause ne "WHERE " && $date_annotated_clause) {
		$where_clause .= " $conjunction $date_annotated_clause";
	}
	
	# Override DateAnnotated range request if IS NULL is checked
	$date_annotated_clause = "DateAnnotated IS NULL" if $dateAnnotatedIsNull;
	if ( $where_clause eq "WHERE " && $dateAnnotatedIsNull ) {
		$where_clause .= $date_annotated_clause;
	}
	elsif ($where_clause ne "WHERE " && $dateAnnotatedIsNull ) {
		$where_clause .= " $conjunction $date_annotated_clause";
	}
	
	#
	# Single Dive Date
	#
	if ( GetFormValue('qDiveDay') && GetFormValue('qDiveMon') && GetFormValue('qDiveYr') ) {
		$dive_date_clause = " (convert(varchar, [Dive].[DiveStartDtg], 101) = ";
		$dive_date_clause .= "'" . GetFormValue('qDiveMon') . '/' . GetFormValue('qDiveDay') . '/' . GetFormValue('qDiveYr') . "')";
		
	}
	if ($where_clause eq "WHERE " && $dive_date_clause) {
		$where_clause .= $dive_date_clause;
	}
	elsif ($where_clause ne "WHERE " && $dive_date_clause) {
		$where_clause .= " $conjunction $dive_date_clause";
	}
	
	#
	# Expedition ShipName
	#
	if ( $where_clause ne "WHERE " && $qShipName ) {
		$where_clause .= " $conjunction [Expedition].[ShipName] ='$qShipName'";
	}
	elsif ( $qShipName ) {
		$where_clause .= " [Expedition].[ShipName] ='$qShipName'";
	}	
	
	#
	# Expedition date range 
	#
	if ( GetFormValue('qDates') ) {
		if ( GetFormValue('qDates') =~ /last(\d+)/ ) {
			$expd_date_clause = "( DateDiff(\"dy\", ScheduledEndDtg, getdate()) < $+ )";
			$expd_date_clause .= " AND ( DateDiff(\"dy\", ScheduledStartDtg, getdate()) >= 0 )";
		}
		elsif ( GetFormValue('qDates') =~ /\d\d\d\d/ ) {
			$expd_date_clause = "( DatePart(\"yy\", ScheduledStartDtg) = " . GetFormValue('qDates') . ")";
		}
		elsif ( GetFormValue('qDates') =~ /Future/i ) {
			$expd_date_clause = "( DateDiff(\"dy\", ScheduledStartDtg, getdate()) <= 0 )";
		}
		else {
			$expd_date_clause = "";
		}
	
	}
	elsif ( GetFormValue('qExpdSmon') && GetFormValue('qExpdSyr') && GetFormValue('qExpdEmon') && GetFormValue('qExpdEyr') ) {
		$expd_date_clause = "( ScheduledStartDtg BETWEEN '1 ";
		$expd_date_clause .= GetFormValue('qExpdSmon') . ' ' . GetFormValue('qExpdSyr');
		$expd_date_clause .= "' AND '1 " . GetFormValue('qExpdEmon') . ' ' . GetFormValue('qExpdEyr');
		$expd_date_clause .= "')";
	}
	
	else {
		$expd_date_clause = '';
	}
	
	if ($where_clause eq "WHERE " && $expd_date_clause) {
		$where_clause .= $expd_date_clause;
	}
	elsif ($where_clause ne "WHERE " && $expd_date_clause) {
		$where_clause .= " $conjunction $expd_date_clause";
	}
	
	#
	# Expedition YYYYDDD query
	#
	if ( GetFormValue('qYYYYDDD') ) {
		my $yyyy = substr(GetFormValue('qYYYYDDD'),0,4);
		my $ddd = substr(GetFormValue('qYYYYDDD'),4,3);
		if ( ($ddd > 0 && $ddd < 366) && $yyyy > 1986 ) {
			$yyyyddd_clause = "((DATEPART(yy,ScheduledStartDtg) = $yyyy ";
			$yyyyddd_clause .= "AND DATEPART(yy,ScheduledEndDtg) = $yyyy ";
			$yyyyddd_clause .= "AND (DATEPART(dy,ScheduledStartDtg) <= $ddd ";
			$yyyyddd_clause .= "AND DATEPART(dy,ScheduledEndDtg) >= $ddd)) ";
			$yyyyddd_clause .= "OR ";
			$yyyyddd_clause .= "(DATEPART(yy,StartDtg) = $yyyy ";
			$yyyyddd_clause .= "AND DATEPART(yy,EndDtg) = $yyyy ";
			$yyyyddd_clause .= "AND (DATEPART(dy,StartDtg) <= $ddd ";
			$yyyyddd_clause .= "AND DATEPART(dy,EndDtg) >= $ddd) ))";
		}
		else {
			$yyyyddd_clause = '';
		}
	}
	else {
		$yyyyddd_clause = '';
	}
	
	if ($where_clause eq "WHERE " && $yyyyddd_clause) {
		$where_clause .= $yyyyddd_clause;
	}
	elsif ($where_clause ne "WHERE " && $yyyyddd_clause) {
		$where_clause .= " $conjunction $yyyyddd_clause";
	}
	
	
	
	$sql =<<EOS;
SELECT 
	[Dive].[DiveChiefScientist],
	convert(varchar, [Dive].[DiveStartDtg], 101) AS DiveStartDate,
	[TapeSummary].[Camera],
	[TapeSummary].[Application],
	[TapeSummary].[Style],
	[TapeSummary].[Notes],
	[TapeSummary].[Mission],
	[TapeSummary].[RovName] AS RovName,
	[TapeSummary].[DiveNumber] AS DiveNumber,
	[TapeSummary].[Annotators],
	[TapeSummary].[NumSDTapes],
	[TapeSummary].[NumHDTapes],
	[TapeSummary].[NumFileSegments],
	[TapeSummary].[DateAnnotated],
	[TapeSummary].[HoursOfVideo],
	[TapeSummary].[HoursAnnotated],
	DATEPART(yy, [Dive].[DiveStartDtg]) AS YYYY,
	DATEPART(dy, [Dive].[DiveStartDtg]) AS JDay
from
	[dbo].[Dive] [Dive] 
		inner join [dbo].[Expedition] [Expedition] 
		on [Dive].[ExpeditionID_FK] = [Expedition].[ExpeditionID] 
			inner join [dbo].[TapeSummary] [TapeSummary] 
			on [Dive].[RovName] = [TapeSummary].[RovName] and
			[Dive].[DiveNumber] = [TapeSummary].[DiveNumber]	
$where_clause 
ORDER BY DiveStartDate
EOS

	Win32::ASP::DebugPrint("\nqueryDB():\nSQL = $sql ") if $resultsAs eq 'diveList';
	$RS = $Conn->Execute($sql);
	
	$Errors = $Conn->Errors();
	if ( keys %$Errors ) { 
		$Response->Write("<p>Database query Error(s):</p>");
		foreach $error (keys %$Errors)
	        {
	      		$Response->Write($error->{Description});
	   	}
	   	$Response->Write("<pre>$sql</pre>");
	   	$Conn->Close;
		return -1;
	 }
	 else {
		%notesHash = ();
		%missionHash = ();
		%annotatorsHash = ();
		%dateAnnotatedHash = ();
		%numSDTapesHash = ();
		%numHDTapesHash = ();
		%numFileSegmentsHash = ();
		%cameraHash = ();
		%applicationHash = ();
		%styleHash = ();
		%hoursOfVideoHash = ();
		%hoursAnnotatedHash = ();
		%diveStartDateHash = ();
		%diveChfSciHash = ();
		%diveYYYYHash = ();
		%jDayHash = ();
	
		$sumHoursOfVideo = 0;
		$sumHoursAnnotated = 0;
		while ( ! $RS->EOF ) {
			$diveStr = sprintf("%s%d", substr($RS->Fields('RovName')->value, 0,1), $RS->Fields('DiveNumber')->value);
			push @diveList, $diveStr;
			
			# Save html-marked-up TapeSummary fields for display
			$notesHash{$diveStr} = $RS->Fields('Notes')->value;
			$notesHash{$diveStr} =~ s/($qNotes)/\<b\>$1\<\/b\>/i if ($qNotes && $resultsAs eq 'diveList');
			$missionHash{$diveStr} = $RS->Fields('Mission')->value;
			$missionHash{$diveStr} =~ s#($qMission)#\<b\>$1\<\/b\>#i if ($qMission && $resultsAs eq 'diveList');
			$annotatorsHash{$diveStr} = $RS->Fields('Annotators')->value;
			$annotatorsHash{$diveStr} =~ s/($qAnnotators)/\<b\>$1\<\/b\>/ if ($qAnnotators && $resultsAs eq 'diveList');
			$dateAnnotatedHash{$diveStr} = $RS->Fields('DateAnnotated')->value;
			$numSDTapesHash{$diveStr} = $RS->Fields('NumSDTapes')->value;
			$numHDTapesHash{$diveStr} = $RS->Fields('NumHDTapes')->value;
			$numFileSegmentsHash{$diveStr} = $RS->Fields('NumFileSegments')->value;
			$cameraHash{$diveStr} = $RS->Fields('Camera')->value;
			$applicationHash{$diveStr} = $RS->Fields('Application')->value;
			$styleHash{$diveStr} = $RS->Fields('Style')->value;
			$hoursOfVideoHash{$diveStr} = $RS->Fields('HoursOfVideo')->value;
			$hoursAnnotatedHash{$diveStr} = $RS->Fields('HoursAnnotated')->value;
			$diveStartDateHash{$diveStr} = $RS->Fields('DiveStartDate')->value;
			$diveChfSciHash{$diveStr} = $RS->Fields('DiveChiefScientist')->value;
			$diveYYYYHash{$diveStr} = $RS->Fields('YYYY')->value;
			$jDayHash{$diveStr} = $RS->Fields('JDay')->value;
			
			$sumHoursOfVideo += $RS->Fields('HoursOfVideo')->value;
			$sumHoursAnnotated += $RS->Fields('HoursAnnotated')->value;
			Win32::ASP::DebugPrint("\nqueryDB():\nPushed dive = $diveStr ") if $resultsAs eq 'diveList';
			$RS->MoveNext;
		}
		
		# Call appropriate method for the response requested
		if ( $resultsAs eq 'diveList' ) {
			diveListResponse(\@diveList, $sql, \%notesHash, \%missionHash, \%annotatorsHash, \%dateAnnotatedHash, 
				\%numSDTapesHash, \%numHDTapesHash, \%numFileSegmentsHash, \%cameraHash, \%applicationHash, \%styleHash,
				\%hoursOfVideoHash, \%hoursAnnotatedHash, \%diveStartDateHash, \%diveChfSciHash, $sumHoursOfVideo, $sumHoursAnnotated);
		}
		elsif ( $resultsAs eq 'excelTxt' ) {
			excelResponse(\@diveList, $sql, \%notesHash, \%missionHash, \%annotatorsHash, \%dateAnnotatedHash, 
				\%numSDTapesHash, \%numHDTapesHash, \%numFileSegmentsHash, \%cameraHash, \%applicationHash, \%styleHash,
				\%hoursOfVideoHash, \%hoursAnnotatedHash, \%diveStartDateHash, \%diveChfSciHash, $sumHoursOfVideo, $sumHoursAnnotated);			
		}
		elsif ( $resultsAs eq 'tapeLabelPdf' ) {		
			tapeLabelsResponse(\@diveList, $sql, \%notesHash, \%missionHash, \%annotatorsHash, \%dateAnnotatedHash, 
				\%numSDTapesHash, \%numHDTapesHash, \%numFileSegmentsHash, \%cameraHash, \%applicationHash, \%styleHash,
				\%hoursOfVideoHash, \%hoursAnnotatedHash, \%diveStartDateHash, \%diveChfSciHash, $sumHoursOfVideo, $sumHoursAnnotated,
				\%diveYYYYHash, \%jDayHash );
		}
	 }	

	$RS->Close;
	$Conn->Close;
	return 0;

}	# End queryDB()


#--------------------------------------------------------------------
#

=head3 diveListResponse()

Return results as a list of dives.

Author: Mike McCann

Date Created: 7/28/2011

=cut

sub diveListResponse{

	( $rdiveList, $tsQuery, $rnotesHash = $_[2], $rmissionHash, $rannotatorsHash, $rdateAnnotatedHash,
		$rnumSDTapesHash, $rnumHDTapesHash, %numFileSegmentsHash, $rcameraHash, $rapplicationHash, $rstyleHash,
		$rhoursOfVideoHash, $rhoursAnnotatedHash, $rdiveStartDateHash, $rdiveChfSciHash, $sumHoursOfVideo , $sumHoursAnnotated ) = @_;

	$numDives = $#$rdiveList + 1;
	%>
	<h3><%=$numDives%> dives with Tape Summary information (<%=$sumHoursOfVideo%> Total hours of video)</h3>

	<form method="GET" action="<%= $this_script%>" id="form1" name="form1">
	  <input type="hidden" name="step" value="QueryDB">
	  <input type="hidden" name="resultsAs" value="excelTxt">
	  <%
	  ##Win32::ASP::DebugPrint("\nForm entries\n");
	  # Must use POST for this to work
	  foreach my $name ( Win32::OLE::in ( $Request->Form ) ) {
		##Win32::ASP::DebugPrint("\n$name = " . $Request->Form($name)->{item} ) if $Request->Form($name)->{item};
		$value = $Request->Form($name)->{item};
		if ( $value ) {
			%>
	    <input type="hidden" name="<%=$name%>" value="<%=$value%>"><%	
		}
	  }
	  %>
	  
	  <input type="submit" name="Continue" value="These results in Excel-importable tab-delimited format"> 
	</form>
	
	<%
	# Count HD tapes
	my $hdCountSum = 0;
	foreach my $hdCount ( values %$rnumHDTapesHash) {
		$hdCountSum += $hdCount;
	}
	
	# Count SD tapes
	my $sdCountSum = 0;
	foreach my $sdCount ( values %$rnumSDTapesHash) {
		$sdCountSum += $sdCount;
	}
	
	# Count File Segments
	my $fsCountSum = 0;
	foreach my $fsCount ( values %$rnumFileSegmentsHash) {
		$fsCountSum += $fsCount;
	}
	%>
	
	<form method="POST" action="<%=$this_script%>" onsubmit="return Field_Validator(this)" name="Field">
	<input type="hidden" name="step" value="QueryDB">
	<input type="hidden" name="resultsAs" value="tapeLabelPdf">
	<%
	# Must use POST for this to work
	foreach my $name ( Win32::OLE::in ( $Request->Form ) ) {
		##Win32::ASP::DebugPrint("\n$name = " . $Request->Form($name)->{item} ) if $Request->Form($name)->{item};
		$value = $Request->Form($name)->{item};
		if ( $value ) {
			%>
		<input type="hidden" name="<%=$name%>" value="<%=$value%>"><%	
		}
	}
	if ($hdCountSum != 0) {
	%>
	<input type="submit" name="Submit" value="PDF of <%=$hdCountSum%> HD Tape labels for these results<%=$rovName%><%=$diveNumber%>" >
	<% }
	elsif ($sdCountSum != 0) {
	%>
	<input type="submit" name="Submit" value="PDF of <%=$sdCountSum%> SD Tape labels for these results<%=$rovName%><%=$diveNumber%>"  >
	<%
	}
	%>
	</form>
	
	<dl>
	<%	

	foreach $d ( @$rdiveList ) {
		Win32::ASP::DebugPrint("\ndiveListResponse():\nd = $d ");
		$rovCode = 'vnta' if $d =~ /^v/i;
		$rovCode = 'tibr' if $d =~ /^t/i;
		$rovCode = 'docr' if $d =~ /^d/i;
		$rovCode = 'mini' if $d =~ /^m/i;

		$diveNumber = substr($d,1);

		# Query dive table for Expedition link info
		$sql =<<EOS;
select ExpeditionID_FK
from dive
where RovName = '$rovCode' and DiveNumber = $diveNumber
EOS

		Win32::ASP::DebugPrint("\ndiveListResponse():\nSQL = $sql ");
		$RS = $Conn->Execute($sql);
	
		$Errors = $Conn->Errors();
		if ( keys %$Errors ) { 
			$Response->Write("Database query Error(s):<pre>$sql</pre>");
			foreach $error (keys %$Errors)
				{
					$Response->Write($error->{Description});
			}
			$Conn->Close;
			return -1;
		 }
		 else {
			$expeditionID = $RS->Fields('ExpeditionID_FK')->Value;
		}

		
		%>
		<dt>
		<b><a href="tapeSummary_karen.asp?rovName=<%=$rovCode%>&diveNumber=<%=$diveNumber%>"><%=$$rhoursAnnotatedHash{$d}%> hours annotated</a></b> for Dive 
		<a href="dive_karen.asp?RovName=<%=$rovCode%>&DiveNumber=<%=$diveNumber%>&edit=yes"><%=$rovCode%><%=$diveNumber%></a> from Expedition
		<a href="postcruise_karen.asp?step=3&ExpeditionID=<%=$expeditionID%>&edit=<%=$expeditionID%>"><%=$expeditionID%></a>
		</dt>
		<dd>
		<b>Date:</b> <%=$$rdiveStartDateHash{$d}%>
		<br>
		<b>Chief Scientist:</b> <%=$$rdiveChfSciHash{$d}%>
		<br>
		<b>Notes:</b> <%=$$rnotesHash{$d}%>
		<br>
		<b>Mission:</b> <%=$$rmissionHash{$d}%>
		<br>
		<b>Annotators:</b> <%=$$rannotatorsHash{$d}%>
		<br>
		<b>DateAnnotated:</b> <%=$$rdateAnnotatedHash{$d}%>
		<br>
		</dd>
		<%

	} # End foreach $d ( @$rdiveList ) 

	%>
	</dl>
	Total Hours of Video = <%=$sumHoursOfVideo%>
	Total Hours Annotated = <%=$sumHoursAnnotated%>
	<hr>
	<h4>Results of SQL query</h4>
	<p><pre><%=$tsQuery%></pre></p>
	<%

} # End diveListResponse()



#--------------------------------------------------------------------
#

=head3 excelResponse()

Return results of TapeSummary query in tab-delimted format.
Note that if any other HTML or debug output happens then the print will not work.

Author: Mike McCann

Date Created: 8/16/2011

=cut

sub excelResponse{

	( $rdiveList, $tsQuery, $rnotesHash = $_[2], $rmissionHash, $rannotatorsHash, $rdateAnnotatedHash,
		$rnumSDTapesHash, $rnumHDTapesHash, $rnumFileSegmentsHash, $rcameraHash, $rapplicationHash, $rstyleHash,
		$rhoursOfVideoHash, $rhoursAnnotatedHash, $rdiveStartDateHash, $rdiveChfSciHash, $sumHoursOfVideo , $sumHoursAnnotated ) = @_;

	$numDives = $#$rdiveList + 1;
	
	# Original headings
	##$header = "Vehicle\tYear\tYear-date\tDate\tChief Scientist(s)\tSD Tapes\tHD Tapes\tCamera\tCameralog Filename\tDiveNum\tNotes";
	##$header .= "\tMission\tApplication\tAnnotation Filename\tRT File Edited?\tStyle\tTotal Tapes Annotated\tAnnotator\tDate Annotated\tHours of Video\n";
	
	$header = "Vehicle\tDate\tChief Scientist(s)\tSD Tapes\tHD Tapes\tFile Segments\tCamera\tDiveNum\tNotes";
	$header .= "\tMission\tApplication\tStyle\tTotal Tapes Annotated\tAnnotator\tDate Annotated\tHours of Video\tHours Annotated\n";
	
	print $header;

	$recs = '';
	foreach $d ( @$rdiveList ) {
		
		$vehicle = 'Ventana' if $d =~ /^v/i;
		$vehicle = 'Tiburon' if $d =~ /^t/i;
		$vehicle = 'Doc Ricketts' if $d =~ /^d/i;

		$diveNumber = substr($d,1);
		$totalNumTapes = $$rnumSDTapesHash{$d} + $$rnumHDTapesHash{$d};
		
		$recs .= "$vehicle\t$$rdiveStartDateHash{$d}\t$$rdiveChfSciHash{$d}\t$$rnumSDTapesHash{$d}\t$$rnumHDTapesHash{$d}";
		$recs .= "\t$$rnumFileSegmentsHash{$d}";
		$recs .= "\t$$rcameraHash{$d}\t$diveNumber\t$$rnotesHash{$d}";
		$recs .= "\t$$rmissionHash{$d}\t$$rapplicationHash{$d}\t$$rstyleHash{$d}\t$totalNumTapes\t$$rannotatorsHash{$d}";
		$recs .= "\t$$rdateAnnotatedHash{$d}";
		$recs .= "\t$$rhoursOfVideoHash{$d}";
		$recs .= "\t$$rhoursAnnotatedHash{$d}\n";

	} # End foreach $d ( @$rdiveList ) 

	print $recs;

} # End excelResponse()



#--------------------------------------------------------------------
#

=head3 edit()

Present form for user to fill out.  Grab a lot of information from VARS
for the tape and present text boxes for the user to add additional information
to the TapeSummary table in Expd.

Author: Mike McCann

Date Created: 7/28/2011

=cut

sub edit {
	
	my $rovName = $_[0];
	my $diveNumber = $_[1];
	my $freeFormFields = $_[2] || 0;

	open_database($dsn);		# Creates $Conn object for Expd as a global variable
	open_vars_database();		# Creates $ConnVARS for VARS global object

	$varsRovName = 'Doc Ricketts' if $rovName eq 'DocRicketts' || $rovName eq 'docr';
	$varsRovName = 'Tiburon' if $rovName eq 'tibr';
	$varsRovName = 'Ventana' if $rovName eq 'vnta';
	$varsRovName = 'MiniROV' if $rovName eq 'mini';
	
	#
	# Query VARS for list of tapes on this dive
	#
	my $sql_old =<<EOS_OLD;
select
	[VideoArchive].[videoArchiveName] as TapeName,
	[EXPDMergeStatus].[MergeDate],
	[VideoArchiveSet].[LAST_UPDATED_TIME],
	[VideoArchiveSet].[FormatCode] 
from
	[dbo].[CameraPlatformDeployment] [CameraPlatformDeployment] 
		inner join [dbo].[VideoArchiveSet] [VideoArchiveSet] 
		on [CameraPlatformDeployment].[VideoArchiveSetID_FK] = [VideoArchiveSet]
		.[id] 
			inner join [dbo].[EXPDMergeStatus] [EXPDMergeStatus] 
			on [EXPDMergeStatus].[VideoArchiveSetID_FK] = [VideoArchiveSet].[id] 
				inner join [dbo].[VideoArchive] [VideoArchive] 
				on [VideoArchive].[VideoArchiveSetID_FK] = [VideoArchiveSet].
				[id] 
where
	([CameraPlatformDeployment].[SeqNumber] = $diveNumber) and
	([VideoArchiveSet].[PlatformName] = '$varsRovName')
order by 
	TapeName
	
EOS_OLD
# Query provided by Brian which ASP complains about syntax error at 'WITH'
#WITH  
#video_tapes AS
#  (
#    SELECT  
#      vas.id,
#      va.videoArchiveName AS TapeName,
#      vas.LAST_UPDATED_TIME,
#      vas.FormatCode
#    FROM
#      CameraPlatformDeployment AS cpd RIGHT JOIN  
#      VideoArchiveSet AS vas ON cpd.VideoArchiveSetID_FK = vas.id  LEFT JOIN  
#      VideoArchive AS va ON va.VideoArchiveSetID_FK = vas.id
#    WHERE
#      cpd.SeqNumber = $diveNumber AND
#      vas.PlatformName = '$varsRovName'
#  ),
#merge_status AS
#  (
#    SELECT TOP 1
#      vt.id,
#      ms.MergeDate
#    FROM  
#      EXPDMergeHistory AS ms JOIN
#      video_tapes AS vt ON vt.id = ms.VideoArchiveSetID_FK
#    ORDER BY
#      ms.MergeDate DESC
#  )
#SELECT
#  video_tapes.TapeName,
#  merge_status.MergeDate,
#  video_tapes.LAST_UPDATED_TIME,
#  video_tapes.FormatCode
#FROM
#  video_tapes LEFT JOIN
#  merge_status ON merge_status.id = video_tapes.id
#ORDER BY
#  TapeName

	my $sql =<<EOS;
select
	Max([EXPDMergeHistory].[MergeDate]) as LastMergeDate,
	[VideoArchiveSet].[FormatCode],
	[VideoArchiveSet].[LAST_UPDATED_TIME],
	[VideoArchive].[videoArchiveName] as TapeName
from
	[dbo].[CameraPlatformDeployment] [CameraPlatformDeployment] 
		inner join [dbo].[VideoArchiveSet] [VideoArchiveSet] 
		on [CameraPlatformDeployment].[VideoArchiveSetID_FK] = [VideoArchiveSet]
		.[id] 
			inner join [dbo].[VideoArchive] [VideoArchive] 
			on [VideoArchive].[VideoArchiveSetID_FK] = [VideoArchiveSet].[id] 
				inner join [dbo].[EXPDMergeHistory] [EXPDMergeHistory] 
				on [EXPDMergeHistory].[VideoArchiveSetID_FK] = [VideoArchiveSet]
				.[id] 
where
	([CameraPlatformDeployment].[SeqNumber] =$diveNumber) and
	([VideoArchiveSet].[PlatformName] ='$varsRovName') 
group by	
	[VideoArchive].[videoArchiveName], [VideoArchiveSet].[FormatCode],
	[VideoArchiveSet].[LAST_UPDATED_TIME]
EOS

	Win32::ASP::DebugPrint("\nedit():\nExecuting VARS SQL: $sql");
	
	my $RSvars = $ConnVARS->Execute($sql);
	
	$Errors = $ConnVARS->Errors();
	if ( keys %$Errors ) { 
		$Response->Write("Database Error(s): ");
		foreach $error (keys %$Errors)
		{
			$Response->Write($error->{Description});
		}
	}
	
	if(!$RSvars) {
	 	$Response->Write("<BR>Empty Return Set: RSvars. for <br>$sql ");
	}	
	
	%mergeDateHash = ();
	%lastUpdatedTimeHash = ();
	%formatCodeHash = ();
	%observersHash = ();
	while ( !$RSvars->EOF ) {
		#
		# Do one-to-many query to get observers (Annotatotrs)
		#
		$tapeName = $RSvars->Fields('TapeName')->Value;
		$sql = <<EOS;
select distinct
	[Observation].[Observer] 
from
	[dbo].[Observation] [Observation] 
		inner join [dbo].[VideoFrame] [VideoFrame] 
		on [Observation].[VideoFrameID_FK] = [VideoFrame].[id] 
			inner join [dbo].[VideoArchive] [VideoArchive] 
			on [VideoFrame].[VideoArchiveID_FK] = [VideoArchive].[id] 
where
	([VideoArchive].[videoArchiveName] ='$tapeName')
EOS
		Win32::ASP::DebugPrint("\nedit():\nExecuting VARS SQL to get observers: $sql");
		
		my $RSobservers = $ConnVARS->Execute($sql);
		if(!$RSvars) {
		 	 $Response->Write("<BR>Empty Return Set: RSobservers. for <br>$sql ");
		}	
		my $observers = '';
		while ( !$RSobservers->EOF ) { 
			$observers .= $RSobservers->Fields('Observer')->value . ', ';
			$RSobservers->MoveNext;
		}
		$observers =~ s/, $//;
		
		#
		# Save the VARS tape infor into a hash for display on the web page
		#
		$mergeDateHash{$tapeName} = $RSvars->Fields('LastMergeDate')->value;
		$lastUpdatedTimeHash{$tapeName} = $RSvars->Fields('LAST_UPDATED_TIME')->Value;
                $formatCodeHash{$tapeName} = $RSvars->Fields('FormatCode')->Value;
		$observersHash{$tapeName} = $observers;
		
		$RSvars->MoveNext;
		
	} # End while ( !$RSvars->EOF ) 
	
	#
	# Query VARS to get duration information
	#
	$sql =<<EOS;
	select * 
from

(
select  top(1)
	[VideoFrame].[RecordedDtg],
	DATEDIFF(second, '1970-01-01', RecordedDtg) AS Epoch,
    	[VideoFrame].[TapeTimeCode]
from
	[dbo].[CameraPlatformDeployment] [CameraPlatformDeployment] 
		inner join [dbo].[VideoArchiveSet] [VideoArchiveSet] 
		on [CameraPlatformDeployment].[VideoArchiveSetID_FK] = [VideoArchiveSet]
		.[id] 
			inner join [dbo].[VideoArchive] [VideoArchive] 
			on [VideoArchive].[VideoArchiveSetID_FK] = [VideoArchiveSet].[id] 
				inner join [dbo].[VideoFrame] [VideoFrame] 
				on [VideoFrame].[VideoArchiveID_FK] = [VideoArchive].[id] 
where
	([CameraPlatformDeployment].[SeqNumber] = $diveNumber) and
	([VideoArchiveSet].[PlatformName] = '$varsRovName') 
order by
	[VideoFrame].[RecordedDtg]

union all

select  top(1)
	[VideoFrame].[RecordedDtg],
	DATEDIFF(second, '1970-01-01', RecordedDtg) AS Epoch,
    	[VideoFrame].[TapeTimeCode]
from
	[dbo].[CameraPlatformDeployment] [CameraPlatformDeployment] 
		inner join [dbo].[VideoArchiveSet] [VideoArchiveSet] 
		on [CameraPlatformDeployment].[VideoArchiveSetID_FK] = [VideoArchiveSet]
		.[id] 
			inner join [dbo].[VideoArchive] [VideoArchive] 
			on [VideoArchive].[VideoArchiveSetID_FK] = [VideoArchiveSet].[id] 
				inner join [dbo].[VideoFrame] [VideoFrame] 
				on [VideoFrame].[VideoArchiveID_FK] = [VideoArchive].[id] 
where
	([CameraPlatformDeployment].[SeqNumber] = $diveNumber) and
	([VideoArchiveSet].[PlatformName] = '$varsRovName') 
order by
	[VideoFrame].[RecordedDtg] desc
)
as a
EOS

	#
	# Comment out this query while VARS is being transitioned to the new M3 system - mpm 5 October 2018
	#
	##Win32::ASP::DebugPrint("\nedit():\nExecuting VARS SQL to get duration: $sql");
	##	
	##$RSduration = $ConnVARS->Execute($sql);
	##if(!$RSduration) {
	## 	 $Response->Write("<BR>Empty Return Set: RSduration. for <br>$sql ");
	##}	
	
	Win32::ASP::DebugPrint("\nedit():\nExecution of VARS SQL to get duration has been commented out because it often delays page loading by 30 seconds");
	
	#
	# Do Expd query to get its TapeSummary information, and ReturnSets for picklist options
	#
	$sql =<<EOS;	
SELECT 
	[Dive].[DiveChiefScientist],
	convert(varchar, [Dive].[DiveStartDtg], 101) AS DiveStartDate,
	[TapeSummary].[Camera],
	[TapeSummary].[Notes],
	[TapeSummary].[Mission],
	[TapeSummary].[RovName] AS RovName,
	[TapeSummary].[DiveNumber] AS DiveNumber,
	[TapeSummary].[Application],
	[TapeSummary].[Style],
	[TapeSummary].[Annotators],
	[TapeSummary].[DateAnnotated],
	[TapeSummary].[NumSDTapes],
	[TapeSummary].[NumHDTapes],
	[TapeSummary].[NumFileSegments],
	[TapeSummary].[HoursOfVideo],
	[TapeSummary].[HoursAnnotated],
	[TapeSummary].[RTFileEdited],
	[TapeSummary].[AnnotationFileName]
FROM 	[dbo].[Dive] [Dive] 
		inner join [dbo].[TapeSummary] [TapeSummary] 
		on [Dive].[DiveNumber] = [TapeSummary].[DiveNumber] and
		[Dive].[RovName] = [TapeSummary].[RovName]
where [TapeSummary].[RovName] = '$rovName' and [TapeSummary].[DiveNumber] = $diveNumber
EOS

	Win32::ASP::DebugPrint("\nedit():\nExecuting Expd SQL: $sql");
	
	my $RSexpd = $Conn->Execute($sql);
	
	$Errors = $Conn->Errors();
	if ( keys %$Errors ) { 
		$Response->Write("Database Load Error(s):<pre>$sql</pre>");
		foreach $error (keys %$Errors)
	        {
	      		$Response->Write($error->{Description});
	   	}
	   	$Conn->Close;
	 }
	
	if(!$RSexpd) {
	 	 $Response->Write("<BR>Empty Return Set: RSexpd. for <br>$sql ");
	}	
	
	$sql =<<EOS;
SELECT distinct camera
  FROM [EXPD].[dbo].[TapeSummary]
  ORDER BY Camera
EOS
	Win32::ASP::DebugPrint("\nedit():\nExecuting Expd SQL: $sql");
	
	my $RScameras = $Conn->Execute($sql);
	if(!$RScameras) {
	 	 $Response->Write("<BR>Empty Return Set: RScameras. for <br>$sql ");
	}	
	  
	$sql =<<EOS;
SELECT distinct application
  FROM [EXPD].[dbo].[TapeSummary]
  ORDER BY application
EOS
	Win32::ASP::DebugPrint("\nedit():\nExecuting Expd SQL: $sql");
	
	my $RSapplications = $Conn->Execute($sql);
	if(!$RSapplications) {
	 	 $Response->Write("<BR>Empty Return Set: RSapplications. for <br>$sql ");
	}	
		  
	$sql =<<EOS;
SELECT distinct style
  FROM [EXPD].[dbo].[TapeSummary]
  ORDER BY style
EOS
	Win32::ASP::DebugPrint("\nedit():\nExecuting Expd SQL: $sql");
	
	my $RSstyles = $Conn->Execute($sql);
	if(!$RSstyles) {
	 	 $Response->Write("<BR>Empty Return Set: RSstyles. for <br>$sql ");
	}	
	
	$sql =<<EOS;
SELECT distinct annotators
  FROM [EXPD].[dbo].[TapeSummary]
  ORDER BY annotators
EOS
	Win32::ASP::DebugPrint("\nedit():\nExecuting Expd SQL: $sql");
	
	my $RSannotators = $Conn->Execute($sql);
	if(!$RSannotators) {
	 	 $Response->Write("<BR>Empty Return Set: RSannotators. for <br>$sql ");
	}	
%>

     
    <h2><FONT color=#408080>Expd Information</FONT></h2>
    <%
    if ( $RSexpd->Fields(0)->value ) {
    ##if ( 1 ) {
    %>
    
    <form method="POST" action="<%=$this_script%>" onsubmit="return Field_Validator(this)" name="Field">
    <input type="hidden" name="step" value="UpdateTapeSummary">
    <input type="hidden" name="rovName" value="<%=$rovName%>">
    <input type="hidden" name="diveNumber" value="<%=$diveNumber%>">
    <table border="0" cellpadding="5" cellspacing="0">
    
    	<tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Chief Scientist:</font></strong></td>
            <td><%=$RSexpd->Fields('DiveChiefScientist')->Value%></td>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Dive Date:</font></strong></td>
            <td><%=$RSexpd->Fields('DiveStartDate')->Value%></td>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Camera:</font></strong></td>
			<% if ( $freeFormFields ) { %>
				<td><input type="text" name="Camera" size="20" value="<%=$RSexpd->Fields('Camera')->Value%>"></td>
			<% }
			else { 
			%>
            <td><select name="Camera" size="1">
            	<%while (!$RScameras->EOF) {
            		$cameraChoice = $RScameras->Fields('Camera')->value;
            		##$cameraChoice =~ s/\s+$//;
					my $camera_option = "<option value=\"$cameraChoice\"";
					$camera_option .= "selected" if $RSexpd->Fields('Camera')->Value eq $cameraChoice;
					$camera_option .= ">$cameraChoice</option>\n";
					$Response->write($camera_option);
					$RScameras->MoveNext;
				 } %>
            	<option 
            </select></td>
			<%
			}
			%>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Notes:</font></strong></td>
            <td><textarea rows="1" name="Notes" cols="80" wrap><%=$RSexpd->Fields('Notes')->Value%></textarea></td>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Mission:</font></strong></td>
            <td><textarea rows="1" name="Mission" cols="32" wrap><%=$RSexpd->Fields('Mission')->Value%></textarea></td>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Application:</font></strong></td>
			<% if ( $freeFormFields ) { %>
				<td><input type="text" name="Application" size="20" value="<%=$RSexpd->Fields('Application')->Value%>"></td>
			<% }
			else { 
			%>
            <td><select name="Application" size="1">
            	<%while (!$RSapplications->EOF) {
            		$applicationChoice = $RSapplications->Fields('Application')->value;
            		##$applicationChoice =~ s/\s+$//;
					my $application_option = "<option value=\"$applicationChoice\"";
					$application_option .= " selected" if $RSexpd->Fields('Application')->Value eq $applicationChoice;
					$application_option .= ">$applicationChoice</option>\n";
					$Response->write($application_option);
					$RSapplications->MoveNext;
				 } %>
            	<option 
            </select></td>
			<%
			}
			%>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Style:</font></strong></td>
			<% if ( $freeFormFields ) { %>
				<td><input type="text" name="Style" size="20" value="<%=$RSexpd->Fields('Style')->Value%>"></td>
			<% }
			else { 
			%>
            <td><select name="Style" size="1">
            	<%while (!$RSstyles->EOF) {
            		$styleChoice = $RSstyles->Fields('Style')->value;
            		##$styleChoice =~ s/\s+$//;
            		##$styleChoice =~ s/"//g;
					my $style_option = "<option value=\"$styleChoice\"";
					$style_option .= "selected" if $RSexpd->Fields('Style')->Value eq $styleChoice;
					$style_option .= ">$styleChoice</option>\n";
					$Response->write($style_option);
					$RSstyles->MoveNext;
				 } %>
            	<option 
            </select></td>
			<%
			}
			%>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">DateAnnotated:</font></strong></td>
            <td><input type="text" name="DateAnnotated" size="12" value="<%=$RSexpd->Fields('DateAnnotated')->Value%>"></td>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Annotators:</font></strong></td>
			<% if ( $freeFormFields ) { %>
				<td><input type="text" name="Annotators" size="20" value="<%=$RSexpd->Fields('Annotators')->Value%>"></td>
			<% }
			else { 
			%>
            <td><select name="Annotators" size="1">
            	<%while (!$RSannotators->EOF) {
            		$annotatorChoice = $RSannotators->Fields('Annotators')->value;
            		##$annotatorChoice =~ s/\s+$//;
            		##$annotatorChoice =~ s/"//g;
					my $annotator_option = "<option value=\"$annotatorChoice\"";
					$annotator_option .= "selected" if $RSexpd->Fields('Annotators')->Value eq $annotatorChoice;
					$annotator_option .= ">$annotatorChoice</option>\n";
					$Response->write($annotator_option);
					$RSannotators->MoveNext;
				 } %>
            	<option 
            </select></td>
			<%
			}
			%>
					
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">RTFileEdited:</font></strong></td>
            <td><%=$RSexpd->Fields('RTFileEdited')->Value%></td>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">AnnotationFileName:</font></strong></td>
            <td><%=$RSexpd->Fields('AnnotationFileName')->Value%></td>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">SD Tapes:</font></strong></td>
            <td><input type="text" name="NumSDTapes" size="2" value="<%=$RSexpd->Fields('NumSDTapes')->Value%>"></td>
        </tr>
        
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">HD Tapes:</font></strong></td>
            <td><input type="text" name="NumHDTapes" size="2" value="<%=$RSexpd->Fields('NumHDTapes')->Value%>"></td>
        </tr>
		
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">File Segments:</font></strong></td>
            <td><input type="text" name="NumFileSegments" size="2" value="<%=$RSexpd->Fields('NumFileSegments')->Value%>"></td>
        </tr>
        
		<tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Hours of Video:</font></strong></td>
            <td><input type="text" name="HoursOfVideo" size="4" value="<%=$RSexpd->Fields('HoursOfVideo')->Value%>"></td>
        </tr>
		
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Hours Annotated:</font></strong></td>
            <td><input type="text" name="HoursAnnotated" size="4" value="<%=$RSexpd->Fields('HoursAnnotated')->Value%>"></td>
        </tr>
        
    </table>
	
	<% if ( ! $freeFormFields ) {
	%>
		<a href= "<%=$this_scrip%>?rovName=<%=$rovName%>&diveNumber=<%=$diveNumber%>&freeFormFields=1">Form with free text fields instead of pick lists</a>
		<br>
	<%}
	else {
	%>
		<a href= "<%=$this_scrip%>?rovName=<%=$rovName%>&diveNumber=<%=$diveNumber%>&freeFormFields=0">Form with pick lists</a>
		<br><br>
	
	<% } %>
	
    <font color="brown"><br>Updates restricted to Video Lab staff</font>
	    <% 
	    if ($vlStaffFlag) {
		%>
		    <input type="submit" class="button blue" value="Submit" name="Step2"> 
		    <input type="reset" class="button gray" value="Reset form fields" name="Reset form">
	    <%
	    }
	    else {
		%>
		    <input disabled type="submit" class="button blue" value="Submit" name="Step2"> 
		    <input disabled type="reset" class="button gray" value="Reset form fields" name="Reset form">
	    <%
	    } %>
    
    
      <!-- Form CSS -->   
	<style type="text/css">
	        button,
            .button {
              font: 12px Arial;
              text-decoration: none;
              color : #fff;
              cursor : pointer;
              border-style : solid;
              border-width : 1px;
            }
            button.blue,
            .button.blue {
              font: 12px Arial;
              text-decoration: none;
              border-color : #2989d8;
              background: #2989d8;
              background: -moz-linear-gradient(top, #2989d8 0%, #1e5799 100%);
              background: -webkit-gradient(linear, left top, left bottom, 
                  color-stop(0%,#2989d8), color-stop(100%,#1e5799));
              background: -webkit-linear-gradient(top, #2989d8 0%,#1e5799 100%);
              background: -o-linear-gradient(top, #2989d8 0%,#1e5799 100%);
              background: -ms-linear-gradient(top, #2989d8 0%,#1e5799 100%);
              background: linear-gradient(top, #2989d8 0%,#1e5799 100%);
                  filter: progid:DXImageTransform.Microsoft.gradient( 
                  startColorstr='#2989d8', endColorstr='#1e5799',GradientType=0 );
            }
            button.gray,
            .button.gray {
               font: 12px Arial;
               font-weight: bold;
               text-decoration: none;
               background-color: #E6E6E6;
               color: #333333;
               padding: 2px 6px 2px 6px;
               border-top: 1px solid #CCCCCC;
               border-right: 1px solid #333333;
               border-bottom: 1px solid #333333;
               border-left: 1px solid #CCCCCC;
            }
	</style>
    </form>
    
    <%}
    else {
    %>
    <p><font color="brown"><strong>No record for this dive in the TapeSummary table.</strong></font></p>
    <form method="GET" action="<%=$this_script%>" onsubmit="return Field_Validator(this)" name="Field">
    <input type="hidden" name="step" value="AddDiveToTapeSummary">
    <input type="hidden" name="rovName" value="<%=$rovName%>">
    <input type="hidden" name="diveNumber" value="<%=$diveNumber%>">
    <input type="submit" class="button blue" name="Submit" value="Add blank record for <%=$rovName%><%=$diveNumber%> to TapeSummary table" >
    <font color="brown">Restricted to Video Lab staff</font>
        <!-- Form CSS -->   
	<style type="text/css">
	        button,
            .button {
              font: 12px Arial;
              text-decoration: none;
              color : #fff;
              cursor : pointer;
              border-style : solid;
              border-width : 1px;
            }
            button.blue,
            .button.blue {
              font: 12px Arial;
              text-decoration: none;
              border-color : #2989d8;
              background: #2989d8;
              background: -moz-linear-gradient(top, #2989d8 0%, #1e5799 100%);
              background: -webkit-gradient(linear, left top, left bottom, 
                  color-stop(0%,#2989d8), color-stop(100%,#1e5799));
              background: -webkit-linear-gradient(top, #2989d8 0%,#1e5799 100%);
              background: -o-linear-gradient(top, #2989d8 0%,#1e5799 100%);
              background: -ms-linear-gradient(top, #2989d8 0%,#1e5799 100%);
              background: linear-gradient(top, #2989d8 0%,#1e5799 100%);
                  filter: progid:DXImageTransform.Microsoft.gradient( 
                  startColorstr='#2989d8', endColorstr='#1e5799',GradientType=0 );
            }
    </style>
    </form>
    <%}
    %>
    
    <h2><FONT color=#408080>VARS information</FONT></h2>
    
    
    
    <table border="1" cellpadding="5" cellspacing="0">
   
   	<tr>
   	    <th>Tape Name</th>
   	    <th>Expd Merge Date</th>
   	    <th>Last Updated Time</th>
   	    <th>Format Code</th>
   	    <th>Annotators</th>
   	</tr>
	<%
	$numberOfVARSTapes = 0;
	@tapeList = ();
	for my $k (sort keys %mergeDateHash) {
		$numberOfVARSTapes++;
		push @tapeList, $k;
	%>
        <tr>
            <td><%=$k%></td>
            <td><%=$mergeDateHash{$k}%></td>
            <td><%=$lastUpdatedTimeHash{$k}%></td>
            <td><%=$formatCodeHash{$k}%></td>
            <td><%=$observersHash{$k}%></td>
        </tr>
        <%
        }
        %>
    </table>
    
    <strong>Total number of tapes:</strong> <%=$numberOfVARSTapes%>
    <br>
    <br>
	
	<%
	if ($RSduration ) {
	%>
    <table>
    	<tr>
   	    <th>&nbsp;</th>
   	    <th>Recorded Time (GMT)</th>
   	    <th>TapeTimeCode</th>
   	</tr>
    <%
    $rowCount = 0;
    while ( !$RSduration->EOF ) { 
    	$rowCount++;
    %>
	<tr>
	<%
	if ($rowCount == 1) {
		$startEpoch = $RSduration->Fields('Epoch')->value;
	%>
		<th>First Annotation</td>
	<%
	}
	elsif ($rowCount == 2) {
		$endEpoch = $RSduration->Fields('Epoch')->value;
	%>
		<th>Last Annotation</td>
	<%
	}
	%>
	<td align="center"><%=$RSduration->Fields('RecordedDtg')->value%></td>
	<td align="center"><%=$RSduration->Fields('TapeTimeCode')->value%></td>
	</tr>
	
    <%	$RSduration->MoveNext;
    }
    $annotationDuration = sprintf('%.2f', ($endEpoch - $startEpoch) / 3600);
    %>
    </table>
	
    <br>
    <table>
    	<tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Elapsed time:</font></strong></td>
            <td><%=$annotationDuration%> hours</td>
        </tr>
        <!--
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Camlog data:</font></strong></td>
            <td>https://mww.mbari.org/expd/queries/camlog-data.asp?ROVName=vnta&DiveNumber=3602</td>
        </tr>
        -->
    </table>
    <%
	}
	%>
    
   
<!--
</td></tr>
</table>
-->


<%
$RSvars->Close;
$RSexpd->Close;
$Conn->Close;
$ConnVARS->Close;

}	# End edit()

%>

<%
#--------------------------------------------------------------------
#

=head3 addDiveToTapeSummary()

Add a blank record to the 

Author: Mike McCann

Date Created: 10/28/98

=cut


sub addDiveToTapeSummary {

	$rovName = $_[0];
	$diveNumber = $_[1];
	
	#
	# Open the database, this time to insert fields
	#
	open_database($dsn);		# Creates $Conn object as a global variable
	
	# Add a dive record to the TapeSummary table
	$sql =<<EOS;
INSERT INTO [EXPD].[dbo].[TapeSummary]
           ([DiveNumber]
           ,[RovName]
           )
           values
           ($diveNumber
           ,'$rovName'
           )
EOS

	Win32::ASP::DebugPrint("\naddDiveToTapeSummary():\nSQL = $sql ");
	$RS = $Conn->Execute($sql);
	
	$Errors = $Conn->Errors();
	if ( keys %$Errors ) { 
		$Response->Write("Database Load Error(s):<pre>$sql</pre>");
		foreach $error (keys %$Errors)
	        {
	      		$Response->Write($error->{Description});
	   	}
	   	$Conn->Close;
		return -1;
	 }
	 else {
	 %>
	 	A blank record for this dive has been added to the TapeSummary table. 
	<%
	 }	


	$RS->Close;
	$Conn->Close;
	return 0;

}	# End addDiveToTapeSummary()


#--------------------------------------------------------------------
#

=head3 updateTapeSummary()

Process the form field entries and update the database record.
Here is where we need to map the names we have on the HTML 
forms for all the fields to the names of the fields that are
actually in the database table.

Author: Mike McCann

Date Created: 8/5/2011

=cut

sub updateTapeSummary {

	$rovName = $_[0];
	$diveNumber = $_[1];
	
	#
	# Open the database, this time to update fields
	#
	open_database($dsn);		# Creates $Conn object as a global variable
	 

	#
	# Loop through all form names and construct SQL update string, exit without
	# complaining if there are no fields to update, an update implies that StatCode
	# goes to plnd.
	#
	$sql = "UPDATE TapeSummary\nSET ";
	$sql .= "\nCamera = '" . GetFormValue('Camera');
	$sql .= "',\nNotes = '" . GetFormValue('Notes');
	$sql .= "',\nMission = '" . GetFormValue('Mission');
	$sql .= "',\nApplication = '" . GetFormValue('Application');
	$sql .= "',\nStyle = '" . GetFormValue('Style');
	$sql .= "',\nAnnotators = '" . GetFormValue('Annotators') . "'";
	$sql .= ",\nDateAnnotated = '" . GetFormValue('DateAnnotated') . "'" if GetFormValue('DateAnnotated');
	$sql .= ",\nNumSDTapes = " . GetFormValue('NumSDTapes') if GetFormValue('NumSDTapes') || GetFormValue('NumSDTapes') eq '0';
	$sql .= ",\nNumHDTapes = " . GetFormValue('NumHDTapes') if GetFormValue('NumHDTapes') || GetFormValue('NumHDTapes') eq '0';
	$sql .= ",\nNumFileSegments = " . GetFormValue('NumFileSegments') if GetFormValue('NumFileSegments') || GetFormValue('NumFileSegments') eq '0';
	$sql .= ",\nHoursOfVideo = " . GetFormValue('HoursOfVideo') if GetFormValue('HoursOfVideo') || GetFormValue('HoursOfVideo') eq '0';
	$sql .= ",\nHoursAnnotated = " . GetFormValue('HoursAnnotated') if GetFormValue('HoursAnnotated') || GetFormValue('HoursAnnotated') eq '0';
	
	$sql .= "\nWHERE ( RovName = '$rovName' AND DiveNumber = $diveNumber)";
	
	
	Win32::ASP::DebugPrint("\nSQL = $sql ");
	
	$RS = $Conn->Execute($sql);
	
	$Errors = $Conn->Errors();
	if ( keys %$Errors ) { 
		$Response->Write("Database Load Error(s): ");
		foreach $error (keys %$Errors)
	        {
	      		$Response->Write($error->{Description});
	   	}
	   	$Conn->Close;
	   	return -1;
	}
	
	%>
	<h3>Updated TapeSummary for <%=$rovName%><%=$diveNumber%></h3>
	
	<%
	
	$RS->Close;
	$Conn->Close;
	return 0;

}	# End updateTapeSummary()


#--------------------------------------------------------------------
#

=head3 tapeLabelsResponse()

Return results of TapeSummary query in as a PDF file of tape labels
Note that if any other HTML or debug output happens then the print will not work.

Author: Mike McCann

Date Created: 8/24/2011

=cut

sub tapeLabelsResponse {
	
	( $rdiveList, $tsQuery, $rnotesHash = $_[2], $rmissionHash, $rannotatorsHash, $rdateAnnotatedHash,
		$rnumSDTapesHash, $rnumHDTapesHash, $rnumFileSegmentsHash, $rcameraHash, $rapplicationHash, $rstyleHash,
		$rhoursOfVideoHash, $rhoursAnnotatedHash, $rdiveStartDateHash, $rdiveChfSciHash, $sumHoursOfVideo, $sumHoursAnnotated,
		$rdiveYYYYHash, $rjDayHash) = @_;

	my $_debug = 0; 		# Make sure MIME-type is set to 'text/html' at line 159 to see the print statements in the browser
	
	print "Inside tapeLabelsResponse()\n<br>" if $_debug;
	
	# Count SD and HD tapes - expect one to be non-zero and the other to be zero.
	# Assign the non-zero hash reference to a hash that is used below
	my $hdCountSum = 0;
	foreach my $hdCount ( values %$rnumHDTapesHash) {
		$hdCountSum += $hdCount;
	}
	my $sdCountSum = 0;
	foreach my $sdCount ( values %$rnumSDTapesHash) {
		$sdCountSum += $sdCount;
	}
	if ($hdCountSum) {
		print "We have $hdCountSum HD tapes\n<br>" if $_debug;
		$rnumtapesHash = $rnumHDTapesHash;
		$HDorSD = "HD";
	}
	if ($sdCountSum) {
		print "We have $sdCountSum SD tapes\n<br>" if $_debug;
		$rnumtapesHash = $rnumSDTapesHash;
		$HDorSD = "SD";
	} 
	
	
	##$Response->{Buffer} = 0;		# Trying to permit pdf files > 3 MB to be delivered
	##$Server->{ScriptTimeout} = 300;		# to the browser...

	# Go through the dives and make tapeInfoHashes for all the HD tapes on each dive.
	# No need to query VARS for tape information as we just use the numHDTapes for the dive.
	my @tapeInfoList = ();
	foreach $d ( @$rdiveList ) {
		
		$vehicle = 'Ventana' if $d =~ /^v/i;
		$vehicle = 'Tiburon' if $d =~ /^t/i;
		$vehicle = 'Doc Ricketts' if $d =~ /^d/i;
		
		$shipName = 'Point Lobos' if $d =~ /^v/i && $$rdiveYYYYHash{$d} < 2012;
		$shipName = 'Rachel Carson' if $d =~ /^v/i && $$rdiveYYYYHash{$d} >= 2012;
		$shipName = 'Western Flyer' if $d =~ /^t/i;
		$shipName = 'Western Flyer' if $d =~ /^d/i;
		
		$diveNumber = substr($d,1);
		
		for (my $tapeNum = 1; $tapeNum <= $$rnumtapesHash{$d}; $tapeNum++ ) {
			print "Building tapeInfor hash for d = $d, tapeNum = $tapeNum\n<br>" if $_debug;
			$tapeInfo = { 	'rovName' => $vehicle,
				'diveNumber' => $diveNumber,
				'shipName' => $shipName,
				'diveDate' => $$rdiveStartDateHash{$d},
				'jDay' => $$rdiveYYYYHash{$d} . '-' . $$rjDayHash{$d},
				'ChiefSci' => $$rdiveChfSciHash{$d},
				'mission' => $$rmissionHash{$d},
				'camera' => $$rcameraHash{$d},
				'tapeNumberHD' => $tapeNum,
				'tapeTotalHD' => $$rnumtapesHash{$d},
				'yrCopyRight' => $$rdiveYYYYHash{$d}
				};
				
				push @tapeInfoList, $tapeInfo;
				
		} # End for (my $tapeNum = 1; $tapeNum <= $$rnumTapesHash{$d}; $tapeNum++ ) 
			
	} # End foreach $d ( @$rdiveList ) 
	print "Number of items in tapeInfoList = " . scalar @tapeInfoList . "\n<br>" if $_debug;
			
	# Create PDF file
	
	use constant mm => 25.4 / 72;
	use constant in => 1 / 72;
	use constant pt => 1;

	# Use writable directory on the server
	
	my $user = 'shore/nouser';
	if ( $Request->ServerVariables('LOGON_USER')->{Item} ) {
		$user = $Request->ServerVariables('LOGON_USER')->{Item};
	}
	$user =~ s/.+[\/\\]//;
	
	my $pdfFile = $Server->mappath("\\expd\\log\\fileup\\data\\tapeLabels_$user.pdf");
	
	print "pdfFile = $pdfFile\n<br>"  if $_debug;
	
	my $pdf = PDF::API2->new( -file => $pdfFile );
	my $now = scalar localtime;
	
	$pdf->info(
	        'Author'       => "Mike McCann",
	        'CreationDate' => "$now",
	        'Creator'      => "$0",
	        'Producer'     => "PDF::API2",
	        'Title'        => "Tape labels for MBARI ROV HD dive video",
	        'Keywords'     => "video tape MBARI underwater scientific oceanographic HD ROV dive"
	    );
	
	print "Defining fonts\n<br>" if $_debug;
	my %font = (
	    Helvetica => {
	        Bold   => $pdf->corefont( 'Helvetica-Bold',    -encoding => 'latin1' ),
	        Roman  => $pdf->corefont( 'Helvetica',         -encoding => 'latin1' ),
	        Italic => $pdf->corefont( 'Helvetica-Oblique', -encoding => 'latin1' ),
	        BoldItalic => $pdf->corefont( 'Helvetica-BoldOblique', -encoding => 'latin1' ),
	    },
	);
	
	my $logoFile = $Server->mappath("\\expd\\log\\MBARI_Logo.pnm");
	die("Unable to find image file $logoFile: $!") unless -e $logoFile;

	
	# Go through all the tapes and print two labels to a page
	
	print "Going through tapeInfolist...\n<br>" if $_debug;
	print "Number of items in tapeInfoList = " . scalar @tapeInfoList . "\n<br>" if $_debug;
	print " #tapeInfoList = " . $#tapeInfoList . "\n<br>" if $_debug;
	for (my $i = 0; $i <= $#tapeInfoList; $i++) {
		print "i = $i,  tapeInfoList[i] = " . $tapeInfoList[$i] . "\n<br>" if $_debug;
		if ( $i % 2 == 0 ) {
			print "Making new page with i = $i\n<br>" if $_debug;
			$page = $pdf->page;
			$page->mediabox( 215.9 / mm, 279.4 / mm );			# 8.5 x 11 "
			makeLabel('top', $pdf, $page, \%font, $logoFile, $tapeInfoList[$i], $HDorSD);
		}
		else {
			print "Making lable with i = $i\n<br>" if $_debug;
			makeLabel('bot', $pdf, $page, \%font, $logoFile, $tapeInfoList[$i], $HDorSD);
		}
	}
	print "Saving pdf\n<br>" if $_debug;
	
	$pdf->saveas;
	$pdf->end();
	
	my $fileSize = -s $pdfFile;
	print "fileSize = $fileSize\n<br>" if $_debug;
	
	my $buf;
	open ( FILE, $pdfFile );
	binmode(FILE);
	read ( FILE, $buf, $fileSize );		
	close FILE;
	
	# This converts to a variant unicode value
	my $variant = Win32::OLE::Variant->new( VT_UI1, $buf );
	$Response->BinaryWrite($variant);

	
} # End tapeLabelsResponse()



#--------------------------------------------------------------------
#

=head3 makeLabel()

print a label on the page.  Assumes that the pdf page has been created 

Author: Mike McCann

Date Created: 8/8/2011

=cut


sub makeLabel {
	
	my ( $topOrBotFlag, $pdf, $page, $rfont, $logoFile, $rtapeHash, $HDorSD ) = @_;
	
	my $_debug = 0; 		# Make sure MIME-type is set to 'text/html' at line 157 to see the print statements in the browser
	
	print "topOrBotFlag = $topOrBotFlag\n<br>" if $_debug;
	
	my %font = %$rfont;
	

	my $logo = $page->gfx;
	my $logoObject = $pdf->image_pnm($logoFile);
	my $logoSpine = $page->gfx;
	my $titleText = $page->text;
	
	my ($x, $y);
	
	if ($_debug) {
		print "rovName = -" . $$rtapeHash{'rovName'} . "-<br>\n";
		print "shipName = -" . $$rtapeHash{'shipName'} . "-<br>\n";
	}
	if ( $topOrBotFlag eq 'bot' ) {
		$logo->image( $logoObject, 80 / mm, 112 / mm, 38 / mm, 23.75 / mm );	# MBARI_Logo has 1.6 aspect ratio (376 x 235 pixels)
		$logoSpine->image( $logoObject, 39 / mm, 122 / mm, 23 / mm, 14.375 / mm );	
		if ( $$rtapeHash{'rovName'} eq 'Doc Ricketts' && $$rtapeHash{'shipName'} eq 'Western Flyer' ) {
			$titleText->translate( 124/mm, 128/mm );
			print "Translating title for Doc Ricketts<br>\n" if $_debug;
		}
		elsif ( $$rtapeHash{'rovName'} eq 'Ventana' && $$rtapeHash{'shipName'} eq 'Point Lobos' ) {
			$titleText->translate( 131/mm, 128/mm );
			print "Translating title for Ventana<br>\n" if $_debug;
		}
		elsif ( $$rtapeHash{'rovName'} eq 'Ventana' && $$rtapeHash{'shipName'} eq 'Rachel Carson' ) {
			$titleText->translate( 131/mm, 128/mm );
			print "Translating title for Ventana<br>\n" if $_debug;
		}
		elsif ( $$rtapeHash{'rovName'} eq 'Tiburon' && $$rtapeHash{'shipName'} eq 'Western Flyer' ) {
			$titleText->translate( 131/mm, 128/mm );
			print "Translating title for Tiburon<br>\n" if $_debug;
		}
		elsif ( $$rtapeHash{'rovName'} eq 'MiniROV' && $$rtapeHash{'shipName'} eq 'Western Flyer' ) {
			$titleText->translate( 124/mm, 128/mm );
			print "Translating title for MiniROV<br>\n" if $_debug;
		}
		elsif ( $$rtapeHash{'rovName'} eq 'MiniROV' && $$rtapeHash{'shipName'} eq 'Rachel Carson' ) {
			$titleText->translate( 124/mm, 128/mm );
			print "Translating title for MiniROV<br>\n" if $_debug;
		}
		##$page->cropbox( 13 / mm, 43 / mm, 177 / mm, 139 / mm );	
		# Lower label
	}
	else {
		$logo->image( $logoObject, 80 / mm, 207 / mm, 38 / mm, 23.75 / mm );	# MBARI_Logo has 1.6 aspect ratio (376 x 235 pixels)
		$logoSpine->image( $logoObject, 39 / mm, 217 / mm, 23 / mm, 14.375 / mm );	
		if ( $$rtapeHash{'rovName'} eq 'Doc Ricketts' && $$rtapeHash{'shipName'} eq 'Western Flyer' ) {
			$titleText->translate( 124/mm, 223/mm );
			print "Translating title for Doc Ricketts<br>\n" if $_debug;
		}
		elsif ( $$rtapeHash{'rovName'} eq 'Ventana' && $$rtapeHash{'shipName'} eq 'Point Lobos' ) {
			$titleText->translate( 127/mm, 223/mm );
			print "Translating title for Ventana<br>\n" if $_debug;
		}
		elsif ( $$rtapeHash{'rovName'} eq 'Ventana' && $$rtapeHash{'shipName'} eq 'Rachel Carson' ) {
			$titleText->translate( 127/mm, 223/mm );
			print "Translating title for Ventana<br>\n" if $_debug;
		}
		elsif ( $$rtapeHash{'rovName'} eq 'Tiburon' && $$rtapeHash{'shipName'} eq 'Western Flyer' ) {
			$titleText->translate( 127/mm, 223/mm );
			print "Translating title for Tiburon<br>\n" if $_debug;
		}
		elsif ( $$rtapeHash{'rovName'} eq 'MiniROV' && $$rtapeHash{'shipName'} eq 'Western Flyer' ) {
			$titleText->translate( 127/mm, 223/mm );
			print "Translating title for MiniROV<br>\n" if $_debug;
		}
		elsif ( $$rtapeHash{'rovName'} eq 'MiniROV' && $$rtapeHash{'shipName'} eq 'Rachel Carson' ) {
			$titleText->translate( 127/mm, 223/mm );
			print "Translating title for MiniROV<br>\n" if $_debug;
		}
		##$page->cropbox( 13 / mm, 139 / mm, 177 / mm, 235 / mm );		# Upper label
	}
	print "Setting fillcolor\n<br>" if $_debug;
	
	$titleText->fillcolor('black');
	print "Setting font: $font{'Helvetica'}{'Bold'}\n<br>" if $_debug;
	$titleText->font( $font{'Helvetica'}{'Bold'}, 16 / pt );
	$titleText->text('ROV ');
	$titleText->font( $font{'Helvetica'}{'BoldItalic'}, 16 / pt );
	$titleText->text($$rtapeHash{'rovName'});
	$titleText->font( $font{'Helvetica'}{'Bold'}, 16 / pt );
	$titleText->text(' Video');
	
	if ( $$rtapeHash{'rovName'} eq 'Doc Ricketts' && $$rtapeHash{'shipName'} eq 'Western Flyer' ) {
		$titleText->distance( 11/mm, -6/mm );
	}
	elsif ( $$rtapeHash{'rovName'} eq 'Ventana' && $$rtapeHash{'shipName'} eq 'Point Lobos' ) {
		$titleText->distance( 7/mm, -6/mm );
	}
	elsif ( $$rtapeHash{'rovName'} eq 'Ventana' && $$rtapeHash{'shipName'} eq 'Rachel Carson' ) {
		$titleText->distance( 5/mm, -6/mm );
	}
	elsif ( $$rtapeHash{'rovName'} eq 'Tiburon' && $$rtapeHash{'shipName'} eq 'Western Flyer' ) {
		$titleText->distance( 6/mm, -6/mm );
	}
	elsif ( $$rtapeHash{'rovName'} eq 'MiniROV' && $$rtapeHash{'shipName'} eq 'Western Flyer' ) {
		$titleText->distance( 7/mm, -6/mm );
	}
	elsif ( $$rtapeHash{'rovName'} eq 'MiniROV' && $$rtapeHash{'shipName'} eq 'Rachel Carson' ) {
		$titleText->distance( 5/mm, -6/mm );
	}
	$titleText->font( $font{'Helvetica'}{'Bold'}, 14 / pt );
	$titleText->text('R/V ');
	$titleText->font( $font{'Helvetica'}{'BoldItalic'}, 14 / pt );
	$titleText->text($$rtapeHash{'shipName'});
	
	my $frontText = $page->text;
	$frontText->font( $font{'Helvetica'}{'Roman'}, 14 / pt );
	
	my $lineSpacing = 12;
	
	if ( $topOrBotFlag eq 'bot' ) {		
		$x = 80;
		$y = 104;
	} 
	else {
		$x = 80;
		$y = 199;
	}
	
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Date: $$rtapeHash{'diveDate'}");
	
	$y -= $lineSpacing;
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Year-Day: $$rtapeHash{'jDay'}");
	
	$y -= $lineSpacing;
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Chief Scientist: $$rtapeHash{'ChiefSci'}");
	
	$y -= $lineSpacing;
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Mission: $$rtapeHash{'mission'}");
	
	$y -= $lineSpacing;
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Camera: $$rtapeHash{'camera'}");
	
	if ( $topOrBotFlag eq 'bot' ) {	
		$x = 147;
		$y = 104;
	}
	else {
		$x = 147;
		$y = 199;
	}

	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Tape # $$rtapeHash{'tapeNumberHD'} $HDorSD of $$rtapeHash{'tapeTotalHD'} $HDorSD");
	
	$y -= $lineSpacing;
	print "Printing label for " . uc(substr($$rtapeHash{'rovName'},0,1)) . $$rtapeHash{'diveNumber'} . "\n" if $_debug;
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Dive Number: " . uc(substr($$rtapeHash{'rovName'},0,1)) . $$rtapeHash{'diveNumber'} );
	
	if ( $topOrBotFlag eq 'bot' ) {	
		$x = 195;
		$y = 50;
	}
	else {
		$x = 195;
		$y = 145;
	}
	$frontText->translate( $x/mm, $y/mm );
	$frontText->font( $font{'Helvetica'}{'Roman'}, 12 / pt );
	$frontText->text_right("© MBARI $$rtapeHash{'yrCopyRight'}");
	
	
	my $spineText = $page->text;
	$spineText->font( $font{'Helvetica'}{'Roman'}, 12 / pt );
	$lineSpacing = 7;
	
	if ( $topOrBotFlag eq 'bot' ) {	
		$x = 61;
		$y = 121;
	}
	else {
		$x = 61;
		$y = 216;
	}
	
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text("$$rtapeHash{'ChiefSci'}");
	
	$x -= $lineSpacing;
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	if ( length($$rtapeHash{'mission'}) > 32 ) {
		$shrtMission = substr($$rtapeHash{'mission'}, 0, 32) . '...';
	}
	else {
		$shrtMission = $$rtapeHash{'mission'};
	}
	$spineText->text("$shrtMission");
	
	$x -= $lineSpacing;
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text("$$rtapeHash{'diveDate'}");
	
	$x -= $lineSpacing;
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text("$$rtapeHash{'jDay'}");
	
	if ( $topOrBotFlag eq 'bot' ) {	
		$x = 61;
		$y = 60;
	}
	else {
		$x = 61;
		$y = 155;
	}
	
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text_right("Dive " . uc(substr($$rtapeHash{'rovName'},0,1)) . $$rtapeHash{'diveNumber'} );
	
	$x -= $lineSpacing * 2;
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text_right("Tape # $$rtapeHash{'tapeNumberHD'}");
	
	$x -= $lineSpacing;
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text_right("© MBARI $yrCopyRight");

} # End makeLabel()


%>

