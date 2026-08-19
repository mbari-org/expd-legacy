<%@ LANGUAGE = PerlScript %>

<!--#include file="dive_hdr_karen.inc"-->
<!--#include file="expd_functions_karen.inc"-->

<% 
=head1 NAME

dive_karen.asp - Dive edit/update/delete Active Server Page

=head1 SYNOPSIS

    > http://expd.mbari.org/expd/log/dive_karen.asp

=head1 DESCRIPTION

Pages for inserting, updating or deleting dives that belong to expeditions.

=head1 FUNCTIONS

=cut


# Don't use CGI, it freezes IIS when a new object is created.
use Win32::ASP;
use OLE;
use Time::Local;
use Net::SMTP;
use POSIX;
use LWP::Simple;

$this_script = 'dive_karen.asp';

init_AppVars();

$delimiter = '<font face="helvetica,Ariel"> <b>&#183;</b> </font>';

#
# Some font strings
#
$font_2 = '<font face="Helvetica,Ariel" size="-2">';
$font_1 = '<font face="Helvetica,Ariel" size="-1">';
$font_err = '<font face="Helvetica,Ariel" size="-1" color="red">';

$dsn = $Application->{'DSN'};

$Session->{'logistics_email'} = GetFormValue('logistics_email') || 
								$Session->{'logistics_email'} || $Application->{'logistics_email'};							
					
$Session->{'RegisteredUser'} = GetFormValue('RegisteredUser') || 
								$Session->{'RegisteredUser'} || $Application->{'RegisteredUser'};
$Session->{'RegisteredUser'} = 0 unless $Session->{'RegisteredUser'} == 1;
##$Session->{'RegisteredUser'} = 0;
Win32::ASP::DebugPrint("\nRegisteredUser = " . $Session->{'RegisteredUser'} . "\n");

#
# Used for both editing a dive and entering a new one (different forms, same logic)
#
$diveCheckDateFunctions = "
function calcDate(form) {	// From http://developer.netscape.com/viewsource/goodman_dateobject.html

	var epoch = false;
	var localdtg = false;
	var days = new Array(\"Sun\", \"Mon\", \"Tue\", \"Wed\", \"Thu\", \"Fri\", \"Sat\");
	
	for( var i = 0; i < document.diveForm.length; i++ ) {
		var e = document.diveForm.elements[i];
		if (e.type == \"radio\" ) {
			//alert (\"e.name = \" + e.name);
			if ( e.name == \"Stime\" ) {
				if ( e.checked && e.value == \"epoch\" ) {
					epoch = true;
					//alert(\"epoch = true\");
				}
				if ( e.checked && e.value == \"localdtg\" ) {
					localdtg = true;
					//alert(\"localdtg = true\");
				}
			}
		}
	}

	if ( localdtg ) { 
		// Start date
		var sm = form.DiveStartDtgMonth.selectedIndex
		var sd = form.DiveStartDtgDay.selectedIndex + 1	
		var sy = form.DiveStartDtgYear.options[form.DiveStartDtgYear.selectedIndex].value	
		var sh = form.DiveStartDtgHr.options[form.DiveStartDtgHr.selectedIndex].value	
		var sn = form.DiveStartDtgMin.options[form.DiveStartDtgMin.selectedIndex].value	
		var ss = 0
		//alert(\"localdtg: sd = \" + sd + \"  form.DiveStartDtgDay.selectedIndex = \" + form.DiveStartDtgDay.selectedIndex)
		//alert(\"localdtg: sd = \" + sh + \"  form.DiveStartDtgHr.selectedIndex = \" + form.DiveStartDtgHr.selectedIndex)
		var sDate = new Date(sy,sm,sd,sh,sn,ss)	
		form.SEpochSecs.value = sDate.getTime() / 1000
		form.SDayOfWeek.value = days[sDate.getDay()]
		form.SGMTString.value = sDate.toGMTString()
		form.SYearDay.value = getYearDay(form.SEpochSecs.value)
		
		// End date
		var em = form.DiveEndDtgMonth.selectedIndex
		var ed = form.DiveEndDtgDay.selectedIndex + 1	
		var ey = form.DiveEndDtgYear.options[form.DiveEndDtgYear.selectedIndex].value
		var eh = form.DiveEndDtgHr.options[form.DiveEndDtgHr.selectedIndex].value
		var en = form.DiveEndDtgMin.options[form.DiveEndDtgMin.selectedIndex].value	
		var es = 0
		var eDate = new Date(ey,em,ed,eh,en,es)	
		form.EEpochSecs.value = eDate.getTime() / 1000
		form.EDayOfWeek.value = days[eDate.getDay()]
		form.EGMTString.value = eDate.toGMTString()
		form.EYearDay.value = getYearDay(form.EEpochSecs.value)
	}
	if ( epoch ) {
		// Start date
		var sDate = new Date(form.SEpochSecs.value * 1000)
		form.SDayOfWeek.value = days[sDate.getDay()]
		var sYearIndex = sDate.getYear() - 1988		// 1988 is when pick list starts - beg of MBARI dives
		if ( sYearIndex < 0 ) {				// Wierd, got to add 1900 back for Netscape 4.7
			sYearIndex = sYearIndex + 1900
		}
		if ( sYearIndex < 0 ) {
			alert(\"There are no MBARI dives before 1988.\")
			form.SEpochSecs.focus()
			return (false)
		}	
		form.DiveStartDtgYear.options[sYearIndex].selected = true
		form.DiveStartDtgMonth.options[sDate.getMonth()].selected = true
		form.DiveStartDtgDay.options[sDate.getDate() - 1].selected = true	
		form.DiveStartDtgHr.options[sDate.getHours()].selected = true
		form.DiveStartDtgMin.options[sDate.getMinutes()].selected = true
		form.SGMTString.value = sDate.toGMTString()
		form.SYearDay.value = getYearDay(form.SEpochSecs.value)
		
		// End date
		var eDate = new Date(form.EEpochSecs.value * 1000)
		form.EDayOfWeek.value = days[eDate.getDay()]
		var eYearIndex = eDate.getYear() - 1988		// 1988 is when pick list starts - beg of MBARI dives
		if ( eYearIndex < 0 ) {				// Wierd, got to add 1900 back for Netscape 4.7
			eYearIndex = eYearIndex + 1900
		}
		if ( eYearIndex < 0 ) {
			alert(\"There are no MBARI dives before 1988.\")
			form.EEpochSecs.focus()
			return (false)
		}	
		form.DiveEndDtgYear.options[eYearIndex].selected = true
		form.DiveEndDtgMonth.options[eDate.getMonth()].selected = true
		form.DiveEndDtgDay.options[eDate.getDate() - 1].selected = true	
		form.DiveEndDtgHr.options[eDate.getHours()].selected = true
		form.DiveEndDtgMin.options[eDate.getMinutes()].selected = true
		form.EGMTString.value = eDate.toGMTString()
		form.EYearDay.value = getYearDay(form.EEpochSecs.value)
	}
	
	if ( !form.SEpochSecs ) {}
    else {
    	// Caluculate dive duration
		form.DiveTimeHours.value = (form.EEpochSecs.value - form.SEpochSecs.value) / 3600
    }


	
	// Set image size timeline depending on start & end times
	timeSpan = form.ExpdEndEsecs.value - form.ExpdStartEsecs.value
	maxWidth = 600
	//alert (\"ExpdStartEsecs = \" + form.ExpdStartEsecs.value + \" maxWidth = \" + maxWidth)
	//alert (\"timeSpan = \" + timeSpan + \" form.SEpochSecs.value = \" +  form.SEpochSecs.value)

	if ( !form.ExpdStartEsecs || !timeSpan ) {}
	else {
		form.StartDiveTimeLine.width = maxWidth * (form.SEpochSecs.value - form.ExpdStartEsecs.value) / timeSpan 
		form.EndDiveTimeLine.width = maxWidth * (form.EEpochSecs.value - form.ExpdStartEsecs.value) / timeSpan - form.StartDiveTimeLine.width
		form.StartDiveLabel.width = maxWidth * (form.SEpochSecs.value - form.ExpdStartEsecs.value) / timeSpan 
    }
	//alert (\"form.StartDiveTimeLine.width = \" + form.StartDiveTimeLine.width + \" form.EndDiveTimeLine.width = \" + form.EndDiveTimeLine.width)
	//alert(\"calcDate: Update Button Value = \" + form.update_dive.value)
	return (true)
	
} // End calcDate


function checkEnter(e){ 	//e is event object passed from function invocation
	var characterCode  	//literal character code will be stored in this variable

	if(e && e.which){ 	//if which property of event object is supported (NN4)
	 	e = e
	 	characterCode = e.which 	//character code is contained in NN4's which property
	}
	else{							
	 	e = event						
	 	characterCode = e.keyCode 	//character code is contained in IE's keyCode property
	}
	//alert(\"characterCode = \" + characterCode)
	
	if(characterCode == 13){ 	//if generated character code is equal to ascii 13 (if enter key)
	 	//document.forms[0].submit() 	//submit the form
	 	calcDate(document.forms[0])
	 	return false 
	}
	else{
	 	return true 
	}
} // End checkEnter 

function pad(number,length) {
	var str = '' + number;
	while (str.length < length)
	str = '0' + str;
	return str;
} // End pad

function LeapYear(year) {
	if ((year/4)   != Math.floor(year/4))   return false;
	if ((year/100) != Math.floor(year/100)) return true;
	if ((year/400) != Math.floor(year/400)) return false;
	return true;
} // End LeapYear

function getJulianDay(year,month,day) {
	var accumulate    = new Array(0, 31, 59, 90,120,151,181,212,243,273,304,334);
	var accumulateLY  = new Array(0, 31, 60, 91,121,152,182,213,244,274,305,335);
	if (LeapYear(year))
		return pad(day + accumulateLY[month],3);
	else
		return pad(day + accumulate[month],3);
} // End getJulianDay

function getYearDay(esecs) {		// getYearDay returns GMT year-day based on epoch seconds -  arghhhh!
	var d = new Date(esecs * 1000);
	
	var reDaMoYe = /[a-zA-Z,\\s]+(\\d+)\\s+([a-zA-Z]+)\\s+(\\d\\d\\d\\d).+/;
	var reTime = /\\s+\\d\\d:\\d\\d:\\d\\d\\s+GMT/;
	var reDate = /[a-zA-Z,\\s]+\\d+[a-zA-Z,\\s]+\\d\\d\\d\\d\\s/;
	var foo = d.toGMTString();
	var bar = foo.replace(reDate,\"\");
	var reGMT = /\\s+GMT/;
	
	strs = foo.match(reDaMoYe);
	//alert(\"Day =\" + strs[1] + \"Month =\" + strs[2] + \"Year =\" + strs[3]);
	if (strs[2] == \"Jan\") m = 0;
	if (strs[2] == \"Feb\") m = 1;
	if (strs[2] == \"Mar\") m = 2;
	if (strs[2] == \"Apr\") m = 3;
	if (strs[2] == \"May\") m = 4;
	if (strs[2] == \"Jun\") m = 5;
	if (strs[2] == \"Jul\") m = 6;
	if (strs[2] == \"Aug\") m = 7;
	if (strs[2] == \"Sep\") m = 8;
	if (strs[2] == \"Oct\") m = 9;
	if (strs[2] == \"Nov\") m = 10;
	if (strs[2] == \"Dec\") m = 11;
	var y=strs[3] - 0;		 // Convert from string to number
	var d=strs[1] - 0;
	return getJulianDay(y,m,d);

} // End getYearDay
";

%>

<%
if ( GetFormValue('RovName') && GetFormValue('DiveNumber') && GetFormValue('edit') ) {
	$rov = GetFormValue('RovName');
	if ( $rov =~ /Ventana/i || $rov =~ /vnta/i ) {
		$rovName = 'vnta';
	}
	elsif ( $rov =~ /Tiburon/i || $rov =~ /tibr/i ) {
		$rovName = 'tibr';
	}
	elsif ( $rov =~ /Doc\s*Ricketts/i || $rov =~ /docr/i ) {
		$rovName = 'docr';
	}
	elsif ( $rov =~ /MiniROV/i || $rov =~ /mini/i ) {
		$rovName = 'mini';
	}
	else {
		%>
		<h2>'<%=$rov%>' is not a valid ROVName.  Please choose 'vnta', 'tibr', 'docr' or 'mini'</h2>
		<%
		return;
	}
	%>
	<h2>Edit Dive <%= GetFormValue('RovName')%><%= GetFormValue('DiveNumber')%></h2>
	<font color="#408080"><strong>Modify time or comments for existing dive.
   </strong></font>
	</td></tr></table>
	<% 
	edit();
}
elsif ( GetFormValue('update_dive') ) {
	%>
	<h2>Updating Dive <%= GetFormValue('RovName')%><%= GetFormValue('DiveNumber')%></h2>
	</td></tr></table>
	<% 
	if ( GetFormValue('RovName') =~ /Ventana/i || GetFormValue('RovName') =~ /vnta/i ) {
		$rovName = 'vnta';
	}
	elsif ( GetFormValue('RovName') =~ /Tiburon/i || GetFormValue('RovName') =~ /tibr/i ) {
		$rovName = 'tibr';
	}
	elsif ( GetFormValue('RovName') =~ /Doc\s*Ricketts/i || GetFormValue('RovName') =~ /docr/i ) {
		$rovName = 'docr';
	}
	elsif ( GetFormValue('RovName') =~ /MiniROV/i || GetFormValue('RovName') =~ /mini/i ) {
		$rovName = 'mini';
	}
	if ( $Session->{'RegisteredUser'} ) {
		update();
		edit();
	}
	else {
		email_dive_insert('update');
	}
	
}
elsif ( GetFormValue('delete_dive') ) {
	if ( $Session->{'RegisteredUser'} ) { %>
		<h2>Deleting Dive <%= GetFormValue('RovName')%><%= GetFormValue('DiveNumber')%></h2>
		</td></tr></table> <%
		delete_rec();
	}
	else { %>
		<h2>Must be a registered user to delete dives.</h2>
		</td></tr></table>
	<%}
}
elsif ( GetFormValue('add_dive') ) {
	%>
	<h2>Add a Dive to ExpeditionID <%= GetFormValue('ExpeditionID')%></h2>
	<font color="#408080"><strong>Enter a new dive for this expedition.
    </strong></font>
	</td></tr></table>
	<% 
	if ( GetFormValue('add_dive') =~ /Another/ ) {
		if ( $Session->{'RegisteredUser'} ) {
			insert_dive();
		}
		else {
			email_dive_insert('');
		}
		add_dive_form();
	}
	else {
		add_dive_form();
	}
}
elsif ( GetFormValue('insert') ) {
	%>
	<h2>Inserting dive <%= GetFormValue('RovName')%><%= GetFormValue('DiveNumber')%></h2>
	</td></tr></table>
	<% 
	if ( $Session->{'RegisteredUser'} ) {
			insert_dive();
	}
	else {
		email_dive_insert('');
	}
	close_session();
}
else {
	%>
	<h2 align="center">Dive List</h2>
	</td></tr></table>
	<% 
	list_dives();
}
%>


<%
#--------------------------------------------------------------------
#

=head3 edit()

Present the info for the dive indicated by the ID passed.

Mike McCann

Date Created: 7/13/99

=cut

sub edit {

#
# Open DB
#
open_database($dsn);		# Creates $Conn object as a global variable

if ( GetFormValue('DiveNumber') && $rovName ) {
	$sql = "SELECT RovName, DiveNumber, DeviceID, DiveStartDtg, DiveEndDtg, ExpeditionID_FK ";
	$sql .= "FROM Dive WHERE DiveNumber = " . GetFormValue('DiveNumber');
	$sql .= " AND RovName = '" . $rovName . "'";
} else { 
	# If editing to add a new dive
	$sql = "SELECT RovName, DiveNumber, DeviceID, DiveStartDtg, DiveEndDtg, ExpeditionID_FK ";
	$sql .= "FROM Dive WHERE ExpeditionID_FK = " . GetFormValue('ExpeditionID');
}
Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RSt = $Conn->Execute($sql);
	if(!$RSt) {
    	$Errors = $Conn->Errors();
	    print "Errors1:\n";
	    foreach $error (keys %$Errors) {
	   		print $error->{Description}, "\n";
	    }	
	die "<BR>Empty Return Set1 for \n$sql \n\nfrom $dsn";
}

#####################################################################################
# Get the record to edit, need to be explicit to get times as epoch seconds
#####################################################################################
if ( $RSt->Fields('ExpeditionID_FK')->value ne 0 ) {	# If editing to add a new dive
	$sql = "SELECT RovName, DiveNumber, DiveStartDtg, DiveEndDtg, ExpeditionID_FK, Dive.DeviceID, Expedition.ShipName, Expedition.ShipSeqNum, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS StartEpoch, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', DiveEndDtg) AS EndEpoch, ";
	$sql .= "DiveChiefScientist, BriefAccomplishments ";
	$sql .= "FROM Dive \n";
	$sql .= "JOIN Expedition ON Dive.ExpeditionID_FK = Expedition.ExpeditionID \n";
	$sql .= "WHERE DiveNumber = " . GetFormValue('DiveNumber');
	$sql .= " AND RovName = '" . $rovName . "'";
	####print"<br>Get all MBARI ROV (including MiniROV) dives to edit (RS):<br> $sql<br>"; 
}
elsif ( GetFormValue('DiveNumber') && $rovName && $RSt->Fields('ExpeditionID_FK')->value eq 0 ) {
	$sql = "SELECT RovName, DiveNumber, ExpeditionID_FK, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS StartEpoch, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', DiveEndDtg) AS EndEpoch, ";
	$sql .= "DiveChiefScientist, BriefAccomplishments ";
	$sql .= "FROM Dive \n";
	$sql .= "JOIN Expedition ON Dive.ExpeditionID_FK = Expedition.ExpeditionID \n";
	$sql .= "WHERE DiveNumber = " . GetFormValue('DiveNumber');
	$sql .= " AND RovName = '" . $rovName . "'";
	####print"<br>Get all NON-MBARI dives to edit (RS): $sql<br>";
}
Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RS = $Conn->Execute($sql);
	if(!$RS) {
    	$Errors = $Conn->Errors();
	    print "Errors2:\n";
	    foreach $error (keys %$Errors) {
	   		print $error->{Description}, "\n";
	    }	
	die "<BR>Empty Return Set2 for \n$sql \n\nfrom $dsn";
}


######################################################################################################
# Get pick list for Dive Chief Scientist - show all in Person table (Cannot remove Chief Scientists!)
######################################################################################################
$sql="SELECT * FROM Person ORDER BY LastName, FirstName";
Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RScs = $Conn->Execute($sql);
if(!$RScs) {
 	die "<BR>Empty Return Set2: RScs. ";
}


######################################################################################################
# Do an explicit select on EXPEDITION to get posssible start & end times for dive 
######################################################################################################
	my $ExpdID = $RSt->Fields('ExpeditionID_FK')->value if $RSt;
	$ExpdID = GetFormValue('ExpeditionID') unless $ExpdID;
	$dn = $RSt->Fields('DiveNumber')->value || GetFormValue('DiveNumber');
	$rov = $RS->Fields('RovName')->value;
if ($RSt->Fields('ExpeditionID_FK')->value eq 0 && $RSt->Fields('RovName')->value eq 'mini') {
	$sql = "SELECT " . join( ", ", @{$Application->{'DiveFields'}} ) . ", 'cmpl' as StatCode, \n";
	$sql .= "DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS StartEpoch, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', DiveEndDtg) AS EndEpoch ";
	$sql .= "FROM Dive \n";
	$sql .= "WHERE ";
	$sql .= "ExpeditionID_FK = " . $ExpdID . " AND DeviceID = 1859 AND DiveNumber = " . $dn . " AND RovName = '" . $rovName . "'";
	####print"<br>SQL1 Start/End Times (RSexpd): <br>$sql<br>";
}
elsif ($RSt->Fields('ExpeditionID_FK')->value ne 0 && $RSt->Fields('RovName')->value eq 'mini') {
	$ExpdID = $RS->Fields('ExpeditionID_FK')->value;
	$diven = $RS->Fields('DiveNumber')->value;
	$sql = "SELECT " . join( ", ", @{$Application->{'DiveFields'}} ) . ", \n";
	$sql .= "DateDiff(\"ss\", '01/01/70', Dive.DiveStartDtg) AS StartEpoch, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', Dive.DiveEndDtg) AS EndEpoch, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', Expedition.ScheduledStartDtg) AS ScheduledStartEpoch, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', Expedition.ScheduledEndDtg) AS ScheduledEndEpoch ";
	$sql .= " FROM Dive \n";
	$sql .= "JOIN Expedition ON Dive.ExpeditionID_FK = Expedition.ExpeditionID \n";
	$sql .= "WHERE ";
	$sql .= "Dive.DiveNumber = " . $dn . " AND Dive.ExpeditionID_FK = " . $ExpdID . " AND Dive.RovName = '" . $rov . "' AND Expedition.ShipName = '" . $RS->Fields('ShipName')->Value . "' AND Expedition.ShipSeqNum = " . $RS->Fields('ShipSeqNum')->Value;
	####print"<br>SQL2 Start/End Times (RSexpd): <br>$sql<br>";
}
if ($RSt->Fields('ExpeditionID_FK')->value ne 0 && $RSt->Fields('RovName')->value ne 'mini') {
	$ExpdID = $RS->Fields('ExpeditionID_FK')->value;
	$ExpdID = GetFormValue('ExpeditionID') unless $ExpdID;
	$sql = "SELECT " . join( ", ", @{$Application->{'ExpeditionFields'}} ) . " FROM Expedition\n";
	$sql .= "WHERE ";
	$sql .= "ExpeditionID = " . $ExpdID;
	####print"<br>SQL3 Start/End Times (RSexpd): <br>$sql<br>";
}

Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: \n$sql");
$RSexpd = $Conn->Execute($sql);

if(!$RSexpd) {
 	die "<BR>Empty Return EXPD Record Set3 for: \n<br>$sql ";
}

######################################################################################################
# Prime start & end times for dive base on 1) Dive records, 2) Expedition times, 3) Scheduled times
######################################################################################################

if ($RSt->Fields('ExpeditionID_FK')->value eq 0 && $RSt->Fields('RovName')->value eq 'mini') {
	$ExpdID = $RSt->Fields('ExpeditionID_FK')->value;  ### Probably don't need this line....
	####print"<br>ExpdID:  $ExpdID<br>";
	$StartEpoch = $RSexpd->Fields('StartEpoch')->value; 
	Win32::ASP::DebugPrint("\nStartEpoch = $StartEpoch");
	Win32::ASP::DebugPrint("\nStartEpoch =". $StartEpoch );

	$EndEpoch = $RSexpd->Fields('EndEpoch')->value;
	Win32::ASP::DebugPrint("\nEndEpoch = $EndEpoch");

	####print"<br>(NON-MBARI Ship MiniROV):<br>Dive: StartEpoch: $StartEpoch <br>Dive: EndEpoch: $EndEpoch<br>";
} 
elsif ($RS->Fields('ExpeditionID_FK')->value ne 0 && $RS->Fields('RovName')->value eq 'mini') {
	$ExpdID = $RSt->Fields('ExpeditionID_FK')->value;
	$StartEpoch = $RSexpd->Fields('StartEpoch')->value; 
	$ExpdStartEsecs = $RSexpd->Fields('ScheduledStartEpoch')->value;
	Win32::ASP::DebugPrint("\nStartEpoch = $StartEpoch");
	Win32::ASP::DebugPrint("\nStartEpoch =". $StartEpoch );

	$EndEpoch = $RSexpd->Fields('EndEpoch')->value;
	$ExpdEndEsecs = $RSexpd->Fields('ScheduledEndEpoch')->value;
	Win32::ASP::DebugPrint("\nEndEpoch = $EndEpoch");

	####print"<br>(MBARI SHIP MiniROV):<br>Dive Start: StartEpoch: $StartEpoch, ExpdStartEsecs: $ExpdStartEsecs, <br>Dive End: EndEpoch: $EndEpoch, ExpdEndEsecs: $ExpdEndEsecs<br>";
}
elsif ($RS->Fields('ExpeditionID_FK')->value ne 0 && $RS->Fields('RovName')->value ne 'mini') {
	####print"<br>ExpdID:  $ExpdID<br>";
	$StartEpoch = $RS->Fields('StartEpoch')->value if $RS; 
	$StartEpoch = $StartEpoch || SelectEpoch("StartDtg", $ExpdID) || SelectEpoch("ScheduledStartDtg", $ExpdID);
	$ExpdStartEsecs = SelectEpoch("StartDtg", $ExpdID) || SelectEpoch("ScheduledStartDtg", $ExpdID);
	Win32::ASP::DebugPrint("\nStartEpoch = $StartEpoch");
	Win32::ASP::DebugPrint("\nStartEpoch =". $StartEpoch );

	$EndEpoch = $RS->Fields('EndEpoch')->value if $RS;
	$EndEpoch =	$EndEpoch || SelectEpoch("EndDtg", $ExpdID) || SelectEpoch("ScheduledEndDtg", $ExpdID);
	$ExpdEndEsecs = SelectEpoch("EndDtg", $ExpdID) || SelectEpoch("ScheduledEndDtg", $ExpdID);
	Win32::ASP::DebugPrint("\nEndEpoch = $EndEpoch");

	####print"<br>(MBARI ROV):<br>Expd Start: <br>StartEpoch: $StartEpoch, ExpdStartEsecs: $ExpdStartEsecs, <br>EndEpoch: $EndEpoch, ExpdEndEsecs: $ExpdEndEsecs<br>";
}

##############################################
# Get all the dives from this expedition
##############################################

if ($RSt->Fields('ExpeditionID_FK')->value eq 0 ) {
	my $ExpdID = $RSt->Fields('ExpeditionID_FK')->value;
	$sql = "SELECT RovName, DeviceID, DiveNumber, DiveStartDtg, DiveEndDtg, ExpeditionID_FK, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS StartEpoch, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', DiveEndDtg) AS EndEpoch ";
	$sql .= " FROM Dive\n";
	$sql .= " WHERE ";
	$sql .= " DeviceID = 1859 AND ExpeditionID_FK = " . $ExpdID;
	$sql .= " ORDER BY DiveNumber";
	####print"<br>EXPDID1: $ExpdID<br>";
	####print"<br>EXPDID=0: $sql<br>";
}
elsif ($RSt->Fields('ExpeditionID_FK')->value ne 0 ) {
	my $ExpdID = $RS->Fields('ExpeditionID_FK')->value if $RS;
	$ExpdID = GetFormValue('ExpeditionID') unless $ExpdID;
	$sql = "SELECT RovName, Dive.DeviceID, DiveNumber, ExpeditionID_FK, DiveStartDtg, DiveEndDtg, DiveChiefScientist, BriefAccomplishments, ";
	$sql .= "Expedition.ShipName, Expedition.ScheduledStartDtg, Expedition.ScheduledEndDtg, Expedition.StartDtg, Expedition.EndDtg, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', Dive.DiveStartDtg) AS StartEpoch, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', Dive.DiveEndDtg) AS EndEpoch, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', Expedition.ScheduledStartDtg) AS ScheduledStartEpoch, ";
	$sql .= "DateDiff(\"ss\", '01/01/70', Expedition.ScheduledEndDtg) AS ScheduledEndEpoch ";	
	$sql .= " FROM Dive\n";
	$sql .= " JOIN Expedition ON Dive.ExpeditionID_FK = Expedition.ExpeditionID";
	$sql .= " WHERE ";
	$sql .= " Dive.ExpeditionID_FK = " . $ExpdID;
	$sql .= " ORDER BY DiveNumber";
####print"<br>EXPDID2: $ExpdID<br>";
}
Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: \n$sql");
$RSdives = $Conn->Execute($sql);
if(!$RSdives) {
 	die "<BR>Empty Return Set4: RSdives. for <br>$sql ";
}
$expdDiveList = '';
while ( !$RSdives->EOF ) { 
	my $diveName = $RSdives->Fields('RovName')->value . $RSdives->Fields('DiveNumber')->value;
	$DiveStartEsecs = $RS->Fields('StartEpoch')->value;
	$DiveEndEsecs = $RS->Fields('EndEpoch')->value;
###################

	if ( $RSt->Fields('ExpeditionID_FK')->value eq 0 ) {
		if ( $RSdives->Fields('DiveNumber')->value == GetFormValue('DiveNumber') ) {
			$sDtl{$diveName} = (600 * ($RSexpd->Fields('StartEpoch')->value ));
			$Dtl{$diveName} = (600 * ($RSexpd->Fields('EndEpoch')->value ));
			$eDtl{$diveName} = $Dtl{$diveName} - $sDtl{$diveName};
			####print"<br>NON-MBARI MiniROV Dive ($diveName) Start and End Times: $diveName: $sDtl{$diveName}, $eDtl{$diveName}<br>";
		}
	}
	elsif ( $RS->Fields('ExpeditionID_FK')->value ne 0 ) {
		if ( $RSdives->Fields('RovName')->value eq 'mini' ) {
			$sDtl{$diveName} = (600 * (($RSdives->Fields('StartEpoch')->value - $RSdives->Fields('ScheduledStartEpoch')->value) / ( $RSdives->Fields('ScheduledEndEpoch')->value - $RSdives->Fields('ScheduledStartEpoch')->value)));
			$eDtl{$diveName} = (600 * (($RSdives->Fields('EndEpoch')->value - $RSdives->Fields('ScheduledStartEpoch')->value) / ( $RSdives->Fields('ScheduledEndEpoch')->value - $RSdives->Fields('ScheduledStartEpoch')->value))) - $sDtl{$diveName};
			####print"<br>MBARI MiniROV $diveName Start and End Times: $sDtl{$diveName}, $eDtl{$diveName}<br>";
		}
		if ( $RSdives->Fields('RovName')->value ne 'mini' ) {
			$sDtl{$diveName} = 600 * ($RSdives->Fields('StartEpoch')->value - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs);
			$eDtl{$diveName} = 600 * ($RSdives->Fields('EndEpoch')->value - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $sDtl{$diveName};
			####print"<br>MBARI $diveName Start and End Times: $sDtl{$diveName}, $eDtl{$diveName}<br>";
		}
	}

	if ( $RSdives->Fields('DiveNumber')->value != GetFormValue('DiveNumber') ) {
		$expdDiveList .= "<a href=\"$this_script?RovName=" . $RSdives->Fields('RovName')->value;
		$expdDiveList .= "&DiveNumber=" . $RSdives->Fields('DiveNumber')->value . "&edit=yes\">$diveName</a>";
		$expdDiveList .= " &#183; ";
		####print"<br>expdDiveList1: $expdDiveList<br>";
	}
	$RSdives->MoveNext;
}
$expdDiveList =~ s/ &#183; $//;
####print"<br>expdDiveList: $expdDiveList<br>";
$RSdives->close;
##$RSt->close;



###############################################################################
# Get times from Pilot's database table for the dives in this Expedition
# If MiniROV, get times from DiveStart and DiveEndDtg from Dive
###############################################################################
Win32::ASP::DebugPrint("\nstep3():\nPSQL: rovName =  $rovName\n");

		
$ExpdID = $RSt->Fields('ExpeditionID_FK')->value;
####print"<br>$ExpdID<br>";

# Check for MiniROV on non-MBARI Ships because data entry is different.
	if ( $RSt->Fields('ExpeditionID_FK')->value eq 0 && $RSt->Fields('RovName')->value eq 'mini') {	
		$psql = "SELECT DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS DiveStartEpoch, ";
		$psql .= "DateDiff(\"ss\", '01/01/70', DiveEndDtg) AS DiveEndEpoch, DiveNumber, DiveStartDtg, DiveEndDtg ";
		$psql .= " FROM dbo.Dive\n";
		$psql .= " WHERE ";
		$psql .= "ExpeditionID_FK = " . $ExpdID . " AND RovName = '". $rovName . "' AND ";
		$psql .= " DiveStartDtg BETWEEN '" . $RSt->Fields('DiveStartDtg')->value . "' AND '" . $RSt->Fields('DiveEndDtg')->value . "' AND ";
		$psql .= " DiveEndDtg BETWEEN '" . $RSt->Fields('DiveStartDtg')->value . "' AND '" . $RSt->Fields('DiveEndDtg')->value . "' ";
		$psql .= " ORDER BY DiveNumber ";
	#	print "<br>PSQL2 = $psql<br>";
	}
	# Check for MiniROV on MBARI Ships because data entry is different.
	elsif ( $RSt->Fields('ExpeditionID_FK')->value ne 0 && $rovName eq 'mini') {
		$psql = "SELECT RovName, DiveNumber, ExpeditionID_FK, DiveStartDtg, DiveEndDtg, DeviceID, \n";
		$psql .= "DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS StartEpoch, DateDiff(\"ss\", '01/01/70', DiveEndDtg) AS EndEpoch ";
		$psql .= " FROM Dive\n";
		$psql .= " WHERE DeviceID = " . $RS->Fields('DeviceID')->value . "\n";
		$psql .= " AND DiveStartDtg BETWEEN '" . $RS->Fields('DiveStartDtg')->value  . "' AND '" . $RS->Fields('DiveEndDtg')->value  . "' AND ";
		$psql .= " DiveEndDtg BETWEEN '" . $RS->Fields('DiveStartDtg')->value  . "' AND '" . $RS->Fields('DiveEndDtg')->value  . "' ";
		$psql .= " ORDER BY DiveNumber";
	#	print "<br>MiniROV on MBARI Ships:<br>  $psql<br>";
	}
	# Check for Non MiniROV Everything else
	elsif ( $RS->Fields('ExpeditionID_FK')->value ne 0 && $rovName ne 'mini') {
		$psql = "SELECT DateDiff(\"ss\", '01/01/70', gmtStart) AS PilotStartEpoch, ";
		$psql .= "DateDiff(\"ss\", '01/01/70', gmtEnd) AS PilotEndEpoch, DiveNumber, Results ";
		if ($rovName eq 'vnta') {
			$psql .= " FROM dbo.VentanaPilotsDive\n";
		}
		elsif ($rovName eq 'tibr') {
			$psql .= " FROM dbo.TiburonPilotsDive\n";
		}
		elsif ($rovName eq 'docr') {
			$psql .= " FROM dbo.DocRickettsPilotsDive\n";
		}
		$psql .= " WHERE ";
		$psql .= " gmtStart BETWEEN '" . $RSexpd->Fields('ScheduledStartDtg')->value . "' AND '" . $RSexpd->Fields('ScheduledEndDtg')->value . "' AND ";
		$psql .= " gmtEnd BETWEEN '" . $RSexpd->Fields('ScheduledStartDtg')->value . "' AND '" . $RSexpd->Fields('ScheduledEndDtg')->value . "' ";
		$psql .= " ORDER BY DiveNumber ";
		####print"<br>MBARI Non-MiniROV Dive <br>$psql<br>";
	}

Win32::ASP::DebugPrint("\nstep3():\nExecuting PSQL: \n$psql");
$PRSdives = $Conn->Execute($psql);
if(!$PRSdives) {
 	die "<BR>Empty Return Set: PRSdives. for <br>$psql ";
}
while ( !$PRSdives->EOF ) {
	my $diveName = $rovName . $PRSdives->Fields('DiveNumber')->value;
	my $exped = $RS->Fields('ExpeditionID_FK')->value;
	####print"<br>$diveName, $exped<br>";

    if ( $RSt->Fields('ExpeditionID_FK')->value eq 0 && $RSt->Fields('RovName')->value eq 'mini') {	
   		####print"<br>NON-MBARI MINI DIVE<br>";
		$DiveStartEsecs = $PRSdives->Fields('DiveStartEpoch')->value;
		$DiveEndEsecs = $PRSdives->Fields('DiveEndEpoch')->value;
		####print"<br>DiveStartEpochsecs: $DiveStartEsecs, DiveEndEpochsecs: $DiveEndEsecs<br>";
	
		$sPDtl{$diveName} = 600 * ($PRSdives->Fields('DiveStartEpoch')->value - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs);
		$ePDtl{$diveName} = 600 * ($PRSdives->Fields('DiveEndEpoch')->value - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs) - $sDtl{$diveName};
		####print"<br>Pilot Startss and End times for NON-MBARI MiniROV...$sPDtl{$diveName}, $ePDtl{$diveName}<br>";

		# Build timeline hash lists from CTD data

		getCTDTimes($rovName, $DiveStartEsecs, $DiveEndEsecs, $PRSdives->Fields('DiveStartDtg')->Value, $PRSdives->Fields('DiveEndDtg')->Value);

	$PRSdives->MoveNext;
	}
  	elsif ( $RS->Fields('ExpeditionID_FK')->value ne 0 && $rovName eq 'mini' ) {
		####print"<br>MBARI MINIROV DIVE:  $diveName, <br>";
			$DiveStartEsecs = $PRSdives->Fields('StartEpoch')->value;
			$DiveEndEsecs = $PRSdives->Fields('EndEpoch')->value;
			####print"<br>DiveStartEpochsecs: $DiveStartEsecs, DiveEndEpochsecs: $DiveEndEsecs<br>";


			$sPDtl{$diveName} = 600 * ($PilotStartEsecs - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs);
			$ePDtl{$diveName} = 600 * ($PilotEndEsecs - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs) - $sPDtl{$diveName};
			####print"<br>Pilot Startsss and End times for MBARI MiniROV...$sPDtl{$diveName}, $ePDtl{$diveName}<br>";

			# Build timeline hash lists from CTD data

			$DStartEpoch = $RSexpd->Fields('DiveStartDtg')->value; 
			$DEndEpoch = $RSexpd->Fields('DiveEndDtg')->value; 
			####print"<br>$rovName, $ExpdStartEsecs, $ExpdEndEsecs, $StartEpoch, $EndEpoch<br>";

			getCTDTimes($rovName, $ExpdStartEsecs, $ExpdEndEsecs, $DStartEpoch, $DEndEpoch);
	
	$PRSdives->MoveNext;
  	}	
	elsif ( $RS->Fields('ExpeditionID_FK')->value ne 0 && $rovName ne 'mini' ) {
  		    ####print"<br>MBARI NON-MINIROV DIVE<br>";
			$PilotStartEsecs = $PRSdives->Fields('PilotStartEpoch')->value;
			$PilotEndEsecs = $PRSdives->Fields('PilotEndEpoch')->value;
			$PilotResults = (defined $PRSdives->Fields('Results')) ? $PRSdives->Fields('Results')->value : '';
	
			$sPDtl{$diveName} = 600 * ($PilotStartEsecs - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs);
			$ePDtl{$diveName} = 600 * ($PilotEndEsecs - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $sPDtl{$diveName};
			####print"<br>Pilot Start and End times for MBARI ROV...$sPDtl{$diveName}, $ePDtl{$diveName}<br>";

			# Build timeline hash lists from CTD data

			getCTDTimes($rovName, $ExpdStartEsecs, $ExpdEndEsecs, $RSexpd->Fields('ScheduledStartDtg')->Value, $RSexpd->Fields('ScheduledEndDtg')->Value);

	$PRSdives->MoveNext;
	}
} ## End of While loop
$PRSdives->Close;


###############################################################################
# Construct month, day, year option lists for start & end entry
###############################################################################
if ($RS->Fields('ExpeditionID_FK')->Value eq 0) {
	#######print"<br>NON-MBARI SHIP DIVE<br>";
	($Smonth_option_list, $Sday_option_list, $Syear_option_list, $Shr_option_list, $Smin_option_list) = construct_option_lists($DiveStartEpoch);

	($Emonth_option_list, $Eday_option_list, $Eyear_option_list, $Ehr_option_list, $Emin_option_list) = construct_option_lists($DiveEndEpoch);
}
else {
	#######print"<br>MBARI SHIP DIVE<br>";
	($Smonth_option_list, $Sday_option_list, $Syear_option_list, $Shr_option_list, $Smin_option_list) = construct_option_lists($StartEpoch);

	($Emonth_option_list, $Eday_option_list, $Eyear_option_list, $Ehr_option_list, $Emin_option_list) = construct_option_lists($EndEpoch);
}
%>
	
<script Language="JavaScript"><!--
<%=$diveCheckDateFunctions%>

function Field_Validator(theForm) {

	//alert("Update Button Value = " + theForm.update_dive.value)
	
	if ( theForm.SEpochSecs.value > theForm.EEpochSecs.value )
	{
	  //alert("SEpochSecs = " + theForm.SEpochSecs.value + "  EEpochSecs = " + theForm.EEpochSecs.value);
	  alert("Start time is after end time.  (Make sure that there are no leading zeros in your epoch secs entries.)");
	  theForm.SEpochSecs.focus();
	  return (false);
	}
	
	if ( (theForm.EEpochSecs.value - theForm.SEpochSecs.value) > 86400 )
	{
	  alert("Dive duration is greater than 24 hours. \nThis is unlikely. Please double check the times.");
	  return (false);
	}
	
	//alert("Debugging, not submitted.")
	return (true)	// make true for production!
	
} // End Field_Validator(theForm)


  //-->
</script>


<form method="post" action="<%=$this_script%>" id="form1" name="diveForm" onsubmit="return Field_Validator(this)">
<input type="hidden" name="updateID" value="<%= GetFormValue('edit')%>">
<input type="hidden" name="deleteID" value="<%= GetFormValue('edit')%>">
<input type="hidden" name="RovName" value="<%= $rovName%>">
<input type="hidden" name="DiveNumber" value="<%= GetFormValue('DiveNumber')%>">
<input type="hidden" name="ExpdStartEsecs" value="<%=$ExpdStartEsecs%>">
<input type="hidden" name="ExpdEndEsecs" value="<%=$ExpdEndEsecs%>">
<input type="hidden" name="DiveStartEsecs" value="<%=$DiveStartEsecs%>">
<input type="hidden" name="DiveEndEsecs" value="<%=$DiveEndEsecs%>">

<!-- EDIT DIVE INFORMATION FORM PAGE HERE _ JUST DIVE PAGE NOT EXPD -->
<table class="a" border="3">
<tr><td>
<!-- Gray Dive Information header table width -->
<table width="815">
    <tr>
        <td colspan="2" bgcolor="silver" align="center">
           <font face="helvetica,Ariel"><b>Dive Information</b></font>
        </td>
    </tr>
</table>

<table class="a" border=0 cellpadding="0" cellspacing="0">
<% if ( ! $Session->{'RegisteredUser'} ) { %>
<tr><td valign="top"><%= $font_1%><b>Your email:</b></font></td>
<td valign="top">
<%= $font_1%><input type="text"  size="8" name="email_addr" value="<%=GetFormValue('email_addr')%>" onKeyPress="return checkEnter(event)">@mbari.org</font></td>
</tr>
<% } %>
<%

#
# Some constants
#
my %days = (0=>31, 1=> $leap ? 29 : 28, 2=>31, 3=>30, 4=>31, 5=>30, 6=>31, 7=>31, 8=>30, 9=>31, 10=>30, 11=>31);
my %iWid = (0=>6, 1=>5, 2=>6, 3=>6, 4=>6, 5=>6, 6=>6, 7=>6, 8=>6, 9=>6);
my %iNam = (0=>'zero.gif', 1=>'one.gif', 2=>'two.gif', 3=>'three.gif', 4=>'four.gif', 5=>'five.gif', 6=>'six.gif', 7=>'seven.gif', 8=>'eight.gif', 9=>'nine.gif');


#
# Look up associated data for this dive and parse for logr.gz file url
#
my $dlinks = dataLinks( ExpeditionID_FK => $ExpdID, DataType => 'rovctd' );
foreach ( split('</a>', $dlinks) ) {
	Win32::ASP::DebugPrint("\nSearch string = $_");
	next unless ( $_ =~ /.*(http:.*rovctdlogr\.dat\.gz).*/ );
	$logrLink = $+;
	Win32::ASP::DebugPrint("\nMatch = $logrLink");
 
	next unless $logrLink;
	push @logrLinks, $logrLink;
	##$Response->Write("<br>$logrLink");
	$logrFile = $logrLink;
	$logrFile =~ s#http://mww.mbari.org/ARCHIVE/rovctd/.+/\d\d\d\d/\d\d\d/##;
	push @logrFiles, $logrFile;
	$logrLink = '';
}

$dlinks =~ /.*(http:.*logr\.dat\.gz).*/;
my $logrLink = $+;
my $logrFile = $logrLink;
$logrFile =~ s#http://mww.mbari.org/ARCHIVE/rovctd/.+/\d\d\d\d/\d\d\d/##;
$diveNumber = GetFormValue('DiveNumber');

###############################################################################
# SETTING DATA LINKS FOR ROV CTD DATA
###############################################################################

my $dlinks = dataLinks( ExpeditionID_FK => $RS->Fields(ExpeditionID_FK)->value, DataType => 'rovctd' );

####print"<br>DLINKS_DIVE: " . $dlinks . "<br>";

###############################################################################
# HERE IS WHERE THE TAPE LINKS ARE SET TO VARSDB
###############################################################################
open_vars_database();
my $tslinks = tapeSummaryLinks( RovName => $rovName, DiveNumber => GetFormValue('DiveNumber') );
$ConnVARS->Close;

while ( !$RS->EOF ) { 
    foreach $f ( @{$Application->{'DiveFields'}} ) { 
	    if ( $f eq 'ExpeditionID_FK' ) { 
	    	if ( $RS->Fields('ExpeditionID_FK')->value eq 0 ) {
				####print"<br>f = ExpeditionID_FK = 0<br>"; %>
				<input type="hidden" name="ShipName" value="othr">
				<tr><td valign="top"><%= $font_1%><b>Non-MBARI Ship:</b></font></td> <
				<td valign="top"><%= $font_1%><a href="postcruise_karen.asp?step=5&RovName=<%=$RSdives->Fields('RovName')->value%>&DiveNumber=<%=$RSdives->Fields('DiveNumber')->value%>ShipName=<%=$RS->Fields('ShipName')->value%>&edit=<%= $RS->Fields($f)->value%>" target="expd">
				<%= $RS->Fields($f)->value%></a>   (Click to view this dive's Expedition information in another window)</font></td>
	    	<%}
	    	elsif ( $RS->Fields('ExpeditionID_FK')->value ne 0 ) {
				####print"<br>f = ExpeditionID_FK != 0<br>";%>
				<tr><td valign="top"><%= $font_1%><b>ExpeditionID:</b></font></td>
				<td valign="top"><%= $font_1%><a href="postcruise_karen.asp?step=3&ShipName=<%=$RS->Fields('ShipName')->value%>&ShipSeqNum=<%=$RS->Fields('ShipSeqNum')->value%>&edit=<%= $RS->Fields($f)->value%>" target="expd">
				<%= $RS->Fields($f)->value%></a> (Click to view this dive's Expedition information in another window)</font></td>
	    	<%}
	    }
	    elsif ( $f eq 'RovName' ) { %>
			<tr><td valign="top" width="1"><%= $font_1%><b><%= $f%>:</b></font></td>
			<td valign="top"><%= $font_1%><%= $RS->Fields($f)->value%></font></td>
	    <%}
	    elsif ( $f eq 'DiveNumber' ) { %>
			<tr><td valign="top"><%= $font_1%><b><%= $f%>:</b></font></td>
			<td valign="top"><%= $font_1%><b><%= $RS->Fields($f)->value%></b>
			<% if ( @logrLinks ) { 
			$Response->Write(" <br>Parse for possible dive number(s): ");
		}
		foreach ( @logrLinks ) { 
			$logrFile = shift @logrFiles; #<br>%>
			<a href="javascript:OpenData('/expd/queries/greplogr_karen.asp?url=<%= $_%>&string=DIVE_NUM.*NEW','logrparse')"><%= $logrFile%></a>
		<% } %>
			</font></td>			
	    <% }
	    elsif ( $f eq 'DiveStartDtg' ) { %>
			<tr><td colspan=2>
			<br> <%
			if ($dlinks) { %>
				<font color=brown><b>Use these data links to help determine the dive's start and end times:</b><br></font>
				<font color=brown size=2>Use GIF plot to view dive profile(s) and Derived link to grab start & end times. 
				Dive start and end times correspond to the<br> ROV's entire time in the water,
				since various science products, such as CTD data, are collected during that whole period.</font>
				<%= $font_1%><br><br><b>Possible ROVCTD Data:</b> <%= $dlinks%> 
				</font><%
			} %>
			</td></tr>
			
			<tr><td colspan=2>
			 <%
			if ($tslinks) { %>
				
				<%= $font_1%><br><b>Tapes:</b> <%= $tslinks%> 
				</font><br><br><%
			} %>
			</td></tr>

			<tr><td valign="top"><%= $font_1%><b>Dive Start:</b></font></td>
			<td valign="top">	

			<table class="b"><tr><td colspan=2>
			<%
			if ( $RS->Fields('StartEpoch')->value || $RS->Fields('DiveStartEpoch')->value )  { %>
				<font color=brown><b>Modify dive start time:</b> </font>
			<%}
			else { %>
				<font color=brown><b>Please enter dive start time:</b> </font>
			<% }%>
			</td></tr>
			<tr><td><%= $font_1%>
			
			<%
			if ( $RS->Fields('ExpeditionID_FK')->value eq 0 ) {%>
				<input type="radio" name="Stime" onFocus="document.diveForm.Etime[0].checked=true;" value="epoch" checked>
				</font>
				</div></td><td><%= $font_1%>
				<input type="text" size="12" STYLE="font-family:monospace" name="SEpochSecs" value="<%=$StartEpoch%>" onKeyPress="return checkEnter(event)" onChange="calcDate(this.form)"  onFocus="document.diveForm.Stime[0].checked=true;document.diveForm.Etime[0].checked=true;">
				(epoch secs - <font color="brown">beginning of good ROVCTD data, use Derived <i>ddd</i> (csv) link above</font>)  
				</font></td></tr>
			<%}
			elsif ( $RS->Fields('ExpeditionID_FK')->value ne 0 ) {%>		
				<input type="radio" name="Stime" onFocus="document.diveForm.Etime[0].checked=true;" value="epoch" checked>
				</font>
				</td><td><%= $font_1%>
				<input type="text" size="12" STYLE="font-family:monospace" name="SEpochSecs" value="<%=$StartEpoch%>" onKeyPress="return checkEnter(event)" onChange="calcDate(this.form)"  onFocus="document.diveForm.Stime[0].checked=true;document.diveForm.Etime[0].checked=true;">
				(epoch secs - <font color="brown">beginning of good ROVCTD data, use Derived <i>ddd</i> (csv) link above</font>)  
				</font></td></tr>
		    <%}
		    my $cell = "<select name=\"${f}Month\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Stime[1].checked=true;document.diveForm.Etime[1].checked=true;\">$Smonth_option_list</select>\n";
			$cell .= "<select name=\"${f}Day\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Stime[1].checked=true;document.diveForm.Etime[1].checked=true;\">$Sday_option_list</select>\n";
			$cell .= "<select name=\"${f}Year\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Stime[1].checked=true;document.diveForm.Etime[1].checked=true;\">$Syear_option_list</select>\n";
			$cell .= "<select name=\"${f}Hr\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Stime[1].checked=true;document.diveForm.Etime[1].checked=true;\">$Shr_option_list</select>\n";
			$cell .= ":<select name=\"${f}Min\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Stime[1].checked=true;document.diveForm.Etime[1].checked=true;\">$Smin_option_list</select>\n";
				    ##$cell .= "<input type=\"text\" name=\"${f}Time\" size=\"4\" value=\"$Shrmn\" maxlength=\"4\" onFocus=\"document.diveForm.Stime[1].checked=true;\">\n";
			$cell .= "(US/Pacific Time <font color=\"brown\">*</font>)\n\n";%>
			<tr><td><%= $font_1%>
			<input type="radio" name="Stime" onFocus="document.diveForm.Etime[1].checked=true;" value="localdtg">
			</font>
			</td><td><%= $font_1%>
			<input type="text" size="3" STYLE="font-family:monospace" name="SDayOfWeek" value="" onKeyPress="return checkEnter(event)">
			<%= $cell %>
			</font></td></tr>
			<tr><td><%= $font_1%>&nbsp;
			</font>
			</td><td><%= $font_1%>
			<input type="text" size="32" align="left" STYLE="font-family:monospace" name="SGMTString" value="" onKeyPress="return checkEnter(event)" readonly="readonly">
			Year day: <input type="text" size="3" STYLE="font-family:monospace" name="SYearDay" value="" onKeyPress="return checkEnter(event)" readonly="readonly">
			</font></tr>
			
			</table>
			</td> <%


		}
elsif ( $f eq 'DiveEndDtg' ) { %>
			<tr><td valign="top"><%= $font_1%><b>Dive End:</b></font></td>
			<td valign="top">
			<table class="b"><tr><td colspan=2>
			<%
		    if ( $RS->Fields('EndEpoch')->value || $RS->Fields('DiveEndEpoch')->value ) {%>
				<font color=brown><b>Modify dive end time:</b> </font>
			<%
			}
			else { %>
				<font color=brown><b>Please enter dive end time:</b> </font>
			<% }%>
			</td></tr>
			<tr><td class="b td:first-child"><%= $font_1%>
			<%
			if ( $RS->Fields('ExpeditionID_FK')->value eq 0 ) {%>
				<input type="radio" name="Etime" onFocus="document.diveForm.Stime[0].checked=true;" value="epoch" checked>
				</font>
				<%= $font_1%>
				<input type="text" size="12" STYLE="font-family:monospace" name="EEpochSecs" value="<%= $EndEpoch%>" onKeyPress="return checkEnter(event)" onChange="calcDate(this.form)" onFocus="document.diveForm.Etime[0].checked=true;document.diveForm.Stime[0].checked=true;"> 
				(epoch secs - <font color="brown">end of good ROVCTD data, use Derived <i>ddd</i> (csv) link above</font>)  
				</font></td></tr>
			<%}
			elsif ( $RS->Fields('ExpeditionID_FK')->value ne 0 ) {%>		
				<input type="radio" name="Etime" onFocus="document.diveForm.Stime[0].checked=true;" value="epoch" checked>
				</font>
				<td><%= $font_1%>
				<input type="text" size="12" STYLE="font-family:monospace" name="EEpochSecs" value="<%= $EndEpoch%>" onKeyPress="return checkEnter(event)" onChange="calcDate(this.form)" onFocus="document.diveForm.Etime[0].checked=true;document.diveForm.Stime[0].checked=true;"> (epoch secs - <font color="brown">end of good ROVCTD data, use Derived <i>ddd</i> (csv) link above</font>)  
				</font></td></tr>
		    <%}
			my $cell = "<select name=\"${f}Month\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Etime[1].checked=true;document.diveForm.Stime[1].checked=true;\">$Emonth_option_list</select>\n";
			$cell .= "<select name=\"${f}Day\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Etime[1].checked=true;document.diveForm.Stime[1].checked=true;\">$Eday_option_list</select>\n";
			$cell .= "<select name=\"${f}Year\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Etime[1].checked=true;document.diveForm.Stime[1].checked=true;\">$Eyear_option_list</select>\n";
			$cell .= "<select name=\"${f}Hr\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Etime[1].checked=true;document.diveForm.Stime[1].checked=true;\">$Ehr_option_list</select>\n";
			$cell .= ":<select name=\"${f}Min\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Etime[1].checked=true;document.diveForm.Stime[1].checked=true;\">$Emin_option_list</select>\n";
			##$cell .= "<input type=\"text\" name=\"${f}Time\" size=\"4\" value=\"$Ehrmn\" maxlength=\"4\" onFocus=\"document.diveForm.Etime[1].checked=true;\">\n";
			$cell .= "(US/Pacific Time <font color=\"brown\">*</font>)\n\n";%>
			<tr><td><%= $font_1%>
			<input type="radio" name="Etime" onFocus="document.diveForm.Stime[1].checked=true;" value="localdtg">
			</font>
			</td><td><%= $font_1%>
			<input type="text" size="3" STYLE="font-family:monospace" name="EDayOfWeek" value="" onKeyPress="return checkEnter(event)">
			<%= $cell %>
			</font></td></tr>
			<tr><td><%= $font_1%>&nbsp;
			</font>
			</td><td><%= $font_1%>
			<input type="text" size="32" align="left" STYLE="font-family:monospace" name="EGMTString" value="" onKeyPress="return checkEnter(event)" readonly="readonly">
			Year day: <input type="text" size="3" STYLE="font-family:monospace" name="EYearDay" value="" onKeyPress="return checkEnter(event)" readonly="readonly">
			</font></tr>
			</table></td>


			<tr><td valign="top"></td>
			<td>
			<!--<table width="680px" class="b"><tr></tr>-->
			<table width="680px"><tr></tr>
			<tr><td colspan="2"><%= $font_1%>
			Dive duration: <input type="text" size="5" STYLE="font-family:monospace" name="DiveTimeHours" value="" onKeyPress="return checkEnter(event)"> hours
			<br>
			<font color="brown">
			<% if ( $expdDiveList eq '' ) {%>
			The brown bar will adjust as you modify above times
			<% }
			else { %>
			Colored bars show dive timeline, (<font color="cyan">Cyan</font>) bars are other dives that have been entered into the Expedition<br>database. The brown bar will adjust as you modify above times (dives must not overlap).<br>
			<% } %>
			</font>
			</font>
			</td></tr>
			
			<% 
			###############################################################################
			# Loop though dives that have already been assigned to this Expedition
			###############################################################################
	    
			foreach my $d ( sort keys %sDtl ) { 
				####print"<br>$d<br>";
				my $r = substr( $d, 0, 1 );
				####print"<br>First letter: $r<br>";
				
				## This highlights the Expedition DiveID chosen by user
				if ( $d eq $RS->Fields(RovName)->value . $RS->Fields(DiveNumber)->value ) {
					####print"<br>TITO: $d<br>";%>
					<tr><td colspan="2" nowrap>
					<img src="sp_wh.gif" width="" height="3" name="StartDiveTimeLine" hspace="0" vspace="0">
					<img src="sp_brown.gif" width="" height="3" name="EndDiveTimeLine" alt="<%=$d%>" hspace="0" vspace="0">
					<br>
					<img src="sp_wh.gif" width="" height="3" name="StartDiveLabel" hspace="0" vspace="0"><%= $font_2%><%=$d%></font>
					</td></tr> <%
					####print"<br>Clicked $d: Dive - sDtl: $sDtl{$d}, eDtl: $eDtl{$d}<br>";
					if ( $sPDtl{$d} && $ePDtl{$d} ) { %>
						<tr><td colspan="2" nowrap>
						<img src="sp_wh.gif" width="<%=$sPDtl{$d}%>" height="3" hspace="0" vspace="0">
						<img src="sp_grey.gif" width="<%=$ePDtl{$d}%>" height="3" alt="Pilot <%=$d%>" hspace="0" vspace="0">
						<br>
						<img src="sp_wh.gif" width="<%=$sPDtl{$d}%>" height="3" alt="Pilot <%=$d%>" hspace="0" vspace="0"><%= $font_2%>Pilot's <%=$d%></font>
						</td></tr>	<%
						####print"<BR>Clicked $d: Pilot's sPDtl: $sPDtl{$d}, ePDtl: $ePDtl{$d}<BR>";
			       	} 
			    }
			    ## This shows all other Expedition DiveIDs to the user
			    else { %>
<!-- #			    	if ( $d ne $RS->Fields(RovName)->value . $RS->Fields(DiveNumber)->value ) {
#			    		####print"<br>NTO: $d<br>";
#       				my @CTD = split('&middot;', $d);
#        				####print"<br>CTD List: @CTD<br>";
#        				foreach my $ctd (@CTD) {
#          					if ($ctd =~ $d) {
#            					$ctd =~ s/\$d.*//;
#            					$ctd =~ s/\.$d*//s;
#            					%><tr><td><font face="helvetica,Ariel" size="-1"><b>Test:</b></font></td>
#            					<td><font face="helvetica,Ariel" size="-1"><%= $ctd %></font></td></tr><% 
#            			}
#            		}%> -->
					<tr><td colspan="2" nowrap>
					<img src="sp_wh.gif" width="<%=$sDtl{$d}%>" height="3" hspace="0" vspace="0">
					<img src="sp_cyan.gif" width="<%=$eDtl{$d}%>" height="3" alt="<%=$d%>" hspace="0" vspace="0">
					<br>
					<img src="sp_wh.gif" width="<%=$sDtl{$d}%>" height="3" alt="<%=$d%>" hspace="0" vspace="0"><%= $font_2%><%=$d%></font>
					</td></tr> <%
					####print"<br>$d: Dive - sDtl: $sDtl{$d}, eDtl: $eDtl{$d}<br>";
						if ( $sPDtl{$d} && $ePDtl{$d} ) { %>
							<tr><td colspan="2" nowrap>
							<img src="sp_wh.gif" width="<%=$sPDtl{$d}%>" height="3" hspace="0" vspace="0"><img src="sp_grey.gif" width="<%=$ePDtl{$d}%>" height="3" alt="Pilot <%=$d%>" hspace="0" vspace="0">
							<br>
							<img src="sp_wh.gif" width="<%=$sPDtl{$d}%>" height="3" alt="Pilot <%=$d%>" hspace="0" vspace="0"><%= $font_2%>Pilot's <%=$d%></font>
							</td></tr><% 
						####print"<br>$d: Pilot's - sPDtl: $sPDtl{$d}, ePDtl: $ePDtl{$d}<br>";
						}
#			  		} 
				} 
			}
			
			###############################################################################
			# Loop though the pilot's dive(s) that have not been assigned yet to this expedition
			###############################################################################
			foreach my $d ( sort keys %sPDtl ) { 
				next if ( $sDtl{$d} && $eDtl{$d} );
				 %>
				<tr><td colspan="2" nowrap>
				<img src="sp_wh.gif" width="<%=$sPDtl{$d}%>" height="3" hspace="0" vspace="0"><img src="sp_grey.gif" width="<%=$ePDtl{$d}%>" height="3" alt="Pilot <%=$d%>" hspace="0" vspace="0">
				<br>
				<img src="sp_wh.gif" width="<%=$sPDtl{$d}%>" height="3" alt="Pilot <%=$d%>" hspace="0" vspace="0"><%= $font_2%>Pilot's <%=$d%></font>
				</td></tr><%	
				####print"<br>$d: Other Pilot's - sPDtl: $sPDtl{$d}, ePDtl: $ePDtl{$d}<br>";
			}
			###############################################################################
			# Display timeline of ROVCTD data
			###############################################################################
			sub numerically {
				$a <=> $b;
			}
			foreach my $d ( sort numerically keys %sCTDtl ) { 
				Win32::ASP::DebugPrint("\nstep3():\nedit(): \nd = $d, s = $sCTDtl{$d}, e = $eCTDtl{$d}\n");%>
				<tr><td colspan="2" nowrap>
				<img src="sp_wh.gif" width="<%=$sCTDtl{$d}%>" height="3" hspace="0" vspace="0"><img src="sp_lightgrey.gif" width="<%=$eCTDtl{$d}%>" height="3" alt="<%=$d%>" hspace="0" vspace="0"><%
				if ($d == 1) { %>
				<br>
				<img src="sp_wh.gif" width="<%=$sCTDtl{$d}%>" height="3" alt="CTD data" hspace="0" vspace="0"><%= $font_2%>CTD data</font>
				<% } %>
				</td></tr><%
				####print"<br>$d: sCTDtl: $sCTDtl{$d}, eCTDtl: $eCTDtl{$d}<br>";
			}%>
			<tr><td colspan="2" width="0" height="3" valign="bottom" nowrap>
			<img src="sp_grey.gif" width="600" height="3" name="ExpdTimeLine" alt="expd timeline" hspace="0" vspace="0">
			</td></tr>
			<%
		
			# Put GMT year-day boundaries on expedition timeline

			if ( $RS->Fields('ExpeditionID_FK')->value eq 0 ) {		
				$YYYY = POSIX::strftime("%Y", POSIX::gmtime($DiveStartEsecs) );
				$leap = ( $YYYY % 4 == 0 ) ? 1 : 0;
				$leap = ( ! $YYYY % 100 == 0 || $YYYY % 400 == 0 ) ? 1 : 0;
				$leap = ( $YYYY % 400 == 0 ) ? 1 : 0;
				$sDDD = POSIX::strftime("%j", POSIX::gmtime($DiveStartEsecs) );
				$eDDD = POSIX::strftime("%j", POSIX::gmtime($DiveStartEsecs) );
				$mo = POSIX::strftime("%m", POSIX::gmtime($DiveStartEsecs) ) - 1;
				$da = POSIX::strftime("%d", POSIX::gmtime($DiveStartEsecs) );	
			}
			else {
				$YYYY = POSIX::strftime("%Y", POSIX::gmtime($ExpdStartEsecs) );
				$leap = ( $YYYY % 4 == 0 ) ? 1 : 0;
				$leap = ( ! $YYYY % 100 == 0 || $YYYY % 400 == 0 ) ? 1 : 0;
				$leap = ( $YYYY % 400 == 0 ) ? 1 : 0;
				$sDDD = POSIX::strftime("%j", POSIX::gmtime($ExpdStartEsecs) );
				$eDDD = POSIX::strftime("%j", POSIX::gmtime($ExpdEndEsecs) );
				$mo = POSIX::strftime("%m", POSIX::gmtime($ExpdStartEsecs) ) - 1;
				$da = POSIX::strftime("%d", POSIX::gmtime($ExpdStartEsecs) );
			}
			
			my $dayB = 0;
			my $barHtml = '';
			my $txtHtml = '';
			my $accumDayB = 0;
			my $accumDayBlabel = 0;
			my $char3width = 10;
			foreach my $ddd ($sDDD..$eDDD) { 

			################################################################################
			# Create image for year-day text
			###############################################################################
			my $wid = 0;
			my $dddImg = '';
			foreach my $d ( split(//, $ddd) ) {
				$wid += $iWid{$d};
				$dddImg .= "<img src=\"$iNam{$d}\">";
			}
				
				###############################################################################
				# Calculate where to put year-day Boundary
				###############################################################################
				$dayBesecs = timegm(0, 0, 0, $da, $mo, $YYYY);
				Win32::ASP::DebugPrint("\nddd = $ddd, d= $da, mo = $mo, YYYY = $YYYY, dayBesecs = $dayBesecs\nwid = $wid\n");

				if ( $RSt->Fields('ExpeditionID_FK')->Value eq 0 && $RSt->Fields('RovName')->Value eq 'mini' ) { 
					####print"<br>Non-MBARI MiniROV Dive.<br>";
					$dayB = 600 * ( $dayBesecs - $StartEpoch ) / ( $EndEpoch - $StartEpoch) - $accumDayB;
					$dayBlabel = 600 * ( $dayBesecs - $StartEpoch ) / ( $EndEpoch - $StartEpoch) - $accumDayBlabel;
					if ( $dayB < 0 ) {
						$dayB = 0;
						$wid = 0;
						$Bwid = 0;
					}
					else {
						$barHtml .= "<img src=\"sp_wh.gif\" width=\"$dayB\" height=\"1\" hspace=\"0\" vspace=\"0\"><img src=\"sp_grey.gif\" width=\"2\" height=\"20\" alt=\"Year-day boundary\">$dddImg";
						$Bwid = 2;
					}
				}
				elsif ($RSt->Fields('ExpeditionID_FK')->Value ne 0 && $RSt->Fields('RovName')->Value eq 'mini' ) { 
					####print"<br>MBARI MiniROV Dive.<br>";
					$dayB = 600 * ( $dayBesecs - $ExpdStartEsecs ) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $accumDayB;
					$dayBlabel = 600 * ( $dayBesecs - $ExpdStartEsecs ) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $accumDayBlabel;
					if ( $dayB < 0 ) {
						$dayB = 0;
						$wid = 0;
						$Bwid = 0;
					}
					else {
						$barHtml .= "<img src=\"sp_wh.gif\" width=\"$dayB\" height=\"1\" hspace=\"0\" vspace=\"0\"><img src=\"sp_grey.gif\" width=\"2\" height=\"20\" alt=\"Year-day boundary\">$dddImg";
						$Bwid = 2;
					}
				}
				elsif ( $RSt->Fields('ExpeditionID_FK')->Value ne 0 && $RSt->Fields('RovName')->Value ne 'mini' ) { 
					####print"<br>MBARI Non-MiniROV Dive.<br>";
					$dayB = 600 * ( $dayBesecs - $ExpdStartEsecs ) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $accumDayB;
					$dayBlabel = 600 * ( $dayBesecs - $ExpdStartEsecs ) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $accumDayBlabel;
					if ( $dayB < 0 ) {
						$dayB = 0;
						$wid = 0;
						$Bwid = 0;
					}
					else {
						$barHtml .= "<img src=\"sp_wh.gif\" width=\"$dayB\" height=\"1\" hspace=\"0\" vspace=\"0\"><img src=\"sp_grey.gif\" width=\"2\" height=\"20\" alt=\"Year-day boundary\">$dddImg";
						$Bwid = 2;
					}
				}
				$txtHtml .= "<img src=\"sp_wh.gif\" width=\"$dayBlabel\" height=\"1\" hspace=\"0\" vspace=\"0\">${font_2}${ddd}</font>";
				$accumDayB = $accumDayB + $dayB + $wid + $Bwid;
				$accumDayBlabel = $accumDayBlabel + $dayB + $wid;
				$da = ($days{$mo} == $da) ? 1 : $da + 1;
				$mo++ if $da == 1;
			} %>
			<tr><td colspan="2" valign="top" nowrap>
			<%=$barHtml%>
			<br><%= $font_1%>
			<font color="brown">
			Grey bar is Expedition timeline with GMT year-day boundary marks
			</font>
			</font>
			</td></tr>
			
			<%
			###############################################################################
			# Query Summary table to estimate gapiness of the CTD data
			###############################################################################
			$rovName = GetFormValue('RovName');
			$diveNumber = GetFormValue('DiveNumber');

			$ctdSummarySql = <<EOS;
declare \@dstrt as bigint
declare \@dend as bigint
declare \@deltap as bigint
declare \@np as bigint
declare \@nt as bigint
declare \@ns as bigint
declare \@nxmiss as bigint
declare \@no2sbe as bigint
declare \@no2optode as bigint

select \@dstrt = expd.dbo.DTGToEpochSeconds(divestartdtg), \@dend=expd.dbo.DTGToEpochSeconds(diveenddtg), 
 \@np=ctdpcount, \@nt = ctdtcount, \@ns=ctdscount, \@no2sbe=ctdo2count, \@no2optode=ctdo2altcount, \@nxmiss = ctdxmisscount
from [EXPD].[dbo].[DiveSummary]
where divenumber = $diveNumber and rovname = '$rovName'

set \@deltap = (\@dend - \@dstrt) / 15
select \@deltap as 'expected', \@np as 'ctdpcount', \@nt as ctdtcount, \@ns as 'ctdscount', \@no2sbe as 'ctdo2sbecount', \@no2optode as 'ctdo2optodecount', \@nxmiss as 'ctdxmisscount'

EOS

			##################################################################################################################
			# Construct HTML text for report that identifies gaps greater than 5 minutes (at 15 seconds that's a count of 20)
			##################################################################################################################
			$minGap = 5;
			$gapCrit = $minGap * 60 / 15;
			Win32::ASP::DebugPrint("\nstep3():\nExecuting ctdSummarySql: \n$ctdSummarySql");
			$RSCTDsummary = $Conn->Execute($ctdSummarySql);
			$ctdReportHtml = '';
			if(!$RSCTDsummary ) {
 				die "<BR>Empty Return Set6: CTDsummary . for <br>$ctdSummarySql";
			}
			else {			
				$ctdExpectedCount = $RSCTDsummary ->Fields('expected')->value;

				foreach my $field ((qw(ctdpcount ctdtcount ctdscount ctdo2sbecount ctdo2optodecount ctdxmisscount))) {
					$field =~ /ctd(.+)count/;
					my $label = uc $1;
					Win32::ASP::DebugPrint("\nCTD Summary():\nfield = $field: \nlabel = $label");
					my $count = $RSCTDsummary ->Fields($field)->value;
					Win32::ASP::DebugPrint("\nCTD Summary():\ncount = $count");
					$ctdReportHtml .= "$label: $count ";

					if ($count == 0)  {
						$ctdReportHtml .= " - $font_err No Data. </font><br>";
					}
					elsif (abs($ctdExpectedCount - $count) > $gapCrit) {
						$ctdReportHtml .= " - $font_err One or more gap(s) totaling more than $minGap minutes detected. </font><br>";
					}			
					else {
						$ctdReportHtml .= '<br>';
					}
				}
			}
			%>
			<tr><td colspan="2" valign="top" nowrap>
			<%= $font_1%>
			<strong>CTD Data Summary:</strong><br>
			(Expecting <%=$ctdExpectedCount%> values from the 15-second binned data)<br>
			</font>
			
			<%= $font_1%><%=$ctdReportHtml%></font?
			</td>
			</tr>
			
			<tr><td colspan="2" valign="top" nowrap>
			
			</td></tr>
			
			<tr><td colspan="2"><div class="wrap">
			<% if ( $expdDiveList ne '' ) {%>
			<%= $font_1%>Edit other dives from this expedition: <%=$expdDiveList%></font></div>
			<% } %>
			</td></tr>
				
		</table>
			</td> 			
			<%
		}
	    elsif ( $f eq 'DiveChiefScientist' ) { %>
			<tr><td valign="top" class="b td:first-child"><%= $font_1%><b><%= $f%>:</font>&nbsp</b></td>
			<td valign="top">
			<table class="a"><tr><td colspan=2>
				<%= $font_1%><select name="<%=$f%>" size="1">
			<option value=""></option>
			<% while ( !$RScs->EOF ) {
				$fullname = $RScs->Fields('FirstName')->value . " " . $RScs->Fields('LastName')->value;
				
				my $cso_list = "<option value=\"$fullname\"";
				$cso_list .= "selected" if $RS->Fields($f)->Value eq $fullname;
				$cso_list .= ">$fullname</option>\n";
				$Response->write($cso_list);
				$RScs->MoveNext;
			 } %>
      		</select></font>
      	</table></td></tr>
	    <%}
	    elsif ( $f eq 'BriefAccomplishments' ) { 
	    %>
	    <tr><td valign="top"><%= $font_1%><b><br>Dive Accomplishments <br>and Comments:</b>
	    <br><font size="-2">
	    (Fill out if different from 
	    <br>Expedition Accomplishments. 
	    <br>You may also use this <br>field to make comments about <br>the data.)</font></td>
	    <td valign="top"><%= $font_1%><textarea width="100%" rows="5" name="<%=$f%>" cols="90" wrap><%=$RS->Fields($f)->value%></textarea>
	    </font><br>  <%
	    	%>
    		<%= $font_1%><font color="brown">
		Results from Pilot's dive log database:
		<br>
		</font></font>
    		<%= $font_1%>
    		<%=$PilotResults%>
    		</font>
    		<br>
    		<%= $font_1%><font color="brown">
		(Dive Accomplishments and Comments should describe science accomplishments.)
		</font></font>
	    	<% 	    
	    %>
	    </td>
	    <% }
	    else { %>
	    <tr><td valign="top"><b><%= $font_1%><%= $f%>:</b></font></td>
		<td valign="top"><input type="text" size="50" name="<%= $f%>" value="<%= $RS->Fields($f)->value%>"></font></td>
	    <% } %>
	   </tr>
    <% } %> 
    <% $RS->MoveNext; %>
<% } %>
</table>
<%


###################################################################################################################
# Construct timeline for encompassing days of this dive with marks showing various events
################################################################################################################### 
my $diveDaysWidth = 680;
my $YYYY = POSIX::strftime("%Y", POSIX::gmtime($StartEpoch) );
my $sDDD = POSIX::strftime("%j", POSIX::gmtime($StartEpoch) );
my $eDDD = POSIX::strftime("%j", POSIX::gmtime($EndEpoch) );
my $sMo = POSIX::strftime("%m", POSIX::gmtime($StartEpoch) ) - 1;
my $sDa = POSIX::strftime("%d", POSIX::gmtime($StartEpoch) );
my $eMo = POSIX::strftime("%m", POSIX::gmtime($EndEpoch) ) - 1;
my $eDa = POSIX::strftime("%d", POSIX::gmtime($EndEpoch) );
Win32::ASP::DebugPrint("\nsDa = $sDa,  eDa = $eDa\n");
#
# Always show 2 days on the timeline, even if dive ends in the dame day it starts
#
if ( $eDa == $sDa ) {
	$eDa = ($days{$mo} == $eDa) ? 1 : $eDa + 1;
	$eMo++ if $eDa == 1;
	$eDDD++;	# Assume no dives across New Years eve
	Win32::ASP::DebugPrint("\nIncrementing eDa,  eDa = $eDa\n");
}
my $daySesecs = timegm(0, 0, 0, $sDa, $sMo, $YYYY);
my $dayEesecs = timegm(59, 59, 23, $eDa, $eMo, $YYYY);

my $utc_off = 8 - (localtime($StartEpoch))[8];  # Becomes 7 if dst.
Win32::ASP::DebugPrint("\nutc_off = $utc_off\n");
my $diveStic = int($diveDaysWidth * ( $StartEpoch - $daySesecs ) / ( $dayEesecs - $daySesecs));
my $diveWidth = int($diveDaysWidth * ( $EndEpoch - $StartEpoch ) / ( $dayEesecs - $daySesecs));

#
# Read Coolpix epoch second file times for the days of this dive
# 				
my $cptSUrl = "http://search.mbari.org/ARCHIVE/digitalImages";
$cptSUrl .= ($rovName eq 'vnta') ? "/Ventana/$YYYY/$sDDD/esecsLookup.txt" : "/Tiburon/$YYYY/$sDDD/esecsLookup.txt";
my $cptEUrl = "http://search.mbari.org/ARCHIVE/digitalImages/Tiburon/$YYYY/$eDDD/esecsLookup.txt";
$cptEUrl .= ($rovName eq 'vnta') ? "/Ventana/$YYYY/$eDDD/esecsLookup.txt" : "/Tiburon/$YYYY/$eDDD/esecsLookup.txt";
my @cptSContent = split('\n', get($cptSUrl) );
my @cptEContent =  split('\n', get($cptEUrl) );
Win32::ASP::DebugPrint("\ncptSUrl = $cptSUrl\ncptEUrl = $cptEUrl\n");
Win32::ASP::DebugPrint("\ncptSContent = " . $#cptSContent + 1 . " \ncptEContent = " . $#cptEContent + 1 . "\n");
$numCoolpix{$sDDD} = $#cptSContent + 1;
$numCoolpix{$eDDD} = $#cptEContent + 1;
$ticHtml = '';
$guessHtml = '';
my $tCount = 0;
my $imgCount = 0;
$accumTic = 0;
$accumGuess = 0;
$ticSum = 0;
foreach ( sort ( @cptSContent,  @cptEContent ) ) {
	($es, undef) = split;
	next unless $es;
	$imgCount++;
	my $tic = int($diveDaysWidth * ( $es - $daySesecs ) / ( $dayEesecs - $daySesecs) + 0.5) - $accumTic;
	my $guess = int($diveDaysWidth * ( ($es + $utc_off * 3600) - $daySesecs ) / ( $dayEesecs - $daySesecs) + 0.5) - $accumGuess;

	##Win32::ASP::DebugPrint("\n_ = $_ : diffBeg = " , $es - $daySesecs , "  diffEnd = " , $dayEesecs - $es , "\n");
	##Win32::ASP::DebugPrint("accumGuess = $accumGuess, guess = $guess < diveDaysWidth = $diveDaysWidth
##accumTic = $accumTic, tic = $tic ");
	next if $tic > $diveDaysWidth || $tic < 0;		# Skip bogus values
	##if ( $tic > 1 && ($accumTic + $tic) < $diveDaysWidth) {
	$tic = 0 if $tic < 0;
	$guess = 0 if $guess < 0;
	
	if ( $tic > 1 ) {
		##Win32::ASP::DebugPrint("\nwidth=\"$tic\"  ticSum = $ticSum\n");
		# No white space here!  puts extra spaces in the graphics!
		$ticHtml .= "<img src=\"sp_wh.gif\" width=\"$tic\" height=\"1\">";
		$ticHtml .= "<img src=\"sp_black.gif\" width=\"1\" height=\"3\">";
		if ( $accumGuess + $guess < $diveDaysWidth ) {
			# No white space here!  puts extra spaces in the graphics!
			$guessHtml .= "<img src=\"sp_wh.gif\" width=\"$guess\" height=\"1\">";
			$guessHtml .= "<img src=\"sp_red.gif\" width=\"1\" height=\"3\">";
		}
		$tCount++;
		$accumTic = $accumTic + $tic + 1;
		$accumGuess = $accumGuess + $guess + 1;
		$ticSum += $tic;
	}
	
	
}
Win32::ASP::DebugPrint("\nimgCount = $imgCount,  tCount = $tCount\n");

#########################################################
# Get number of Coolpix images assembled for this dive
#########################################################
my $diveName = $rovName . GetFormValue('DiveNumber');
my $cpDiveUrl = "http://search.mbari.org/ARCHIVE/digitalImages/Tiburon/$YYYY/$diveName/";
my $cpDayUrl = "http://search.mbari.org/ARCHIVE/digitalImages/Tiburon/$YYYY";
my $cpDiveImgCount = 0;
foreach (split('\n', get($cpDiveUrl) ) ) {
	##Win32::ASP::DebugPrint("\n_ = $_\n");
	next unless /T.jpg/;	# Count just the preview jpgs
	$cpDiveImgCount++;
}


##############################################################
# Read VARS epoch second file times for the days of this dive
##############################################################
my $vtSUrl = "http://search.mbari.org/ARCHIVE/frameGrabs/";			
$vtSUrl .= ($rovName eq 'vnta') ? "/Ventana/images/" : "/Tiburon/images/";
$vtSUrl .= sprintf("%04d", GetFormValue('DiveNumber')) . "/vfcbydive.log";
my @vtSContent = split('\n', get($vtSUrl) );
Win32::ASP::DebugPrint("\nvtSUrl = $vtSUrl\n");
Win32::ASP::DebugPrint("\nvtSContent = " . $#vtSContent + 1 . "\n");

$ticVarsHtml = '';
$imgVarsCount = 0;
$tCount = 0;
$accumTic = 0;
$ticSum = 0;
foreach ( sort @vtSContent ) {
	/Frame.+es\s\=\s(\d+)/;
	$es = $1;
	next unless $es;
	$imgVarsCount++;
	
	my $tic = int($diveDaysWidth * ( $es - $daySesecs ) / ( $dayEesecs - $daySesecs) + 0.5) - $accumTic;
	my $guess = int($diveDaysWidth * ( ($es + $utc_off * 3600) - $daySesecs ) / ( $dayEesecs - $daySesecs) + 0.5) - $accumGuess;

	##Win32::ASP::DebugPrint("\n_ = $_ : diffBeg = " , $es - $daySesecs , "  diffEnd = " , $dayEesecs - $es , "\n");
	##accumTic = $accumTic, tic = $tic ");
	next if $tic > $diveDaysWidth || $tic < 0;		# Skip bogus values
	##if ( $tic > 1 && ($accumTic + $tic) < $diveDaysWidth) {
	$tic = 0 if $tic < 0;
	
	if ( $tic > 1 ) {
		##Win32::ASP::DebugPrint("\nwidth=\"$tic\"  ticSum = $ticSum\n");
		# No white space here!  puts extra spaces in the graphics!
		$varsStartWidth = $tic if $tCount == 0;
		$ticVarsHtml .= "<img src=\"sp_wh.gif\" width=\"$tic\" height=\"1\">";
		$ticVarsHtml .= "<img src=\"sp_black.gif\" width=\"1\" height=\"3\">";
		
		$tCount++;
		$accumTic = $accumTic + $tic + 1;
		$ticSum += $tic;
	}
}

Win32::ASP::DebugPrint("\nimgVarsCount = $imgVarsCount,  tCount = $tCount\n");


#############################################################
# Code block to label days on timeline (borrowed from above)
#############################################################
my $dayB = 0;
my $barHtml = '';
my $numHtml = '';

my $accumDayB = 0;
my $accumDayBlabel = 0;
my $char3width = 10;
my $da = $sDa;
my $mo = $sMo;
foreach my $ddd ($sDDD..$eDDD) { 

	#
	# Create image for year-day text
	#
	my $wid = 0;
	my $dddImg = '';
	foreach my $d ( split(//, $ddd) ) {
		$wid += $iWid{$d};
		$dddImg .= "<img src=\"$iNam{$d}\">";
	}
	
	#
	# Calculate where to put year-day Boundary
	#
	$dayBesecs = timegm(0, 0, 0, $da, $mo, $YYYY);
	Win32::ASP::DebugPrint("\nddd = $ddd, d= $da, mo = $mo, YYYY = $YYYY, dayBesecs = $dayBesecs\nwid = $wid\n");
	$dayB = $diveDaysWidth * ( $dayBesecs - $daySesecs ) / ( $dayEesecs - $daySesecs) - $accumDayB;
	$dayBlabel = $diveDaysWidth * ( $dayBesecs - $daySesecs ) / ( $dayEesecs - $daySesecs) - $accumDayBlabel;
	if ( $dayB < 0 ) {
		$dayB = 0;
		$wid = 0;
		$Bwid = 0;
	}
	else {
		$barHtml .= "<img src=\"sp_wh.gif\" width=\"$dayB\" height=\"1\" hspace=\"0\" vspace=\"0\"><img src=\"sp_grey.gif\" width=\"2\" height=\"20\" alt=\"Year-day boundary\">$dddImg";
		$numHtml .= "<img src=\"sp_wh.gif\" width=\"$dayB\" height=\"1\" hspace=\"0\" vspace=\"0\">${font_2}<a href=\"${cpDayUrl}/${ddd}\" target=\"day\">${numCoolpix{$ddd}} coolpix during year-day $ddd</a></font>";
		Win32::ASP::DebugPrint("\nnumHtml = $numHtml\n");
		$Bwid = 2;
	}
	
	$accumDayB = $accumDayB + $dayB + $wid + $Bwid;
	$accumDayBlabel = $accumDayBlabel + $dayB + $wid;
	$da = ($days{$mo} == $da) ? 1 : $da + 1;
	$mo++ if $da == 1;
} 

#
# Show coolpix timeline only if we have coolpix images ($ticHtml)
#
if ( $ticHtml && $diveStic > 0 ) {
%>
	<br>
	<%= $font_1%>
	<font color="brown">
	Coolpix image snap times are show here for the days that this dive spans.  The dive timeline here will not adjust 
	when the dive times are changed above.  Use this display to confirm that the coolpix image times are within the bounds 
	of a defined dive. (Only one dive is show here - the dive being edited.)
	</font></font>
	
	<!-- Begin Dive - Coolpix info ***************************************** -->
	    
	
	<br>
	<img src="sp_wh.gif" width="<%=$diveStic%>" height="3"><img src="sp_brown.gif" width="<%=$diveWidth%>" height="3">
	<br>
	<img src="sp_wh.gif" width="<%=$diveStic%>" height="3"><%=$font_2%><%=$diveName%>
	<% if ( $cpDiveImgCount > 0 ) { %>
	 - <a href="<%=$cpDiveUrl%>" target="diveCP"><%=$cpDiveImgCount%> coolpix images</a></font>
	<% } %>
	<br>
	<%=$ticHtml%>	
	<br>
	<%=$numHtml%>
	<br>
	
	<!-- Commented out 22 July 2003 as the Coolpix camera seems to be on GMT now
	
	<%=$guessHtml%>
	<BR>
	<%=$font_2%><font color="red">Adding GMT offset of <%=$utc_off%> hours to coolpix times</font></font>
	<br>
	<img src="sp_grey.gif" width="<%=$diveDaysWidth%>" height="3" hspace="0" vspace="0">
	<br>	
	<%=$barHtml%>
	<br>
	
	End  Dive - Coolpix info ***************************************** -->
	
<%
}
elsif ( $ticVarsHtml ) {  %>
	<br>
	<img src="sp_wh.gif" width="<%=$diveStic%>" height="3"><img src="sp_brown.gif" width="<%=$diveWidth%>" height="3">
	<br>
<%
} # End if ( $ticHtml )

#
# Show tic marks for the VARS image times
#
if ( $ticVarsHtml ) {  
	$varsImageDir = sprintf("%04d", GetFormValue('DiveNumber')); 
%>
	<img src="sp_wh.gif" width="<%=$varsStartWidth%>" height="3"><%=$font_2%><%=$varsImageDir%> - <a href="<%=$vtSUrl%>" target="diveVARS"><%=$imgVarsCount%> VARS images</a></font>
	<br>
	<%=$ticVarsHtml%>	
	<BR>
<%
}

%>
		
<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="submit" class="button delete" value="Delete Dive" name="delete_dive">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="reset" class="button gray" value="Undo Form Changes" name="reset">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="submit" class="button blue" name="update_dive" value="Update Dive"><br>
</form>
<%
if ( $Session->{'RegisteredUser'} ) {

}
else { %>
	You will get a blind copy of the mail that is sent to <%=$Application->{'logistics_email'}%>.
	<br>
<%
}
%>
<br>
<font class="td smalltimesText" color="brown">* Local time is what is in the pilot's dive log.  You can not enter GMT time on this form.  
It is displayed in a text box on this form, but it cannot be edited.  The preferred method of entering time is
in epoch seconds as read from the ROVCTD data file.  This accurately defines the start and end of the dive based
on when we logged good CTD data. Note: The time calculations on this form use your local computer's time zone setting.  
Your computer's time zone needs to be set to "(GMT-08:00) US/Pacific" with daylight savings time adjustment for this 
all to work right.</font><br><br>
<hr width="75%" align="center">
<!--#include file="postcruise_ftr_karen.inc"-->
</td></tr>
</table>

<script Language="JavaScript"><!--
	calcDate(document.diveForm)
//-->
</script>	
<%

$RS->Close;
$RSt->Close;
$RSexpd->Close;
$RScs->Close;
$Conn->Close;
}	# End edit()
%>

<%
#--------------------------------------------------------------------
#

=head3 list_dives()

Present list of dives with an opportunity to edit
each record.

Mike McCann

Date Created: 7/13/99

=cut

sub list_dives {

#
# Open DB
#
open_database($dsn);		# Creates $Conn object as a global variable

#
# Get all the records
#
$sql="SELECT * FROM Dive ORDER BY RovName, DiveNumber";
Win32::ASP::DebugPrint("\nExecuting SQL for $dsn: $sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
 	die "<BR>Empty Return Set from $dsn:<pre>\n$sql</pre>";
}

#
# Print as HTML table - no, not a table, it takes too long
#
for ($i = 0; $i < $RS->Fields->Count; $i++) {
	push @fields, $RS->Fields($i)->Name;
} 

$str = join(' ', @fields);
Win32::ASP::DebugPrint("\nfields: $str ");
while ( !$RS->EOF ) {
	$line = "";
	$line .= "<a href=$this_script?DiveNumber=" . $RS->Fields('DiveNumber')->Value;
	$line .= "&RovName=". $RS->Fields('RovName')->Value . "&edit=yes>";
	$line .= $RS->Fields('RovName')->Value . $RS->Fields('DiveNumber')->Value . "</a><br>";
	$Response->Write("<font face=\"Helvetica,Ariel\"> &#183; $line</font>");
	$RS->MoveNext;
} 

$RS->Close;
$Conn->Close;
}	# End list_dives()
%>

<%

#--------------------------------------------------------------------
#

=head3 update()

Update Dive record with changes made on the form.

Mike McCann

Date Created: 7/15/99

=cut

sub update {

	#
	# Construct the GMT start and End times for the database update
	# Save things to be updated in the database as session variables
	#
	if ( GetFormValue('Stime') eq 'localdtg' ) {
		@sl = (	0, GetFormValue('DiveStartDtgMin'), GetFormValue('DiveStartDtgHr'),
				GetFormValue('DiveStartDtgDay'), GetFormValue('DiveStartDtgMonth')-1, GetFormValue('DiveStartDtgYear')-1900 );
		$Session->{'DiveStartDtg'} = toGMT(@sl);
		$Session->{'StartEpoch'} = timelocal(@sl);		# For original check
		$Session->{'DiveStartDtgLocal'} = GetFormValue('DiveStartDtgYear') . '-' .
					GetFormValue('DiveStartDtgMonth') . '-' .
					GetFormValue('DiveStartDtgDay') . ' ' . 
					GetFormValue('DiveStartDtgHr') .
					GetFormValue('DiveStartDtgMin');
	}
	elsif ( GetFormValue('Stime') eq 'epoch' ) {
		$Session->{'StartEpoch'} = GetFormValue('SEpochSecs');
		Win32::ASP::DebugPrint("\nStartEpoch = " . $Session->{'StartEpoch'} );
		$Session->{'DiveStartDtg'} = toGMTfromES($Session->{'StartEpoch'});
	}
	
	if ( GetFormValue('Etime') eq 'localdtg' ) {
		@el = (	0, GetFormValue('DiveEndDtgMin'), GetFormValue('DiveEndDtgHr'),
				GetFormValue('DiveEndDtgDay'), GetFormValue('DiveEndDtgMonth')-1, GetFormValue('DiveEndDtgYear')-1900 );
				
		$Session->{'DiveEndDtg'} = toGMT(@el);
		$Session->{'EndEpoch'} = timelocal(@el);		# For original check
		$Session->{'DiveEndDtgLocal'} = GetFormValue('DiveEndDtgYear') . '-' .
					GetFormValue('DiveEndDtgMonth') . '-' .
					GetFormValue('DiveEndDtgDay') . ' ' . 
					GetFormValue('DiveStartDtgHr') .
					GetFormValue('DiveStartDtgMin');
	}
	elsif ( GetFormValue('Etime') eq 'epoch' ) {
		$Session->{'EndEpoch'} = GetFormValue('EEpochSecs');
		Win32::ASP::DebugPrint("\nEndEpoch = " . $Session->{'EndEpoch'} );
		$Session->{'DiveEndDtg'} = toGMTfromES($Session->{'EndEpoch'});
	}

	$Session->{'DiveChiefScientist'} = GetFormValue('DiveChiefScientist');
	$Session->{'BriefAccomplishments'} = GetFormValue('BriefAccomplishments');

#
# Open connection to DataBase
#
open_database($dsn);		# Creates $Conn object as a global variable

#
# Open Recrord Set that we'll modify values based on Web form entries and
# then use the Update method on the RS object to update this record in the DB.
#
$sql = "SELECT * FROM Dive WHERE RovName = '" . GetFormValue('RovName');
$sql .= "' AND DiveNumber = " . GetFormValue('DiveNumber');
Win32::ASP::DebugPrint("\nExecuting on $dsn, SQL=\n$sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
 	die "<BR>Empty Return Set from $dsn:<pre>\n$sql</pre>";
}


$sql = "UPDATE Dive\nSET ";

foreach $f ( @{$Application->{'DiveFields'}} ) {
	
	Win32::ASP::DebugPrint("$f: " . $RS->Fields($f)->Value . "\n");
	next if ($f =~ 'RovName' || 
			$f =~ 'DiveNumber' ||
			$f =~ 'ExpeditionID_FK');
	
	if ( $Session->{$f} eq '') {
		$sql .= "$f=NULL,\n";
	}
	else {
		$sql .= "$f=" . FixString( $Session->{$f}, 'text') . ",\n";
	}
	Win32::ASP::DebugPrint("After, $f: " . $RS->Fields($f)->Value . "\n");
}
$sql =~ s/,$//;
$sql .= "WHERE (RovName = '" . GetFormValue('RovName') . "' AND DiveNumber = " . GetFormValue('DiveNumber') . ")";

Win32::ASP::DebugPrint("Executing \n$sql\n");
$Conn->Execute($sql);


$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("Database Load Error(s):<pre>$sql</pre>");
	foreach $error (keys %$Errors) {
      		$Response->Write($error->{Description});
   	}
}
else {
	%>
	<p><B>Dive <%= GetFormValue('RovName')%> <%= GetFormValue('DiveNumber')%> has been updated.</B></p>
<%
}
%>


<%
$RS->Close;
$Conn->Close;
}	# End update()
%>


<%
#--------------------------------------------------------------------
#

=head3 delete_rec()

Delete a record.

Mike McCann

Date Created: 7/14/99

=cut

sub delete_rec {

#
# Open DB
#
open_database($dsn);		# Creates $Conn object as a global variable

#
# Update the records
#
$sql = "DELETE Dive \nWHERE (DiveNumber = " . GetFormValue('DiveNumber');
$sql .= " AND RovName = '" . GetFormValue('RovName') . "') \n";
Win32::ASP::DebugPrint("\nExecuting SQL:\n$sql");

$Conn->Execute($sql);

$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("Database Load Error(s):<pre>$sql</pre>");
	foreach $error (keys %$Errors) {
      		$Response->Write($error->{Description});
   	}
}
else {
	$Response->Write("<b>Dive " . GetFormValue('RovName') . GetFormValue('DiveNumber') . " successfully deleted.</b><p>");
}

$Conn->Close;
%>

<%

}	# End delete_rec()
%>
<%
#--------------------------------------------------------------------
#

=head3 SelectEpoch()

Select epoch seconds from database field specified.

Mike McCann

Date Created: 10/28/98

=cut

sub SelectEpoch {
	my $time_field = $_[0];
	my $ExpeditionID = $_[1];

    ####print"<br>timefield: " . $time_field . "<br>";
    ####print"<br>ExpeditionID: " . $ExpeditionID . "<br>";

 ####print"<br>timefield: $time_field,  ExpeditionID: $ExpeditionID<br>";
	if ($time_field == 'StartDtg' && $RSt->Fields('RovName')->value ne 'mini')  {
	    $ExpeditionID = $RSt->Fields('ExpeditionID_FK')->value unless $ExpeditionID;
			####print"<br>MBARI NON MINI DIVE<br>";
			$sql = "SELECT DateDiff(\"ss\", '01/01/70', $time_field) AS Epoch ";
			$sql .= "FROM Expedition\nWHERE ";
			$sql .= "ExpeditionID = " . $ExpeditionID . "\n";
		    ####print"<br> SelectEpoch expd:  $time_field, $ExpeditionID<br>";
		Win32::ASP::DebugPrint("\nSelectEpoch():\nExecuting SQL: \n$sql");
		$RSdate = $Conn->Execute($sql);
		if(!$RSdate) {
			die "<BR>Empty Return Set9: RSdate. for <br>$sql ";
		}
		my $epoch = $RSdate->Fields('Epoch')->Value;
		####print"<br>EPOCH: $epoch<br>";
		$RSdate->Close;	
	return($epoch);
	}
	if ($time_field == 'DiveStartDtg' && $RSt->Fields('RovName')->value ne 'mini' && $RSt->Fields('ExpeditionID_FK')->value ne 0 ) {
		####print"<br>SelectEpoch = DiveStartDtg Expd=$ExpdID<br>";
		$sql = "SELECT DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS Epoch ";
		$sql .= "FROM Dive\nWHERE ";
		$sql .= "ExpeditionID_FK = " . $ExpeditionID . "\n";
		####print"<br> SelectEpoch expd:  $time_field, $ExpeditionID<br>";
		Win32::ASP::DebugPrint("\nSelectEpoch():\nExecuting SQL: \n$sql");
		$RSdate = $Conn->Execute($sql);
		if(!$RSdate) {
			die "<BR>Empty Return Set SE: RSdate. for <br>$sql ";
		}
		my $epoch = $RSdate->Fields('Epoch')->Value;
		####print"<br>EPOCH: $epoch<br>";
		$RSdate->Close;	
	return($epoch);
	}
	if ($time_field == 'DiveStartDtg' && $RSt->Fields('RovName')->value eq 'mini' && $RSt->Fields('ExpeditionID_FK')->value ne 0 )  {
	    $ExpeditionID = $RSt->Fields('ExpeditionID_FK')->value;
			####print"<br>MBARI MINI DIVE, $ExpeditionID<br>";
			$sql = "SELECT DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS Epoch ";
			$sql .= "FROM Dive\nWHERE ";
			$sql .= "ExpeditionID_FK = " . $ExpeditionID . "\n";
			####print"<br>$sql<br>";
		Win32::ASP::DebugPrint("\nSelectEpoch():\nExecuting SQL: \n$sql");
		$RSdate = $Conn->Execute($sql);
		if(!$RSdate) {
			die "<BR>Empty Return Set9: RSdate. for <br>$sql ";
		}
		my $epoch = $RSdate->Fields('Epoch')->Value;
		####print"<br>EPOCH: $epoch<br>";
		$RSdate->Close;	
	return($epoch);
	}
	if ($time_field == 'DiveStartDtg' && $RSt->Fields('RovName')->value eq 'mini' && $RSt->Fields('ExpeditionID_FK')->value eq 0 )  {
		$ExpeditionID = $RSt->Fields('ExpeditionID_FK')->value;
			####print"<br>NON-MBARI MINI DIVE<br>";
			$sql = "SELECT DateDiff(\"ss\", '01/01/70', $time_field) AS Epoch ";
			$sql .= "FROM Dive\nWHERE ";
			$sql .= "ExpeditionID_FK = " . $ExpeditionID . " AND DeviceID = 1859\n";
			####print"<br>$sql<br>";
		Win32::ASP::DebugPrint("\nSelectEpoch():\nExecuting SQL: \n$sql");
		$RSdate = $Conn->Execute($sql);
		if(!$RSdate) {
			die "<BR>Empty Return Set9: RSdate. for <br>$sql ";
		}
		my $epoch = $RSdate->Fields('Epoch')->Value;
		####print"<br>EPOCH: $epoch<br>";
		$RSdate->Close;	
	return($epoch);
	}
}  # End of SelectEpoch
%>

<%
#--------------------------------------------------------------------
#

=head3 construct_option_lists()

Construct option lists for use in the DTG Month/day/year pull-down
boxes.

Mike McCann

Date Created: 11/19/98

=cut

sub construct_option_lists {

my ($Epoch) = $_[0];

my $lstr = localtime(time);	# Arghhh! get 2 digit yr from above!
my @lstr2 = split('\s+', $lstr);
my $NextYear = $lstr2[4] + 1;

my @l = localtime($Epoch);
my $day = $l[3];
my $mo = $l[4] + 1;
my $month = ($mo < 10) ? "0" . $mo : $mo;
my $lstr = localtime($Epoch);	# Arghhh! get 2 digit yr from above!
my @lstr2 = split('\s+', $lstr);
my $year = $lstr2[4];

%MONTH_ARRAY = ('01', 'Jan',     '02', 'Feb',
                '03', 'Mar',     '04', 'Apr', 
               	'05', 'May',     '06', 'Jun', 
               	'07', 'Jul',     '08', 'Aug', 
               	'09', 'Sep',     '10', 'Oct', 
               	'11', 'Nov',     '12', 'Dec');
#
# Construct option list for months
#
$month_option_list = '';
foreach $num (sort keys (%MONTH_ARRAY)) {
   if ( $num eq $month ) {
	$month_option_list .= "<option value=\"$num\" selected>" . 
				$MONTH_ARRAY{$num} . "</option>\n";
   }
   else {
	$month_option_list .= "<option value=\"$num\">" . 
				$MONTH_ARRAY{$num} . "</option>\n";
   }
}
##Win32::ASP::DebugPrint("month_option_list = \n$month_option_list ");

#
# Construct option list for days
#
$day_option_list = '';
for $num (1..31) {
   ##$val = ($num < 10) ? '0'.$num : $num;	## May not work right.
   if ( $num eq $day ) {
	$day_option_list .= "<option value=\"" . sprintf("%02s",$num) . "\" selected>" . 
				$num . "</option>\n";
   }
   else {
	$day_option_list .= "<option value=\"" . sprintf("%02s",$num) . "\">" . 
				$num . "</option>\n";
   }
}
##Win32::ASP::DebugPrint("day_option_list = \n$day_option_list ");

#
# Construct option list for years
#
$year_option_list = '';

for ($y = 1988; $y <= $NextYear; $y++) {
   if ( $y eq $year ) {
	$year_option_list .= "<option value=\"$y\" selected>" . 
				$y . "</option>\n";
   }
   else {
	$year_option_list .= "<option value=\"$y\">" . 
				$y . "</option>\n";
   }
}

my $hr = ($l[2] < 10 ) ? "0" . $l[2] : $l[2];
my $min = ($l[1] < 10 ) ? "0" . $l[1] : $l[1];

#
# Construct option list for hours
#
$hr_option_list = '';

for ($h = 0; $h <= 23; $h++) {
   if ( $h eq $hr ) {
	$hr_option_list .= "<option value=\"" . sprintf("%02s",$h) . "\" selected>" . 
				sprintf("%02s",$h) . "</option>\n";
   }
   else {
	$hr_option_list .= "<option value=\"" . sprintf("%02s",$h) . "\">" . 
				sprintf("%02s",$h) . "</option>\n";
   }
}

#
# Construct option list for minutes
#
$min_option_list = '';

for ($mi = 0; $mi <= 59; $mi++) {
   if ( $mi eq $min ) {
	$min_option_list .= "<option value=\"" . sprintf("%02s", $mi). " \" selected>" . 
				sprintf("%02s", $mi) . "</option>\n";
   }
   else {
	$min_option_list .= "<option value=\"" . sprintf("%02s",$mi) . "\">" . 
				sprintf("%02s",$mi) . "</option>\n";
   }
}
Win32::ASP::DebugPrint("hr_option_list = \n$hr_option_list ");


my $hrmn = $hr.$min;


return ($month_option_list, $day_option_list, $year_option_list, $hr_option_list, $min_option_list );

}	# End construct_option_lists()
%>

<%
#--------------------------------------------------------------------
#

=head3 add_dive_form()

Present blank form for user to enter dive info.

Mike McCann

Date Created: 10/30/99

=cut

sub add_dive_form {
#
# Open DB
#
open_database($dsn);		# Creates $Conn object as a global variable
##########################################
# Get pick list for Dive Chief Scientist
##########################################
$sql="SELECT * FROM Person WHERE (DisplayChiefSciPickList = 1) ORDER BY LastName, FirstName";
Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RScs = $Conn->Execute($sql);
if(!$RScs) {
 	die "<BR>Empty Return Set: RScs. ";
}

#####################################################################################
# Do an explicit select on Expedition to get possible start & end times for dive  
# - This will not work with Non MBARI ship cruises (e.g., miniROV)
#####################################################################################
$ExpdID = GetFormValue('ExpeditionID');
$sql = "SELECT " . join( ", ", @{$Application->{'ExpeditionFields'}} ) . " FROM Expedition\n";
$sql .= "WHERE ";
$sql .= "ExpeditionID = " . $ExpdID;

Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: \n$sql");
$RSexpd = $Conn->Execute($sql);
if(!$RSexpd) {
 	die "<BR>Empty Return Set: RSexpd. for <br>$sql ";
}

if ( $ExpdID != 0 ) {
	$sql = "SELECT Dive.ExpeditionID_FK AS ExpeditionID_FK, Dive.RovName AS RovName, Dive.DeviceID AS DeviceIDD, Dive.DiveNumber AS DiveNumber, " . join( ", ", @{$Application->{'ExpeditionFields'}} ) . " FROM Expedition\n";
    $sql .= "JOIN Dive ON Expedition.ExpeditionID = Dive.ExpeditionID_FK \n";
    $sql .= "WHERE ";
    $sql .= "ExpeditionID = " . $ExpdID;
}
Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RSt = $Conn->Execute($sql);
  if(!$RSt) {
      $Errors = $Conn->Errors();
      print "Errors:\n";
      foreach $error (keys %$Errors) {
        print $error->{Description}, "\n";
      } 
  die "<BR>Empty Return Set1 for \n$sql \n\nfrom $dsn";
}

#########################################################################################################
# Prime start & end times for dive base on 0) Passed dive times, 1) Expedition times, 2) Scheduled times
#########################################################################################################
if ( $RSt->Fields('RovName')->Value && $RSt->Fields('DiveNumber')->Value ) {
  $sql = "SELECT RovName, DiveNumber, DeviceID, DiveStartDtg, DiveEndDtg, ExpeditionID_FK ";
  $sql .= "FROM Dive WHERE DiveNumber = " . $RSt->Fields('DiveNumber')->Value ; 
  $sql .= " AND RovName = '" . $RSt->Fields('RovName')->value . "'";
}
Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RS = $Conn->Execute($sql);
  if(!$RS) {
      $Errors = $Conn->Errors();
      print "Errors:\n";
      foreach $error (keys %$Errors) {
        print $error->{Description}, "\n";
      } 
      die "<BR>Empty Return Set2 for \n$sql \n\nfrom $dsn";
  }

if ( $RSexpd->Fields('StartDtg')->Value && $RSexpd->Fields('EndDtg')->Value  ) {
	$StartEpoch = SelectEpoch("StartDtg", $ExpdID) || SelectEpoch("ScheduledStartDtg", $ExpdID);
	$ExpdStartEsecs = SelectEpoch("StartDtg", $ExpdID) || SelectEpoch("ScheduledStartDtg", $ExpdID);
	Win32::ASP::DebugPrint("\nStartEpoch = $StartEpoch");

	$EndEpoch =  SelectEpoch("EndDtg", $ExpdID) || SelectEpoch("ScheduledEndDtg", $ExpdID);
	$ExpdEndEsecs = SelectEpoch("EndDtg", $ExpdID) || SelectEpoch("ScheduledEndDtg", $ExpdID);
	Win32::ASP::DebugPrint("\nEndEpoch = $EndEpoch");
} else {
	$StartEpoch = GetFormValue('DiveStartEpoch') || SelectEpoch("StartDtg", $ExpdID) || SelectEpoch("ScheduledStartDtg", $ExpdID);
	$ExpdStartEsecs = SelectEpoch("StartDtg", $ExpdID) || SelectEpoch("ScheduledStartDtg", $ExpdID);
	Win32::ASP::DebugPrint("\nStartEpoch = $StartEpoch");

	$EndEpoch =  GetFormValue('DiveEndEpoch') || SelectEpoch("EndDtg", $ExpdID) || SelectEpoch("ScheduledEndDtg", $ExpdID);
	$ExpdEndEsecs = SelectEpoch("EndDtg", $ExpdID) || SelectEpoch("ScheduledEndDtg", $ExpdID);
	Win32::ASP::DebugPrint("\nEndEpoch = $EndEpoch");
}

$passedDiveNumber = GetFormValue('passedDiveNumber');

$rovName = 'vnta' if $RSexpd->Fields('ShipName')->Value eq 'ptlo';
$rovName = 'vnta' if $RSexpd->Fields('ShipName')->Value eq 'rcsn';
$rovName = 'tibr' if ($RSexpd->Fields('ShipName')->Value eq 'wfly' && $ExpdStartEsecs < 1230768000);	# before 1/1/2009
$rovName = 'docr' if ($RSexpd->Fields('ShipName')->Value eq 'wfly' && $ExpdStartEsecs > 1230768000);	# after 1/1/2009
$rovName = 'mini' if ($RSexpd->Fields('ShipName')->Value eq 'othr' && ($RSexpd->Fields('DeviceIDD')->Value eq 1859 || $qRovName eq 'mini'));

$rovName = ( $RSexpd->Fields('RovName')->Value ) unless $rovName;

#######print"<br>$StartEpoch, $ExpdStartEsecs, $EndEpoch, $ExpdEndEsecs " . $RSexpd->Fields('ShipName')->Value . $ExpdID . " DiveNumber: " . $passedDiveNumber . "<br>";
		
###########################################
# Get all the dives from this expedition
###########################################
$ExpdID = GetFormValue('ExpeditionID') unless $ExpdID;
$sql = "SELECT ExpeditionID_FK, RovName, DiveNumber, DeviceID, DiveStartDtg, DiveEndDtg, ";
$sql .= "DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS StartEpoch, ";
$sql .= "DateDiff(\"ss\", '01/01/70', DiveEndDtg) AS EndEpoch ";
$sql .= " FROM Dive\n";
$sql .= " WHERE ";
$sql .= " ExpeditionID_FK = " . $ExpdID;
$sql .= " ORDER BY DiveNumber";

Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: \n$sql");
$RSdives = $Conn->Execute($sql);
if(!$RSdives) {
 	die "<BR>Empty Return Set: RSdives. for <br>$sql ";
}

$expdDiveList = '';
my @diveNumsInExpd = ();
while ( !$RSdives->EOF ) { 	
	my $diveName = $RSdives->Fields('RovName')->value . $RSdives->Fields('DiveNumber')->value;

	if ( $RSt->Fields('ExpeditionID_FK')->value eq 0 ) {
		$sDtl{$diveName} = 600 * ($RSdives->Fields('StartEpoch')->value - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs);
		$eDtl{$diveName} = 600 * ($RSdives->Fields('EndEpoch')->value - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs) - $sDtl{$diveName};
	}
	elsif ( $RSt->Fields('ExpeditionID_FK')->value ne 0 && $RSt->Fields('RovName')->value eq 'mini' ) {
		$sDtl{$diveName} = 600 * ($RSdives->Fields('StartEpoch')->value - $DiveStartEsecs) / ( $EDiveEndEsecs - $DiveStartEsecs);
		$eDtl{$diveName} = 600 * ($RSdives->Fields('EndEpoch')->value - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs) - $sDtl{$diveName};
	}
	elsif ( $RSt->Fields('ExpeditionID_FK')->value ne 0 && $RSt->Fields('RovName')->value ne 'mini' ) {
		$sDtl{$diveName} = 600 * ($RSdives->Fields('StartEpoch')->value - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs);
		$eDtl{$diveName} = 600 * ($RSdives->Fields('EndEpoch')->value - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $sDtl{$diveName};
	}

	$sDes{$RSdives->Fields('DiveNumber')->value} = $RSdives->Fields('StartEpoch')->value;
	$eDes{$RSdives->Fields('DiveNumber')->value} = $RSdives->Fields('EndEpoch')->value;
	#######print"<br>sDtl-DN: $sDtl{$diveName}, <br>eDtl-DN: $eDtl{$diveName}, <br>sDes: $sDes, <br>eDes: $eDes<br> ";

	if ( $RSdives->Fields('DiveNumber')->value != GetFormValue('DiveNumber') ) {
		$expdDiveList .= "<a href=\"$this_script?RovName=" . $RSdives->Fields('RovName')->value;
		$expdDiveList .= "&DiveNumber=" . $RSdives->Fields('DiveNumber')->value . "&edit=yes\">$diveName</a>";
		$expdDiveList .= " &#183; ";
	}
	push @diveNumsInExpd, $RSdives->Fields('DiveNumber')->value;
	$RSdives->MoveNext;
}
$expdDiveList =~ s/ &#183; $//;

#######print"<br>rovName: $rovName<br>";

###########################################################################
# Get times from Pilot's database table for the dives in this Expedition
###########################################################################
Win32::ASP::DebugPrint("\nstep3():\nPSQL: rovName =  $rovName\n");
# Check for MiniROV on non-MBARI Ships because data entry is different.
	if ( $ExpdID eq 0 && $rovName eq 'mini') {	
		$psql = "SELECT DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS DiveStartEpoch, ";
		$psql .= "DateDiff(\"ss\", '01/01/70', DiveEndDtg) AS DiveEndEpoch, DiveNumber ";
		$psql .= " FROM Dive\n";
		$psql .= " WHERE ";
		$psql .= "DeviceID = 1859 AND ExpeditionID_FK = " . $ExpdID . " AND RovName = '". $rovName . "' AND ";
		$psql .= " DiveStartDtg between '" . $RSdives->Fields('DiveStartDtg')->value . "' AND '" . $RSdives->Fields('DiveEndDtg')->value . "' AND ";
		$psql .= " DiveEndDtg BETWEEN '" . $RSdives->Fields('DiveStartDtg')->value . "' AND '" . $RSdives->Fields('DiveEndDtg')->value . "' ";
		$psql .= " ORDER BY DiveNumber ";
		#######print"<br>PSQL2 = $psql<br>";
	}
	# Check for MiniROV on MBARI Ships because data entry is different.
	elsif ( $ExpdID ne 0 && $rovName eq 'mini') {
		$psql = "SELECT DiveNumber, RovName, DeviceID, DiveStartDtg, DiveEndDtg, ExpeditionID_FK, \n";
		$psql .= "DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS StartEpoch, DateDiff(\"ss\", '01/01/70', DiveEndDtg) AS EndEpoch ";
		$psql .= " FROM Dive\n";
		$psql .= " WHERE DeviceID = 1859\n";
		$psql .= " AND DiveStartDtg BETWEEN '" . $RSdives->Fields('DiveStartDtg')->value . "' AND '" . $RSdives->Fields('DiveEndDtg')->value . "' AND ";
		$psql .= " DiveEndDtg BETWEEN '" . $RSdives->Fields('DiveStartDtg')->value . "' AND '" . $RSdives->Fields('DiveEndDtg')->value . "' ";
		$psql .= " ORDER BY DiveNumber ";
		#######print"<b>MiniROV on MBARI Ships<br>";
	}
	#Check for Non MiniROV Everything else
	elsif ( $ExpdID ne 0 && $rovName ne 'mini') {
		$psql = "SELECT DateDiff(\"ss\", '01/01/70', gmtStart) AS PilotStartEpoch, ";
		$psql .= "DateDiff(\"ss\", '01/01/70', gmtEnd) AS PilotEndEpoch, DiveNumber, Results ";
		if ($rovName eq 'vnta' || $RSdives->Fields('RovName')->Value eq 'vnta') {
			$psql .= "  FROM dbo.VentanaPilotsDive\n";
		}
		elsif ($rovName eq 'tibr' || $RSdives->Fields('RovName')->Value eq 'tibr') {
			$psql .= " FROM dbo.TiburonPilotsDive\n";
		}
		elsif ($rovName eq 'docr' || $RSdives->Fields('RovName')->Value eq 'docr') {
			$psql .= " FROM dbo.DocRickettsPilotsDive\n";
		}
		$psql .= " WHERE ";
		$psql .= " gmtStart between '" . $RSexpd->Fields('ScheduledStartDtg')->value . "' AND '" . $RSexpd->Fields('ScheduledEndDtg')->value . "' AND ";
		$psql .= " gmtEnd between '" . $RSexpd->Fields('ScheduledStartDtg')->value . "' AND '" . $RSexpd->Fields('ScheduledEndDtg')->value . "' ";
		$psql .= " ORDER BY DiveNumber ";
		#######print"<br>PSQL3 = $psql<br>";
	}

Win32::ASP::DebugPrint("\nstep3():\nExecuting PSQL: \n$psql");
$PRSdives = $Conn->Execute($psql);
if(!$PRSdives) {
 	die "<h2>Error: Could not read Pilot's Database</h2>Empty Return Set: PRSdives. for <br>$psql ";
}

@diveNumsInPilot = ();
while ( !$PRSdives->EOF ) {

	my $diveName = $rovName . $PRSdives->Fields('DiveNumber')->value;

    if ( $ExpdID eq 0 && $rovName eq 'mini') {	
 		#######print"<br>NON-MBARI MINI DIVE<br>";
		$DiveStartEsecs = $PRSdives->Fields('DiveStartEpoch')->value;
		$DiveEndEsecs = $PRSdives->Fields('DiveEndEpoch')->value;

		$sPDtl{$diveName} = 600 * ($PRSdives->Fields('StartEpoch')->value - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs);
		$ePDtl{$diveName} = 600 * ($PRSdives->Fields('EndEpoch')->value - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs) - $sDtl{$diveName};
		#######print"<br>Pilot Start and End times for MBARI MiniROV...$sPDtl{$diveName}, $ePDtl{$diveName}<br>";
		#######print"<br>sDtl-DN: $sDtl{$diveName}, <br>eDtl-DN: $eDtl{$diveName}, <br>sDes: $sDes, <br>eDes: $eDes<br> ";

		# Build timeline hash lists from CTD data

		getCTDTimes($rovName, $DiveStartEsecs, $DiveEndEsecs, $RSdives->Fields('DiveStartDtg')->Value, $RSdives->Fields('DiveEndDtg')->Value);

	$PRSdives->MoveNext;
	}
	if ( $ExpdID ne 0 && $rovName ne 'mini' ) {
  			#######print"<br>MBARI NON-MINIROV DIVE2<br>";
			$PilotStartEsecs = $PRSdives->Fields('PilotStartEpoch')->value;
			$PilotEndEsecs = $PRSdives->Fields('PilotEndEpoch')->value;
			$PilotResults = (defined $PRSdives->Fields('Results')) ? $PRSdives->Fields('Results')->value : '';
			$PilotResultsHash{$diveName} = $PilotResults;
	
			$sPDtl{$diveName} = 600 * ($PilotStartEsecs - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs);
			$ePDtl{$diveName} = 600 * ($PilotEndEsecs - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $sPDtl{$diveName};
	
			push @diveNumsInPilot, $PRSdives->Fields('DiveNumber')->value;

			# Build timeline hash lists from CTD data
			getCTDTimes($rovName, $ExpdStartEsecs, $ExpdEndEsecs, $RSexpd->Fields('ScheduledStartDtg')->Value, $RSexpd->Fields('ScheduledEndDtg')->Value);

	$PRSdives->MoveNext;
	}
  	if ( $ExpdID ne 0 && $rovName eq 'mini' ) {
			#######print"<br>MBARI MINI DIVE<br>";
			$DiveStartEsecs = $PRSdives->Fields('StartEpoch')->value;
			$DiveEndEsecs = $PRSdives->Fields('EndEpoch')->value;
			#######print"<br>DiveStartEpochsecs: $DiveStartEsecs, DiveEndEpochsecs: $DiveEndEsecs<br>";

			$sPDtl{$diveName} = 600 * ($PRSdives->Fields('StartEpoch')->value - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs);
			$ePDtl{$diveName} = 600 * ($PRSdives->Fields('EndEpoch')->value - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs) - $sDtl{$diveName};
			#######print"<br>Pilot Starts and End times for MBARI MiniROV...$sPDtl{$diveName}, $ePDtl{$diveName}<br>";
			#######print"<br>sDtl-DN: $sDtl{$diveName}, <br>eDtl-DN: $eDtl{$diveName}, <br>sDes: $sDes, <br>eDes: $eDes<br> ";

			$DStartEpoch = $RSexpd->Fields('DiveStartDtg')->value; 
			$DEndEpoch = $RSexpd->Fields('DiveEndDtg')->value; 
			#######print"<br>$rovName, $ExpdStartEsecs, $ExpdEndEsecs, $StartEpoch, $EndEpoch<br>";

			# Build timeline hash lists from CTD data

			getCTDTimes($rovName, $ExpdStartEsecs, $ExpdEndEsecs, $DStartEpoch, $DEndEpoch);
			####getCTDTimes($rovName, $ExpdStartEsecs, $ExpdEndEsecs, $RSexpd->Fields('ScheduledStartEpoch')->Value, $RSexpd->Fields('ScheduledEndEpoch')->Value);		

		$PRSdives->MoveNext;
  	}	
} ## End of While loop
$PRSdives->close;


#################################################################
# Set information for form entry values for next dive addition
#################################################################
$diveIndex = 1;
undef $diveNumToAdd;
foreach my $dn (@diveNumsInPilot) {
	$notAdded = 1;
	Win32::ASP::DebugPrint("\ndiveNumToAdd:\ndn = $dn");
	foreach my $d (@diveNumsInExpd) {
		Win32::ASP::DebugPrint("\ndiveNumToAdd:\nd = $d");
		Win32::ASP::DebugPrint("\ndiveIndex = $diveIndex\n rovName.d = $rovName$d\neDes{d} = $eDes{$d}\n");
		Win32::ASP::DebugPrint("\neCTDes{diveIndex} = $eCTDes{$diveIndex}\n");
		if ($dn == $d) {
			$notAdded = 0;
			$diveIndex++;
			Win32::ASP::DebugPrint("\ndiveNumToAdd:\ndiveIndex = $diveIndex");
			## Put in this code for messy dives: inserted for T946 and commented out for T950-T951
			##if ( $eDes{$d} == $eCTDes{$diveIndex} ) {	# Increment diveIndex if we skipped CTD data dive
			##	Win32::ASP::DebugPrint("\nIncrementing diveIndex as CTD dive data was skipped\n");
			##	$diveIndex++;	
			##}
			last;
		}
		else {
			$notAdded = 1;
		}
	}
	if ($notAdded) {
		$diveNumToAdd = $dn;
		#######print"<br>$diveNumToAdd<br>";
		last;
	}
}



###########################################################
# Set initial form field values for the dive to be added 
###########################################################
$initialDiveNumber = $passedDiveNumber || $diveNumToAdd;
$initialStartEpoch = $sCTDes{$diveIndex} || $StartEpoch;
$initialEndEpoch = $eCTDes{$diveIndex} || $EndEpoch;
Win32::ASP::DebugPrint("\ninitial values:\ninitialStartEpoch = $initialStartEpoch, diveIndex = $diveIndex\n");


#######print"<br>initialStartEpoch = $initialStartEpoch, diveIndex = $diveIndex, EXPD = " . $RSexpd->Fields('ExpeditionID')->Value . "<br>";

###############################################################################
# Construct month, day, year option lists for start & end entry
###############################################################################
if ($RSdives->Fields('ExpeditionID_FK')->Value eq 0 || $RSexpd->Fields('ExpeditionID')->Value eq 0) {
	#######print"<br>NON-MBARI SHIP DIVE<br>";
	($Smonth_option_list, $Sday_option_list, $Syear_option_list, $Shr_option_list, $Smin_option_list) = construct_option_lists($DiveStartEpoch);

	($Emonth_option_list, $Eday_option_list, $Eyear_option_list, $Ehr_option_list, $Emin_option_list) = construct_option_lists($DiveEndEpoch);
}
else {
	#######print"<br>MBARI SHIP DIVE<br>";
	($Smonth_option_list, $Sday_option_list, $Syear_option_list, $Shr_option_list, $Smin_option_list) = construct_option_lists($StartEpoch);

	($Emonth_option_list, $Eday_option_list, $Eyear_option_list, $Ehr_option_list, $Emin_option_list) = construct_option_lists($EndEpoch);
}
%>
<table width="680" border="3">
<tr><td>
<table width="815">
        <tr>
                <td colspan="2" bgcolor="silver" align="center">
                <font face="helvetica,Ariel"><b>Dive Information</b></font></td>
        </tr>
</table>
<!-- EDIT DIVE INFORMATION FORM PAGE HERE _ JUST DIVE PAGE NOT EXPD -->
<!--<table class="a" border="3">
<tr><td>
<table width="815">
    <tr>
        <td colspan="2" bgcolor="silver" align="center">
           <font face="helvetica,Ariel"><b>Dive Information</b></font>
        </td>
    </tr>
</table>-->

<script Language="JavaScript"><!--
<%=$diveCheckDateFunctions%>

function Field_Validator(theForm) {
	
	<% if ( ! $Session->{'RegisteredUser'} ) { %>
		if (theForm.email_addr.value == "") {
			alert("Please enter your MBARI email address.");
			theForm.email_addr.focus();
			return (false);
		}	
	<% } %> 

	if (theForm.DiveNumber.value == "") {
		alert("Please enter a Dive Number.");
		theForm.DiveNumber.focus();
		return (false);
	}
	
	if (theForm.EDayOfWeek.value == "") {		// Only gets set if times are fiddled with
		alert("Please enter Dive times.  \nCannot have a dive set to the same time as the Expedition");
		return (false);
	}

  	if ( theForm.SEpochSecs.value > theForm.EEpochSecs.value ) {
		alert("Start time is after end time.  This can't be right.");
		theForm.SEpochSecs.focus();
		return (false);
	}
	
	if ( (theForm.EEpochSecs.value - theForm.SEpochSecs.value) > 86400 ) {
		alert("Dive duration is greater than 24 hours. \nThis is unlikely. Please double check the times.");
		return (false);
	}
	
	//alert("Debugging, not submitted.")
	return (true)	// make true for production!
	
} // End Field_Validator(theForm)

// Force calculation of form fields
calcDate(document.forms[0])
  //-->
</script>

<form method="post" action="<%=$this_script%>" id="form1" name="diveForm" onsubmit="return Field_Validator(this)">
<input type="hidden" name="updateID" value="<%= GetFormValue('edit')%>">
<input type="hidden" name="deleteID" value="<%= GetFormValue('edit')%>">
<input type="hidden" name="ExpdStartEsecs" value="<%=$ExpdStartEsecs%>">
<input type="hidden" name="ExpdEndEsecs" value="<%=$ExpdEndEsecs%>">

<table border=0 cellpadding="0" cellspacing="0">
<!--<table class="a" border=0 cellpadding="0" cellspacing="0">-->
<% if ( ! $Session->{'RegisteredUser'} ) { %>
<tr><td valign="top"><%= $font_1%><b>Your email:</b></font></td>
<td valign="top">
<%= $font_1%><input type="text"  size="8" name="email_addr" value="<%=GetFormValue('email_addr')%>" onKeyPress="return checkEnter(event)">@mbari.org</font></td>
</tr>
<% }
#
# Look up associated data for this dive and parse for logr.gz file url
#
my $logrFile, $logrLink;
my $dlinks = dataLinks( ExpeditionID_FK => $ExpdID, DataType => 'rovctd' );
foreach ( split('</a>', $dlinks) ) {
	/.*(http:.*logr\.dat\.gz).*/;
	$logrLink = $+;
	next unless $logrLink;
	next if $oldLogrLink eq $logrLink;
	push @logrLinks, $logrLink;
	##$Response->Write("<br>$logrLink");
	$logrFile = $logrLink;
	$logrFile =~ s#http://mww.mbari.org/ARCHIVE/rovctd/.+/\d\d\d\d/\d\d\d/# #;		# Shorten the name (old path)
	$logrFile =~ s#http://mww.mbari.org/ARCHIVE/logger/.+/# #;				# Shorten the name (new path)
	push @logrFiles, $logrFile;
	$oldLogrLink = $logrLink;
}

foreach $f ( @{$Application->{'DiveFields'}} ) { 
    if ( $f eq 'ExpeditionID_FK' ) { %>
		<tr><td valign="top"><%= $font_1%><br><b><%= $f%>:</b></font></td>
		<td valign="top"><%= $font_1%><br><a href="postcruise.asp?step=3&ExpeditionID=<%= $ExpdID%>&edit=<%= $ExpdID%>" target="expd">
		<%= $ExpdID%></a> (Click to view this dive's Expedition information in another window)</font></td>
    <%}
    elsif ( $f eq 'RovName' ) {
		my $checked_vnta = '';
		my $checked_tibr = '';
		my $checked_docr = '';
		my $checked_mini = '';
		$checked_vnta = 'checked' if $RSexpd->Fields('ShipName')->Value eq 'ptlo';
		$checked_vnta = 'checked' if $RSexpd->Fields('ShipName')->Value eq 'rcsn';
		$checked_tibr = 'checked' if ($RSexpd->Fields('ShipName')->Value eq 'wfly' && $ExpdStartEsecs < 1230768000);	# before 1/1/2009
		$checked_docr = 'checked' if ($RSexpd->Fields('ShipName')->Value eq 'wfly' && $ExpdStartEsecs > 1230768000);	# after 1/1/2009
		#$checked_mini = 'checked' if $RSexpd->Fields('ShipName')->Value eq 'rcsn' && $RSdives->Fields('DeviceID')->Value eq 1859 );
		##$checked_mini = 'checked' if ($RSexpd->Fields('ShipName')->Value eq 'othr' && $RSdives->Fields('DeviceID')->Value eq 1859 || $RSexpd->Fields('ShipName')->Value eq 'bhrzn' && $RSdives->Fields('DeviceID')->Value eq 1859 || $qRovName eq 'mini');		
		$checked_mini = 'checked' if ($RSexpd->Fields('ShipName')->Value eq 'othr' && $RSdives->Fields('DeviceID')->Value eq 1859 || $qRovName eq 'mini');			
		%>
		<tr><td valign="center"><%= $font_1%><br><b><%= $f%>:</b></td>
		<td valign="top"><br><%= $font_1%>
		<input type="radio" name="RovName" value="vnta" <%= $checked_vnta%>>Ventana&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="radio" name="RovName" value="tibr" <%= $checked_tibr%>>Tiburon&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="radio" name="RovName" value="docr" <%= $checked_docr%>>Doc Ricketts&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="radio" name="RovName" value="mini" <%= $checked_mini%>>MiniROV		
		<br><br></font></td>
    <% }
    elsif ( $f eq 'DiveNumber' ) {%>
		<tr><td valign="center"><%= $font_1%><b><%= $f%>:</b></td>
		<td valign="top"><%= $font_1%><input type="text"  size=5 name="DiveNumber" value="<%= $initialDiveNumber%>" onKeyPress="return checkEnter(event)">
		<% if ( @logrLinks ) { 
			$Response->Write(" <font color=brown>Parse for possible dive number(s): </font>");
		}
		foreach ( @logrLinks ) { 
			$logrFile = shift @logrFiles; %>
			<a href="javascript:OpenData('/expd/queries/greplogr.asp?url=<%= $_%>&string=DIVE_NUM.*NEW','logrparse')"><%= $logrFile%></a> <b>&#183;</b> 
		<% } %>
		</font></td>
    <% }
    elsif ( $f eq 'DiveStartDtg' ) { %>
		<tr><td colspan=2>
		<br> <%	
		if ($dlinks) { %>
			<font class="td timesText" color=brown size="3"><b>Use these data links to help determine the dive's start and end times:</b>
			<br>Use GIF plot to view dive profile(s) and Derived link to grab start & end times.  Dive start and end times correspond to the ROV's entire time in the water,
			since various science products, such as CTD data, are collected during that whole period.</font>
			<div class="wrap2"><%= $font_1%><br><b>Possible ROVCTD data:</b><div class="wrap2"> <%= $dlinks%>
			</font></div><br><%
		} %></div>
		</td></tr>

		<tr><td valign="top"><%= $font_1%><b>Dive Start:</b></font></td>
		<td valign="top">

		<table><tr><td colspan="2">		


		<font color=brown>Please enter dive start time: </font>
		</td></tr>
		<tr><td width="21"><%= $font_1%>
		<input type="radio" name="Stime" onFocus="document.diveForm.Etime[0].checked=true;" value="epoch" checked>
		</font>
		</td><td><%= $font_1%>
		<input type="text" size="12" STYLE="font-family:monospace" name="SEpochSecs" value="<%= $initialStartEpoch%>" onKeyPress="return checkEnter(event)" onChange="calcDate(this.form)" onFocus="document.diveForm.Stime[0].checked=true;document.diveForm.Etime[0].checked=true;"> (epoch secs - <font color="brown">beginning of good ROVCTD data, use Derived <i>ddd</i> (csv) link above</font>)  
		</font></td></tr>
		
		<% my $cell = "<select name=\"${f}Month\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Stime[1].checked=true;document.diveForm.Etime[1].checked=true;\">$Smonth_option_list</select>\n";
		$cell .= "<select name=\"${f}Day\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Stime[1].checked=true;document.diveForm.Etime[1].checked=true;\">$Sday_option_list</select>\n";
		$cell .= "<select name=\"${f}Year\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Stime[1].checked=true;document.diveForm.Etime[1].checked=true;\">$Syear_option_list</select>\n";
		$cell .= "<select name=\"${f}Hr\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Stime[1].checked=true;document.diveForm.Etime[1].checked=true;\">$Shr_option_list</select>\n";
		$cell .= ":<select name=\"${f}Min\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Stime[1].checked=true;document.diveForm.Etime[1].checked=true;\">$Smin_option_list</select>\n";
			    ##$cell .= "<input type=\"text\" name=\"${f}Time\" size=\"4\" value=\"$Shrmn\" maxlength=\"4\" onFocus=\"document.diveForm.Stime[1].checked=true;\">\n";
		$cell .= "(US/Pacific Time <font color=\"brown\">*</font>)\n\n";%>
		<tr><td><%= $font_1%>
		<input type="radio" name="Stime" onFocus="document.diveForm.Etime[1].checked=true;" value="localdtg">
		</font>
		</td><td><%= $font_1%>
		<input type="text" size="3" STYLE="font-family:monospace" name="SDayOfWeek" onKeyPress="return checkEnter(event)">
		<%= $cell %>
		</font></td></tr>
		<tr><td><%= $font_1%>&nbsp;
		</font>
		</td><td><%= $font_1%>
		<input type="text" size="32" STYLE="font-family:monospace" name="SGMTString" value="(enter time in above fields)" onKeyPress="return checkEnter(event)">
		Year day: <input type="text" size="3" STYLE="font-family:monospace" name="SYearDay" value="" onKeyPress="return checkEnter(event)">
		</font></td></tr>
		</table>
		</td> <%
	}  # End elsif ( $f eq 'DiveStartDtg' ...
	elsif ( $f eq 'DiveEndDtg' ) { %>
		<tr><td valign="top"><%= $font_1%><b>Dive End:</b></font></td>
		<td valign="top">

		<table><tr><td colspan=2>

		<font color=brown>Please enter dive end time: </font>	
		</td></tr>
		<tr><%= $font_1%>
		<input type="radio" name="Etime" onFocus="document.diveForm.Stime[0].checked=true;" value="epoch" checked>
		</font>
		<%= $font_1%>
		<input type="text"  size="10" STYLE="font-family:monospace" name="EEpochSecs" value="<%= $initialEndEpoch%>" onKeyPress="return checkEnter(event)" onChange="calcDate(this.form)" onFocus="document.diveForm.Etime[0].checked=true;document.diveForm.Stime[0].checked=true;"> (epoch secs - <font color="brown">end of good ROVCTD data, use Derived <i>ddd</i> (csv) link above</font>)  
		</font></td></tr>
		<%
		my $cell = "<select name=\"${f}Month\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Etime[1].checked=true;document.diveForm.Stime[1].checked=true;\">$Emonth_option_list</select>\n";
		$cell .= "<select name=\"${f}Day\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Etime[1].checked=true;document.diveForm.Stime[1].checked=true;\">$Eday_option_list</select>\n";
		$cell .= "<select name=\"${f}Year\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Etime[1].checked=true;document.diveForm.Stime[1].checked=true;\">$Eyear_option_list</select>\n";
		$cell .= "<select name=\"${f}Hr\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Etime[1].checked=true;document.diveForm.Stime[1].checked=true;\">$Ehr_option_list</select>\n";
		$cell .= ":<select name=\"${f}Min\" size=\"1\" STYLE=\"font-family:monospace\" onChange=\"calcDate(this.form)\" onFocus=\"document.diveForm.Etime[1].checked=true;document.diveForm.Stime[1].checked=true;\">$Emin_option_list</select>\n";
		##$cell .= "<input type=\"text\" name=\"${f}Time\" size=\"4\" value=\"$Ehrmn\" maxlength=\"4\" onFocus=\"document.diveForm.Etime[1].checked=true;\">\n";
		$cell .= "(US/Pacific Time <font color=\"brown\">*</font>)\n\n";%>
		<tr><td><%= $font_1%>
		<input type="radio" name="Etime" onFocus="document.diveForm.Stime[1].checked=true;" value="localdtg">
		</font>
		</td><td align="left"><%= $font_1%>
		<input type="text" size="3" STYLE="font-family:monospace" name="EDayOfWeek" onKeyPress="return checkEnter(event)">
		<%= $cell %>
		</font></td>
		</tr>
		<tr><td><%= $font_1%>&nbsp;
		</font>
		</td><td align="left"><%= $font_1%>
		<input type="text" size="32" STYLE="font-family:monospace" name="EGMTString" value="(enter time in above fields)" onKeyPress="return checkEnter(event)">
		Year day: <input type="text" size="3" STYLE="font-family:monospace" name="EYearDay" value="" onKeyPress="return checkEnter(event)">
		</font></td></tr>

		<tr><td colspan="2"><%= $font_1%>
		Dive duration: <input type="text" size="5" STYLE="font-family:monospace" name="DiveTimeHours" onKeyPress="return checkEnter(event)"> hours
		<br>
		<font color="brown">
		<% if ( $expdDiveList eq '' ) {%>
		The brown bar will adjust as you modify above times
		<% }
		else { %>
		Colored bars show dive timeline (<font color="cyan">Cyan</font> bars are dives that have been entered into the Expedition database), the brown bar will adjust as you modify above times (dives must not overlap).
		<% } %>
		<br><br></font>
		</font></td></tr>
		
		<tr><td colspan="2">
		
		<img src="sp_wh.gif" width="" height="3" name="StartDiveTimeLine"><img src="sp_brown.gif" width="" height="3" name="EndDiveTimeLine">
		<br>
		<img src="sp_wh.gif" width="" height="3" name="StartDiveLabel"><%= $font_2%>dive to be added: <%=$initialDiveNumber%></font>
		</td></tr>
		
		<% foreach my $d ( sort keys %sDtl ) { %>
			<tr><td colspan="2">
			
			<img src="sp_wh.gif" width="<%=$sDtl{$d}%>" height="3"><img src="sp_cyan.gif" width="<%=$eDtl{$d}%>" height="3" alt="<%=$d%>">
			<br>
			<img src="sp_wh.gif" width="<%=$sDtl{$d}%>" height="3" alt="<%=$d%>"><%= $font_2%><%=$d%></font>
			</td></tr>
		<% } 
		
		#
		# Loop though the pilot's dive(s) that have not been assigned yet to this expedition
		#
		foreach my $d ( sort keys %sPDtl ) { 
			next if ( $sDtl{$d} && $eDtl{$d} );
			 %>
			<tr><td colspan="2" nowrap>
			<img src="sp_wh.gif" width="<%=$sPDtl{$d}%>" height="3" hspace="0" vspace="0"><img src="sp_grey.gif" width="<%=$ePDtl{$d}%>" height="3" alt="Pilot <%=$d%>" hspace="0" vspace="0">
			<br>
			<img src="sp_wh.gif" width="<%=$sPDtl{$d}%>" height="3" alt="Pilot <%=$d%>" hspace="0" vspace="0"><%= $font_2%>Pilot's <%=$d%></font>
			</td></tr>
			<%
		}
		
		#
		# Display timeline of ROVCTD data
		#
		sub numerically {
				$a <=> $b;
		}
		foreach my $d ( sort numerically keys %sCTDtl ) { 
			Win32::ASP::DebugPrint("\nstep3():\nadd_dive_form(): \nd = $d, s = $sCTDtl{$d}, e = $eCTDtl{$d}\n");
			 %>
			<tr><td colspan="2" nowrap>
			<img src="sp_wh.gif" width="<%=$sCTDtl{$d}%>" height="3" hspace="0" vspace="0"><img src="sp_lightgrey.gif" width="<%=$eCTDtl{$d}%>" height="3" alt="<%=$d%>" hspace="0" vspace="0">
			<%
			if ($d == 1) { %>
			<br>
			<img src="sp_wh.gif" width="<%=$sCTDtl{$d}%>" height="3" alt="CTD data" hspace="0" vspace="0"><%= $font_2%>CTD data</font>
			<% } %>
			</td></tr>
			<%
		}%>
		
		<tr><td colspan="2">
		<img src="sp_grey.gif" width="600" height="3" name="ExpdTimeLine" alt="expd timeline">
		</td></tr>
			<% # Put GMT year-day boundaries on expedition timeline

			$YYYY = POSIX::strftime("%Y", POSIX::gmtime($ExpdStartEsecs) );
			$leap = ( $YYYY % 4 == 0 ) ? 1 : 0;
			$leap = ( ! $YYYY % 100 == 0 || $YYYY % 400 == 0 ) ? 1 : 0;
			$leap = ( $YYYY % 400 == 0 ) ? 1 : 0;
			my %days = (0=>31, 1=> $leap ? 29 : 28, 2=>31, 3=>30, 4=>31, 5=>30, 6=>31, 7=>31, 8=>30, 9=>31, 10=>30, 11=>31);
			my %iWid = (0=>6, 1=>5, 2=>6, 3=>6, 4=>6, 5=>6, 6=>6, 7=>6, 8=>6, 9=>6);
			my %iNam = (0=>'zero.gif', 1=>'one.gif', 2=>'two.gif', 3=>'three.gif', 4=>'four.gif', 5=>'five.gif', 6=>'six.gif', 7=>'seven.gif', 8=>'eight.gif', 9=>'nine.gif');
			$sDDD = POSIX::strftime("%j", POSIX::gmtime($ExpdStartEsecs) );
			$eDDD = POSIX::strftime("%j", POSIX::gmtime($ExpdEndEsecs) );
			$mo = POSIX::strftime("%m", POSIX::gmtime($ExpdStartEsecs) ) - 1;
			$da = POSIX::strftime("%d", POSIX::gmtime($ExpdStartEsecs) );
			
			my $dayB = 0;
			my $barHtml = '';
			my $txtHtml = '';
			my $accumDayB = 0;
			my $accumDayBlabel = 0;
			my $char3width = 10;
			foreach my $ddd ($sDDD..$eDDD) { 
				###############################
				# Create image for year-day text
				################################
				my $wid = 0;
				my $dddImg = '';
				foreach my $d ( split(//, $ddd) ) {
					$wid += $iWid{$d};
					$dddImg .= "<img src=\"$iNam{$d}\">";
				}
				#
				# Calculate where to put year-day Boundary
				#
				$dayBesecs = timegm(0, 0, 0, $da, $mo, $YYYY);
				Win32::ASP::DebugPrint("\nddd = $ddd, d= $da, mo = $mo, YYYY = $YYYY, dayBesecs = $dayBesecs\nwid = $wid\n");
				$dayB = 600 * ( $dayBesecs - $ExpdStartEsecs ) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $accumDayB;
				$dayBlabel = 600 * ( $dayBesecs - $ExpdStartEsecs ) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $accumDayBlabel;

				if ( $dayB < 0 ) {
					$dayB = 0;
					$wid = 0;
					$Bwid = 0;
				}
				else {
					$barHtml .= "<img src=\"sp_wh.gif\" width=\"$dayB\" height=\"1\" hspace=\"0\" vspace=\"0\"><img src=\"sp_grey.gif\" width=\"2\" height=\"20\" alt=\"Year-day boundary\">$dddImg";
					$Bwid = 2;
				}
				$txtHtml .= "<img src=\"sp_wh.gif\" width=\"$dayBlabel\" height=\"1\" hspace=\"0\" vspace=\"0\">${font_2}${ddd}</font>";
				$accumDayB = $accumDayB + $dayB + $wid + $Bwid;
				$accumDayBlabel = $accumDayBlabel + $dayB + $wid;
				$da = ($days{$mo} == $da) ? 1 : $da + 1;
				$mo++ if $da == 1;
			} 
			%>
			<tr><td colspan="2" valign="top" nowrap>
			<%=$barHtml%>
			</td></tr>
		
		
		<tr><td colspan="2">
		<% if ( $expdDiveList ne '' ) {%>
		<br><%= $font_1%>Edit Other dives from this expedition: <%=$expdDiveList%></font><br><br>
		<% } %>
		</td></tr>
		
		</table>
		</td> 
		<%
	} # End elsif ( $f eq 'DiveEndDtg' ...
    elsif ( $f eq 'DiveChiefScientist' ) { %>
		<tr><td valign="top"><%= $font_1%><b><%= $f%>:</b></font></td>
		<td valign="top">

		<!--<table class="a"><tr><td colspan=2>-->

			<%= $font_1%><select name="<%=$f%>" size="1">
		<option value=""></option>
		<% while ( !$RScs->EOF ) {
			$fullname = $RScs->Fields('FirstName')->value . " " . $RScs->Fields('LastName')->value;
				
			my $cso_list = "<option value=\"$fullname\"";
			$cso_list .= "selected" if $RSexpd->Fields('ExpdChiefScientist')->Value eq $fullname;
			$cso_list .= ">$fullname</option>\n";
			$Response->write($cso_list);
			$RScs->MoveNext;
		 } %>
  		</select></font><br><br>
  	<!--</table>-->
  </td></tr>
    <%}
    elsif ( $f eq 'BriefAccomplishments' ) { %>
    <tr><td valign="top">
		<!--<table class="a"><tr><td colspan=2>-->
    	<%= $font_1%><b>Dive Accomplishments and Comments:</b>
    <br><font size="-2">(If different from Expedition Accomplishments. For data coments too.)</font>
    </font></td>
    <td valign="top"><%= $font_1%><textarea rows="5" name="<%=$f%>" cols="90" wrap><%=$PilotResultsHash{$rovName.$initialDiveNumber}%></textarea>
    </font><br>  <%
	    	if ( $PilotResultsHash{$rovName.$initialDiveNumber} ) { %>
	    		<%= $font_1%><font color="brown">
			Above text inserted from Pilot's database.  Alter to include science accomplishments too.
			</font></font>
	    	<%
	    	}	
	%>
    <!--</table>-->
</td>
    <% }
    else { %>
    <tr><td valign="top"><%= $font_1%><b><%= $f%>:</b></td>
	<td valign="top"><input type="text" size="50" name="<%= $f%>" value=""></font></td>
    <% } # End if ( $f eq ...%>
   </tr>
<% }  %> 

</table>
<br>
<input type="hidden" name="logistics_email" value="<%= $Session->{'logistics_email'}%>">
<input type="hidden" name="ShipName" value="<%= $RSexpd->Fields('ShipName')->Value%>">
<input type="hidden" name="ShipSeqNum" value="<%= $RSexpd->Fields('ShipSeqNum')->Value%>">
<input type="hidden" value="<%=GetFormValue('ExpeditionID')%>" name="ExpeditionID">
<input type="hidden" value="<%=GetFormValue('ExpeditionID')%>" name="ExpeditionID_FK">
<center>
<input type="submit" class="button green" value="Submit This and Add Another Dive to This Expedition" name="add_dive"><br><br>
<input type="submit" class="button blue" value="Submit This Dive" name="insert"><br><br>
<input type="reset" class="button gray" value="Undo Form Changes" name="reset">
</CENTER>


</form>

<%
if ( $Session->{'RegisteredUser'} ) {

}
else { %>
	You will get a blind copy of the mail that is sent to <%=$Application->{'logistics_email'}%>.
	<br>
<%
}
%>
<br>
<font class="td smalltimesText" color="brown">* Local time is what is in the pilot's dive log.  You can not enter GMT time on this form.  
It is displayed in a text box on this form, but it cannot be edited.  The preferred method of entering time is
in epoch seconds as read from the ROVCTD data file.  This accurately defines the start and end of the dive based
on when we logged good CTD data. Note: The time calculations on this form use your local computer's time zone setting.  
Your computer's time zone needs to be set to "(GMT-08:00) US/Pacific" with daylight savings time adjustment for this 
all to work right.</font><br><br>
<hr width="75%" align="center">
<!--#include file="postcruise_ftr_karen.inc"-->
</td></tr>
</table>

<%
$RSdives->close;
$RS->Close;
$RSt->Close;
$RSexpd->Close;
$RScs->Close;
$Conn->Close;
}	# End add_dive_form()
%>



<%
#--------------------------------------------------------------------
#

=head3 insert_dive()

add a dive to the dive table using stored procedure that
prevents duplicate entries.

Mike McCann

Date Created: 11/20/98

=cut

sub insert_dive {
	#################################################################
	# Construct the GMT start and End times for the database update
	# Save things to be updated in the database as session variables
	#################################################################
	Win32::ASP::DebugPrint("\ninsert_dive():\nStime= " . GetFormValue('Stime') . "\n");
	if ( GetFormValue('Stime') eq 'localdtg' ) {
		@sl = (	0, GetFormValue('DiveStartDtgMin'), GetFormValue('DiveStartDtgHr'),
				GetFormValue('DiveStartDtgDay'), GetFormValue('DiveStartDtgMonth')-1, GetFormValue('DiveStartDtgYear')-1900 );
		$Session->{'DiveStartDtg'} = toGMT(@sl);
		Win32::ASP::DebugPrint("\ninsert_dive():\nSession->{'DiveStartDtg'} = $Session->{'DiveStartDtg'}\n");
		$Session->{'StartEpoch'} = timelocal(@sl);		# For original check
		$Session->{'DiveStartDtgLocal'} = GetFormValue('DiveStartDtgYear') . '-' .
					GetFormValue('DiveStartDtgMonth') . '-' .
					GetFormValue('DiveStartDtgDay') . ' ' . 
					GetFormValue('DiveStartDtgTime');
	}
	elsif ( GetFormValue('Stime') eq 'epoch' ) {
		$Session->{'StartEpoch'} = GetFormValue('SEpochSecs');
		Win32::ASP::DebugPrint("\nStartEpoch = " . $Session->{'StartEpoch'} );
		$Session->{'DiveStartDtg'} = toGMTfromES($Session->{'StartEpoch'});
	}
	
	if ( GetFormValue('Etime') eq 'localdtg' ) {
		@el = (	0, GetFormValue('DiveEndDtgMin'), GetFormValue('DiveEndDtgHr'),
				GetFormValue('DiveEndDtgDay'), GetFormValue('DiveEndDtgMonth')-1, GetFormValue('DiveEndDtgYear')-1900 );
				
		$Session->{'DiveEndDtg'} = toGMT(@el);
		$Session->{'EndEpoch'} = timelocal(@el);		# For original check
		$Session->{'DiveEndDtgLocal'} = GetFormValue('DiveEndDtgYear') . '-' .
					GetFormValue('DiveEndDtgMonth') . '-' .
					GetFormValue('DiveEndDtgDay') . ' ' . 
					GetFormValue('DiveEndDtgHr') .
					GetFormValue('DiveEndDtgMin');
	}
	elsif ( GetFormValue('Etime') eq 'epoch' ) {
		$Session->{'EndEpoch'} = GetFormValue('EEpochSecs');
		Win32::ASP::DebugPrint("\nEndEpoch = " . $Session->{'EndEpoch'} );
		$Session->{'DiveEndDtg'} = toGMTfromES($Session->{'EndEpoch'});
	}

#
# Open the database to insert fields
#
open_database($dsn);		# Creates $Conn object as a global variable

#
# Use stored procedure, depends on order....
#
# @RovName varchar(4) = '', @DiveNumber int = 0, @ExpeditionID_FK int = null,
# /* Date and time are not inserted by the postcruise...
# @DiveStartDtg datetime = '', @DiveEndDtg datetime = '',
# @DiveChiefScientist varchar(50) = null, @BriefAccomplishments varchar(255) = null
#


$sql = "exec insertDive ";
$sql .= FixString( GetFormValue('RovName'), 'text') . ",";
$sql .= FixString( GetFormValue('DiveNumber'), 'num') . ",";
$sql .= FixString( GetFormValue('ExpeditionID_FK'), 'num') . ",";
$sql .= FixString( $Session->{'DiveStartDtg'}, 'text') . ",";
$sql .= FixString( $Session->{'DiveEndDtg'}, 'text') . ",";
$sql .= FixString( GetFormValue('DiveChiefScientist'), 'text') . ",";
$sql .= FixString( GetFormValue('BriefAccomplishments'), 'text');

Win32::ASP::DebugPrint("\ninsert_dive():\nSQL = $sql ");
$RS = $Conn->Execute($sql);

$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("Database Load Error(s):<pre>$sql</pre>");
	foreach $error (keys %$Errors)
        {
      		$Response->Write($error->{Description});
   	}
}
else { %>
	<h3>Dive <%= GetFormValue('DiveNumber')%> has been added to Expedition <%= GetFormValue('ExpeditionID_FK')%></h3>
	<a href="postcruise_karen.asp?step=2&Conjunction=AND&search=advanced&Continue=Search+for+expeditions&qDiveNumber=<%= GetFormValue('DiveNumber')%>&qRovName=<%= GetFormValue('RovName')%>">
	Search database for dive number <%= GetFormValue('DiveNumber')%></a>
	<br>
	
	<%
	$shrtRovName = GetFormValue('RovName');
	$diveNumber = GetFormValue('DiveNumber');
	# Add a dive record to the TapeSummary table
	$sql =<<EOS;
INSERT INTO [EXPD].[dbo].[TapeSummary]
           ([DiveNumber]
           ,[RovName]
           )
           values
           ($diveNumber
           ,'$shrtRovName'
           )
EOS

	Win32::ASP::DebugPrint("\ninsert_dive():\nSQL = $sql ");
	$RS = $Conn->Execute($sql);
	
	$Errors = $Conn->Errors();
	if ( keys %$Errors ) { 
		$Response->Write("Database Load Error(s):<pre>$sql</pre>");
		foreach $error (keys %$Errors)
	        {
	      		$Response->Write($error->{Description});
	   	}
	 }
	 else {
	 %>
	 	A blank record for this dive has been added to the TapeSummary table. 
	<%
	 }	
}

if (defined $RS) { 
	$RS->Close;
	}

$Conn->Close;

 } # End insert_dive()
%>

<%
#--------------------------------------------------------------------
#

=head3 email_dive_insert()

Process the form field entries and mail to the Logistics
coordinator.

Mike McCann

Date Created: 11/5/99

=cut

sub email_dive_insert {

	#
	# Use libnet package to send mail message.
	#
	use Net::SMTP;
	use Text::Wrap;
	
	
	#
	# Use IIS LOGON_USER to "know" who is using this & to provide email address
	#
	$logonEmail = $Request->ServerVariables('LOGON_USER')->{Item};
	$logonEmail =~ s/.+[\/\\]//;
	Win32::ASP::DebugPrint("\nlogonEmail = $logonEmail\n");
	$logonEmail .= "\@mbari.org";
	Win32::ASP::DebugPrint("\nlogonEmail = $logonEmail\n");

	$formEmail = GetFormValue('email_addr') . "\@mbari.org" if GetFormValue('email_addr');
	
	#
	# Check for valid email address
	#
	$smtp = Net::SMTP->new('mail.shore.mbari.org'); 	# connect to an SMTP server
	$senderEmail ='';
	Win32::ASP::DebugPrint("\n$formEmail, verify = " . $smtp->verify($formEmail) . "\n");
	$senderEmail = $formEmail if $smtp->verify($formEmail);
	Win32::ASP::DebugPrint("\n$logonEmail, verify = " . $smtp->verify($logonEmail) . "\n");
	$senderEmail = $logonEmail if ( $smtp->verify($logonEmail) && $senderEmail eq '');
	if ( ! $smtp->verify($senderEmail) ) {%>
		<h3>Failed to send email because senderEmail is not valid.</h3>
		<%
		return;
	} 
	
	
	$smtp->mail( $senderEmail );     	# use the sender's address here
	$smtp->to( $Session->{'logistics_email'}, $senderEmail );       # recipient's address
	$smtp->data();                      # Start the mail
	
	# Send the header.
	#
	$smtp->datasend("To: $Session->{'logistics_email'}\n");
	
	$smtp->datasend("From: " . $senderEmail . "\n");
	$smtp->datasend("Subject: Request: Postcruise Dive " .  GetFormValue('RovName') . GetFormValue('DiveNumber') . " (Web form $this_script)\n");
	$smtp->datasend("\n");
	
	# Send the body.
	#
	my $mail_msg =  '';
	$mail_msg .= "Postcruise database load request from " . $senderEmail . "\n";
	
	Win32::ASP::DebugPrint("\n_[0] = $_[0]\n");
	if ($_[0] eq 'update' ) {
		$mail_msg .= "Request to update a dive on ExpeditionID " . GetFormValue('ExpeditionID') . "\n\n";
	}
	else {
		$mail_msg .= "Request to add a dive to ExpeditionID " . GetFormValue('ExpeditionID') . "\n\n";
	}

	$mail_msg .= "============================================\n\n";
	$mail_msg .= "Fields to be added:\n";
	$mail_msg .= "-------------------\n\n";
	
	$Text::Wrap::columns = 60;
	
	foreach $f ( @{$Application->{'DiveFields'}} ) {
	   	$mail_msg .= "$f:\n";
	   	if ($f eq 'DiveStartDtg') {
   			$mail_msg .= "    Use this >>>  Start epoch seconds = " . GetFormValue('SEpochSecs') . "\n";
   		
   			$mail_msg .= "    Local Month = " . GetFormValue('DiveStartDtgMonth') . "\n";
   			$mail_msg .= "    Local Day   = " . GetFormValue('DiveStartDtgDay') . "\n";
   			$mail_msg .= "    Local Year  = " . GetFormValue('DiveStartDtgYear') . "\n";
   			$mail_msg .= "    Local Hour  = " . GetFormValue('DiveStartDtgHr') . "\n";
   			$mail_msg .= "    Local Min   = " . GetFormValue('DiveStartDtgMin') . "\n";
   			$mail_msg .= "    GMT Time    = " . GetFormValue('SGMTString') . "\n";
	   	}
	   	elsif ($f eq 'DiveEndDtg') {
   			$mail_msg .= "    Use this >>>  End epoch seconds = " . GetFormValue('EEpochSecs') . "\n";

   			$mail_msg .= "    Local Month = " . GetFormValue('DiveEndDtgMonth') . "\n";
   			$mail_msg .= "    Local Day   = " . GetFormValue('DiveEndDtgDay') . "\n";
   			$mail_msg .= "    Local Year  = " . GetFormValue('DiveEndDtgYear') . "\n";
   			$mail_msg .= "    Local Hour  = " . GetFormValue('DiveEndDtgHr') . "\n";
   			$mail_msg .= "    Local Min   = " . GetFormValue('DiveEndDtgMin') . "\n";
   			$mail_msg .= "    GMT Time    = " . GetFormValue('EGMTString') . "\n";
	   	}
	   	else {
	   		$mail_msg .= "    " . GetFormValue($f) . "\n";
	   	}
	   	$mail_msg .= "\n";
	}
	
	if ($_[0] eq 'update' ) {
		$mail_msg .= "Click below to edit this dive. Copy above fields into the web form\n";
		$baseurl = $Application->{'BaseUrl'};
		$hyperlink = "$baseurl/dive_karen.asp?RovName=" . GetFormValue('RovName') . "&DiveNumber=" . GetFormValue('DiveNumber') .
			"&edit=yes";
	}
	else {
		$mail_msg .= "Click below to edit this postcruise and add the dive. Copy above fields into the web form\n";
		$baseurl = $Application->{'BaseUrl'};
		$hyperlink = "$baseurl/dive_karen.asp?ExpeditionID=" . GetFormValue('ExpeditionID') . 
			"&passedDiveNumber=" . GetFormValue('DiveNumber') .
			"&DiveStartEpoch=" . GetFormValue('SEpochSecs') . "&DiveEndEpoch=" . GetFormValue('EEpochSecs') .
			"&add_dive=Continue+to+enter+new+dive+information";
	}
	
	$mail_msg .= "For the Logistics Coordinator:\n";
	$mail_msg .= "-----------------------------:\n";
	$mail_msg .= "\n" . $hyperlink . "\n\n";
	
	
	$smtp->datasend($mail_msg) || Win32::ASP::DebugPrint("\nFailed to send mail\n");
	
	$smtp->dataend();                   # Finish sending the mail
	$smtp->quit;                        # Close the SMTP connection
	%>

<h3>The email below has been sent from <%=$senderEmail%> to <%= $Session->{'logistics_email'}%></h3>
Subscribe to the 'postcruise' alias to get notification of your cruise
being entered into the database.
<pre>
<%=$mail_msg%>
</pre>

<%
}	# End email_dive_insert()
%>


<%
#--------------------------------------------------------------------
#

=head3 getCTDTimes()

Return 2 hash lists of pseudo dives based on gaps in the CTD data.  To be used for the timeline display.
Pass in Expedition start and end DTGs.

Mike McCann

Date Created: 11/18/05

=cut

sub getCTDTimes {

	local ($rovName, $ExpdStartEsecs, $ExpdEndEsecs, $eStartDTG, $eEndDTG) = @_;
	my $gapCrit = 18 * 60; 		# 18 minutes
	my $durationCrit = 10 * 60;	# 10 minutes
	

	#######print"<br>$rovName, $ExpdStartEsecs, $ExpdEndEsecs, $eStartDTG, $eEndDTG<br>";
	
	
	my $sql =  " SELECT DatetimeGMT, DateDiff(\"ss\", '01/01/70', DatetimeGMT) AS Epoch ";
	if ( $rovName eq 'vnta' ) {
		$sql .= " FROM  dbo.VentanaRovctdData ";
	}
	elsif ( $rovName eq 'tibr' ) {
		$sql .= " FROM  dbo.TiburonRovctdData ";
	}
	elsif ( $rovName eq 'docr' ) {
		$sql .= " FROM  dbo.DocRickettsRovctdData ";
	}
	elsif ( $rovName eq 'mini' ) {
		$sql .= " FROM  dbo.MinirovRovctdData ";
	}
	$sql .= " WHERE (DatetimeGMT BETWEEN '" . $eStartDTG . "' AND '" . $eEndDTG . "')";
	$sql .= " ORDER BY DatetimeGMT";

	#######print"<br>getCTDTimes: <br>$sql<br>";
	
	Win32::ASP::DebugPrint("\ngetCTDTimes():\nExecuting SQL: \n$sql");
	$RSctd = $Conn->Execute($sql);
	if(!$RSctd) {
	 	die "<h2>Error: Did not read any data from the ___RovctdData view</h2>Empty Return Set: RSctd. for <br>$sql ";
	}
	
	my %sCTDes0 = ();
	my %eCTDes0 = ();
	my %sCTDtl0 = ();
	my %eCTDtl0 = ();
	$lastEs = 0;
	$dataSegment = 0;
	while ( !$RSctd->EOF ) { 
		$es = $RSctd->Fields('Epoch')->value;
		
		if (($es - $lastEs) > $gapCrit) {
		
			$dataSegment++;
			
			$sCTDes0{$dataSegment} = $es;
			$eCTDes0{$dataSegment-1} = $lastEs if $dataSegment > 1;
			
			if ( $RS->Fields('ExpeditionID_FK') eq 0 ) {
				$sCTDtl0{$dataSegment} = 600 * ($es - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs);
				$eCTDtl0{$dataSegment-1} = 600 * ($lastEs - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs) - $sCTDtl0{$dataSegment-1} if $dataSegment > 1;
			}
			elsif ( $RS->Fields('ExpeditionID_FK') ne 0 ) {	
				$sCTDtl0{$dataSegment} = 600 * ($es - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs);
				$eCTDtl0{$dataSegment-1} = 600 * ($lastEs - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $sCTDtl0{$dataSegment-1} if $dataSegment > 1;
			}
		}
		$lastEs = $es;
		$RSctd->MoveNext;
	} # End while()

	if ( $RS->Fields('ExpeditionID_FK') eq 0 ) {
		$eCTDtl0{$dataSegment} = 600 * ($lastEs - $DiveStartEsecs) / ( $DiveEndEsecs - $DiveStartEsecs) - $sCTDtl0{$dataSegment};
	}
	elsif ( $RS->Fields('ExpeditionID_FK') ne 0 ) {	
		$eCTDtl0{$dataSegment} = 600 * ($lastEs - $ExpdStartEsecs) / ( $ExpdEndEsecs - $ExpdStartEsecs) - $sCTDtl0{$dataSegment};
	}
	$eCTDes0{$dataSegment} = $lastEs;
	
	##############################################################################
	# Do a final check and reject dives that are not at least $durationCrit long
	##############################################################################
	%sCTDtl = ();
	%eCTDtl = ();
	my $seg = 0;
	foreach my $k ( sort keys %sCTDes0 ) {
		Win32::ASP::DebugPrint("\ngetCTDTimes():\nk = $k \nsCTDes0{k}, eCTDes0{k} = $sCTDes0{$k}, $eCTDes0{$k} \n");
		if ( ($eCTDes0{$k} - $sCTDes0{$k}) > $durationCrit ) {
			$seg++;
			$sCTDes{$seg} = $sCTDes0{$k};
			$eCTDes{$seg} = $eCTDes0{$k};
			
			$sCTDtl{$seg} = $sCTDtl0{$k};
			$eCTDtl{$seg} = $eCTDtl0{$k};
			Win32::ASP::DebugPrint("\ngetCTDTimes():\nExceeds durationCrit = $durationCrit: sCTDtl{seg}, eCTDtl{seg} = $sCTDtl{$seg}, $eCTDtl{$seg} \n");
		}
	}
	
	return (%sCTDtl, %eCTDtl);

}	# getCTDTimes()
%>
