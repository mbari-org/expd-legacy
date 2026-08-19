<%@ LANGUAGE = PerlScript %>

<!--#include file="expd_functions.inc"-->
<!--#include file="precruise_hdr.inc"-->

<% 
=head1 NAME

precruise.asp - Application for submitting precruises

=head1 SYNOPSIS

    http://expd.mbari.org/expd/log/precruise.asp

=head1 DESCRIPTION

Precruise processing Active Server Page.
Collect info for planed expedition and insert new record to
the Expedition table.

Mike McCann MBARI

November 1998
December 2001

=head1 FUNCTIONS

=cut


# Warning: Don't use CGI, it freezes IIS when a new object is created.
use Win32::ASP;
use OLE;
use Time::Local;
use POSIX;

$this_script = 'precruise.asp';

init_AppVars();

$Session->{'logistics_email'} = $Application->{'logistics_email'};
##$Session->{'logistics_email'} = 'mccann@mbari.org';			# For debugging
$dsn = $Application->{'DSN'};

$Response->Write($Application->{'Warning'});
$Response->Write($Application->{'MoreInfo'});
##$Session->{'DBadministrator'} = 1;					# For debugging.  Make sure to comment out when done!
##if ( $Session->{'DBadministrator'} ) {
##	$Response->Write("Admin = " . $Session->{'DBadministrator'});		# For debugging
##}
##$Response->Write("logistics_email = " . $Session->{'logistics_email'});


if ( GetFormValue('step') eq '1' || ! GetFormValue('step') ) {
	%><h2>Pre-cruise entry: Step 1</h2>
	<P align=left><STRONG><FONT color=#408080>Precruise reports should be filed at 
least 10 working days in advance of your cruise. Use this form for entering a 
new precruise. Click &quot;Continue&quot; button when complete and follow 
further instructions. To edit an existing precruise, <A 
href="http://expd.mbari.org/expd/log/postcruise.asp?search=advanced">search 
for the expedition</A> and click on the &quot;Edit Precruise&quot; button. 
Changes are subject to approval by the Logistics Coordinator. See DMO's <A 
href="http://www.mbari.org/dmo/cruise_planning/cruise.htm">cruise planning</A> 
site for more info.</FONT></STRONG></P>
	</td></tr></table>
	 <% 
	step1();
}
elsif ( GetFormValue('step') eq 'getWaypoints') {		# Needs to be before step 2
	%><h2 align="center">Visiting waypoint management system to collect & organize waypoints</h2>
	</td></tr></table><% 
	getWaypoints();
}
elsif ( GetFormValue('step') eq '2') {
	%><h2 align="center">Pre-cruise entry: Step 2</h2>
	</td></tr></table><% 
	step2();
}
elsif ( GetFormValue('step') eq '3') {
	%><h2 align="center">Pre-cruise entry: Checking Values</h2>
	</td></tr></table><% 
	step3();
}
elsif ( GetFormValue('step') eq 'orderwpts') {
	%><h2 align="center">Pre-cruise entry: Order Waypoints</h2>
	</td></tr></table><% 
	step2('orderwpts');
}
elsif ( GetFormValue('step') eq '4') {
	%><h2 align="center">Pre-cruise entry: Finishing Up</h2>
	</td></tr></table><font face="Arial,Helvetica" >
	It is the responsibility of the Chief Scientist to notify Operations 
	of any special arrangements in advance of your cruise, such as computer 
	systems needs, diving activities, participant's dietary restrictions or 
	pertinent medical information. Please submit an e-mail to the Logistics 
	<a href="mailto:<%= $Session->{'logistics_email'}%>">Support Specialist</a>
	describing any special requirements at this time.</font><% 
	if ( GetFormValue('finish') =~ /Load/ ) {
		my $ret = load();
		precruise_mail_out() if $ret == 0;
	}
	else {
		email();
	}
	close_session();		# Forget all that I know
	%>
	<p><a href="<%= $this_script%>">Enter another precruise</a></p><%
}
elsif ( GetFormValue('step') eq 'revise') {
	%><h2 align="center">Revise Precruise <%=GetFormValue('ShipName')%><%=GetFormValue('ShipSeqNum')%></h2>
	</td></tr></table><% 
	revise();
}
elsif ( GetFormValue('step') eq 'update') {
	%><h2 align="center">Results of <%=GetFormValue('ShipName')%><%=GetFormValue('ShipSeqNum')%> revision</h2>
	</td></tr></table><% 
	update();
	precruise_mail_out('REVISED:');
}

else {
	%><h2 align="center">Unknown step <%= GetFormValue('step')%></h2>
	</td></tr></table><% 
}
%>


<!--#include file="precruise_ftr.inc"-->



<%
#--------------------------------------------------------------------
#

=head3 step1()

Present first form for user to fill out.  This is where the 
ship, time, chief scientist, purpose, etc. for cruise is entered.  
Don't go onto step 2 until all these fields are properly filled.

Author: Mike McCann

Date Created: 10/21/98

Redesigned: 9 Feb 1999 (combined with old step2())

=cut

sub step1 {
	
	require 'timelocal.pl';
	
        my @l = localtime(time+864000);	# Add 10 days
	my $month_10 = $l[4] + 1;	# Because range is 0..11
	my $day_10 = $l[3];

	my $lstr = localtime(time+864000);	# Arghhh! get 2 digit yr from above!
	@lstr2 = split('\s+', $lstr);
	my $year_10 = $lstr2[4];

	@month_names = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov",
                  "Dec");

	%MONTH_ARRAY = ('01', 'Jan',     '02', 'Feb',
                  	'03', 'Mar',     '04', 'Apr', 
                  	'05', 'May',     '06', 'Jun', 
                  	'07', 'Jul',     '08', 'Aug', 
                  	'09', 'Sep',     '10', 'Oct', 
                  	'11', 'Nov',     '12', 'Dec');

	
%>



<script Language="JavaScript"><!--
function Field_Validator(theForm)
{
  if ( theForm.ShipName.value == "" )
  {
    alert("Please select a Ship");
    theForm.ShipName.focus();
    return (false);
  }
  
  if (theForm.ExpdPrincipalInvestigator.value == "")
  {
    alert("Please select a Principle Investigator");
    theForm.ExpdPrincipalInvestigator.focus();
    return (false);
  }
  if (theForm.ExpdChiefScientist.value == "")
  {
    alert("Please select a Chief Scientist");
    theForm.ExpdChiefScientist.focus();
    return (false);
  }
  if (theForm.ProjNum.value == "")
  {
    alert("Please enter an MBARI Project Number");
    theForm.ProjNum.focus();
    return (false);
  }
  if (theForm.Purpose.value == "")
  {
    alert("Please enter Cruise Purpose");
    theForm.Purpose.focus();
    return (false);
  }
  if (theForm.PlannedTrackDesc.value == "")
  {
    alert("Please enter Planned Track Description");
    theForm.PlannedTrackDesc.focus();
    return (false);
  }
  if (theForm.Participants.value == "")
  {
    alert("Please enter Participants");
    theForm.Participants.focus();
    return (false);
  }
  if (theForm.EquipmentDesc.value == "")
  {
    alert("Please enter Required Equipment Description");
    theForm.EquipmentDesc.focus();
    return (false);
  }
  
  return (true);
}

function FlyerTime(select)
{
	//alert("index=" + select.selectedIndex + ", value1=" + select.options[1].value);
	if (select.options[select.selectedIndex].value == "wfly") {
		alert("For Western Flyer cruises please confirm Departure and Arrival times with the ship's captain before entering them here.");
		return (true);
	}
	if (select.options[select.selectedIndex].value == "zphr") {
		alert("For Zephyr cruises please confirm Departure and Arrival times with the ship's captain before entering them here.");
		return (true);
	}
	if (select.options[select.selectedIndex].value == "prgn") {
		alert("For Paragon reservations please confirm that your captain is an approved operator.");
		return (true);
	}
	return (false);
}

//-->
</script>

<% 

open_database($dsn);		# Creates $Conn object as a global variable

$sql="SELECT * FROM Person WHERE (DisplayPIPickList = 1) ORDER BY LastName, FirstName";
Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RSpm = $Conn->Execute($sql);

$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("Database Query Error(s): ");
	foreach $error (keys %$Errors)
        {
      		$Response->Write($error->{Description});
   	}
}


if(!$RSpm) {
	$RSpm->Close;
 	die "<BR>Empty Return Set: RSpm. ";
}

$sql="SELECT * FROM Person WHERE (DisplayChiefSciPickList = 1) ORDER BY LastName, FirstName";
##$sql="SELECT * FROM Person";
Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RScs = $Conn->Execute($sql);
if(!$RScs) {
	$RSpm->Close;
	$RScs->Close;
 	die "<BR>Empty Return Set: RScs. ";
}

#
# Construct option list for months
#
Win32::ASP::DebugPrint("\nmonth_10 = $month_10");
$month_option_list = '';
foreach $num (sort keys (%MONTH_ARRAY)) {
   if ( $num == $month_10 ) {
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
   $val = ($num < 10) ? '0'.$num : $num;
   if ( $num eq $day_10 ) {
	$day_option_list .= "<option value=\"$val\" selected>" . 
				$num . "</option>\n";
   }
   else {
	$day_option_list .= "<option value=\"$val\">" . 
				$num . "</option>\n";
   }
}
##Win32::ASP::DebugPrint("day_option_list = \n$day_option_list ");

#
# Construct option list for years
#
$year_option_list = '';

for ($y = 1988; $y <= $year_10 + 1; $y++) {
   if ( $y eq $year_10 ) {
	$year_option_list .= "<option value=\"$y\" selected>" . 
				$y . "</option>\n";
   }
   else {
	$year_option_list .= "<option value=\"$y\">" . 
				$y . "</option>\n";
   }
}
##Win32::ASP::DebugPrint("year_option_list = \n$year_option_list ");

	  

%>


<form method="POST" action="<%=$this_script%>" onsubmit="return Field_Validator(this)" name="Field">
<font face="Arial,Helvetica" size="-1" color="brown">All fields below must be completed.</font>
<table border="3">
<tr><td>
   <input type="hidden" name="step" value="2">
   <table border="0" cellpadding="5" cellspacing="0">
        <tr>
            <td valign="top">
            
            <br><strong><font face="Arial,Helvetica">Ship Name</font></strong></td>
            <td>
            <select name="ShipName" size="1" onChange="return FlyerTime(this)">
	            <option value=""> </option>
	            <option value="prgn">Paragon</option>
	            <option value="wfly">Western Flyer</option>
	            <option value="rcsn">Rachel Carson</option>
            </select>
            
	    </td>
        </tr>
        <tr>
            <td valign="top">
            <strong><font face="Arial,Helvetica">Cruise Departure Date - Time</font></strong>
            </td>

            <td valign="top">
                <select name="dMonth" size="1">
                  <%= $month_option_list%>
                </select> 
                <select name="dDay" size="1">
                  <%= $day_option_list%>
                </select>
	        <select name="dYear" size="1">
                  <%= $year_option_list%>
                </select>
	        <input type="text" name="dTime" size="4" value="0700" maxlength="4"> 
                (Local Moss Landing time)
	   </td>
        </tr>
        <tr>
            <td valign="top">
            <strong><font face="Arial,Helvetica">Cruise Arrival Date - Time</font></strong>
            </td>

            <td valign="top">
                <select name="aMonth" size="1">
                  <%= $month_option_list%>
                </select> 
                <select name="aDay" size="1">
                  <%= $day_option_list%>
                </select>
	        <select name="aYear" size="1">
                  <%= $year_option_list%>
                </select>
	        <input type="text" name="aTime" size="4" value="1630" maxlength="4"> 
                (Local Moss Landing time)
	   </td>

        </tr>
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Principal Investigator</font></strong></td>
            <td>
	    <select name="ExpdPrincipalInvestigator" size="1">
	    <option value=""></option>
	    <% while ( !$RSpm->EOF ) {
		$fullname = $RSpm->Fields('FirstName')->value . " " . $RSpm->Fields('LastName')->value;
		%><option value="<%= $fullname%>"><%= $fullname%></option>
                <%
		$RSpm->MoveNext;
	     } %>
      	</select>
	The person who ship time is charged against
	</td>
        </tr>

	<tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Chief Scientist</font></strong></td>
            <td>
	    <select name="ExpdChiefScientist" size="1">
	    <option value=""></option>
	    <% while ( !$RScs->EOF ) {
		$fullname = $RScs->Fields('FirstName')->value . " " . $RScs->Fields('LastName')->value;
		%><option value="<%= $fullname%>"><%= $fullname%></option>
                <%
		$RScs->MoveNext;
	     } %>
      	</select>
	<a href="person.asp">Add/Edit person list</a>
	</td>
        </tr>
        
        
	<tr>
            <td valign="top"><strong><font face="Arial,Helvetica">MBARI Project Number</font></strong></td>
            <td>
	    <input type="text" name="ProjNum" size="10" value="" maxlength="10">
	    </td>
        </tr>
       
    </table>
    <p><strong><font face="Arial,Helvetica">Cruise Purpose</font></strong> <br>
    <font face="Arial,Helvetica" size="-1" color="brown">Enter objective for the expedition.</font><br>
    <textarea rows="10" name="Purpose" cols="80" wrap></textarea></p>
    

    <p><strong><font face="Arial,Helvetica">Required Equipment Description</font></strong><br>
    <font face="Arial,Helvetica" size="-1" color=brown>1) Please list hazardous materials in the box below.</font><br>
    <font face="Arial,Helvetica" size="-1" color=brown>2) If your cruise will require a winch (MacArtney, small black Dynacon, or borehole winch) please contact <a href="mailto:efitzgerald@mbari.org">Eric Fitzgerald</a> and describe which winch below.</font><br>
    <font face="Arial,Helvetica" size="-1" color=brown>3) List <a href="http://mww.mbari.org/exportregs/items.htm">export-controlled items</a> (ITAR, EAR). For ITAR items include serial number and part number.</font><br>
	<font face="Arial,Helvetica" size="-1" color=brown>4) If the Wave Glider will be launched or recovered, please indicate in the box below. Note that Wave Glider cruise plans are due three weeks prior to activity. This cruise plan will be sent to <a href="mailto:waveglider@mbari.org">waveglider@mbari.org</a>.</font><br>
	<textarea rows="10" name="EquipmentDesc" cols="80" wrap></textarea></p>
    	
    <p><strong><font face="Arial,Helvetica">Participants</font></strong><br>
    <font face="Arial,Helvetica" size="-1" color=brown>
 
1) For Western Flyer cruises please send any special meal requests (vegetarian, etc.) 
    to the <A href="mailto:mitts@mbari.org">Ship's steward</a>.
    <br>
2) The Western Flyer is regulated by the USCG under 33 CFR 104 for security measures. Participants will require a TWIC card for unescorted access to restricted areas of the vessel. Participants who do not have a TWIC card will be required to be escorted and monitored by a TWIC cardholder. 
    <strong>Please indicate with an asterisk the participants who are TWIC cardholders.</strong>
    <br>

	
3) For <em>Western Flyer</em> cruises: The <strong>chief scientist</strong> will need to contact <A href="mailto:tara@mbari.org">Tara Vadas</A> (Toni Mackenzie or Meilina Dalit 
as backup) to make arrangements for non-MBARI participants to obtain Temporary Cruise Participant badges to keep 
for the duration of the cruise.  If a participant must retrieve a badge after hours or at a remote port, the Vessel 
Security Officer (see captain or mates) onboard can also issue badges. Participants must show current photo 
identification (drivers’ license or passport) to obtain a badge. <br>

4) For all vessels: Clearly identify all <strong>foreign nationals</strong> in the <em>Participants</em> box below for export regulation screening 
purposes. The chief scientist must notify his/her division administrator of all foreign participants at least one week 
prior to the start of the cruise to allow time for screening. <em>Western Flyer</em> badges will not be issued until the screening process is completed.

<br></font>

    <textarea rows="8" name="Participants" cols="80" wrap></textarea></p>
    
    <p><strong><font face="Arial,Helvetica">Planned Track Description</font></strong><br>
    <font face="Arial,Helvetica" size="-1" color=brown>1) Enter text for the cruise track (must fit this box), OR <br>2) select CONTINUE for the waypoint database and select your waypoints - they will be inserted in the Planned Track Description.
<br>NOTE: Please indicate if track extends out of Monterey Bay (beyond M2).
</font><br>
    <textarea rows="10" name="PlannedTrackDesc" cols="80" wrap="off"></textarea></p>
   
    <p> </p>
    <input type="submit" value="Continue" name="Step2"> 
    <input type="reset" value="Reset form fields" name="Reset form">
   


</td></tr>
</table>
</form>

<%
$RSpm->Close;
$RScs->Close;
$Conn->Close;

}	# End step1()

#--------------------------------------------------------------------
#

=head3 step2()

Present entries from step1() for the user to view and check before submitting.

Author: Mike McCann

Date Created: 10/21/98

=cut

sub step2 {

	local $flag = $_[0];
	
	#
	# Save all input variables as session variables Convert local time to GMT
	# Calcuations should be good until Jan 1 2038.
	#

	#
	# If coming back from waypoint selection don't try to get these form values
	#
	if ( $flag ne 'orderwpts' ) {
		@dl = (	0, substr(GetFormValue('dTime'),2,2), substr(GetFormValue('dTime'),0,2),
			GetFormValue('dDay'), GetFormValue('dMonth')-1, GetFormValue('dYear')-1900 );
		@al = (	0, substr(GetFormValue('aTime'),2,2), substr(GetFormValue('aTime'),0,2),
			GetFormValue('aDay'), GetFormValue('aMonth')-1, GetFormValue('aYear')-1900 );
	
		$dGMT = toGMT(@dl);
		$aGMT = toGMT(@al);
	
		$Session->{'ShipName'} = GetFormValue('ShipName');
		$Session->{'ScheduledStartDtg'} = $dGMT;
		$Session->{'ScheduledStartDtgLocal'} = GetFormValue('dYear') . '-' .
					GetFormValue('dMonth') . '-' .
					GetFormValue('dDay') . ' ' . 
					GetFormValue('dTime');
	
		$Session->{'ScheduledEndDtg'} = $aGMT;
		$Session->{'ScheduledEndDtgLocal'} = GetFormValue('aYear') . '-' .
					GetFormValue('aMonth') . '-' .
					GetFormValue('aDay') . ' ' . 
					GetFormValue('aTime');
	
		$Session->{'YYYYDDD'} = toYDfromDate(GetFormValue('dMonth'),
						GetFormValue('dDay'),
						GetFormValue('dYear'),
						int(GetFormValue('dTime')/100), 0);
		
		$Session->{'ExpdChiefScientist'} = GetFormValue('ExpdChiefScientist');
		$Session->{'ExpdPrincipalInvestigator'} = GetFormValue('ExpdPrincipalInvestigator');
		$Session->{'ProjNum'} = GetFormValue('ProjNum');
		$Session->{'StatCode'} = 'plnd';
		$Session->{'Purpose'} = GetFormValue('Purpose');
	
		$Session->{'PlannedTrackDesc'} = GetFormValue('PlannedTrackDesc');
		$Session->{'Participants'} = GetFormValue('Participants');
		$Session->{'EquipmentDesc'} = GetFormValue('EquipmentDesc');
		
		%>
	   <script Language="JavaScript"><!--
		function Field_Validator(theForm)
		{
		return(true);
		}
		//-->
	    </script>
	    <%
		
	}
	else {
	%>
	   <script Language="JavaScript"><!--
		function Field_Validator(theForm)
		{
		     //alert("length = " + theForm.length);
		     for (i=0; i<theForm.length; i++) {
		          var e = theForm.elements[i];
		          //alert("type = " + e.type);
		          if ( e.type != "select-one") {
		          	continue;
		          }
		          // Check that something is entered
		          //alert("options length = " + e.options.length);
		          //alert("selectedIndex = " + e.options.selectedIndex);
		          //alert("i="+i+", e.options[e.options.selectedIndex].value="+e.options[e.options.selectedIndex].value);
		          
			  if ( e.options[e.options.selectedIndex].value == "" || e.options[e.options.selectedIndex].value == null )
			  {
				alert("Please choose the order of visiting the waypoints.");
				e.focus();
				return (false);
			  }
			  // Check that numbers are unique
			  for (j=i+1; j<theForm.length; j++) {
			  	var f = theForm.elements[j];
			  	if ( f.type != "select-one") {
		          		continue;
		          	}
			        //alert("i="+i+", j="+j+", e.value="+e.value+", f.value="+f.value);
			  	if ( f.options[f.options.selectedIndex].value == e.options[e.options.selectedIndex].value )
			  	{
			  		alert("Please set unique numbers for the order.");
			    		f.focus();
			    		return (false);
			  	}
			  }
		     }
		  return(true);
		}
		//-->
	     </script>
	<%
	}
	

	#
	# Save step 2 variables for use in email message
	#
	$step2args = '';
	my $tmp = '';
	foreach my $f ( Win32::OLE::in ($Request->QueryString) ) {  
		#
		# For text fields, only show first 5 characters
		#
		if ($f eq 'EquipmentDesc' || $f eq 'Participants' ||
			$f eq 'PlannedTrackDesc' || $f eq 'Purpose' ) {
			$tmp .= "$f=" . substr($Request->QueryString($f)->{Item},0,5) . "&";
		}
		else {
			$tmp .= "$f=" . $Request->QueryString($f)->{Item} . "&";
		}
   	} 
   	$step2args = $tmp;
	$step2args =~ s/\&$//;
	$Session->{'Step2Args'} = URLEncode($step2args);

	#
	# Get next Ship sequence number
	#
	open_database($dsn);		# Creates $Conn object as a global variable

	$sql = "SELECT max(ShipSeqNum) FROM Expedition WHERE (ShipName = '" . $Session->{'ShipName'} . "')";
	Win32::ASP::DebugPrint("\nSQL = $sql ");
	$RS = $Conn->Execute($sql);
	$Session->{'ShipSeqNum'} = $RS->Fields(0)->value + 1;

	$RS->Close;
	$Conn->Close;
%>


<h3>Check these entries...</h3>
<!-- Need to use "GET" for waypoint string creation to work -->
<form method="GET" action="<%= $this_script%>" onsubmit="return Field_Validator(this)" name=order>

<table border="0">
<tr>
  <td valign="top"><font face="Helvetica,Arial" size="-1"><strong>MBARI Project Number:</strong></font></td>
  <td valign="top"><font face="Helvetica,Arial" size="-1"><%= $Session->{ProjNum}%></font>
</tr>
<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
<%
#
# Loop through fields list and print out all session variables that match
#
foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
	next if $f eq 'ExpeditionID';
	last if $f eq 'StartDtg';		# Following fields come from Post Cruise entry
	%>
	<tr>
	
	<% if ( $f =~ /StartDtg$/ && $Session->{$f} ) {	# Add local time %>
	     <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td>
	     <td valign="top"><font face="Helvetica,Arial" size="-1"><%= $Session->{$f}%></font>
	     <font face="Helvetica,Arial" size="-1">GMT (<%= $Session->{'ScheduledStartDtgLocal'}%> local time)</font>
	<% } %>
	<% elsif ( $f =~ /EndDtg$/ && $Session->{$f} ) {	# Add local time %>
	     <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td>
	     <td valign="top"><font face="Helvetica,Arial" size="-1"><%= $Session->{$f}%></font>
	     <font face="Helvetica,Arial" size="-1">GMT (<%= $Session->{'ScheduledEndDtgLocal'}%> local time)</font>
	<% } %>
	<% elsif ( $f =~ /PlannedTrackDesc$/ && $Session->{$f} ) {	# Preseve format %>
	     <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td>
	     <td valign="top"><PRE><%= $Session->{'PlannedTrackDesc'}%></PRE>
	<% }  
	   else { %>
	     <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td>
	     <td valign="top"><font face="Helvetica,Arial" size="-1"><%= $Session->{$f}%></font>
	<% } %>   
	</td>
	</tr>	
<%
}
#
# Need to get PersonID numbers for waypoint.asp
#
open_database($dsn);

# PI
($pFirstName, $pLastName) = split('\s', $Session->{'ExpdPrincipalInvestigator'});
$sql="SELECT PersonID FROM Person WHERE LastName = '$pLastName' AND FirstName = '$pFirstName'";
Win32::ASP::DebugPrint("\nstep2():\nExecuting SQL: $sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
	$RS->Close;
 	die "<BR>Step2(): Empty Return Set: RS. ";
}
while ( !$RS->EOF ) {
	$piID = $RS->Fields(0)->value;
	$RS->MoveNext;
}
$RS->Close;
$Session->{'piID'} = $piID;

# CS
($pFirstName, $pLastName) = split('\s', $Session->{'ExpdChiefScientist'});
$sql="SELECT PersonID FROM Person WHERE LastName = '$pLastName' AND FirstName = '$pFirstName'";
Win32::ASP::DebugPrint("\nstep2():\nExecuting SQL: $sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
	$RS->Close;
 	die "<BR>Step2(): Empty Return Set: RS. ";
}
while ( !$RS->EOF ) {
	$csID = $RS->Fields(0)->value;
	$RS->MoveNext;
}
$RS->Close;
$Session->{'csID'} = $csID;


$Conn->Close;
 
%>
<tr><td><font face="Helvetica,Arial" size="-1" color="#8F8F8F"><b>Ship Sequence Number:</b></font></td>
    <td><font face="Helvetica,Arial" size="-1" color="#8F8F8F"><%= $Session->{'ShipSeqNum'}%></font></td>
</tr>

<tr>
<td colspan="2"><font color="red">
If something is not right here then press your browser's Back button and fix it.
<br>
If the above is O.K. then select your waypoints and finish the pre-cruise.
</font></td>
</tr>
<tr><td colspan="2"><br><h3>Select Waypoints to be visited...</h3><td></tr>

	<%
	if ( $flag eq 'orderwpts' ) {
	
	  %>
		<tr><th align="right">Order of visit</th><th align="left">&nbsp;&nbsp;Waypoint Name</th></tr>
		
		<%
		my @wpts = Win32::OLE::in ($Request->QueryString('selectedWaypoints'));
		##print "#wpts = $#wpts<br>";
		$iWpts = 0;
		if ( $#wpts ) {  	# If more than 2 then demand order
			$wptOrder_list = '<option></option>';
		}
		
		foreach my $s ( sort @wpts ) {
			$iWpts++;
			$wptOrder_list .= "<option value=\"$iWpts\">$iWpts</option>";
			##print "<br>$iWpts: ", $s;	
		}
		##print "wptOrder_list= ", $wptOrder_list, "\n";
		
		foreach my $s ( sort @wpts ) {
			%>
			<tr><td align="right">
			<select name="<%= $s%>" size="1">
	                  <%= $wptOrder_list%>
	                </select> 
	                </td>
			<td>&nbsp;&nbsp;<%=$s%></td></tr><% 
		}%>
		<tr>
		<td valign="top" colspan="2">	
		<a href="waypoint.asp?action=select&piID=<%= $Session->{'piID'}%>&csID=<%= $Session->{'csID'}%>">
		Select different Waypoints</a>		
		</td>
		</tr><%
	}
	else {
	%>
		<tr>
		<td valign="top" colspan="2">	
		<a href="waypoint.asp?action=select&piID=<%= $Session->{'piID'}%>&csID=<%= $Session->{'csID'}%>">
		<img src="compassrose_small.jpg" border="0">Select Waypoints
		</a> (Select from list of PI's and Chief Scientist's waypoints, then return here to finish up.)	
		</td>
		</tr><%
	}
	%>


</table>



<hr align="left" width="50%">


    <input type="hidden" name="step" value="4">
<% if ( $Session->{'DBadministrator'} ) { %>  
	<font face="helvetica,Ariel" size="-1">Send email to 
	<input type="text" name="pre_email" value="precruise" size="9" maxlength="9">
	@mbari.org and </font> <input type="submit" name="finish" value="Load Data Base"> 
<% }
   else { 	# We email the form %>
        <b>Enter your email address to get a copy of this precruise:</b>
        <input type="text"  size=8 name="sender_email_addr" value="<%=GetFormValue('sender_email_addr')%>">@mbari.org
        <br>
	<input type="submit" name="finish" value="Email Form to Logistics Coordinator"> 
<% } %>
</form>

<%


}	# End step2()
%>


<%
#
#/*======================================================================
#	getWaypoints()
#	
#	Description: 
#	     Pass chief scientists & PI to waypoints.asp in order to
#		 have user select a list of waypoints to add to the cruise
#		 plan.
#	     
#	
#	Author: Mike McCann
#	Date Created: 4/16/01
#====================================================================== */
sub getWaypoints {

	%>ExpdChiefScientist = <%= GetFormValue('ExpdChiefScientist')%> <br>
	<%
	if ( ! GetFormValue('ExpdChiefScientist') ) { 
		%><h2>Must select a Chief Scientist before selecting waypoints</h2>
	<%	return;
	}
	%>

<%
} # End getWaypoints
%>

<%
#
#/*======================================================================
#	step3()
#	
#	Description: 
#	     Chance to verify form entries  (I don't think this function
#		 is used anymore.... 11/8/99)
#	     
#	
#	Author: Mike McCann
#	Date Created: 10/21/98
#====================================================================== */
sub step3 {

	#
	# At this point we should have all the name/value pairs
	# properly entered.  These name/values correspond to the
	# column/values in the Expedition table.
	#
	# This should work:  Query Expedtion for the field names,
	# Loop through this list, pick off Session variables that
	# have the same name.




	#
	# Loop through fields list and print out all session variables that match
	#
	%><table border="0">
	<h3>Check the values and submit...</h3><%
	foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
		next if $f =~ /ExpeditionID/;
		last if $f =~ /^StartDtg/;
		%><input type="hidden" name="<%= $f%>" value="<%= $Session->{$f}%>">
		<tr>
		<td valign="top"><font face="Helvetica,Arial" size="-1"><b><%= $f%>:</b></font></td>
		<td valign="top"><font face="Helvetica,Arial" size="-1"><%= $Session->{$f}%></font>
		<% if ( $f =~ /StartDtg$/ && $Session->{$f} ) {	# Add local time %>
	     	   <font face="Helvetica,Arial" size="-1">GMT (<%= $Session->{'ScheduledStartDtgLocal'}%> local time)</font>
		<% } %>

		<% if ( $f =~ /EndDtg$/ && $Session->{$f} ) {	# Add local time %>
	     	   <font face="Helvetica,Arial" size="-1">GMT (<%= $Session->{'ScheduledEndDtgLocal'}%> local time)</font>
		<% } %>
		
		</td>
		</tr><%
	}
	%></table><%
%>

<br>
<font color="red">
If something is not right here then press your browser's Back button and fix it.
</font>
<hr align="left" width="50%">

<p><font face="helvetica,Ariel" size="-1">Send email to 
	<input type="text" name="pre_email" value="precruise" size="9" maxlength="9">
	@mbari.org when finished.</font></p>
	

<%
}	# End step3()
%>

<%
#--------------------------------------------------------------------
#

=head3 load()

Process the form field entries and insert into the database
Here is where we need to map the names we have on the HTML 
forms for all the fields to the names of the fields that are
actaully in the database table.

Author: Mike McCann

Date Created: 10/28/98

=cut


sub load {

#
# Open the database, this time to insert fields
#

open_database($dsn);		# Creates $Conn object as a global variable


#
# Loop through all form names and construct SQL insert string
#
#$sql = "INSERT INTO Expedition\n(";
$sql = "exec insert_expedition ";
#foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
#	next unless $Session->{$f};
#	$sql .= "$f,";
#}
#$sql =~ s/,$//;
#$sql .= ")\nVALUES(";
Win32::ASP::DebugPrint("Fields are:");
foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
	next if ($f eq 'ExpeditionID');
	if ( $f eq 'ShipSeqNum' ) {
		$sql .= FixString( $Session->{$f}, 'num') . ",";
	}
	elsif ( $f eq 'ismodified' ) {
		$sql .= "\'1\',";
	}
	elsif ( $f eq 'PlannedTrackDesc' ) {
		$Session->{$f} .= procWaypoints();
		$sql .= FixString( $Session->{$f}, 'text') . ",";
	}
	else {
		$sql .= FixString( $Session->{$f}, 'text') . ",";
	}
}
$sql =~ s/,$//;
#$sql .= ")";


Win32::ASP::DebugPrint("\nSQL = $sql ");

$RS = $Conn->Execute($sql);
if(!$RS) {
		Win32::ASP::DebugPrint("Empty Return Set: RS. for $sql ");
}
	
$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("<p>Database Load Error(s): ");
	foreach $error (keys %$Errors)
        {
      		$Response->Write($error->{Description});
   	}
   	%>
	<h3>Pre-cruise entry failed! <br> Please contact information applications support.</h3>
	<%
	$RS->Close;
	$Conn->Close;
	return 1;
}
else {

	%>
	<h3>Inserted record to the Expedition table</h3>

	<%
	$RS->Close;
	
	#$sql = "SELECT max(ExpeditionID) FROM Expedition";
	$sql = "SELECT max(ShipSeqNum) FROM Expedition WHERE (ShipName = '" . $Session->{'ShipName'} . "')";
	Win32::ASP::DebugPrint("\nSQL = $sql ");
	$RS = $Conn->Execute($sql);
	
	%>

	<br>ShipSeqNum = <%= $RS->Fields(0)->value%>

	<%
	$RS->Close;
	$Conn->Close;
	return 0;
}

}	# End load()
%>


<%
#--------------------------------------------------------------------
#

=head3 procWaypoints()

Process the waypoint order into a string that can be put into the precruise.

Author: Mike McCann

Date Created: 1/22/02

=cut

sub procWaypoints {

	my @wpts = Win32::OLE::in ($Request->QueryString);
	my %wptVisits = {};
	my $wptString = "";
		 
	foreach my $f ( @wpts ) {
		next if ( $f eq 'step' && $Request->QueryString($f)->Item == 4); 	#Hope that no waypoints are named 'step'
		next if ( $f eq 'finish' && ( $Request->QueryString($f)->Item eq 'Email Form to Logistics Coordinator'
		                         ||   $Request->QueryString($f)->Item eq 'Load Data Base' ) );
		next if ( $f eq 'pre_email' );
		next if ( $f eq 'sender_email_addr' );
		next unless $f;
		print "<br> waypoint=", $f, " order=", $Request->QueryString($f)->Item;	
		$wptVisits{$Request->QueryString($f)->Item} = $f;
	}
	
	$wptString = "\nWaypoints to be visited\r\n";
	$wptString .="=======================\r\n";
	$wptString .= sprintf("%5s %-25s %11s %11s %7s\r\n", "Order", "Waypoint Name", "Latitude", "Longitude", "Depth");
	$wptString .= sprintf("%5s %-25s %11s %11s %7s\r\n", "-----", "-----------------------", "---------", "-----------", "-----");

	#
	# Open DB, Need separate QConn object as load() has one, and load() calls this
	#
	$QConn = CreateObject OLE "ADODB.Connection";
	$QConn->{'Provider'} = "sqloledb";
	$QConn->Open($dsn);
	
	Win32::ASP::AddDeathHook( sub { $QConn->Close } );	# In case asp dies do some cleanup
	
	if(!QConn) {
	        $Errors = $QConn->Errors();
	        print "<B>Error in opening connection to $dsn:</B>";
	        foreach $error (keys %$Errors) {
	                print $error->{Description}, "\n";
	        }
	        die "<BR>Empty Connection object.";
	}	
	
	$Errors = $QConn->Errors();
	if ( keys %$Errors ) { 
		$Response->Write("<p>Database open Error(s): ");
		foreach $error (keys %$Errors)
	        {
	      		$Response->Write($error->{Description});
	   	}
	   	%>
		<h3>Waypoint query failed! <br> Please contact information applications support.</h3>
		<%
		$QConn->Close;
		return 1;
	}
	my $gotWaypoints = 0;
	foreach my $n ( sort keys %wptVisits ) {
		next unless $wptVisits{$n};		# Getting some weird hash that we need to skip.
		$gotWaypoints = 1;
		
		#
		# Get Waypoint info from DB
		#
		
		my $sql = "SELECT WaypointName, Latitude, Longitude, Depth FROM waypoint WHERE ";
		$sql .= "WaypointName = '$wptVisits{$n}'\n";
		Win32::ASP::DebugPrint("\nSQL = $sql ");
		$RS = $QConn->Execute($sql);
		
		$Errors = $QConn->Errors();
		if ( keys %$Errors ) { 
			$Response->Write("<p>Database query Error(s): ");
			foreach $error (keys %$Errors)
		        {
		      		$Response->Write($error->{Description});
		   	}
		}
		
		$wptString .= sprintf("%-5d %-25s %11.5f %11.5f %5.0f m\r\n", $n, $wptVisits{$n}, $RS->Fields('Latitude')->Value, 
					$RS->Fields('Longitude')->Value, $RS->Fields('Depth')->Value);
					
		#
		# Also show the position in DDD MM.MMM format
		#
		my $latDD = int($RS->Fields('Latitude')->Value);
		my $latMM = sprintf("%5.3f", ($RS->Fields('Latitude')->Value - $latDD) * 60);
		my $lonDD = int($RS->Fields('Longitude')->Value);
		my $lonMM = sprintf("%5.3f", ($RS->Fields('Longitude')->Value - $lonDD) * 60);
		
		$wptString .= sprintf("%30s %11s %11s\r\n", 'in DDD MM.MMM format ==>', "$latDD $latMM", 
					"$lonDD $lonMM");
		$RS->Close;
	}
	$QConn->Close;
	
	$wptString = "" unless $gotWaypoints;
	##$Response->Write("<pre>$wptString</pre>\n");
	return($wptString);
	
} # End procWaypoints()
%>


<%
#--------------------------------------------------------------------
#

=head3 email()

Process the form field entries and mail to the Logistics
coordinator.

Author: Mike McCann

Date Created: 11/5/98

=cut

sub email {

#
# Use libnet package to send
# mail message.
#
        use Net::SMTP;
	use Text::Wrap;


    $smtp = Net::SMTP->new('mail.shore.mbari.org'); 	        # connect to an SMTP server
    $smtp->mail( $Session->{'logistics_email'} );     	# use the sender's, for the email, address here
    my $submitter_email = '';
    $submitter_email = GetFormValue('sender_email_addr') . "\@mbari.org" if GetFormValue('sender_email_addr');
    $smtp->to( $Session->{'logistics_email'}, $submitter_email );        # recipient's address
    $smtp->data();                      # Start the mail

    # Send the header.
    #
    $smtp->datasend("To: $Session->{'logistics_email'}\n");
    $smtp->datasend("From: $Session->{'logistics_email'} \n");
    $smtp->datasend("Cc: $submitter_email \n" );
	$smtp->datasend("Subject: Precruise $Session->{'ScheduledStartDtg'} (Web form $this_script)\n");
    $smtp->datasend("\n");

    # Send the body.
    #
    $smtp->datasend("Precruise database load request\n");
	$smtp->datasend("===============================\n\n");

	##$hyperlink = "$Application->{'BaseUrl'}/$this_script?" . $Session->{'Step2Args'};
	$hyperlink = "$Application->{'BaseUrl'}/$this_script?step=1";
	
	$smtp->datasend("MBARI Project Number: " . $Session->{'ProjNum'} . "\n");
	
	$Text::Wrap::columns = 60;
	$smtp->datasend("\nLink to load this expedition into the database:\n" . 
	$hyperlink . "\n\n(Copy from the fields below into the web form fields)\n\n");
	foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
		next if $f eq 'StatCode';
		last if $f eq 'StartDtg';
		if ( $f =~ /ScheduledStartDtg$/ ) {	# E-mail local time
	     		$str1 = "$Session->{'ScheduledStartDtgLocal'} Local Moss Landing time";
	     		$str2 = wrap("", "", $str1);
		}
		elsif ( $f =~ /ScheduledEndDtg$/ ) {	# E-mail local time
	     		$str1 = "$Session->{'ScheduledEndDtgLocal'} Local Moss Landing time";
	     		$str2 = wrap("", "", $str1);
		}
		elsif ( $f =~ /PlannedTrackDesc$/ ) {	# Append waypoints
	     		$str1 = $Session->{'PlannedTrackDesc'};
	     		$str2 = wrap("", "", $str1);
	     		$str2 .= procWaypoints();
		}
		
		else {
			$str1 = $Session->{$f};
			$str2 = wrap("", "", $str1);
		}

		if ( $str2 =~ /\n/ ) {	# Put new line in front 
	   		$smtp->datasend("$f:\n$str2\n");
	   	}
	   	else {
	   		$smtp->datasend("$f: $str2\n");
	   	}
	   	 
	}
	
	
        $smtp->dataend();                   # Finish sending the mail
        $smtp->quit;                        # Close the SMTP connection
	%>

<h3>Email has been sent to <%= $Session->{'logistics_email'}%></h3>
Subscribe to the 'precruise' alias to get notification of your cruise
being entered into the database.

<%
}	# End email()
%>

<%
#--------------------------------------------------------------------
#

=head3 precruise_mail_out()

Process the form field entries and mail to the address
specified, precruise@mbari.org is the default.

Author: Mike McCann

Date Created: 9/2/99

=cut

sub precruise_mail_out {

#
# Use libnet package to send
# mail message.
#
    use Net::SMTP;
    use Text::Wrap;

    my $message = $_[0];		# Catches 'REVISED:' string

    $smtp = Net::SMTP->new('mail.shore.mbari.org'); 	# connect to an SMTP server
    
    $smtp->mail( $Session->{'logistics_email'} );     	# use the sender's address here
    $smtp->to( GetFormValue('pre_email') . "\@mbari.org" );        # recipient's address
    $smtp->data();                      # Start the mail

    # Send the header.
    #
    $smtp->datasend("To: " . GetFormValue('pre_email') . "\@mbari.org\n");
    $smtp->datasend("From: " . $Session->{'logistics_email'} . " (via Web form $this_script)\n");
    $smtp->datasend("Subject: $message Precruise: $Session->{'ExpdChiefScientist'} $Session->{'ScheduledStartDtg'} GMT\n");
    $smtp->datasend("\n");

    # Send the body.
    #
    $smtp->datasend("$message \n") if $message;
    $smtp->datasend(sprintf("%-30s %s\n", "Precruise:", $Session->{'ExpdChiefScientist'}));

	$hyperlink = "$Application->{'BaseUrl'}/$this_script?step=1";
	
	$Text::Wrap::columns = 90;
	my $labl;

	foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
		next if $f eq 'StatCode';
		next if $f eq 'ExpeditionID';
		last if $f eq 'StartDtg';
		if ( $f =~ /ScheduledStartDtg$/ ) {	# E-mail local time
	     		$str1 = "$Session->{'ScheduledStartDtgLocal'} Local Moss Landing time"
		}
		elsif ( $f =~ /ScheduledEndDtg$/ ) {	# E-mail local time
	     		$str1 = "$Session->{'ScheduledEndDtgLocal'} Local Moss Landing time"
		}
		else {
			$str1 = $Session->{$f};
		}
		$str1 =~ s/\n\s+/\n/g;				# Remove leading spaces put in by email to logistics coordinator (arghhh!!!)
		$str2 = wrap("\t", "\t", $str1);
		if ( $str2 =~ /\n/ ) {
	   		$smtp->datasend("$f:\n$str2\n");
	   	}
	   	else {
	   		$labl = $f . ":";
	   		$smtp->datasend(sprintf("%-30s %s\n", $labl, $str1));
	   	}
	   	 
	}
	
	#
	# Send trailing message
	#
	my $trailmsg = "
--------------------
Cruise participants:
";

	$trailmsg .= "
Visiting scientists who will participate on R/V Western Flyer cruises will 
need a TEMPORARY R/V WESTERN FLYER BADGE for the duration of their cruise 
(after positive identification is established). Badges will not have to be 
worn offshore, but will be required while the vessel is docked. These badges 
will only be issued to personnel listed on the precruise information submitted 
by the chief scientist. The badges for external participants will be 
maintained and issued by Tara Smallwood. All badges are numbered and must 
be returned. 
 
also, before going to sea:
" if $Session->{'ShipName'} eq 'wfly';

	$trailmsg .= "
1. Pick up blank videotapes from the video lab (MBARI Building A, room 140).
2. Call video lab (831-775-1829) one week prior to your cruise to schedule 
   training on how to use our videotape recorders and VARS program (used for 
   framegrabbing and sample annotations).
3. After your cruise, return all videotapes to the video lab (room A140).
4. Report all problems related to video, annotation, and samples to the video 
   lab staff IMMEDIATELY.
 
More info at: http://www.mbari.org/dmo/cruise_planning/cruise.htm. 
If you have any questions, please call 831-775-1829. 
";
	$smtp->datasend("$trailmsg");
	
    $smtp->dataend();                   # Finish sending the mail
    $smtp->quit;                        # Close the SMTP connection
%>

<h3>Email has been sent to <%= GetFormValue('pre_email')%>.</h3>

<%
}	# End precruise_mail_out()
%>

<%
#--------------------------------------------------------------------
#

=head3 revise()

Present form for user to fill out to revise an exiting precruise record.

Author: Mike McCann

Date Created: 9/2/99

=cut

sub revise {
	
	require 'timelocal.pl';
	
	if ( $Session->{'DBadministrator'} != 1 ) { 
	%>
		<h2>Precruises can be revised only by Logistics Coordinator</h2>
	<%	return;
	}
		
	#
	# Open database to get specified expedition
	#
	open_database($dsn);

	# Do an explicit select (not select *) to avoid problems...
	$sql = "SELECT " . join( ", ", @{$Application->{'ExpeditionFields'}} ) . " FROM Expedition\n";
	$sql .= "WHERE ";
	$sql .= "ShipName = '" . GetFormValue('ShipName') . "' ";
	$sql .= "AND ShipSeqNum = " . GetFormValue('ShipSeqNum');

	Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: \n$sql");
	$RS = $Conn->Execute($sql);
	if(!$RS) {
	 	die "<BR>Empty Return Set: RS. for <br>$sql ";
	}

	#
	# Save the Expedition field values into an assoc. array.  Need to do this because
	# Access/ODBC forgets memo fields once they've been read.  (Nice, hunh?  This took
	# about 4 hours to discover, I wonder if SQL Server has the same behaviour.)
	# Well one thing that is wierd with MSSQL is that I had to move this code block up
	# here.  It seems that just accessing one of the RS values removes it!  ?!?
	#
	Win32::ASP::DebugPrint("\nstep3(): Saving Expedition Values:\n");


	foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
		$ExpeditionValue{$f} = $RS->Fields($f)->Value;
	    Win32::ASP::DebugPrint("$f: $ExpeditionValue{$f}");
	}
#
# Construct option list for ships
#
$prgn_sel = ($RS->Fields('ShipName')->Value eq "prgn") ? "selected" : "";
$wfly_sel = ($RS->Fields('ShipName')->Value eq "wfly") ? "selected" : "";
$rcsn_sel = ($RS->Fields('ShipName')->Value eq "rcsn") ? "selected" : "";
$ship_option_list = "<option value=\"prgn\"${prgn_sel}>Paragon</option>\n";
$ship_option_list .= "<option value=\"wfly\"${wfly_sel}>Western Flyer</option>\n";
$ship_option_list .= "<option value=\"rcsn\"${rcsn_sel}>Rachel Carson</option>\n";

#
# Get the actual start and end date info save as session vars for check of original
#

$StartEpoch = SelectEpoch("ScheduledStartDtg");
$EndEpoch = SelectEpoch("ScheduledEndDtg");

$Session->{'StartEpoch_orig'} = $StartEpoch;
$Session->{'EndEpoch_orig'} = $EndEpoch;

#
# Construct month, day, year option lists for start & end entry
#
($Smonth_option_list, $Sday_option_list, $Syear_option_list, $Shrmn) = 
	construct_option_lists($StartEpoch);

($Emonth_option_list, $Eday_option_list, $Eyear_option_list, $Ehrmn) = 
	construct_option_lists($EndEpoch);


#
# Get Chief Scientists from Person table
#
$sql="SELECT * FROM Person WHERE (DisplayChiefSciPickList = 1) ORDER BY LastName, FirstName";
Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: $sql");
$RScs = $Conn->Execute($sql);
if(!$RScs) {
	$RScs->Close;
 	die "<BR>Empty Return Set: RScs. ";
}
$cs_option_list = "<option value=\"\"</option>";
while ( !$RScs->EOF ) {
	$fullname = $RScs->Fields('FirstName')->value . " " . $RScs->Fields('LastName')->value;
	$cs_option_list .= "<option value=\"$fullname\"";
	$cs_option_list .= "selected" if $RS->Fields('ExpdChiefScientist')->Value eq $fullname;
	$cs_option_list .= ">$fullname</option>\n";
	$RScs->MoveNext;
}
$RScs->Close;
 
#
# Get Principal Investigators from Person table
#
$sql="SELECT * FROM Person WHERE (DisplayPIPickList = 1 OR DisplayChiefSciPickList = 1) ORDER BY LastName, FirstName";
Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: $sql");
$RSpi = $Conn->Execute($sql);
if(!$RSpi) {
	$RSpi->Close;
 	die "<BR>Empty Return Set: RSpi. ";
}
$pi_option_list = "<option value=\"\"</option>";
while ( !$RSpi->EOF ) {
	$fullname = $RSpi->Fields('FirstName')->value . " " . $RSpi->Fields('LastName')->value;
	$pi_option_list .= "<option value=\"$fullname\"";
	$pi_option_list .= "selected" if $RS->Fields('ExpdPrincipalInvestigator')->Value eq $fullname;
	$pi_option_list .= ">$fullname</option>\n";
	$RSpi->MoveNext;
}
$RSpi->Close;



%>

<font face="Arial,Helvetica" color="Brown">
Change the fields, indicating the revisions and click Continue.
</font>

<script Language="JavaScript"><!--
function Field_Validator(theForm)
{
  if (theForm.ShipName.value == "")
  {
    alert("Please enter a value for the \"Ship Name\" field.");
    theForm.ShipName.focus();
    return (false);
  }
  if (theForm.Purpose.value == "")
  {
    alert("Please enter a value for the \"Cruise Purpose\" field.");
    theForm.Purpose.focus();
    return (false);
  }
  if (theForm.PlannedTrackDesc.value == "")
  {
    alert("Please enter a value for the \"Planned Track Description\" field.");
    theForm.PlannedTrackDesc.focus();
    return (false);
  }
  if (theForm.Participants.value == "")
  {
    alert("Please enter a value for the \"Participants\" field.");
    theForm.Participants.focus();
    return (false);
  }
  if (theForm.EquipmentDesc.value == "")
  {
    alert("Please enter a value for the \"Required Equipment Description\" field.");
    theForm.EquipmentDesc.focus();
    return (false);
  }
  return (true);
}

function FlyerTime(select)
{
	if (select.options[select.selectedIndex].value == "wfly") {
		alert("For Western Flyer cruises please confirm Departure and Arrival times with the ship's captain before entering them here.");
		return (true);
	}
	if (select.options[select.selectedIndex].value == "zphr") {
		alert("For Zephyr cruises please confirm Departure and Arrival times with the ship's captain before entering them here.");
		return (true);
	}
	if (select.options[select.selectedIndex].value == "prgn") {
		alert("For Paragon reservations please confirm that your captain is an approved operator.");
		return (true);
	}
	return (false);
}
//-->
</script>

<%
	  
%>
<form method="POST" action="<%= $this_script%>" onsubmit="return Field_Validator(this)" name="Field">
   
   <table border="0" cellpadding="5" cellspacing="0">
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Ship Name</font></strong></td>
            <td>
            <select name="ShipName" size="1" onChange="return FlyerTime(this)">
	            <%=$ship_option_list%>
            </select>
            </td>
        </tr>
        <tr>
            <td valign="top">
            <strong><font face="Arial,Helvetica">Cruise Departure Date - Time</font></strong>
            </td>

            <td valign="top">
            
                <select name="dMonth" size="1">
                  <%= $Smonth_option_list%>
                </select> 
                <select name="dDay" size="1">
                  <%= $Sday_option_list%>
                </select>
	        <select name="dYear" size="1">
                  <%= $Syear_option_list%>
                </select>
	        <input type="text" name="dTime" size="4" value="<%= $Shrmn%>" maxlength="4"> 
                (Local Moss Landing time)
	   </td>
        </tr>
        <tr>
            <td valign="top">
            <strong><font face="Arial,Helvetica">Cruise Arrival Date - Time</font></strong>
            </td>

            <td valign="top">
                <select name="aMonth" size="1">
                  <%= $Emonth_option_list%>
                </select> 
                <select name="aDay" size="1">
                  <%= $Eday_option_list%>
                </select>
	        <select name="aYear" size="1">
                  <%= $Eyear_option_list%>
                </select>
	        <input type="text" name="aTime" size="4" value="<%= $Ehrmn%>" maxlength="4"> 
                (Local Moss Landing time)
	   </td>

        </tr>
        <tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Principal Investigator</font></strong></td>
            <td>
	    <select name="ExpdPrincipalInvestigator" size="1">
	    <%= $pi_option_list %>
      	</select>
	The person who ship time is charged against
	</td>
        </tr>

	<tr>
            <td valign="top"><strong><font face="Arial,Helvetica">Chief Scientist</font></strong></td>
            <td>
	    <select name="ExpdChiefScientist" size="1">
	    <%= $cs_option_list %>
      	</select>
	<a href="person.asp">Add/Edit person list</a>
	</td>
        </tr>
       
    </table>
	<p><strong><font face="Arial,Helvetica">Cruise Purpose
    </font></strong> (Indicate reason for revision here)<br>
    <textarea rows="3" name="Purpose" cols="80" wrap><%= $ExpeditionValue{Purpose}%></textarea></p>

    <p><strong><font face="Arial,Helvetica">Required Equipment Description</font></strong><br>
    <textarea rows="3" name="EquipmentDesc" cols="80" wrap><%= $ExpeditionValue{EquipmentDesc}%></textarea></p>
    	
    <p><strong><font face="Arial,Helvetica">Participants</font></strong><br>
    <textarea rows="3" name="Participants" cols="80" wrap><%= $ExpeditionValue{Participants}%></textarea></p>
    
	

	<p><strong><font face="Arial,Helvetica">Planned Track Description</font></strong><br>
    <textarea rows="3" name="PlannedTrackDesc" cols="80" wrap><%= $ExpeditionValue{PlannedTrackDesc}%></textarea></p>
    
    <input type="hidden" name="step" value="update">
    <input type="hidden" name="ShipSeqNum" value="<%= GetFormValue(ShipSeqNum)%>">
    
    <font face="helvetica,Ariel" size="-1">Send email to 
	<input type="text" name="pre_email" value="precruise" size="9" maxlength="9">
	@mbari.org and </font>
	
    <p><input type="submit" value="Update database and send mail" name="update"> 
    <input type="reset" value="Reset form fields" name="Reset form"></p>


	

</form>
<%
$RS->Close;
$Conn->Close;

}	# End revise()
%>

<%
#--------------------------------------------------------------------
#

=head3 SelectEpoch()

Select epoch seconds from database field specified.

Author: Mike McCann

Date Created: 10/28/98

=cut

sub SelectEpoch {

	my $time_field = $_[0];

	##$sql = "SELECT DateDiff('s', '01/01/70', $time_field) AS Epoch ";		# -The way Access wants it
	$sql = "SELECT DateDiff(\"ss\", '01/01/70', $time_field) AS Epoch ";
	$sql .= "FROM Expedition\nWHERE ";
	$sql .= "ExpeditionID = " . GetFormValue('edit') . "\n";
	Win32::ASP::DebugPrint("\nSelectEpoch():\nExecuting SQL: \n$sql");
	$RSdate = $Conn->Execute($sql);
	if(!$RSdate) {
		die "<BR>Empty Return Set: RSdate. for <br>$sql ";
	}

	my $epoch = $RSdate->Fields('Epoch')->Value;
	$RSdate->Close;
	
	return($epoch);
}
%>

<%
#--------------------------------------------------------------------
#

=head3 construct_option_lists()

Construct option lists for use in the DTG Month/day/year pull-down
boxes.

Author: Mike McCann

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
my$month = ($mo < 10) ? "0" . $mo : $mo;
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
   $val = ($num < 10) ? '0'.$num : $num;
   if ( $num eq $day ) {
	$day_option_list .= "<option value=\"$val\" selected>" . 
				$num . "</option>\n";
   }
   else {
	$day_option_list .= "<option value=\"$val\">" . 
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
##Win32::ASP::DebugPrint("year_option_list = \n$year_option_list ");

my $hr = ($l[2] < 10 ) ? "0" . $l[2] : $l[2];
my $min = ($l[1] < 10 ) ? "0" . $l[1] : $l[1];
my $hrmn = $hr.$min;


return ($month_option_list, $day_option_list, $year_option_list, $hrmn);

}	# End construct_option_lists()
%>

<%
#--------------------------------------------------------------------
#

=head3 update()

Process the form field entries and update the database record.
Here is where we need to map the names we have on the HTML 
forms for all the fields to the names of the fields that are
actually in the database table.

Author: Mike McCann

Date Created: 9/3/98

=cut

sub update {

#
# Take the date/time fields and construct UTC DTGs & save as session vars
#
@dl = (	0, substr(GetFormValue('dTime'),2,2), substr(GetFormValue('dTime'),0,2),
		GetFormValue('dDay'), GetFormValue('dMonth')-1, GetFormValue('dYear')-1900 );
@al = (	0, substr(GetFormValue('aTime'),2,2), substr(GetFormValue('aTime'),0,2),
	GetFormValue('aDay'), GetFormValue('aMonth')-1, GetFormValue('aYear')-1900 );

$dGMT = toGMT(@dl);
$aGMT = toGMT(@al);

$Session->{'ScheduledStartDtg'} = $dGMT;
$Session->{'ScheduledStartDtgLocal'} = GetFormValue('dYear') . '-' .
			GetFormValue('dMonth') . '-' .
			GetFormValue('dDay') . ' ' . 
			GetFormValue('dTime');

$Session->{'ScheduledEndDtg'} = $aGMT;
$Session->{'ScheduledEndDtgLocal'} = GetFormValue('aYear') . '-' .
			GetFormValue('aMonth') . '-' .
			GetFormValue('aDay') . ' ' . 
			GetFormValue('aTime');
			
$Session->{'ExpdChiefScientist'} = GetFormValue('ExpdChiefScientist');
$Session->{'ExpdPrincipalInvestigator'} = GetFormValue('ExpdPrincipalInvestigator');
$Session->{'StatCode'} = 'plnd';
$Session->{'Purpose'} = GetFormValue('Purpose');


$Session->{'PlannedTrackDesc'} = GetFormValue('PlannedTrackDesc');
$Session->{'Participants'} = GetFormValue('Participants');
$Session->{'EquipmentDesc'} = GetFormValue('EquipmentDesc');
$Session->{'ShipName'} = GetFormValue('ShipName');
$Session->{'ShipSeqNum'} = GetFormValue('ShipSeqNum');

#
# Open the database, this time to update fields
#
open_database($dsn);		# Creates $Conn object as a global variable
 
#
# Test where clause to see that it updates only one record.  Bail if it matches more than 1.
#
$sql = "SELECT count(*) AS cnt \nFROM Expedition\n";
$sql .= "WHERE ( StatCode != 'cmpl' AND ShipName = '";
$sql .= $Session->{ShipName} . "' AND ShipSeqNum = ";
$sql .= $Session->{ShipSeqNum} . ")";

Win32::ASP::DebugPrint("\nSQL = $sql ");

$RS = $Conn->Execute($sql);

$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("Database Load Error(s): ");
	foreach $error (keys %$Errors)
        {
      		$Response->Write($error->{Description});
   	}
}

Win32::ASP::DebugPrint("\ncnt = " . $RS->Fields(cnt)->Value );

if ( $RS->Fields(cnt)->Value != 1 ) {
	%><h2><%= $RS->Fields(cnt)->Value%> records would be updated!</h2>
	Only one record should be updated.  Please contact your database support to have this problem fixed.
	<%
	return;
}
$RS->Close;

#
# Loop through all form names and construct SQL update string, exit without
# complaining if there are no fields to update, an update implies that StatCode
# goes to plnd.
#
$sql = "UPDATE Expedition\nSET ";
$sql .= "\nStatCode = 'plnd',";
my $nfields = 0;
foreach $f ( @{$Application->{'PreUpdateExpeditionFields'}} ) {
	next if $f eq 'ExpeditionID';
	next unless $Session->{$f};
	$sql .= "\n$f = " . FixString( $Session->{$f}, 'text') . ",";
	$nfields++;
}
return unless $nfields;
$sql =~ s/,$//;

$sql .= "WHERE ( StatCode != 'cmpl' AND ShipName = '";
$sql .= $Session->{ShipName} . "' AND ShipSeqNum = ";
$sql .= $Session->{ShipSeqNum} . ")";

Win32::ASP::DebugPrint("\nSQL = $sql ");

$RS = $Conn->Execute($sql);

$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("Database Load Error(s): ");
	foreach $error (keys %$Errors)
        {
      		$Response->Write($error->{Description});
   	}
}


%>
<h3>Updated precruise <%= $Session->{'ShipName'}%><%= $Session->{'ShipSeqNum'}%></h3>

<%

$RS->Close;
$Conn->Close;

}	# End update()
%>

