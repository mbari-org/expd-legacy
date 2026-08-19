<%@ LANGUAGE = PerlScript %>

<!--#include file="postcruise_hdr.inc"-->
<!--#include file="expd_functions.inc"-->

<% 
=head1 NAME

postcruise.asp - Application for searching and entering post-cruise information

=head1 SYNOPSIS

    #http://expd.mbari.org/expd/log/postcruise.asp

    https://mww.mbari.org/cruises/

=head1 DESCRIPTION

Postcruise processing Active Server Page.
Search for existing precruise record and offer editing of
fields that can be filled in during the postcruise process.

Mike McCann MBARI

November 1998

=head1 FUNCTIONS

=cut


# Warning: Don't use CGI, it freezes IIS when a new object is created.
use Win32::ASP;
use OLE;
use Time::Local;
use POSIX;
use Benchmark;

$this_script = 'postcruise.asp';

init_AppVars();

$dsn = $Application->{'DSN'};

$_db_type = 'mssql';

##$tmp1 = $Application->{'Warning'};
##$Response->Write("App vars Warning: ". $tmp1);
##$Response->Write("<p>App vars MoreInfo: ".$Application->{'MoreInfo'});
##$Response->Write("<p>Application vars DSN: ".$Application->{'DSN'});
##$Response->Write("<p>dsn: ".$dsn);
$Session->{'logistics_email'} = GetFormValue('logistics_email') || 
								$Session->{'logistics_email'} || $Application->{'logistics_email'};
$Session->{'RegisteredUser'} = GetFormValue('RegisteredUser') || 
								$Session->{'RegisteredUser'} || $Application->{'RegisteredUser'};
$Session->{'RegisteredUser'} = 0 unless $Session->{'RegisteredUser'} == 1;

Win32::ASP::DebugPrint("\nlogistics_email=" . $Session->{'logistics_email'} . "\n" . "\nRegisteredUser=" . $Session->{'RegisteredUser'} . "\n");
if ( GetFormValue('step') eq '1' || ! GetFormValue('step') ) {
	if (GetFormValue('search') eq 'advanced' ) {
		%>
	
		<h2>Advanced Search of Expedition Database</h2>
		<strong><font color="#408080">Enter any parameters below to restrict search for expeditions, or use <a href="<%=$this_script%>">keyword search</a>.</font></strong><BR><BR>
  
		</td></tr></table><% 
		advanced_search();
	}
	else {
		%><h2>Simple Search of Expedition Database</h2>
		<strong><font color="#408080">Enter a keyword or number to query for
        expeditions or use <a href="<%= $this_script%>?search=advanced">advanced search</a></font></strong>
		</td></tr></table><% 
		step1();
	}
}
elsif ( GetFormValue('step') eq '2') {
	%><h2>Results of Expedition Search</h2>
	<font color="#408080"><strong>Scroll through cruises and find associated images and
    data.&nbsp; <br>
    Submit or edit postcruise reports by clicking on the Edit Postcruise button.</strong></font>
	</td></tr></table><% 
	if ( GetFormValue('data_table') eq 'yes') {
		step2('data_table');
	}
	else {
		step2();
	}
}
elsif ( GetFormValue('step') eq '3' ) {
    %><h2>Edit Expedition <%= GetFormValue('ShipName')%><%= GetFormValue('ShipSeqNum')%></h2>
    <font color="#408080"><strong>Fill in or edit the fields below, following
    the instructions in <span class="browntext">brown</span><font color="#408080">. Then submit
    the form to the Logistics Coordinator.</font></strong></font>
    </td></tr></table><% 
    step3();
}
elsif ( GetFormValue('step') eq '4' ) {
  step4();
}
elsif ( GetFormValue('step') eq '5' ) {
    %><h2>Edit Non-MBARI Ship Dive <%= GetFormValue('RovName')%><%= GetFormValue('DiveNumber')%></h2>
    <font color="#408080"><strong>Fill in or edit the fields below, following
    the instructions in <span class="browntext">brown</span><font color="#408080">. Then submit
    the form to the Logistics Coordinator.</font></strong> </font>
    </td></tr></table><% 
    step5();
}
elsif (  GetFormValue('step') eq 'display' && GetFormValue('ExpeditionID') ) {
  %><h2 align="center">ExpeditionID <%= GetFormValue('ExpeditionID')%></h2>
  </td></tr></table><%
    display_expd();
}
elsif (  GetFormValue('finish') =~ /done/i ) {    # Really finish this expedition entry
  %><h2 align="center">Finished with Expedition <%= GetFormValue('ShipName')%><%= GetFormValue('ShipSeqNum')%></h2>
  </td></tr></table><%
  process_form_entries(); 
  if ( $Session->{'RegisteredUser'} ) {
    update_expedition();
    display_expd();
    finished_email();
  }
  else {
    email();
  }
  if ( !GetFormValue('ExpeditionID')) {
    $edit_again_url = $this_script . "?step=5&ShipName=" . GetFormValue('ShipName') . 
    "&ShipSeqNum=" . GetFormValue('ShipSeqNum') . "&edit=" . GetFormValue('ExpeditionID');
    close_session();
  }
  else {
    $edit_again_url = $this_script . "?step=3&ShipName=" . GetFormValue('ShipName') . 
    "&ShipSeqNum=" . GetFormValue('ShipSeqNum') . "&edit=" . GetFormValue('ExpeditionID');
    close_session();
}
%>

<br>
<b>Finished with this expedition.</b>
	<ul>
	<li><a href="<%= $edit_again_url%>">Edit this expedition again</a>
	<li><a href="<%= $this_script%>">Enter another post-cruise starting with simple search</a>
	<li><a href="<%= $this_script%>?search=advanced">Enter another post-cruise starting with advanced search options</a>
	</ul>

<%
} # End elsif ( ... /done/i
 
elsif (  GetFormValue('finish') =~ /dive information/i ) {		# Continue to enter dives using dive.asp%>
	<h2>Updating Expedition information</h2>
	<font color="#408080"><strong>Click on the button below to continue entering Dive information.
    </strong></font>
	</td></tr></table>
	<%
	process_form_entries(); 
	if ( $Session->{'RegisteredUser'} ) {
		update_expedition();
		finished_email();
	}
	else {
		email();
	}
	close_session();
	%>
	
	<table width="680">
	<form method="get" action="dive.asp">
	<input type="hidden" name="email_addr" value="<%= GetFormValue('email_addr')%>">
	<input type="hidden" name="ExpeditionID" value="<%= GetFormValue('ExpeditionID')%>">
  <input type="hidden" name="ShipName" value="<%= GetFormValue('ShipName')%>">
  <input type="hidden" name="DiveNumber" value="<%= GetFormValue('DiveNumber')%>">
  <input type="hidden" name="RovName" value="<%= GetFormValue('RovName')%>">

  <%
  if (GetFormValue('ShipName') eq 'prgn' || GetFormValue('ShipName') eq 'zphr') {%>
  <br>
  <font size="+3">
  <a type="submit" class="button big" href="<%= $this_script%>">Return to Simple Search</a>&nbsp;&nbsp; 
  <a type="submit" class="button big" href="<%= $this_script%>?search=advanced">Return to Advanced Search</a> 
  <br>
  </font>
  <br><b>ShipName: <%= GetFormValue('ShipName')%></b>
  <br>Will NOT add a new ROV Dive to ExpeditionID <%= GetFormValue('ExpeditionID')%>
  <br>No ROV Dives to enter.<br>
  <br>logistics_email = <%= $Session->{'logistics_email'}%>
  <br>
  </form>
  </table>
  
  </a>
  <%
  }
  else {
  %>
  <font size="+3"><br>
  &nbsp;<input type="submit" class="button blue" name="add_dive" value="Enter New Dive Information">
  <br>
</font>
	<br>(Will add a dive to ExpeditionID <%= GetFormValue('ExpeditionID')%>)
	<br>logistics_email = <%= $Session->{'logistics_email'}%>
  <font size="+3"><br><br>
  &nbsp;<a type="submit" class="button big" href="<%= $this_script%>">Return to Simple Search</a> &nbsp;&nbsp; 
  <a type="submit" class="button big" href="<%= $this_script%>?search=advanced">Return to Advanced Search</a> 
  </font>
	</form>
	</table>
	
	</a>
	<%
  }
} # End elsif ( ... /continue/i
else {
	%><h2 align="center">Unknown step <%= GetFormValue('step')%></h2>
	</td></tr></table><% 
}
%>

<!--#include file="postcruise_ftr.inc"-->





<%
#--------------------------------------------------------------------
#

=head3 step1()

Offer simple way to search for existing cruises in the database.

Author: Mike McCann

Date Created: 11/16/98

=cut

sub step1 {

%>

<font face="Arial,Helvetica">
<form method="POST" action="<%= $this_script%>" id="form1" name="form1">

<table border="4" cellpadding="5" cellspacing="1">
<tr><td>
    <table border="0" cellpadding="5" cellspacing="0">
            <tr>
              <td bgcolor="#B3D9D9"><big><strong>Search for: </strong></big></td>
            </tr>
            <tr>
              <td><input type="text" name="qString" size="20"> <br>
              <font color="#408080">Enter keyword or number<br>
              <font size="2">(Searches for text string in the following fields: Purpose, SiteTrackDesc,
              Participants, ChiefScientist, Accomplishments, Comments, ShipSequenceNum, DiveNumber.) </font></font><p>
              <input type="submit" class="button blue" name="Continue" value="Search Expeditions"> &nbsp;&nbsp;<input type="reset" class="button gray" name="Reset form" value="Reset">
              </td>
            </tr>
            <tr>
              <td align="right">&gt;&gt;&nbsp;<a type="submit" class="button big" href="<%= $this_script%>?search=advanced">Advanced Search</a> 
              <br><br>
              </strong><font size="2" color="brown">Search the EXPD database for other <br> parameters (e.g., Date, Ship, ROV, etc)</font></strong></input></div>
              </td>
            </tr>
    </table>
</td></tr>
</table>

  <input type="hidden" name="step" value="2">
</form>
<%
}
%>


<%
#--------------------------------------------------------------------
#

=head3 advanced_search()

Offer ways to search for existing cruises in the database.

Author: Mike McCann

Date Created: 11/16/98

=cut

sub advanced_search {

#
# Open database to get person list
#
	
open_database($dsn);		# Creates $Conn object as a global variable

$sql="SELECT * FROM Person ORDER BY LastName, FirstName";
##$sql="SELECT * FROM Person WHERE (DisplayChiefSciPickList = 1) ORDER BY LastName, FirstName";
##$sql="SELECT * FROM Person";
Win32::ASP::DebugPrint("\nExecuting SQL: \n$sql\n");
Win32::ASP::DebugPrint("\nFrom DSN: $dsn\n");
$RScs = $Conn->Execute($sql);
if(!$RScs) {
 	die "<BR>Empty Return Set11: RScs.<br> from $dsn ";
}

#
# Construct month & year pick lists. For year, start at 1988 and go to current year.
#
@month_names = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov",
                  "Dec");
my $lstr = localtime(time+864000);	# The year for 10 days from now
my @lstr2 = split('\s+', $lstr);
my $year_10 = $lstr2[4];
$lstr = localtime(time+86400*180);	# The year for 180 days from now
@lstr2 = split('\s+', $lstr);
my $year_180 = $lstr2[4];


for (my $day = 1; $day <= 31; $day++ ) {
	$qday_option_list .= "<option value=\"$day\">$day</option>";
}
foreach my $mon ( @month_names ) {
	$qmonth_option_list .= "<option value=\"$mon\">$mon</option>";
}
for (my $y = $year_180; $y >= 1988; $y--) {
	$qyear_option_list .= "<option value=\"$y\">$y</option>";
}
for (my $y = $year_10; $y >= 1988; $y--) {
	$qrev_year_option_list .= "<option value=\"$y\">$y</option>";
}


$sql="SELECT * FROM Region";
Win32::ASP::DebugPrint("\nExecuting SQL: \n$sql\n");
Win32::ASP::DebugPrint("\nFrom DSN: $dsn\n");
$RSrgn = $Conn->Execute($sql);
if(!$RSrgn) {
 	die "<BR>Empty Return Set12: RSrgn.<br> from $dsn ";
}
%>

<!-- for IE support -->
<script src="./es6-promise/promise.min.js"></script>
<!-- SET UP SWEETALERTS2 -->
<script src="./sweetalert2/sweetalert2.min.js"></script>
<link rel="stylesheet" href="./sweetalert2/sweetalert2.min.css">
<!-- SET UP FONT AWESOME FOR ARROWS -->
<!-- <script src="https://use.fontawesome.com/a26ffc8b21.js"></script> -->

<script language="JavaScript" type="text/javascript"><!--
function SelectRov(theForm) {
	alert("Must select an ROV with Dive Number.");
	theForm.qRovName.focus();
	return (true);
}
//-->

function ShipSelect(select) {
	//Pass in new function to configure the form
  	ConfigForm(select.options[select.selectedIndex].value);
  	
    if (select.options[select.selectedIndex].value == "othr") {
    	// Show alert only once.
    	if (!sessionStorage.returning) {
    	// SWEETALERT
     	swal({
        html:
          '<br> ' +
          '<b>For Unspecified Non-MBARI ship cruises there are no ExpeditionIDs to search on.  As a result, certain MBARI ship-specific query variables are <br>Not Allowed (N/A).</b>',
        showCloseButton: true,
        showCancelButton: false,
        confirmButtonText:
          'Got it',
        allowOutsideClick: false
      		}).then(function() {  // Yes Button
        // Close the Alert.....
        }, function(dismiss) {
          if (dismiss === 'cancel') {  //No Button
        }
      }) // End of swal()
      sessionStorage.returning = true; // set returning
    } // End of if
    return(true)
	}
	if (select.options[select.selectedIndex].value == "wfly" || "rcsn" || "ptlo" || "zphr" || "prgn" || "bhrzn") {
  		document.getElementById("qPurpose").value = "";
    	document.getElementById("qParticipants").value = "";
    	document.getElementById("qShipSeqNum").value = "";
    	document.getElementById("qPlannedTrackDesc").value = "";
    	document.getElementById("qAccomplishments").value = "";
    	document.getElementById("qScientistsComments").value = "";
    	document.getElementById("qOperatorComments").value = "";
    	document.getElementById("qOtherComments").value = "";   
      document.getElementById("qDiveName").value = "";
    	return(true);
    }
}

// Initial Landing Page Configuration
function ConfigForm() {
	  var selectedShip = document.getElementById('qShipName').value;
	  //console.log("selectedShip = " + document.getElementById('qShipName').value);

	  if (selectedShip == "othr") {
	  	    //console.log("othr");
  			document.getElementById("qPurpose").disabled = true;
    		document.getElementById("qParticipants").disabled = true;
    		document.getElementById("qShipSeqNum").disabled = true;
    		document.getElementById("qPlannedTrackDesc").disabled = true;
    		document.getElementById("qAccomplishments").disabled = true;
    		document.getElementById("qScientistsComments").disabled = true;
    		document.getElementById("qOperatorComments").disabled = true;
    		document.getElementById("qOtherComments").disabled = true; 
    		document.getElementById("qExpdPrincipalInvestigator").disabled = true;

			  document.getElementById("qPurpose").value = "N/A";
   	 		document.getElementById("qParticipants").value = "N/A";
   	 		document.getElementById("qShipSeqNum").value = "N/A";
   	 		document.getElementById("qPlannedTrackDesc").value = "N/A";
   	 		document.getElementById("qAccomplishments").value = "N/A";
   	 		document.getElementById("qScientistsComments").value = "N/A";
    		document.getElementById("qOperatorComments").value = "N/A";
    		document.getElementById("qOtherComments").value = "N/A"; 

    		document.getElementById("qPurpose").style.background = "lightgray";
    		document.getElementById("qParticipants").style.background = "lightgray";
   	 		document.getElementById("qShipSeqNum").style.background = "lightgray";
   	 		document.getElementById("qPlannedTrackDesc").style.background = "lightgray";
   	 		document.getElementById("qAccomplishments").style.background = "lightgray";
   	 		document.getElementById("qScientistsComments").style.background = "lightgray";
    		document.getElementById("qOperatorComments").style.background = "lightgray";
    		document.getElementById("qOtherComments").style.background = "lightgray";
    		document.getElementById("qExpdPrincipalInvestigator").style.background = "lightgray";

	    }  
	    //else if (selectedShip == "wfly" || "rcsn" || "ptlo" || "zphr" || "prgn") {
      else {
              document.getElementById("qPurpose").disabled = false;
              document.getElementById("qParticipants").disabled = false;
              document.getElementById("qShipSeqNum").disabled = false;
              document.getElementById("qPlannedTrackDesc").disabled = false;
              document.getElementById("qAccomplishments").disabled = false;
              document.getElementById("qScientistsComments").disabled = false;
              document.getElementById("qOperatorComments").disabled = false;
              document.getElementById("qOtherComments").disabled = false;
              document.getElementById("qExpdPrincipalInvestigator").disabled = false;
              document.getElementById("qDiveName").disabled = false;

              document.getElementById("qDiveName").value = "";

              document.getElementById("qPurpose").style.background = "";
              document.getElementById("qParticipants").style.background = "";
              document.getElementById("qShipSeqNum").style.background = "";
              document.getElementById("qPlannedTrackDesc").style.background = "";
              document.getElementById("qAccomplishments").style.background = "";
              document.getElementById("qScientistsComments").style.background = "";
              document.getElementById("qOperatorComments").style.background = "";
              document.getElementById("qOtherComments").style.background = "";
              document.getElementById("qExpdPrincipalInvestigator").style.background = "";
              document.getElementById("qDiveName").style.background = "";
      }
}
//-->

function ResetForm() {
			document.getElementById("qPurpose").disabled = false;
    	document.getElementById("qParticipants").disabled = false;
    	document.getElementById("qShipSeqNum").disabled = false;
    	document.getElementById("qPlannedTrackDesc").disabled = false;
    	document.getElementById("qAccomplishments").disabled = false;
    	document.getElementById("qScientistsComments").disabled = false;
    	document.getElementById("qOperatorComments").disabled = false;
    	document.getElementById("qOtherComments").disabled = false;
    	document.getElementById("qExpdPrincipalInvestigator").disabled = false;
      document.getElementById("qDiveName").disabled = false;

      document.getElementById("qDiveName").value = ""; 

    	document.getElementById("qPurpose").style.background = "";
    	document.getElementById("qParticipants").style.background = "";
   	 	document.getElementById("qShipSeqNum").style.background = "";
   	 	document.getElementById("qPlannedTrackDesc").style.background = "";
   	 	document.getElementById("qAccomplishments").style.background = "";
   	 	document.getElementById("qScientistsComments").style.background = "";
    	document.getElementById("qOperatorComments").style.background = "";
    	document.getElementById("qOtherComments").style.background = "";
    	document.getElementById("qExpdPrincipalInvestigator").style.background = "";
      document.getElementById("qDiveName").style.background = "";
}
//-->

// Hitting the Back Button  ////
// Set Location to return to on the screen one the back button has been hit.
function scrollWin() {
    window.scrollTo(0, 0);
}
//-->

// Reload the page prior to the first Submit
window.onload = function (e) {
  ConfigForm();
  setTimeout(scrollWin, 10);
}
//-->
</script>

<form method="POST" action="<%= $this_script%>" id="form1" name="form1">
  <input type="hidden" name="step" value="2">
  <table table-layout="fixed" border="0" cellpadding="5" cellspacing="0">
	<tr>
	  <td colspan="8" align="left" valign="top" bgcolor="#F0F0F0">
        <b>Join fields below with:</b>&emsp;<input type="radio" name="Conjunction" value="AND" checked>AND 
								  &emsp;<input type="radio" name="Conjunction" value="OR">OR
	  </td>
	</tr>
    <tr>
      <br><td valign="middle"><font face="Arial,Helvetica" size="-1"><b>ShipName</b></font></td>
      <td>
        <div class="shipSelect">
        <select id="qShipName" name="qShipName" text-align-last="center" onChange="return ShipSelect(this)">
		      <option value=""><--Select--></option>		
        <optgroup label="MBARI Ship">
		      <option text-align="right" value="ptlo">&emsp;&emsp;&emsp;Point Lobos</option>
		      <option value="wfly">&emsp;&emsp;&emsp;Western Flyer</option>
		      <option value="zphr">&emsp;&emsp;&emsp;Zephyr</option>
		      <option value="rcsn">&emsp;&emsp;&emsp;Rachel Carson</option>
		      <option value="prgn">&emsp;&emsp;&emsp;Paragon</option>
        </optgroup>
        <optgroup class="optgroup label" label="Non-MBARI Ship - MiniROV">
          <option value="bhrzn">&emsp;&emsp;&emsp;Bold Horizon (CalCOFI)</option>
          <option value="othr">&emsp;&emsp;&emsp;Other Ship</option>
        </optgroup>
		    </select>
      </div>
      </td>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Sequence Number</font></strong></td>
      <td><input type="text" id="qShipSeqNum" name="qShipSeqNum" size="5">
      <span class="smallbrownText">for legacy sake</span></td>      
    </tr>
    <tr>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">DiveNumber</font></strong></td>
      <td><input type="text" id="qDiveName" name="qDiveName" size="6">
      <span class="smallbrownText">May precede with first ROV Name Letter</span><br>
      <span class="smallbrownText">&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;(e.g., V3296, T1001, D1, M32)</span>
      </td>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Status</font></strong></td>
      <td><select name="qStatCode">
		<option value=""></option>
		<option value="cmpl">Completed</option>
		<option value="abrt">Aborted</option>
		<option value="plnd">Planned</option>
		<option value="cncl">Canceled</option>
		</select>
	  </td>
    </tr>
    <tr>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Date Choices</font></strong></td>
      <td><select name="qDates">
		<option value=""></option>
		<option value="Future">Future</option>
		<option value="last7">Last 7 days</option>
		<option value="last30">Last 30 days</option>
		<option value="last60">Last 60 days</option>
		<option value="last180">Last 180 days</option>
		<%= $qrev_year_option_list %>
		</select> 
		<span class="smallbrownText">Based on scheduled start time</span>
	  </td>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">YYYYDDD</font></strong></td>
      <td><input type="text" name="qYYYYDDD" size="7" value=""> 
      <span class="smallbrownText">e.g. 1999075</span></td>
    </tr>
    <tr>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Start Date</font></strong></td>
      <td>
        <font face="Arial,Helvetica" size="-1">First of </font><select name="qSmon" size="1">
		<option value=""></option>
		<%= $qmonth_option_list %>
		</select>
		<font face="Arial,Helvetica" size="-1"> </font><select name="qSyr">
		<option value=""></option>
		<%= $qyear_option_list %>
		</select> 
		<span class="smallbrownText">Specify End Date too ---&gt;</span>
	  </td>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">End Date</font></strong></td>
      <td>
        <font face="Arial,Helvetica" size="-1">First of </font><select name="qEmon" size="1">
		<option value=""></option>
		<%= $qmonth_option_list %>
		</select> 
		<font face="Arial,Helvetica" size="-1"> </font><select name="qEyr">
		<option value=""></option>
		<%= $qyear_option_list %>
		</select> 
	  </td>
    </tr>
    <tr>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Purpose</font></strong></td>
      <td><input type="text" id="qPurpose" name="qPurpose" size="50" value=""></td>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Single Date</font></strong></td>
      <td>
      	<select name="qSingleDay" size="1">
      	<option value=""></option>
		<%= $qday_option_list %>
		</select> 
		<select name="qSingleMon" size="1">
		<option value=""></option>
		<%= $qmonth_option_list %>
		</select> 
		<select name="qSingleYr" size="1">
		<option value=""></option>
		<%= $qyear_option_list %>
		</select> 
      
      </td>
    </tr>
    <tr>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Site/Track</font></strong></td>
      <td><input type="text" id="qPlannedTrackDesc" name="qPlannedTrackDesc" size="50" value=""></td>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Participants</font></strong></td>
      <td><input type="text" id="qParticipants" name="qParticipants" size="50" value=""></td>
    </tr>
    <tr>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Accomplishments</font></strong></td>
      <td><input type="text" id="qAccomplishments" name="qAccomplishments" size="50" value=""></td>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Chief Scientist</font></strong></td>
      <td>
	<select name="qExpdChiefScientist" size="1">
      <option value=""></option>
	    <% while ( !$RScs->EOF ) {
		$fullname = $RScs->Fields('FirstName')->value . " " . $RScs->Fields('LastName')->value;
		%><option value="<%= $fullname%>"><%= $fullname%></option>
                <%
		$RScs->MoveNext;
	     } %>
      	</select>
      </td>
    </tr>
    <tr>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Operator Comments</font></strong></td>
      <td><input type="text" id="qOperatorComments" name="qOperatorComments" size="50" value=""></td>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Principal Investigator</font></strong></td>
      <td>
	<select id="qExpdPrincipalInvestigator" name="qExpdPrincipalInvestigator" size="1" disabled="false">
            <option value=""></option>
	    <% 
	    $RScs->MoveFirst;
	    while ( !$RScs->EOF ) {
		$fullname = $RScs->Fields('FirstName')->value . " " . $RScs->Fields('LastName')->value;
		%><option value="<%= $fullname%>"><%= $fullname%></option>
                <%
		$RScs->MoveNext;
	     } %>
      	</select>
      </td>
    </tr>
    <tr>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Scientist Comments</font></strong></td>
      <td><input type="text" id="qScientistsComments" name="qScientistComments" size="50" value=""></td>
      <td valign="middle"><strong><font face="Arial,Helvetica" size="-1">Data Comments</font></strong></td>
      <td><input type="text" id="qOtherComments" name="qOtherComments" size="50" value=""></td>
    </tr>
  </table><hr><br>
  <font face="Arial,Helvetica" size="3"<b>Return results as:&emsp;</b> <input type="radio" name="data_table" value="no" checked>Cruise Report List
  &emsp;<input type="radio" name="data_table" value="yes">Data Table
  </font>
  <br>
  <p></p><br>
  <p class="left">
  <input type="hidden" name="search" value="advanced"> 
  &nbsp;<input type="reset" class="button gray" value="Reset Form Fields" name="Reset form" style="width: auto" onclick="ResetForm()">&nbsp;&nbsp;&nbsp;&nbsp;
  <input type="submit" class="button blue" value="Search For Expeditions" name="Continue" style="width: auto">

  </p><br><br>
  </form>

<%
$RSrgn->Close;
$RScs->Close;
$Conn->Close;
}	# End advanced_search()
%>

<%
#--------------------------------------------------------------------
#

=head3 step2()

Return list of cruises from query specified in step1() or advanced_search().
If argument of 'data_table' passed then return results as a table.

Author: Mike McCann

Date Created: 11/16/98

=cut

sub step2 {

my $startBM = new Benchmark;

#
# Open database to get do query and produce list of  expeditions
#
open_database($dsn);		# Creates $Conn object as a global variable

#
# Open VARS database to get data for dataLinks()
#
open_vars_database();		# Creates $ConnVARS object as a global variable

#
# Open Samples database to get data for dataLinks()
#
open_samples_database();		# Creates $ConnSamples object as a global variable

$conjunction = GetFormValue('Conjunction');
my $has3Dlinks = 0;

#
# Copy form values into lookup list so that we can accept them from either
# step1() or advanced_search()
#
@qFields = ('qStatCode', 'qPurpose', 'qParticipants', 'qPlannedTrackDesc', 
			'qAccomplishments', 'qExpdChiefScientist', 'qShipSeqNum', 'qExpdPrincipalInvestigator',
			'qYYYYDDD', 'qScientistComments', 'qOperatorComments', 'qOtherComments' );

if ( GetFormValue('search') eq 'advanced' ) {
	foreach $q ( @qFields ) {
		$qvalue{$q} = GetFormValue($q) if GetFormValue($q);
	}
}
else {
	$conjunction = 'OR';
	foreach $q ( @qFields ) {
		$qvalue{$q} = GetFormValue('qString') if GetFormValue('qString');
	}
}

################################################
#
# Construct where clause for a dive number query
#
################################################
if ( GetFormValue('qDiveName') ) {
  GetFormValue('qDiveName') =~ /(^[tvdmTVDM]*)(\d+)$/;
  Win32::ASP::DebugPrint("\nqDiveName match#" . GetFormValue('qDiveName') . ":  1 = $1, 2 = $2\n");
  $qDiveNumber = $2;  # Got to pick off 2 first for some reason (?)
  if ( $1 =~ /T/i || GetFormValue('qRovName') eq 'tibr' ) {
    $qRovName = 'tibr';
  }
  elsif ( $1 =~ /V/i || GetFormValue('qRovName') eq 'vnta' ) {
    $qRovName = 'vnta';
  }
  elsif ( $1 =~ /D/i || GetFormValue('qRovName') eq 'docr' ) {
    $qRovName = 'docr';
  }
  elsif ( $1 =~ /M/i || GetFormValue('qRovName') eq 'mini' ) {
    $qRovName = 'mini';
  }
#  else {
#    %>
<!--    <h2><font color="#408080">'<%=GetFormValue('qDiveName') %>' is not a valid ROV Name.   Please choose 'V', 'T', 'D' or 'M' to precede DiveNumber for specific MBARI ROV dives (Ventana, Tiburon, Doc Ricketts, or MiniROV).</font></h2>-->
<%
#    return;
#  }
  Win32::ASP::DebugPrint("\nqDiveName match#" . GetFormValue('qDiveName') . ":  qRovName = $qRovName, qDiveNumber = $qDiveNumber\n");

  $dive_clause = "( ExpeditionID IN (SELECT ExpeditionID_FK FROM Dive \n";
  $dive_clause .= "WHERE DiveNumber = " . $qDiveNumber;
    if ( $qRovName && $qRovName == 'tibr') {
       $dive_clause .= " AND RovName = '" . $qRovName . "'))";
    } elsif ( $qRovName && $qRovName == 'vnta') {
       $dive_clause .= " AND RovName = '" . $qRovName . "'))";
    } elsif ( $qRovName && $qRovName == 'docr') {
       $dive_clause .= " AND RovName = '" . $qRovName . "'))";
    } elsif ( $qRovName && $qRovName == 'mini') {
       $sql = "SELECT * FROM Dive \n";
       $dive_clause = "WHERE DiveNumber = " . $qDiveNumber;
       $dive_clause .= " AND DeviceID = 1859 AND RovName = '" . $qRovName . "' AND DiveNumber = " . $qDiveNumber . "'))";   
    } else {
       $dive_clause .= "))";
    }
}
elsif ( GetFormValue('qDiveNumber') =~ /^\d+$/ ) {    # Case where advanced form is not used
  $dive_clause = "( ExpeditionID IN (SELECT ExpeditionID_FK FROM Dive \n";
  $dive_clause .= "WHERE DiveNumber = " . GetFormValue('qDiveNumber');
  if ( GetFormValue('qRovName') =~ /^[tvdTVD]/) {
    $dive_clause .= " AND RovName = '" . GetFormValue('qRovName') . "'))";
  }
  elsif ( GetFormValue('qRovName') =~ /^[mM]/) {
    $sql = "SELECT * FROM Dive \n";
    $dive_clause = "WHERE DiveNumber = " . GetFormValue('qDiveNumber');
    $dive_clause .= " AND DeviceID = 1859 AND RovName = '" . GetFormValue('qRovName') . "'))";
  }
  else {
    $dive_clause .= "))";
  }
}
elsif ( GetFormValue('qString') =~ /^\d+$/ ) {    # Look for dive if number specified
  $dive_clause = "( ExpeditionID IN (SELECT ExpeditionID_FK FROM Dive \n";
  $dive_clause .= "WHERE DiveNumber = " . GetFormValue('qString') . "))";
}
else {
  $dive_clause = '';
}
#print "<br>Dive_Clause: $dive_clause<br>";
###########################################################
#
# Construct where clause if a date range was specified
#
###########################################################
if ( GetFormValue('qShipName') eq 'othr' ) {
	if ( GetFormValue('qDates') ) {
		if ($_db_type eq 'mssql') {
			if ( GetFormValue('qDates') =~ /last(\d+)/ ) {
				$date_clause = "( DateDiff(\"dy\", DiveEndDtg, getdate()) < $+ )";
				$date_clause .= " AND ( DateDiff(\"dy\", DiveStartDtg, getdate()) >= 0 )";
			}
			elsif ( GetFormValue('qDates') =~ /\d\d\d\d/ ) {
				$date_clause = "( DatePart(\"yy\", DiveStartDtg) = " . GetFormValue('qDates') . ")";
			}
			elsif ( GetFormValue('qDates') =~ /Future/i ) {
				$date_clause = "( DateDiff(\"dy\", DiveStartDtg, getdate()) <= 0 )";
			}
			else {
				$date_clause = "";
			}
		}
		elsif ($_db_type eq 'access') {
			$date_clause = "( DateDiff('d', DiveStartDtg, date()) < 7 )" 
				if GetFormValue('qDates') =~ 'last7';
			$date_clause = "( DateDiff('d', DiveStartDtg, date()) < 30 )" 
				if GetFormValue('qDates') =~ 'last30';
			$date_clause = "( DatePart('yyyy', DiveStartDtg) = " . GetFormValue('qDates') . ")" 
				if GetFormValue('qDates') =~ /\d\d\d\d/;
		}
		else {
			die "<br>db_type '$_db_type' not implemented.";
		} # if ($_db_type eq ...
	}
	elsif ( GetFormValue('qSmon') && GetFormValue('qSyr') && GetFormValue('qEmon') && GetFormValue('qEyr') ) {
		$date_clause = "( DiveStartDtg BETWEEN '1 ";
		$date_clause .= GetFormValue('qSmon') . ' ' . GetFormValue('qSyr');
		$date_clause .= "')";
	}
	elsif ( GetFormValue('qSingleDay') && GetFormValue('qSingleMon') && GetFormValue('qSingleYr') ) {
		$date_clause = "('" . GetFormValue('qSingleDay') . ' ' . GetFormValue('qSingleMon') . ' ' . GetFormValue('qSingleYr') . "' ";
		$date_clause .= " BETWEEN DiveStartDtg AND DiveEndDtg)";
		$date_clause .= " OR ('" . (GetFormValue('qSingleDay') + 1) . ' ' . GetFormValue('qSingleMon') . ' ' . GetFormValue('qSingleYr') . "' ";
		$date_clause .= " BETWEEN DiveStartDtg AND DiveEndDtg)";
	}
}
elsif ( GetFormValue('qShipName') ne 'othr') {
	if ( GetFormValue('qDates') ) {
		if ($_db_type eq 'mssql') {
			if ( GetFormValue('qDates') =~ /last(\d+)/ ) {
				$date_clause = "( DateDiff(\"dy\", ScheduledEndDtg, getdate()) < $+ )";
				$date_clause .= " AND ( DateDiff(\"dy\", ScheduledStartDtg, getdate()) >= 0 )";
			}
			elsif ( GetFormValue('qDates') =~ /\d\d\d\d/ ) {
				$date_clause = "( DatePart(\"yy\", ScheduledStartDtg) = " . GetFormValue('qDates') . ")";
			}
			elsif ( GetFormValue('qDates') =~ /Future/i ) {
				$date_clause = "( DateDiff(\"dy\", ScheduledStartDtg, getdate()) <= 0 )";
			}
			else {
				$date_clause = "";
			}
		}
		elsif ($_db_type eq 'access') {
			$date_clause = "( DateDiff('d', ScheduledStartDtg, date()) < 7 )" 
				if GetFormValue('qDates') =~ 'last7';
			$date_clause = "( DateDiff('d', ScheduledStartDtg, date()) < 30 )" 
				if GetFormValue('qDates') =~ 'last30';
			$date_clause = "( DatePart('yyyy', ScheduledStartDtg) = " . GetFormValue('qDates') . ")" 
				if GetFormValue('qDates') =~ /\d\d\d\d/;
		}
		else {
			die "<BR>db_type '$_db_type' not implemented.";
		} # if ($_db_type eq ...
	}
	elsif ( GetFormValue('qSmon') && GetFormValue('qSyr') && GetFormValue('qEmon') && GetFormValue('qEyr') ) {
		$date_clause = "( ScheduledStartDtg BETWEEN '1 ";
		$date_clause .= GetFormValue('qSmon') . ' ' . GetFormValue('qSyr');
		$date_clause .= "' AND '1 " . GetFormValue('qEmon') . ' ' . GetFormValue('qEyr');
		$date_clause .= "')";
	}
	elsif ( GetFormValue('qSingleDay') && GetFormValue('qSingleMon') && GetFormValue('qSingleYr') ) {
		$date_clause = "('" . GetFormValue('qSingleDay') . ' ' . GetFormValue('qSingleMon') . ' ' . GetFormValue('qSingleYr') . "' ";
		$date_clause .= " BETWEEN StartDtg AND EndDtg)";
		$date_clause .= " OR ('" . (GetFormValue('qSingleDay') + 1) . ' ' . GetFormValue('qSingleMon') . ' ' . GetFormValue('qSingleYr') . "' ";
		$date_clause .= " BETWEEN StartDtg AND EndDtg)";
		$date_clause = "('" . GetFormValue('qSingleDay') . ' ' . GetFormValue('qSingleMon') . ' ' . GetFormValue('qSingleYr') . "' ";
		$date_clause .= " BETWEEN ScheduledStartDtg AND ScheduledEndDtg)";
		$date_clause .= " OR ('" . (GetFormValue('qSingleDay') + 1) . ' ' . GetFormValue('qSingleMon') . ' ' . GetFormValue('qSingleYr') . "' ";
		$date_clause .= " BETWEEN ScheduledStartDtg AND ScheduledEndDtg)";
	}
}
else {
	$date_clause = '';
}

################################################
#
# Construct where clause for a YYYYDDD query
#
#################################################
if ( GetFormValue('qShipName') eq 'othr') {
	if ( GetFormValue('qYYYYDDD') ) {
		my $yyyy = substr(GetFormValue('qYYYYDDD'),0,4);
		my $ddd = substr(GetFormValue('qYYYYDDD'),4,3);
		if ( ($ddd > 0 && $ddd < 366) && $yyyy > 1986 ) {
			$yyyyddd_clause = "((DATEPART(yy,DiveStartDtg) = $yyyy ";
			$yyyyddd_clause .= "AND DATEPART(yy,DiveEndDtg) = $yyyy ";
			$yyyyddd_clause .= "AND DATEPART(dy,DiveStartDtg) <= $ddd ";
			$yyyyddd_clause .= "AND DATEPART(dy,DiveEndDtg) >= $ddd)) ";
		}
		else {
			$yyyyddd_clause = '';
		}
	}
}
elsif ( GetformValue('qShipName') ne 'othr') {
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
}
else {
	$yyyyddd_clause = '';
}



################################################
#
# Construct where clause for ship name(s) - need to identify if Ship is an MBARI Ship or a Ship of Opportunity
#
################################################

if ( GetFormValue('qShipName') ) {
    if ( GetFormValue('qShipName') eq 'othr' ) {
        if ( $qDiveNumber && $qRovName ) {
          $dive_clause = "ExpeditionID_FK = 0 AND DeviceID = 1859 AND RovName = '" . $qRovName . "' AND DiveNumber = " . $qDiveNumber;
          $ship_clause = '';
        } 
        elsif ( $qDiveNumber && !$qRovName ) {
          $dive_clause = "ExpeditionID_FK = 0 AND DeviceID = 1859 AND DiveNumber = " . $qDiveNumber;
          $ship_clause = '';
        }
        else {
          $ship_clause = "ExpeditionID_FK = 0 AND DeviceID = 1859 AND RovName = 'mini'"; 
        	 if (GetFormValue('qExpdChiefScientist') ) {
      			$DCS = (GetFormValue('qExpdChiefScientist') );
      			$dive_clause .= " DiveChiefScientist = '" . $DCS . "'";
      	      }
        }

        $sql = "SELECT DiveID, DeviceID, RovName, DiveNumber, DiveChiefScientist, DiveStartDtg, DiveEndDtg, ExpeditionID_FK, 'cmpl' as StatCode FROM Dive\n";        
        $sql .= "WHERE\n";
        $sql .= $ship_clause . "\nAND " if $ship_clause;
        $sql .= $date_clause . "\n$conjunction " if $date_clause;
        $sql .= $dive_clause . "\n$conjunction " if $dive_clause;
        $sql .= $yyyyddd_clause . "\n$conjunction " if $yyyyddd_clause;

        $sql .= "(";
        foreach $q ( @qFields ) {
        	# Non-MBARI Ship Expeditions Do not have ExpdChiefScientist but rather DiveChiefScientists
        	if ($q ne 'qExpdChiefScientist') {
        		$f = $q; $f =~ s/^q//;
          		next unless $qvalue{$q};
          		next if $f =~ /YYYYDDD/;
          		if ( $f =~ /Num$/ || $f =~ /ID$/ ) {
            		$sql .= "$f = " . $qvalue{$q} . "\n$conjunction " if $qvalue{$q} =~ /^\d+$/;
          		}
          		else {
            		$sql .= "$f LIKE '%" . $qvalue{$q} . "%'\n$conjunction ";
          		}
      		}
      		else {
      			next;
  			}
  		}

      $sql =~ s/\nAND\s*$//;
      $sql =~ s/\nOR\s*$//;
      $sql .= ")";  
      $sql =~ s/\s*AND\s*\(\s*\)\s*$//;
      $sql =~ s/\s*OR\s*\(\s*\)\s*$//;
      $sql .= "\nORDER BY DiveNumber";
    }
    elsif ( GetFormValue('qShipName') ne 'othr' ) {
        $ship_clause = "(ShipName = '" . GetFormValue('qShipName') . "')\n";
        $sql = "SELECT * FROM Expedition\n";
        #$sql = "SELECT Dive.DeviceID, Dive.RovName, * FROM Expedition\n";
        ## Adding this (below JOIN) to help with DataLinks call later on
        #$sql .= " JOIN Dive ON Expedition.ExpeditionID = Dive.ExpeditionID_FK \n";
        $sql .= "WHERE\n";
        $sql .= $ship_clause . "\nAND " if $ship_clause;
        $sql .= $date_clause . "\n$conjunction " if $date_clause;
        $sql .= $dive_clause . "\n$conjunction " if $dive_clause;
        $sql .= $yyyyddd_clause . "\n$conjunction " if $yyyyddd_clause;

        $sql .= "(";
        foreach $q ( @qFields ) {
        $f = $q; $f =~ s/^q//;
          next unless $qvalue{$q};
          next if $f =~ /YYYYDDD/;
          if ( $f =~ /Num$/ || $f =~ /ID$/ ) {
            $sql .= "$f = " . $qvalue{$q} . "\n$conjunction " if $qvalue{$q} =~ /^\d+$/;
          }
          else {
            $sql .= "$f LIKE '%" . $qvalue{$q} . "%'\n$conjunction ";
          }
        }

      $sql =~ s/\nAND\s*$//;
      $sql =~ s/\nOR\s*$//;
      $sql .= ")";  
      $sql =~ s/\s*AND\s*\(\s*\)\s*$//;
      $sql =~ s/\s*OR\s*\(\s*\)\s*$//;
      $sql .= "\nORDER BY ScheduledStartDtg";
    }
} 
else {	
  $ship_clause = "";

    $sql = "SELECT * FROM Expedition\n";
    #$sql .= " JOIN Dive ON Expedition.ExpeditionID = Dive.ExpeditionID_FK \n";
    $sql .= "WHERE\n";
  	$sql .= $ship_clause . "\nAND " if $ship_clause;
  	$sql .= $date_clause . "\n$conjunction " if $date_clause;
  	$sql .= $dive_clause . "\n$conjunction " if $dive_clause;
  	$sql .= $yyyyddd_clause . "\n$conjunction " if $yyyyddd_clause;

  $sql .= "(";
  foreach $q ( @qFields ) {
  $f = $q; $f =~ s/^q//;
    next unless $qvalue{$q};
    next if $f =~ /YYYYDDD/;
    if ( $f =~ /Num$/ || $f =~ /ID$/ ) {
      $sql .= "$f = " . $qvalue{$q} . "\n$conjunction " if $qvalue{$q} =~ /^\d+$/;
    }
    else {
      $sql .= "$f LIKE '%" . $qvalue{$q} . "%'\n$conjunction ";
    }
  }

  $sql =~ s/\nAND\s*$//;
  $sql =~ s/\nOR\s*$//;
  $sql .= ")";  
  $sql =~ s/\s*AND\s*\(\s*\)\s*$//;
  $sql .= $order_by;
}
#print "<br>RS->Fields: $sql<br>";
#
# Count the records returned by the query and offer to bail if too many
#
my $count_sql = $sql;

 if ( GetFormValue('qShipName') eq 'othr' ) {
        $count_sql =~ s/SELECT \Dive.*/SELECT count\(\*\), 'cmpl' as StatCode FROM Dive /;
 } 
 else {
        $count_sql =~ s/SELECT \*/SELECT count\(\*\)/;
}
$count_sql =~ s/ORDER BY.*//;

#print "<br>Count_sql: $count_sql<br>";

################################
Win32::ASP::DebugPrint("\nExecuting count_sql: $count_sql");
$RS = $Conn->Execute($count_sql);
if (!$RS) {
 	  die "<BR>Empty Return Set1: RS. for $count_sql ";
}
# Record Set exists
#elsif ($RS->Fields('Dive.RovName') eq 'mini') { 
#  $num_expd = $count;
#  print "<br>$num_expd<br>";
#} 
else {
	$num_expd = $RS->Fields(0)->Value;
}
#print "<br>Count: $num_expd<br>";


Win32::ASP::DebugPrint("\ncount_check value = " . GetFormValue('count_check') );
if ( $num_expd > 100 &&  GetFormValue('count_check') !~ /yes/i ) { %>
	<h2>
	The query you have chosen will return <%= $num_expd%> expeditions. <br><br>Do you want to continue?
	</h2>
	<form method="POST" id="form2" name="form2">
	   <input type="button" class="button gray" name="count_check" value="No. Go Back To Search" onclick="history.go( -1 );return true;">
	</form>

	<form><%
	Win32::ASP::DebugPrint("\nAssigning hidden variables:\n" );
	# If step1() or advanced_search uses GET use QueryString, if POST then Form
	# Actually, QueryString() does not work, so I switched back to POST.  13 Oct 2006
	##foreach my $name ( Win32::OLE::in ( $Request->QueryString() ) ) {
	foreach my $name ( Win32::OLE::in ( $Request->Form() ) ) {
		##Win32::ASP::DebugPrint("\nname value = " . Request->QueryString($name) );
		Win32::ASP::DebugPrint("\nname value = " . $Request->Form($name)->{item} );
		%>
		<input type="hidden" name="<%= $name%>" value="<%= $Request->Form($name)->{item} %>">
		<%
	} %>
	<input type="hidden" name="step" value="3"> 
	   <input type="submit" class="button blue" name="count_check" value="Yes. Show Me All. <%= $num_expd%> expeditions">
	</form>
	<br>
	<font color=red><b>NOTE: If you display this long list please wait for your browser to finish loading the page.</b></font>
	<%
	exit;
}
  ####print "<br>Step2: $sql<br>";
$RS->Close;

########################################################
#Branch off to query NON MBARI SHIPS HERE (step4) otherwise, keep going.
if ( GetFormValue('qShipName') eq 'othr' ) {
	   #print "<br><br>----------------------<br>STEP4: ExpeditionID_FK = 0";
	   #print "<br>$qShipName<br>$qRovName<br>";
	  step4();
}
########################################################

else {
  ####print "<br>Step3: $sql<br>";
Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
 	die "<BR>Empty Return Set2: RS. for $sql ";
}

#
# Header for returned expeditions, save sql for display at end of list
#
my $search_sql = $sql;
if ( $num_expd > 0 ) {
%>
<h3>Displaying <%= $num_expd%> 
<% if ( $num_expd > 1 ) { %>
expeditions
<% } 
else { %>
expedition
<%}
if ($_[0] eq 'data_table') { %>
 in table format
 <% if ($num_expd > 10) { %>
	(all data must transfer before table is shown)
 <%}%>
<%}%>
...</h3> 
<% }

#foreach my $v (Win32::OLE::in ($Request->ServerVariables)) {
#	$Response->Write("$v: " . $Request->ServerVariables($v)->{Item} . "<br>");
#}

#
# Now cursor through the record set and print the expeditions
#
if ($_[0] eq 'data_table') { %>
	<form action="/3Dreplay/3D.asp" target="3Dreplay">
	<table BORDER="1">
	<tr>
  <th BGCOLOR="#00C0C0">Date</th>
  <th BGCOLOR="#00C0C0">Expedition Report</th>
  <th BGCOLOR="#00C0C0">3D Replay<sup>0</sup></th>
  <th BGCOLOR="#00C0C0">Annotations<font COLOR="#404040"><sup>1</sup></font></th>
  <th BGCOLOR="#00C0C0">Samples<font COLOR="#404040"><sup>2</sup></font></th>
  <th BGCOLOR="#00C0C0">Images<font COLOR="#404040"><sup>3</sup></font></th>
  <th BGCOLOR="#00C0C0">ROVCTD <font COLOR="#404040"><sup>4</sup></font></th>
  <th BGCOLOR="#00C0C0">Camera Log<font COLOR="#404040"><sup>5</sup></font></th>
  <th BGCOLOR="#00C0C0">Navigation<font COLOR="#404040"><sup>6</sup></font></th>
  </tr>
<%} else { %>
	<dl>
<%}
$delimiter = '<font face="helvetica,Ariel"> <b>&middot;</b> </font>';
$delimiter .= "<br>\n" if ($_[0] eq 'data_table');

$count = 0;
while ( !$RS->EOF ) {
	$count++;
	$Status = $RS->Fields('StatCode')->Value;
	$Status = "Aborted" if $Status eq 'abrt';
	$Status = "Completed" if $Status eq 'cmpl';
	$Status = "Canceled" if $Status eq 'cncl';
	$Status = "Planned" if $Status eq 'plnd';
	
  ##########################  Determining Start and End Times #######################
	my @l = ($Status eq 'Completed') ?  split(' ', $RS->Fields('StartDtg')->Value) : 
		split(' ', $RS->Fields('ScheduledStartDtg')->Value);
	$StartDate = $l[0];
	my ($mo, $da, $yr) = split('/', $StartDate);
	#print "<br>step2(): yr= $yr";
	my ($hr, $mn) = split(':', $l[1]);
	$hr = ($hr + 12) if $l[2] =~ /PM/i;
	my $YYYYDDD = toYDfromDate($mo,$da,$yr,$hr,$mn);
  #########################################
	
	my @e = ($Status eq 'Completed') ?  split(' ', $RS->Fields('EndDtg')->Value) : 
		split(' ', $RS->Fields('ScheduledEndDtg')->Value);
	$EndDate = $e[0];
	($mo, $da, $yr) = split('/', $EndDate);
	#print "<br>step2(): yr= $yr";
	($hr, $mn) = split(':', $e[1]);
	$hr = ($hr + 12) if $e[2] =~ /PM/i;
	my $endYYYYDDD = toYDfromDate($mo,$da,$yr,$hr,$mn);

  #print $endYYYYDDD . "<br>";
	
	##$YYYYDDD = "$l[2]";
  ##################################################################################


  ###################################################
  # Start SQL Dive Search Query for this Expedition #
  ###################################################

  $sql = "SELECT * FROM Dive ";
  $sql .= " WHERE ExpeditionID_FK = " . $RS->Fields('ExpeditionID')->Value . "";
  $sql .= " ORDER BY DiveNumber ";

  Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
  $RSdive = $Conn->Execute($sql);

  if(!$RSdive) {
    die "<BR>Empty Return Set2: RS. for $sql ";
  }

  if( $RSdive->Fields('DiveNumber')->Value > 0 ) {
    $RovName = $RSdive->Fields('RovName')->Value;
    $dive_text = '<BR><b>Edit Dives: </b>';
    while ( !$RSdive->EOF ) {
      $dive_text .= "<a href=\"dive.asp?RovName=" . $RSdive->Fields('RovName')->Value;
      $dive_text .= "&amp;DiveNumber=" . $RSdive->Fields('DiveNumber')->Value . "&amp;edit=yes\" target=\"dive_edit\">";
      $dive_text .= $RSdive->Fields('RovName')->Value;
      
      if ( $RSdive->Fields('DiveNumber')->Value == GetFormValue('qString') || 
           $RSdive->Fields('DiveNumber')->Value == $qDiveNumber ) {
        $dive_text .= "<b>" . $RSdive->Fields('DiveNumber')->Value . "</b>";
      }
      else {
        $dive_text .= $RSdive->Fields('DiveNumber')->Value;
      }
      $dive_text .= "</a> ";
      $dive_text .= " ";

      #print "<br>$dive_text<br>";
      $dive_text .= $delimiter;

      $RSdive->MoveNext;
    }
    $dive_text =~ s/$delimiter$//;
    $dive_text .= " <font color=brown>(These links also show CTD summary information and timeline of framegrabs and coolpix images)</font>";
  }
  else {
    if ($Status eq 'Completed') {
      $dive_text = "<br><br><font color=brown>No dives entered in database</font>";
    }
    else {
      $dive_text = '';
    }
  }

	$RSdive->Close;
 # $RSexdive->Close;
	$dive_text =~ s/$delimiter$//;

  #
  # Data definition text
  #
    my $expd_text = $RS->Fields('Purpose')->Value . $delimiter .
      $RS->Fields('ExpdChiefScientist')->Value . $delimiter .
      $RS->Fields('EquipmentDesc')->Value . $delimiter .
      $RS->Fields('Participants')->Value . $delimiter .
      $RS->Fields('RegionDesc')->Value . $delimiter .
      $RS->Fields('PlannedTrackDesc')->Value . $delimiter .
      $RS->Fields('Accomplishments')->Value . $delimiter .
      $RS->Fields('ScientistComments')->Value . $delimiter .
      $RS->Fields('OperatorComments')->Value . $delimiter .
      $RS->Fields('OtherComments')->Value . $delimiter .
      $RS->Fields('ShipName')->Value .
      $RS->Fields('ShipSeqNum')->Value . $delimiter .
      $RS->Fields('ExpeditionID')->Value;


    my $expdDeviceID = $RS->Fields('DeviceID')->Value;
#print "<br>PostCruise: $expdDeviceID<br>";
  
  #
  # Bold the found text, be careful not to get the inside of href's
  #
  if ( GetFormValue('search') eq 'advanced' ) {
    foreach $q ( @qFields ) {
      next if GetFormValue($q) < 0;
      $str = GetFormValue($q);
      $expd_text =~ s#($str)#\<b\>$1\<\/b\>#ig if $str;
    }
  }
  else {
    $str = GetFormValue('qString');
    $expd_text =~ s#($str)#\<b\>$1\<\/b\>#ig if $str;
  }
  
  $expd_text .= $dive_text;
  
  #
  # Replace any î characters with º
  #
  $expd_text =~ s/î/º/g; 

  my $ExpeditionID = $RS->Fields('ExpeditionID')->Value;
  if ($Status eq 'Completed') {
    if ( $RS->Fields('ShipName')->Value eq 'bhrzn') {
      $StatusString = "<font color=\"red\"><strong>";
      $StatusString .= $RS->Fields('ShipName')->Value  . "</strong></font> &middot; ";
      $StatusString .= "$StartDate &middot; $YYYYDDD ";
    }
    else {
      $StatusString = $RS->Fields('ShipName')->Value  . " &middot; ";
      $StatusString .= "$StartDate &middot; $YYYYDDD ";
    }
    if ($endYYYYDDD - $YYYYDDD > 1) {
      $StatusString .= " &#150; $EndDate &middot; $endYYYYDDD <em>&bull;</em> " ;
    }
    else {
      $StatusString .= " <em>&bull;</em> ";
    }
    $StatusString .= $RS->Fields('ExpdChiefScientist')->Value . " &middot; $Status";
  }
  else {
    $StatusString = "<font color=\"red\">";
    $StatusString .= $RS->Fields('ShipName')->Value . " &middot; ";
    $StatusString .= "$StartDate &middot; $YYYYDDD ";
    if ($endYYYYDDD - $YYYYDDD > 1) {
      $StatusString .= " &#150; $EndDate &middot; $endYYYYDDD <em>&bull;</em> " ;
    }
    else {
      $StatusString .= " <em>&bull;</em> ";
    }
    $StatusString .= $RS->Fields('ExpdChiefScientist')->Value . " &middot; $Status </font>";
    $StatusString .= " &middot; <a href=precruise.asp?step=revise&ShipName=" . $RS->Fields('ShipName')->Value;
    $StatusString .= "&ShipSeqNum=" . $RS->Fields('ShipSeqNum')->Value;
    $StatusString .= "&edit=" . $ExpeditionID . ">";
    $StatusString .= "<img src=\"revise_pre.gif\" border=0></a>";
  }
  ###########################################################
  # Add links to data in the EXPEDITION SEARCH RESULTS text #
  ###########################################################

  if ( $RovName ) {
      my $rdbdlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'rovctddb', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
      my $proclinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'rovctd', expdDeviceID_FK => $expdDeviceID );

    if ( $rdbdlinks ) {

#print "<br>RDBDLINKS:  $rdbdlinks<br>";
#print "<br>Have Step2 ROVCTDDB Datalinks (rdbdlinks).<br>";
      $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>ROVCTD Data:</b> $rdbdlinks</font>" if $rdbdlinks;
#print "<brEXPD Text: " . $expd_text . "<br>";
    }
    elsif ( $proclinks ) {
#print "<br>PROCLINKS:  $proclinks<br>";
#print "<br>Have Step2 ROVCTD Data links (proclinks).<br>";
      $expd_text .= "\n<br><font face=\"helvetica,Ariel\" color=\"#c0c0c0\"><b>ROVCTD Data:</b></font> ";
      $expd_text .= "<font color=\"brown\">";
      $expd_text .= "<A HREF=\"dive.asp?ExpeditionID=$ExpeditionID&add_dive=Continue+to+enter+new+dive+information\" target=\"dive_edit\">";
      $expd_text .= "Add dive(s)</A> to enter dive number(s), start and end times so that additional data products may be produced. (<i>This message is replaced by links after dive(s) have been entered and nightly job has run.</i>)</font>" if $proclinks;
    }
  $rdlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'rovctd', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
#print "<br>RDLINKS:  $rdlinks<br>";
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>ROVCTD Processing:</b> <font size=\"-1\">$rdlinks</font></font>" if $rdlinks;

  }
  
  $tslinks = tapeSummaryLinks( ExpeditionID_FK => $ExpeditionID, SearchDiveNumber => $qDiveNumber );
#print "<br>TSLINKS:  $tslinks<br>";
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Tapes:</b> $tslinks</font>" if $tslinks;
  
  $adlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'annotations', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
#print "<br>ADLINKS:  $adlinks<br>";
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Annotations:</b> $adlinks</font>" if $adlinks;
  
  $sdlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'samplesdb', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
#print "<br>SDLINKS:  $sdlinks<br>";
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Samples:</b> $sdlinks</font>" if $sdlinks;
  
  $fdlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'frameGrabs', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
#print "<br>FDLINKS:  $fdlinks<br>";
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Frame grabs:</b> $fdlinks</font>" if $fdlinks;
  
  $ddlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'digitalImages', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Coolpix images:</b> $ddlinks</font>" if $ddlinks;
  
  $cdlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'CameraLog', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
  my $camdbLinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'camlogdb', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
  
  $cdlinks .= "<b>&middot;</b>" . $camdbLinks if $camdbLinks;
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Camera Log:</b> $cdlinks</font>" if $cdlinks;
  
  $ndlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'Navigation', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
  my $rawNavdbLinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'RawNavdb', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
  #
  $ndlinks .= $rawNavdbLinks . "<br>" if $rawNavdbLinks;
  
  my $cleanNavdbLinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'CleanNavdb', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
  $ndlinks .=  $cleanNavdbLinks . "<b>&middot;</b>" if $cleanNavdbLinks;
#print "<br>NDLINKS:  $ndlinks<br>";
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Navigation:</b> $ndlinks</font>" if $ndlinks;
  
  $auvdlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'auvctdSurvey', expdDeviceID_FK => $expdDeviceID );
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>AUVCTD Survey:</b> $auvdlinks</font>" if $auvdlinks;
  
  ##$expd_text .= "<br><a href=\"addlink.asp\" target=\"addlink\"><img src=\"add_a_link.gif\" border=0></a>";
  
  $d3links_orig = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => '3Dreplay', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
  $d3links = $d3links_orig;
  $d3links =~ s/<input type="checkbox" name="dives" value="\S+" checked>//g;  # Remove checkbox for listing display

  $GElinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'GoogleEarth', SearchDiveNumber => $qDiveNumber, expdDeviceID_FK => $expdDeviceID );
#print "<br>GOOGLINKS:  $GElinks<br>";
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Google Earth:</b> $GElinks </font><font color=\"brown\">(If prompted enter your CanyonHead credentials, e.g.: shore\\&lt;username&gt; &amp; password)</font>" if $GElinks;

  $expd_text .= "<br><br>";
	
	if ( $_[0] eq 'data_table' ) { 
		$d3links = $d3links_orig;
		$d3links = '&nbsp;' unless $d3links;
		$adlinks = '&nbsp;' unless $adlinks;
		$sdlinks = '&nbsp;' unless $sdlinks;
		$fdlinks = '&nbsp;' unless $fdlinks;
		$rdlinks = '&nbsp;' unless $rdlinks;
		$cdlinks = '&nbsp;' unless $cdlinks;
		$ndlinks = '&nbsp;' unless $ndlinks;
		$auvdlinks = '&nbsp;' unless $auvdlinks;
		
		#
		# Check for multiple day cruises & display end date if more than 1 day long
		#
		if ($endYYYYDDD - $YYYYDDD > 1) {
			$DateString = "$StartDate &#150; $EndDate";
			$YYYYDDDString = "$YYYYDDD &#150; $endYYYYDDD";
		}
		else {
			$DateString = "$StartDate";
			$YYYYDDDString = "$YYYYDDD";
		}
		
		#
		# Create Expedition Link, alter if past cruise is still planned
		#
		$ElinkString = $RS->Fields('ExpdChiefScientist')->Value . "<br><br>\n";
		$ElinkString .=	"<a href=\"$this_script?step=3&ShipName=" . $RS->Fields('ShipName')->Value;
		$ElinkString .= "&ShipSeqNum=" . $RS->Fields('ShipSeqNum')->Value . "&edit=" . $RS->Fields('ExpeditionID')->Value;
		$ElinkString .= "\" target=\"edit\">";
		$ElinkString .=	$RS->Fields('ShipName')->Value . $YYYYDDDString . "</a>";
		
		if ( $Status eq 'Planned' ) {
			$ElinkString .=	"\n<br><font color=\"red\">$Status</font>";
			if ( $mydive_text =~ /vnta/ || $mydive_text =~ /tibr/ ) {
			}
			else {
				$adlinks = "<font color=\"brown\">See Missing Links note below.</font>" if $adlinks eq '&nbsp;';
				$sdlinks = "<font color=\"brown\">See Missing Links note below.</font>" if $sdlinks eq '&nbsp;';
			}
		}
		else {
			$ElinkString .=	"\n<br>$Status";
		}
		$ElinkString .=	"<BR>\n";
		
		$mydive_text = $dive_text;
		$mydive_text =~ s/DIVES://;
		
		if ( ($mydive_text eq '' || $mydive_text =~ /No/) && $Status !~ /Canc/i && $Status !~ /Aborted/i 
				&& $RS->Fields('ExpdChiefScientist')->Value ne "Tim Pennington" && $RS->Fields('ShipName')->Value ne "zphr" ) {
		  $mydive_text .= "<br>";
		  $mydive_text .= "<A HREF=\"dive.asp?ExpeditionID=" . $RS->Fields('ExpeditionID')->Value . 
						"&add_dive=Continue+to+enter+new+dive+information\" target=\"dive_edit\">" .
						 "<img src=\"add_dives.gif\" alt=\"Add dive(s)\" border=\"0\"></A>";
		}
		$ElinkString .= $mydive_text;
		
		$tslinks = tapeSummaryLinks( ExpeditionID_FK => $ExpeditionID, SearchDiveNumber => $qDiveNumber );
		$ElinkString .= "\n<br><font face=\"helvetica,Ariel\">Tapes: $tslinks</font>" if $tslinks;
	
		
		$ElinkString .= "<br><p align=\"right\" class=\"smalltext\">" . $RS->Fields('ExpeditionID')->Value . "</p>\n";
		
		# Put a break (dot) after each data link for table display (easier to read)
		$adlinks =~ s/&middot;/&middot;<br>/g;
		$sdlinks =~ s/&middot;/&middot;<br>/g;
		$fdlinks =~ s/^(.)/<b>Video&nbsp;framegrabs<\/b><br>$1/ unless $fdlinks eq '&nbsp;';
		$fdlinks =~ s/&middot;/&middot;<br>/g;
		$ddlinks =~ s/^(.)/<br><br><b>Coolpix&nbsp;images<\/b><br>$1/;
		$ddlinks =~ s/&middot;/&middot;<br>/g;
		$rdlinks =~ s/&middot;/&middot;<br>/g;
		$rdbdlinks =~ s/&middot;/&middot;<br>/g;
		$cdlinks =~ s/&middot;/&middot;<br>/g;
		$ndlinks =~ s/&middot;/&middot;<br>/g;
		$auvdlinks =~ s/&middot;/&middot;<br>/g;
		$d3links =~ s/&middot;/<br>/g;
		$d3links =~ s/checked>\s/checked>/g;
		
		##$d3links = "<form action=\"/3Dreplay/3D.asp\">" . $d3links;
		$d3links =~ s/<font size=\"-2/<br><br><center><font size=\"-2/;
		$d3links .= "</center>";
		##$d3links .= "</form>";
		$has3Dlinks = 1 if $d3links =~ /3D\.asp/;
		
		#
		# Message if can't run Cosmo with GeoVRML
		#
		##my $brows = $Request->ServerVariables(HTTP_USER_AGENT)->{Item};
		##if ( ($brows =~ /mac/i) && $d3links_orig =~ /3D.asp/ ) {
		##	$d3links = "<font color=\"brown\">3D Replay does not yet work on Macintoshes</font>";
		##	$has3Dlinks = 0;	# Flag for display button at bottom of table
		##}
		%>
		<tr>
		<td valign="top" class="dateCell"><font face="helvetica,Ariel" size="-1"><%=$DateString%></font></td>
		<td valign="top"><%=$ElinkString%></td>
		<td valign="top"><font face="helvetica,Ariel" size="-1"><%=$d3links%></font></td>
		<td valign="top"><font face="helvetica,Ariel" size="-2"><%=$adlinks%></font></td>
		<td valign="top"><font face="helvetica,Ariel" size="-2"><%=$sdlinks%></font></td>
		<td valign="top"><font face="helvetica,Ariel" size="-2"><%=$fdlinks%><%=$ddlinks%><%=$auvdlinks%></font></td>
		<td valign="top"><font face="helvetica,Ariel" size="-2"><%=$rdbdlinks%><%=$rdlinks%></font></td>  
		<td valign="top"><font face="helvetica,Ariel" size="-2"><%=$cdlinks%></font></td>
		<td valign="top"><font face="helvetica,Ariel" size="-2"><%=$ndlinks%></font></td>
		</tr>
	<%} else { # List format %>
		<dt><font face="helvetica,Ariel">

		<%= $StatusString%>
		&middot;
		<a href="<%= $this_script%>?step=3&amp;ShipName=<%= $RS->Fields('ShipName')%>&amp;ShipSeqNum=<%= $RS->Fields('ShipSeqNum')%>&amp;edit=<%= $RS->Fields('ExpeditionID')%>" target="edit"><img src="edit_post.gif" border="0" WIDTH="114" HEIGHT="21"></a>
		<font face="helvetica,Ariel" size="-1">
		<a href="<%= $this_script%>?step=display&ExpeditionID=<%= $RS->Fields('ExpeditionID')%>" target="expd">Printable format</a>
		</font>
		</dt></font><dd>
		<font face="helvetica,Ariel" size="-1"><%= $expd_text%></font></dd>

	<% }
	$RS->MoveNext;
} ## End of While loop
if ( $_[0] eq 'data_table' ) { %>
	</table>
	<br>
	<%
	if ( $has3Dlinks == 1 ) { %>
		<input type="submit" class="button blue" name="3D" value="Continue to Select Terrain For All Above Selected 3D Replays.">
	<% } %>
	</form>
	<hr ALIGN="left" WIDTH="30%">


<b>Missing Links:</b> <p class="brownText">Annotations and samples may exist for dives of an 
expedition even though they do not appear in this table.  The links will be provided here after dive
numbers, start, and end times have been entered.  To
enter dive information click on the expedition date(s) link and complete the Postcruise report; Science team
members are encouraged to complete this report as soon as possible following the cruise as this information enhances
the quality of the data.  The final database entry is made by the logistics coordinator based on your input. 
Annotations and samples are entered into their respective databases by the
video lab staff and samples coordinator and will be found if you query those databases directly.  
If you see
that other data links may be missing please contact Mike McCann.</p>
<br>
<br>
<b>Data link information:</b><br>
0. The 3D Replay requires a browser plugin and special software.  See 
<a href="/3Dreplay/geoVRMLreqts.html">3D Replay requirements and FAQ</a> for details.
<br>
1. <a href="/itd/video/AnnotationRetrieval.htm">Video Annotation retrieval</a>
<br>
2. <a href="/samples/">MBARI Sample Collection information</a> and
<a href="/samplesDB/queries/default.asp">Query the Samples Database</a>
<br>
3. <a href="/Ventana/stills/about.html">Background on Frame Grabs</a>
<br>
4. Links named like &quot;Data for dive xxxx (xxx points)&quot; are returned from the Rovctd table and 
include quality flags.  This links will not appear if dive numbers, start, and end times
have not been entered. Please see <a href="/expd/UserDocs/userdocs.htm"> 
Information about ROVCTD data</a>.
<br>
5. <a href="/expd/UserDocs/camlog_format.txt">Camera log file format</a>
<br>
6. <a href="/expd/UserDocs/NavRoadmap.txt">Original roadmap to Navigation data</a>
and <a href="/itd/DataArchives/navigation/index.htm">Navigation Data Archive</a>
<br>
<br>
For more details on MBARI's data archive, please see
<a href="/itd/DataArchives/ArchiveInfo/dataLinks-ships.htm">
Links to files, documentation, and file maintainers</a>.


<%} else { %>
	</dl>
<% } 
my $endBM = new Benchmark;
my $diffBM = timediff($endBM, $startBM);
Win32::ASP::DebugPrint("\nstep2(): Time taken for query and data links = ", timestr($diffBM, 'all'), "\n");

if ($qRovName eq 'mini') {
	%>
	<p><br><br><b>Total of <%= $count%> expeditions found with query:</b>
	<br><font face="helvetica,Ariel" size="-3"><%= $search_sql %> [<%= timestr($diffBM, 'all') %> ]</font></p>
    <br><%
    if ($count == 0) {
	%><h2><font color="#408080"><strong>**You are searching on MiniROV dives. If your query resulted in no matches,
    please <br>consider revising your <a href="https://mww.mbari.org/expd/log/postcruise.asp?search=advanced">EXPD search(es)</a> to include both MBARI and Non-MBARI vessels.<br>
    <br><br>****If you feel your search is correct and the MBARI Expedition DataBase (EXPD) does not <br>reflect recent MiniROV data and/or cruises, please contact:<br>&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;</font><font color="blue">---> Dale Graves (x-1913, grda@mbari.org)</font><font color="#408080"><br>to request that newer MiniROV records be added to the MBARI EXPD.</strong></font>
   </h2>
	<%}
} else {
    %>
	<p><br><br><b>Total of <%= $count%> expeditions found with query:</b>
	<br><font face="helvetica,Ariel" size="-3"><%= $search_sql %> [<%= timestr($diffBM, 'all') %> ]</font></p>
	<%	
}

#$RSdive->Close;
$RS->Close;
$Conn->Close;
$ConnVARS->Close;
$ConnSamples->Close;

}
}	# End step2()
%>


<%
#--------------------------------------------------------------------
#

=head3 step3()

Present cruise info for editing.

Author: Mike McCann

Date Created: 11/18/98

=cut

sub step3 {

#print "<br>Step3<br>";

#
# Open database to get specified expedition
#
open_database($dsn);

#
# Open VARS database to get data for dataLinks()
#
open_vars_database();

#
# Open Samples database to get data for dataLinks()
#
open_samples_database();		# Creates $ConnSamples object as a global variable

# Do an explicit select (not select *) to avoid problems...
if ( GetFormValue('ShipName') && GetFormValue('ShipSeqNum') ) {
  $sql = "SELECT " . join( ", ", @{$Application->{'ExpeditionFields'}} ) . " FROM Expedition\n";
  $sql .= "WHERE ";
  $sql .= "ShipName = '" . GetFormValue('ShipName') . "' ";
  $sql .= "AND ShipSeqNum = " . GetFormValue('ShipSeqNum');
}
elsif ( GetFormValue('ExpeditionID') ) {
  $sql = "SELECT " . join( ", ", @{$Application->{'ExpeditionFields'}} ) . " FROM Expedition\n";
  $sql .= "WHERE ";
  $sql .= "ExpeditionID = " . GetFormValue('ExpeditionID');
}

#print "<br>RS_CONN_STEP#: $sql<br>";

Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: \n$sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
  die "<BR>Empty PC Return Set1: RS. for <br>$sql ";
}

#
# Save the Expedition field values into an assoc. array.  Need to do this because
# Access/ODBC forgets memo fields once they've been read.  (Nice, huh?  This took
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
# Get Chief Scientists from Person table
#
$sql="SELECT * FROM Person WHERE (DisplayChiefSciPickList = 1) ORDER BY LastName, FirstName";
Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: $sql");
$RScs = $Conn->Execute($sql);
if(!$RScs) {
	$RScs->Close;
 	die "<BR>Empty Return Set4: RScs. ";
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
 	die "<BR>Empty Return Set5: RSpi. ";
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

#
# Get Regions list from Region table
#
$sql="SELECT * FROM Region WHERE (DisplayInPickList = 1) ORDER BY RegionName, RegionOwner";
Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: $sql");
$RSrgn = $Conn->Execute($sql);
if(!$RSrgn) {
	$RSrgn->Close;
 	die "<BR>Empty Return Set6: RSrgn. ";
}
@regions = sort split( ',', $RS->Fields('RegionDesc') );
$region = pop @regions;
$rgn_option_list = "<option value=\"\"</option>";
while ( !$RSrgn->EOF ) {
	$rgn_option_list .= "<option value=\"" . $RSrgn->Fields('RegionName')->Value . "\"";
	if ( $region eq $RSrgn->Fields('RegionName')->Value ) {
		$rgn_option_list .= "selected";
		$region = pop @regions;
	}
	$rgn_option_list .= ">" . $RSrgn->Fields('RegionName')->Value . "</option>\n";
	$RSrgn->MoveNext;
}
$RSrgn->Close;

###################################################################################
# Get the actual start and end date info save as session vars for check of original
##################################################################################

  $StartEpoch = SelectEpoch("ScheduledStartDtg");

if ( $RS->Fields('StatCode')->Value eq 'plnd' ) {
	$StartEpoch = SelectEpoch("ScheduledStartDtg");
	$EndEpoch = SelectEpoch("ScheduledEndDtg");
}
else {
	#
	# Make sure that we have an epoch seconds, if the database has Null for the actual start & end times
	# then Start & End epoch are blank, in this case just use scheduled start & end times.
	#
	$StartEpoch = SelectEpoch("StartDtg") || SelectEpoch("ScheduledStartDtg");
	$EndEpoch = SelectEpoch("EndDtg") || SelectEpoch("ScheduledEndDtg");
}
$Session->{'StartEpoch_orig'} = $StartEpoch;
$Session->{'EndEpoch_orig'} = $EndEpoch;


##################################################################
# Construct month, day, year option lists for start & end entry
##################################################################
($Smonth_option_list, $Sday_option_list, $Syear_option_list, $Shrmn) =
	construct_option_lists($StartEpoch);

($Emonth_option_list, $Eday_option_list, $Eyear_option_list, $Ehrmn) = 
	construct_option_lists($EndEpoch);


#
# Check for email address entered if user get e-mail form
#
if ( ! $Session->{'RegisteredUser'} ) {
%>
	<script language="JavaScript" type="text/javascript"><!--
	function Field_Validator(theForm)
	{
	if (theForm.email_addr.value == "")
	{
	  alert("Please enter your MBARI email address.");
	  theForm.email_addr.focus();
	  return (false);
	}
	return (true);
	}
	//-->
	</script>
	<form method="POST" action="<%= $this_script%>" onsubmit="return Field_Validator(this)">
<%
}
else {
%>
	<form method="POST" action="<%= $this_script%>">
<%
}
%>


<!-- SETTING FORM TABLE AND BORDER WIDTHS HERE -->
<table width="830" border="3">
<tr><td> 
<table>
  <td colspan="2" bgcolor="silver" align="center">
  <font face="helvetica,Ariel"><b>Expedition Information</b></font>
</td></td>
</tr>
<tr></tr></td> 
<%
#
# Loop through all the fields and offer for update the ones we change in a Post Cruise
# Once a value is read vie $RS->Fields($f)->Value from a memo field in the database
# you can't read it again!  That is why $ExpeditionValue{$f} is used.
#
foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
	if ( $ExpeditionValue{$f} ) {
		$Session->{$f .'_orig'} = $ExpeditionValue{$f} ;		# Remember original
	}
	else {
		$Session->{$f .'_orig'} = ''							# Clear session space;
	}
	Win32::ASP::DebugPrint("\nstep3():\n$f = " . $ExpeditionValue{$f} );


####print "<br>$f: <font style='bold' color=blue>ExpdVal:</font> $ExpeditionValue{$f}, <font style='bold' color=blue>Original:</font> $Session->{$f .'_orig'}<br>";

	
	next if $f eq 'UpdatedBy';
	next if $f eq 'TimeStamp';
	
	if ($f eq 'Accomplishments' || $f =~ /Comments$/ || $f eq 'Participants' ) {	# Standard text area
		if ($f eq 'OtherComments') {
		$ffff = "DataComments";
			$cell = "<font color=brown>Please add your comment, about anything affecting the data from this expd with your (name and date).</font><br>";
			$cell .= "<textarea name=\"$f\" cols=\"80\" rows=\"3\" wrap>" . 
			$ExpeditionValue{$f} . "</textarea>";
		}
		else {
			$cell = "<textarea name=\"$f\" cols=\"80\" rows=\"3\" wrap>" . 
			$ExpeditionValue{$f} . "</textarea>";
		}
	}
	elsif ($f eq 'StatCode') {					# Drop down list
		$cell = $ExpeditionValue{$f};
		$cell .= "<font color=brown> Enter outcome of cruise -> </font>";
		$cell .= "<select name=\"StatCode\" size=\"1\">\n";
		$cell .= "<option value=\"plnd\" ";
		$cell .= "selected" if $ExpeditionValue{$f} eq 'plnd';
		$cell .= ">Planned</option>\n";
		$cell .= "<option value=\"cmpl\" ";
		$cell .= "selected" if $ExpeditionValue{$f} eq 'cmpl';
		$cell .= ">Completed</option>\n";
		$cell .= "<option value=\"abrt\" ";
		$cell .= "selected" if $ExpeditionValue{$f} eq 'abrt';
		$cell .= ">Aborted</option>\n";
		$cell .= "<option value=\"cncl\" ";
		$cell .= "selected" if $ExpeditionValue{$f} eq 'cncl';
		$cell .= ">Canceled</option>\n</select>\n";
		$cell .= "<font color=brown size=\"-2\"> (Mark as completed if filing postcruise report)</font>";
	}
	elsif ($f eq 'ExpdChiefScientist') {		# Drop down list
		$cell = $ExpeditionValue{$f};
		$cell .= "<font color=brown> If different -> </font>";
		$cell .= "<select name=\"ExpdChiefScientist\" size=\"1\">\n";
		$cell .= $cs_option_list;
		$cell .= "</select>\n";
		$cell .= "<a href=person.asp>Edit person list</a>";
	}
	elsif ($f eq 'ExpdPrincipalInvestigator') {	# Drop down list
		$cell = $ExpeditionValue{$f};
		$cell .= "<font color=brown> If different -> </font>";
		$cell .= "<select name=\"ExpdPrincipalInvestigator\" size=\"1\">\n";
		$cell .= $pi_option_list;
		$cell .= "</select>\n";
	}
	
	elsif ($f eq 'SciObjectivesMet' || $f eq 'AllEquipmentFunctioned' ) {		# yes/no (1/0) Radio buttons
		$cell = "<input type=\"radio\" name=\"$f\" value=\"yes\" ";
		$cell .= "checked" if $ExpeditionValue{$f} eq 'yes';
		$cell .= ">Yes ";
		$cell .= "<input type=\"radio\" name=\"$f\" value=\"no\" ";
		$cell .= "checked" if $ExpeditionValue{$f} eq 'no';
		$cell .= ">No";
		$cell .= " <font color=brown><i>(Chief Scientist to answer this)</i></font>" if $f =~ /^Sci/;
		$cell .= " <font color=brown><i>(Chief ROV pilot to answer this)</i></font>" if $f =~ /^All/;
	}
	elsif ( $f eq 'ScheduledStartDtg' ) {
		$cell = $ExpeditionValue{$f} . " UTC &nbsp;";
	}
	elsif ( $f eq 'ScheduledEndDtg' ) {
		$cell = $ExpeditionValue{$f} . " UTC &nbsp;";
	}
  elsif ( $f eq 'StartDtg' ) {
    if ( $RS->Fields('StatCode')->Value eq 'cmpl' && $ExpeditionValue{$f} ) {
      $cell = $ExpeditionValue{$f} . " UTC &nbsp;";
      $cell .= "<font color=brown>Actual start: </font>";
    }
    else {
      $cell = "<font color=brown>Enter actual start time -> </font>";
    }
    $cell .= "<select name=\"${f}Month\" size=\"1\">$Smonth_option_list</select>\n";
    $cell .= "<select name=\"${f}Day\" size=\"1\">$Sday_option_list</select>\n";
    $cell .= "<select name=\"${f}Year\" size=\"1\">$Syear_option_list</select>\n";
    $cell .= "<input type=\"text\" name=\"${f}Time\" size=\"4\" value=\"$Shrmn\" maxlength=\"4\">\n";
    $cell .= "<font size=-2>(hhmm - Pacific Time)</font>\n\n";
  }
  elsif ( $f eq 'EndDtg' ) {
    if ( $RS->Fields('StatCode')->Value eq 'cmpl' && $ExpeditionValue{$f} ) {
      $cell = $ExpeditionValue{$f} . " UTC &nbsp;";
      $cell .= "<font color=brown>Actual end: </font>";
    }
    else {
      $cell = "<font color=brown>Enter actual end time -> </font>";
    }
    $cell .= "<select name=\"${f}Month\" size=\"1\">$Emonth_option_list</select>\n";
    $cell .= "<select name=\"${f}Day\" size=\"1\">$Eday_option_list</select>\n";
    $cell .= "<select name=\"${f}Year\" size=\"1\">$Eyear_option_list</select>\n";
    $cell .= "<input type=\"text\" name=\"${f}Time\" size=\"4\" value=\"$Ehrmn\" maxlength=\"4\">\n" if $f eq 'EndDtg';
    $cell .= "<font size=-2>(hhmm - Pacific Time)</font>\n\n";
	}
	elsif ( $f eq 'PlannedTrackDesc' ) {
		$cell = "<pre>$ExpeditionValue{$f}</pre>";
	}
	else {
		$cell = $ExpeditionValue{$f};
	}	# End if(
	
	#
	# Replace any î characters with º
	#
	$cell =~ s/î/º/g;
	
	$ffff = $f;
	if ($f eq 'OtherComments') {
	$ffff = "DataComments";
	}
	
	
%><tr>
	<td valign="top" width="30%"><font face="helvetica,Ariel" size="-1"><b><%= $ffff%>
	</b></font></td>
	<td valign="top" width="70%"><font face="helvetica,Ariel" size="-1"><%= $cell%></font></td>
</tr>
	<%
}	# End foreach(

  ###################################################
  # Add links to data in the expedition text        #
  ###################################################
	my $dlinks;
  #print "<br>EXPD: $ExpeditionValue{ExpeditionID}<br>";
  #print "<br>RovName: $RovName<br>";
	##$dlinks = dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => 'rovctddb' );
	##$Response->Write("<tr><td valign=top><font face=\"helvetica,Ariel\" size=-1><b>ROVCTD Data:</b></font></td><td><font face=\"helvetica,Ariel\" size=-1>$dlinks</font></td></tr>") if $dlinks;

  if ( $RovName ) {
    $dlinks = dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => 'rovctd' );
    $Response->Write("<tr><td valign=top><font face=\"helvetica,Ariel\" size=-1><b>ROVCTD Processing:</b></font></td><td><font face=\"helvetica,Ariel\" size=-2>$dlinks</font></td></tr>") if $dlinks;
  }
  $dlinks = dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => 'frameGrabs' );
  $Response->Write("<tr><td valign=top><font face=\"helvetica,Ariel\" size=-1><b>Frame grabs:</b></font></td><td><font face=\"helvetica,Ariel\" size=-1>$dlinks</font></td></tr>") if $dlinks;
  $dlinks = dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => 'digitalImages' );
  $Response->Write("<tr><td valign=top><font face=\"helvetica,Ariel\" size=-1><b>Coolpix images:</b></font></td><td><font face=\"helvetica,Ariel\" size=-1>$dlinks</font></td></tr>") if $dlinks;
  $dlinks = dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => 'CameraLog' ) . dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => 'camlogdb' );
  $Response->Write("<tr><td valign=top><font face=\"helvetica,Ariel\" size=-1><b>Camera Log:</b></font></td><td><font face=\"helvetica,Ariel\" size=-1>$dlinks</font></td></tr>") if $dlinks;

  $dlinks = dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => 'Navigation' );
  my $rawNavdbLinks = dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => 'RawNavdb' );
  $dlinks .= $rawNavdbLinks if $rawNavdbLinks;
  my $cleanNavdbLinks = dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => 'CleanNavdb' );
  $dlinks .= "<b>&middot;<br></b>" . $cleanNavdbLinks if $cleanNavdbLinks;
  $Response->Write("<tr><td valign=top><font face=\"helvetica,Ariel\" size=-1><b>Navigation:</b></font></td><td><font face=\"helvetica,Ariel\" size=-1>$dlinks</font></td></tr>") if $dlinks;

  $dlinks = dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => '3DReplay' );
  $Response->Write("<tr><td valign=top><font face=\"helvetica,Ariel\" size=-1><b>3D Replay:</b></font></td><td><font face=\"helvetica,Ariel\" size=-1>$dlinks</font></td></tr>") if $dlinks;
  $dlinks = dataLinks( ExpeditionID_FK => $ExpeditionValue{ExpeditionID}, DataType => 'GoogleEarth' );
  $Response->Write("<tr><td valign=top><font face=\"helvetica,Ariel\" size=-1><b>Google Earth:</b></font></td><td><font face=\"helvetica,Ariel\" size=-1>$dlinks</font></td></tr>") if $dlinks;


%>
</table>

<%
###################################################
# Now check and ask for dive info
#
# THIS IS WHEN YOU EDIT POSTCRUISE DIVES
###################################################

$sql = "SELECT Expedition.DeviceID AS expdDeviceID, * FROM Dive\n";
$sql .= "JOIN Expedition ON Dive.ExpeditionID_FK = Expedition.ExpeditionID\n";
$sql .= "WHERE ";
$sql .= "Dive.ExpeditionID_FK = " . GetFormValue('edit');
$sql .= " ORDER BY DiveNumber";
Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: $sql");
$RSdive = $Conn->Execute($sql);
if(!$RSdive) {
   $Response->Write("<BR>Empty Return Set7: RSdive. for <br>$sql ");
}
%>
<form method="POST" action="<%= $this_script%>">
<table width ="820"><br><br><hr width="100%" size="2" color="black" align="center">
<tr><td>
<table>
  <td colspan="2" bgcolor="silver" align="center">
  <font face="helvetica,Ariel"><b>Dive Information</b></font>
</td></td></tr>
  <tr><td colspan="2"><font face="helvetica,Ariel" color="brown" size="-1">
  Dives are defined as part of the postcruise report process. Click on 
  the appropriate button.  Dive start and end times must be entered to be
  able to get ROVCTDO data from the database, have the frame grabs organized 
  by dive number, and to view 3D Replays.<br><br><br>
  </font></td></tr>
<%
while ( !$RSdive->EOF ) {     # List existing dives in the database
  my $editlink = "";
  for (my $i = 0; $i < 20; $i++) {
    $editlink .= "&nbsp;";
  }

#print "<br>EDITLINK1: " . $editlink  . "<br>";
#print "<br>DIVE: " . $RSdive->Fields('RovName')->Value .  $RSdive->Fields('DiveNumber')->Value . "<br>";

  # SEND THIS OVER TO DIVE.ASP
  $editlink .= "<a href=dive.asp?RovName=" . $RSdive->Fields('RovName')->Value . "\&";
  $editlink .= "DiveNumber=" . $RSdive->Fields('DiveNumber')->Value . "&edit=yes target=\"dive_edit\">"; 
  $editlink .= "<img src=\"edit_this_dive.gif\" class=img border=0></a>";

#print "<br>EDITLINK_Dive.asp: " . $editlink  . "<br>";

  
  foreach $f ( @{$Application->{'DiveFields'}} ) {
    if ( $RSdive->Fields($f)->Value ) {
      $Session->{$f .'_orig'} = $RSdive->Fields($f)->Value ;    # Remember original
    }
    else {
      $Session->{$f .'_orig'} = ''              # Clear session space;
    }
    %>
    <tr>
    <%if ( $f eq 'BriefAccomplishments' ) {   # Geeze, this is a mess! %>
      <td valign="top" width="200"><font face="helvetica,Ariel" size="-1"><b>
      Dive Accomplishments <br>and Comments:</b></font></td>

    <%} else { %>
      <td valign="top"><font face="helvetica,Ariel" size="-1"><b><%= $f%></b></font></td>
    <% }
    if ( $f eq 'RovName') { %>
      <td valign="top"><font face="helvetica,Ariel" size="-1">
      <%= $RSdive->Fields($f)->Value%></font></td>
    <%
    }
    elsif ( $f eq 'DiveStartDtg' ) { %>
      <td valign="bottom"><font face="helvetica,Ariel" size="-1">
      <% if ( $RSdive->Fields($f)->Value ) {%>
            <%= $RSdive->Fields($f)->Value%> UTC<%
         }
         else { %>
            Edit dive to enter a dive start time
       <%}%>
      <%= $editlink%>
      </font></td><%
    }
    elsif ( $f eq 'DiveEndDtg' ) {
      if ( $RSdive->Fields($f)->Value ) {%>
        <td valign="top"><font face="helvetica,Ariel" size="-1">
        <%= $RSdive->Fields($f)->Value%> UTC</font></td><%
      }
      else {%>
        <td><font face="helvetica,Ariel" color="brown" size="-2">
        Edit dive to enter a dive end time</font></td>
      <%
      }
    }
    else {%>
      <td valign="top"><font face="helvetica,Ariel" size="-1">
      <%= $RSdive->Fields($f)->Value%></font></td>
    <%}%>
    
    </tr><%
  } # End foreach( 

#################################
#
#            LINKS                
#
#################################            

  ## ROVCTD DB  
  my $dlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value,
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => 'rovctddb');
  if ( $dlinks ) { 
    #print "<br>HAVE Step3 ROVCTDDB ROVCTD Datalinks for $dlinks.<br>";%>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>ROVCTD Data:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $dlinks %></font></td></tr>
  <%}
  
  ## TAPES 
  $tslinks = tapeSummaryLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value,
                               RovName_FK => $RSdive->Fields('RovName')->Value );
  if ( $tslinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Tapes:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $tslinks %></font></td></tr>
  <%}
  
  ## ANNOTATIONS  
  $dlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value,
                       RovName_FK => $RSdive->Fields('RovName')->Value,
                       expdDeviceID_FK => $RSdive->Fields('expdDeviceID')->Value,
                       DataType => 'annotations');
  if ( $dlinks ) { %>
   <input type="hidden" name="DeviceID" value="<%= $RSdive->Fields('expdDeviceID')->Value%>">
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Annotations:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $dlinks %></font></td></tr>
  <%}

  ## SAMPLES DB  
  $dlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value,
                       RovName_FK => $RSdive->Fields('RovName')->Value,
                       expdDeviceID_FK => $RSdive->Fields('expdDeviceID')->Value,
                       DataType => 'samplesdb');
  if ( $dlinks ) { %>
    <input type="hidden" name="DeviceID" value="<%= $RSdive->Fields('expdDeviceID')->Value%>">
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Samples:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $dlinks %></font></td></tr>
  <%}

  ## FRAMEGRABS 
  my $flinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
                          RovName_FK => $RSdive->Fields('RovName')->Value,
                          DataType => 'frameGrabs');
  if ( $flinks ) { 
#print "<br>HAVE Step3 flinks.<br>";%>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Frame Grabs:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $flinks %></font></td></tr>
  <%}

  ## DIGITAL IMAGES  
  my $clinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
                          RovName_FK => $RSdive->Fields('RovName')->Value,
                          DataType => 'digitalImages');
  if ( $clinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Coolpix images:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $clinks %></font></td></tr>
  <%}
  
  ## CAMLOG DB  
  my $camlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
                            RovName_FK => $RSdive->Fields('RovName')->Value,
                            DataType => 'camlogdb');
  if ( $camlinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Camlog:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $camlinks %></font></td></tr>
  <%}
  



  ## NAVLINKS
  my $navlinks = '';
  my $rawNavlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => 'RawNavdb');
  $navlinks .= $rawNavlinks if $rawNavlinks;
  my $cleanNavlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => 'CleanNavdb');
  $navlinks .= "<b>&middot;</b>" . $cleanNavlinks if $cleanNavlinks;
  if ( $navlinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Navigation:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $navlinks %></font></td></tr>
  <%}




  ## 3D REPLAY AND GOOGLE EARTH LINKS ##

  my %dd = $RSdive->Fields('RovName')->Value . $RSdive->Fields('DiveNumber')->Value;
  foreach my $dd ( sort keys %dd ) { 

    my $d3links = dataLinks( ExpeditionID_FK => $RSdive->Fields('ExpeditionID_FK')->Value,
                             DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
                             RovName_FK => $RSdive->Fields('RovName')->Value,
                             DataType => '3DReplay');

    if ( $d3links ) {   
        my @TD = split('&middot;', $d3links);
        foreach my $TDR (@TD) {
          if ($TDR =~ $dd) {
            $TDR =~ s/\$dd.*//;
            $TDR =~ s/\.$dd*//s;
            %><tr><td><font face="helvetica,Ariel" size="-1"><b>3D Replay:</b></font></td>
            <td><font face="helvetica,Ariel" size="-1"><%= $TDR %></font></td></tr><%  
          }
      }   
    } 
  }

  my $r = uc substr( $RSdive->Fields('RovName')->Value, 0, 1 );
  my %dn = $r . $RSdive->Fields('DiveNumber')->Value . " Tour";
    foreach my $dn ( sort keys %dn ) { 
    ## GOOGLE EARTH LINKS
    my $GElinks = dataLinks( ExpeditionID_FK => $RSdive->Fields('ExpeditionID_FK')->Value, 
                             DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value,  
                             RovName_FK => $RSdive->Fields('RovName')->Value,
                             DataType => 'GoogleEarth');

    if ( $GElinks ) { 
        my @GE = split('&middot;', $GElinks);
        foreach my $GL (@GE) {
          if ($GL =~ $dn) {
            $GL =~ s/\$dn.*//;
            $GL =~ s/\.$dn*//s;
            %><tr><td><font face="helvetica,Ariel" size="-1"><b>Google Earth:</b></font></td>
            <td><font face="helvetica,Ariel" size="-1"><%= $GL %></font></td></tr> <% 
          }
        }   
    } 
}
%><tr><td>&nbsp;</td><td></td></tr><%
  $RSdive->MoveNext;
  
} # End while ( !$RSdive->EOF

if ( $RSdive->EOF == $RSdive->BOF ) { %>
<tr><td>
<font face="helvetica,Ariel" size="-1">No dives entered for this expedition</font>
</td></tr>
<% }
# END OF DIVE INFORMATION TABLE
%>

</table>  

<hr width="75%" align="center">

<br>
<input type="hidden" name="step" value="6"> <!-- Step 6 doesn't exist.  Simply calls to through steps again for processing.-->
<input type="hidden" name="ExpeditionID_FK" value="<%= $RS->Fields('ExpeditionID')->Value%>">
<input type="hidden" name="ExpeditionID" value="<%= $RS->Fields('ExpeditionID')->Value%>">
<input type="hidden" name="ShipName" value="<%= $RS->Fields('ShipName')->Value%>">
<input type="hidden" name="RovName" value="<%= $RSdive->Fields('RovName')->Value%>">
<input type="hidden" name="ShipSeqNum" value="<%= $RS->Fields('ShipSeqNum')->Value%>">
<input type="hidden" name="DeviceID" value="<%= $RSdive->Fields('expdDeviceID')->Value%>">
<input type="hidden" name="edit" value="<%= $RS->Fields('ExpeditionID')->Value%>">

<% if ( $Session->{'RegisteredUser'} ) { %> 
  <strong><font face="helvetica,Ariel" size="-1" color="#408080">Send postcruise updates email to 
  <input type="text" name="post_email" value="postcruise" size="9" maxlength="15">
  @mbari.org.</font></strong></p>
  <br>
  <center>
  <input type="submit" class="button green" name="finish" value="Update Database with Above Changes and Continue with Dive Information"> 
  <br><br>
  <input type="submit" class="button blue" name="finish" value="Done with this Expedition. Update Changes Made Above.">
  <br><br>
  <input type="reset" class="button gray" name="reset" value="Reset Form Fields"><br>
  </center>
<%}
else {
%>
  <font face="helvetica,Ariel" size="-1"><b>Your MBARI email address:</b>
  <input type="text" name="email_addr" value size="9">@mbari.org</font>
  <br><br>
  <center>
  <input type="submit" class="button blue" name="finish" value="Continue with this Expedition and Enter New Dive Information">
  <br>
  <font face="helvetica,Ariel" size="-1">
  Email is sent only if you have updated the above Expedition Information.
  <br>
  You will get a blind copy of the mail that is sent to <%=$Application->{'logistics_email'}%>.
  </font>
  <br><br>
  <input type="submit" class="button blue" name="finish" value="Done with this Expedition. Submit to Logistics Coordinator.">
  <br><br>
  <input type="reset" class="button gray" name="reset" value="Reset Form Fields"><br>
  </center><br>

<%}
%>  
</form>
</td></tr>
</table>
<%
$RSdive->Close;
$RS->Close;
$Conn->Close;
$ConnVARS->Close;
$ConnSamples->Close;
} # End step3()
%>



<%
#--------------------------------------------------------------------
#

=head3 step4()


Return list of cruises from query specified in step1() or advanced_search()
SPECIFICALLY For NON-MBARI SHIPS ($EXPD = 0, $RovName = miniROV) searches.
If argument of 'data_table' passed then return results as a table.

Author: Karen Salamy

Date Created: 11/16/98

=cut

sub step4 {

my $startBM = new Benchmark;

#
# Open database to get do query and produce list of  expeditions
#
open_database($dsn);    # Creates $Conn object as a global variable

#
# Open VARS database to get data for dataLinks()
#
open_vars_database();   # Creates $ConnVARS object as a global variable

#
# Open Samples database to get data for dataLinks()
#
open_samples_database();    # Creates $ConnSamples object as a global variable

$conjunction = GetFormValue('Conjunction');
my $has3Dlinks = 0;

#
# Copy form values into lookup list so that we can accept them from either
# step1() or advanced_search()
#
@qFields = ( 'qShipName', 'qDiveNumber', 'qStatusCode', 'qExpeditionID_FK', 
             'qDiveChiefScientist', 'qDiveStartDtg', 'qDiveEndDtg', 'qYYYYDDD' );

if ( GetFormValue('search') eq 'advanced' ) {
  foreach $q ( @qFields ) {
    $qvalue{$q} = GetFormValue($q) if GetFormValue($q);
  }
}
else {
  $conjunction = 'OR';
  foreach $q ( @qFields ) {
    $qvalue{$q} = GetFormValue('qString') if GetFormValue('qString');
  }
}

Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
  die "<BR>Empty PC Return Set3: Step4 RS for $sql ";
}

#
# Header for returned expeditions, save sql for display at end of list
#
my $search_sql = $sql;
if ( $num_expd > 0 ) {
%>
<h3>Displaying <%= $num_expd%> 
<% if ( $num_expd > 1 ) { %>
expeditions
<% } 
else { %>
expedition
<%}
if ($_[0] eq 'data_table') { %>
 in table format
 <% if ($num_expd > 10) { %>
  (all data must transfer before table is shown)
 <%}%>
<%}%>
...</h3> 
<% }

#
# Now cursor through the record set and print the expeditions
#
if ($_[0] eq 'data_table') { %>
  <form action="/3Dreplay/3D.asp" target="3Dreplay">
  <table BORDER="1">
  <tr>
  <th BGCOLOR="#00C0C0">Date</th>
  <th BGCOLOR="#00C0C0">Expedition Report</th>
  <th BGCOLOR="#00C0C0">3D Replay<sup>0</sup></th>
  <th BGCOLOR="#00C0C0">Annotations<font COLOR="#404040"><sup>1</sup></font></th>
  <th BGCOLOR="#00C0C0">Samples<font COLOR="#404040"><sup>2</sup></font></th>
  <th BGCOLOR="#00C0C0">Images<font COLOR="#404040"><sup>3</sup></font></th>
  <th BGCOLOR="#00C0C0">ROVCTD <font COLOR="#404040"><sup>4</sup></font></th>
  <th BGCOLOR="#00C0C0">Camera Log<font COLOR="#404040"><sup>5</sup></font></th>
  <th BGCOLOR="#00C0C0">Navigation<font COLOR="#404040"><sup>6</sup></font></th>
  </tr>
<%} else { %>
  <dl>
<%}
$delimiter = '<font face="helvetica,Ariel"> <b>&middot;</b> </font>';
$delimiter .= "<br>\n" if ($_[0] eq 'data_table');

$count = 0;
while ( !$RS->EOF ) {
  $count++;
  $Status = $RS->Fields('StatCode')->Value;
  $Status = "Aborted" if $Status eq 'abrt';
  $Status = "Completed" if $Status eq 'cmpl';
  $Status = "Canceled" if $Status eq 'cncl';
  $Status = "Planned" if $Status eq 'plnd';
  
  my @l = ($Status eq 'Completed') ?  split(' ', $RS->Fields('DiveStartDtg')->Value) : 
    split(' ', $RS->Fields('DiveStartDtg')->Value);
  $StartDate = $l[0];
  my ($mo, $da, $yr) = split('/', $StartDate);
  my ($hr, $mn) = split(':', $l[1]);
  $ahr = ($hr + 12) if $l[2] =~ /PM/i;
  if ($ahr < 12) { $clock = "AM"; }else {$clock = "PM";}
  my $starthrmn = "$hr:$mn$clock";  
  my $YYYYDDD = toYDfromDate($mo,$da,$yr,$ahr,$mn);
  
  my @e = ($Status eq 'Completed') ?  split(' ', $RS->Fields('DiveEndDtg')->Value) : 
    split(' ', $RS->Fields('DiveEndDtg')->Value);
  $EndDate = $e[0];
  ($hr, $mn) = split(':', $e[1]);
  $ehr = ($hr + 12) if $e[2] =~ /PM/i;
  if ($ehr < 12) { $clock = "AM"; }else {$clock = "PM";}
  my $endhrmn = "$hr:$mn$clock";
  my $endYYYYDDD = toYDfromDate($mo,$da,$yr,$ehr,$mn);
  

  if ( $RS->Fields('DiveID')->Value ) {
    $DiveID = $RS->Fields('DiveID')->Value;
  }
  if ( $RS->Fields('DeviceID')->Value ne null ) {
    $DeviceID = $RS->Fields('DeviceID')->Value;
  }
  if ( $RS->Fields('DiveChiefScientist')->Value ne null ) {
    $DiveChiefScientist = $RS->Fields('DiveChiefScientist')->Value;
  }
  if ( $RS->Fields('RovName')->Value ne null ) {
    $RovName = $RS->Fields('RovName')->Value;
  }
  if ( $RS->Fields('DiveNumber')->Value ne null ) {
    $DiveNumber = $RS->Fields('DiveNumber')->Value;
  }
  if ( $RS->Fields('DiveStartDtg')->Value ne null ) {
    $DiveStartDtg = $RS->Fields('DiveStartDtg')->Value;
  }
  if ( $RS->Fields('DiveEndDtg')->Value ne null ) {
    $DiveEndDtg = $RS->Fields('DiveEndDtg')->Value;
  }
  if ( $RS->Fields('ExpeditionID_FK')->Value ne null ) {
    $ExpeditionID_FK = $RS->Fields('ExpeditionID_FK')->Value;
  } 

  #
  # Find any dives for this expedition, need to do combined query with ROV Name
  #
  if ( GetFormValue('qShipName') eq 'othr' ) {
    if ($DCS) {
      $DiveChiefSci = (GetFormValue('qExpdChiefScientist'));
      $DiveChiefScientist = $DiveChiefSci;        
    }
  } 
  else {
        die "<BR>Empty Return Set STEP4: Non-MBARI Ship for <br>$sql ";
  } 

################################

  Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
  #$RSdive = $Conn->Execute($sql);   ##### Removed / changed all linking to RS instead of RSdive.
if ( $DiveNumber > 0) {
    $dive_text = '<br><font color=blue><strong>Cannot Edit Dive:  ';

    if ( $ExpeditionID_FK == 0 ) {
          if ( $RS->Fields('DiveNumber')->Value == GetFormValue('qString') || $RS->Fields('DiveNumber')->Value == $qDiveNumber ) {   ##### Removed / changed all linking to RS instead of RSdive.
              $dive_text .= $RS->Fields('RovName')->Value;
              $dive_text .= "<b>" . $DiveNumber . "</b></font></strong>";
              $dive_text .= "  ";
              $dive_text .= $delimiter;
          } 
          else {
              $dive_text .= $RS->Fields('RovName')->Value;
              $dive_text .= "<b>" . $DiveNumber . "</b></font></strong>";
              $dive_text .= "  <br><br>";
              $dive_text .= $delimiter;
          }
#          else {
#            # Loop through the Dives on MBARI ships 
#            for ($d=0; $d<1; $d++) {
#              $dive_text .= "<a href=\"dive.asp?RovName=" . $RovName;
#              $dive_text .= "&amp;DiveNumber=" . $DiveNumber . "&amp;edit=yes\" target=\"dive_edit\">";
#              $dive_text .= $RS->Fields('RovName')->Value;   ##### Removed / changed all linking to RS instead of RSdive.
#              $dive_text .= "<b>" . $DiveNumber . "</b>";
#              $dive_text .= "</a> ";
#              $dive_text .= " ";
#              $dive_text .= $delimiter;
#             }
#          }
#          $dive_text =~ s/$delimiter$//;
#          $dive_text .= " <font color=brown>(These links also show CTD summary information and timeline of framegrabs and coolpix images)</font><br><br>";
    }
  else {
    if ($Status eq 'Completed') {
      $dive_text = "<br><br><font color=brown>No dives entered in database</font>";
    }
    else {
      $dive_text = '';
    }
  }  # End of inner If statement
}  # End of outer If statement

  #$RSdive->Close;

#########################################

  $dive_text =~ s/$delimiter$//;

  #
  # Data definition text
  #
  my $expd_text = 
    "DiveID: " . $DiveID . $delimiter .
    "DeviceID: " . $DeviceID . $delimiter .
    "RovName: " . $RovName . $delimiter .
    "DiveNumber: " . $DiveNumber . $delimiter .
    "DiveChiefScientist: " . $DiveChiefScientist . $delimiter .
    "DiveStart: " . $DiveStartDtg . $delimiter .
    "DiveEnd: " . $DiveEndDtg . $delimiter .
    "ExpeditionID: " . $ExpeditionID_FK;

  #
  # Bold the found text, be careful not to get the inside of href's
  #
  if ( GetFormValue('search') eq 'advanced' ) {
    foreach $q ( @qFields ) {
      next if GetFormValue($q) < 0;
      $str = GetFormValue($q);
      $expd_text =~ s#($str)#\<b\>$1\<\/b\>#ig if $str;
    }
  }
  else {
    $str = GetFormValue('qString');
    $expd_text =~ s#($str)#\<b\>$1\<\/b\>#ig if $str;
  }
  
  # Combine both expd and dive text together
  $expd_text .= $dive_text;
  
  #
  # Replace any î characters with º
  #
  $expd_text =~ s/î/º/g;

  if ($Status eq 'Completed') {
    $StatusString = "<font color=\"red\"><strong>";
    $StatusString .= "Non-MBARI Ship </strong></font> &middot; ";
    $StatusString .= "$StartDate &middot; $starthrmn &middot; $YYYYDDD ";
    if ($endYYYYDDD - $YYYYDDD < 1) {
      #$StatusString .= "$StartDate &middot; $starthrmn&#150;$endhrmn &middot; $endYYYYDDD <em>&bull;</em> " ;
      $StatusString .= " &#150; $EndDate - $endhrmn &middot; $endYYYYDDD <em>&bull;</em> " ;
    }
    else {
      $StatusString .= "$StartDate &middot; $YYYYDDD  &#150; $EndDate &middot; $endYYYYDDD <em>&bull;</em> ";
    }
    $StatusString .= $DiveChiefScientist . " &middot; $Status ";
  } 
  else {
    $StatusString = "<font color=\"red\"><strong>";
    $StatusString .= "Non-MBARI Ship </strong></font> &middot; ";
    $StatusString .= "$StartDate - $starthrmn &middot; $YYYYDDD ";
    if ($endYYYYDDD - $YYYYDDD > 1) {
      $StatusString .= " &#150; $EndDate - $endhrmn &middot; $endYYYYDDD <em>&bull;</em> " ;
    }
    else {
      $StatusString .= " <em>&bull;</em> ";
    }
    $StatusString .= $DiveChiefScientist . " &middot; $Status </font>";
    $StatusString .= " &middot; <a href=precruise.asp?step=revise&RovName=" . $RovName;
    $StatusString .= "&edit=" . $DiveID . ">";
    $StatusString .= "<img src=\"revise_pre.gif\" border=0></a>";
  } 

  #
  # Add links to data in the expedition text
  #
  
  my $RovName = $RS->Fields('RovName')->Value;
  if ( $RovName ) {
      my $rdbdlinks = dataLinks ( ExpeditionID_FK => 0, DataType => 'rovctddb', SearchDiveNumber => $qDiveNumber );
      my $proclinks = dataLinks( ExpeditionID_FK => 0, DataType => 'rovctd' );

    if ( $rdbdlinks ) {
#print "<br>Have Step4 ROVCTDDB Datalinks.<br>";
      $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>ROVCTD Data:</b> $rdbdlinks</font>";
    }
    elsif ( $proclinks ) {
#print "<br>Have Step4 ROVCTD Data links (proclinks).<br>";
      $expd_text .= "\n<br><font face=\"helvetica,Ariel\" color=\"#c0c0c0\"><b>ROVCTD Data:</b></font> ";
      $expd_text .= "<font color=\"brown\">";
      $expd_text .= "<A HREF=\"dive.asp?ExpeditionID=0&add_dive=Continue+to+enter+new+dive+information\" target=\"dive_edit\">";
      $expd_text .= "Add dive(s)</A> to enter dive number(s), start and end times so that additional data products may be produced. (<i>This message is replaced by links after dive(s) have been entered and nightly job has run.</i>)</font>";
    }

    $rdlinks = dataLinks( ExpeditionID_FK => 0, DataType => 'rovctd', SearchDiveNumber => $qDiveNumber ); 
#print "<br>Have Step4 ROVCTD Data links (rdlinks).<br>";
    $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>ROVCTD Processing:</b> <font size=\"-2\">$rdlinks</font></font>" if $rdlinks;
  }
  
  $tslinks = tapeSummaryLinks( ExpeditionID_FK => $ExpeditionID, SearchDiveNumber => $qDiveNumber );
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Tapes:</b> $tslinks</font>" if $tslinks;
  $adlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'annotations', SearchDiveNumber => $qDiveNumber );
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Annotations:</b> $adlinks</font>" if $adlinks;
  $sdlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'samplesdb', SearchDiveNumber => $qDiveNumber );
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Samples:</b> $sdlinks</font>" if $sdlinks;
  $fdlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'frameGrabs', SearchDiveNumber => $qDiveNumber );
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Frame grabs:</b> $fdlinks</font>" if $fdlinks;
  $ddlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'digitalImages', SearchDiveNumber => $qDiveNumber );
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Coolpix images:</b> $ddlinks</font>" if $ddlinks;
  $cdlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'CameraLog', SearchDiveNumber => $qDiveNumber );
  my $camdbLinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'camlogdb', SearchDiveNumber => $qDiveNumber );
  $cdlinks .= "<b>&middot;</b>" . $camdbLinks if $camdbLinks;
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Camera Log:</b> $cdlinks</font>" if $cdlinks;
  $ndlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'Navigation', SearchDiveNumber => $qDiveNumber );
  
  my $rawNavdbLinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'RawNavdb', SearchDiveNumber => $qDiveNumber );
  $ndlinks .= "<b>&middot;</b>" . $rawNavdbLinks if $rawNavdbLinks;
  my $cleanNavdbLinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'CleanNavdb', SearchDiveNumber => $qDiveNumber );
  $ndlinks .= "<b>&middot;</b>" . $cleanNavdbLinks if $cleanNavdbLinks;
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Navigation:</b> $ndlinks</font>" if $ndlinks;
  
  $auvdlinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'auvctdSurvey');
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>AUVCTD Survey:</b> $auvdlinks</font>" if $auvdlinks;
  
  $d3links_orig = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => '3Dreplay', SearchDiveNumber => $qDiveNumber );
  $d3links = $d3links_orig;
  $d3links =~ s/<input type="checkbox" name="dives" value="\S+" checked>//g;  # Remove checkbox for listing display
  


  #
  # Message if browser can't run Cosmo with GeoVRML
  #
  
  $GElinks = dataLinks( ExpeditionID_FK => $ExpeditionID, DataType => 'GoogleEarth', SearchDiveNumber => $qDiveNumber );
  $expd_text .= "\n<br><font face=\"helvetica,Ariel\"><b>Google Earth:</b> $GElinks </font><font color=\"brown\">(If prompted enter your CanyonHead credentials, e.g.: shore\\&lt;username&gt; &amp; password)</font>" if $GElinks;

  $expd_text .= "<br>";
  
  if ( $_[0] eq 'data_table' ) { 
    $d3links = $d3links_orig;
    $d3links = '&nbsp;' unless $d3links;
    $adlinks = '&nbsp;' unless $adlinks;
    $sdlinks = '&nbsp;' unless $sdlinks;
    $fdlinks = '&nbsp;' unless $fdlinks;
    $rdlinks = '&nbsp;' unless $rdlinks;
    $cdlinks = '&nbsp;' unless $cdlinks;
    $ndlinks = '&nbsp;' unless $ndlinks;
    $auvdlinks = '&nbsp;' unless $auvdlinks;
    
    #
    # Check for multiple day cruises & display end date if more than 1 day long
    #
    if ($endYYYYDDD - $YYYYDDD > 1) {
      $DateString = "$StartDate &#150; $EndDate";
      $YYYYDDDString = "$YYYYDDD &#150; $endYYYYDDD";
    }
    else {
      $DateString = "$StartDate";
      $YYYYDDDString = "$YYYYDDD";
    }
    
    #
    # Create Expedition Link, alter if past cruise is still planned
    #
    $ShipName = "Non-MBARI Ship";
    $ElinkString = $RSdive->Fields('DiveChiefScientist')->Value . "<br><br>\n";
    $ElinkString .= "<a href=\"$this_script?step=5&RovName=" . $RovName;
    $ElinkString .= "&ShipSeqNum=" . $RSdive->Fields('DiveID')->Value . "&edit=" . $RSdive->Fields('ExpeditionID_FK')->Value;
    $ElinkString .= "\" target=\"edit\">";
    $ElinkString .= $RSdive->Fields('ShipName')->Value . $YYYYDDDString . "</a>";
    
    if ( $Status eq 'Planned' ) {
      $ElinkString .= "\n<br><font color=\"red\">$Status</font>";
      if ( $mydive_text =~ /vnta/ || $mydive_text =~ /tibr/ ) {
      }
      else {
        $adlinks = "<font color=\"brown\">See Missing Links note below.</font>" if $adlinks eq '&nbsp;';
        $sdlinks = "<font color=\"brown\">See Missing Links note below.</font>" if $sdlinks eq '&nbsp;';
      }
    }
    else {
      $ElinkString .= "\n<br>$Status";
    }
    $ElinkString .= "<BR>\n";
    
    $mydive_text = $dive_text;
    $mydive_text =~ s/DIVES://;
    
    if ( ($mydive_text eq '' || $mydive_text =~ /No/) && $Status !~ /Canc/i && $Status !~ /Aborted/i 
        && $RSdive->Fields('DiveChiefScientist')->Value ne "Tim Pennington" && $RSdive->Fields('ShipName')->Value ne "zphr" ) {
      $mydive_text .= "<br>";
      $mydive_text .= "<A HREF=\"dive.asp?ExpeditionID=" . $RSdive->Fields('ExpeditionID_FK')->Value . 
            "&add_dive=Continue+to+enter+new+dive+information\" target=\"dive_edit\">" .
             "<img src=\"add_dives.gif\" alt=\"Add dive(s)\" border=\"0\"></A>";
    }
    $ElinkString .= $mydive_text;
    
    $tslinks = tapeSummaryLinks( ExpeditionID_FK => $ExpeditionID, SearchDiveNumber => $qDiveNumber );
    $ElinkString .= "\n<br><font face=\"helvetica,Ariel\">Tapes: $tslinks</font>" if $tslinks;
  
    
    $ElinkString .= "<br><p align=\"right\" class=\"smalltext\">" . $RS->Fields('ExpeditionID')->Value . "</p>\n";
    
    # Put a break (dot) after each data link for table display (easier to read)
    $adlinks =~ s/&middot;/&middot;<br>/g;
    $sdlinks =~ s/&middot;/&middot;<br>/g;
    $fdlinks =~ s/^(.)/<b>Video&nbsp;framegrabs<\/b><br>$1/ unless $fdlinks eq '&nbsp;';
    $fdlinks =~ s/&middot;/&middot;<br>/g;
    $ddlinks =~ s/^(.)/<br><br><b>Coolpix&nbsp;images<\/b><br>$1/;
    $ddlinks =~ s/&middot;/&middot;<br>/g;
    $rdlinks =~ s/&middot;/&middot;<br>/g;
    $rdbdlinks =~ s/&middot;/&middot;<br>/g;
    $cdlinks =~ s/&middot;/&middot;<br>/g;
    $ndlinks =~ s/&middot;/&middot;<br>/g;
    $auvdlinks =~ s/&middot;/&middot;<br>/g;
    $d3links =~ s/&middot;/<br>/g;
    $d3links =~ s/checked>\s/checked>/g;
    
    ##$d3links = "<form action=\"/3Dreplay/3D.asp\">" . $d3links;
    $d3links =~ s/<font size=\"-2/<br><br><center><font size=\"-2/;
    $d3links .= "</center>";
    ##$d3links .= "</form>";
    $has3Dlinks = 1 if $d3links =~ /3D\.asp/;
    
    #
    # Message if can't run Cosmo with GeoVRML
    #
    ##my $brows = $Request->ServerVariables(HTTP_USER_AGENT)->{Item};
    ##if ( ($brows =~ /mac/i) && $d3links_orig =~ /3D.asp/ ) {
    ##  $d3links = "<font color=\"brown\">3D Replay does not yet work on Macintoshes</font>";
    ##  $has3Dlinks = 0;  # Flag for display button at bottom of table
    ##}
    %>
    <tr>
    <td valign="top" class="dateCell"><font face="helvetica,Ariel" size="-1"><%=$DateString%></font></td>
    <td valign="top"><%=$ElinkString%></td>
    <td valign="top"><font face="helvetica,Ariel" size="-1"><%=$d3links%></font></td>
    <td valign="top"><font face="helvetica,Ariel" size="-2"><%=$adlinks%></font></td>
    <td valign="top"><font face="helvetica,Ariel" size="-2"><%=$sdlinks%></font></td>
    <td valign="top"><font face="helvetica,Ariel" size="-2"><%=$fdlinks%><%=$ddlinks%><%=$auvdlinks%></font></td>
    <td valign="top"><font face="helvetica,Ariel" size="-2"><%=$rdbdlinks%><%=$rdlinks%></font></td> 
    <td valign="top"><font face="helvetica,Ariel" size="-2"><%=$cdlinks%></font></td>
    <td valign="top"><font face="helvetica,Ariel" size="-2"><%=$ndlinks%></font></td>
    </tr>
  <%} else { # List format %>
    <dt><font face="helvetica,Ariel">

    <%= $StatusString%>
    &middot;

    <%if ( GetFormValue('qShipname') eq 'othr' ) {
      $shnm = ( $RS->Fields('RovName')->Value);
      $rvnm = ($RS->Fields('Divenumber')->Value);
      my $diveName = $RS->Fields('RovName')->value . $RS->Fields('DiveNumber')->value;
      #print "<br>RovName= $shnm$rvnm<br>"; %>
      <a href="<%= $this_script%>?step=5&amp;RovName=<%= $RS->Fields('RovName')%>&amp;DiveNumber=<%= $RS->Fields('DiveNumber')%>&amp;edit=<%= $diveName%>" target="edit"><img src="view_post2.gif" border="0" WIDTH="114" HEIGHT="21"></a>
      <font face="helvetica,Ariel" size="-1">
      <!--<a href="<%= $this_script%>?step=display&RovName=<%= $RS->Fields('RovName')%>&amp;DiveNumber=<%= $RS->Fields('DiveNumber')%>&amp;edit=<%= $RS->Fields('DiveNumber')%>" target="expd">Printable format</a>-->
      </font></dt></font><dd>
      <font face="helvetica,Ariel" size="-1"><%= $expd_text%></font></dd>
   <%}
    else { 
      $shnm = ( GetFormValue('qShipname') ne 'othr');
      #print "<br>qShipname1= $shnm<br>";%>
      <a href="<%= $this_script%>?step=3&amp;ShipName=<%= $RSdive->Fields('ShipName')%>&amp;ShipSeqNum=<%= $RSdive->Fields('ShipSeqNum')%>&amp;edit=<%= $RSdive->Fields('ExpeditionID_FK')%>" target="edit"><img src="edit_post.gif" border="0" WIDTH="114" HEIGHT="21"></a>
      <font face="helvetica,Ariel" size="-1">
      <a href="<%= $this_script%>?step=display&ExpeditionID=<%= $RSdive->Fields('ExpeditionID_FK')%>" target="expd">Printable format</a>
      </font>
      </dt></font><dd>
      <font face="helvetica,Ariel" size="-1"><%= $expd_text%></font></dd>
 <% }
}
  #$RSdive->MoveNext;  ##### Removed / changed all linking to RS instead of RSdive.
  $RS->MoveNext;
} ## End of While loop
if ( $_[0] eq 'data_table' ) { %>
  </table>
  <br>
  <%
  if ( $has3Dlinks == 1 ) { %>
    <input type="submit" class="button blue" name="3D" value="Continue to Select Terrain for All Above Selected 3D Replays">
  <% } %>
  </form>
  <hr ALIGN="left" WIDTH="30%">


<b>Missing Links:</b> <p class="brownText">Annotations and samples may exist for dives of an 
expedition even though they do not appear in this table.  The links will be provided here after dive
numbers, start, and end times have been entered.  To
enter dive information click on the expedition date(s) link and complete the Postcruise report; Science team
members are encouraged to complete this report as soon as possible following the cruise as this information enhances
the quality of the data.  The final database entry is made by the logistics coordinator based on your input. 
Annotations and samples are entered into their respective databases by the
video lab staff and samples coordinator and will be found if you query those databases directly.  
If you see
that other data links may be missing please contact Mike McCann.</p>
<br>
<br>
<b>Data link information:</b><br>
0. The 3D Replay requires a browser plugin and special software.  See 
<a href="/3Dreplay/geoVRMLreqts.html">3D Replay requirements and FAQ</a> for details.
<br>
1. <a href="/itd/video/AnnotationRetrieval.htm">Video Annotation retrieval</a>
<br>
2. <a href="/samples/">MBARI Sample Collection information</a> and
<a href="/samplesDB/queries/default.asp">Query the Samples Database</a>
<br>
3. <a href="/Ventana/stills/about.html">Background on Frame Grabs</a>
<br>
4. Links named like &quot;Data for dive xxxx (xxx points)&quot; are returned from the Rovctd table and 
include quality flags.  This links will not appear if dive numbers, start, and end times
have not been entered. Please see <a href="/expd/UserDocs/userdocs.htm"> 
Information about ROVCTD data</a>.
<br>
5. <a href="/expd/UserDocs/camlog_format.txt">Camera log file format</a>
<br>
6. <a href="/expd/UserDocs/NavRoadmap.txt">Original roadmap to Navigation data</a>
and <a href="/itd/DataArchives/navigation/index.htm">Navigation Data Archive</a>
<br>
<br>
For more details on MBARI's data archive, please see
<a href="/itd/DataArchives/ArchiveInfo/dataLinks-ships.htm">
Links to files, documentation, and file maintainers</a>.


<%} else { %>
  </dl>
<% } 
my $endBM = new Benchmark;
my $diffBM = timediff($endBM, $startBM);
Win32::ASP::DebugPrint("\nstep2(): Time taken for query and data links = ", timestr($diffBM, 'all'), "\n");



if ($qRovName == 'mini') {
  %>
  <p><br><br><b>Total of <%= $count%> expeditions found with query:</b>
  <br><font face="helvetica,Ariel" size="-3"><%= $search_sql %> [<%= timestr($diffBM, 'all') %> ]</font></p>
    <br><%
    if ($count == 0) {
  %><h2><font color="#408080"><strong>**You are searching on MiniROV dives. If your query resulted in no matches,
    please <br>consider revising your <a href="https://mww.mbari.org/expd/log/postcruise.asp?search=advanced">EXPD search(es)</a> to include both MBARI and Non-MBARI vessels.<br>
    <br><br>****If you feel your search is correct and the MBARI Expedition DataBase (EXPD) does not <br>reflect recent MiniROV data and/or cruises, please contact:<br>&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;</font><font color="blue">---> Dale Graves (x-1913, grda@mbari.org)</font><font color="#408080"><br>to request that newer MiniROV records be added to the MBARI EXPD.</strong></font>
   </h2>
  <%}
} else {
    %>
  <p><br><br><b>Total of <%= $count%> expeditions found with query:</b>
  <br><font face="helvetica,Ariel" size="-3"><%= $search_sql %> [<%= timestr($diffBM, 'all') %> ]</font></p>
  <%  
}

#$RSdive->Close;
$RS->Close;
$Conn->Close;
$ConnVARS->Close;
$ConnSamples->Close;


} # End step4()


%>
  

<%
#--------------------------------------------------------------------
#

=head3 step5()

Present Non-MBARI Ship / MiniROV dive info for editing.  Currently all editing ofNon-MBARI 
dives is not allowed.  Leaving "hooks" in for future discussion.

Author: Karen Salamy

Date Created: 11/18/98

=cut

sub step5 {

#
# Open database to get specified expedition
#
open_database($dsn);

#
# Open VARS database to get data for dataLinks()
#
open_vars_database();

#
# Open Samples database to get data for dataLinks()
#
open_samples_database();    # Creates $ConnSamples object as a global variable

# Do an explicit select (not select *) to avoid problems...
  $sql = "SELECT " . join( ", ", @{$Application->{'DiveFields'}} ) . " FROM Dive\n";
  $sql .= "WHERE RovName = '" . GetFormValue('RovName') . "' ";
  $sql .= "AND DiveNumber = " . GetFormValue('DiveNumber');
  $sql .= " ORDER BY DiveNumber";

#print "<br> $sql<br>";

Win32::ASP::DebugPrint("\nstep5():\nExecuting SQL: $sql");
$RSdive = $Conn->Execute($sql);
if(!$RSdive) {
   $Response->Write("<BR>Empty PC Return Set2: RSdive. for <br>$sql ");
}

%>

<table width="830" border="3">
<tr><td>
    <h3><font color="#408080"><strong><br><br>&nbsp;<u>NOTE:</u> <br><br>&nbsp;Postcruise editing of MiniROV Dives on Non-MBARI Vessels is not currently
    supported within the<br>&nbsp;MBARI Expedition Database.<br><br></strong></font>
   </h3>
<table width="830">
  <tr>
    <td colspan="2" bgcolor="silver" align="center">
    <font face="helvetica,Ariel"><b>Dive Information</b></font></td>
  </tr>
  <tr><td colspan="2"><font face="helvetica,Ariel" color="brown" size="-1">
  <br>Dives are defined as part of the postcruise report process. Click on 
  the appropriate button.  Dive start and end times must be entered to be
  able to get ROVCTDO data from the database, have the frame grabs organized 
  by dive number, and to view 3D Replays.
  </font></td></tr>
<%
while ( !$RSdive->EOF ) {     # List existing dives in the database
  
  my $editlink = "";
  for (my $i = 0; $i < 20; $i++) {
    $editlink .= "&nbsp;";
  }
  
  # SEND THIS OVER TO DIVE.ASP - CANNOT EDIT Non-MBARI Dives
  $editlink .= "<a href=dive.asp?RovName=" . $RSdive->Fields('RovName')->Value . "\&";
  $editlink .= "DiveNumber=" . $RSdive->Fields('DiveNumber')->Value . "&edit=yes target=\"dive_edit\">"; 
  
  foreach $f ( @{$Application->{'DiveFields'}} ) {
    if ( $RSdive->Fields($f)->Value ) {
      $Session->{$f .'_orig'} = $RSdive->Fields($f)->Value ;    # Remember original
    }
    else {
      $Session->{$f .'_orig'} = ''              # Clear session space;
    }
    %>
    <tr>
    <%if ( $f eq 'BriefAccomplishments' ) {   # Geeze, this is a mess! %>
      <td valign="top"><font face="helvetica,Ariel" size="-1"><b>
      Dive Accomplishments and Comments:</b></font></td>
    <%} else { %>
      <td valign="top"><font face="helvetica,Ariel" size="-1"><b><%= $f%></b></font></td>
    <% }
    if ( $f eq 'RovName') { %>
      <td valign="top"><font face="helvetica,Ariel" size="-1">
      <%= $RSdive->Fields($f)->Value%></font></td>
    <%
    }
    elsif ( $f eq 'DiveStartDtg' ) { %>
      <td valign="bottom"><font face="helvetica,Ariel" size="-1">
      <% if ( $RSdive->Fields($f)->Value ) {%>
            <%= $RSdive->Fields($f)->Value%> UTC<%
         }
         else { %>
            Edit dive to enter a dive start time
       <%}%>
      <%= $editlink%>
      </font></td><%
    }
    elsif ( $f eq 'DiveEndDtg' ) {
      if ( $RSdive->Fields($f)->Value ) {%>
        <td valign="top"><font face="helvetica,Ariel" size="-1">
        <%= $RSdive->Fields($f)->Value%> UTC</font></td><%
      }
      else {%>
        <td><font face="helvetica,Ariel" color="brown" size="-2">
        Edit dive to enter a dive end time</font></td>
      <%
      }
    }
    else {%>
      <td valign="top"><font face="helvetica,Ariel" size="-1">
      <%= $RSdive->Fields($f)->Value%></font></td>
    <%}%>
    
    </tr><%
  } # End foreach( 
  my $dlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value,
               RovName_FK => $RSdive->Fields('RovName')->Value,
               DataType => 'rovctddb' );
  if ( $dlinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>ROVCTD Data:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $dlinks %></font></td></tr>
  <%}
  
  $tslinks = tapeSummaryLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value,
              RovName_FK => $RSdive->Fields('RovName')->Value );
  if ( $tslinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Tapes:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $tslinks %></font></td></tr>
  <%}
  
  $dlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value,
              RovName_FK => $RSdive->Fields('RovName')->Value,
               DataType => 'annotations');
  if ( $dlinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Annotations:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $dlinks %></font></td></tr>
  <%}
  $dlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value,
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => 'samplesdb');
  if ( $dlinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Samples:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $dlinks %></font></td></tr>
  <%}
  my $flinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => 'frameGrabs');
  if ( $flinks ) { 
    #print "<br>HAVE Step5 flinks.<br>";
    %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Frame Grabs:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $flinks %></font></td></tr>
  <%}
  my $clinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => 'digitalImages');
  if ( $clinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Coolpix images:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $clinks %></font></td></tr>
  <%}
  
  my $camlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => 'camlogdb');
  if ( $camlinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Camlog:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $camlinks %></font></td></tr>
  <%}
  
  my $navlinks = '';
  my $rawNavlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => 'RawNavdb');
  $navlinks .= $rawNavlinks if $rawNavlinks;
  my $cleanNavlinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => 'CleanNavdb');
  $navlinks .= "<b>&middot;</b>" . $cleanNavlinks if $cleanNavlinks;
  if ( $navlinks ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Navigation:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $navlinks %></font></td></tr>
  <%}
  
  my $d3links = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => '3DReplay');
  if ( $d3links ) { %>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>3D Replay:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $d3links %></font></td></tr>
  <%}
    
  my $GElinks = dataLinks( DiveNumber_FK => $RSdive->Fields('DiveNumber')->Value, 
              RovName_FK => $RSdive->Fields('RovName')->Value,
              DataType => 'GoogleEarth');
  if ( 1 ) { 
    #print "<br>HAVE Step5 GElinks.<br>";%>
    <tr><td><font face="helvetica,Ariel" size="-1"><b>Google Earth:</b></font></td>
    <td><font face="helvetica,Ariel" size="-1"><%= $GElinks %></font></td></tr>
    
  <%}
  %>
  <tr><td>&nbsp;</td><td></td></tr>
  <%
  $RSdive->MoveNext;
  %><%  
} # End while ( !$RSdive->EOF

if ( $RSdive->EOF == $RSdive->BOF ) { %>
<tr><td>
<font face="helvetica,Ariel" size="-1">No dives entered for this expedition</font>
</td></tr>
<% }
# END OF DIVE INFORMATION TABLE
%>
</table>  
<br>
<hr width="75%" align="center">
<!--#include file="postcruise_ftr.inc"-->
<input type="hidden" name="step" value="4">
<!--<input type="hidden" name="ExpeditionID_FK" value="<%= $RS->Fields('ExpeditionID')->Value%>"> -->
<input type="hidden" name="ExpeditionID" value="<%= $RS->Fields('ExpeditionID')->Value%>">
<input type="hidden" name="ShipName" value="<%= $RS->Fields('ShipName')->Value%>">
<input type="hidden" name="ShipSeqNum" value="<%= $RS->Fields('ShipSeqNum')->Value%>">
<input type="hidden" name="edit" value="<%= $RS->Fields('ExpeditionID')->Value%>">
<input type="hidden" name="RovName" value="<%= $RSdive->Fields('RovName')->Value%>">
<input type="hidden" name="DeviceID" value="<%= $RSdive->Fields('DeviceID')->Value%>">
<input type="hidden" name="DiveNumber" value="<%= $RSdive->Fields('DiveNumber')->Value%>">

<% if ( $Session->{'RegisteredUser'} ) { %> 
  <strong><font face="helvetica,Ariel" size="-1" color="#408080">Send postcruise updates email to 
  <input type="text" name="post_email" value="postcruise" size="9" maxlength="15">
  @mbari.org.</font></strong></p>
  <br>
  <center>
  <input type="submit" class="button green" name="finish" value="Update Database with Above Changes and Continue with Dive Information."> 
  <br><br>
  <input type="submit" class="button blue" name="finish" value="Done with this Expedition. Update Changes Made Above.">
  <br><br>
  <input type="reset" class="button gray" name="reset" value="Reset Form Fields"> 
  </center>
<%}
else {
%>
  <font face="helvetica,Ariel" size="-1"><b>Your MBARI email address:</b>
  <input type="text" name="email_addr" value size="9">@mbari.org</font>
  <br><br>
  <center>
  <input type="submit" class="button blue" name="finish" value="Continue with this Expedition and Enter New Dive Information">
  <br>
  <font face="helvetica,Ariel" size="-1">
  Email is sent only if you have updated the above Expedition Information.
  <br>
  You will get a blind copy of the mail that is sent to <%=$Application->{'logistics_email'}%>.
  </font>
  <br><br>
  <input type="submit" class="button blue" name="finish" value="Done with this Expedition. Submit to Logistics Coordinator">
  <br><br>
  <input type="reset" class="button gray" name="reset" value="Reset Form Fields">
  </center>
<%}
%>
</form>
</td></tr>
</table>
<%
$RSdive->Close;
$RS->Close;
$Conn->Close;
$ConnVARS->Close;
$ConnSamples->Close;
} # End step5()
%>

<%
#--------------------------------------------------------------------
#

=head3 display_expd()

Same as step3(), but don't present form fields.  Save text in session
variable for postcruise email message.

Author: Mike McCann

Date Created: 12/4/98

=cut

sub display_expd {

use Text::Wrap;

$Text::Wrap::columns = 65;

open_database($dsn);

if ( GetFormValue('qShipName') ne 'othr' ) {
  $sql = "SELECT " . join(",", @{$Application->{'ExpeditionFields'}}) . " FROM Expedition\n";
  $sql .= "WHERE ";
    if ( GetFormValue('ShipName') && GetFormValue('ShipSeqNum') ) {
      $sql .= "ShipName = '" . GetFormValue('ShipName') . "' ";
      $sql .= "AND ShipSeqNum = " . GetFormValue('ShipSeqNum');
    }
    elsif ( GetFormValue('ExpeditionID') ) {
      $sql .= "ExpeditionID = " . GetFormValue('ExpeditionID');
    }
}

#print "<br>Display_EXPD: $sql<br>";
Win32::ASP::DebugPrint("\nstep3():\nExecuting SQL: $sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
  $RS->Close;
  die "<BR>Empty Return Set: RS1. for <br>$sql ";
}

%>

<table width="775">
<tr>
	<td colspan="2" bgcolor="silver" align="center">
	<font face="helvetica,Ariel"><b>Expedition Information</b></font></td>
</tr>
<%
#
# Loop through all the fields and offer for update the ones we change in a Post Cruise
# Once a value is read vie $RS->Fields($f)->Value from a memo field in the database
# you can't read it again!  That is why $ExpeditionValue{$f} is used.
#
$mail_msg = "                    Expedition Information\n";
$mail_msg .= "                    ----------------------\n";
foreach $f ( @{$Application->{'ExpeditionFields'}} ) {

	if ($f eq 'StartDtg') {
	   		$mail_msg .= "\nPost Cruise report:\n-------------------\n";
	}

	my $val = $RS->Fields($f)->Value;
	#
	# Replace any î characters with º
	#
	$val =~ s/î/º/g;
	$val = wrap("", "", $val);
	next if $f eq 'TimeStamp';   # Screws up $mail_msg somehow, just skip it.
	$GMTstring = "";
	$GMTstring = " GMT" if ( $f =~ /dtg/i );
%>
<tr>
	<td valign="top"><font face="helvetica,Ariel" size="-1"><b><%= $f%></b></font></td>
	<td valign="top"><pre><%= $val%><%=$GMTstring%></pre></td>
</tr>
<%
	my $val = wrap("  ", "  ", $val);
		$mail_msg .= "\n" . $f . ":\n" . $val;
		$mail_msg .= " UTC" if $f =~ /Dtg$/i;
		$mail_msg .= "\n";
		if ($f eq 'ScheduledStartDtg') {
	   		$mail_msg .= "  $Session->{'StartDtgLocal'} Local time\n";
	   	}
	   	if ($f eq 'ScheduledEndDtg') {
	   		$mail_msg .= "  $Session->{'EndDtgLocal'} Local time\n";
	   	}
	   	if ($f eq 'StartDtg') {
	   		$mail_msg .= "  $Session->{'StartDtgLocal'} Local time\n";
	   	}
	   	if ($f eq 'EndDtg') {
	   		$mail_msg .= "  $Session->{'EndDtgLocal'} Local time\n";
	   	}
}
%>

<%
#
# Now check for and display dive information
#
if ( GetFormValue('qShipName') eq 'othr' && GetFormValue('qRovName') eq 'mini' ) {
  $sql = "SELECT * FROM Dive\n";
  $sql .= "WHERE ";
  $sql .= "RovName = '" . (GetFormValue('qRovName')) . "' AND DeviceID = " . ( $RS->Fields('DeviceID')->Value) . " AND \n";
  $sql .= "ExpeditionID_FK = " . $RS->Fields('ExpeditionID_FK')->Value;
  $sql .= " ORDER BY DiveStartDtg";
}
elsif ( GetFormValue('qShipName') ne 'othr' && GetFormValue('qRovName') ne 'mini' ) {
  $sql = "SELECT * FROM Dive\n";
  $sql .= "WHERE ";
  $sql .= "ExpeditionID_FK = " . $RS->Fields('ExpeditionID')->Value;
  $sql .= " ORDER BY DiveStartDtg";
}

#print "<br>Dive Info Display _Expd: $sql<br>";

Win32::ASP::DebugPrint("\ndisplay_expd():\nExecuting SQL: $sql");
$RSdive = $Conn->Execute($sql);
if(!$RSdive) {
 	 $Response->Write("<BR>Empty Return Set9: RSdive. for <br>$sql ");
}

%>

<tr>
	<td colspan="2" bgcolor="silver" align="center">
	<font face="helvetica,Ariel"><b>Dive Information</b></font></td>
</tr>
<%
$mail_msg .= "                       Dive Information\n";
$mail_msg .= "                       ----------------\n";

##Win32::ASP::DebugPrint("\nfoobar\n$mail_msg\n barrrrr");

# Confusing and not really needed.  Don't include with the postcruise report.
##$mail_msg .= "\nNo dive(s) entered.\n" if ($RSdive->EOF == $RSdive->BOF);
while ( !$RSdive->EOF ) {
	foreach $f ( @{$Application->{'DiveFields'}} ) {
		%><tr>
		<td valign="top"><font face="helvetica,Ariel" size="-1"><b><%= $f%></b></font></td>
		<td valign="top"><font face="helvetica,Ariel" size="-1">
		<%= $RSdive->Fields($f)->Value%></font>
		<%
		if ( $f eq 'DiveStartDtg' ) {
			my $link = "dive.asp?DiveNumber=" . $RSdive->Fields(DiveNumber)->Value;
			$link .= "&RovName=" . $RSdive->Fields(RovName)->Value;
			$Response->Write("<font face=\"helvetica,Ariel\" size=-2>
			<a href=$link>Edit dive to set start time</a></font>") 
				unless $RSdive->Fields($f)->Value
		}
		%>
		</td>
		
		</tr>
		
		<%
		my $val = wrap("  ", "  ", $RSdive->Fields($f)->Value);
		$mail_msg .= "\n" . $f . ":\n" . $val;
		$mail_msg .= " UTC" if $f =~ /Dtg$/i && $RSdive->Fields($f)->Value;
		$mail_msg .= "\n";
	} # End foreach(
	$mail_msg .= "\n";
	$RSdive->MoveNext;
	%>
	<tr><td colspan="2">
	<hr width="50%" align="center">
	</td></tr>
	<%
	Win32::ASP::DebugPrint("\nfoobar\n$mail_msg\n barrrrr");
} # End While(
%></table><%

#
# Save mail text for use by finished_email()
#
$Session->{'mail_msg'} = $mail_msg;


$RSdive->Close;
$RS->Close;
$Conn->Close;

}	# End display_expd()
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

=head3 SelectEpoch()

Select epoch seconds from database field specified.

Author: Mike McCann

Date Created: 10/28/98

=cut

sub SelectEpoch {

	if ($RS->Fields('ExpeditionID')->value eq 0) {
		my $time_field = $_[0];

		$sql = "SELECT DateDiff(\"ss\", '01/01/70', DiveStartDtg) AS Epoch ";
		$sql .= "FROM Dive\nWHERE ";
		$sql .= "ExpeditionID_FK = " . GetFormValue('edit') . "\n";
	}
	else {
	my $time_field = $_[0];

		##$sql = "SELECT DateDiff('s', '01/01/70', $time_field) AS Epoch ";		# -The way Access wants it
		$sql = "SELECT DateDiff(\"ss\", '01/01/70', $time_field) AS Epoch ";
		$sql .= "FROM Expedition\nWHERE ";
		$sql .= "ExpeditionID = " . GetFormValue('edit') . "\n";
	}
	Win32::ASP::DebugPrint("\nSelectEpoch():\nExecuting SQL: \n$sql");
	$RSdate = $Conn->Execute($sql);
	if(!$RSdate) {
		die "<BR>Empty Return Set10: RSdate. for <br>$sql ";
	}

	my $epoch = $RSdate->Fields('Epoch')->Value;
	$RSdate->Close;
	
	return($epoch);
}
%>

<%
#--------------------------------------------------------------------
#

=head3 process_form_entries()

Verify entries before submitting them, does the work of creating
all the DB session variables.  Session variables are not assigned 
unless it is meant to be there, i.e. the field requires updating
because it was entered on the form and is different from what was
in the database.

Author: Mike McCann

Date Created: 10/20/98

=cut

sub process_form_entries {

	#
	# Construct the UTC start and End times for the database update
	#
	my $hhmm = sprintf('%04s', GetFormValue('StartDtgTime'));
	@sl = (	0, substr($hhmm,2,2), substr($hhmm,0,2),
			GetFormValue('StartDtgDay'), GetFormValue('StartDtgMonth')-1, GetFormValue('StartDtgYear')-1900 );
	@el = (	0, substr(GetFormValue('EndDtgTime'),2,2), substr(GetFormValue('EndDtgTime'),0,2),
			GetFormValue('EndDtgDay'), GetFormValue('EndDtgMonth')-1, GetFormValue('EndDtgYear')-1900 );
			
	$Session->{'StartDtg'} = toGMT(@sl);
	$Session->{'EndDtg'} = toGMT(@el);
	$Session->{'StartEpoch'} = timelocal(@sl);		# For original check
	$Session->{'EndEpoch'} = timelocal(@el);		# For original check
	
	
	$Session->{'StartDtgLocal'} = GetFormValue('StartDtgYear') . '-' .
				GetFormValue('StartDtgMonth') . '-' .
				GetFormValue('StartDtgDay') . ' ' . 
				GetFormValue('StartDtgTime');
				
	$Session->{'EndDtgLocal'} = GetFormValue('EndDtgYear') . '-' .
				GetFormValue('EndDtgMonth') . '-' .
				GetFormValue('EndDtgDay') . ' ' . 
				GetFormValue('EndDtgTime');
	
	#
	# Save all input variables as session variables
	#
	$ActualRegions = '';
	foreach $s (Win32::OLE::in ($Request->QueryString('ActualRegions'))) {
		$s =~ s/,/_/;		# Replace any commas
		$ActualRegions .= $s . ',';
	}
	$ActualRegions =~ s/,$//;		# Remove last comma
	$Session->{'RegionDesc'} = $ActualRegions;
	
	#
	# For values that may have been entered, assign to Session variables, if they are
	# different than their original, special check on Epoch seconds for ...Dtg
	#
	foreach $f ( @{$Application->{'PostUpdateExpeditionFields'}} ) {
		if ( $f eq 'StartDtg' ) {
			if ( $Session->{'StartEpoch_orig'} == $Session->{'StartEpoch'} ) {
				#
				# If status changed from plnd to cmpl then use scheduled times, as people aren't
				# really entering the actual times anyway.
				# Do as well if status goes to abrt.
				#
				if ( $Session->{'StatCode_orig'} eq 'plnd' && $Session->{'StatCode'} eq 'cmpl' ) {
				}
				elsif ( $Session->{'StatCode_orig'} eq 'plnd' && $Session->{'StatCode'} eq 'abrt' ) {
				}
				else {
					$Session->{'StartDtg'} = '';
				}
			}
		}
		elsif ( $f eq 'EndDtg' ) {
			if ( $Session->{'EndEpoch_orig'} == $Session->{'EndEpoch'} ) {
				#
				# If status changed from plnd to cmpl then use scheduled times, as people aren't
				# really entering the actual times anyway
				#
				if ( $Session->{'StatCode_orig'} eq 'plnd' && $Session->{'StatCode'} eq 'cmpl' ) {
				}
				else {
					$Session->{'EndDtg'} = '';
				}
			}
		}
		elsif ( $f eq 'RegionDesc' ) {
			if ( $Session->{'RegionDesc_orig'} eq $Session->{'RegionDesc'} ) {
				$Session->{'RegionDesc'} = '';
			}
		}
		elsif ( GetFormValue($f) ) {
			if ( GetFormValue($f) ne $Session->{$f . '_orig'} ) {
				$Session->{$f} = GetFormValue($f);		
			}
			else {
				$Session->{$f} = '';		# Clear this out to forget fields between entries in same session
			}
		}
		else {
			$Session->{$f} = '';
		}
		# Debugging lines
		Win32::ASP::DebugPrint(	"\n$f: " . $Session->{$f} .
								"\n$f" . '_orig: ' . $Session->{$f . '_orig'} );
		##$Response->Write("\n<br><br><b>$f:</b> " . $Session->{$f});
		##$Response->Write("\n<br><b>$f" . '_orig:</b> ' . $Session->{$f . '_orig'});
		
	}	# End foreach PostUpdateExpeditionFields
	
	#
	# Print things out for checking
	#
	%>
	<table><%
	my $i = 0;
	foreach $f ( @{$Application->{'PostUpdateExpeditionFields'}} ) {
		next if $f eq 'ExpeditionID';
		next unless $Session->{$f};
				$i++;
		%>
		<tr>
		<%
		if ($i == 1) {
			$Response->Write("<td valign=\"top\" bgcolor=\"silver\"><font face=\"helvetica,Ariel\">");
			if ( $Session->{'RegisteredUser'} ) {
				$Response->Write("<b>Expedition fields updated:</b> ");
			}
			else {
				$Response->Write("<b>Expedition fields to be updated:</b> ");
			}
		}
		else {
			$Response->Write("<td valign=\"top\"><font face=\"helvetica,Ariel\">");
			$Response->Write("&nbsp;");
		}
		%>
		</td>
			<td valign="top"><font face="helvetica,Ariel" size="-1"><b><%= $f%></b></font></td>
			<td valign="top"><font face="helvetica,Ariel" size="-1"><%= $Session->{$f}%></font></td></tr>
		<%
	}
	%></table>
	
<%

}	# End process_form_entries()
%>


<%
#--------------------------------------------------------------------
#

=head3 update_expedition()

Process the form field entries and insert into the database
Here is where we need to map the names we have on the HTML 
forms for all the fields to the names of the fields that are
actaully in the database table.

Author: Mike McCann

Date Created: 10/28/98

=cut

sub update_expedition {
#
# Loop through all form names and construct SQL update string, exit without
# complaining if there are no fields to update
#
$sql = "UPDATE Expedition\nSET ";
my $nfields = 0;
foreach $f ( @{$Application->{'PostUpdateExpeditionFields'}} ) {
	next if $f eq 'ExpeditionID';
	##$Response->Write("<br>$f");
	next unless $Session->{$f};
	if ($f eq 'DiveNumber' || $f eq 'ExpeditionID_FK' ) {
		$sql .= "\n$f = " . FixString( $Session->{$f}, 'num') . ",";
	}
	else {
		$sql .= "\n$f = " . FixString( $Session->{$f}, 'text') . ",";
		$nfields++;
	}
}
return unless $nfields;

#
# Open the database, this time to update fields
#
open_database($dsn);		# Creates $Conn object as a global variable

$sql =~ s/,$//;
$sql .= "\nWHERE ExpeditionID = " . GetFormValue('ExpeditionID');
Win32::ASP::DebugPrint("\nupdate_expedition():\nSQL = $sql ");
$RS = $Conn->Execute($sql);

$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("Database Load Error(s): ");
	foreach $error (keys %$Errors)
        {
      		$Response->Write($error->{Description});
   	}
}

$RS->Close;
$Conn->Close;

}	# End update_expedition()
%>

<%
#--------------------------------------------------------------------
#

=head3 email()

Process the form field entries and send email to the Logistics
coordinator.

Author: Mike McCann

Date Created: 11/5/98

=cut

sub email {

	#
	# Use libnet package to send
	# mail message to logistics coordinator.  Composes the mail message here.
	# Do not use $mail_msg!
	#
  use Net::SMTP;
	use Text::Wrap;


    $smtp = Net::SMTP->new('mail.shore.mbari.org'); 	# connect to an SMTP server
	  $smtp->mail( GetFormValue('email_addr') . "\@mbari.org" );     	# use the sender's address (mccann?) here
    $smtp->to( $Application->{'logistics_email'}, GetFormValue('email_addr') . "\@mbari.org" );        # recipient's address
    $smtp->data();                      # Start the mail

    # Send the header.
    #
    $smtp->datasend("To: $Session->{'logistics_email'}\n");
    $smtp->datasend("From: " . GetFormValue('email_addr') . "\@mbari.org \n");
	  $smtp->datasend("Subject: Request: Postcruise $Session->{'StartDtg'} (mail from web form $this_script)\n");
    $smtp->datasend("\n");

    # Send the body.
    #
    $smtp->datasend("Postcruise database load request from " . GetFormValue('email_addr') . "\n");
	  $smtp->datasend("============================================\n\n");
	  $smtp->datasend("Fields to be modified:\n");
	  $smtp->datasend("----------------------\n\n");
	
	$Text::Wrap::columns = 65;
	my $fieldCount = 0;
	foreach $f ( @{$Application->{'PostUpdateExpeditionFields'}} ) {
		$str1 = $Session->{$f};
		next unless $str1;
		$fieldCount++;
		##$str2 = wrap("    ", "    ", $str1);
	   	$smtp->datasend("$f:\n$str1\n") if $Session->{$f};
	   	if ($f eq 'StartDtg') {
	   		$smtp->datasend("    $Session->{'StartDtgLocal'} Local time\n");
	   	}
	   	if ($f eq 'EndDtg') {
	   		$smtp->datasend("    $Session->{'EndDtgLocal'} Local time\n");
	   	}
	}
	if ( ! $fieldCount ) {
		$Response->Write("<H3>fieldCount = $fieldCount.  No mail message to send.</H3>");
		return;
	}
	
	$smtp->datasend("\nClick below to edit this postcruise. Copy above fields into the web form:");
	$baseurl = $Application->{'BaseUrl'};
	$hyperlink = "$baseurl/$this_script?step=3&edit=" . GetFormValue('ExpeditionID') .
		"&ShipName=" . GetFormValue('ShipName') . "&ShipSeqNum=" . GetFormValue('ShipSeqNum');

	  $smtp->datasend("\n" . $hyperlink . "\n\n");
	
    $smtp->dataend();                   # Finish sending the mail
    $smtp->quit;                        # Close the SMTP connection
	%>

<h3>Email has been sent to <%= $Session->{'logistics_email'}%> and cc'ed to <%=GetFormValue('email_addr')%></h3>
Subscribe to the 'postcruise' alias to get notification of your cruise
being entered into the database.

<%
}	# End email()
%>


<%
#--------------------------------------------------------------------
#

=head3 finished_email()

Send message to the postcruise alias (or whatever was entered
on the step3() form.

Author: Mike McCann

Date Created: 12/7/98

=cut

sub finished_email {

	return unless $mail_msg;		# Do nothing unless we have a message.
	
	#
	# Use libnet package to send
	# mail message to address indicated in the form that the Logistcs Coordinator sees, usually 'postcruise'.
	#
	use Net::SMTP;
	use Text::Wrap;
	
	$smtp = Net::SMTP->new('mail.shore.mbari.org'); 		# connect to an SMTP server
	$smtp->mail( $Application->{'logistics_email'} );   # use the sender's address here
	$smtp->to( GetFormValue('post_email') . "\@mbari.org" );		# recipient's address
	$smtp->data();									# Start the mail

	# Send the header.
	#
	my $startTime = ($Session->{'StartDtg'} ne '') ? $Session->{'StartDtg'} : $Session->{'StartDtg_orig'};
	my $chiefScientist = ($Session->{'ExpdChiefScientist'} ne '') ? $Session->{'ExpdChiefScientist'} : $Session->{'ExpdChiefScientist_orig'};
	$smtp->datasend("To: " . GetFormValue('post_email') . "\@mbari.org" . "\n");
	$smtp->datasend("From: " . $Application->{'logistics_email'} . " (Web form $this_script)\n");
	$smtp->datasend("Subject: " . $chiefScientist . ": " . $startTime . "\n");
	
	$smtp->datasend("\n");
	
	# Send the body.
	#
	$smtp->datasend("Postcruise report\n");
	$smtp->datasend("=================\n\n");
	
	$smtp->datasend( $Session->{'mail_msg'} );		# Session var constructed in display_expd()
	$smtp->datasend("\nClick below to edit this expedition:");
	$hyperlink = "$Application->{'BaseURL'}/$this_script?step=3&ShipName=" .
		GetFormValue('ShipName') . "&ShipSeqNum=" . GetFormValue('ShipSeqNum') . 
		"&edit=" . GetFormValue('edit');
	$smtp->datasend("\n" . $hyperlink . "\n\n");
	$smtp->dataend();                   # Finish sending the mail
	$smtp->quit;                        # Close the SMTP connection
	%>
<h3>The following email has been sent to <%= GetFormValue('post_email')%>@mbari.org</h3>
	<pre>
<%= $mail_msg%>
	</pre>
<%
}	# End finished_email()
%>



