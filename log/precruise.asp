<%@ LANGUAGE = PerlScript %>

<!--#include file="expd_functions.inc"-->
<!--#include file="precruise_hdr.inc"-->

<%
=head1 NAME

precruise.asp - Application for submitting precruises

=head1 SYNOPSIS

    http://expd.mbari.org/expd/log/precruise-telepresence.asp

=head1 DESCRIPTION

Precruise processing Active Server Page.
Collect info for planed expedition and insert new record to
the Expedition table.

Mike McCann MBARI

November 1998
December 2001

Edited by: Karen Salamy, MBARI

February 2016
Edited: Precruise email updates: 4/30/2025; Karen Salamy

=head1 FUNCTIONS

=cut


# Warning: Do NOT use CGI, it freezes IIS when a new object is created.
use Win32::ASP;
use OLE;
use Time::Local;
use POSIX;

$this_script = 'precruise.asp';

init_AppVars();

$Session->{'logistics_email'} = $Application->{'logistics_email'};
##$Session->{'logistics_email'} = 'mccann@mbari.org';     # For debugging
##$Session->{'logistics_email'} = 'salamy@mbari.org';     # For debugging

$dsn = $Application->{'DSN'};

$Response->Write($Application->{'Warning'});
$Response->Write($Application->{'MoreInfo'});
##$Session->{'DBadministrator'} = 1;          # For debugging.  Make sure to comment out when done!
##if ( $Session->{'DBadministrator'} ) {
##  $Response->Write("Admin = " . $Session->{'DBadministrator'});   # For debugging
##}
##$Response->Write("logistics_email = " . $Session->{'logistics_email'});

if ( GetFormValue('step') eq '1' || ! GetFormValue('step') ) {
   %><h2>Precruise entry: Step 1</h2>
</td></tr></table>
    <ul style="list-style: none; padding-left: 10px; width: 1200px; color:#408080; line-height: 0.85; font-weight: bold; font-size: 20px" >
    <!-- CLEAR UP ANY CONFUSION ON WHERE TO BEGIN FORM ENTRY  -->
    <br><br><br><b>&#8667; SELECT A SHIP from the dropdown menu <font color="red">below</font> to begin your precruise form entry.</b>
    </ul>
   <%
  step1(); %>
   <p>
    <ul style="list-style: none; padding-left: 10px; width: 1200px; color:#408080; line-height: 0.5; font-weight: bold; font-size: 16px" >
<table>
  <tr>
    <td style="margin: 10px; padding: 50px;"></td>
  </tr>
</table>
    <b>NOTE: Important Cruise Planning Links</b>
    </ul>

   <ul style="width: 1200px; color:#408080; line-height: 0.8; font-weight: bold; font-size: 16px" >

    <br>
    <li><a target="_blank" href="https://mww.mbari.org/dmo/Pages/Covid19.html">MBARI <font color="red">COVID-19</font> shipboard plans</a> and <a target="_blank" href="https://mww.mbari.org/dmo/Pages/ForSciencePlanning.html"> precruise questionnaire link and information.</a></li><br><br>
    <!--<li><a target="_blank" href="http://mww.mbari.org/dmo/">MBARI <font color="red">COVID-19</font> Ship and Policy</a> link.</li><br><br> -->
     <!--<li>Use this form for entering a new MBARI cruise plan.</li><br>-->
    <li>Timeline for submitting cruise plans:</li>
      <ul style="line-height: 1.0">
         <li><a target="_blank" href="https://mww.mbari.org/dmo/Pages/DPPlanning.html">R/V <i>David Packard</i> plans</a> are due three (3) weeks in advance.</li>
         <li><a target="_blank" href="https://mww.mbari.org/dmo/Pages/RCPlanning.html">R/V <i>Rachel Carson</i> plans</a> are due two (2) weeks in advance.</li>
         <li><a target="_blank" href="https://mww.mbari.org/dmo/Pages/Paragon.html">R/V <i>Paragon</i> plans</a> are due prior to departure.</li>
         <li><a target="_blank" href="http://www.mbari.org/technology/emerging-current-tools/wave-glider/">Wave Glider plans</a> are due three (3) weeks in advance.</li></ul>
         <br><br>
        <li>To modify an existing cruise plan email requested changes to the <a href="mailto:<%= $Session->{'logistics_email'}%>?subject=Request: Modify an Existing Precruise Plan">logistics coordinator</a>.</li><br><br>

    <li><a target="_blank" href="https://mww.mbari.org/dmo/Pages/CruisePlanning.html">DMO cruise planning</a> link.</li>
    <br><br>
    <li><a target="_blank" href="https://mww.mbari.org/cruises/">MBARI expedition database</a> link.</li>
    <br>

  </ul>
  </p>
<!-- MOVE THE ENTIRE FORM LEFT - NOT SURE WHY THIS WAS ORIGINALLY DONE-->
</td></tr></table>
   <%
}
elsif ( GetFormValue('step') eq 'getWaypoints') {   # Needs to be before step 2
  %><h2 align="center">Visiting waypoint management system to collect & organize waypoints</h2>
  </td></tr></table><%
  getWaypoints();
}
elsif ( GetFormValue('step') eq '2') {
  %><h2 align="center">Precruise entry: Step 2</h2>
  </td></tr></table><%
  step2();
}
elsif ( GetFormValue('step') eq '3') {
  %><h2 align="center">Precruise entry: Checking Values</h2>
  </td></tr></table><%
  step3();
}
elsif ( GetFormValue('step') eq 'orderwpts') {
  %><h2 align="center">Precruise entry: Order Waypoints</h2>
  </td></tr></table><%
  step2('orderwpts');
}
elsif ( GetFormValue('step') eq '4') {  
  %><h2 align="center">Precruise entry: Finishing Up</h2>
  </td></tr></table><font face="Arial,Helvetica" >
  It is the responsibility of the Chief Scientist to notify Operations
  of any special arrangements in advance of your cruise, such as computer
  systems needs, diving activities, participant's dietary restrictions or
  pertinent medical information. Please submit an e-mail to the Logistics
  <a href="mailto:<%= $Session->{'logistics_email'}%>">Support Specialist</a>
  describing any special requirements at this time.</font><%

  ### If this is being submitted by the logistics coordinator, send a precruise_mail_out
  # Clicking the "Load Data Base"
  if ( GetFormValue('finish') =~ /Load/ ) {
    my $ret = load();
    precruise_mail_out() if $ret == 0;
  }

  ## Otherwise, this is an initial email from a user being sent to the logistics coordinator
  # Clicking the "Email Form to Logistics Coordinator"
  else {
    email();
  }
  close_session();    # Forget all that I know
  %>
  <p><a href="<%= $this_script%>">Enter another precruise</a></p><%
}

## This is a request to the logistics coordinator to update an existing precruise - not used often
elsif ( GetFormValue('step') eq 'revise') {
  %><h2 align="center">Revise Precruise <%=GetFormValue('ShipName')%><%=GetFormValue('ShipSeqNum')%></h2>
  </td></tr></table><%
  revise();
}

## This is a revision email sent out by the logistics coordinator once a precruise has been updated in EXPD
elsif ( GetFormValue('step') eq 'update') {
  %><h2 align="center">Results of <%=GetFormValue('ShipName')%><%=GetFormValue('ShipSeqNum')%> revision</h2>
  </td></tr></table><%
  precruise_mail_out('REVISED:') unless update();
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
Do not go onto step 2 until all these fields are properly filled.

Author: Mike McCann

Date Created: 10/21/98

Redesigned: 9 Feb 1999 (combined with old step2())

Edited by: Karen Salamy, MBARI

February 2016

=cut

sub step1 {

  require 'timelocal.pl';

  my @l = localtime(time+864000); # Add 10 days
  my $month_10 = $l[4] + 1; # Because range is 0..11
  my $day_10 = $l[3];

  my $lstr = localtime(time+864000);  # Arghhh! get 2 digit yr from above!
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
<!-- for IE support -->
<script src="./es6-promise/promise.min.js"></script>
<!-- SET UP SWEETALERTS2 -->
<script src="./sweetalert2_old/sweetalert2.min.js"></script>
<link rel="stylesheet" href="sweetalert2_old/sweetalert2.min.css">
<!--<script src="./sweetalert2/dist/sweetalert2.min.js"></script>
<link rel="stylesheet" href="@sweetalert2/themes/borderless/borderless.css">
<link rel="stylesheet" href="@sweetalert2/themes/bootstrap-4/bootstrap-4.css"> -->
<!--<link rel="stylesheet" href="./sweetalert2/dist/sweetalert2.min.css"> -->
<!-- SET UP FONT AWESOME FOR ARROWS -->
<!-- <script src="https://use.fontawesome.com/a26ffc8b21.js"></script> -->


<!-- Javascript src="http://code.jquery.com/jquery-1.10.1.min.js"-->
<script Language="JavaScript"><!--
var nonMBARI = new Boolean(false);
var foreignNationals = new Boolean(false);
var equip = Array();
var winch = Array();
var platform = Array();
var isotopeEquip = "";
var itarEquip = "";
var hazmatEquip = "";
var mooringRecoveryEquip = "";
var mooringDeployEquip = "";
var nonMBARIPersonnel = "";
var foreignNationalPersonnel = "";
var telepresencePersonnel = "";
var plannedTrackExtent = "";
var plannedTrackRestricted = "";

var mbnmsPermitCheck = "";
var samplePermitCheck = "";
var smpaPermitCheck = "";
var smpaCollectionPermitCheck = "";
var cdfwPermitCheck = "";
var diver = "NO";
var purposeStr = "";
var dc = "";
var telepres = "";

// Field_Validator Function Checks required empty text boxes on Submit
function Field_Validator(theForm)
{
  if (theForm.ShipName.value == "" )
  {
    theForm.ShipName.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please select a Ship.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }

  if (theForm.ExpdPrincipalInvestigator.value == "")
  {
    theForm.ExpdPrincipalInvestigator.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please select a principle investigator.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
  if (theForm.ExpdChiefScientist.value == "")
  {
    theForm.ExpdChiefScientist.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please select a Chief Scientist.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
  if (theForm.ProjNum.value == "")
  {
    theForm.ProjNum.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please enter an MBARI Project Number.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
////////////////////////////////////////////  ADDITIONS  ///////////////////////////


////////////////////////////////  CRUISE PURPOSE AND DIVER REQUEST  /////////////////////////
  if (theForm.Purpose.value == "")
  {
    theForm.Purpose.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please enter cruise purpose.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
  if (theForm.Purpose.value != "") {
    joinCruisePurpose();
  }

  // EMAIL ALIASING FOR DIVER REQUESTED
  if (document.getElementById('diverRequested').value == "YES") {
    // Add to cruise purpose text area.
    joinCruisePurpose();
  }

   if (document.getElementById('diverRequested').value == "NO") {
    // Add to cruise purpose text area.
    joinCruisePurpose();
  }


  if (document.getElementById('telepresenceRequested').value == "YES") {
    // Add to cruise purpose text area.
    joinCruisePurpose();
  }

   if (document.getElementById('diverRequested').value == "NO") {
    // Add to cruise purpose text area.
    joinCruisePurpose();
  }


////////////////////////////    EQUIPMENT  ////////////////////////////////////////
  // Text entered into text box but no equipment checkboxes checked
/*  if (theForm.EquipmentDesc.value != "") {
    if (equip.length == 0) {
      document.getElementById('other-equipment').focus();
      // SWEETALERT
      swal({
        html:
          '<br> ' +
          '<b>Please choose all of the appropriate equipment you are requesting for this cruise. If none, please leave text area blank.</b>',
        showCloseButton: true,
        showCancelButton: false,
        confirmButtonText:
          'Ok',
        allowOutsideClick: false
      })
      return (false);
    }
  }*/

  if (theForm.EquipmentDesc.value == "") {
    if (equip.length == 0) {
        // var ed = "None";
        // document.getElementById('EquipmentDesc').value = ed;
        joinEquipment();
      }
  }

  // Only Other checked but no equipment details entered into Text Box
  if ((equip.length == 1) && (document.getElementById('other-equipment').checked == true )) {
    if (theForm.EquipmentDesc.value == "") {
      theForm.EquipmentDesc.focus();
      // SWEETALERT
      swal({
        html:
          '<br> ' +
          '<b>Please describe "other" equipment requested in text area.</b>',
        showCloseButton: true,
        showCancelButton: false,
        confirmButtonText:
          'Ok',
        allowOutsideClick: false
      })
      return (false);
    }
  }


  if (theForm.EquipmentDesc.value != "") {
    joinEquipment();
  }

  if (document.getElementById('asv').checked == true) {
    joinEquipment();
  }

  if (document.getElementById('asv').checked == false) {
    joinEquipment();
  }

  if (document.getElementById('ctd').checked == true) {
    joinEquipment();
  }

  if (document.getElementById('ctd').checked == false) {
    joinEquipment();
  }


////////////////////////////    WINCHES  ////////////////////////////////////////
  // Text entered into text box but no winch checkboxes checked
  // if (theForm.WinchDesc.value != "") {
  //   if (winch.length == 0) {
  //     document.getElementById('other-winch').focus();
  //     // SWEETALERT
  //     swal({
  //       html:
  //         '<br> ' +
  //         '<b>Please choose all of the winch equipment you are requesting for this cruise. If none, please leave text area blank.</b>',
  //       showCloseButton: true,
  //       showCancelButton: false,
  //       confirmButtonText:
  //         'Ok',
  //       allowOutsideClick: false
  //     })
  //     return (false);
  //   }
  // }

  if (theForm.WinchDesc.value == "") {
    if (winch.length == 0) {
        // var wd = "None";
        // document.getElementById('WinchDesc').value = wd;
        joinEquipment();
      }
  }

  // Only Other checked by no winch details entered into Text Box
  if ((winch.length == 1) && (document.getElementById('other-winch').checked == true )) {
    if (theForm.WinchDesc.value == "") {
      theForm.WinchDesc.focus();
      // SWEETALERT
      swal({
        html:
          '<br> ' +
          '<b>Please describe "other" winch requested in text area.</b>',
        showCloseButton: true,
        showCancelButton: false,
        confirmButtonText:
          'Ok',
        allowOutsideClick: false
      })
      return (false);
    }
  }


  if (theForm.WinchDesc.value != "") {
      joinEquipment();
  }

   if ((document.getElementById('paragonWinch').checked == true) || (document.getElementById('downriggerWinch').checked == true) || (document.getElementById('black-dynacon-winch').checked == true) || (document.getElementById('borehole-winch').checked == true) || (document.getElementById('macArtney-mooring-winch').checked == true) || (document.getElementById('other-winch').checked == true) || (document.getElementById('packard-dynacon-winch').value == '.322_CTD_WINCH') || (document.getElementById('packard-dynacon-winch').value == '.25_TOW_WIRE_WINCH') || (document.getElementById('packard-dynacon-winch').value == 'SYNTHETIC_ROPE_WINCH')) {
    joinEquipment();
  }


  //////////////////////////    TELEPRESENCE     /////////////////////
  ////////////////////////////////////////////////////////////////////
  if (theForm.TelepresenceParticipantsDesc.value == "") {
    if (platform.length == 0) {
        // var ed = "None";
        // document.getElementById('EquipmentDesc').value = ed;
        joinEquipment();
      }
  }

  // Only Other checked but no equipment details entered into Text Box
  if (document.getElementById('other-telepresence-equipment').checked == true ) {
    if (theForm.BroadcastPlatformDesc.value == "") {
      theForm.BroadcastPlatformDesc.focus();
      // SWEETALERT
      swal({
        html:
          '<br> ' +
          '<b>Please describe "other" broadcast media requested in text area.</b>',
        showCloseButton: true,
        showCancelButton: false,
        confirmButtonText:
          'Ok',
        allowOutsideClick: false
      })
      return (false);
    }
  }

////////////////////////////////////////////////////////////////////////////////////////////////
  ///////////////////////////////////////    ISOTOPES   ////////////////////////////////////////
  if (((document.getElementById('isotopes').value=="stable-isotope") || (document.getElementById('isotopes').value=="radio-isotope") || (document.getElementById('isotopes').value=="radio-stable-isotope")) && (theForm.IsotopeDesc.value == "")) {
    theForm.IsotopeDesc.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please enter Stable or Radioactive Isotopes planned for the cruise.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }
  if ((document.getElementById('isotopes').value=="no-isotope") && (theForm.IsotopeDesc.value != "")) {
    document.getElementById('isotopes').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please select an option in the Isotopic Description dropdown box.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
  return(false);
  }

  if ((document.getElementById('isotopes').value!="no-isotope") || (document.getElementById('hazMat').value!="no-hazMat")) {
    joinEquipment();
  }

  if (theForm.IsotopeDesc.value != "") {
    joinEquipment();
  }

  ///////////////////////////////////////    HAZMAT   ////////////////////////////////////////
  if ((document.getElementById('hazMat').value=="yes-hazMat") && (theForm.HazMatDesc.value == "")) {
    theForm.HazMatDesc.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please enter the Hazardous Material that you plan to bring on the cruise.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }
  if ((document.getElementById('hazMat').value=="no-hazMat") && (theForm.HazMatDesc.value != "")) {
    document.getElementById('hazMat').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please select Yes in the Hazardous Materials dropdown box.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }
  if (theForm.HazMatDesc.value != "") {
      joinEquipment();
  }

  ///////////////////////////////////////    ITAR   ////////////////////////////////////////
  if ((document.getElementById('itar').value=="yes-itar") && (theForm.ItarDesc.value == "")) {
    theForm.ItarDesc.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please enter the Equipment Requested that is ITAR Controlled that you plan to bring on the cruise.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }
  if ((document.getElementById('itar').value=="no-itar") && (theForm.ItarDesc.value != "")) {
    document.getElementById('itar').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please select Yes in the ITAR Equipment dropdown box.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }
  if (theForm.ItarDesc.value != "") {
      joinEquipment();
  }

  ///////////////////////////////////////  MOORING  /////////////////////////////////////
  // RECOVERY  //
  if (((document.getElementById('mooring-recovery').value=="yes-mooring-recovery-surface") || (document.getElementById('mooring-recovery').value=="yes-mooring-recovery-subsea")) && (theForm.MooringRecoveryDesc.value == "")) {
    theForm.MooringRecoveryDesc.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please describe the MOORING RECOVERY effort in detail.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }
  if (((document.getElementById('mooring-recovery').value!="yes-mooring-recovery-surface") && (document.getElementById('mooring-recovery').value!="yes-mooring-recovery-subsea")) && (theForm.MooringRecoveryDesc.value != "")) {
    document.getElementById('mooring-recovery').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please select either Surface or Subsea Mooring Recovery in dropdown box.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }
  if (theForm.MooringRecoveryDesc.value != "") {
      joinEquipment();
  }

  // DEPLOY ///
  if (((document.getElementById('mooring-deploy').value=="yes-mooring-deploy-surface") || (document.getElementById('mooring-deploy').value=="yes-mooring-deploy-subsea")) && (theForm.MooringDeployDesc.value == "")) {
    theForm.MooringDeployDesc.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please descibe the MOORING DEPLOYMENT effort in detail.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }
  if (((document.getElementById('mooring-deploy').value!="yes-mooring-deploy-surface") && (document.getElementById('mooring-deploy').value!="yes-mooring-deploy-subsea")) && (theForm.MooringDeployDesc.value != "")) {
    document.getElementById('mooring-deploy').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please select either Surface or Subsea Mooring Deployment in dropdown box.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }
  if (theForm.MooringDeployDesc.value != "") {
      joinEquipment();
  }

  /////////////////////////////////////  NON-MBARI AND FOREIGN NATIONAL PARTICIPANTS  /////////////////////////////////////
  if (theForm.Participants.value == "")
  {
    theForm.Participants.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please enter MBARI Participants. List an (*) by TWIC cardholders.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
  if (theForm.Participants.value != "") {
    joinPersonnel();
  }

  if ((document.getElementById('ForeignNationalParticipantsDesc').value!="") || (document.getElementById('foreign-national-personnel').value=="yes-foreign-national-personnel")) {
    joinPersonnel();   //EMAIL ALIASING FOR PARTICIPANTS SECTION
  }

  if ((document.getElementById('non-mbari-personnel').value=="yes-non-mbari-personnel") && (theForm.nonMBARIParticipantsDesc.value == "")) {
    theForm.nonMBARIParticipantsDesc.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please enter ALL Non-MBARI Personnel (US Citizens).</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }

  if ((document.getElementById('non-mbari-personnel').value=="no-non-mbari-personnel") && (theForm.nonMBARIParticipantsDesc.value != "")) {
    document.getElementById('non-mbari-personnel').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please choose YES in the Non MBARI dropdown box to proceed.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }

  if (theForm.nonMBARIParticipantsDesc.value != "") {
    joinPersonnel();
  }

  if ((document.getElementById('foreign-national-personnel').value=="yes-foreign-national-personnel") && (theForm.ForeignNationalParticipantsDesc.value == "")) {
    theForm.ForeignNationalParticipantsDesc.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please enter ALL Foreign National Personnel scheduled to sail on the cruise.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }

  if ((document.getElementById('foreign-national-personnel').value=="no-foreign-national-personnel") && (theForm.ForeignNationalParticipantsDesc.value != "")) {
    document.getElementById('foreign-national-personnel').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please choose YES in the Foreign Nationals dropdown box to proceed.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }

  if (theForm.ForeignNationalParticipantsDesc.value != "") {
    joinPersonnel();
  }



  /////////////////////////////////////// TELEPRESENCE PARTICIPANTS CHECKS ////////////////////////

  if ((document.getElementById('telepresenceRequested').value == "YES") && (theForm.TelepresenceParticipantsDesc.value == "")) {
    //console.log("Test Telepresence Participants");
    theForm.TelepresenceParticipantsDesc.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please list ALL participants that will be participating on shore.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }

//////////////////////////////////
  if ((document.getElementById('broadcast-publicly').value == "yes-broadcast") && (platform == undefined || platform.length == 0)) {
        //console.log("Test Telepresence Participants");
        theForm.BroadcastPlatformDesc.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please check ALL of the public media platforms you intend to use for this expedition.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }


  if ((document.getElementById('telepresenceRequested').value == "YES") && (theForm.TelepresenceParticipantsDesc.value != "") && (document.getElementById('broadcast-publicly').value == "none")) {
    document.getElementById('broadcast-publicly').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please choose YES or NO in the Telepresence Public Broadcast dropdown box to proceed.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }

  if ((document.getElementById('other-telepresence-equipment').checked == true) && (theForm.TelepresenceParticipantsDesc.value == "")) {
    theForm.TelepresenceParticipantsDesc.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please describe other public media platform requested.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }


  /////////////////////////////////////// PLANNED TRACK DESCRIPTION ENTRIES ///////////////////////

  if (theForm.PlannedTrackDesc.value != "") {
    joinPlannedTrackInfo();
  }
  if ((document.getElementById('planned-track-extent').value!="NO") && (document.getElementById('planned-track-extent').value!="YES")) {
    document.getElementById('planned-track-extent').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please state whether your planned track will extend outside of Monterey Bay (beyond M2) and involve underwater operations.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
  if (document.getElementById('planned-track-extent').value != "none") {
    joinPlannedTrackInfo();
  }

  if ((document.getElementById('planned-track-extent').value == "track-outside-MB") || (document.getElementById('planned-track-restricted').value == "yes-restricted-track")) {
    joinPlannedTrackInfo();   //EMAIL ALIASING SPECIAL OPS CHECK
  }


  if ((document.getElementById('planned-track-restricted').value!="NO") && (document.getElementById('planned-track-restricted').value!="YES")) {
    document.getElementById('planned-track-restricted').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please state whether you are planning on operating in a US Naval Restricted area, Pacific Missile Range, or shipping lanes.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
    if (document.getElementById('planned-track-restricted').value != "none") {
    joinPlannedTrackInfo();
  }


  // PERMITS CHECKING
  //if ((document.getElementById('mbnms-permit-check').value!="YES") && (document.getElementById('mbnms-permit-check').value!="NOT_SURE")) {
  if (document.getElementById('mbnms-permit-check').value=="none") {
    document.getElementById('mbnms-permit-check').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please state whether this operation complies with MBARI\'s MBNMS permit.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
    if (document.getElementById('mbnms-permit-check').value != "none") {
    joinPlannedTrackInfo();
  }

  if (document.getElementById('mbnms-permit-check').value == "NOT_SURE")  {
    joinPlannedTrackInfo();   //EMAIL ALIASING PERMITS CHECK
  }

  //if ((document.getElementById('cdfw-permit-check').value!="YES") && (document.getElementById('cdfw-permit-check').value!="NOT_SURE")) {
    if ((document.getElementById('cdfw-permit-check').value!="YES") && (document.getElementById('cdfw-permit-check').value!="NO") && (document.getElementById('cdfw-permit-check').value!="NOT_SURE") && (document.getElementById('cdfw-permit-check').value!="NOT_APPLICABLE")) {
    document.getElementById('cdfw-permit-check').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please state whether you have a valid California Department of Fish & Wildlife permit to collect samples during your cruise.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
  if (document.getElementById('cdfw-permit-check').value != "none") {
    joinPlannedTrackInfo();
  }
  if (document.getElementById('cdfw-permit-check').value == "YES")  {
    joinPlannedTrackInfo();   //EMAIL ALIASING PERMITS CHECK
  }
  if (document.getElementById('cdfw-permit-check').value == "NO")  {
    joinPlannedTrackInfo();   //EMAIL ALIASING PERMITS CHECK
  }
  if (document.getElementById('cdfw-permit-check').value == "NOT_SURE")  {
    joinPlannedTrackInfo();   //EMAIL ALIASING PERMITS CHECK
  }
  if (document.getElementById('cdfw-permit-check').value == "NOT_APPLICABLE")  {
    joinPlannedTrackInfo();   //EMAIL ALIASING PERMITS CHECK
  }


  //if ((document.getElementById('smpa-collection-permit-check').value!="YES") && (document.getElementById('smpa-collection-permit-check').value!="NOT_SURE")) {
  if (document.getElementById('smpa-collection-permit-check').value=="none") {
    document.getElementById('smpa-collection-permit-check').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please state whether you have a proper CDFW permit to operate or take samples in a State Marine Protected Area (SMPA).</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
    if (document.getElementById('smpa-collection-permit-check').value != "none") {
    joinPlannedTrackInfo();
  }

  if (document.getElementById('smpa-collection-permit-check').value == "NOT_SURE")  {
    joinPlannedTrackInfo();   //EMAIL ALIASING PERMITS CHECK
  }


  ///////////////////////////////////////  AGREE TO TERMS /////////////////////////////////////
  if (((document.getElementById('isotopes').value!="no-isotope") || (document.getElementById('hazMat').value!="no-hazMat")) && (document.getElementById('isotopePolicyAgree').checked == false)) {
    document.getElementById('isotopePolicyAgree').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please indicate that you agree to the MBARI HAZMAT and Isotope Policy for Research Vessels before submitting this precruise.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }


  if (document.getElementById('MBARIPolicyAgree').checked == false) {
    theForm.MBARIPolicyAgree.focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please indicate that you agree to all MBARI Precruise Policies and Procedures before submitting this precruise.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return(false);
  }


 // ALERT IF NO PLANNED TRACK IS ENTERED BEFORE HITTING CONTINUE BUTTON
  if (theForm.PlannedTrackDesc.value == "") {
    // Show alert only once.
    if (!sessionStorage.returning_ptr_pre) {
    // SWEETALERT
    swal({
      //title: "Are you sure?",
      html:
         '<br> ' +
         '<b>NOTE: By not entering a Planned Track Description here in Step 1, you are required to Select your cruise Waypoints in Step 2 after clicking the Continue Button below.</b>',
      showCloseButton: true,
      showCancelButton: true,
      confirmButtonText:
        'Got it!',
        cancelButtonText:
        'Edit Planned Track',
      allowOutsideClick: false
    }).then(function (dismiss) {
    if (dismiss === 'cancel' || dismiss == 'close') {
      theForm.PlannedTrackDesc.focus(); // Probably can delete this line....
    }
  }).catch(function() {
    theForm.PlannedTrackDesc.focus();
  });
      sessionStorage.returning_ptr_pre = true; // set returning
      return (false);
    } // End of sessionStorage.returning if
    return (true);
  } // End of if
} // End of Field_Validator function



// Base form on which ship is selected at beginning page.
function ShipSelect(select)
{
  //Pass in new function to configure the form
  ConfigForm(select.options[select.selectedIndex].value);
  restartFlash();

  if (select.options[select.selectedIndex].value == "dpkd") {
    document.getElementById("planned-track-extent").value = "none";
    document.getElementById("planned-track-restricted").value = "none";
    document.getElementById("precruise_plans_link").href = "https://mww.mbari.org/dmo/Pages/DPPlanning.html";
    document.getElementById("dmo_safety_link").href = "https://mww.mbari.org/dmo/Pages/SafetyManagementSystem.html";
    document.getElementById("mbari_ship_policy_link").href = "https://mww.mbari.org/dmo/Pages/DMOPolicies.html";
    // Show alert only once.
    if (!sessionStorage.returning_dpkd_pre) {
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b><i>R/V David Packard</i> Cruises, Departure & Arrival times are listed </b>' +
        '<a target="_blank" href="https://www.mbari.org/our-work/schedules/">HERE.</a> ',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Got it!',
      allowOutsideClick: false
    })  // END OF swal()
      sessionStorage.returning_dpkd_pre = true; // set returning
    } // End of sessionStorage.returning if
    return (true);
  } // End of if
  if (select.options[select.selectedIndex].value == "rcsn") {
    document.getElementById("planned-track-extent").value = "none";
    document.getElementById("planned-track-restricted").value = "none";
    document.getElementById("precruise_plans_link").href = "https://mww.mbari.org/dmo/Pages/RCPlanning.html";
    document.getElementById("dmo_safety_link").href = "https://mww.mbari.org/dmo/Pages/SafetyManagementSystem.html";
    document.getElementById("mbari_ship_policy_link").href = "https://mww.mbari.org/dmo/Pages/DMOPolicies.html";

// Show alert only once per session
    if (!sessionStorage.returning_rcsn_pre) {
      swal({
        html:
          '<br> ' +
          '<b>For <i>R/V Rachel Carson</i> cruises, please confirm Departure and Arrival times with the ship\'s captain before entering them here. </b>',
        showCloseButton: true,
        showCancelButton: false,
        confirmButtonText: 'Ok',
        allowOutsideClick: false
      });
      sessionStorage.returning_rcsn_pre = true; // Set the flag so it doesn't show again
    }
    return (true);
  }

  //   swal({
  //     html:
  //       '<br> ' +
  //       '<b>For <i>R/V Rachel Carson</i> cruises, please confirm Departure and Arrival times with the ship\'s captain before entering them here. </b>',
  //     showCloseButton: true,
  //     showCancelButton: false,
  //     confirmButtonText:
  //       'Ok',
  //     allowOutsideClick: false
  //   })
  //   return (true);
  // }



  if (select.options[select.selectedIndex].value == "prgn") {
    document.getElementById("planned-track-extent").value = "NO";
    document.getElementById("planned-track-restricted").value = "NO";
    document.getElementById("precruise_plans_link").href = "https://mww.mbari.org/dmo/Pages/Paragon.html";
    document.getElementById("dmo_safety_link").href = "https://mww.mbari.org/dmo/Pages/SafetyManagementSystem.html";
    document.getElementById("mbari_ship_policy_link").href = "https://mww.mbari.org/dmo/Pages/Paragon.html";
    // Show alert only once.
    if (!sessionStorage.returning_prgn_pre) {
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>For <i>R/V Paragon</i> reservations, please first confirm that your captain is an approved operator </b>' +
        //'<a target="_blank" href="//www.mbari.org/at-sea/ships/research-vessel-paragon/paragon-cruise-planning/">HERE.</a> ',
        '<a target="_blank" href="https://docs.google.com/spreadsheets/d/14MTXwKS2jCNEKfAs1nFwCT5Wc0FaL7FY7Rw5zzmkakg/edit?usp=sharing/">HERE.</a> ',


      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Got it!',
      allowOutsideClick: false
    }) // End of swal()
      sessionStorage.returning_prgn_pre = true; // set returning
  }
    return (true);
  }
  else {
    swal({
      html:
        '<br> ' +
        '<b>You must first set the Ship Name field to continue.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
    return (false);
  }
}


// CLEAR LINKS
// function clearLinks() {
//     document.getElementById("precruise-link").href = "javascript: void(0)";
//     document.getElementById("precruise-link").style.color = "gray";
//     document.getElementById("precruise-link").style.textDecoration = "none";
// }



// Update Diver Request
function updateDiverRequest(text) {
  var diver = text;
  joinCruisePurpose();
}


// Update Telepresence Request
function updateTelepresenceRequest(text) {
  var telepresence = text;
  //console.log("telepresence = " + telepresence);
  joinCruisePurpose();
  updateTelepresenceParticipant(telepresence);
  joinPersonnel();
}

/*function translateEntities(str){
    var degreeMin = str.replace('deg', '&deg;').replace('min', '&rsquo;');
    //console.log("Degree/Min: " + degreeMin);
    return str.replace('deg', '&deg;').replace('min', '&rsquo;');*/


    // var text, p=document.createElement('p');
    // p.innerHTML=string;
    // text= p.innerText || p.textContent;
    // p.innerHTML='';
    // return text;
 //}



function strip_tags(str, allow) {
    // Making sure the allowed arg is a string containing only tags in lowercase (<a><b><c>)
    allow = (((allow || "") + "").toLowerCase().match(/<[a-z][a-z0-9]*>/g) || []).join('');

    var tags = /<\/?([a-z][a-z0-9]*)\b[^>]*>/gi;
    var commentsAndPhpTags = /<!--[\s\S]*?-->|<\?(?:php)?[\s\S]*?\?>/gi;

    return str.replace(/\u00B0/g, 'deg').replace(/\u2019/g, 'min').replace(commentsAndPhpTags, '').replace(tags, function ($0, $1) {
        return allow.indexOf('<' + $1.toLowerCase() + '>') > -1 ? $0 : ''; });
}



// Join all Cruise purpose information into one text area.
function joinCruisePurpose(){
  var y = "";
  var z = "";
  var tp = "";
  //var cruisePurposeTxt = "";
  var diverTxt = "\nDiver Requested: ";
  var telepresenceTxt = "\nTelepresence Requested: ";

  // Cruise Purpose text
  var p = document.getElementById("Purpose").value;
  if (document.getElementById('Purpose').value != "") {
    y = p;
  }

  // Diver Requested checkbox checked?
  dc = document.getElementById('diverRequested').value;  //Yes or No
  //console.log("DC = " + dc);
  if (dc == "YES") {
    z = diverTxt.bold() + dc;
  }
  if (dc == "NO") {
  }

    // Telepresence Requested checkbox checked?
  telepres = document.getElementById('telepresenceRequested').value;  //Yes or No
  //console.log("DC = " + dc);
  if (telepres == "YES") {
    // var telepresenceYES = "YES";
    // telepres = diveYES;
    tp = telepresenceTxt.bold() + telepres;
    //console.log("Bolded???  " + tp);
    //localStorage.setItem('telepresenceRequestedStore', 'YES');
  }
  if (telepres == "NO") {  //If no, don't add
     //document.getElementById('telepresence-requested').value = "NO";
  }

  //console.log("DC2 = " + dc);
  document.getElementById('allCruisePurposeInfo').value = y + " " + z + " " + tp;
  var acp = document.getElementById('allCruisePurposeInfo').value;
  purposeStr = strip_tags(acp);
  document.getElementById('allCruisePurposeInfo_cleanHTML').value = purposeStr;
  //console.log("CleanHTML Purpose: " + purposeStr);
}



// Function called everytime Other-Equipment Checkbox is checked
function enableDisableEquip() {
  var cbequip = document.getElementById("other-equipment");
  var txtequip = document.getElementById("EquipmentDesc");
     if (cbequip.checked == true) {
        //console.log("Other Equipment Checkbox Clicked! Textbox ENABLED");
        txtequip.readOnly = false;
        txtequip.focus();
        updateRequestedEquip(this.value);
     }
     else if (cbequip.checked == false) {
        updateRequestedEquip(this.value);
      }
    updateRequestedEquip(this.value);
    }


// Function called everytime Other-Winch Checkbox is checked
function enableDisableWinch() {
  var cbwinch = document.getElementById("other-winch");
  var txtwinch = document.getElementById("WinchDesc");
     if (cbwinch.checked == true) {
        //console.log("Other Winch Checkbox Clicked! Textbox ENABLED");
        txtwinch.readOnly = false;
        txtwinch.focus();
        updateRequestedWinch(this.value);
     } else if (cbwinch.checked == false) {
        updateRequestedWinch(this.value);
      }
  updateRequestedWinch(this.value);
}


// Function called everytime Other-Broadcast TELEPRESENCE Platform Checkbox is checked
function enableDisablePlatforms() {
  var cbplats = document.getElementById("other-telepresence-equipment");
  var txtplats = document.getElementById("BroadcastPlatformDesc");
     if (cbplats.checked == true) {
        //document.getElementById('BroadcastPlatformDesc').value = "";
        //console.log("Other Broadcast Platforms Checkbox Clicked! Textbox ENABLED");
        txtplats.readOnly = false;
        txtplats.focus();
        updateRequestedPlatforms(this.value);
     } else if (cbplats.checked == false) {
        updateRequestedPlatforms(this.value);
      }
    updateRequestedPlatforms(this.value);
}

function checkEquipOther() {
/*  if (document.getElementById('other-equipment').checked == false) {
    document.getElementById('other-equipment').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please select Other Equipment Checkbox first.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
  }*/
}

function checkWinchOther() {
/*  if (document.getElementById('other-winch').checked == false) {
    document.getElementById('other-winch').focus();
    // SWEETALERT
    swal({
      html:
        '<br> ' +
        '<b>Please select Other Winch Checkbox first.</b>',
      showCloseButton: true,
      showCancelButton: false,
      confirmButtonText:
        'Ok',
      allowOutsideClick: false
    })
  }*/
}


// Add up all of the Vehicle Equipment requests
function updateRequestedEquip(text) {

  // Array to store value of selected boxes
  equip = [];

  // Get all the elements with equipSelect[] array name tag and determine length
  var selectEquip = document.getElementsByName('equipSelect[]');
  var nr_selectEquip = selectEquip.length;

  // Traverse the selectEquip elements and add (or remove) the value of selected checked checkboxes to equip[]
  for (var i=0; i<nr_selectEquip; i++) {
    if(selectEquip[i].type == 'checkbox' && selectEquip[i].checked == true) {
      equip.push(selectEquip[i].value);
      //console.log("selectEquip = " + selectEquip[i].value);
    } else if (selectEquip[i].value == "RAD_VAN" || selectEquip[i].value == "CHEM_VAN" ) {
      equip.push(selectEquip[i].value);
      //console.log("selectEquip = " + selectEquip[i].value);
    }
    //console.log("updateRequestedEquip() = " + equip);
  }
    joinEquipment();
}


// Add up all of the Telepresence Platform Broadcast Equipment requests
function updateRequestedPlatforms() {
    // Array to store value of selected boxes
    platform = [];

    // Get all the elements with equipSelect[] array name tag and determine length
    var selectTPPlatforms = document.getElementsByName('platformSelect[]');
    var nr_selectTPPlatforms = selectTPPlatforms.length;

    // Traverse the selectPlatforms elements and add (or remove) the value of selected checked checkboxes to platform[]
    for (var i=0; i<nr_selectTPPlatforms; i++) {
    if(selectTPPlatforms[i].type == 'checkbox' && selectTPPlatforms[i].checked == true) {
      platform.push(selectTPPlatforms[i].value);
    }
  }
    joinEquipment();
}


function updateBroadcastPublicly(text) {
  //console.log("UPDATE: updateBroadcastPublicly: " + text);
  var ubp = text;

  if (ubp == "NO") {
    document.getElementById('broadcastPlatformsRequested').style.display = "none";
    document.getElementById('listTPPlatforms').style.display = "none";
    document.getElementById('BroadcastPlatformDesc').style.display = "none";
    document.getElementById('broadcastTP-check').value = ubp;
  } else if (ubp == "YES") {
    document.getElementById('broadcastPlatformsRequested').style.display = "inline";
    document.getElementById('listTPPlatforms').style.display = "inline";
    document.getElementById('BroadcastPlatformDesc').style.display = "inline";
    document.getElementById('broadcastTP-check').value = ubp;
  } else if (ubp == "none") {
    document.getElementById('broadcastPlatformsRequested').style.display = "none";
    document.getElementById('listTPPlatforms').style.display = "none";
    document.getElementById('BroadcastPlatformDesc').style.display = "none";
    document.getElementById('broadcastTP-check').value = ubp;
  }

//document.getElementById('broadcastTP-check').value = "YES";
  checkBroadcastPlatform();
  ConfigForm();
  joinCruisePurpose();
  joinEquipment();
}

function checkBroadcastPlatform() {
   //Try and get return from Step 2 variables
   //console.log("CHECKING BROADCAST PLATFORMS");
  if (document.getElementById('broadcastTP-check').value == "YES") {
        updateRequestedPlatforms("YES");
        document.getElementById('youTube-media-label').style.display = "inline";
        document.getElementById('broadcastPlatformsRequested').style.display = "inline";
        document.getElementById('listTPPlatforms').style.display = "inline";
        document.getElementById('BroadcastPlatformDesc').style.display = "inline";
  } else if (document.getElementById('broadcastTP-check').value == "NO") {
        updateRequestedPlatforms("NO");
        document.getElementById('youTube-media-label').style.display = "none";
        document.getElementById('broadcastPlatformsRequested').style.display = "none";
        document.getElementById('listTPPlatforms').style.display = "none";
        document.getElementById('BroadcastPlatformDesc').style.display = "none";
  } else if (document.getElementById('broadcastTP-check').value == "none") {
        updateRequestedPlatforms("none");
        document.getElementById('youTube-media-label').style.display = "none";
        document.getElementById('broadcastPlatformsRequested').style.display = "none";
        document.getElementById('listTPPlatforms').style.display = "none";
        document.getElementById('BroadcastPlatformDesc').style.display = "none";  }
      //console.log("Broadcast value: " + document.getElementById("broadcast-publicly").value);
}

function checkBroadcastPlatformOther() {
  }


// Add up all of the Winch requests
function updateRequestedWinch() {

  // Array to store value of selected boxes
  winch = [];

  // Get all the elements with winchSelect[] array name tag and determine length
  var selectWinch = document.getElementsByName('winchSelect[]');
  var nr_selectWinch = selectWinch.length;

  // Traverse the selectWinch elements and add (or remove) the value of selected checked checkboxes to winch[]
  for (var i=0; i<nr_selectWinch; i++) {
    if(selectWinch[i].type == 'checkbox' && selectWinch[i].checked == true) {
      winch.push(selectWinch[i].value);
    } else if (selectWinch[i].value == ".322_CTD_WINCH" || selectWinch[i].value == ".25_TOW_WIRE_WINCH" || selectWinch[i].value == "SYNTHETIC_ROPE_WINCH") {
      winch.push(selectWinch[i].value);
    }
    //console.log("updateRequestedWinch() = " + winch);
  }
    joinEquipment();
}



// Add ISOTOPE MATERIALS to requests dialog
function updateIsotopeDeclaration(text) {
  isotopeEquip = text;
  if (document.getElementById('isotopes').value != "no-isotope") {
    IsotopeDesc.focus();
  }
  joinEquipment();
}


// Add HAZMAT MATERIALS to requests dialog
function updateHazMatDeclaration(text) {
  hazmatEquip = text;
  if (document.getElementById('hazMat').value != "no-hazMat") {
    HazMatDesc.focus();
  }
  joinEquipment();
}


// Add ITAR controlled equipment to requests dialog
function updateItarDeclaration(text) {
  itarEquip = text;
  if (document.getElementById('itar').value != "no-itar") {
    ItarDesc.focus();
  }
  joinEquipment();
}


// Add Mooring Recovery Description to requests dialog
function updateMooringRecoveryDeclaration(text) {
  mooringRecoveryEquip = text;
  if (document.getElementById('mooring-recovery').value != "no-mooring-recovery") {
    MooringRecoveryDesc.focus();
  }
  joinEquipment();
}


// Add ITAR controlled equipment to requests dialog
function updateMooringDeployDeclaration(text) {
  mooringDeployEquip = text;
  if (document.getElementById('mooring-deploy').value != "no-mooring-deploy") {
    MooringDeployDesc.focus();
  }
  joinEquipment();
}


// Join all Equipment chosen (vehicles and winches) into one hidden input section
function joinEquipment() {
  //console.log("CALLING joinEquipment()");

  // Grab whatever is in Equipment text area.
  var a = "";
  var b = "";
  var c = "";
  var d = "";
  var e = "";
  var f = "";
  var g = "";
  var h = "";
  var mrd1 = "";
  var mrd2 = "";
  var iso1 = "";
  var equipTxt = "Equipment Requested: ";
  var winchTxt = "\nWinch Requested: ";
  var isotopeTxt = "\nIsotopes: ";
  var hazmatTxt = "\nHazardous Materials: ";
  var itarTxt = "\nITAR Controlled Equipment: ";
  var mooringRecoveryTxt = "\nMooring Recovery Details: ";
  var mooringDeployTxt = "\nMooring Deployment Details: ";
  var broadcastTPPublicTxt = "\nBroadcast Telepresence Publicly: ";
  var telepTxt = "\nTelepresence Media Requested: ";
  var otherE = "\nEquipment Details: ";
  var otherW = "\nWinch Details: ";
  var otherT = "\nTelepresence Media Details: ";


  //Grab the string values of ALL Equipment text areas
  // Other Equipment Desc
  var v = document.getElementById("EquipmentDesc").value;
  //console.log("v = " + v );

  // Other Winch Desc
  var w = document.getElementById('WinchDesc').value;
  //console.log("w = " + w );

  // Telepresence
  // Broadcast Or Other Telepresence Desc
  var bp = document.getElementById('BroadcastPlatformDesc').value;
  //console.log("bp = " + bp );

  // Isotope Desc
  if (document.getElementById("isotopes").value != "no-isotope") {
    if (document.getElementById("isotopes").value == "stable-isotope") {
      iso1 = "STABLE_ISOTOPES \n";
    }
    if (document.getElementById("isotopes").value == "radio-isotope") {
      iso1 = "RADIO_ISOTOPES \n";
    }
    if (document.getElementById("isotopes").value == "radio-stable-isotopes") {
      iso1 = "RADIO_AND_STABLE_ISOTOPES \n";
    }
  }
  var iso = iso1 + document.getElementById('IsotopeDesc').value;
  //console.log("ISOTOPES: " + iso);
  //var iso = document.getElementById('IsotopeDesc').value;

  // Hazard Materials Desc
  var haz = document.getElementById('HazMatDesc').value;

  // ITAR Controlled Equipment Desc
  var it = document.getElementById('ItarDesc').value;

  // Mooring Recovery Desc
  if (document.getElementById("mooring-recovery").value != "no-mooring-recovery") {
    if (document.getElementById("mooring-recovery").value == "yes-mooring-recovery-surface") {
      mrd1 = "SURFACE_MOORING_RECOVERY \n";
    } else if (document.getElementById("mooring-recovery").value == "yes-mooring-recovery-subsea") {
      mrd1 = "SUBSEA_MOORING_RECOVERY \n";
    }
  }
  var mr = mrd1 + document.getElementById('MooringRecoveryDesc').value;
  //console.log("Mooring Recovery: " + mr);

  // Mooring Deployment Desc
  if (document.getElementById("mooring-deploy").value != "no-mooring-deploy") {
    if (document.getElementById("mooring-deploy").value == "yes-mooring-deploy-surface") {
      mrd2 = "SURFACE_MOORING_DEPLOYMENT \n";
    } else if (document.getElementById("mooring-deploy").value == "yes-mooring-deploy-subsea") {
      mrd2 = "SUBSEA_MOORING_DEPLOYMENT \n ";
    }
  }
  var md = mrd2 + document.getElementById('MooringDeployDesc').value;
  //console.log("Mooring Deployment: " + md);

// Equipment Checks
//////////////////////////
  // if (document.getElementById('EquipmentDesc').value == "") {
  //   document.getElementById('EquipmentDesc').value = "None";
  // }

  // Absolutely NOTHING ENTERED in Equipment section
  if ((equip == undefined || equip.length == 0) && (document.getElementById('EquipmentDesc').value == "")) {
    v = "None";
    // Nothing was chosen for equipment
    a = equipTxt.toString().bold() + v + otherE.toString().bold() + v;
  }
  // Nothing entered in Checkboxes - Text Entered in text area
  else if ((equip == undefined || equip.length == 0) && (document.getElementById('EquipmentDesc').value != "")) {
    // Nothing was chosen for equipment
    a = equipTxt.toString().bold() + "CUSTOM" + otherE.toString().bold() + v;
  }
  // Something entered in CheckBoxes - Nothing in Textarea
  else if ((equip.length > 0) && (document.getElementById('EquipmentDesc').value == "")) {
    v = "None";
    a = equipTxt.toString().bold() + equip.join(", ") + otherE.toString().bold() + v;
  }
  // Something entered in CheckBoxes (including Other) or Custom Equipment was chosen
  else if ((equip.length > 0) && (document.getElementById('EquipmentDesc').value != "")) { //&& (document.getElementById('EquipmentDesc').value != ""))&& (document.getElementById('other-equipment').checked == true)){
    a = equipTxt.toString().bold() + equip.join(", ") + otherE.toString().bold() + v;
    //console.log("2: " + a);
  }


// Winch Checks.
///////////////////////
  // if (document.getElementById('WinchDesc').value == "") {
  //   document.getElementById('WinchDesc').value = "None";
  // }

  // Absolutely NOTHING ENTERED in Equipment section
  if ((winch == undefined || winch.length == 0) && (document.getElementById('WinchDesc').value == "")) {
    w = "None";
    // Nothing was chosen for equipment
    b = winchTxt.toString().bold() + w + otherW.toString().bold() + w;
  }

  // Nothing entered in Checkboxes - Text Entered in textarea
  else if ((winch == undefined || winch.length == 0) && (document.getElementById('WinchDesc').value != "")) {
    // Nothing was chosen for equipment
    b = winchTxt.toString().bold() + "CUSTOM" + otherW.toString().bold() + w;
  }

  // Something entered in CheckBoxes - Nothing in Textarea
  else if ((winch.length > 0) && (document.getElementById('WinchDesc').value == "")) {
    w = "None";
    b = winchTxt.toString().bold() + winch.join(", ") + otherW.toString().bold() + w;
  }

  // Something entered in CheckBoxes (including Other) or Custom Equipment was chosen
  else if ((winch.length > 0) && (document.getElementById('WinchDesc').value != "")) { //&& (document.getElementById('WinchDesc').value != ""))&& (document.getElementById('other-winch').checked == true)){
    b = winchTxt.toString().bold() + winch.join(", ") + otherW.toString().bold() + w;
    //console.log("2: " + a);
  }




if (document.getElementById('broadcast-publicly').value=="yes-broadcast") {
      document.getElementById('broadcastTPCheck').value = "YES";
  }
  else if (document.getElementById('broadcast-publicly').value=="no-broadcast") {
      document.getElementById('broadcastTPCheck').value = "NO";
  }
  else if (document.getElementById('broadcast-publicly').value == "none") {
  }

if (document.getElementById("broadcast-publicly").value == "no-broadcast") {
  document.getElementById('broadcastTP-check').value = "NO";
}
if (document.getElementById("broadcast-publicly").value == "yes-broadcast") {
  document.getElementById('broadcastTP-check').value = "YES";
}
if (document.getElementById("broadcast-publicly").value == 'none') {
}


// Telepresence AND BROADCAST Checks.
  // Absolutely NOTHING ENTERED in Telepresence section
  if ((platform == undefined || platform.length == 0) && (document.getElementById('BroadcastPlatformDesc').value == "") && (document.getElementById("telepresenceRequested").value == "no-telepresence")) {
    bp = "None";
    // Nothing was chosen for equipment
    h = broadcastTPPublicTxt.toString().bold() + document.getElementById("telepresenceRequested").value + telepTxt.toString().bold() + bp + otherT.toString().bold() + bp;
  }

  // Nothing entered in Checkboxes - Text Entered in textarea
  else if ((platform == undefined || platform.length == 0) && (document.getElementById('BroadcastPlatformDesc').value != "")) {
    // Nothing was chosen for telepresence
    h = broadcastTPPublicTxt.toString().bold() + document.getElementById("telepresenceRequested").value + telepTxt.toString().bold() + "CUSTOM" + otherT.toString().bold() + bp;
  }

  // Something entered in CheckBoxes - Nothing in Textarea
  else if ((platform.length > 0) && (document.getElementById('BroadcastPlatformDesc').value == "")) {
    bp = "None";
    h = broadcastTPPublicTxt.toString().bold() + document.getElementById("telepresenceRequested").value + telepTxt.toString().bold() + platform.join(", ") + otherT.toString().bold() + bp;
  }

  // Something entered in CheckBoxes (including Other) or Custom Equipment was chosen
  else if ((platform.length > 0) && (document.getElementById('BroadcastPlatformDesc').value != "") && (document.getElementById("telepresenceRequested").value != "no-telepresence")) {
    h = broadcastTPPublicTxt.toString().bold() + document.getElementById("telepresenceRequested").value + telepTxt.toString().bold() + platform.join(", ") + otherT.toString().bold() + bp;
  }


//////////////////////////////////////////////////////////////
  // Isotopes
  if (document.getElementById('IsotopeDesc').value != "") {
    c = isotopeTxt.bold() + iso;
  }

  // HazMat
  if (document.getElementById('HazMatDesc').value != "") {
    d = hazmatTxt.bold() + haz;
  }

  // Itar
  if (document.getElementById('ItarDesc').value != "") {
    e = itarTxt.bold() + it;
  }

  // Mooring Recovery and Deploy
  if (document.getElementById('MooringRecoveryDesc').value != "") {
    f = mooringRecoveryTxt.bold() + mr;
  }

  if (document.getElementById('MooringDeployDesc').value != "") {
    g = mooringDeployTxt.bold() + md;
  }

  // Telepresence

  // Add them all up, grab the string, strip out the html characters and set value
  document.getElementById('allEquip').value = a + b + c + d + e + f + g + h;
  var ae = document.getElementById('allEquip').value;
  var equipStr = strip_tags(ae);
  document.getElementById('allEquip_cleanHTML').value = equipStr;


  ////////////////////////////////// EMAIL ALIASING ////////////////////////////////////////
  // BROADCAST TELEPRESENCE PUBLICLY CHECK?
  if (document.getElementById('broadcast-publicly').value != "none") {
    var broadcastTPCheck = document.getElementById('broadcast-publicly').value; //Yes or No
    document.getElementById('broadcastTP-check').value = broadcastTPCheck;

    if (document.getElementById('broadcast-publicly').checked == true) {
      document.getElementById('broadcastTP-check').value = "YES";
    }
    if (document.getElementById('broadcast-publicly').checked == false) {
      document.getElementById('broadcastTP-check').value = "NO";
    }
  }



  // WAVEGLIDER
  var waveG = document.getElementById('asv').value; //Yes or No
  document.getElementById('waveG').value = waveG;

  if (document.getElementById('asv').checked == true) {
    document.getElementById('waveG').value = "YES";
  }
  if (document.getElementById('asv').checked == false) {
    document.getElementById('waveG').value = "NO";
  }

  // CTD
  var ctdCheck = document.getElementById('ctd').value; //Yes or No
  document.getElementById('ctdCheck').value = ctdCheck;

  if (document.getElementById('ctd').checked == true) {
    document.getElementById('ctdCheck').value = "YES";
  }
  if (document.getElementById('ctd').checked == false) {
    document.getElementById('ctdCheck').value = "NO";
  }

  //ISOTOPES AND HAZMAT
  if ((document.getElementById('isotopes').value != "no-isotope") || (document.getElementById('hazMat').value != "no-hazMat"))  {
    document.getElementById('hazIsoCheck').value = "YES";
  } else {
    document.getElementById('hazIsoCheck').value = "NO";
  }

  //WINCH
  if ((document.getElementById('paragonWinch').checked == true) || (document.getElementById('downriggerWinch').checked == true) || (document.getElementById('black-dynacon-winch').checked == true) || (document.getElementById('borehole-winch').checked == true) || (document.getElementById('macArtney-mooring-winch').checked == true) || (document.getElementById('other-winch').checked == true) || (document.getElementById('packard-dynacon-winch').value == '.322_CTD_WINCH') || (document.getElementById('packard-dynacon-winch').value == '.25_TOW_WIRE_WINCH') || (document.getElementById('packard-dynacon-winch').value == 'SYNTHETIC_ROPE_WINCH')) {
      document.getElementById('winchCheck').value = "YES";
  } else {
      document.getElementById('winchCheck').value = "NO";
  }

  // ITAR
  if ((document.getElementById('itar').value == "yes-itar") || (document.getElementById('ItarDesc').value != "")) {
    document.getElementById('itarCheck').value = "YES";
  } else {
    document.getElementById('itarCheck').value = "NO";
  }

 // FOREIGN NATIONAL CHECK
 if ((document.getElementById('ForeignNationalParticipantsDesc').value!="") || (document.getElementById('foreign-national-personnel').value=="yes-foreign-national-personnel")) {
   document.getElementById('foreignNationalCheck').value = "YES";
 } else {
    document.getElementById('foreignNationalCheck').value = "NO";
 }
} // End of joinEquipment()




// Add Non MBARI Personnel to Participants database
function updateNonMBARIParticipant(text) {
  nonMBARIPersonnel = text;
  if (document.getElementById('non-mbari-personnel').value != "no-non-mbari-personnel") {
    nonMBARIParticipantsDesc.focus();
  }
  joinPersonnel();
}


// Add Foreign Nationals to Participants database
function updateForeignNationalParticipant(text) {
  foreignNationalPersonnel = text;
  if (document.getElementById('foreign-national-personnel').value != "no-foreign-national-personnel") {
    ForeignNationalParticipantsDesc.focus();
    //cruise-participants-section.focus();
  }
  joinPersonnel();
}


// Add Shore-Based Personnel to Participants database
function updateTelepresenceParticipant(text) {
  telepresencePersonnel = text;
  if (document.getElementById('telepresenceRequested').value == "no-telepresence") {
    document.getElementsByName('telepresence-requested').value = "NO";
  }
  if (document.getElementById('telepresenceRequested').value == "yes-telepresence") {
    document.getElementsByName('telepresence-requested').value = "YES";
  }
    ConfigForm();
    joinCruisePurpose();
    joinEquipment();
    joinPersonnel();
}

// Join all Equipment chosen (vehicles and winches) into one hidden input section
function joinPersonnel() {
  var j = "";
  var k = "";
  var l = "";
  var tpr = "";
  var mbariTxt = "MBARI Personnel: ";
  var nonMbariTxt = "\nNon-MBARI Personnel: ";
  var telepresenceParticipantsTxt = "\nTelepresent On-Shore Personnel: ";
  var foreignNationalTxt = "\nForeign Nationals: ";

  //Grab the string values of ALL Participant text areas
  // MBARI Participants Desc
  var p = document.getElementById('Participants').value;

  // NON MBARI Participants (US Citizens)
  var non = document.getElementById('nonMBARIParticipantsDesc').value;

   // Telepresent Shore-based Participants
  var tosp = document.getElementById('TelepresenceParticipantsDesc').value;

  // FOREIGN NATIONAL Participants
  var fn = document.getElementById('ForeignNationalParticipantsDesc').value;


  // Add all text areas together and set value to ID=allParticipants.  This will be placed in the Participants database field.
  if (document.getElementById('Participants').value != "") {
    j = mbariTxt.bold() + p;
  }

  if (document.getElementById('nonMBARIParticipantsDesc').value != "") {
    k = nonMbariTxt.bold() + non;
  }

  if (document.getElementById('ForeignNationalParticipantsDesc').value != "") {
    l = foreignNationalTxt.bold() + fn;
  }

  if (document.getElementById('TelepresenceParticipantsDesc').value != "") {
    tpr = telepresenceParticipantsTxt.bold() + tosp;
  }



  // Add them all up, grab the string, strip out the html characters and set value
  document.getElementById('allParticipants').value = j + k + l + tpr;
  var ap = document.getElementById('allParticipants').value;
  var participantsStr = strip_tags(ap);
  document.getElementById('allParticipants_cleanHTML').value = participantsStr;
}


// Add Planned Track Extent info to Track database
function updatePlannedTrackExtent(text) {
  plannedTrackExtent = text;
  joinPlannedTrackInfo();
}

// Add Planned Track Restricted info to Track database
function updatePlannedTrackRestricted(text) {
  plannedTrackRestricted = text;
  joinPlannedTrackInfo();
}

// Add MBNMS Permit Check verification to Permits database
function updateMBNMSPermitCheck(text) {
  mbnmsPermitCheck = text;
  joinPlannedTrackInfo();
}

// Add INFO WHETHER SAMPLES WILL BE TAKEN ON CRUISE
 function updateSamplePermitCheck(text) {
//   samplePermitCheck = text;
//   joinPlannedTrackInfo();
 }


// Add SMPA Permit Check verification to Permits database
function updateSMPAPermitCheck(text) {
  smpaPermitCheck = text;
  joinPlannedTrackInfo();
}

// Add SMPA Sample Collection Permit Check verification to Permits database
function updateSMPACollectionPermitCheck(text) {
//   smpaCollectionPermitCheck = text;
//   joinPlannedTrackInfo();
 }


// Add California Department of Fish and Wildlife Permit Check verification to Permits database
function updateCDFWPermitCheck(text) {
  cdfwPermitCheck = text;
  joinPlannedTrackInfo();
}




// Join all Equipment chosen (vehicles and winches) into one hidden input section
function joinPlannedTrackInfo() {
  var s = "";
  var t = "";
  var u = "";
  var v = "";
  var v0 = "";
  var v1 = "";
  var v2 = "";
  var v3 = "";
  var w = "";
  var plannedTxt = "Track Info: \n";
  var extentTxt = "\nTrack Outside Monterey Bay: ";
  var restrictedTxt = "\nTrack In Restricted Areas: ";
  var mbnmsPermitChk = "\nMBNMS Permit Compliance: ";
  //var samplePermitChk = "\nPlanned Sample Collection: ";
  var cdfwPermitChk = "\nCDFW Permit Compliance: ";
  //var smpaPermitChk = "\nTrack in SMP Area: ";
  var smpaCollectionPermitChk = "\nSMPA Collection Permit Compliance: ";

  var mbnmsPermitTxt = "\nPermit Comments: ";

  //Grab the string values of ALL Planned Track Info areas

  // Planned Track Desc
  var pt = document.getElementById('PlannedTrackDesc').value;

  // DOES PLANNED TRACK EXTEND BEYOND MONTEREY BAY AND INVOLVE UNDERWATER OPERATIONS?
  var te = document.getElementById('planned-track-extent').value;

  // DOES PLANNED TRACK INCLUDE RESTRICTED AREAS?
  var tr = document.getElementById('planned-track-restricted').value;


  // PERMITTING

  // IS PLANNED TRACK COMPLIANT WITH MBARI MBNMS PERMIT?
  var mp = document.getElementById('mbnms-permit-check').value;

  // ANIMAL, PLANT, OR GEOLOGICAL SAMPLES TO BE COLLECTED ON CRUISE?
  //var spc = document.getElementById('sample-permit-check').value;

  // IS PLANNED TRACK COMPLIANT WITH State Marine Protected Area (SMPA) PERMIT?
  //var sm = document.getElementById('smpa-permit-check').value;

  // IS PLANNED TRACK COLLECTING SAMPLES IN SMPA?
  var smc = document.getElementById('smpa-collection-permit-check').value;
  //console.log("SMC = " + smc);

  // IS COLLECTING SAMPLES IN SMPA PLANNED TRACK, DO YOU HAVE CSFW PERMIT?
  var cd = document.getElementById('cdfw-permit-check').value;


  //ANY PERMIT COMMENTS OR DETAILS?
  var pd = document.getElementById('PermitDesc').value;

  // Add both text areas together and set value to ID=allParticipants.  This will be placed in the Participants database field.
  if (document.getElementById('PlannedTrackDesc').value != "") {
    s = pt;
    //s = plannedTxt.bold() + pt;
  } else {
    s = "Selected Waypoints from Database.";
    //s = plannedTxt.bold() + "Selected Waypoints from Database.";
  }

  if (document.getElementById('planned-track-extent').value != "") {
    t = extentTxt.bold() + te;
  }

  if (document.getElementById('planned-track-restricted').value != "") {
    u = restrictedTxt.bold() + tr;
  }


  // PERMITS
  if (document.getElementById('mbnms-permit-check').value != "") {
    v = mbnmsPermitChk.bold() + mp;
  }

  // New Permitting requests.  Going to make it easy and use sequential var numbers......Not pretty but .....
  /*  if (document.getElementById('sample-permit-check').value != "") {
    v0 = samplePermitChk.bold() + spc;
  }*/

  if (document.getElementById('cdfw-permit-check').value != "") {
    v1 = cdfwPermitChk.bold() + cd;
  }

/*  if (document.getElementById('smpa-permit-check').value != "") {
    v2 = smpaPermitChk.bold() + sm;
  }*/

  if (document.getElementById('smpa-collection-permit-check').value != "") {
    v3 = smpaCollectionPermitChk.bold() + smc;
  }

  // PERMIT TEXT
  if (document.getElementById('PermitDesc').value != "") {
    w = mbnmsPermitTxt.bold() + pd;
  }

  // Add them all up, grab the string, strip out the html characters and set value
  document.getElementById('allPlannedTrackInfo').value = s + t + u + v + v0 + v1 + v2 + v3 + w;
  var ptd = document.getElementById('allPlannedTrackInfo').value;
  var planTrackStr = strip_tags(ptd);
  //var planTrackStr = translateEntities(planTrackStr);
  document.getElementById('allPlannedTrackInfo_cleanHTML').value = planTrackStr;


  ////////////////////////////////// EMAIL ALIASING ////////////////////////////////////////
  // SPECIAL OPS CHECK -  Set up Aliasing of Special Operations in the Planned Track Description
  if ((document.getElementById('planned-track-extent').value == "YES") || (document.getElementById('planned-track-restricted').value == "YES")) {
    document.getElementById('specialOpsCheck').value = "YES";
  } else {
    document.getElementById('specialOpsCheck').value = "NO";
  }


  // PERMITS CHECK -- Mandy Allen has asked that SHE ALWAYS RECEIVE copies of the Precruise for Permits - no exceoptions.  Sending them all via the alias.
/*  if ((document.getElementById('mbnms-permit-check').value!="none") || (document.getElementById('sample-permit-check').value!="none") || (document.getElementById('smpa-permit-check').value!="none") || (document.getElementById('smpa-collection-permit-check').value!="none") || (document.getElementById('cdfw-permit-check').value!="none")) {*/
  if ((document.getElementById('mbnms-permit-check').value!="none") || (document.getElementById('smpa-collection-permit-check').value!="none") || (document.getElementById('cdfw-permit-check').value!="none")) {
    document.getElementById('permitsCheck').value = "YES";
    //console.log("PERMITS CHECK: " + document.getElementById('permitsCheck').value);
  } else {
    document.getElementById('permitsCheck').value = "NO";
    //console.log("PERMITS CHECK: " + document.getElementById('permitsCheck').value);
  }
} // End of joinPlannedTrackInfo()



// Configure Form to the Selected Ship.  This step must come before all others.
// Function ConfigForm(selectedShip)
function ConfigForm()
{
  //console.log("Changing ships.");
  //console.log("CONFIGFORM()");
  // update Equipment should back button have been chosen.
  updateRequestedEquip();
  updateRequestedWinch();
  checkBroadcastPlatform();
  updateRequestedPlatforms();

  var selectedShip = document.getElementById('ShipName').value;

  if (selectedShip == "dpkd" || "rcsn" || "prg") {
    //Sections to Display
    document.getElementById("fill_fields").style.display = "inline";
    document.getElementById("cruise_specific_info_section").style.display = "table-header-group";
    document.getElementById("cruise-purpose-section").style.display = "inline";
    document.getElementById("required-equipment-section").style.display = "inline";
    document.getElementById("equipmentRequested").style.display = "inline";
    document.getElementById("hazardous-materials-section").style.display = "inline";
    document.getElementById("winch-requirements-section").style.display = "inline";
    document.getElementById("itar-section").style.display = "inline";
    document.getElementById("mooring-section").style.display = "inline";
    document.getElementById("cruise-participants-section").style.display = "inline";
    document.getElementById("planned-track-section").style.display = "inline";
    document.getElementById("sanctuary-permitting-section").style.display = "inline";
    document.getElementById("chief-sci-agreement-section").style.display = "inline";
    document.getElementById("submit-button-section").style.display = "inline";
    document.getElementById("precruise-entry-link").style.display = "inline";

    //Links
    document.getElementById("diver_select_label").style.display = "inline";
    //document.getElementById("diver-cruise-planning-link").href = "https://mww.mbari.org/safety/Diving/MBARI_Dive_Manual_2013.pdf";
    document.getElementById("mbari-dive-manual-link").href = "https://mww.mbari.org/safety/Diving/MBARI_Dive_Manual_2013.pdf";
    document.getElementById("dive-plan-form-link").href = "https://mww.mbari.org/safety/Diving/MBARIdiveplan.doc";
  }

  if (selectedShip == "dpkd")
  {
    // Remove equip in array not used by Packard if switching requested ships.
    if (equip.indexOf('LRAUV')) {
      for (var i = equip.length; i >=0; i--) {
        if(equip[i] === 'LRAUV') {
          equip.splice(i, 1);
          document.getElementById('lrauv').checked = false;
        }
        if(equip[i] === 'AUV') {
          equip.splice(i, 1);
          document.getElementById('auv').checked = false;
        }
      }
    }

    // Remove winch in array not used by Packard if switching requested ships.
    if (winch.indexOf('PARAGON_WINCH') || ('DOWNRIGGER_WINCH')) {
      for (var i = winch.length; i >=0; i--) {
        if(winch[i] === 'PARAGON_WINCH') {
          winch.splice(i, 1);
          document.getElementById('paragonWinch').checked = false;
        }
        if(winch[i] === 'DOWNRIGGER_WINCH') {
          winch.splice(i, 1);
          document.getElementById('downriggerWinch').checked = false;
        }
      }
    }

    // Add Telepresence Section
    if (document.getElementById("telepresenceRequested").value == "none") {
      document.getElementById("TelepresenceParticipantsDesc").style.display = "none";
      document.getElementById("broadcastPlatformsRequested").style.display = "none";
      document.getElementById("broadcastTP").style.display = "none";
      document.getElementById("listTPPlatforms").style.display = "none";
    }
    else if (document.getElementById("telepresenceRequested").value == "YES") {
      document.getElementById("broadcastTP").style.display = "inline";
      document.getElementById("telepresence-participants-section").style.display = "inline";
      document.getElementById("telepresence_participants_statement").style.display = "inline";
      document.getElementById("TelepresenceParticipantsDesc").style.display = "inline";
      document.getElementById("BroadcastPlatformDesc").style.display = "hidden";
        if (document.getElementById("broadcast-publicly").value == "yes-broadcast") {
          //console.log("Public Broadcast Requested.");
          document.getElementById("listTPPlatforms").style.display = "inline";
          document.getElementById("BroadcastPlatformDesc").style.display = "inline";
          updateBroadcastPublicly();
          updateRequestedPlatforms();
          //ConfigForm();
        }
    }
      else if (document.getElementById("telepresenceRequested").value == "NO") {
      document.getElementById("telepresence-participants-section").style.display = "none";
      document.getElementById("telepresence_participants_statement").style.display = "none";
      document.getElementById("youTube-media-label").style.display = "none";
      document.getElementById("broadcastPlatformsRequested").style.display = "none";
      document.getElementById("broadcastTP").style.display = "none";
      document.getElementById("listTPPlatforms").style.display = "none";
    }


    // Sections To Display
    document.getElementById('starlink-warning').style.display = "inline";
    document.getElementById("DPDivePlanRequired").style.display = "inline";
    document.getElementById("DPTelepresencePlanRequired").style.display = "inline";
    document.getElementById("telepresence_select_label").style.display = "inline";
    document.getElementById("meal-requests-section").style.display = "inline";
    document.getElementById("isotopes").style.display = "inline";
    document.getElementById("isotope_select_label").style.display = "inline";
    document.getElementById("IsotopeDescDiv").style.display = "inline";
    document.getElementById("DP_Participants_statement").style.display = "inline";
    document.getElementById("DP_Non_MBARI_Participants_Statement").style.display = "inline";
    document.getElementById("Foreign_participants_statement").style.display = "inline";
    document.getElementById("DP_foreign_participants_statement").style.display = "inline";
    document.getElementById("plannedTrack").style.display = "inline";

    // Links
    document.getElementById("precruise-link").href = "https://mww.mbari.org/dmo/Pages/DPPlanning.html";
    document.getElementById("mbari-haz-mat-policy-link").href = "https://mww.mbari.org/dmo/Pages/ChemLabHazardousMaterial.html";

    // Equipment
    document.getElementById("rov").style.display = "inline";
    document.getElementById("rov-label").style.display = "inline";
    document.getElementById("mini-rov").style.display = "inline";
    document.getElementById("minirov-label").style.display = "inline";
    document.getElementById("shipboard-vans").style.display = "inline";
    document.getElementById("shipboardVan-label").style.display = "inline";
    document.getElementById("asv").style.display = "inline";
    document.getElementById("asv-label").style.display = "inline";
    document.getElementById("ctd").style.display = "inline";
    document.getElementById("ctd-label").style.display = "inline";
    document.getElementById("elevator").style.display = "inline";
    document.getElementById("elevator-label").style.display = "inline";
    document.getElementById("other-equipment").style.display = "inline";
    document.getElementById("otherEquip-label").style.display = "inline";

    //Winches
    document.getElementById("black-dynacon-winch").style.display = "inline";
    document.getElementById("black-dynacon-winch-label").style.display = "inline";
    document.getElementById("macArtney-mooring-winch").style.display = "inline";
    document.getElementById("macArtney-mooring-winch-label").style.display = "inline";
    document.getElementById("borehole-winch").style.display = "inline";
    document.getElementById("borehole-winch-label").style.display = "inline";
    document.getElementById("packard-dynacon-winch").style.display = "inline";
    document.getElementById("packard-dynacon-winch_label").style.display = "inline";
    document.getElementById("other-winch").style.display = "inline";
    document.getElementById("otherWinch-label").style.display = "inline";

    document.getElementById("RCDivePlanRequired").style.display = "none";
    document.getElementById("ParagonDivePlanRequired").style.display = "none";
    document.getElementById("RCTelepresencePlanRequired").style.display = "none";
    document.getElementById("paragonWinch").style.display = "none";
    document.getElementById("paragon-winch-label").style.display = "none";
    document.getElementById("downriggerWinch").style.display = "none";
    document.getElementById("downrigger-winch-label").style.display = "none";
    document.getElementById("lrauv").style.display = "none";
    document.getElementById("lrauv-label").style.display = "none";
    document.getElementById("auv").style.display = "none";
    document.getElementById("auv-label").style.display = "none";
  }
  else if (selectedShip == "rcsn")
  {
    // Remove equip in array not used by Rachel Carson if switching requested ships.
    if (equip.indexOf('LRAUV') || ('RAD_VAN') || ('CHEM_VAN')) {
      for (var i = equip.length; i >=0; i--) {
        if(equip[i] === 'LRAUV') {
          equip.splice(i, 1);
          document.getElementById('lrauv').checked = false;
        }
        if(equip[i] === 'RAD_VAN') {
          equip.splice(i, 1);
          document.getElementById('rad-van').value = "";
        }
        if(equip[i] === 'CHEM_VAN') {
          equip.splice(i, 1);
          document.getElementById('chem-van').value = "";
        }
      }
    }

    // Remove winch in array not used by Rachel Carson if switching requested ships.
    if (winch.indexOf('PARAGON_WINCH') || ('DOWNRIGGER_WINCH') || ('.322_CTD_WINCH') || ('.25_TOW_WIRE_WINCH') || ('SYNTHETIC_ROPE_WINCH')) {
      for (var i = winch.length; i >=0; i--) {
        if(winch[i] === 'PARAGON_WINCH') {
          winch.splice(i, 1);
          document.getElementById('paragonWinch').checked = false;
        }
        if(winch[i] === 'DOWNRIGGER_WINCH') {
          winch.splice(i, 1);
          document.getElementById('downriggerWinch').checked = false;
        }
        if(winch[i] === '.322_CTD_WINCH') {
          winch.splice(i, 1);
          document.getElementById('ctd-winch').value = "";
        }
        if(winch[i] === '.25_TOW_WIRE_WINCH') {
          winch.splice(i, 1);
          document.getElementById('tow-wire-winch').value = "";
        }
        if(winch[i] === 'SYNTHETIC_ROPE_WINCH') {
          winch.splice(i, 1);
          document.getElementById('synthetic-rope-winch').value = "";
        }
      }
    }

    // Add Telepresence Section
    if (document.getElementById("telepresenceRequested").value == "none") {
      document.getElementById("TelepresenceParticipantsDesc").style.display = "none";
      document.getElementById("broadcastPlatformsRequested").style.display = "none";
      document.getElementById("broadcastTP").style.display = "none";
      document.getElementById("listTPPlatforms").style.display = "none";
    }
    else if (document.getElementById("telepresenceRequested").value == "YES") {
      document.getElementById("broadcastTP").style.display = "inline";
      document.getElementById("telepresence-participants-section").style.display = "inline";
      document.getElementById("telepresence_participants_statement").style.display = "inline";
      document.getElementById("TelepresenceParticipantsDesc").style.display = "inline";
      document.getElementById("BroadcastPlatformDesc").style.display = "hidden";
      if (document.getElementById("broadcastTP").value == "yes-broadcast") {
          //console.log("Public Broadcast Requested.");
          document.getElementById("broadcastTP").style.display = "inline";
          document.getElementById("telepresence-participants-section").style.display = "inline";
          document.getElementById("telepresence_participants_statement").style.display = "inline";
          document.getElementById("listTPPlatforms").style.display = "inline";
          document.getElementById("BroadcastPlatformDesc").style.display = "inline";
          ConfigForm();
          updateRequestedPlatforms();
          updateBroadcastPublicly("YES");
      }
    }
      else if (document.getElementById("telepresenceRequested").value == "NO") {
      document.getElementById("telepresence-participants-section").style.display = "none";
      document.getElementById("telepresence_participants_statement").style.display = "none";
      document.getElementById("youTube-media-label").style.display = "none";
      document.getElementById("broadcastPlatformsRequested").style.display = "none";
      document.getElementById("broadcastTP").style.display = "none";
      document.getElementById("listTPPlatforms").style.display = "none";
    }

    // Sections To Display
    document.getElementById('starlink-warning').style.display = "inline";
    document.getElementById("RCDivePlanRequired").style.display = "inline";
    document.getElementById("RCTelepresencePlanRequired").style.display = "inline";
    document.getElementById("telepresence_select_label").style.display = "inline";
    document.getElementById("meal-requests-section").style.display = "none";
    document.getElementById("isotopes").style.display = "inline";
    document.getElementById("isotope_select_label").style.display = "inline";
    document.getElementById("IsotopeDescDiv").style.display = "inline";
    document.getElementById("plannedTrack").style.display = "inline";
    document.getElementById("Foreign_participants_statement").style.display = "inline";
    document.getElementById("DP_Participants_statement").style.display = "none";
    document.getElementById("DP_Non_MBARI_Participants_Statement").style.display = "none";
    document.getElementById("DP_foreign_participants_statement").style.display = "none";



    // Links
    document.getElementById("precruise-link").href = "https://mww.mbari.org/dmo/Pages/RCPlanning.html";
    document.getElementById("mbari-haz-mat-policy-link").href = "https://mww.mbari.org/dmo/Pages/ChemLabHazardousMaterial.html";


    // Equipment
    document.getElementById("lrauv").style.display = "none";
    document.getElementById("lrauv-label").style.display = "none";
    document.getElementById("rov").style.display = "inline";
    document.getElementById("rov-label").style.display = "inline";
    document.getElementById("mini-rov").style.display = "inline";
    document.getElementById("minirov-label").style.display = "inline";
    document.getElementById("shipboard-vans").style.display = "none";
    document.getElementById("shipboardVan-label").style.display = "none";
    document.getElementById("auv").style.display = "inline";
    document.getElementById("auv-label").style.display = "inline";
    document.getElementById("asv").style.display = "inline";
    document.getElementById("asv-label").style.display = "inline";
    document.getElementById("ctd").style.display = "inline";
    document.getElementById("ctd-label").style.display = "inline";
    document.getElementById("other-equipment").style.display = "inline";
    document.getElementById("otherEquip-label").style.display = "inline";
    document.getElementById("elevator").style.display = "inline";
    document.getElementById("elevator-label").style.display = "inline";
    document.getElementById("DPDivePlanRequired").style.display = "none";
    document.getElementById("ParagonDivePlanRequired").style.display = "none";
    document.getElementById("DPTelepresencePlanRequired").style.display = "none";

    //Winches
    document.getElementById("black-dynacon-winch").style.display = "inline";
    document.getElementById("black-dynacon-winch-label").style.display = "inline";
    document.getElementById("macArtney-mooring-winch").style.display = "inline";
    document.getElementById("macArtney-mooring-winch-label").style.display = "inline";
    document.getElementById("borehole-winch").style.display = "inline";
    document.getElementById("borehole-winch-label").style.display = "inline";
    document.getElementById("other-winch").style.display = "inline";
    document.getElementById("otherWinch-label").style.display = "inline";

    document.getElementById("packard-dynacon-winch").style.display = "none";
    document.getElementById("packard-dynacon-winch_label").style.display = "none";
    document.getElementById("paragonWinch").style.display = "none";
    document.getElementById("paragon-winch-label").style.display = "none";
    document.getElementById("downriggerWinch").style.display = "none";
    document.getElementById("downrigger-winch-label").style.display = "none";

  }
  else if (selectedShip == "prgn")
  {
    // Remove equip in array not used by Paragon if switching requested ships.
    if (equip.indexOf('ROV') || ('RAD_VAN') || ('CHEM_VAN') || ('MINI_ROV')) {
      for (var i = equip.length; i >=0; i--) {
        if(equip[i] === 'ROV') {
          equip.splice(i, 1);
          document.getElementById('rov').checked = false;
        }
        if(equip[i] === 'MINI_ROV') {
          equip.splice(i, 1);
          document.getElementById('mini-rov').checked = false;
        }
        if(equip[i] === 'RAD_VAN') {
          equip.splice(i, 1);
          document.getElementById('rad-van').value = "";
        }
        if(equip[i] === 'CHEM_VAN') {
          equip.splice(i, 1);
          document.getElementById('chem-van').value = "";
        }
      }
    }

    // Remove winch in array not used by Paragon if switching requested ships.
    if (winch.indexOf('BLACK_DYNACON_WINCH') || ('BOREHOLE_WINCH') || ('MACARTNEY_MOORING_WINCH') || ('.322_CTD_WINCH') || ('.25_TOW_WIRE_WINCH') || ('SYNTHETIC_ROPE_WINCH')) {
      for (var i = winch.length; i >=0; i--) {
        if(winch[i] === 'BLACK_DYNACON_WINCH') {
          winch.splice(i, 1);
          document.getElementById('black-dynacon-winch').checked = false;
        }
        if(winch[i] === 'BOREHOLE_WINCH') {
          winch.splice(i, 1);
          document.getElementById('borehole-winch').checked = false;
        }
        if(winch[i] === 'MACARTNEY_MOORING_WINCH') {
          winch.splice(i, 1);
          document.getElementById('macArtney-mooring-winch').checked = false;
        }
        if(winch[i] === '.322_CTD_WINCH') {
          winch.splice(i, 1);
          document.getElementById('ctd-winch').value = "";
        }
        if(winch[i] === '.25_TOW_WIRE_WINCH') {
          winch.splice(i, 1);
          document.getElementById('tow-wire-winch').value = "";
        }
        if(winch[i] === 'SYNTHETIC_ROPE_WINCH') {
          winch.splice(i, 1);
          document.getElementById('synthetic-rope-winch').value = "";
        }
      }
    }

    // Sections To Display
    document.getElementById("ParagonDivePlanRequired").style.display = "inline";
    document.getElementById("telepresence_select_label").style.display = "none";
    document.getElementById("meal-requests-section").style.display = "none";
    document.getElementById("DP_Participants_statement").style.display = "none";
    document.getElementById("DP_Non_MBARI_Participants_Statement").style.display = "none";
    document.getElementById("Foreign_participants_statement").style.display = "inline";
    document.getElementById("DP_foreign_participants_statement").style.display = "none";

    // Links
    document.getElementById("precruise-link").href = "https://mww.mbari.org/dmo/Pages/Paragon.html";
    document.getElementById("mbari-haz-mat-policy-link").href = "https://mww.mbari.org/dmo/Pages/ChemLabHazardousMaterial.html";

    // Equipment
    document.getElementById("auv").style.display = "inline";
    document.getElementById("auv-label").style.display = "inline";
    document.getElementById("asv").style.display = "inline";
    document.getElementById("asv-label").style.display = "inline";
    document.getElementById("ctd").style.display = "inline";
    document.getElementById("ctd-label").style.display = "inline";
    document.getElementById("lrauv").style.display = "inline";
    document.getElementById("lrauv-label").style.display = "inline";
    document.getElementById("other-equipment").style.display = "inline";
    document.getElementById("otherEquip-label").style.display = "inline";


    document.getElementById('starlink-warning').style.display = "none";
    document.getElementById("rov").style.display = "none";
    document.getElementById("rov-label").style.display = "none";
    document.getElementById("mini-rov").style.display = "none";
    document.getElementById("minirov-label").style.display = "none";
    document.getElementById("shipboard-vans").style.display = "none";
    document.getElementById("shipboardVan-label").style.display = "none";
    document.getElementById("isotopes").style.display = "none";
    document.getElementById("isotope_select_label").style.display = "none";
    document.getElementById("IsotopeDescDiv").style.display = "none";
    document.getElementById("elevator").style.display = "none";
    document.getElementById("elevator-label").style.display = "none";

    //Winches
    document.getElementById("other-winch").style.display = "inline";
    document.getElementById("otherWinch-label").style.display = "inline";
    document.getElementById("paragonWinch").style.display = "inline";
    document.getElementById("paragon-winch-label").style.display = "inline";
    document.getElementById("downriggerWinch").style.display = "inline";
    document.getElementById("downrigger-winch-label").style.display = "inline";

    document.getElementById("black-dynacon-winch").style.display = "none";
    document.getElementById("black-dynacon-winch-label").style.display = "none";
    document.getElementById("macArtney-mooring-winch").style.display = "none";
    document.getElementById("macArtney-mooring-winch-label").style.display = "none";
    document.getElementById("borehole-winch").style.display = "none";
    document.getElementById("borehole-winch-label").style.display = "none";
    document.getElementById("packard-dynacon-winch").style.display = "none";
    document.getElementById("packard-dynacon-winch_label").style.display = "none";
    document.getElementById("plannedTrack").style.display = "none";
    document.getElementById("RCDivePlanRequired").style.display = "none";
    document.getElementById("DPDivePlanRequired").style.display = "none";
    document.getElementById("RCTelepresencePlanRequired").style.display = "none";
    document.getElementById("DPTelepresencePlanRequired").style.display = "none";


    //Planned Track Description
    // document.getElementById("planned-track-extent").style.display = "none";
    // document.getElementById("planned-track-restricted").style.display = "none";
  }
  else if (selectedShip == "unknownShip")
  {
    //clearLinks();
    document.getElementById("fill_fields").style.display = "none";
    document.getElementById("cruise_specific_info_section").style.display = "none";
    document.getElementById("cruise-purpose-section").style.display = "none";
    document.getElementById("required-equipment-section").style.display = "none";
    document.getElementById("equipmentRequested").style.display = "none";
    document.getElementById("hazardous-materials-section").style.display = "none";
    document.getElementById("winch-requirements-section").style.display = "none";
    document.getElementById("itar-section").style.display = "none";
    document.getElementById("mooring-section").style.display = "none";
    document.getElementById("cruise-participants-section").style.display = "none";
    document.getElementById("meal-requests-section").style.display = "none";
    document.getElementById("planned-track-section").style.display = "none";
    document.getElementById("sanctuary-permitting-section").style.display = "none";
    document.getElementById("chief-sci-agreement-section").style.display = "none";
    document.getElementById("submit-button-section").style.display = "none";
    document.getElementById("precruise-entry-link").style.display = "none";
    document.getElementById("RCDivePlanRequired").style.display = "none";
    document.getElementById("DPDivePlanRequired").style.display = "none";
    document.getElementById("ParagonDivePlanRequired").style.display = "none";
    document.getElementById("DPTelepresencePlanRequired").style.display = "none";
    document.getElementById("RCTelepresencePlanRequired").style.display = "none";
    document.getElementById("telepresence-participants-section").style.display = "none";
    return (false);
  }
} // End of ConfigForm();

function restartFlash() {
  var flashElements = document.getElementsByClassName('flash');
  for (var i = 0; i < flashElements.length; i++) {
    var el = flashElements[i];
    
    // 1. Remove the animation
    el.style.animation = 'none';
    
    // 2. Trigger "reflow" (The Magic Trick)
    // Accessing offsetHeight forces the browser to refresh the element's state
    void el.offsetHeight; 
    
    // 3. Re-enable the animation
    el.style.animation = null; 
  }
}


// Hitting the Back Button  ////
// Set Location to return to on the screen one the back button has been hit.
function scrollWin() {
    window.scrollTo(0, 0);
}

// Reload the page prior to the first Submit
window.onload = function (e) {
  ConfigForm();
  setTimeout(scrollWin, 10);
}
</script>


<%
open_database($dsn);    # Creates $Conn object as a global variable

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



<!-- SET UP HTML FORM -->
<!-- <form method="POST" action="<%=$this_script%>" autocomplete="on" onsubmit="return Field_Validator(this)" name="Field" id="Form"> -->
<form method="POST" action="<%=$this_script%>" onsubmit="return Field_Validator(this)" name="Field" id="Form" accept-charset="utf-8">
<br>
<font id="fill_fields" style="display:none;" face="Arial,Helvetica" size="3" color="brown"><strong>All fields below must be completed.</strong></font>
<br>
<table border="1" style="vertical-align:top">
  <tr>
    <td>
      <input type="hidden" name="step" value="2">

      <!-- Begin Top Table -->
      <table border="0" cellpadding="5" cellspacing="0">

              <!-- Section 1.0 -->
              <!-- Set Ship Name - Must do this first -->
              <thead>
                <tr align="left">
                  <th colspan="1" width="250"><strong><font face="Arial,Helvetica">Ship name:</font></strong></th>
                  <th colspan="1">
                    <div class="shipSelect">
                      <select id="ShipName" name="ShipName" size="1" onChange="return ShipSelect(this)">
                        <option value="unknownShip">Please Select A Ship</option>
                        <option value="dpkd">R/V David Packard</option>
                        <option value="prgn">R/V Paragon</option>
                        <option value="rcsn">R/V Rachel Carson</option>
                      </select>
                    </div></th>
                </tr>
              </thead>

              <tbody id="cruise_specific_info_section" style="display: none;">

                <!-- Cruise Departure Date -->
                <tr id="cruise-departure-date">
                  <td valign="top"><strong><font face="Arial,Helvetica">Cruise departure date - time</font></strong></td>
                  <td valign="top"><font face="Arial,Helvetica">
                    <select name="dMonth" size="1"><%= $month_option_list%></select>
                    <select name="dDay" size="1"><%= $day_option_list%></select>
                    <select name="dYear" size="1"><%= $year_option_list%></select>
                    <input type="text" name="dTime" size="4" value="0700" maxlength="4" style="width: 40px;"><font size="-1"> (Local Moss Landing time)</font>
                  </td>
                </tr>

                <!-- Cruise Arrival Date -->
                <tr id="cruise-arrival-date">
                  <td valign="top"><strong><font face="Arial,Helvetica">Cruise arrival date - time</font></strong></td>
                  <td valign="top"><font face="Arial,Helvetica">
                    <select name="aMonth" size="1"><%= $month_option_list%></select>
                    <select name="aDay" size="1"><%= $day_option_list%></select>
                    <select name="aYear" size="1"><%= $year_option_list%></select>
                    <input type="text" name="aTime" size="4" value="1630" maxlength="4" style="width: 40px;"><font size="-1"> (Local Moss Landing time)</font>
                  </td>
                </tr>

                <!-- Principal Investigator Entry: Lookup to Database -->
                <tr id="princial-investigator">
                  <td valign="top"><strong><font face="Arial,Helvetica">Principal investigator</font></strong></td>
                  <td valign="top">
                    <select name="ExpdPrincipalInvestigator" size="1">
                      <option value=""></option>
                      <% while ( !$RSpm->EOF ) {
                        $fullname = $RSpm->Fields('FirstName')->value . " " . $RSpm->Fields('LastName')->value;
                        %><option value="<%= $fullname%>"><%= $fullname%></option>
                        <%
                        $RSpm->MoveNext;
                      } %>
                    </select>
                    Project PI awarded ship time
                  </td>
                </tr>

                <!-- Chief Scientist Entry: Lookup to Database -->
                <tr id="chief-scientist">
                  <td valign="top"><strong><font face="Arial,Helvetica">Chief scientist</font></strong></td>
                  <td valign="top">
                    <select name="ExpdChiefScientist" size="1">
                      <option value=""></option>
                      <% while ( !$RScs->EOF ) {
                        $fullname = $RScs->Fields('FirstName')->value . " " . $RScs->Fields('LastName')->value;
                        %><option value="<%= $fullname%>"><%= $fullname%></option>
                        <%
                        $RScs->MoveNext;
                      } %>
                    </select>
                    <a href="person.asp">Add/edit person to list</a>
                  </td>
                </tr>

                <!-- Input MBARI Project Number for reference -->
                <tr id="project-number">
                  <td valign="top"><strong><font face="Arial,Helvetica">MBARI project number</font></strong></td>
                  <td><input type="text" name="ProjNum" size="10" value="" maxlength="10" style="width: 85px;"></td>
                </tr>

                <!-- Setup New Precruise link based on ship choice -->
                <tr id="precruise-entry-link">
                  <td><a id="precruise-link" target="_blank"><font size=2 face="Arial,Helvetica">Ship precruise information</a><br><br><br></td>
                </tr>
              </tbody>
            </table>



          <!-- Section 2.0 -->
          <!-- Setup Cruise Purpose Section -->
          <div class="padding" id="cruise-purpose-section" style="display:none;">
            <hr>
            <p class="padding"><strong><font face="Arial,Helvetica" size=5 color=green>I. Cruise purpose</font></strong><br><br>
              <font face="Arial,Helvetica" size="3" color="black"><strong>Enter objective for the expedition</strong></font><font face="Arial,Helvetica" size="-1" color="brown"> (required).</font><br><br>
            <textarea class="textfield" input="text" rows="10" id="Purpose" name="Purpose" value="" cols="160" placeholder="MAXIMUM Entry Length: 8000 character limit." maxlength="8000" onKeyDown="textCounter(this.form.Purpose,this.form.countDisplay);" onKeyUp="textCounter(this.form.Purpose,this.form.countDisplay);" wrap></textarea><br>
            <!-- <input readonly type="text" name="countDisplay" color="brown" size="4" maxlength="5" value="8000"> Characters Remaining         -->
            <input readonly type="text" name="countDisplay" color="brown" size="4" maxlength="5" value="8000" style="font-weight: bold; color: brown; width: 38px;"> characters remaining

            <script Language="Javascript">
                var maxLength = 8000;
              function textCounter(textField, showCountField) {
                  if (textField.value.length > maxLength) {
                     textField.value = textField.value.substring(0, maxLength);
                  }
                  else {
                  showCountField.value = maxLength - textField.value.length;
                  }
              }
            </script>


            <!--  <textarea rows="10" name="Purpose" value="" id="Purpose" cols="160" wrap></textarea>-->
            </p>
            <br><br><br>

            <!-- <p class="padding" id="diver" style="display:none;"><font class="nudge" face="Arial,Helvetica" size="3"><strong>Will there be diving operations?</strong><input type="checkbox" id="diverRequested" name="diverRequested" value="" onclick="updateDiverRequest(this.value)"> -->

            <p class="padding">
              <label id="diver_select_label" for="diverRequested"><font class="nudge" face="Arial,Helvetica" size="3" color="black"><strong>Will there be diving operations?</strong>
                <select style="margin-left:5px" id="diverRequested" name="diverRequested" value="" onchange='updateDiverRequest(this.value)'>
                  <option value="NO" id="no-diver">No</option>
                  <option value="YES" id="yes-diver">Yes</option>
                </select>
              </label>

            <br><a id="mbari-dive-manual-link" target="_blank"><font size=2 face="Arial,Helvetica">Dive manual</font></a>
            <br><br>

            <font face="Arial,Helvetica" size="-1" color="brown">If yes, please complete this form: </font><a class="padding" id="dive-plan-form-link" target="_blank"><font size=2 face="Arial,Helvetica">Dive plan form</font></a></p>

            <p class="padding" id="DPDivePlanRequired"><font face="Arial,Helvetica" size="-1" color="brown"><strong>NOTE:</strong> Dive plans for R/V <i>David Packard</i> are due <strong>3 weeks</strong> in advance of cruise.</font></p>
            <p class="padding" id="ParagonDivePlanRequired"><font face="Arial,Helvetica" size="-1" color="brown"><strong>NOTE:</strong> Dive plans for R/V <i>Paragon</i> are due <strong>prior</strong> to departure.</font><br></p>
            <p class="padding" id="RCDivePlanRequired"><font face="Arial,Helvetica" size="-1" color="brown"><strong>NOTE:</strong> Dive plans for R/V <i>Rachel Carson</i> are due <strong>2 weeks</strong> in advance of cruise.</font></p>


            <!-- Telepresence Request Section -->
            <p class="padding">
              <!-- <div class="flash"> -->
              <label id="telepresence_select_label" for="telepresenceRequested"><font face="Arial,Helvetica" size="3" color="black"><a name="telepresencelink"></a>
                <strong><br><br><br><br><span class="flash">Is telepresence required for this expedition?</span></strong>
                <select style="margin-left:5px" id="telepresenceRequested" name="telepresenceRequested" value="" onchange='updateTelepresenceRequest(this.value)'>
                  <option value="NO" id="no-telepresence">No</option>
                  <option value="YES" id="yes-telepresence">Yes</option>
                </select>
              </label>
              <!--</div> -->
            <p class="padding" id="DPTelepresencePlanRequired"><font face="Arial,Helvetica" size="-1" color="brown"><strong>NOTE:</strong> Telepresence may ONLY be requested for R/V <i>David Packard</i> and R/V <i>Rachel Carson</i>.</font><br><br></p>
            <p class="padding" id="RCTelepresencePlanRequired"><font face="Arial,Helvetica" size="-1" color="brown"><strong>NOTE:</strong> Telepresence may ONLY be requested for R/V <i>David Packard</i> and R/V <i>Rachel Carson</i>.</font>
            <br><br>
            </p>

            <input id="allCruisePurposeInfo" type="hidden" style="width: auto" name="allCruisePurposeInfo" value="" />
            <input id="allCruisePurposeInfo_cleanHTML" type="hidden" style="width: auto" name="allCruisePurposeInfo_cleanHTML" value="" />
          </div>


            <!-- Section 3.0 -->
            <!-- Setup Required Equipment Description Section -->
            <div class="padding" id="required-equipment-section" style="display:none;">
              <hr>
              <p class="padding"><strong><font face="Arial,Helvetica" size=5 color=green>II. Required equipment description</font></strong><br><br>

                <!-- Section 3.0.1 -->
                <!-- Strong Encouragement to all Scientists to coordinate a FULL STARLINK TEST (BEFORE) departure with IS and ship's crew if they are bringing any equipment onboard that depends on the network. -->
                <div id="starlink-warning" style="display:none;">
                  <span class="flash" style="display:inline-block; font-family:Arial,Helvetica; font-size: 12pt;">
                  <!-- <span class="flash" face="Arial,Helvetica" size="3"> -->
                    <font color="red"><strong>Attention Scientists:</strong></font> 
                    If you plan to bring equipment requiring network access, please contact 
                    <strong><a href="mailto:lmurphy@mbari.org" style="color: blue;">Leland Murphy (DMO)</a></strong> and 
                    <strong><a href="mailto:asisfontes@mbari.org" style="color: blue;">Adriana Sisfontes (IS)</a></strong> 
                    to schedule a full Starlink test. <br>This must be coordinated with the ship&apos;s crew and ROV pilots at least <strong>48 hours prior to departure.</strong>
                  </span>
                  <br><br><br>
              </div>

                <!-- Section 3.1 -->
                <!-- Declaration all Required Equipment Section -->
                <font face="Arial,Helvetica" size="3" color="black"><strong>Check all that apply</strong>.</font><br><br>
                <fieldset class="equipmentRequested" id="equipmentRequested" style="display:none;" name="equipmentRequested">
                  <label id="asv-label" for="asv" face="Arial,Helvetica" size="-1" color="black"><strong>ASV</strong><input type="checkbox" id="asv" name="equipSelect[]" value="ASV" onclick='updateRequestedEquip(this.value)'></label>
                  <label id="auv-label" for="auv" face="Arial,Helvetica" size="-1" color="black"><strong>AUV</strong><input type="checkbox" id="auv" name="equipSelect[]" value="AUV" onclick='updateRequestedEquip(this.value)'></label>
                  <label id="ctd-label" for="ctd" face="Arial,Helvetica" size="-1" color="black"><strong>CTD</strong><input type="checkbox" id="ctd" name="equipSelect[]" value="CTD" onclick='updateRequestedEquip(this.value)'></label>
                  <label id="elevator-label" for="elevator" face="Arial,Helvetica" size="-1" color="black"><strong>Elevator</strong><input type="checkbox" id="elevator" name="equipSelect[]" value="ELEVATOR" onclick='updateRequestedEquip(this.value)'></label>
                  <label id="lrauv-label" for="lrauv" face="Arial,Helvetica" size="-1" color="black"><strong>LRAUV</strong><input type="checkbox" id ="lrauv" name="equipSelect[]" value="LRAUV" onclick='updateRequestedEquip(this.value)'></label>
                  <label id="minirov-label" for="mini-rov" face="Arial,Helvetica" size="-1" color="black"><strong>MiniROV</strong><input type="checkbox" id="mini-rov" name="equipSelect[]" value="MINI_ROV" onclick='updateRequestedEquip(this.value)'></label>
                  <label id="rov-label" for="rov" face="Arial,Helvetica" size="-1" color="black"><strong>ROV</strong><input type="checkbox" id="rov" name="equipSelect[]" value="ROV" onclick='updateRequestedEquip(this.value)'></label>
                  <label id="shipboardVan-label" for="shipboard-vans" face="Arial,Helvetica" size="-1" color="black"><strong>Shipboard van</strong>
                    <select style="margin-left:5px" id="shipboard-vans" name="equipSelect[]" type="text" value="SHIPBOARDVANS" onchange='updateRequestedEquip(this.value)'>
                      <option id="no-van"></option>
                      <option value="RAD_VAN" id="rad-van" type="text" name="rad-van">Radiation Van</option>
                      <option value="CHEM_VAN" id="chem-van" type="text" name="chem-van">Paull Chem Van</option>
                    </select></label>
                  <label id="otherEquip-label" for="other-equipment" face="Arial,Helvetica" size="-1" color="black"><strong>Other</strong><input type="checkbox" id="other-equipment" name="equipSelect[]" value="OTHER_EQUIP" onclick='enableDisableEquip()'></label>
                  <br>
                  <font size=2 color="brown"> *If <strong>"Other"</strong>, please explain in text area below.</font>
                </fieldset>

                <!-- ALIASING PLACEHOLDERS -->
                <input id="allEquip" type="hidden" name="allEquip" value="" />
                <input id="allEquip_cleanHTML" type="hidden" name="allEquip_cleanHTML" value="" />
                <input id="waveG" type="hidden" name="waveG" value="" />
                <input id="ctdCheck" type="hidden" name="ctdCheck" value="" />
                <!--<input id="broadcastTP-check" type="hidden" name="broadcastTP-check" value="" />-->
              </p>

              <p class="padding"><br>
                <font face="Arial,Helvetica" size="-1" color=brown>If the Wave Glider (ASV) will be launched or recovered, please indicate in the box below. <br>&nbsp;- Wave Glider cruise plans are due three (3) weeks prior to activity.<br><br>
                  <strong> NOTE:</strong>  Planned ASV operations should establish a <a target="_blank" href="https://www.pacificarea.uscg.mil/Our-Organization/District-11/Prevention-Division/LnmRequest/">Local Notice to Mariners (LNM) Request</a> <b>14 days in advance</b> if not under a blanket LNM.
                </font><br><br>

                <textarea class="textfield" input="text" rows="10" id="EquipmentDesc" name="EquipmentDesc" value="EquipmentDesc" cols="160" placeholder="Please enter equipment details here. If none, please leave empty.
(MAXIMUM Entry Length: 8000 character limit)." maxlength="8000" onclick='checkEquipOther()' onKeyDown="textCounter1(this.form.EquipmentDesc,this.form.countDisplay1)" onKeyUp="textCounter1(this.form.EquipmentDesc,this.form.countDisplay1)" wrap></textarea><br>

                <!--<textarea class="textfield" input="text" rows="10" id="EquipmentDesc" name="EquipmentDesc" value="" cols="160" placeholder="Please enter equipment details here. If none, please leave empty." maxlength="8000" onclick='checkEquipOther()' onKeyDown="textCounter1(this.form.EquipmentDesc,this.form.countDisplay1);" onKeyUp="textCounter1(this.form.EquipmentDesc,this.form.countDisplay1);" wrap>
                </textarea>
                <br>-->
                <input readonly type="text" name="countDisplay1" color="brown" size="4" maxlength="5" value="8000" style="font-weight: bold; color: brown; width: 38px;"> characters remaining

                <script Language="Javascript">
                var maxLength = 8000;
                function textCounter1(textField, showCountField) {
                    if (textField.value.length > maxLength) {
                        textField.value = textField.value.substring(0, maxLength);
                    }
                    else {
                        showCountField.value = maxLength - textField.value.length;
                    }
                }
                </script>


                <br><br></p>
            </div>



          <!-- Section 3.2 -->
          <!-- Winch Requirements -->
          <div id="winch-requirements-section" style="display:none;">
          <hr>
          <p class="padding"><strong><font face="Arial,Helvetica" size=5 color="green">III. Winch requirements</font></strong><br><br>
          <font face="Arial,Helvetica" size="3" color="black"><strong>Check all that apply</strong>.</font><br><br>
          <fieldset class="WinchSelected" id="WinchSelected" name="WinchSelected" value="test">
            <label id="paragon-winch-label" for="paragonWinch" face="Arial,Helvetica" size="-1" color="black"><strong>Paragon Winch</strong><input type="checkbox" id="paragonWinch" name="winchSelect[]" value="PARAGON_WINCH" onclick='updateRequestedWinch(this.value)'></label>
            <label id="downrigger-winch-label" for="downriggerWinch" face="Arial,Helvetica" size="-1" color="black"><strong>Downrigger Winch</strong><input type="checkbox" id="downriggerWinch" name="winchSelect[]" value="DOWNRIGGER_WINCH" onclick='updateRequestedWinch(this.value)'></label>
            <label id="black-dynacon-winch-label" for="black-dynacon-winch" face="Arial,Helvetica" size="-1" color="black"><strong>Black Dynacon</strong><input type="checkbox" id="black-dynacon-winch" name="winchSelect[]" value="BLACK_DYNACON_WINCH" onclick='updateRequestedWinch(this.value)'></label>
            <label id="borehole-winch-label" for="borehole-winch" face="Arial,Helvetica" size="-1" color="black"><strong>Borehole</strong><input type="checkbox" id="borehole-winch" name="winchSelect[]" value="BOREHOLE_WINCH" onclick='updateRequestedWinch(this.value)'></label>
            <label id="macArtney-mooring-winch-label" for="macArtney-mooring-winch" face="Arial,Helvetica" size="-1" color="black"><strong>MacArtney</strong><input type="checkbox" id="macArtney-mooring-winch" name="winchSelect[]" value="MACARTNEY_MOORING_WINCH" onclick='updateRequestedWinch(this.value)'></label>
            <label id="packard-dynacon-winch_label" for="packard-dynacon-winch" face="Arial,Helvetica" size="-1" color="black"><strong><i>Packard</i> Dynacon</strong>
              <select style="margin-left:5px" id="packard-dynacon-winch" name="winchSelect[]" value="PACKARD_DYNACON_WINCH" onchange='updateRequestedWinch(this.value)'>
                <option value=" "></option>
                <option value=".322_CTD_WINCH" id="ctd-winch" type="text" name="ctd-winch">.322 CTD Winch</option>
                <option value=".25_TOW_WIRE_WINCH" id="tow-wire-winch" type="text" name="tow-wire-winch">.25 Tow Wire</option>
                <option value="SYNTHETIC_ROPE_WINCH" id="synthetic-rope-winch" type="text" name="synthetic-rope-winch">High Strength Synthetic Rope</option>
              </select></label>
             <label id="otherWinch-label" for="other-winch" face="Arial,Helvetica" size="-1" color="black"><strong>Other</strong><input type="checkbox" id="other-winch" name="winchSelect[]" value="OTHER_WINCH" onclick='enableDisableWinch()'></label>
             <br>
             <font size=2 color="brown"> *If <strong>"Other"</strong>, please explain in text area below.
          </fieldset>
          </p>

          <p class="padding">
          <br><font face="Arial,Helvetica" size="-1" color=brown>Please describe any additional winch requests or requirements in the box below:</font><br><br>
          <textarea class="textfield" input="text" rows="10" id="WinchDesc" name="WinchDesc" value="WinchDesc" cols="160" placeholder="Please enter winch details here. If none, please leave empty." onclick='checkWinchOther()' wrap></textarea><br>
          <br><br><br><br>
            <input id="winchCheck" type="hidden" name="winchCheck" value="" />
          </p>
          <hr>
          </div>



          <!-- Section 3.3 -->
          <!-- Declaration of Isotopes Section -->
          <div id="hazardous-materials-section" style="display:none;">
          <p class="padding">
            <strong><font face="Arial,Helvetica" size=5 color="green">IV. Hazardous materials declaration and description</font></strong><br><br>

            <label id="isotope_select_label" for="isotopes"><font face="Arial,Helvetica" size="3" color="black"><strong>Do you plan to use any enriched stable or radioactive isotopes?</strong><select style="margin-left:5px" id="isotopes" name="isotopes" value="" onchange="updateIsotopeDeclaration(this.value)">
              <option value="no-isotope">No</option>
              <option value="stable-isotope" id="stable-isotope">Yes - Stable Isotopes</option>
              <option value="radio-isotope" id="radio-isotope">Yes - Radioisotopes</option>
              <option value="radio-stable-isotopes" id="radio-stable-isotope">Yes - Both Stable & Radioisotopes
            </select><br><br>
            <font face="Arial,Helvetica" size="-1" color=brown>If yes, please provide isotope list in the box below:</font><br></label>
          </p>
            <div class="padding" id="IsotopeDescDiv"><textarea rows="10" id="IsotopeDesc" name="IsotopeDesc" value="IsotopeDesc"cols="160" wrap></textarea>
            <br><br><br></div>



          <!-- Section 3.4 -->
          <!-- Declaration of Hazardous Materials Section -->
          <p class="padding">
            <label id="hazmat_select_label" for="hazMat"><font face="Arial,Helvetica" size="3" color="black"><strong>Are you bringing any hazardous materials onboard?</strong><select style="margin-left:5px" id="hazMat" name="hazMat" value="isotopeprecruise@listserver.mbari.org" onchange='updateHazMatDeclaration(this.value)'>
              <option value="no-hazMat" id="no-hazMat">No</option>
              <option value="yes-hazMat" id="yes-hazMat">Yes</option>
            </select><br><br>
            <font face="Arial,Helvetica" size="-1" color=brown>If yes, please provide a complete list of hazardous materials in the box below:</font><br></label>
          </p>
            <div class="padding" id="HazMatDiv"><textarea rows="10" name="HazMatDesc" id="HazMatDesc" value="HazMatDesc" cols="160" wrap></textarea>
            <br><br></div>

            <!-- Acceptance of MBARI Isotope and Haz Mat Policy Section -->
            <p class="padding">
            <font face="Arial,Helvetica" size="2" color=red>PLEASE READ:</font>
            <a id="mbari-haz-mat-policy-link" name="mbari-haz-mat-policy-link" target="_blank"><font size=2 color="blue" face="Arial,Helvetica" style="margin-left:20px">MBARI hazardous, radioactive and stable isotope materials policies</font></a>
            <br><br>
            <font face="Arial,Helvetica" size="3" color="black"><strong>I have read and agree to this policy</strong><input type="checkbox" id="isotopePolicyAgree" name="isotopePolicyAgree" value="isotopePolicyAgree"><br></font>
            <br><br><br>

            <input id="hazIsoCheck" type="hidden" name="hazIsoCheck" value="" />
          </p>
          <hr>
        </div>



          <!-- Section 3.5 -->
          <!-- ITAR Declaration Section -->
          <div id="itar-section" style="display:none;">
          <p class="padding">
            <strong><font face="Arial,Helvetica" size=5 color="green">V. International Traffic in Arms Regulations (ITAR) declaration</font></strong><br><br><br>
            <font face="Arial,Helvetica" size="3" color="black"><strong>Is your equipment ITAR-controlled and going beyond <a target="_blank" href="http://mww.mbari.org/expdlog/expd/log/maps/12nm_range_cropped.pdf">12 nautical miles</a> from shore?</strong><select style="margin-left:5px" id="itar" name="itar" value="mandy@mbari.org" onchange="updateItarDeclaration(this.value)">
              <option value="no-itar">No</option>
              <option value="yes-itar">Yes</option>
            </select><br>
            <!-- <font face="Arial,Helvetica" size="-1" color="brown"><strong>NOTE:</strong> Export Administration Regulatons (EAR) items are not controlled in international waters. <br>-->
            <br><br>
            <font face="Arial,Helvetica" size="-1" color=brown>If yes, please list <a target="_blank" href="http://mww.mbari.org/exportregs/items.htm">ITAR-controlled items</a> in box below and include <strong>serial number</strong> and <strong>part number</strong>:
            <br><br>
            <textarea rows="10" name="ItarDesc" id="ItarDesc" value="ItarDesc" cols="160" wrap></textarea>
            <br><a id="export-controlled-regulations-list" target="_blank" href="http://mww.mbari.org/exportregs/index.htm"><font size=2 color="blue" face="Arial,Helvetica">Export controlled regulations checklist</a>
            <br><a id="export-controlled-policy" target="_blank" href="http://mww.mbari.org/exportregs/policy.htm"><font size=2 color="blue" face="Arial,Helvetica">Export controlled policy</a>
            <br><a id="export-controlled-items" target="_blank" href="http://mww.mbari.org/exportregs/items.htm"><font size=2 color="blue" face="Arial,Helvetica">Export controlled items</a>
            <br><br><br><br>

            <input id="itarCheck" type="hidden" name="itarCheck" value="" />
          </p>
          <hr>
        </div>


          <!-- Section 3.6 -->
          <!-- Mooring Recovery or Deployment Declaration Section -->
          <div id="mooring-section" style="display:none;">
          <p class="padding">
            <strong><font face="Arial,Helvetica" size=5 color="green">VI. Mooring recovery / deployment declaration</font></strong><br><br><br>
            <font face="Arial,Helvetica" size="3" color="black"><strong>Will this cruise be RECOVERING a mooring?</strong><select style="margin-left:5px" id="mooring-recovery" name="mooring-recovery" onchange="updateMooringRecoveryDeclaration(this.value)">
              <option value="no-mooring-recovery">No</option>
              <option value="yes-mooring-recovery-surface">Yes - Surface Mooring Recovery</option>
              <option value="yes-mooring-recovery-subsea">Yes - Subsea Mooring Recovery</option>
            </select><br><br>
            <font face="Arial,Helvetica" size="-1" color=brown>If yes, please describe the recovery in detail below:</font><br><br>
            <textarea rows="10" name="MooringRecoveryDesc" id="MooringRecoveryDesc" value="MooringRecoveryDesc" cols="160" wrap></textarea><br>

            <a id="dmo-policies" target="_blank" href="https://mww.mbari.org/dmo/Pages/DMOPolicies.html"><font size=2 color="blue" face="Arial,Helvetica">Related DMO policies</a>
<!--             <br><a id="mbari-nilspin-policy" target="_blank" href="http://www.mbari.org/at-sea/marine-operations-policies/marine-operations-policies-nilspin"><font size=2 color="blue" face="Arial,Helvetica">MBARI NILSPIN policy</a> -->
            <br><br><br><br><br>

            <font face="Arial,Helvetica" size="3" color="black"><strong>Will this cruise be DEPLOYING a mooring?</strong><select style="margin-left:5px" id="mooring-deploy" name="mooring-deploy" onchange="updateMooringDeployDeclaration(this.value)">
              <option value="no-mooring-deploy">No</option>
              <option value="yes-mooring-deploy-surface">Yes - Surface Mooring Deployment</option>
              <option value="yes-mooring-deploy-subsea">Yes - Subsea Mooring Deployment</option>
            </select><br><br>
            <font face="Arial,Helvetica" size="-1" color=brown>If yes, please describe the deployment in detail below:<br><br><strong>NOTE:</strong>  For deployment operations, you will need to request publication in the <br>Local Notice to Mariners (LNM) 14 days prior to start of operation: <a target="_blank" href="https://https://www.pacificarea.uscg.mil/Our-Organization/District-11/Prevention-Division/LnmRequest/">LNM Request</a></font><br><br>
            <textarea rows="10" name="MooringDeployDesc" id="MooringDeployDesc" value="MooringDeployDesc" cols="160" wrap></textarea><br>
            <a id="dmo-policies" target="_blank" href="https://mww.mbari.org/dmo/Pages/DMOPolicies.html"><font size=2 color="blue" face="Arial,Helvetica">Related DMO policies</a>
<!--             <br><a id="mbari-nilspin-policy" target="_blank" href="http://www.mbari.org/at-sea/marine-operations-policies/marine-operations-policies-nilspin"><font size=2 color="blue" face="Arial,Helvetica">MBARI NILSPIN policy</a> -->
            <br><br><br>
          </p>
          <hr>
        </div>



        <!-- Section 4.0 -->
        <!-- Setup Cruise Onboard Participants Section -->
        <div id="cruise-participants-section" style="display:none;">
        <p class="padding"><strong><font face="Arial,Helvetica" size=5 color="green">VII. Participants</font></strong><br><br>
          <font face="Arial,Helvetica" size="2" id="DP_Participants_statement" color=brown><strong>MBARI TWIC/security policy:</strong><br>Participants will require a TWIC card for unescorted access to restricted areas of the vessel.  <br>Participants who do not have a TWIC card will be required to be escorted and monitored by a TWIC cardholder when in restricted areas of the vessel. <br><br><strong>Please indicate with an asterisk (*) the participants who are TWIC cardholders in <u>all</u> cases below.</strong><br><strong><a target="_blank" href="https://mww.mbari.org/dmo/Pages/DPPlanning.html">DP-Berthing Assignments</a><br></strong></font><br>


        <!-- Section 4.1 -->
        <!-- Declare MBARI Participants Section -->
          <font style="width: 1200px" face="Arial,Helvetica" size="3" color=black><strong>Please list all participating onboard MBARI personnel below.</strong></font><font face="Arial,Helvetica" size="2" color=brown> (Required)
          <br><strong>Note: </strong>MBARI employees who are Foreign Nationals should be listed here.</font><br>
          <textarea rows="8" id="Participants" name="Participants" value="Participants" cols="160" wrap></textarea><br><br><br><br><br>


        <!-- Section 4.2 -->
        <!-- Declare non MBARI Participants Section -->
            <font face="Arial,Helvetica" size="3" color=black><strong>Will you have non-MBARI personnel onboard?</strong><select style="margin-left:5px" id="non-mbari-personnel" name="non-mbari-personnel" onchange="updateNonMBARIParticipant(this.value)">
              <option value="no-non-mbari-personnel">No</option>
              <option value="yes-non-mbari-personnel">Yes</option>
            </select><br>
             <font face="Arial,Helvetica" size="2" id="DP_Non_MBARI_Participants_Statement" color=brown><br><strong>NOTE:</strong> The <strong>chief scientist </strong>will need to contact <a href="mailto:lisa@mbari.org">Lisa Emanuelson</a> (or the <a href="mailto:frontdesk@mbari.org">MBARI front desk</a> as backup) to make arrangements for non-MBARI participants to obtain Temporary Cruise Participant badges to keep for the duration of the cruise.  <br>If a participant must retrieve a badge after hours or at a remote port, the Vessel Security Officer (see captain or mates) onboard can also issue badges. <br>Participants must show current photo identification (drivers' license or passport) to obtain a badge.<br></font>
             <br>
             <font face="Arial,Helvetica" size=2 color=black><strong>If yes, please provide a complete list of non-MBARI personnel below. It is the chief scientist's<br> responsibility to ensure that each non-MBARI participant completes: <br>1) a <a target="_blank" href="https://mww.mbari.org/dmo/Pages/VisitingCruiseParticipants.html">Visiting participants release form and information</a> and, <br>2) a <a target="_blank" href="https://www.mbari.org/wp-content/uploads/2015/12/Medical-1C.doc">medical history form</a> <font color="red"> (only if planning for overnight trips)</font>. <br>These completed forms should be given to the captain prior to departure.</strong></font><br>
            <textarea rows="8" id="nonMBARIParticipantsDesc" name="nonMBARIParticipantsDesc" value="nonMBARIParticipantsDesc" cols="160" wrap></textarea><br><br><br><br>


        <!-- Section 4.3 -->
        <!-- Declare Foreign National Participants Section -->
            <font face="Arial,Helvetica" size="3" color=black><strong>Will you have Foreign National personnel onboard?</strong><select style="margin-left:5px" id="foreign-national-personnel" name="foreign-national-personnel" onchange="updateForeignNationalParticipant(this.value)">
              <option value="no-foreign-national-personnel">No</option>
              <option value="yes-foreign-national-personnel">Yes</option>
            </select><br><br>
             <font face="Arial,Helvetica" id="Foreign_participants_statement" size="2" color=brown>Clearly identify all <strong>foreign nationals </strong>in the box below for export regulation screening purposes.  <br>The chief scientist must notify his/her division administrator of all foreign participants at least one (1) week prior to the start of the cruise to allow time for screening.</font>

             <font face="Arial,Helvetica" id="DP_foreign_participants_statement" size="2" color=brown><br>R/V <i>David Packard</i> badges will not be issued until the screening process is completed.</font>

             <br><br>
            <font face="Arial,Helvetica" size=2 color=black><strong>If yes, please provide a complete list of foreign national personnel and affiliation below.</strong></font><br>
            <textarea rows="8" name="ForeignNationalParticipantsDesc" id="ForeignNationalParticipantsDesc" value="ForeignNationalParticipantsDesc"cols="160" wrap></textarea><br><br><br><br>
            <input id="allParticipants" type="hidden" name="allParticipants" value="" />
            <input id="allParticipants_cleanHTML" type="hidden" name="allParticipants_cleanHTML" value="" />
            <input id="foreignNationalCheck" type="hidden" name="foreignNationalCheck" value="" />
            </div>


        <!-- Section 4.4 -->
        <!-- Declare Meal Requests Section -->
        <div id="meal-requests-section" style="display:none;">
          <p class="padding">
             <!--<br><br> -->
             <font face="Arial,Helvetica" size="3" color=black><strong>Meal requests:</strong><br>
             <font face="Arial,Helvetica" size="2" color=brown>For R/V <i>David Packard</i>, please send any special meal requests (vegetarian, etc.) to the <a href="mailto:mealspc@listserver.mbari.org">ship's steward</a>.
             <br><br><br><br></p>
         <!--<hr> -->
        </div>




        <!-- Section 4.5 -->
        <!-- Declare Telepresence On-Shore Participants Section ONLY IF REQUESTED-->
        <div id="telepresence-participants-section" style="display:none;"><hr>
        <p class="padding"><strong><font class="flash" face="Arial,Helvetica" size=5 color="green"><span style='font-size:30px;'>&#8594;</span> Telepresence Participants and Media Broadcast Info</font></strong><font face="Arial,Helvetica" size="2" color=brown> (Required)</font>
          <br><br>
          <font face="Arial,Helvetica" size="2" id="telepresence_participants_statement" color=brown>You indicated above in Section I that Telepresence will be needed for this MBARI expedition. <strong><br><br>PLEASE LIST: 1) ALL INDIVIDUALS, 2) THEIR RESEARCH AFFILIATIONS, AND 3) EMAIL ADDRESSES FOR THOSE THAT WILL BE PARTICIPATING ONSHORE IN THE TEXT AREA BELOW.</strong>   <br>NOTE: You <strong>MUST</strong> provide a completed list of onshore participant names and all information at least <strong>3 weeks</strong> prior to the cruise start date.</font><br>
          <textarea rows="8" id="TelepresenceParticipantsDesc" name="TelepresenceParticipantsDesc" value="TelepresenceParticipantsDesc" cols="160" placeholder="Example: Dr. Albert Einstein, CalTech, aEinstein@caltech.edu" wrap></textarea><br><br>

        <span id="broadcastTP"><br><br>
          <font face="Arial,Helvetica" size="3" color=black><strong>Do you wish to broadcast any feeds publicly?</strong><font face="Arial,Helvetica" size="2" color=brown> (Required)
            <select style="margin-left:5px" id="broadcast-publicly" name="broadcast-publicly" onchange="updateBroadcastPublicly(this.value)">
              <option id="none" value="none">Select One</option>
              <option id="no-broadcast" name="no-broadcast" value="NO">No</option>
              <option id="yes-broadcast" name="yes-broadcast" value="YES">Yes</option>
            </select>
          <br><br><br>
        <!--</span> -->

          <font face="Arial,Helvetica" size="3"  id="listTPPlatforms" color=black><strong>On what public broadcast media platforms?
            (Check all that apply).</strong><font face="Arial,Helvetica" size="2" color=brown> (Required)
          <br><br></font>
          <!--<font face="Arial,Helvetica" size="3" color="black"><strong>Check all that apply</strong>.<br><br> -->
          <fieldset class="broadcastPlatformsRequested" id="broadcastPlatformsRequested" style="display:none;" name="broadcastPlatformsRequested">
            <label id="youTube-media-label" for="youTube" face="Arial,Helvetica" size="-1" color="black"><strong>YouTube</strong><input type="checkbox" id="youTube" name="platformSelect[]" value="YOUTUBE" onclick='updateRequestedPlatforms(this.value)'></label>
            <label id="facebook-media-label" for="facebook" face="Arial,Helvetica" size="-1" color="black"><strong>Facebook Live</strong><input type="checkbox" id="facebook" name="platformSelect[]" value="FACEBOOK_LIVE" onclick='updateRequestedPlatforms(this.value)'></label>
            <label id="twitch-media-label" for="twitch" face="Arial,Helvetica" size="-1" color="black"><strong>Twitch</strong><input type="checkbox" id="twitch" name="platformSelect[]" value="TWITCH" onclick='updateRequestedPlatforms(this.value)'></label>
            <label id="other-telepresence-label" for="other-telepresence-equipment" face="Arial,Helvetica" size="-1" color="black"><strong>Other</strong><input type="checkbox" id="other-telepresence-equipment" name="platformSelect[]" value="OTHER_MEDIA" onclick='enableDisablePlatforms()'></label>
            <br><br>
            <font size=2 color="brown"> *If <strong>"Other"</strong>, please explain in text area below.</font><br>
          <!--</fieldset></p> -->
          <!-- <p class="padding"> -->
          <textarea class="textfield" input="text" rows="10" id="BroadcastPlatformDesc" name="BroadcastPlatformDesc" value="BroadcastPlatformDesc" cols="160" placeholder="Please enter telepresence platform details here. If none, please leave empty." onclick='checkBroadcastPlatformOther()' wrap></textarea><br>
          </fieldset>
          <br>
          <input id="broadcastTP-check" type="hidden" name="broadcastTP-check" value="" />
          <!--<input id="broadcastTPCheck" type="text" name="broadcastTPCheck" value="" />-->
          <br><br>
        </span>
          <br>
         </p>
          <!--<font face="Arial,Helvetica" size="2" id="telepresence_participants_statement" color=brown>In the Cruise purpose section of this form <a href="#telepresencelink">(section I)</a>, you indicated that onshore telepresence would be needed for this cruise. From the Participants list section above, <strong>please list ALL shore-based participants</strong>for this cruise.</font><br>These completed forms should be given to the captain prior to departure.</strong></font><br>-->
        </div>





        <!-- Section 5.0 -->
        <!-- Declare Planned Track Description Section -->
        <div id="planned-track-section" style="display:none;"><hr>
          <p class="padding"><strong><font face="Arial,Helvetica" size=5 color="green">VIII. Planned track description</font></strong><br>
            <span id="plannedTrack"><br><br><font face="Arial,Helvetica" size="3" color=black><strong>Will your planned track extend outside of Monterey Bay (<a target="_blank" href="http://mww.mbari.org/expdlog/expd/log/maps/cropped_LL_navy_ops.pdf">beyond Navy Notification Radius</a>) and involve underwater operations?</strong><font face="Arial,Helvetica" size="2" color=brown> (Required for Navy notification)
              <select style="margin-left:5px" id="planned-track-extent" name="planned-track-extent" value="" onchange="updatePlannedTrackExtent(this.value)">
                <option value="none">Select One</option>
                <option id="track-within-MB" name="track-within-MB" value="NO">No</option>
                <option id="track-outside-MB" name="track-outside-MB" value="YES">Yes</option>
              </select><br>
              <br><br><br>
              <font face="Arial,Helvetica" size="3" color=black><strong>Are you planning on operating in a US Naval restricted area, Pacific Missile Range, or shipping lanes?</strong><font face="Arial,Helvetica" size="2" color=brown> (Required)
                <select style="margin-left:5px" id="planned-track-restricted" name="planned-track-restricted" onchange="updatePlannedTrackRestricted(this.value)">
                  <option value="none">Select One</option>
                  <option id="non-restricted-track" name="non-restricted-track" value="NO">No</option>
                  <option id="yes-restricted-track" name="yes-restricted-track" value="YES">Yes</option>
                </select><br></span>
                <br><br>

              <font face="Arial,Helvetica" size="-1" color=brown>1) Enter text for the cruise track below (must fit in this box), OR <br>2) Select blue CONTINUE button below after you have completed this form, then click 'Select Waypoints' for the waypoint database. Select your waypoints - they will be inserted in the planned track description.<br>
              </font>
              <textarea rows="10" name="PlannedTrackDesc" id="PlannedTrackDesc" value="PlannedTrackDesc" cols="160" wrap="off"></textarea><br><br><br></p>

            <input id="specialOpsCheck" type="hidden" name="specialOpsCheck" value="" />
            <input id="allPlannedTrackInfo" type="hidden" name="allPlannedTrackInfo" value="" />
            <input id="allPlannedTrackInfo_cleanHTML" type="hidden" name="allPlannedTrackInfo_cleanHTML" value="" />
            <hr>
        </div>



        <!-- Section 6.0 -->
        <!-- Compliance with Monterey Bay Sanctuary Permitting Section -->
        <div id="sanctuary-permitting-section" style="display:none;">
          <p class="padding"><strong><font face="Arial,Helvetica" size=5 color="green">IX. Permits</font></strong><br>
                <font face="Arial,Helvetica" size="-1" color=brown><strong>(Please contact <a href="mailto:mandy@mbari.org">Mandy Allen</a> with any and all questions you may have regarding the permitting process.)  NOTE: Water samples do not require permitting.</strong><br>
                </font>
                <br>

            <br><br><font face="Arial,Helvetica" size="3" color=black><strong>Does this operation comply with MBARI's <a id="mbnms-link" target="_blank" href="http://mww.mbari.org/pco/permits/mbnms-2015-002_feb2015.pdf">Monterey Bay National Marine Sanctuary (MBNMS) permit</a>?</strong><font face="Arial,Helvetica" size="2" color=brown> (Required)
                <select style="margin-left:5px" id="mbnms-permit-check" name="mbnms-permit-check" onchange="updateMBNMSPermitCheck(this.value)">
                  <option value="none">Select One</option>
                  <option id="yes-mbnms-compliant" name="yes-mbnms-compliant" value="YES">Yes</option>
                  <option id="notsure-mbnms-compliant" name="notsure-mbnms-compliant" value="NOT_SURE">Not Sure</option>
                </select>

<!--             <br><br><br><br><font face="Arial,Helvetica" size="3" color=black><strong>Will you be collecting any animal, plant, or geological samples on your cruise?</strong><font face="Arial,Helvetica" size="2" color=brown> (Required)
                <select style="margin-left:5px" id="sample-permit-check" name="sample-permit-check" onchange="updateSamplePermitCheck(this.value)">
                  <option value="none">Select One</option>
                  <option id="yes-sample-compliant" name="yes-sample-compliant" value="YES">Yes</option>
                  <option id="no-sample-compliant" name="no-sample-compliant" value="NO">No</option>
                  <option id="notsure-sample-compliant" name="notsure-sample-compliant" value="NOT_SURE">Not Sure</option>
                </select>  -->

            <br><br><br><br><font face="Arial,Helvetica" size="3" color=black><strong>If collecting samples, regardless of location, do you have a valid permit from the <a id="cdfw-permit-link" target="_blank"
            href="https://www.wildlife.ca.gov/Licensing/Scientific-Collecting">California Department of Fish & Wildlife (CDFW)</a>?</strong><font face="Arial,Helvetica" size="2" color=brown> (Required)
                <select style="margin-left:5px" id="cdfw-permit-check" name="cdfw-permit-check" onchange="updateCDFWPermitCheck(this.value)">
                  <option value="none">Select One</option>
                  <option id="yes-cdfw-compliant" name="yes-cdfw-compliant" value="YES">Yes</option>
                  <option id="no-cdfw-compliant" name="no-cdfw-compliant" value="NO">No</option>
                  <option id="notsure-cdfw-compliant" name="notsure-cdfw-compliant" value="NOT_SURE">Not Sure</option>
                  <option id="NA-cdfw-compliant" name="NA-cdfw-compliant" value="NOT_APPLICABLE">Not Applicable</option>
                </select>

<!--             <br><br><br><br><font face="Arial,Helvetica" size="3" color=black><strong>Will you be operating in a <a id="smpa-link" target="_blank" href="https://nrm.dfg.ca.gov/FileHandler.ashx?DocumentID=105423&inline">State Marine Protected Area (SMPA)</a>?</strong><font face="Arial,Helvetica" size="2" color=brown> (Required)
                <select style="margin-left:5px" id="smpa-permit-check" name="smpa-permit-check" onchange="updateSMPAPermitCheck(this.value)">
                  <option value="none">Select One</option>
                  <option id="yes-smpa-compliant" name="yes-smpa-compliant" value="YES">Yes</option>
                  <option id="no-smpa-compliant" name="no-smpa-compliant" value="NO">No</option>
                  <option id="notsure-smpa-compliant" name="notsure-smpa-compliant" value="NOT_SURE">Not Sure</option>
                </select> -->

            <br><br><br><br><font face="Arial,Helvetica" size="3" color=black><strong>If operating or collecting in a <a id="smpa-collection-link" target="_blank"
            href="https://nrm.dfg.ca.gov/FileHandler.ashx?DocumentID=105423&inline">State Marine Protected Area (SMPA)</a>, do you have a valid permit from the <a id="smpa-cdfw-permit-link" target="_blank"
            href="https://www.wildlife.ca.gov/Licensing/Scientific-Collecting">CDFW</a> to do so?</strong><font face="Arial,Helvetica" size="2" color=brown> (Required)
                <select style="margin-left:5px" id="smpa-collection-permit-check" name="smpa-collection-permit-check" onchange="updateSMPAPermitCheck(this.value)">
                  <option value="none">Select One</option>
                  <option id="yes-smpa-collection-permit-compliant" name="yes-smpa-collection-permit-compliant" value="YES">Yes</option>
                  <option id="no-smpa-collection-permit-compliant" name="no-smpa-collection-permit-compliant" value="NO">No</option>
                  <option id="notsure-smpa-collection-permit-compliant" name="notsure-smpa-collection-permit-compliant" value="NOT_SURE">Not Sure</option>
                  <option id="NA-smpa-collection-permit-compliant" name="NA-smpa-collection-permit-compliant" value="NOT_APPLICABLE">Not Applicable</option>
                </select>

              <br><br><br>
              <font face="Arial,Helvetica" size="-1" color=brown>Please enter any additional comments or details below (must fit in this box).<br>
              </font>
              <textarea rows="10" name="PermitDesc" id="PermitDesc" value="PermitDesc" cols="160" wrap="off"></textarea>
              <br><br><br>

              <input id="permitsCheck" type="hidden" name="permitsCheck" value="" />
          </p>
          <hr>
        </div>



        <!-- Section 7.0 -->
        <!-- Chief Scientist Agreement to follow Precruise outlined here Section -->
        <div id="chief-sci-agreement-section" style="display:none;">
          <p class="padding"><font face="Arial,Helvetica" size=5 color="green"><strong>X. Agreement to precruise and MBARI policies/procedures.</font></strong><br><br>

             <font face="Arial,Helvetica" size="3" color=black><strong>If the submitted plan requires changing while offshore this must be discussed with the Captain for review.<br><br>
              <font face="Arial,Helvetica" size="3" color=brown><input type="checkbox" name="MBARIPolicyAgree" value="MBARIPolicyAgree" id="MBARIPolicyAgree" /> I agree to follow <a id="precruise_plans_link" target="_blank">precruise plans</a>, <a id="dmo_safety_link" target="_blank">DMO vessel safety procedures</a>, and MBARI <a id="mbari_ship_policy_link" target="_blank">policies/procedures</a>.</strong>
            </form>
          </p>
        </div>

        <!-- Section 8.0 -->
        <!-- Submit Button for Precruise Entry: Step 1.  On to Step 2 from here. -->
        <div id="submit-button-section" style="display:none;">
          <p></p><br>
          <p class="center"><input type="submit" class="button blue" value="Continue" name="Step2" style="width: auto">
            <input type="reset" class="button gray" value="Reset form fields" name="Reset form" style="width: auto">
          </p><br><br>
        </div>
    </td>
  </tr>
</table>


<!-- Form CSS -->
        <style type="text/css">
           /*label {
            display: block;
            padding-left: 15px;
            padding-right: 30px;
            text-indent: -15px;
           }*/
            pre {
                font-family: "courier new";
            }
            input {
                width: 30px;
                /*overflow:hidden;*/
                /*height: 13px;
                padding: 0;
                margin:0;
                vertical-align: bottom;
                position: relative;
                top: -1px;
                *overflow: hidden;*/
            }
            input:hover {
                background: beige;
            }
            ::-webkit-input-placeholder { /* WebKit, Blink, Edge */
              /*color:    #000000;*/
              /*color: #C7C7CD;*/
              color: #A4A4b2;
              /*font-weight: bold;*/
              font-size: 12px;
            }
            :-moz-placeholder { /* Mozilla Firefox 4 to 18 */
              /*color:    #000000;*/
              color: #A4A4b2;
              opacity:  1;
              /*font-weight: bold;*/
              font-size: 12px;
            }
            ::-moz-placeholder { /* Mozilla Firefox 19+ */
              /*color:    #000000;*/
              color: #A4A4b2;
              opacity:  1;
              /*font-weight: bold;*/
              font-size: 12px;
            }
            :-ms-input-placeholder { /* Internet Explorer 10-11 */
              /*color:    #000000;*/
              color: #A4A4b2;
              /*color: #C7C7CD;*/
              /*font-weight: bold;*/
              font-size: 12px;
            }
            label input {
              margin-right: 25px;
            }
            select {
              margin-right: 20px;
            }
            select:hover {
              background: beige;
            }
            textarea {}
            textarea:hover {
                background: beige;
            }
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
            button.grayblue,
            .button.grayblue {
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
            p.button {
              text-align: center;
            }
            p.center {
              text-align: center;
            }
            p.padding {
              padding-top: 0px;
              padding-right: 0px;
              padding-bottom: 0px;
              padding-left: 6px;
            }
            div.padding {
              padding-top: 0px;
              padding-right: 0px;
              padding-bottom: 0px;
              padding-left: 6px;
            }
            a.padding {
              padding-top: 0px;
              padding-right: 0px;
              padding-bottom: 0px;
              padding-left: 6px;
            }
            fieldset {
              border-width: 0px;
              margin-bottom: 0px;
              padding-top: 0px;
              padding-right: 0px;
              padding-bottom: 0px;
              padding-left: 6px;
            }
            a {
              color: blue;
            }
            a.link {
            }
            .flash {
              background-color:yellow;
              animation-name:bckanim;
              animation-fill-mode:forwards;
              animation-duration:5s;
              animation-delay:0s;
              display: inline-block; /* Crucial for some browsers to animate */
            }
            @keyframes bckanim {
              0% {background-color:yellow;}
              100% { background-color:transparent;}
            }
            .hidden
            {
              display: none;
            }
           .shipSelect select {
            /* Styling */
              width: 200px;
              text-indent:5px;
              -webkit-appearance: none;
              -moz-appearance: none;
              color: red;
              border-color: gray;
              border-width: 1px;
              font-weight: bold;
              font-size: 12px;

              /* SET UP DROPDOWN LIST ARROW */
              background: url("./font-awesome/fa-sort-red_smallHR9.png") no-repeat 99%;
            }
            ul {
              word-wrap: break-word;
            }
            /* Apply padding to td elements that are direct children of the tr elements with class spaceUnder. */
            td {

            }
            .txtbox {
              font-size: 8pt;
              text-align: left;
              padding: 2px;
              padding-bottom: 160px;
              width: 1098px;
              height: 180px;
              line-height: 180px;
            }
            .txtbox.placeholder {
              font-size: 8pt;
              text-align: left;
              padding: 2px;
              width: 1098px;
              height: 180px;
              line-height: 180px;
            }
            tr.row2 td {
              padding-top: 40px;
            }
            table td > div {
              position: relative;
            }
        </style>
  </form>

<%
$RSpm->Close;
$RScs->Close;
$Conn->Close;

} # End step1()

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
  # If coming back from waypoint selection, do not try to get these form values!!!
  # NEED THIS TO GO FROM STEP1 to STEP2

  if ( $flag ne 'orderwpts' ) {
    @dl = ( 0, substr(GetFormValue('dTime'),2,2), substr(GetFormValue('dTime'),0,2),
      GetFormValue('dDay'), GetFormValue('dMonth')-1, GetFormValue('dYear')-1900 );
    @al = ( 0, substr(GetFormValue('aTime'),2,2), substr(GetFormValue('aTime'),0,2),
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

   # $Session->{'Purpose'} = GetFormValue('allCruisePurposeInfo');
   # $Session->{'PlannedTrackDesc'} = GetFormValue('allPlannedTrackInfo');
   # $Session->{'Participants'} = GetFormValue('allParticipants');
   # $Session->{'EquipmentDesc'} = GetFormValue('allEquip');

    $Session->{'Purpose'} = GetFormValue('allCruisePurposeInfo_cleanHTML');
    $Session->{'PlannedTrackDesc'} = GetFormValue('allPlannedTrackInfo_cleanHTML');
    $Session->{'Participants'} = GetFormValue('allParticipants_cleanHTML');
    $Session->{'EquipmentDesc'} = GetFormValue('allEquip_cleanHTML');

    $Session->{'diverRequested'} = GetFormValue('diverRequested');
    $Session->{'telepresenceRequested'} = GetFormValue('telepresenceRequested');
    $Session->{'broadcastTPCheck'} = GetFormValue('broadcastTPCheck');
    $Session->{'waveG'} = GetFormValue('waveG');
    $Session->{'ctdCheck'} = GetFormValue('ctdCheck');
    $Session->{'hazIsoCheck'} = GetFormValue('hazIsoCheck');
    $Session->{'winchCheck'} = GetFormValue('winchCheck');
    $Session->{'itarCheck'} = GetFormValue('itarCheck');
    $Session->{'foreignNationalCheck'} = GetFormValue('foreignNationalCheck');
    $Session->{'specialOpsCheck'} = GetFormValue('specialOpsCheck');
    $Session->{'permitsCheck'} = GetFormValue('permitsCheck');



    %>
     <script Language="JavaScript"><!--
        function Field_Validator(theForm)
        {
          return(true);
        }
      </script>
      <%
  }
  # OTHERWISE RETURNING FROM WAYPOINTS - FORGETS SESSION INFORMATION!!
  else {
%>
     <script Language="JavaScript">

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
       </script>
  <%

  }  # End of if/else statement

  ################################################
  # Save step 2 variables for use in email message
  ################################################
  $step2args = '';
  my $tmp = '';
  foreach my $f ( Win32::OLE::in ($Request->QueryString) ) {
    #
    # For text fields, only show first 5 characters
    #
    if ($f eq 'Participants' || $f eq 'PlannedTrackDesc' || $f eq 'Purpose' || $f eq 'EquipmentDesc' || $f eq 'diverRequested' || $f eq 'telepresenceRequested' || $f eq 'waveG' || $f eq 'ctdCheck' || $f eq 'hazIsoCheck' || $f eq 'winchCheck' || $f eq 'itarCheck' || $f eq 'foreignNationalCheck' || $f eq 'specialOpsCheck' || $f eq 'permitsCheck') {
      $tmp .= "$f=" . substr($Request->QueryString($f)->{Item},0,5) . "&";
    } else {
      $tmp .= "$f=" . $Request->QueryString($f)->{Item} . "&";
    }
  }

  $step2args = $tmp;
  $step2args =~ s/\&$//;
  # $step2args =~ s|</b>||ig;
  # $step2args =~ s|<b>||ig;

  $Session->{'Step2Args'} = URLEncode($step2args);

  #
  # Get next Ship sequence number
  #
  open_database($dsn);    # Creates $Conn object as a global variable

  $sql = "SELECT max(ShipSeqNum) FROM Expedition WHERE (ShipName = '" . $Session->{'ShipName'} . "')";
  Win32::ASP::DebugPrint("\nSQL = $sql ");
  $RS = $Conn->Execute($sql);
  $Session->{'ShipSeqNum'} = $RS->Fields(0)->value + 1;

  $RS->Close;
  $Conn->Close;
%>

<!-- CHECK THESE ENTRIES FORM - STEP 2 -->

<h3>Check these entries...</h3>

<!-- STEP 2 Form CSS -->
        <style type="text/css">
            input {
                width: 100px;
            }
            button,
            .button {
              color : #fff;
              cursor : pointer;
              border-style : solid;
              border-width : 1px;
            }
            button.blue,
            .button.blue {
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
               font: 11px Arial;
               text-decoration: none;
               background-color: #D5D5D5;
               color: #333333;
               padding: 2px 6px 2px 6px;
               border-top: 1px solid #CCCCCC;
               border-right: 1px solid #333333;
               border-bottom: 1px solid #333333;
               border-left: 1px solid #CCCCCC;
            }
            button.grayblue,
            .button.grayblue {
              font: 11px Arial;
              text-decoration: none;
              background-color: #D5D5D5;
              color: #333333;
              padding: 2px 6px 2px 6px;
              border-top: 1px solid #CCCCCC;
              border-right: 1px solid #333333;
              border-bottom: 1px solid #333333;
              border-left: 1px solid #CCCCCC;
            }
            p.button {
              text-align: left;
            }
        </style>

<!-- ########################################################################################### -->
<!-- Need to use "GET" for waypoint string creation to work -->
<!-- ########################################################################################### -->

<form method="GET" action="<%= $this_script%>" onsubmit="return Field_Validator(this)" accept-charset="utf-8" name=order>

<table border="0">
<tr>
  <td valign="top"><font face="Helvetica,Arial" size="-1"><strong>MBARI Project Number:</strong></font></td>
  <td valign="top"><font face="Helvetica,Arial" size="-1"><%= $Session->{ProjNum}%></font>
</tr>
<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
<%
#
# Loop through fields list and PRINT out all session variables that match
#
foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
  next if $f eq 'ExpeditionID';
  last if $f eq 'StartDtg';   # Following fields come from Post Cruise entry
  %>
  <tr>

  <% if ( $f =~ /StartDtg$/ && $Session->{$f} ) { # Add local time %>
       <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td>
       <td valign="top"><font face="Helvetica,Arial" size="-1"><%= $Session->{$f}%></font>
       <font face="Helvetica,Arial" size="-1">GMT (<%= $Session->{'ScheduledStartDtgLocal'}%> local time)</font>
  <% } %>
  <% elsif ( $f =~ /EndDtg$/ && $Session->{$f} ) {  # Add local time %>
       <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td>
       <td valign="top"><font face="Helvetica,Arial" size="-1"><%= $Session->{$f}%></font>
       <font face="Helvetica,Arial" size="-1">GMT (<%= $Session->{'ScheduledEndDtgLocal'}%> local time)</font>
  <% } %>
  <% elsif ( $f =~ /Purpose$/ && $Session->{$f} ) {  # Preseve format %>
       <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td>
       <td valign="top"><pre><%= $Session->{'Purpose'}%></pre>
  <% } %>
    <% elsif ( $f =~ /EquipmentDesc$/ && $Session->{$f} ) {  # Preseve format %>
       <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td>
       <td valign="top"><pre><%= $Session->{'EquipmentDesc'}%></pre>
  <% } %>
  <% elsif ( $f =~ /PlannedTrackDesc$/ && $Session->{$f} ) {  # Preseve format %>
       <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td>
       <td valign="top"><pre><%= $Session->{'PlannedTrackDesc'}%></pre>
  <% } %>
    <% elsif ( $f =~ /Participants$/ && $Session->{$f} ) {  # Preseve format %>
       <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td><br>
       <td valign="top"><pre><%= $Session->{'Participants'}%></pre>
  <% }
     else { %>
       <td valign="top"><font face="Helvetica,Arial" size="-1"><strong><%= $f%>:</strong></font></td>
       <td valign="top"><pre><%= $Session->{$f}%></pre>
  <% } %>
  </td>
  </tr>
<%
}
#
# Need to get PersonID numbers for waypoint.asp
#
open_database($dsn);

# Principal Investigator
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

# Chief Scientist
($pFirstName, $pLastName) = split('\s', $Session->{'ExpdChiefScientist'});
# Check for names that may contain apostrophes (e.g., O'Reilly).  The string sent to SQL must look like: 'O''Reilly'
if ($pLastName =~ "'") {
   $pLastName = s/"'"/"''"/;
 }
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

<tr></tr>
<tr>
<td><font face="Helvetica,Arial" size="-1" color="#8F8F8F"><b>Ship Sequence Number:</b></font></td>
<td><font face="Helvetica,Arial" size="-1" color="#8F8F8F"><%= $Session->{'ShipSeqNum'}%></font></td>
</tr>

<tr>
<td colspan="2"><font face="Helvetica,Arial" size="-1" color="red">
<br>If something is not right here then press your browser's <strong>BACK</strong> button to the <strong>Step 1 long form entry page</strong> and fix it.
<br>If the above is OK then select your waypoints and finish the precruise.
</font></td>
</tr>

<%
# THIS IS CAUSING ALL THE ISSUES WITH WAYPOINTS
  # cleanHTML();
%>

<br>
<tr><td colspan="2"><br><h3>Select Waypoints to be visited...</h3><td></tr>

<%

 ###############################################################################
 ############   RETURNING FROM WAYPOINTS SECTION HERE ##########################
 ###############################################################################

%>
         <script Language="JavaScript">
         // DO SOME LOCAL STORAGE MAGIC IN ORDER TO ALLOW ALIASING AFTER RETURNING FROM WAYPOINTS ACCESS.
         // SAVE PREVIOUSLY ANSWERED RESPONSES FOR STORAGE AND RECALL AFTER WAYPOINT VISIT TO ALLOW EMAIL ALIASING.
          function saveStorage() {
            //DIVER - Never null.
            //console.log("DIVER-CHECK = " + document.getElementById('diver-requested').value);
            var diverFieldStorage = document.getElementById('diver-requested').value;
            //console.log("DiverFieldStorage = " + diverFieldStorage);
            localStorage.setItem('diverStorage', JSON.stringify(diverFieldStorage));
            //console.log("Value Saved As = " + diverFieldStorage);


            //TELEPRESENCE - Never null.
            var telepresenceFieldStorage = document.getElementById('telepresence-requested').value;
            //console.log("DiverFieldStorage = " + diverFieldStorage);
            localStorage.setItem('telepresenceStorage', JSON.stringify(telepresenceFieldStorage));
            //console.log("Value Saved As = " + telepresenceFieldStorage);

/*            if (document.getElementById('telepresence-required').value != null) {
              if (document.getElementById('telepresence-required').value == "YES") {
                var telepresenceFieldStorage = "YES";
                localStorage.setItem('telepresenceStorage', JSON.stringify(telepresenceFieldStorage));
              } else {
                var telepresenceFieldStorage = "NO";
                localStorage.setItem('telepresenceStorage', JSON.stringify(telepresenceFieldStorage));
                }
              //console.log("telepresence-required = " + document.getElementById('telepresence-required').value);
              //console.log("TelepresenceFieldStorage = " + telepresenceFieldStorage);
              //console.log("Value Saved As = " + telepresenceFieldStorage);
              } else {
                /////  Do nothing.
              }*/


            // BROADCAST TELEPRESENCE PUBLICLY
            if (document.getElementById('broadcastTP-check').value != null) {
              if (document.getElementById('broadcastTP-check').value == "YES") {
                var broadcastTPFieldStorage = "YES";
                localStorage.setItem('broadcastTPStorage', JSON.stringify(broadcastTPFieldStorage));
              } else {
                var broadcastTPFieldStorage = "NO";
                localStorage.setItem('broadcastTPStorage', JSON.stringify(broadcastTPFieldStorage));
                }
              } else {
                /////  Do nothing.
              }


            // ASV
            //console.log("ASV-CHECK = " + document.getElementById('asv-check').value);
            if (document.getElementById('asv-check').value != null) {
              if (document.getElementById('asv-check').value == "YES") {
                var asvFieldStorage = "YES";
                localStorage.setItem('asvStorage', JSON.stringify(asvFieldStorage));
              } else {
                var asvFieldStorage = "NO";
                localStorage.setItem('asvStorage', JSON.stringify(asvFieldStorage));
                }
              } else {
                /////  Do nothing.
              }

            // CTD
            //console.log("CTD-CHECK = " + document.getElementById('ctd-check').value);
            if (document.getElementById('ctd-check').value != null) {
              if (document.getElementById('ctd-check').value == "YES") {
                var ctdFieldStorage = "YES";
                localStorage.setItem('ctdStorage', JSON.stringify(ctdFieldStorage));
              } else {
                var ctdFieldStorage = "NO";
                localStorage.setItem('ctdStorage', JSON.stringify(ctdFieldStorage));
              }
            }

            // HAZMAT ISOTOPE
            //console.log("HAZISO-CHECK = " + document.getElementById('hazIso-check').value);
            if (document.getElementById('hazIso-check').value != null) {
              if (document.getElementById('hazIso-check').value == "YES") {
                var hazIsoFieldStorage = "YES";
                localStorage.setItem('hazIsoStorage', JSON.stringify(hazIsoFieldStorage));
              } else {
                var hazIsoFieldStorage = "NO";
                localStorage.setItem('hazIsoStorage', JSON.stringify(hazIsoFieldStorage));
              }
            }

            // WINCH
            //console.log("WINCH-CHECK = " + document.getElementById('winch-check').value);
            if (document.getElementById('winch-check').value != null) {
              if (document.getElementById('winch-check').value == "YES") {
                var winchFieldStorage = "YES";
                localStorage.setItem('winchStorage', JSON.stringify(winchFieldStorage));
              } else {
                var winchFieldStorage = "NO";
                localStorage.setItem('winchStorage', JSON.stringify(winchFieldStorage));
              }
            }

            // ITAR
            //console.log("ITAR-CHECK = " + document.getElementById('itar-check').value);
            if (document.getElementById('itar-check').value != null) {
              if (document.getElementById('itar-check').value == "YES") {
                var itarFieldStorage = "YES";
                localStorage.setItem('itarStorage', JSON.stringify(itarFieldStorage));
              } else {
                var itarFieldStorage = "NO";
                localStorage.setItem('itarStorage', JSON.stringify(itarFieldStorage));
              }
            }

            // FOREIGN NATIONAL
            //console.log("FOREIGN-NATIONAL-CHECK = " + document.getElementById('foreignNational-check').value);

            if (document.getElementById('foreignNational-check').value != null) {
              if (document.getElementById('foreignNational-check').value == "YES") {
                var foreignNationalFieldStorage = "YES";
                localStorage.setItem('foreignNationalStorage', JSON.stringify(foreignNationalFieldStorage));
              } else {
                var foreignNationalFieldStorage = "NO";
                localStorage.setItem('foreignNationalStorage', JSON.stringify(foreignNationalFieldStorage));
              }
            }

            // SPECIAL OPS
            //console.log("SPECIAL-OPS-CHECK = " + document.getElementById('specialOps-check').value);
            if (document.getElementById('specialOps-check').value != null) {
              if (document.getElementById('specialOps-check').value == "YES") {
                var specialOpsFieldStorage = "YES";
                localStorage.setItem('specialOpsStorage', JSON.stringify(specialOpsFieldStorage));
              } else {
                var specialOpsFieldStorage = "NO";
                localStorage.setItem('specialOpsStorage', JSON.stringify(specialOpsFieldStorage));
              }
            }

            // PERMITS
            //console.log("PERMITS-CHECK = " + document.getElementById('permits-check').value);
            if (document.getElementById('permits-check').value != null) {
              if (document.getElementById('permits-check').value == "YES") {
                var permitsFieldStorage = "YES";
                localStorage.setItem('permitsStorage', JSON.stringify(permitsFieldStorage));
              } else {
                var permitsFieldStorage = "NO";
                localStorage.setItem('permitsStorage', JSON.stringify(permitsFieldStorage));
              }
            }
        } // End of ELSE statement
       </script>
       <%


 ###############################################################################
 ##########     Section under SELECT WAYPOINTS TO BE VISITED   #################
 ###############################################################################

  if ( $flag eq 'orderwpts' ) {
    %>
    <tr><th align="right">Order of visit</th><th align="left">&nbsp;&nbsp;Waypoint Name</th></tr>
    <%
    my @wpts = Win32::OLE::in ($Request->QueryString('selectedWaypoints'));
    ##print "#wpts = $#wpts<br>";
    $iWpts = 0;
    if ( $#wpts ) {   # If more than 2 then demand order
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
      <tr>
        <td align="right">
          <select name="<%= $s%>" size="1">
              <%= $wptOrder_list%>
          </select>
        </td>
      <td>&nbsp;&nbsp;<%=$s%></td>  <!-- Make a space -->
      </tr><%
    }%>
      <td>&nbsp;&nbsp;<%=$s%></td>  <!-- Make a space -->
      <td>&nbsp;&nbsp;<%=$s%></td>  <!-- Make a space -->
      <td>&nbsp;&nbsp;<%=$s%></td>  <!-- Make a space -->
    <tr>
      <td valign="top" colspan="2">
        <img src="compassrose_small.jpg" border="0">
        <a href="waypoint.asp?action=select&piID=<%= $Session->{'piID'}%>&csID=<%= $Session->{'csID'}%>" class="button gray" font:"size=3" letter-spacing:"0.5px" style="width: auto"><strong>
          Select different waypoints</strong></a>
      </td>
    </tr>
       <td>&nbsp;&nbsp;<%=$s%></td>  <!-- Make a space -->
       <td>&nbsp;&nbsp;<%=$s%></td>  <!-- Make a space -->


        <!-- <br> Return from Waypoints:  Setting storedField alias values.<br> -->
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedDiverField" id="storedDiverField" value=""/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedTelepresenceField" id="storedTelepresenceField" value=""/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedBroadcastTPField" id="storedBroadcastTPField" value=""/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedASVField" id="storedASVField" value=""/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedCTDField" id="storedCTDField" value=""/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedHazIsoField" id="storedHazIsoField" value=""/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedWinchField" id="storedWinchField" value=""/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedITARField" id="storedITARField" value=""/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedForeignNationalField" id="storedForeignNationalField" value=""/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedSpecialOpsField" id="storedSpecialOpsField" value=""/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="storedPermitsField" id="storedPermitsField" value=""/>

       <script Language="JavaScript">
       var storedDiverField = localStorage.getItem('diverStorage');
       storedDiverField = storedDiverField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedDiverField').value = storedDiverField;
       //console.log("DIVER STORAGE = " + storedDiverField);

       var storedTelepresenceField = localStorage.getItem('telepresenceStorage');
       storedTelepresenceField = storedTelepresenceField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedTelepresenceField').value = storedTelepresenceField;
       //console.log("TELEPRESENCE STORAGE = " + storedTelepresenceField);

       var storedBroadcastTPField = localStorage.getItem('broadcastTPStorage');
       storedBroadcastTPField = storedBroadcastTPField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedBroadcastTPField').value = storedBroadcastTPField;
       //console.log("BROADCAST TP STORAGE = " + storedBroadcastTPField);

       var storedASVField = localStorage.getItem('asvStorage');
       storedASVField = storedASVField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedASVField').value = storedASVField;
       //console.log("ASV STORAGE = " + storedASVField);

       var storedCTDField = localStorage.getItem('ctdStorage');
       storedCTDField = storedCTDField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedCTDField').value = storedCTDField;
       //console.log("CTD STORAGE = " + storedCTDField);

       var storedHazIsoField = localStorage.getItem('hazIsoStorage');
       storedHazIsoField = storedHazIsoField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedHazIsoField').value = storedHazIsoField;
       //console.log("HAZ ISO STORAGE = " + storedHazIsoField);

       var storedWinchField = localStorage.getItem('winchStorage');
       storedWinchField = storedWinchField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedWinchField').value = storedWinchField;
       //console.log("WINCH STORAGE = " + storedWinchField);

       var storedITARField = localStorage.getItem('itarStorage');
       storedITARField = storedITARField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedITARField').value = storedITARField;
       //console.log("ITAR STORAGE = " + storedITARField);

       var storedForeignNationalField = localStorage.getItem('foreignNationalStorage');
       storedForeignNationalField = storedForeignNationalField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedForeignNationalField').value = storedForeignNationalField;
       //console.log("FOREIGN NATIONAL STORAGE = " + storedForeignNationalField);

       var storedSpecialOpsField = localStorage.getItem('specialOpsStorage');
       storedSpecialOpsField = storedSpecialOpsField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedSpecialOpsField').value = storedSpecialOpsField;
       //console.log("SPECIAL OPS STORAGE = " + storedSpecialOpsField);

       var storedPermitsField = localStorage.getItem('permitsStorage');
       storedPermitsField = storedPermitsField.replace(/^"(.+)"$/,'$1');
       document.getElementById('storedPermitsField').value = storedPermitsField;
       //console.log("PERMITS STORAGE = " + storedPermitsField);
       </script>

        <!-- <br> Return from Waypoints:  Setting storedField values to ALIASES.<br> -->
        <input type="hidden"style="width: auto" letter-spacing:"0.5px" name="diver-requested" id="diver-requested" onchange="Field_Validator(this)" value="<%=GetFormValue('storedDiverField')%>"/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="telepresence-requested" id="telepresence-requested" onchange="Field_Validator(this)" value="<%=GetFormValue('storedTelepresenceField')%>"/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="broadcastTP-check" id="broadcastTP-check" onchange="Field_Validator(this)" value="<%=GetFormValue('storedBroadcastTPField')%>"/>
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="asv-check" id="asv-check" onchange="Field_Validator(this)" value="<%=GetFormValue('storedASVField')%>">
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="ctd-check" id="ctd-check" onchange="Field_Validator(this)" value="<%=GetFormValue('storedCTDField')%>">
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="hazIso-check" id="hazIso-check" onchange="Field_Validator(this)" value="<%=GetFormValue('storedHazIsoField')%>">
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="winch-check" id="winch-check" onchange="Field_Validator(this)" value="<%=GetFormValue('storedWinchField')%>">
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="itar-check" id="itar-check" onchange="Field_Validator(this)" value="<%=GetFormValue('storedITARField')%>">
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="foreignNational-check" id="foreignNational-check" onchange="Field_Validator(this)" value="<%=GetFormValue('storedForeignNationalField')%>">
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="specialOps-check" id="specialOps-check" onchange="Field_Validator(this)" value="<%=GetFormValue('storedSpecialOpsField')%>">
        <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="permits-check" id="permits-check" onchange="Field_Validator(this)" value="<%=GetFormValue('storedPermitsField')%>">


        <script Language="Javascript">
        if (document.getElementById('broadcastTP-check').value == "YES") {
          var btpc = GetFormValue('storedBroadcastTPField');
          updateBroadcastPublicly(btpc);
        }

        document.getElementById('diver-requested').value = storedDiverField;
        //console.log("STORED DIVER STORAGE = " + document.getElementById('storedDiverField').value);
        //console.log("DIVER STORAGE = " + document.getElementById('diver-requested').value);

        document.getElementById('telepresence-requested').value = storedTelepresenceField;
        //console.log("TELEPRESENCE STORAGE = " + document.getElementById('telepresence-requested').value);

        document.getElementById('broadcastTP-check').value = storedBroadcastTPField;
        //console.log("BROADCAST TELEPRESENCE STORAGE = " + document.getElementById('broadcastTP-check').value);

        document.getElementById('asv-check').value = storedASVField;
        //console.log("ASV STORAGE = " + document.getElementById('asv-check').value);

        document.getElementById('ctd-check').value = storedCTDField;
        //console.log("CTD STORAGE = " + document.getElementById('ctd-check').value);

        document.getElementById('hazIso-check').value = storedHazIsoField;
        //console.log("HAZ ISO STORAGE = " + document.getElementById('hazIso-check').value);

        document.getElementById('winch-check').value = storedWinchField;
        //console.log("WINCH STORAGE = " + document.getElementById('winch-check').value);

        document.getElementById('itar-check').value = storedITARField;
        //console.log("ITAR STORAGE = " + document.getElementById('itar-check').value);

        document.getElementById('foreignNational-check').value = storedForeignNationalField;
        //console.log("FOREIGN NATIONAL STORAGE = " + document.getElementById('foreignNational-check').value);

        document.getElementById('specialOps-check').value = storedSpecialOpsField;
        //console.log("SPECIAL OPS STORAGE = " + document.getElementById('specialOps-check').value);

        document.getElementById('permits-check').value = storedPermitsField;
        //console.log("PERMITS STORAGE = " + document.getElementById('permits-check').value);

        function Field_Validator(theForm) {
          //console.log("FIELD FORM TESTING " + document.getElementById('diver-requested'));
          return(true);
        }
        </script>
<%  }
else { %>
      <!--<br> Did NOT return from Waypoints: Original alias values.<br> -->
      <!-- EMAIL ALIASING - A BIT CLUGEY BUT WORKS - DISPLAYS AT TOP OF STEP 2 PAGE WHEN NOT HIDDEN -->
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="diver-requested" id="diver-requested" value="<%=GetFormValue('diverRequested')%>">
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="telepresence-requested" id="telepresence-requested" value="<%=GetFormValue('telepresenceRequested')%>">
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="broadcastTP-check" id="broadcastTP-check" value="<%=GetFormValue('broadcastTP-check')%>">
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="asv-check" id="asv-check" value="<%=GetFormValue('waveG')%>">
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="ctd-check" id="ctd-check" value="<%=GetFormValue('ctdCheck')%>">
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="hazIso-check" id="hazIso-check" value="<%=GetFormValue('hazIsoCheck')%>">
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="winch-check" id="winch-check" value="<%=GetFormValue('winchCheck')%>">
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="itar-check" id="itar-check" value="<%=GetFormValue('itarCheck')%>">
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="foreignNational-check" id="foreignNational-check" value="<%=GetFormValue('foreignNationalCheck')%>">
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="specialOps-check" id="specialOps-check" value="<%=GetFormValue('specialOpsCheck')%>">
          <input type="hidden" style="width: auto" letter-spacing:"0.5px" name="permits-check" id="permits-check" value="<%=GetFormValue('permitsCheck')%>">



    <tr>
    <td valign="top" colspan="2">
    <img src="compassrose_small.jpg" border="0">
    <a href="waypoint.asp?action=select&piID=<%= $Session->{'piID'}%>&csID=<%= $Session->{'csID'}%>" class="button gray" font:"size=3" letter-spacing:"0.5px" style="width: auto" id="saveLocalStorage" onclick="saveStorage()"><strong>Select Waypoints</strong></font></a>
    (Select from list of PI's and Chief Scientist's waypoints, then return here to finish up.)
    </td>
    </tr>
  <% }  %>


</table>

<!-- Redirect to Step 4:  "Email Form to Logistics Coordinator"
 ###############################################################################
 ############   RETURNING FROM WAYPOINTS SECTION HERE ##########################
 ###############################################################################
-->
<hr align="left" width="1078">
<input type="hidden" name="step" value="4"><br>
</td></tr>

<% if ( $Session->{'DBadministrator'} ) { %>
        <font face="Helvetica,Ariel" size="-1">Send email to
        <input type="text" name="pre_email" value="precruise" size="9" maxlength="9">
        @listserver.mbari.org and </font> <input type="submit" class="button blue" name="finish" letter-spacing:"0.5px" style="width: auto" value="Load Data Base">
<% }
   else {   # We email the form
%>
        <br><b>Enter your USERNAME <font color="red">(not complete email address)</font> to get a copy of this precruise:</b>
        <input type="text"  size=8 name="sender_email_addr" value="<%=GetFormValue('sender_email_addr')%>"><b>@mbari.org</b>
        <p class="button"><input type="submit" class="button blue" name="finish" letter-spacing:"0.5px" style="width: auto" value="Email Form to Logistics Coordinator"></p>
<% } %>

</form>   <!-- END OF STEP 2 FORM -->
<%
} # End step2()
%>




<%
#
#/*======================================================================
# getWaypoints()
#
# Description:
#      Pass chief scientists & PI to waypoints.asp in order to
#    have user select a list of waypoints to add to the cruise
#    plan.
#
#
# Author: Mike McCann
# Date Created: 4/16/01
#====================================================================== */
sub getWaypoints {

  %>ExpdChiefScientist = <%= GetFormValue('ExpdChiefScientist')%> <br>
  <%
  if ( ! GetFormValue('ExpdChiefScientist') ) {
    %><h2>Must select a Chief Scientist before selecting waypoints</h2>
  <%  return;
  }
  %>

<%
} # End getWaypoints
%>

<%
#
#/*======================================================================
# step3()
#
# Description:
#      Chance to verify form entries  (I don't think this function
#    is used anymore.... 11/8/99)
#
#
# Author: Mike McCann
# Date Created: 10/21/98
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
    <% if ( $f =~ /StartDtg$/ && $Session->{$f} ) { # Add local time %>
           <font face="Helvetica,Arial" size="-1">GMT (<%= $Session->{'ScheduledStartDtgLocal'}%> local time)</font>
    <% } %>

    <% if ( $f =~ /EndDtg$/ && $Session->{$f} ) { # Add local time %>
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

<p><font face="Helvetica,Ariel" size="-1">Send email to
  <input type="text" name="pre_email" value="precruise" size="9" maxlength="9">
  @listserver.mbari.org when finished.</font></p>


<%
} # End step3()
%>

<%
#--------------------------------------------------------------------
#

=head3 load()

Process the form field entries and insert into the database
Here is where we need to map the names we have on the HTML
forms for all the fields to the names of the fields that are
actually in the database table.

Author: Mike McCann

Date Created: 10/28/98

=cut


sub load {

#
# Open the database, this time to insert fields
#

open_database($dsn);    # Creates $Conn object as a global variable

#
# Loop through all form names and construct SQL insert string
#
#$sql = "INSERT INTO Expedition\n(";
$sql = "exec insert_expedition ";
#foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
# next unless $Session->{$f};
# $sql .= "$f,";
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
  <h3>Precruise entry failed! <br> Please contact information applications support.</h3>
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

} # End load()
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
    next if ( $f eq 'step' && $Request->QueryString($f)->Item == 4);  #Hope that no waypoints are named 'step'
    next if ( $f eq 'finish' && ( $Request->QueryString($f)->Item eq 'Email Form to Logistics Coordinator'
                             ||   $Request->QueryString($f)->Item eq 'Load Data Base' ) );
    # SKIP THESE SO THEY ARE NOT CONFUSED IN WAYPOINTS SECTION
    next if ( $f eq 'pre_email' );
    next if ( $f eq 'sender_email_addr' );
    next if ( $f eq 'diver-requested' );
    next if ( $f eq 'telepresence-requested' );
    next if ( $f eq 'broadcastTP-check' );
    next if ( $f eq 'asv-check' );
    next if ( $f eq 'ctd-check' );
    next if ( $f eq 'hazIso-check' );
    next if ( $f eq 'winch-check' );
    next if ( $f eq 'itar-check' );
    next if ( $f eq 'foreignNational-check' );
    next if ( $f eq 'specialOps-check' );
    next if ( $f eq 'permits-check' );
    next if ( $f eq 'storedDiverField' );
    next if ( $f eq 'storedTelepresenceField' );
    next if ( $f eq 'storedBroadcastTPField' );
    next if ( $f eq 'storedASVField' );
    next if ( $f eq 'storedCTDField' );
    next if ( $f eq 'storedWinchField' );
    next if ( $f eq 'storedHazIsoField');
    next if ( $f eq 'storedITARField');
    next if ( $f eq 'storedForeignNationalField');
    next if ( $f eq 'storedSpecialOpsField');
    next if ( $f eq 'storedPermitsField');
    next unless $f;
   # print "<br> waypoint=", $f, " order=", $Request->QueryString($f)->Item;
    $wptVisits{$Request->QueryString($f)->Item} = $f;
  }

  $wptString = "\n\nWaypoints to be visited\r\n";
  $wptString .="=======================\r\n";
  $wptString .= sprintf("%5s %-25s %11s %11s %7s\r\n", "Order", "Waypoint Name", "Latitude", "Longitude", "Depth");
  $wptString .= sprintf("%5s %-25s %11s %11s %7s\r\n", "-----", "-----------------------", "---------", "-----------", "-----");

  #
  # Open DB, Need separate QConn object as load() has one, and load() calls this
  #
  $QConn = CreateObject OLE "ADODB.Connection";
  $QConn->{'Provider'} = "sqloledb";
  $QConn->Open($dsn);

  Win32::ASP::AddDeathHook( sub { $QConn->Close } );  # In case asp dies do some cleanup

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
    next unless $wptVisits{$n};   # Getting some weird hash that we need to skip.
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
          "$lonDD $lonMM\r\n");
    $RS->Close;
  }
  $QConn->Close;

  $wptString = "" unless $gotWaypoints;
  ##$Response->Write("<pre>$wptString</pre>");
  return($wptString);
} # End procWaypoints()
%>


<%
#--------------------------------------------------------------------
#

=head3 email()

The initial USER email for original precruise request.
Process the form field entries and mail to the Logistics
coordinator.

Author: Mike McCann

Date Created: 11/5/98
Edited: Precruise email updates: 4/30/2025; Karen Salamy

=cut

sub email {

#
# Use libnet package to send
# mail message.
#
  use Net::SMTP;
  use Text::Wrap;
  #use strict;
  #use warnings;

    ################################
    #          HEADER              #
    ################################
    my $smtp = Net::SMTP->new('relay.mbari.org');     # connect to an SMTP server
    $smtp->mail( $Session->{'logistics_email'});       # use the sender's, for the email, address here


    ###############################
    #        ALIASING             #
    ###############################
    # Initialize an array for listserver recipient emails
    my @listserver_recipients;

    # Diver Requested Alias
    if (GetFormValue('diver-requested') eq "YES") {
        push @listserver_recipients, "divingpc\@listserver.mbari.org";
    }

    # Telepresence Requested Alias
    if (GetFormValue('telepresence-requested') eq "YES") {
        push @listserver_recipients, "telepresencepc\@listserver.mbari.org";
    }

    # Waveglider Requested Alias
    if (GetFormValue('asv-check') eq "YES") {
        push @listserver_recipients, "waveglider\@listserver.mbari.org";
    }

    # CTD Requested Alias
    if (GetFormValue('ctd-check') eq "YES") {
        push @listserver_recipients, "ctdpc\@listserver.mbari.org";
    }

    # Isotope and or HazMat Declared Alias
    if (GetFormValue('hazIso-check') eq "YES") {
        push @listserver_recipients, "isotopepc\@listserver.mbari.org";
    }

    # Winch Requested Alias
    if (GetFormValue('winch-check') eq "YES") {
        push @listserver_recipients, "winchpc\@listserver.mbari.org";
    }

    # ITAR Declared Alias
    if (GetFormValue('itar-check') eq "YES") {
        push @listserver_recipients, "itarpc\@listserver.mbari.org";
    }

    # Foreign Nationals Declared Alias
    if (GetFormValue('foreignNational-check') eq "YES") {
        push @listserver_recipients, "foreignpc\@listserver.mbari.org";
    }

    # Planned Track Description Alias
    if (GetFormValue('specialOps-check') eq "YES") {
        push @listserver_recipients, "specialopspc\@listserver.mbari.org";
    }

    # MBNMS Permit Declaration Alias
    if (GetFormValue('permits-check') eq "YES") {
        push @listserver_recipients, "permitpc\@listserver.mbari.org";
    }

    # Special Meal Requests Alias - Commented out, add logic if needed
    # if (GetFormValue('mealRequest-check') eq "YES") {
    #     push @listserver_recipients, "mealspc\@listserver.mbari.org";
    # }


    # Remove duplicate listserver emails
    #my %seen_listserver;
    #my @unique_listserver_recipients = grep { !$seen_listserver{$_}++ } @listserver_recipients;


    # Get the sender's email address and format properly

    my $sender_email = GetFormValue('sender_email_addr');
    my $formatted_sender_email; # Declare the variable

    if ($sender_email) {
        $formatted_sender_email = "$sender_email\@mbari.org";
    } else {
        $formatted_sender_email = undef;
    }


    # Get the logistics email address
    my $logistics_email = $Session->{'logistics_email'};



    ##################################
    #        Start the mail          #
    ##################################
    $smtp->to( $logistics_email ); # Primary recipient

    # Add other recipients to 'Cc:' if there are any
    if (@listserver_recipients || $formatted_sender_email) {
        my @cc_recipients;
        push @cc_recipients, @listserver_recipients;
        push @cc_recipients, $formatted_sender_email if $formatted_sender_email;
        $smtp->cc(@cc_recipients); # Use $smtp->cc for cleaner handling of multiple CCs
    }

    $smtp->data(); #Start the mail




    ################################
    #    SEND THE HEADER           #
    ################################
    $smtp->datasend("Message-ID: <" . time() . "@typhoon.shore.mbari.org>\n");
    $smtp->datasend("To: $logistics_email\n");
    $smtp->datasend("From: $logistics_email\n"); # Or use the actual 'user@mbari.org' if desired
    if (@listserver_recipients || $formatted_sender_email) {
        my @cc_header_values = @listserver_recipients;
        push @cc_header_values, $formatted_sender_email if $formatted_sender_email;
        $smtp->datasend("Cc: " . join(", ", @cc_header_values) . "\n");
    }
    $smtp->datasend("Subject: Precruise $Session->{'ScheduledStartDtg'} (Web form $this_script)\n");
    $smtp->datasend("\n");




    ################################
    #        SEND THE BODY         #
    ################################
    $smtp->datasend("Precruise database load request\n");
    $smtp->datasend("===============================\n\n");




    ################################
    #   ACCESS EXPEDITION FIELDS   #
    ################################
    ##$hyperlink = "$Application->{'BaseUrl'}/$this_script?" . $Session->{'Step2Args'};
    $hyperlink = "$Application->{'BaseUrl'}/$this_script?step=1";

    $smtp->datasend("MBARI Project Number: " . $Session->{'ProjNum'} . "\n");

    # Was 60
    $Text::Wrap::columns = 100;
    $smtp->datasend("\nLink to load this expedition into the database:\n" .
    $hyperlink . "\n\n(Copy from the fields below into the web form fields)\n\n");
    foreach $f ( @{$Application->{'ExpeditionFields'}} ) {
      next if $f eq 'StatCode';
      last if $f eq 'StartDtg';
      if ( $f =~ /ScheduledStartDtg$/ ) { # E-mail local time
          $str1 = "$Session->{'ScheduledStartDtgLocal'} Local Moss Landing time";
          $str2 = wrap("", "", $str1);
      }
      elsif ( $f =~ /ScheduledEndDtg$/ ) {  # E-mail local time
          $str1 = "$Session->{'ScheduledEndDtgLocal'} Local Moss Landing time";
          $str2 = wrap("", "", $str1);
      }
      elsif ( $f =~ /PlannedTrackDesc$/ ) { # Append waypoints
          $str1 = $Session->{'PlannedTrackDesc'};
          $str2 = wrap("", "", $str1);
          $str2 .= procWaypoints();
      }

      else {
        $str1 = $Session->{$f};
        $str2 = wrap("", "", $str1);
      }

      if ( $str2 =~ /\n/ ) {  # Put new line in front
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

} # End email()

%>

<%
#--------------------------------------------------------------------
#

=head3 precruise_mail_out()

This is the Email sent out by this script after a precruise has 
been Created OR Revised within the EXPD by the logistics coordinator.

Process the form field entries and mail to the address
specified, precruise@listserver.mbari.org is the default.

Author: Mike McCann

Date Created: 9/2/99
Edited: 3/7/2017; Karen Salamy
Edited: Precruise email updates: 5/7/2025; Karen Salamy

=cut

sub precruise_mail_out {

#
# Use libnet package to send
# mail message.
#
    use Net::SMTP;
    use Text::Wrap;

    ################################################################
    #            STATE IF THE EMAIL HAS BEEN REVISED               #
    ################################################################
    my $message = $_[0];    # Catches 'REVISED:' string
    my $is_revised = 0;
    if ($message =~ /^REVISED:/) {
      $is_revised = 1;
      }


    #########################################################
    #                       HEADER                          #
    # INITIALIZE SMTP CONNECTION WITH PROPER ERROR HANDLING #
    #########################################################

    my $smtp = Net::SMTP->new('relay.mbari.org') or do {
       $Response->Write("<h3>Error connecting to email server</h3>");
       return;
    };
    
    # Set mail sender (FROM)
    if (!$smtp->mail($Session->{'logistics_email'})) {
        $Response->Write("<h3>Error setting sender address</h3>");
        $smtp->quit;
        return;
    }


    ####################################################
    #                      ALIASING                    #
    #                BUILD RECIPIENT LISTS             #
    ####################################################

    # Initialize an array for listserver recipient emails
    my @listserver_recipients = ();

    # Add recipients based on form values
    ####################################################
    # Diver Requested Alias
    push @listserver_recipients, "divingpc\@listserver.mbari.org" if GetFormValue('diver-requested') eq "YES";
    
    # Telepresence Requested Alias
    push @listserver_recipients, "telepresencepc\@listserver.mbari.org" if GetFormValue('telepresence-requested') eq "YES";
    
    # Waveglider Requested Alias
    push @listserver_recipients, "waveglider\@listserver.mbari.org" if GetFormValue('asv-check') eq "YES";
    
    # CTD Requested Alias
    push @listserver_recipients, "ctdpc\@listserver.mbari.org" if GetFormValue('ctd-check') eq "YES";
    
    # Isotope and or HazMat Declared Alias
    push @listserver_recipients, "isotopepc\@listserver.mbari.org" if GetFormValue('hazIso-check') eq "YES";
    
    # Winch Requested Alias
    push @listserver_recipients, "winchpc\@listserver.mbari.org" if GetFormValue('winch-check') eq "YES";
    
    # ITAR Declared Alias
    push @listserver_recipients, "itarpc\@listserver.mbari.org" if GetFormValue('itar-check') eq "YES";
    
    # Foreign Nationals Declared Alias
    push @listserver_recipients, "foreignpc\@listserver.mbari.org" if GetFormValue('foreignNational-check') eq "YES";
    
    # Planned Track Description Alias
    push @listserver_recipients, "specialopspc\@listserver.mbari.org" if GetFormValue('specialOps-check') eq "YES";
    
    # MBNMS Permit Declaration Alias - Commented out because permitpc is always sent (part of precruise listserver)
    # push @listserver_recipients, "permitpc\@listserver.mbari.org" if GetFormValue('permits-check') eq "YES";
    
    # Special Meal Requests Alias - Commented out, add logic if needed
    # push @listserver_recipients, "mealspc\@listserver.mbari.org" if GetFormValue('mealRequest-check') eq "YES";
    # Uncomment the following line if you want to include the permits-check condition
   

    ####################################################
    #       CREATE REVISED LISTSERVER RECIPIENTS       #
    #            IF A REVISION, SEND ALL               #
    ####################################################
    # Initialize an array for ALL listserver recipient emails
    my @revised_listserver_recipients = (
        "divingpc\@listserver.mbari.org",
        "telepresencepc\@listserver.mbari.org",
        "waveglider\@listserver.mbari.org",
        "ctdpc\@listserver.mbari.org",
        "isotopepc\@listserver.mbari.org",
        "winchpc\@listserver.mbari.org",
        "itarpc\@listserver.mbari.org",
        "foreignpc\@listserver.mbari.org",
        "specialopspc\@listserver.mbari.org",
        ###"permitpc\@listserver.mbari.org",
        ###"mealspc\@listserver.mbari.org",
    );


    ####################################################
    #        Get the logistics email address           #
    ####################################################
    my $logistics_email = $Session->{'logistics_email'};


    ##################################################################
    # Get the precruise listserver email address and format properly #
    ##################################################################
    ###my $precruise_email = GetFormValue('pre_email');
    ###my $formatted_pre_email = $precruise_email ? "$precruise\@listserver.mbari.org" : undef;

    #############################################################
    #   Hardcode sendout to precruise\@listserver.mbari.org     #
    #############################################################
    my $formatted_pre_email = "precruise\@listserver.mbari.org";



    ######################################################
    # SET PRIMARY RECIPIENTS
    ######################################################
    
    # Always send to logistics coordinator
    my @to_recipients = ($logistics_email);
    
    # Add precruise listserver if specified
    push @to_recipients, $formatted_pre_email;
    
    # Set TO recipients
    if (!$smtp->to(@to_recipients)) {
        $Response->Write("<h3>Error setting primary recipients</h3>");
        $smtp->quit;
        return;
    }
    
    ######################################################
    #                   SET CC RECIPIENTS                #
    ######################################################
    # Use appropriate CC list based on whether this is a revision

    my @cc_recipients = $is_revised ? @revised_listserver_recipients : @listserver_recipients;
    
    # Set CC recipients if we have any
    if (@cc_recipients) {
        if (!$smtp->cc(@cc_recipients)) {
            $Response->Write("<h3>Warning: Could not add CC recipients</h3>");
            # Continue anyway with primary recipients
        }
    }

   ######################################################
   #        BEGIN DATA SECTION AND SEND EMAIL           #
   ######################################################
    
    # Start the data section
    if (!$smtp->data()) {
        $Response->Write("<h3>Error starting email data section</h3>");
        $smtp->quit;
        return;
    }

    ####################################################
    #                SEND THE HEADER                  #
    ####################################################
    $smtp->datasend("Message-ID: <" . time() . "@typhoon.shore.mbari.org>\n");
    
    # Send To: header with non-empty recipients
    my $to_header = join(", ", grep {$_} @to_recipients);
    $smtp->datasend("To: $to_header\n") if $to_header;
    
    # Send From: header
    $smtp->datasend("From: $Session->{'logistics_email'} (via Web form $this_script)\n");
    
    # Send Cc: header if there are any CC recipients
    if (@cc_recipients) {
        my $cc_header = join(", ", @cc_recipients);
        $smtp->datasend("Cc: $cc_header\n") if $cc_header;
    }
    
    # Send Subject: header
    my $subject = "Precruise: $Session->{'ExpdChiefScientist'} $Session->{'ScheduledStartDtg'} GMT";
    if ($message) {
        $subject = "$message $subject";
    }
    $smtp->datasend("Subject: $subject\n");
    $smtp->datasend("\n"); # Blank line to separate headers from body
    

    
    ####################################################
    #                  SEND THE BODY                  #
    ####################################################
    $smtp->datasend("$message\n") if $message;
    $smtp->datasend(sprintf("%-30s %s\n", "Precruise:", $Session->{'ExpdChiefScientist'}));




    ####################################################
    #   ACCESS EXPEDITION FIELDS AND START THE EMAIL   #
    ####################################################
    $hyperlink = "$Application->{'BaseUrl'}/$this_script?step=1";

    # Was 90
    $Text::Wrap::columns = 140;

    foreach my $f (@{$Application->{'ExpeditionFields'}}) {
        next if $f eq 'StatCode';  # Skip over this
        next if $f eq 'ExpeditionID';  # Skip over this
        last if $f eq 'StartDtg';  # Skip over this
        
        my $str1;
        if ($f =~ /ScheduledStartDtg$/) { 
            # Email local time
            $str1 = "$Session->{'ScheduledStartDtgLocal'} Local Moss Landing time";
        }
        elsif ($f =~ /ScheduledEndDtg$/) {  
            # Email local time
            $str1 = "$Session->{'ScheduledEndDtgLocal'} Local Moss Landing time";
        }
        else {
            $str1 = $Session->{$f};
        }
        
        # Skip empty fields
        next unless defined $str1 && $str1 ne '';
        
        $str1 =~ s/\n\s+/\n/g; # Remove leading spaces
        my $str2 = wrap("\t", "\t", $str1);
        
        if ($str2 =~ /\n/) {
            $smtp->datasend("$f:\n$str2\n");
        }
        else {
            my $labl = $f . ":";
            $smtp->datasend(sprintf("%-30s %s\n", $labl, $str1));
        }
    }



    ####################################################
    #               SEND TRAILING MESSAGE              #
    ####################################################
    my $trailmsg = "\n--------------------\nCruise participants:\n";

    if ($Session->{'ShipName'} eq 'dpkd') {
        $trailmsg .= "\nVisiting scientists who will participate on R/V David Packard cruises will " .
                    "need a TEMPORARY R/V David Packard BADGE for the duration of their cruise " .
                    "(after positive identification is established). Badges will not have to be " .
                    "worn offshore, but will be required while the vessel is docked. These badges " .
                    "will only be issued to personnel listed on the precruise information submitted " .
                    "by the chief scientist. The badges for external participants will be " .
                    "maintained and issued by Lisa Emanuelson. All badges are numbered and must " .
                    "be returned.\n\nAlso, before going to sea:\n";
    }
    
    $trailmsg .= "\nCall the video lab (831-775-1829) or enter a video lab help request " .
                "(https://mww.mbari.org/video/cgi-bin/ttx.pl) one week prior to your cruise to " .
                "schedule training on how to use our video file recordering system (m3rs) and " .
                "VARS program (used for framegrabbing and video annotations).\n\n" .
                "Please report all problems related to cameras, video recordings, and annotation " .
                "to the video lab staff IMMEDIATELY.\n\n" .
                "More info at: https://www.mbari.org/at-sea/cruise-planning/.\n\n" .
                "If you have any questions, please call 831-775-1829.\n";
    
    $smtp->datasend($trailmsg);
    
    # End data section and check for errors
    if (!$smtp->dataend()) {
        $Response->Write("<h3>Error completing email transmission</h3>");
        $smtp->quit;
        return;
    }
    
    # Close the connection
    $smtp->quit;
    
    # Success message
    %>
    <h3>Email has been sent to <%= $formatted_pre_email || $logistics_email %>.</h3>
    <%
} # End precruise_mail_out()
%>

<%
#--------------------------------------------------------------------
#

=head3 revise()

Present form for user to fill out to revise an exiting precruise record.

Author: Mike McCann

Date Created: 9/2/99
Edited: Precruise email updates: 4/30/2025; Karen Salamy

=cut

sub revise {

  require 'timelocal.pl';

  if ( $Session->{'DBadministrator'} != 1 ) {
  %>
    <h2>Precruises can be revised only by Logistics Coordinator</h2>
  <%  return;
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
# Construct option list for ships
#
$prgn_sel = ($RS->Fields('ShipName')->Value eq "prgn") ? "selected" : "";
$dpkd_sel = ($RS->Fields('ShipName')->Value eq "dpkd") ? "selected" : "";
$rcsn_sel = ($RS->Fields('ShipName')->Value eq "rcsn") ? "selected" : "";
$ship_option_list = "<option value=\"prgn\"${prgn_sel}>Paragon</option>\n";
$ship_option_list .= "<option value=\"rcsn\"${rcsn_sel}>Rachel Carson</option>\n";
$ship_option_list .= "<option value=\"dpkd\"${dpkd_sel}>David Packard</option>\n";


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

<script Language="JavaScript">
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
    /*alert("Please enter a value for the \"Planned Track Description\" field.");
    theForm.PlannedTrackDesc.focus();
    return (false);*/
  }
  if (theForm.Participants == "")
  {
    alert("Please enter a value for the \"MBARI Participants\" field.");
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


function ShipSelect(select)
{
  if (select.options[select.selectedIndex].value == "dpkd") {
    alert("For David Packard cruises please confirm Departure and Arrival times with the ship's captain before entering them here.");
    return (true);
  }
  if (select.options[select.selectedIndex].value == "rcsn") {
    alert("For Rachel Carson cruises please confirm Departure and Arrival times with the ship's captain before entering them here.");
    return (true);
  }
  if (select.options[select.selectedIndex].value == "prgn") {
    alert("For Paragon reservations please confirm that your captain is an approved operator.");
    return (true);
  }
  if (select.options[select.selectedIndex].value == "") {
    return (false);
  }
}
</script>

<%

%>
<form method="POST" action="<%= $this_script%>" onsubmit="return Field_Validator(this)" accept-charset="utf-8" name="Field">

  <table border="0" cellpadding="5" cellspacing="0">
    <tr>
      <td valign="top"><strong><font face="Arial,Helvetica">Ship Name</font></strong></td>
      <td>
        <select name="ShipName" size="1" onChange="return ShipSelect(this)">
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
          <a href="person.asp">Add/Edit person list</a>
        </td>
      </tr>
    </table>


    <p><strong><font face="Arial,Helvetica">Cruise Purpose
    </font></strong> (Indicate reason for revision here)<br>
    <textarea rows="3" name="Purpose" cols="120" wrap><%= $ExpeditionValue{Purpose}%></textarea></p>

    <p><strong><font face="Arial,Helvetica">Required Equipment Description</font></strong><br>
    <textarea rows="3" name="EquipmentDesc" cols="120" wrap><%= $ExpeditionValue{EquipmentDesc}%></textarea></p>

    <p><strong><font face="Arial,Helvetica">Participants</font></strong><br>
    <textarea rows="3" name="Participants" cols="120" wrap><%= $ExpeditionValue{Participants}%></textarea></p>

    <p><strong><font face="Arial,Helvetica">Planned Track Description</font></strong><br>
    <textarea rows="3" name="PlannedTrackDesc" cols="120" wrap><%= $ExpeditionValue{PlannedTrackDesc}%></textarea></p>

    <input type="hidden" name="step" value="update">
    <input type="hidden" name="ShipSeqNum" value="<%= GetFormValue(ShipSeqNum)%>">

    <font face="Helvetica,Ariel" size="-1">Send email to
    <input type="text" name="pre_email" value="precruise" size="9" maxlength="9">
    @listserver.mbari.org and </font>

    <p><input type="submit" class="button gray" value="Update database and send mail" name="update">
    <input type="reset" class="button gray" value="Reset form fields" name="Reset form"></p>
</form>
<%
$RS->Close;
$Conn->Close;

} # End revise()
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

  ##$sql = "SELECT DateDiff('s', '01/01/70', $time_field) AS Epoch ";   # -The way Access wants it
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

my $lstr = localtime(time); # Arghhh! get 2 digit yr from above!
my @lstr2 = split('\s+', $lstr);
my $NextYear = $lstr2[4] + 1;

my @l = localtime($Epoch);
my $day = $l[3];
my $mo = $l[4] + 1;
my$month = ($mo < 10) ? "0" . $mo : $mo;
my $lstr = localtime($Epoch); # Arghhh! get 2 digit yr from above!
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

} # End construct_option_lists()
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
Edited: Precruise email updates: 4/30/2025; Karen Salamy

=cut

sub update {

#
# Take the date/time fields and construct UTC DTGs & save as session vars
#
@dl = ( 0, substr(GetFormValue('dTime'),2,2), substr(GetFormValue('dTime'),0,2),
    GetFormValue('dDay'), GetFormValue('dMonth')-1, GetFormValue('dYear')-1900 );
@al = ( 0, substr(GetFormValue('aTime'),2,2), substr(GetFormValue('aTime'),0,2),
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



#########################################################################
#$Session->{'ShipName'} = GetFormValue('ShipName');
#$Session->{'ShipSeqNum'} = GetFormValue('ShipSeqNum');
#$Session->{'StatCode'} = 'plnd';

#$Session->{'ExpdChiefScientist'} = GetFormValue('ExpdChiefScientist');
#$Session->{'ExpdPrincipalInvestigator'} = GetFormValue('ExpdPrincipalInvestigator');

#$Session->{'Purpose'} = GetFormValue('allCruisePurposeInfo');
#$Session->{'PlannedTrackDesc'} = GetFormValue('allPlannedTrackInfo');
#$Session->{'Participants'} = GetFormValue('allParticipants');
#$Session->{'EquipmentDesc'} = GetFormValue('allEquip');
#########################################################################

#
# Open the database, this time to update fields
#
open_database($dsn);    # Creates $Conn object as a global variable

#
# Test where clause to see that it updates only one record.  Bail if it matches more than 1.
#
$sql = "SELECT count(*) AS cnt \nFROM Expedition\n";
$sql .= "WHERE ( StatCode != 'cmpl' AND ShipName = '";
$sql .= $Session->{ShipName} . "' AND ShipSeqNum = ";
$sql .= $Session->{ShipSeqNum} . ")";

Win32::ASP::DebugPrint("\nSQL = $sql ");

$RS = $Conn->Execute($sql);
if (!$RS) {
  $Response->Write("<h3>Precruise revision failed. The cruise was not saved.</h3>");
  $Response->Write("<p>The record-count check failed. Please contact database support.</p>");
  $Response->Write("Database Load Error(s): ");
  $Errors = $Conn->Errors();
  foreach $error (keys %$Errors) {
    $Response->Write($error->{Description});
  }
  $Conn->Close if $Conn;
  return 1;
}

Win32::ASP::DebugPrint("\ncnt = " . $RS->Fields(cnt)->Value );

if ( $RS->Fields(cnt)->Value != 1 ) {
  %><h2><%= $RS->Fields(cnt)->Value%> records would be updated!</h2>
  Only one record should be updated.  Please contact your database support to have this problem fixed.
  <%
  $RS->Close if $RS;
  $Conn->Close if $Conn;
  return 1;
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
if (!$nfields) {
  $Conn->Close if $Conn;
  return 1;
}
$sql =~ s/,$//;

$sql .= "WHERE ( StatCode != 'cmpl' AND ShipName = '";
$sql .= $Session->{ShipName} . "' AND ShipSeqNum = ";
$sql .= $Session->{ShipSeqNum} . ")";

Win32::ASP::DebugPrint("\nSQL = $sql ");

$RS = $Conn->Execute($sql);
if (!$RS) {
  $Response->Write("<h3>Precruise revision failed. The cruise was not saved.</h3>");
  $Response->Write("<p>The database update failed. Please contact database support and include the details below.</p>");
  $Response->Write("Database Load Error(s): ");
  $Errors = $Conn->Errors();
  foreach $error (keys %$Errors) {
    $Response->Write($error->{Description});
  }
  $Response->Write("<pre>");
  $Response->Write("Purpose length: " . length($Session->{'Purpose'}) . "\n");
  $Response->Write("EquipmentDesc length: " . length($Session->{'EquipmentDesc'}) . "\n");
  $Response->Write("Participants length: " . length($Session->{'Participants'}) . "\n");
  $Response->Write("PlannedTrackDesc length: " . length($Session->{'PlannedTrackDesc'}) . "\n");
  $Response->Write("</pre>");
  $Conn->Close if $Conn;
  return 1;
}

%>
<h3>Updated precruise <%= $Session->{'ShipName'}%><%= $Session->{'ShipSeqNum'}%></h3>
<%

$RS->Close;
$Conn->Close;
return 0;

} # End update()
%>

<%
#--------------------------------------------------------------------
#

=head3 cleanHTML()

Present entries from step1() for the user to view and check before submitting.

Author: Karen Salamy

Date Created: 6/13/2016

=cut

sub cleanHTML {

    $Session->{'Purpose'} = GetFormValue('allCruisePurposeInfo_cleanHTML');
    $Session->{'PlannedTrackDesc'} = GetFormValue('allPlannedTrackInfo_cleanHTML');
    $Session->{'Participants'} = GetFormValue('allParticipants_cleanHTML');
    $Session->{'EquipmentDesc'} = GetFormValue('allEquip_cleanHTML');
%>
     <script Language="JavaScript">
      function Field_Validator(theForm)
      {
        return(true);
      }
      </script>
<%

} # End cleanHTML()
%>
