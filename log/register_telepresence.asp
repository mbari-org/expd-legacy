<%@ LANGUAGE = PerlScript %>

<!--#include file="postcruise_hdr_telepresence.inc"-->

<% 
# -----------------------------------------------------------
# Register Active Server Page.
# Send E-Mail with requesting user's IP address to IAG group member
#
# Mike McCann MBARI
# November 1998
# -----------------------------------------------------------

# Warning: Don't use CGI, it freezes IIS when a new object is created.
use Win32::ASP;
use OLE;

$this_script="register_telepresence.asp";

init_AppVars();

if ( GetFormValue('LoginID') ) {
	%><h2 align="center">Application Accepted</h2>
		</td></tr></table><% 
		send_email();
}
else {
		%><h2 align="center">Register for Expedition Log Privileges</h2>
		</td></tr></table><% 
		present_form();
}
%>

<!--#include file="postcruise_ftr_telepresence.inc"-->


<%
#/*======================================================================
#	present_form()
#	
#	Description: 
#			Form entry fields for register request
#	     
#	
#	Author: Mike McCann
#	Date Created: 11/30/98
#====================================================================== */
sub present_form {


%>
	<script Language="JavaScript" ><!--
	function Field_Validator(theForm)
	{
	if (theForm.LoginID.value == "")
	{
	  alert("Please enter your MBARI Login ID name (same as your email address).");
	  theForm.LoginID.focus();
	  return (false);
	}
	return (true);
	}
	//-->
	</script>
	
<form method="get" action="<%= $this_script%>" onsubmit="return Field_Validator(this)" >
  <font face="helvetica,Ariel"><b>First Name: </b><input type="text" name="FirstName" size="20"></font>
  <p><font face="helvetica,Ariel"><b>Last Name: </b><input type="text" name="LastName" size="20"></font>
  <p><font face="helvetica,Ariel"><b>System login ID: </b><input type="text" name="LoginID" size="9"></font>
  <p><input type="submit" value="Submit" name="B1"> 
  <input type="reset" value="Reset"name="B2"></p>
</form>

<%
}	# End present_form()
%>

<%
#/*======================================================================
#	send_email()
#	
#	Description: 
#			Send email along with use's IP address
#	     
#	
#	Author: Mike McCann
#	Date Created: 11/30/98
#====================================================================== */
sub send_email {

	#
	# Use libnet package to send
	# mail message.
	#
    use Net::SMTP;
	use Text::Wrap;


    $smtp = Net::SMTP->new('mail.shore.mbari.org'); 	# connect to an SMTP server
	$smtp->mail( GetFormValue('LoginID') . "\@mbari.org" );     	# use the sender's address here
    $smtp->to( $Application->{'logistics_email'} );        # recipient's address
    $smtp->data();                      # Start the mail

    # Send the header.
    #
    $smtp->datasend("To: $Application->{'logistics_email'}\n");
    $smtp->datasend("Cc: " . GetFormValue('LoginID') . "\n");
    $smtp->datasend("From: " . GetFormValue('LoginID') . "\@mbari.org (". GetFormValue('FirstName') . " " . GetFormValue('LastName') . ")\n");
	$smtp->datasend("Subject: ExpdLog Register request\n");
    $smtp->datasend("\n");

    # Send the body.
    #
    $smtp->datasend("Request to register " . GetFormValue('FirstName') . " " . GetFormValue('LastName') . "\n");
	$smtp->datasend("as an Expedition Log user. \n\n");
	$smtp->datasend("System LoginID: " . GetFormValue('LoginID') . "\n" );
	$smtp->datasend("User's IP Address: "  . $Request->ServerVariables('REMOTE_ADDR')->{Item} . "\n" );
    $smtp->datasend("\n\n----------------\n\n");
    $smtp->datasend("Logistics coordinator, \n\nIf you approve this request please forward this message");
    $smtp->datasend("\nto mailto:mccann\@mbari.org.");
    
    $smtp->dataend();                   # Finish sending the mail
    $smtp->quit;                        # Close the SMTP connection
	%>

<h3>Email has been sent to <%= $Application->{'logistics_email'}%></h3>
You should receive a copy of this email at the address <%= GetFormValue('LoginID')%>@mbari.org.
<br>Please wait 24 hours for notification on your request being processed.

<%
}	# End email()
%>