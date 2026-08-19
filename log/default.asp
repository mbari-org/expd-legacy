<%@ LANGUAGE = PerlScript %>

<!--#include file="expd_functions.inc"-->

<% 
=head1 NAME

default.asp - Default Active Server Page for Expedition Log

=head1 SYNOPSIS

    http://expd.mbari.org/expd/log/default.asp

=head1 DESCRIPTION

Default page for expedition log site.  Allows presentation
of info and collection of info for beta testers.

If you change the location of this application be sure to
edit the 'AppDir' application variable in this file.

Mike McCann MBARI
December 1998

=head1 FUNCTIONS

=cut


# Warning: Don't use CGI, it freezes IIS when a new object is created.
use Win32::ASP;
use OLE;
use Win32::OLE qw( in );

init_AppVars();

$DBAdministrator = $Session->{'DBAdministrator'};	# 1 if host can update DB, else 0


if ( GetFormValue('debug') ) {
	beta_intro();
	#$Response->Write("$Application->{'AppDir'}/postcruise.asp");
}
elsif ( GetFormValue('goto') =~ /Precruise/i ) {
	process_beta();
	$Response->Redirect("$Application->{'AppDir'}/precruise.asp");
}
elsif ( GetFormValue('goto') =~ /Postcruise/i ) {
	process_beta();
	my $url = "$Application->{'AppDir'}/postcruise.asp";
	$url .= "?logistics_email=" . $Session->{'logistics_email'};
	$url .= "&RegisteredUser=" . $Session->{'RegisteredUser'};
	$Response->Redirect($url);
}
elsif ( GetFormValue('goto') =~ /Advanced/i ) {
	process_beta(); %>
	Click to search for expedition:<br>
	<A href="<%=$Application->{'AppDir'}%>/postcruise.asp?step=1&search=advanced&logistics_email=<%=$Session->{'logistics_email'}%>">Search postcruise</a>
	<%
	##$Response->Redirect("$Application->{'AppDir'}/postcruise.asp");
}
else {
	release_intro();
}
%>



<%
#--------------------------------------------------------------------
#

=head3 beta_intro()

Let users set email address of the Logistics Coordinator
and be a registered user or not.
     

Author: Mike McCann

Date Created: 12/4/98

=cut

sub beta_intro {


%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html> 
<head>

<title>Expedition Log</title>
</head>
<body bgcolor="#FFFFFF">
<table border="0" cellpadding="5">
<tr>
<td valign="top" width="20%">
<a href="http://www.mbari.org">
<img src="http://www.mbari.org/images/mbarilogo-120_sh.gif" border="0" width="120" height="70"></a>
</td>
<td valign="top" width="80%"><h2 align="center">Welcome to Expedition Log on the Web</h2>
<h2 align="center">Database entry - Beta 3</h2>
</td></tr></table>


<font face="Arial,Helvetica">
<%
#foreach my $env (in ($Request->ServerVariables)) { 
#	$Response->Write("<br><b>$env:</b> " . $Request->ServerVariables($env)->{Item} ); 
#}
%>



<p align="left">These pages are currently in beta test.  From this site you
can enter Precruise information and your Postcruise reports.  You may also
search for information on any MBARI shipboard expedition that we have in
our database.</p>
<p>
<p>With this beta test version we are working with a copy of the actual
database. Please feel free to make whatever entries you want in order to
test the the effectiveness of the system.  Whatever changes are made will
not affect the production database.</p>
<font size="-1">
The site has 2 modes:
<ul>
<li>Unregistered user: All updates are sent as an e-mail message to the 
MBARI Logistics Coordinator where she will enter them into the database.
<li>Registered user: All Postcruise web form entries are directly entered into the 
database.  
</ul> 
</font>

</font>
<form method="GET" action="default.asp">
  <table border="0" cellpadding="5" cellspacing="0">
    
    <tr>
	  <td valign="top"><font face="Arial,Helvetica" color="brown">
	  <b>For testing purposes you can set these options >>>></b></font>
	  </td>
	  
      <td valign="top" align="right" nowrap><strong>
      <font face="Arial,Helvetica">Logistics e-mail</font></strong></td>
      
      <td valign="top"><input type="text" name="LogisticsEmail"><br>
      <font face="Arial,Helvetica" size="-1"> 
      (Enter your e-mail address to receive the e-mail that is sent if the 
      box to the right is not checked. The default address is currently
      <%= $Application->{'logistics_email'}%>) </font>
      
	  </td>
      
      <td valign="top" align="right"><font face="Arial,Helvetica" nowrap><strong>Behave as a registered user
      </strong></font></strong></td>
      <td valign="top"><input type="checkbox" name="RegisteredUser" value="yes" checked>
      <font face="Arial,Helvetica" size="-1"> 
      (Leave checked to see your postcruise updates go directly to the database) </font></td>
      <td valign="top" align="right"><font face="Arial,Helvetica" nowrap><strong>Behave as DB Administrator
      </strong></font></strong></td>
      <td valign="top"><input type="checkbox" name="DBadministrator" value="yes">
      <font face="Arial,Helvetica" size="-1"> 
      (Check to see your precruise updates go directly to the database) </font></td>
    </tr>
    
    
    <tr>
    <td colspan="2">
    <u>Important links</u>
	  <ul>
	  <li><a href="http://www.mbari.org/dmo/cruise_planning/cruise.htm">
	  Cruise planning & scheduling</a>
	  <li><a href="/itd/video/shipprocedures.htm">
	  Shipboard video procedures</a>
	  </ul>
    </td>
    <td colspan="3">
		<center>
		<input type="submit" name="goto" value="Precruise entry">
		<p>
		<input type="submit" name="goto" value="Postcruise search for expeditions">
		</center>
    </td>
    </tr> 

  </table>
  
</form>
<font face="Arial,Helvetica" size="-1">We have also not yet implemented
the hyperlinks to data (e.g. Frame Grabs, Nav, ROVCTD, Samples) 
associated with each dive. Also, graphical tools for displaying and editing 
regions and waypoints is needed. These web pages demonstrate that the institutional
database structure is in place and that we can add more capability as time goes on. 
</font>
<hr>
<font size="-2">
Precruise web database entry development team: 
Holly Baum, Lisa deQuatro, Gerry Hatcher, Mike McCann, Debbie Meyer, Rich Schramm, Dan Wilkin.
</font>
</body>
</html>
<%
}
%>


<%
#--------------------------------------------------------------------
#

=head3 process_beta()

Set session variables for beta testing purposes.

Author: Mike McCann

Date Created: 12/7/98

=cut

sub process_beta {

	#
	# Abaondon session if user comes to this page...
	#
	$Session->Abandon();
	
	#
	# Who gets logistics e-mail
	#
	if ( GetFormValue(LogisticsEmail) ) {
		$Session->{'logistics_email'} = GetFormValue(LogisticsEmail);
		##<br><font color="red">Beta testing with logistics_email set to = </font>
	}
	else {
		$Session->{'logistics_email'} = $Application->{'logistics_email'};
	}

	#
	# Force registered user
	#
	if ( GetFormValue(RegisteredUser) eq 'yes' ) {
		$Session->{'RegisteredUser'} = 1 ;
		##<br><font color="red">Behaving as a registered user - your updates will not affect the production database.</font>
	}
	else {
		$Session->{'RegisteredUser'} = 0;
	}
	
	#
	# Force DB Administrator (Can enter precruises)
	#
	if ( GetFormValue(DBadministrator) eq 'yes' ) {
		$Session->{'DBadministrator'} = 1 ; 
		##<br><font color="red">Behaving as a DB administrator</font>
	}
	else {
		$Session->{'DBadministrator'} = 0;
	}

} # End process_beta()
%>


<%
#--------------------------------------------------------------------
#

=head3 release_intro()

Actual introduction and Administrative page for the site.

Author: Mike McCann

Date Created: 9/24/99

=cut

sub release_intro {


%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html> 
<head>

<title>Expedition Log</title>
</head>
<body bgcolor="#FFFFFF">
<table border="0" cellpadding="5">
<tr>
<td valign="top" width="20%">
<a href="http://www.mbari.org">
<img src="http://www.mbari.org/images/mbarilogo-120_sh.gif" border="0" width="120" height="70"></a>
</td>



 <td valign="top"><h2 align="center">Expedition Database Administrative Page</h2>
    <h2 align="center">Release 1.2</h2>
</td></tr></table>
      <div align="left">
      
      <table border="3" cellpadding="5" cellspacing="0" height="267">
        <tr>
          <td valign="top" align="left"><strong><a href="postcruise.asp">Search Expeditions </a></strong></td>
          <td valign="top" align="left">Search for upcoming, recent, and all expeditions over the
          years. Enter precruise information and postcruise reports. Find data and images from any
          MBARI shipboard expedition. </td>
        </tr>
        <tr>
          <td valign="top" align="left"><a href="precruise.asp"><strong>Precruise entry</strong></a></td>
          <td valign="top" align="left">Enter new precruise information. Form will be sent to
          Logistics Coordinator to be entered into the database (direct entry restricted to
          Logistics Coordinator). </td>
        </tr>
        <tr>
          <td valign="top" align="left"><strong><a href="register.asp">Registration request</a></strong></td>
          <td valign="top" align="left">Request permission to edit database directly as a power
          user. </td>
        </tr>
        <tr>
          <td valign="top" align="left"><strong><a href="person.asp">Person edit</a></strong></td>
          <td valign="top" align="left">Add person as Chief Scientist or PI to database (restricted
          to power users).</td>
        </tr>
        <tr>
          <td valign="top" align="left"><strong><a href="region.asp">Region edit</a></strong></td>
          <td valign="top" align="left">Add waypoint or region to database (restricted to power
          users).<font size="3"> We have not yet implemented graphical tools for displaying and
          editing regions and waypoints. Please stay tuned. </font></td>
        </tr>
      </table>
      </div><font size="-1" face="Arial,Helvetica"><p></font><font size="3"><strong>The
      expedition log site has 2 user modes: </strong><ul>
        <li><strong>Unregistered user:</strong> All updates are sent as an e-mail message to the
          MBARI Logistics Coordinator who will enter them into the database. </font>It is
          recommended that most users remain unregistered.<font size="3"></li>
        <li><strong>Registered user:</strong> All Postcruise web form entries are directly entered
          into the database. <a href="register.asp">Request permission to edit database directly</a>.
          </font></li>
      </ul>
      <p><em>If you are operating as a registered power user, changes made via these web forms
      will affect the production database immediately.</em> You will know that you are a
      registered user if the buttons at the bottom of the post cruise forms say things like
      &quot;Finish&quot; and &quot;Add another dive&quot;. Unregistered users will have a button
      saying &quot;Submit to Logistics Coordinator&quot;. </p>


<p>We have also not yet implemented
graphical tools for displaying and editing 
regions and waypoints.  Please stay tuned.</p>

<p><a href="/expd/docs">Expedition log and ROVCTD Documentation</a></p>
    </tr> 

  </table>
  

Last updated: 27 October 1999, Mike McCann
<br>
<hr>
<font size="-2">
Precruise web database entry development team: 
Holly Baum, Gerry Hatcher, Mike McCann, Debbie Meyer, Rich Schramm, Dan Wilkin.
</font>
</body>
</html>
<%
}
%>