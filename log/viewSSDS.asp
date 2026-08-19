<%@ LANGUAGE = PerlScript %>

<!--#include file="expd_functions.inc"-->

<% 
=head1 NAME

viewSSDS.asp - Present page of SSDS DataProducers, DataContainers, and Resources

=head1 SYNOPSIS

    http://mww.mbari.org/expd/log/viewSSDS.asp:ProcessRunID=1234

=head1 DESCRIPTION

View SSDS data Active Server Page.
Show ProcessRun, its inputs, outputs and associated Resources

Mike McCann MBARI

September 2005

=head1 FUNCTIONS

=cut

BEGIN { unshift(@INC,  'D:\\Inetpub\\wwwroot-nfp\\expdlog\\expd\\log') };


# Warning: Don't use CGI, it freezes IIS when a new object is created.

use Win32::ASP;
use OLE;
use Time::Local;
use POSIX;

use SSDS;

$this_script = 'viewSSDS.asp';
$debug = 0;

$ssds = new SSDS();
$server = "http://polyp:8080/";
$ssds->ssdsServer($server);
if ( $ssds->ssdsServer() ) {
	print "Running tests using server: " . $ssds->ssdsServer() . "\n" if $debug;
}
else {
	die "Cannot get response from server $server.  Perhaps it is temporarily down (?)\n";
}
             

if ( GetFormValue('ProcessRunID') ) {
	my $prAccess = new SSDS::ProcessRunAccess();
        my $pr = $prAccess->findByPK(GetFormValue('ProcessRunID'));
        %>
        <%=$pr->id()%>: <%=$pr->name()%>
        <%
        
	
}
else {
	%><h2 align="center">Unknown step <%= GetFormValue('step')%></h2>
	</td></tr></table><% 
}
%>


