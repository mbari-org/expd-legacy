<%@ LANGUAGE = PerlScript %>


<% 
#
# Grep for lines in a logr file that is gzipped
#

# Don't use CGI, it freezes IIS when a new object is created.
use Win32::ASP;
use Compress::Zlib;
use LWP::Simple;


if ( GetFormValue('url') &&  GetFormValue('string') ) {
	 
	urlgrep(GetFormValue('url'), GetFormValue('string'));
}
else {
	%>
	<h3>Nothing to search</h3>	
	<% 
}
%>


<%

#/*======================================================================
#	logrgrep()
#	
#	Description: 
#	     	Search file for string.
#	     
#	
#	Author: Mike McCann
#	Date Created: 11/2/99
#====================================================================== */
sub urlgrep {

	my $url = $_[0];
	$url =~ s/mww/search/;		# LWP getstore() does not follow the redirection.
	my $string = $_[1];
	
	%>
	<h5>Searching <%=$url%> for <font color="blue"><%=$string%></font></h5>
	<%
	
	my $dir = $Request->ServerVariables(APPL_PHYSICAL_PATH)->{Item};
	##$dir =~ s/\\$//;
	$dir .= "temp";
	my $tmpfile = "logr$$.gz";
	
	##$Response->Write("<br>f= $dir\\$tmpfile" );
	##$Response->Write("<br>Removing it..." );
	
	unlink "$dir\\$tmpfile" || die "Can't unlink $dir\\$tmpfile: $!";
	getstore( $url, "$dir\\$tmpfile" ) || die "Can't open $url: $!";
	$gz = gzopen("$dir\\$tmpfile", "rb")|| die "Can't open $dir\\$tmpfile: $!";

	
	$Response->Write("<pre>");
	$foundNum = 0;
	while ($gz->gzreadline($_) > 0) {
		##$Response->Write("<br>$_");
		if ( /$string/i ) {
			$foundNum = 1;
			$Response->Write("<br>$_");
		}
	}
	if ( $foundNum ) {
		$Response->Write("                                    ^      ^  \n");
		$Response->Write("                               EpochSecs Dive#\n");
		$Response->Write("                                \n");
		$Response->Write("          (Time when dive number entered)(Possible number)\n");
	}
	else {
		$Response->Write("String '$string' not found in $dir\\$tmpfile\n");
		$Response->Write("Dive number probably not entered.\n");
	}
	$Response->Write("</pre>");
	$gz->gzclose();
	
	system("delete $dir\\$tmpfile");
	unlink "$dir\\$tmpfile" || die "Can't unlink $dir\\$tmpfile: $!";
	
}	# urlgrep()
%>