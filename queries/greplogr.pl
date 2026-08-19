#
# Grep for lines in a logr file that is gzipped
#

# Don't use CGI, it freezes IIS when a new object is created.
use Compress::Zlib;
use LWP::Simple;


if ( $ARGV[0] &&  $ARGV[1] ) {
	
	print "<h5>Searching $ARGV[0] for $ARGV[1]</h5>\n";
	
	urlgrep($ARGV[0], $ARGV[1]);
}
else {
	
	print "<h3>Nothing to search</h3>\n";	
	
}





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
	my $string = $_[1];
	
	##my $dir = $Request->ServerVariables(APPL_PHYSICAL_PATH)->{Item};
	$dir = "..\\";
	##die "dir = $dir\n";
	##$dir =~ s/\\$//;
	$dir .= "temp";
	$dir = "";
	my $tmpfile = "logr$$.gz";
	
	print "<br>f= $tmpfile \n" ;
	##print "<br>Removing it..." ;
	
	##unlink "$dir\\$tmpfile" || die "Can't unlink $dir\\$tmpfile: $!";
	print "Exeucting getstore( $url, $tmpfile )...\n";
	$content = get($url);
	print $content;
	getstore( $url, $tmpfile ) || die "Can't open $url: $!";
	die "Should have written $tmpfile\n";
	$gz = gzopen("$dir\\$tmpfile", "rb")|| die "Can't open $dir\\$tmpfile: $!";

	
	print "<pre>";
	$foundNum = 0;
	while ($gz->gzreadline($_) > 0) {
		##print "<br>$_");
		if ( /$string/i ) {
			$foundNum = 1;
			print "<br>$_";
		}
	}
	if ( $foundNum ) {
		print "                                    ^      ^  \n";
		print "                               EpochSecs Dive#\n";
		print "                                \n";
		print "          (Time when dive number entered)(Possible number)\n";
	}
	else {
		print "String '$string' not found in $dir\\$tmpfile\n";
		print "Dive number probably not entered.\n";
	}
	print "</pre>";
	$gz->gzclose();
	
	system("delete $dir\\$tmpfile");
	unlink "$dir\\$tmpfile" || die "Can't unlink $dir\\$tmpfile: $!";
	
}	# urlgrep()