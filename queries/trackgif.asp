<%@ LANGUAGE = PerlScript %>

<% 
	# Create gif image of ship archive track on the fly.  
	# Requires ActiveScript Perl and the GD module.
	# GD doc available at http://stein.cshl.org/WWW/software/GD/GD.html
	#
	#										Mike McCann
	#										MBARI  7 June 2000
    use GD;
    use OLE;
    use Win32::ASP;
    use Win32::OLE qw( in );		# Needed to iterate through list
    
    $Response->{'Expires'} = 0;
    $Response->{'ContentType'} = 'image/GIF';
    
    #
	# Set image parameters 
	#
	$Maxlat = 37.04672;
	$Minlat = 36.43817;
	$Maxlon = -121.75527;
	$Minlon = -122.55204;
	$xpix = 167;			# 1/5 od 768-wide image
	$ypix = 154;			#
	
	#
    # Start with thumbnail MB image
    #
    $baseMap = $Server->mappath("mbaflyrec154.png");
	open (PNG, $baseMap) || die "Can't open $baseMap";
    $myImage = newFromPng GD::Image(PNG) || die "Can't create myImage";
   
    # allocate some colors
    $red = $myImage->colorAllocate(255,0,0);
    $white = $myImage->colorAllocate(255,255,255);
 
	#
	# Open nav file & get lat/lon
	#
	$dir = ".";
	$pltfrm = GetFormValue('pltfrm');
	$YYYYDDD = GetFormValue('YYYYDDD');
	$YYYY = substr($YYYYDDD, 0, 4);
	##$file = "nav". $pltfrm . $YYYYDDD . ".txt";
	##$file = $Server->mappath('nav1999075ptlo.txt');
	$file = "\\tornado\\TempNav\\${YYYY}\\${pltfrm}\\nav${YYYYDDD}${pltfrm}.txt";
	open(NAV, "<$file") || die "Can't open $file: $!\n";
	$i = 0;
	while( <NAV> ) {
		next if /[a-zA-Z]/;
		$i++;
		next unless $i % 60 == 0;
		(undef, undef, undef, $esecs, $LatD, $LonD, 
			undef, undef, undef, $head, undef, undef, $roll) = split(',',$_);
		
		if ($LonD > $Maxlon || $LonD < $Minlon) {
		   next;
		}
		if ($LatD > $Maxlat || $LatD < $Minlat) {
		   next;
		}
		
		#
		# Translate the lat lon to pixels
		#
		$DX=($Maxlon - $Minlon )/$xpix;
		$DY=($Maxlat - $Minlat )/$ypix;
		$xcoord=($LonD-$Minlon)/$DX;
		$ycoord=($LatD-$Minlat)/$DY;
		
		$myImage->arc( $xcoord, $ypix-$ycoord, 2, 2, 0, 360, $red);
	}

	#
	# Put pltfrmYYYYDDD in upper right
	#
	$myImage->string(gdMediumBoldFont, 90, 0, "$pltfrm$YYYYDDD", $white);
	
	##$Response->Write("<!DOCTYPE HTML PUBLIC \"-//IETF//DTD HTML//EN\">");
	
	binmode STDOUT;	
	##print STDOUT $myImage->gif;
	$gifdata = $myImage->gif;
	$Response->BinaryWrite($gifdata);
	##
	##print "Length = ", length($gifdata);
	##syswrite STDOUT, $gifdata, length($gifdata);
	##print $gifdata;
	##print "\nfoobar\n";
	##close STDOUT;
	
	# Need absolute (Fully Qualified) path to the gif file, get path from a server variable
    # make sure we are writing to a binary stream
    $dir = $Request->ServerVariables(APPL_PHYSICAL_PATH)->{Item};
    $gfile = 'where_is_lobos.gif';
    $fq_gfile = "$dir\\$gfile";
    ##$Response->Write("<p>Writing to $fq_gfile");
    ##unlink "$fq_file";			# Remove the file before writing to make sure we're always getting a new image
    open (OUT, ">$fq_gfile") || die "Can't open $fq_gfile";
    binmode OUT;

    # Convert the image to GIF and print it to a temporary file

	print OUT $myImage->gif;

	close OUT;
	
	$Response->Redirect("$gfile");
	
	##binmode IN;
	##open(IN, "<$fq_gfile") || die "Can't open $fq_gfile";
	##while ( <IN> ) {
	##	print;
	##}
	##close IN;
%>
