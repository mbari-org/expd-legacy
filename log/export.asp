<%@ LANGUAGE = PerlScript %>

<% 

# -----------------------------------------------------------
# Export Regions as rgn files.
# Export Waypoints as wpt files. 
#
# Dan Wilkin MBARI
# November 1998
# -----------------------------------------------------------

# -----------------------------------------------------------
##  To Do:
##
##  Eliminate duplicates in the waypoint list, if regions overlap
##
# -----------------------------------------------------------

# Warning: Don't use CGI, it freezes IIS when a new object is created.
use Win32::ASP;
use OLE;
use Time::Local;
use POSIX;

$this_script = 'export.asp';
$dsn = $Application->{'DSN'};

$Response->{Expires} = 0;

if ( GetFormValue('export') eq 'rgn') {
	export_regions();
}
elsif ( GetFormValue('export') eq 'wpt') {
	export_waypoints();
}
else {
	%><h2 align="center">You have not choosen to create Region or Waypoint files.</h2><% 
}
%>




<%
#/*======================================================================
#	open_database()
#	
#	Description: 
#	     Create the database connection object.  Receive the DSN.
#	     Does a little error checking. $Conn is global.
#	
#	Author: Mike McCann
#	Date Created: 11/3/98
#====================================================================== */
sub open_database {

my $dsn = $_[0];

$Conn = CreateObject OLE "ADODB.Connection";
$Conn->{'Provider'} = "sqloledb";
$Conn->Open($dsn);
Win32::ASP::AddDeathHook( sub { $Conn->Close } );	# In case asp dies do some cleanup

if(!Conn) {
        $Errors = $Conn->Errors();
        print "<B>Error in opening connection to $dsn:</B>";
        foreach $error (keys %$Errors) {
                print $error->{Description}, "\n";
        }
        die "<BR>Empty Connection object.";
}

}	# End open_database()
%>




<%
#/*======================================================================
#	FixString()
#	
#	Description: 
#	     	Put quotes around text values and nothing around numbers
#		escape embedded quotes.
#	     
#	
#	Author: Mike McCann
#	Date Created: 10/28/98
#====================================================================== */
sub FixString {
	
#
# String, type
#

local ($str, $type) = @_;

if ($type =~ /text/i ) {
	$str =~ s/\'/\'\'/g;
	$fixed_string = "'$str'";
}

if ($type =~ /num/i || $str eq 'NULL' ) {
	$fixed_string = "$str";
}

if ($type =~ /num/i && $str eq '' ) {
	$fixed_string = "NULL";
}


return $fixed_string;

}	# End FixString()
%>






<%
#/*======================================================================
#	FixOutput()
#	
#	Description: 
#		Format output for Winfrog.
#		Escape embedded quotes in strings.
#	
#	Author: Dan Wilkin
#	Date Created: 11/25/98
#====================================================================== */
sub FixOutput {
	
#
# String, type
#

local ($str, $type) = @_;

# handle strings

if ($type =~ /text/i ) {
	$str =~ s/\'/\'\'/g;
	$fixed_output = "$str";
}

# handle numbers

if ($type =~ /num/i || $str eq 'NULL' ) {
	$fixed_output = "$str";
}

if ($type =~ /num/i && $str eq '' ) {
	$fixed_output = "NULL";
}

# handle lats and longs

$is_neg = 0;
$num = 0;

if ($type =~ /lat/i || $type =~ /long/i)
{
    # if the number is negative, chop off the sign
    if ($str < 0)
    {
        $is_neg = 1;
        $rev = reverse $str;
        chop $rev;
        $num = reverse $rev;
	}
	else {$num = $str};
	
    @num_fields = split /\./, $num;

    $dec_min = "0." . $num_fields[1];
	$dec_min = $dec_min * 60;

	$num = $num_fields[0] . " " . $dec_min;
	
	if ($type =~ /lat/i)
	{
	
		if ($is_neg)
		{
			$fixed_output = "S" . "$num";
		}
		else
		{
			$fixed_output = "N" . "$num";
		}
	
	}
	else
	{
		if ($is_neg)
		{
			$fixed_output = "W" . "$num";
		}
		else
		{
			$fixed_output = "E" . "$num";
		}
	}
}

return $fixed_output;

}	# End FixOutput()
%>






<%
#/*======================================================================
#	URLEncode()
#	
#	Description: 
#	     	Make variables safe for a URL
#	     
#	
#	Author: http://eclectic.kluge.net/codesnippets/perl/URLEncode.html
#	Date Created: 10/28/98
#====================================================================== */

sub URLEncode
{
    my($url)=@_;
    my(@characters)=split(/(\%[0-9a-fA-F]{2})/,$url);

    foreach(@characters)
    {
        if ( /\%[0-9a-fA-F]{2}/ ) # Escaped character set ...
        {
            # IF it is in the range of 0x00-0x20 or 0x7f-0xff
            #    or it is one of  "<", ">", """, "#", "%",
            #                     ";", "/", "?", ":", "@", "=" or "&"
            # THEN preserve its encoding
            #"
            unless ( /(20|7f|[0189a-fA-F][0-9a-fA-F])/i
                    || /2[2356fF]|3[a-fA-F]|40/i )
            {
                s/\%([2-7][0-9a-fA-F])/sprintf "%c",hex($1)/e;
            }
        }
        else # Other stuff
        {
            # 0x00-0x20, 0x7f-0xff, <, >, and " ... "
            s/([\000-\040\177-\377\074\076\042])
             /sprintf "%%%02x",unpack("C",$1)/egx;
        }
    }
    return join("",@characters);
}
%>





<%
#/*======================================================================
#	export_regions()
#	
#	Description: 
#	     	Gets the info for each of the regions the user has selected
#	     
#	
#	Author: Dan Wilkin
#	Date Created: 11/18/98
#====================================================================== */

sub export_regions
{

$rgnSQL="SELECT * FROM Region WHERE ";

@regions = split /,/, $Session->{'RegionDesc'};

foreach $rgn (@regions)
{
	$rgnSQL .= "RegionName=";
	$rgnSQL .= FixString($rgn, 'text');
	$rgnSQL .= " OR ";
}

# chop off the last "OR ", but leave the " "
chop $rgnSQL;
chop $rgnSQL;
chop $rgnSQL;

open_database($dsn);		# Creates $Conn object as a global variable

$rgnSQL .= "ORDER BY RegionName";

Win32::ASP::DebugPrint("\nExecuting SQL: $rgnSQL");

$rgnRS = $Conn->Execute($rgnSQL);

if(!$rgnRS)
{
        $Errors = $Conn->Errors();
        print "Errors:\n";
       	foreach $error (keys %$Errors)
	{
   		print $error->{Description}, "\n";
       	}	
 	die "<BR>Empty Return Set for SELECT \* FROM Region.";
}

$rgnRS->MoveFirst;

# create a list of the fields we want
@fields = ('RegionName', 'MaxLatitude', 'MinLatitude',
	   'MaxLongitude', 'MinLongitude', 'RegionOwner', 'Comment');

while ( !$rgnRS->EOF )
{
	foreach $f (@fields)
	{ 
		$val = ($rgnRS->Fields($f)->value) ? $rgnRS->Fields($f)->value : '';
		%><%= $val%>
		<%if (!($f eq 'Comment')) {%>,<%}	# print a comma if this is not the last field
	}
        %><br><%
	$rgnRS->MoveNext;
}

$rgnRS->Close;
$Conn->Close;

}  # end export_regions
%>






<%
#/*======================================================================
#	export_waypoints()
#	
#	Description: 
#	     	Gets the info for each of the regions the user has selected
#	     
#	
#	Author: Dan Wilkin
#	Date Created: 11/18/98
#====================================================================== */

sub export_waypoints
{

$Response->{ContentType} = "text/WPT";

$rgnSQL="SELECT * FROM Region WHERE ";

@regions = split /,/, $Session->{'RegionDesc'};

foreach $rgn (@regions)
{
	$rgnSQL .= "RegionName=";
	$rgnSQL .= FixString($rgn, 'text');
	$rgnSQL .= " OR ";
}

# chop off the last "OR ", but leave the " "
chop $rgnSQL;
chop $rgnSQL;
chop $rgnSQL;

open_database($dsn);		# Creates $Conn object as a global variable

$rgnSQL .= "ORDER BY RegionName";

##### no debugging!! Win32::ASP::DebugPrint("\nExecuting SQL: $rgnSQL");

$rgnRS = $Conn->Execute($rgnSQL);

if(!$rgnRS)
{
        $Errors = $Conn->Errors();
        print "Errors:\n";
       	foreach $error (keys %$Errors)
	{
   		print $error->{Description}, "\n";
       	}	
 	die "<BR>Empty Return Set for SELECT \* FROM Region.";
}

$rgnRS->MoveFirst;

# create a list of the fields we want
@fields = ('RegionName', 'MaxLatitude', 'MinLatitude',
	   'MaxLongitude', 'MinLongitude', 'RegionOwner', 'Comment');

# start the waypoint SQL statement
$wptSQL="SELECT DISTINCT WaypointName, Latitude, Longitude, Icon, ";
$wptSQL .= "Radius, Northing, Easting, Color, Elevation FROM Waypoint WHERE ";

# loop through the region list and add to the waypoint SQL statement

while ( !$rgnRS->EOF )
{
    # get all of the waypoints for the current region

	$wptSQL .= "((Latitude BETWEEN ";

	$wptSQL .= FixString($rgnRS->Fields('MinLatitude')->value, 'num');
	$wptSQL .= " AND ";
	$wptSQL .= FixString($rgnRS->Fields('MaxLatitude')->value, 'num');
		
	$wptSQL .= ") AND ((Longitude BETWEEN ";

	$wptSQL .= FixString($rgnRS->Fields('MinLongitude')->value, 'num');
	$wptSQL .= " AND ";
	$wptSQL .= FixString($rgnRS->Fields('MaxLongitude')->value, 'num');

	$wptSQL .= ") OR (Longitude BETWEEN ";

	$wptSQL .= FixString($rgnRS->Fields('MaxLongitude')->value, 'num');
	$wptSQL .= " AND ";
	$wptSQL .= FixString($rgnRS->Fields('MinLongitude')->value, 'num');
	
	$wptSQL .= "))) OR ";

    	# advance to the next region
	$rgnRS->MoveNext;

}  ## end while ( !$rgnRS->EOF )

# chop off the last "OR ", but leave the " "
chop $wptSQL;
chop $wptSQL;
chop $wptSQL;

# end the waypoint SQL statement
$wptSQL .= "ORDER BY WaypointName";

# print the Waypoint selection SQL statement for debugging
##### no debugging!! Win32::ASP::DebugPrint("\nExecuting SQL: $wptSQL");

# execute the wpt SQL
$wptRS = $Conn->Execute($wptSQL);

# check for errors
if(!$wptRS)
{
       	$Errors = $Conn->Errors();
       	print "Errors:\n";
	foreach $error (keys %$Errors)
	{
		print $error->{Description}, "\n";
       	}	
 	die "<BR>Empty Return Set for selecting waypoints from Waypoint table.";
}

# print the list of waypoints for the wpt file

$wptRS->MoveFirst;

## # create a list of the fields we want
## @fields = ('WaypointName', 'Latitude', 'Longitude', 'Icon',
##	      'Radius', 'Northing', 'Easting', 'Color', 'Elevation');

while ( !$wptRS->EOF )
{
	$Response->Write(FixOutput($wptRS->Fields('WaypointName')->value, 'text'));
	$Response->Write(',');
	$Response->Write(FixOutput($wptRS->Fields('Latitude')->value, 'lat'));
	$Response->Write(',');
	$Response->Write(FixOutput($wptRS->Fields('Longitude')->value, 'long'));
	$Response->Write(',');
	$Response->Write(FixOutput($wptRS->Fields('Icon')->value, 'num'));
	$Response->Write(',');
	$Response->Write(FixOutput($wptRS->Fields('Radius')->value, 'num'));
	$Response->Write(',');
	$Response->Write(FixOutput($wptRS->Fields('Northing')->value, 'num'));
	$Response->Write(',');
	$Response->Write(FixOutput($wptRS->Fields('Easting')->value, 'num'));
	$Response->Write(',');
	$Response->Write(FixOutput($wptRS->Fields('Color')->value, 'num'));
	$Response->Write(',');
	$Response->Write(FixOutput($wptRS->Fields('Elevation')->value, 'num'));

###!!!  !!!
# amazingly this works...it adds a carriage return (or newline), at least for Notepad
$Response->Write('
');
###!!!  !!!

	$wptRS->MoveNext;
}

$wptRS->Close;
$rgnRS->Close;
$Conn->Close;

}  # end export_waypoints
%>