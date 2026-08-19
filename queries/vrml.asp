<%@ LANGUAGE = PerlScript %>

<%

=head1 NAME

vrml.asp - VRML Active Server Page

=head1 SYNOPSIS

    > http://expd.mbari.org/expd/queries/vrml.asp?ShipName=ptlo&YYYYDDD=1999075

=head1 DESCRIPTION

This script generates a of VRML response given a set of options chosen by the user.
It basically combines a set of already created base vrml files using the inline node.

Possibilities include:

	- ship track
	- ROV Navigation
	- CTDO
	- Frame grabbed images
	- annotations from VIMS
	- samples from samples database

=cut

use Win32::ASP;
use VRML::VRML2;

$ShipName = GetFormValue('ShipName');
$YYYYDDD = GetFormValue('YYYYDDD');
$this_script = 'vrml.asp';

#
# Some font strings
#
$font_2 = '<font face="Helvetica,Ariel" size="-2">';
$font_1 = '<font face="Helvetica,Ariel" size="-1">';

if (GetFormValue('genVRML') ){
	genVRML();
}
else {
	entry_page();
}

%>


<%
#--------------------------------------------------------------------
#

=head3 entry_page()

First page a user comes to from (e.g.) the Expedition search pages.  Instruct user
where to get plugin and present options for VRML generation.  Present only the options
that are available (a base VRML file exists).

=cut

sub entry_page {
%>
<html>
<head>
<title>3D Dive Replay for <%= $ShipName%><%= $YYYYDDD%></title>
</head>
<body bgcolor="#808080" text="#FFFFFF" link="#FFFF00" vlink="#FF0000" alink="#00FF00">
<h2> 3D Dive Replay for <%= $ShipName%><%= $YYYYDDD%></h2>

<form method="GET" action="<%=$this_script%>">
<input type="hidden" name="ShipName" value="<%=$ShipName%>">
<input type="hidden" name="YYYYDDD" value="<%=$YYYYDDD%>">

<table border="0" cellpadding="5">
<tr><td><font face="Helvetica,Ariel" size="-1"><b>
Select features to be included in the 3D display, then click on the button below.  
(The image at right is a sample visualization that you can make from this tool.)</b>
</font></td>
<td bgcolor="#404040" rowspan="9" width="427"><img src="vrml_intro.jpg" WIDTH="427" HEIGHT="406">
<%=$font_2%>
Sample screen shot from R/V <i>Point Lobos</i> expedition on 1999273 (Ventana dive #1684).
Terrain is shown with 25% transparency, notice how you can see the ROV model and dive track
go below the surface. The vertical bar on the left is a time-control slider.  You can click 
and drag the blue ball to control the positions of the ship and the ROV accordint to the 
navigation records. The red rectangle on top of the slider control can be clicked to turn
on and off the replay. When it's on it will be green and the Ship and ROV models will move
along their tracks at a normal (un-accelerated) speed.
</font></td></tr>

<tr><td><%=$font_1%>
<input type="checkbox" name="terrain" value="yes"> Terrain
&nbsp; <%=$font_2%><select name="transp">
<option value="0.0" selected>0%</option>
<option value="0.25">25%</option>
<option value="0.50">50%</option>
<option value="0.75">75%</option>
</select></font> Terrain transparency. (A transparency greater than 0% is good for benthic dives 
where the dive track may go beneath our course - 800m resolution - terrain.
Warning: terrain increases the polygon count, a 3D accelerated graphics card is recommended.)
</font></td></tr>
<tr><td><%=$font_1%>
<input type="checkbox" name="tracks" value="yes" checked> Ship and filtered ROV Tracks
</font></td></tr>
<tr><td><%=$font_1%>
<input type="checkbox" name="orig_rov" value="yes"> Original ROV points
</font></td></tr>
<tr><td><%=$font_1%>
<input type="checkbox" name="rpt" value="yes"> Embed cruise report as text
</font></td></tr>
<tr><td><%=$font_1%>
<input type="checkbox" name="frameGrabs" value="yes"> Top 10 viewed frame grabs
</font></td></tr>
<tr><td><%=$font_1%>
Map variable on ROV track: &nbsp; 
<input type="radio" name="var" value="temp"> Temperature
<input type="radio" name="var" value="sal"> salinity
<input type="radio" name="var" value="oxy"> Oxygen

</font></td></tr>
<tr><td><%=$font_1%>
<input type="checkbox" name="fullscrceen" value="yes"> Full screen (for high-end systems)
</font></td></tr>

<tr><td><%=$font_1%>
<input type="submit" name="genVRML" value="Generate VRML"> (You will need a plugin for 
this to work. Click on the link below if you have not installed CosmoPlayer 2.1.)
</font></tr></td>
</table></form>

<p><%=$font_1%><a href="http://www.karmanaut.com/cosmo/player/"><img src="cplayer21.gif" alt="Download Cosmo Player" border="0" WIDTH="88" HEIGHT="31"></a>
<a href="http://www.karmanaut.com/cosmo/player/">
Download Cosmo Player 2.1.1 for PC/Mac/Irix (Karmanaut mirror site)</a>

<p>Go to the <a href="http://www.mbari.org/~mccann/vrml/ROVvis_proposal/">ROV Data
Visualization page</a></p>

<p><img src="vrmllogo.gif" WIDTH="470" HEIGHT="85"></p>
<hr>
<p>Last updated: <!--webbot bot="Timestamp" S-Type="EDITED" S-Format="%d %b %Y" startspan -->15 Nov 1999<!--webbot bot="Timestamp" endspan i-checksum="15178" -->, Mike McCann</td>
</font></body>
</html>
<%
} # End entry_page()
%>

<%
#--------------------------------------------------------------------
#

=head3 genVRML()

Generate the VRML (.wrl) file based on user's selections.  Uses VRML.pm.

=cut

sub genVRML {

my $dir = $Request->ServerVariables(APPL_PHYSICAL_PATH)->{Item};	# Need to know curr dir
chdir($dir);
my $file = 'terrain.wrl';
$Response->{'ContentType'} = 'x-world/x-vrml';

$vrml = new VRML::VRML2;
$vrml->insert("
EXTERNPROTO MBTerrain [
	field SFFloat terrainTransparency
]
[ \"MBTerrainProto.wrl\", \"gis/MBTerrainProto.wrl\",
  \"http://nauplius.mbari.org/~mccann/vrml/ROVDataVis/gis/MBTerrainProto.wrl\",
  \"http://zoea/ROVDataVis/gis/MBTerrainProto.wrl\"
]
");

$vrml->comment("Autogenerated vrml from vrml.asp");
$vrml->at;				# Begin Encompassing transform
$vrml->navigationinfo(["EXAMINE","FLY"], 1000, 100, 150000)
  ->browser('Cosmo Player 2.0','Netscape')
  ->info("ROV Data Visualization Project. Mike McCann, mccann\@mbari.org.  Copyright (c) 1999 MBARI")
  ->backgroundcolor('0 0 0');
  

if ( GetFormValue('terrain') eq 'yes' ) {
	$vrml->viewpoint("Monterey Bay Entry View", '46851.7 80084.1 87885.9', '-0.986227 0.143928 0.0814914  59.7', 45)
		->insert("MBTerrain {
		terrainTransparency " . GetFormValue('transp') . "
	}");
} # End if( terrain

if ( GetFormValue('tracks') eq 'yes' ) {
	if ( GetFormValue('terrain') ne 'yes' ) {
		$vrml->comment("Begin transform for Entry viewpoint");
		$vrml->insert("       Transform {
       translation 46851.7 80084.1 87885.9
       rotation -0.986227 0.143928 0.0814914  1.047
       children [
		");
		$vrml->viewpoint("Monterey Bay Entry View", '0 0 0', '0 0 1 0', 45)
			->at('t=-3 -1 -8')
			->text("Press PageDown to zoom into dive area.",'yellow','0.5 SANS PLAIN')
			->back;
		$vrml->insert("       ]\n       }");
		$vrml->comment("End transform for Entry viewpoint");
	}
	$vrml->transform_begin('s=1 1 1')
		->Inline(GetFormValue('ShipName').GetFormValue('YYYYDDD')."tracks.wrl")
		->transform_end;
} # End if( tracks


if ( GetFormValue('rpt') eq 'yes' ) {
	$vrml->insert("       Transform {
       translation 46851.7 80084.1 87885.9
       rotation -0.986227 0.143928 0.0814914  1.047
       children [
		");
	$vrml->transform_begin('s=1 1 1')
		->Inline(GetFormValue('ShipName').GetFormValue('YYYYDDD')."rpt.wrl")
		->transform_end;
	$vrml->insert("       ]\n       }");
	$vrml->comment("End transform for Report viewpoint");
} # End if( rpt

if ( GetFormValue('frameGrabs') eq 'yes' ) {
	$vrml->transform_begin('s=1 1 1')
		->Inline(GetFormValue('ShipName').GetFormValue('YYYYDDD')."fg.wrl")
		->transform_end;
} # End if( frameGrabs

if ( GetFormValue('orig_rov') eq 'yes' ) {
	$vrml->transform_begin('s=1 1 1')
		->Inline(GetFormValue('ShipName').GetFormValue('YYYYDDD')."orig_rov.wrl")
		->transform_end;
} # End if( frameGrabs
$vrml->back;			# End Encompassing transform 
$vrml->save("$dir\\$file");

#
# Create the vrml file and write it back to the response object
#
open (FILE, "$dir\\$file");
while (<FILE>) {
	$Response->Write($_);
}
close FILE;
unlink "$dir\\$file";
	

} # End genVRML()
%>