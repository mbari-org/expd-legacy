<% @ LANGUAGE = PerlScript %>

<%
#/*======================================================================
#	region.asp
#	
#	Description: 
#	     	Present the table of regions with an opportunity to edit,
#		insert, and delete records.
#	     
#	
#	Author: Dan Wilkin
#	Date Created: 11/2/98
#====================================================================== */
%>

<!-- put the mbari logo at the top of the page -->
<!--#include file="precruise_hdr.inc"-->

<% 
# Don't use CGI, it freezes IIS when a new object is created.
use Win32::ASP;		# ActivePerl module
use OLE;		# database connectivity

$this_script = 'region.asp';
#
# The database is opened and closed in each step.
# $dsn (data source name) object is an Application variable.
#
$dsn=$Application->{'DSN'};				# data source name

%>

<%
# check the hidden cgi variable to see which step we are on

# if edit was given, go to it

if ( GetFormValue('edit') ) {
	%>
	<h2 align="center">Edit Region with RegionID <%= GetFormValue('edit')%></h2>
	</td></tr></table>
	<% 
	edit();
}
elsif ( GetFormValue('update') ) {
	%>
	<h2 align="center">Updating Region</h2>
	</td></tr></table>
	<% 
	update();
}
elsif ( GetFormValue('add') ) {
	%>
	<h2 align="center">Add new Region</h2>
	</td></tr></table>
	<% 
	add();
}
elsif ( GetFormValue('insert') ) {
	%>
	<h2 align="center">Inserting new Region</h2>
	</td></tr></table>
	<% 
	insert();
}
else {
	%>
	<h2 align="center">MBARI Region List</h2>
	</td></tr></table>
	<% 
	list_table();
}
%>

<!--#include file="precruise_ftr.inc"-->

<%
#/*======================================================================
#	list_table()
#	
#	Description: 
#	     	Present the table of regions with an opportunity to edit,
#		insert, and delete records.
#	     
#	
#	Author: Dan Wilkin
#	Date Created: 10/30/98
#====================================================================== */

sub list_table {

#
# Open DB
#
$Conn = CreateObject OLE "ADODB.Connection";
$Conn->{'Provider'} = "sqloledb";
$Conn->Open($dsn);
Win32::ASP::AddDeathHook( sub { $Conn->Close } );	# In case asp dies do some cleanup

if(!Conn) {
        $Errors = $Conn->Errors();
        print "<B>Errors:</B>";
        foreach $error (keys %$Errors) {
                print $error->{Description}, "\n";
        }
        die "<BR>Empty Connection object for $dsn";
}

#
# Get all the records
#
$sql="SELECT * FROM Region ORDER BY RegionName";

Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
        $Errors = $Conn->Errors();
        print "Errors:\n";
       	foreach $error (keys %$Errors) {
   		print $error->{Description}, "\n";
       	}	
 	die "<BR>Empty Return Set \n$sql\n from \n$dsn ";
}

#
# Print as HTML table
#

$RS->MoveFirst;

%>
<b>Click on a RegionID hyperlink below to edit a record or 
<a href="region.asp?add=new">click here to add a region to the list</a>.</b>
<p>
<table border="1" style="font-family: sans-serif; font-size: smaller">

<tr>
<% for ($i = 0; $i < $RS->Fields->Count; $i++) { %>
	<th bgcolor="AAFFAA"><%= $RS->Fields($i)->Name%></th>
	<% push @fields, $RS->Fields($i)->Name; %>
<% } %>
</tr>
<%
$str = join(' ', @fields);
Win32::ASP::DebugPrint("\nfields: $str ");
%>
<tr>
<% while ( !$RS->EOF ) { %>
	<tr>
	<% foreach $f (@fields) { 
		$val = ($RS->Fields($f)->value) ? $RS->Fields($f)->value : '&nbsp;';
		$cell = '<td bgcolor="DDDDDD">';
		if ($f eq 'RegionID') {
			$val = "<a href=\"region.asp?edit=$val\">$val</a>";
		} %><%= $cell%><%= $val%></td>
	<% } %>
	<% $RS->MoveNext; %>
<% } %>
</tr>

</table>

<%
$RS->Close;
$Conn->Close;
}	# End list_table()
%>






<%

#/*======================================================================
#	edit()
#	
#	Description: 
#	     	Present the info for the region indicated by the ID passed
#		in.
#	     
#	
#	Author: Dan Wilkin
#	Date Created: 11/2/98
#====================================================================== */
sub edit {
	
#
# Open DB
#
$Conn = CreateObject OLE "ADODB.Connection";
$Conn->{'Provider'} = "sqloledb";
$Conn->Open($dsn);
Win32::ASP::AddDeathHook( sub { $Conn->Close } );	# In case asp dies do some cleanup

if(!Conn) {
        $Errors = $Conn->Errors();
        print "<B>Errors:</B>";
        foreach $error (keys %$Errors) {
                print $error->{Description}, "\n";
        }
        die "<BR>Empty Connection object. Does some other app have $dsn open?";
}

#
# Get the record to edit
#
$sql="SELECT RegionName, MaxLatitude, MinLatitude, MaxLongitude, MinLongitude, Comment " . 
     "FROM Region where RegionID = " . GetFormValue('edit');

Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
        $Errors = $Conn->Errors();
        print "Errors:\n";
       	foreach $error (keys %$Errors) {
   		print $error->{Description}, "\n";
       	}	
 	die "<BR>Empty Return Set. Does some other app have $dsn open?";
}

for ($i = 0; $i < $RS->Fields->Count; $i++) { 
	push @fields, $RS->Fields($i)->Name;
}
$Session->{'\@fields'} = @fields;
%>


<form method="get" action="region.asp">
<input type="hidden" name="update" value="<%= GetFormValue('edit')%>">

<table border=0 cellpadding="0" cellspacing="0">
<%
while ( !$RS->EOF ) { 
    foreach $f (@fields) { 
	%><tr><td valign="top"><b><%= $f%>:</b></td>
		<td valign="top"><input type="text" size="50" name="<%= $f%>" value="<%= $RS->Fields($f)->value%>"></td>
	  </tr>
    <% } %> 
    <% $RS->MoveNext; %>
<% } %>
</table>
<br>
<input type="submit" value="Update Region Table" name="update">
<input type="reset" value="Undo form changes" name="reset">
</form>
	
<%

$RS->Close;
$Conn->Close;
}	# End edit()
%>









<%

#/*======================================================================
#	add()
#	
#	Description: 
#	     	Insert a Region record.
#	     
#	
#	Author: Dan Wilkin
#	Date Created: 11/2/98
#====================================================================== */
sub add {

@fields = ('RegionName', 'MaxLatitude', 'MinLatitude',
	   'MaxLongitude', 'MinLongitude', 'Comment');

%>

<form method="get" action="region.asp">
<input type="hidden" name="insert" value="<%= GetFormValue('add')%>">
<table border=0 cellpadding="0" cellspacing="0">
<%

foreach $f (@fields) { 
	%><tr><td valign="top"><b><%= $f%>:</b></td>
		<td valign="top"><input type="text" size="50" name="<%= $f%>" value=""></td>
	  </tr>
<% } %> 
</table>
<br>
<input type="submit" value="Add Region" name="insert">
<input type="reset" value="Reset form" name="reset">
</form>
	
<%
} 	# End add()
%>





<%

#/*======================================================================
#	update()
#	
#	Description: 
#	     	Update Region record with changes made on the form.
#	     
#	
#	Author: Dan Wilkin
#	Date Created: 11/2/98
#====================================================================== */
sub update {
	
#
# Open DB
#
$Conn = CreateObject OLE "ADODB.Connection";
$Conn->{'Provider'} = "sqloledb";
$Conn->Open($dsn);
Win32::ASP::AddDeathHook( sub { $Conn->Close } );	# In case asp dies do some cleanup

if(!Conn) {
        $Errors = $Conn->Errors();
        print "<B>Errors:</B>";
        foreach $error (keys %$Errors) {
                print $error->{Description}, "\n";
        }
        die "<BR>Empty Connection object. Does some other app have $dsn open?";
}

#
# Update the records
#

$sql = "UPDATE Region\nSET ";

@fields = ('RegionName', 'MaxLatitude', 'MinLatitude',
	   'MaxLongitude', 'MinLongitude', 'Comment');

foreach $f (@fields) {
	
	if (GetFormValue($f) eq '') {
		$sql .= "$f=NULL,\n";
	}
	else {
		$sql .= "$f=" . FixString(GetFormValue($f), 'text') . ",\n";
	}
}
$sql =~ s/,$//;
$sql .= "WHERE (RegionID=" . GetFormValue('update') . ")";

Win32::ASP::DebugPrint("\nExecuting SQL:\n$sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
        $Errors = $Conn->Errors();
        print "Errors:\n";
       	foreach $error (keys %$Errors) {
   		print $error->{Description}, "\n";
       	}	
 	die "<BR>Empty Return Set.";
}

%>
<p>The record for <%= GetFormValue('RegionName')%> has been updated.</p>
#<p><a href="mike_precruise.asp?step=1">Return to Pre-Cruise entry form step 1</a></p>
<p><a href="precruise.asp?step=1">Return to Pre-Cruise entry form step 1</a></p>
<p><a href="region.asp">Return to Region List</a></p>

<%
}	# End update()
%>






<%

#/*======================================================================
#	insert()
#	
#	Description: 
#	     	Add region to the table.
#	     
#	
#	Author: Dan Wilkin
#	Date Created: 11/2/98
#====================================================================== */
sub insert {
	
#
# Open DB
#
$Conn = CreateObject OLE "ADODB.Connection";
$Conn->{'Provider'} = "sqloledb";
$Conn->Open($dsn);
Win32::ASP::AddDeathHook( sub { $Conn->Close } );	# In case asp dies do some cleanup

if(!Conn) {
        $Errors = $Conn->Errors();
        print "<B>Errors:</B>";
        foreach $error (keys %$Errors) {
                print $error->{Description}, "\n";
        }
        die "<BR>Empty Connection object. Does some other app have $dsn open?";
}

#
# Update the records
#
$sql = "INSERT INTO Region\n (";

@fields = ('RegionName', 'MaxLatitude', 'MinLatitude',
	   'MaxLongitude', 'MinLongitude', 'Comment');

foreach $f (@fields) {
	$sql .= "$f,";
}
$sql =~ s/,$//;
$sql .= ")\nVALUES(";
foreach $f (@fields) {

	$val = (GetFormValue($f) eq '') ? 'NULL' : GetFormValue($f);
	
	$sql .= FixString($val, 'text') . ",";
	
}

$sql =~ s/,$//;
$sql .= ")";

Win32::ASP::DebugPrint("\nExecuting SQL:\n$sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
        $Errors = $Conn->Errors();
        print "Errors:\n";
       	foreach $error (keys %$Errors) {
   		print $error->{Description}, "\n";
       	}	
 	die "<BR>Empty Return Set.";
}

%>
<p>Region <%= GetFormValue('RegionName')%> has been added.</p>
#<p><a href="mike_precruise.asp?step=1">Return to Pre-Cruise entry form step 1</a></p>
<p><a href="precruise.asp?step=1">Return to Pre-Cruise entry form step 1</a></p>
<p><a href="region.asp">Return to Region List</a></p>

<%
}	# End insert()
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
	$str =~ s/'/\'/g;
	$fixed_string = "'$str'";
}

if ($type =~ /num/i || $str eq 'NULL' ) {
	$fixed_string = "$str";
}

return $fixed_string;

}
%>