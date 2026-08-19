<%@ LANGUAGE = PerlScript %>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Person Form</title>
</head>
<body bgcolor="#FFFFFF">
<table border="0" cellpadding="5" width="680">
<tr>
<td valign="top" width="20%">
<a href="http://mww.mbari.org">
<img src="/expd/log/images/mbarilogo-120_sh.gif" border="0" width="120" height="70"></a>
</td>
<td valign="top" width="80%">
<!--#include file="expd_functions_telepresence.inc"-->

<% 

# Don't use CGI, it freezes IIS when a new object is created.
use Win32::ASP;
##use Win32::ADO;
use OLE;

$this_script = 'person_telepresence.asp';

init_AppVars();

$dsn = $Application->{'DSN'};

#
# Check that user is registered & kick out if not
#
if ( ! $Session->{'RegisteredUser'} ) { %>
<h2 align="center">Cannot edit person information</h2>
	</td></tr></table><br>
	<font size=5 color=brown><b>You must be a registered user to edit this information. </b></font>
	<br><br><br><br>
	<font><strong>To Add/Edit a person to the database list, please contact the <a href="mailto:mschultz@mbari.org">MBARI Logistics Coordinator</a>.</strong></font>
	<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
	<!--#include file="postcruise_ftr_telepresence.inc"-->
	<% 
	exit;
}

#
# Open DB
#
open_database($dsn);

#
# Read the actual fields in the Person table
#
$sql="SELECT * FROM Person";
$RS = $Conn->Execute($sql);

for ($i = 0; $i < $RS->Fields->Count; $i++) { 
	push @PersonFields, $RS->Fields($i)->Name;
}
$RS->Close;

$Conn->Close;
%>


<%
if ( GetFormValue('edit') ) {
	%>
	<h2 align="center">Edit Person ID <%= GetFormValue('edit')%></h2>
	</td></tr></table>
	<% 
	edit();
}
elsif ( GetFormValue('update') ) {
	%>
	<h2 align="center">Updating Person ID <%= GetFormValue('updateID')%></h2>
	</td></tr></table>
	<% 
	update();
}
elsif ( GetFormValue('add') ) {
	%>
	<h2 align="center">Add Person </h2>
	</td></tr></table>
	<% 
	add();
}
elsif ( GetFormValue('insert') ) {
	%>
	<h2 align="center">Inserting new Person</h2>
	</td></tr></table>
	<% 
	insert();
}
elsif ( GetFormValue('delete') ) {
	%>
	<h2 align="center">Deleting record <%= GetFormValue('deleteID')%></h2>
	</td></tr></table>
	<% 
	delete_rec();
}
else {
	%>
	<h2 align="center">Person List</h2>
	</td></tr></table>
	<% 
	list_table();
}
%>


<!--#include file="precruise_ftr_telepresence.inc"-->

<%

#/*======================================================================
#	edit()
#	
#	Description: 
#	     	Present the info for the person indicated by the ID passed
#		in.
#	     
#	
#	Author: Mike McCann
#	Date Created: 10/28/98
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
$sql="SELECT * FROM Person where PersonID = " . GetFormValue('edit');
Win32::ASP::DebugPrint("\nExecuting SQL: $sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
    $Errors = $Conn->Errors();
    print "Errors:\n";
    foreach $error (keys %$Errors) {
   		print $error->{Description}, "\n";
  	}	
	$RS->Close;
 	die "<BR>Empty Return Set. Does some other app have $dsn open?";
}

%>


<form method="get" action="person_telepresence.asp">
<input type="hidden" name="updateID" value="<%= GetFormValue('edit')%>">
<input type="hidden" name="deleteID" value="<%= GetFormValue('edit')%>">

<table border=0 cellpadding="0" cellspacing="0">
<%
while ( !$RS->EOF ) { 
    foreach $f ( @PersonFields ) { 
	  next if $f eq 'PersonID';
	  next if $f eq 'rowguid';
	  if ( $f =~ /SystemAccount/ ) { %>
		<tr><td valign="top"><b>Email Address:</b></td>
	  <% } else { %>
		<td valign="top"><b><%= $f%>:</b></td>
	  <% } 
	    if ( $f =~ /PickList/ ) {%>
		<td valign="top"><input type="text" size="2" name="<%= $f%>" value="<%= $RS->Fields($f)->value%>">
		 (0: don't show, 1: show)</td>
	    <% }
	    else { %>
		
		<td valign="top"><input type="text" size="50" name="<%= $f%>" value="<%= $RS->Fields($f)->value%>"></td>
		
	    <% } %>
	   </tr>
    <% } %> 
    <% $RS->MoveNext; %>
<% } %>
</table>
<br>
<input type="submit" value="Update Person Table" name="update">
<input type="submit" value="Delete this record" name="delete">
<input type="reset" value="Undo form changes" name="reset">
</form>

<font color=brown><b>Note: Do not Delete anyone who has been a Dive Chief Scientist.  Please set Display fields to 0 instead of deleting. </b></font>
<%

$RS->Close;
$Conn->Close;
}	# End edit()
%>

<%
#/*======================================================================
#	list_table()
#	
#	Description: 
#	     	Present the table of people with an opportunity to edit
#		each record.
#	     
#	
#	Author: Mike McCann
#	Date Created: 10/28/98
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
        die "<BR>Empty Connection object. Does some other app have $dsn open?";
}

#
# Get all the records
#
$sql="SELECT * FROM Person ORDER BY LastName, FirstName";
Win32::ASP::DebugPrint("\nExecuting SQL for $dsn: $sql");
$RS = $Conn->Execute($sql);
if(!$RS) {
	$RS->Close;
 	die "<BR>Empty Return Set from $dsn:<pre>\n$sql</pre>";
}

#
# Print as HTML table - no, not a table, it takes too long
#
%>
<font face="Helvetica,Ariel">
<b>Click on Person's name to edit the record or 
<a href="person_telepresence.asp?add=new">click here to add a person to the list</a></b>
<p>
People that appear in the <i>Principal Investigator</i> pick lists are indicated in
italic, those who appear in the <b>Chief Scientist</b> pick lists are
in bold. The name is bold and italic if the <b><i>person is in both lists</i></b>.</font>
<p>

<% for ($i = 0; $i < $RS->Fields->Count; $i++) {
	push @fields, $RS->Fields($i)->Name;
} 

$str = join(' ', @fields);
Win32::ASP::DebugPrint("\nfields: $str ");
while ( !$RS->EOF ) {
	$fullname = $RS->Fields('FirstName')->Value . "&nbsp;" . $RS->Fields('LastName')->Value;
	$fullname = "<b>$fullname</b>" if $RS->Fields('DisplayChiefSciPickList')->Value;
	$fullname = "<i>$fullname</i>" if $RS->Fields('DisplayPIPickList')->Value;
	$line = "";
	$line .= "<a href=$this_script?edit=" . $RS->Fields('PersonID')->Value . ">";
	$line .= "$fullname</a>$moreinfo";
	$Response->Write("<font face=\"Helvetica,Ariel\" size=-1>$line &#183; </font>");
	$RS->MoveNext;
} 

$RS->Close;
$Conn->Close;
}	# End list_table()
%>

<%


#/*======================================================================
#	update()
#	
#	Description: 
#	     	Update Person record with changes made on the form.
#	     
#	
#	Author: Mike McCann
#	Date Created: 10/28/98
#====================================================================== */
sub update {

#
# Open connection to DataBase
#
$Conn = CreateObject OLE "ADODB.Connection";
$Conn->{'Provider'} = "sqloledb";
$Conn->Open($dsn);
Win32::ASP::AddDeathHook( sub { $Conn->Close } );	# In case asp dies do some cleanup

#
# Open Recrord Set that we'll modify values based on Web form entries and
# then use the Update method on the RS object to update this record in the DB.
#
$RS = CreateObject OLE "ADODB.Recordset";
$sql = "SELECT * FROM Person WHERE PersonID = " . GetFormValue('updateID');
Win32::ASP::DebugPrint("\nExecuting on $dsn, SQL=\n$sql");
$RS = $Conn->Execute($sql);
#$RS->{LockType} = adLockOptimistic;
#$RS->{CursorType} = adOpenStatic;


$sql = "UPDATE Person\nSET ";

foreach $f ( @PersonFields ) {
	
	Win32::ASP::DebugPrint("$f: " . $RS->Fields($f)->Value . "\n");
	next if $f =~ 'PersonID';
	next if $f eq 'rowguid';
	
	if ($f =~ /PickList/) {
		$sql .= "$f=" . FixString(GetFormValue($f), 'num'). ",\n";
	}
	elsif (GetFormValue($f) eq '') {
		$sql .= "$f=NULL,\n";
	}
	else {
		$sql .= "$f=" . FixString(GetFormValue($f), 'text') . ",\n";
	}
	Win32::ASP::DebugPrint("After, $f: " . $RS->Fields($f)->Value . "\n");
}
$sql =~ s/,$//;
$sql .= "WHERE (PersonID=" . GetFormValue('updateID') . ")";

Win32::ASP::DebugPrint("Executing \n$sql\n");
$Conn->Execute($sql);


$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("Database Load Error(s):<pre>$sql</pre>");
	foreach $error (keys %$Errors)
        {
      		$Response->Write($error->{Description});
   	}
}
else {
	%>
	<p><%= GetFormValue('FirstName')%> <%= GetFormValue('LastName')%>'s record has been updated.</p>
<%
}
%>


<p><a href="person_telepresence.asp">Return to Person List</a></p>

<%
$RS->Close;
$Conn->Close;
}	# End update()
%>

<%

#/*======================================================================
#	add()
#	
#	Description: 
#	     	Insert a Person record.
#	     
#	
#	Author: Mike McCann
#	Date Created: 10/28/98
#====================================================================== */
sub add {

%>

<form method="get" action="person_telepresence.asp">
<input type="hidden" name="insert" value="<%= GetFormValue('add')%>">
<table border=0 cellpadding="0" cellspacing="0">
<%
foreach $f ( @PersonFields ) {
	  next if $f eq 'PersonID';
	  next if $f eq 'rowguid';
	%><tr><td valign="top"><b><%= $f%>:</b></td><%
	    if ( $f =~ /PickList/ ) {%>
		<td valign="top"><input type="text" size="2" name="<%= $f%>" value="">
		 (0: don't show, 1: show)</td>
	    <% }
	    else { %>
		<td valign="top"><input type="text" size="50" name="<%= $f%>" value=""></td>
	    <% } %>
	   </tr>
<% } %> 
</table>
<br>
<input type="submit" value="Add Person" name="insert">
<input type="reset" value="Reset form" name="reset">
</form>
	
<%
} 	# End add()
%>

<%

#/*======================================================================
#	insert()
#	
#	Description: 
#	     	Add person to the table.
#	     
#	
#	Author: Mike McCann
#	Date Created: 10/28/98
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
        die "<BR>Empty Connection object.";
}

#
# Update the records
#
$sql = "INSERT INTO Person\n (";

foreach $f ( @PersonFields ) {
	next if $f eq 'PersonID';
	next if $f eq 'rowguid';
	$sql .= "$f,";
}
$sql =~ s/,$//;
$sql .= ")\nVALUES(";
foreach $f ( @PersonFields ) {
	next if $f eq 'PersonID';
	next if $f eq 'rowguid';
	$val = (GetFormValue($f) eq '') ? 'NULL' : GetFormValue($f);
	
	if ($f =~ /PickList/ && $val ne 'NULL') {
		$sql .= FixString($val, 'num') . ",";
	}
	else {
		$sql .= FixString($val, 'text') . ",";
	}
}
$sql =~ s/,$//;
$sql .= ")";

Win32::ASP::DebugPrint("\nExecuting SQL:\n$sql");
$RS = $Conn->Execute($sql);

$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("Database Load Error(s):<pre>$sql</pre>");
	foreach $error (keys %$Errors) {
      		$Response->Write($error->{Description});
   	}
}
else {
	%>
	<p><%= GetFormValue('FirstName')%>&nbsp;<%= GetFormValue('LastName')%> has been added.</p>
<%
}
%>

<p><a href="precruise_telepresence.asp?step=1">Return to Pre-Cruise entry form step 1</a></p>
<p><a href="person_telepresence.asp">Return to Person List</a></p>

<%
	$RS->Close;
	$Conn->Close;
}	# End insert()
%>


<%
#/*======================================================================
#	delete_rec()
#	
#	Description: 
#	     	Delete a record.
#	     
#	
#	Author: Mike McCann
#	Date Created: 11/4/98
#====================================================================== */
sub delete_rec {

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
        die "<BR>Empty Connection object.";
}

#
# Update the records
#
$sql = "DELETE Person \nWHERE (PersonID = " . GetFormValue('deleteID') . ") \n";
Win32::ASP::DebugPrint("\nExecuting SQL:\n$sql");

$Conn->Execute($sql);

$Errors = $Conn->Errors();
if ( keys %$Errors ) { 
	$Response->Write("Database Load Error(s):<pre>$sql</pre>");
	foreach $error (keys %$Errors) {
      		$Response->Write($error->{Description});
   	}
}
else {
	$Response->Write("Record " . GetFormValue('deleteID') . " successfully deleted.<p>");
}

$Conn->Close;
%>
<a href="person_telepresence.asp">Return to person list</a>
<%

}	# End delete_rec()
%>