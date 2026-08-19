<%@ LANGUAGE = PerlScript %>

<%

use ASP;
use OLE;
use Win32::OLE qw( in );


strWrld="Hello World";

$Session->{'DBAdministrator'} = 0;

%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html> 
	<head>
		<title>ASP Test Page</title>
	</head>
	<body>

		If you see "Hello World" below, then ASP pages in this Web application are serving properly.
		<br>
	     	<STRONG><%=$strWrld%></STRONG>
	     	<br>
	     	USER_AGENT = <%=$Request->ServerVariables(HTTP_USER_AGENT)->{Item}%>
	     	<br>
	     	LOGON_USER = <%=$Request->ServerVariables('LOGON_USER')->{Item}%>
	     	<br>
	     	DBAdministrator = <%=$Session->{'DBAdministrator'}%>
	     	<%
	     	if ( 	$Request->ServerVariables('LOGON_USER')->{Item} =~ /mschultz/ ){
	     		Session->{'DBAdministrator'} = 1;
	     	}
	     	%>
	     	<br>
	     	DBAdministrator = <%=Session->{'DBAdministrator'}%>
	</body>
</html>