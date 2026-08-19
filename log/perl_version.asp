<%@ language="PerlScript" %>
   <% $Response->Write("Perl: $^V (\$] = $]) on $^O. ActiveState build: " . (eval { Win32::BuildNumber() } || 'n/a')); %>