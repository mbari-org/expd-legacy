package SSDS;

use strict;
use Carp;


#
# Private variables
#
my $_ssdsServer = 'http://ssds-test.shore.mbari.org:8080/';
my $_baseServlet = 'access/MetadataAccessServlet?replyInHTML=false&';
my $_delim = '|';
my $_baseUrl = $_ssdsServer . $_baseServlet . "&delimiter=" . $_delim . "&";
my $_debug = 0;

my %_debugLevel = ( ERROR => 1, WARN => 2,  INFO => 3 );

=head1 SSDS

SSDS.pm - Access SSDS Metadata

=head1 SYNOPSIS

	use SSDS;
	
	my $ssds = new SSDS();
	my $ssds->ssdsServer("http://ssds.shore.mbari.org:8080/");
	
	my $deplAccess = new SSDS::DeploymentAccess();
	my @deplAttr = (${$depls}[0]->get_attribute_names());
	my $depls = $deplAccess->findByName("CIMT Mooring Deployment");
	foreach my $d (@{$depls}) {
		foreach my $a (@deplAttr) {
			next unless $depl->$a;
			print "  $a = " . $depl->$a . "\n";
		}	
	}
        	
	Will print out:
	
	  id = 2356
	  name = CIMT Mooring Deployment
	  startDate = Tue May 18 15:30:00 PDT 2004
	  role = platform


=head1 DESCRIPTION

The SSDS class provides a Perl interface to the SSDS metadata services
provided by the SSDS Web Services.  Much of the Java interface to 
the SSDS model and service objects is implemented here.  This module 
was written to make it easy to write external data processing scripts 
for SSDS data.

The syntax used to derefernce the return values from the Access methods
looks a little hairy, but it's really not that bad.  If the method is designed
to return a single object (such as C<findByPK()>) then the return value will 
be an object which you may immediatley use to call its attributes.  
If the method returns multiple objects (such as
C<listOutputs()>) then the return value will be a reference to a list of
objects, which you must dereference with a '$' or '@' before using it.  
Plenty of examples are provided in the method descriptions of this document.

C<$Id: SSDS.pm,v 1.26 2005/08/29 19:43:58 mccann Exp $>

The SSDS package contains these methods which are inherited by all the SSDS objects declared in the other packages.

=over 4

=item *

ssdsServer - set to base SSDS server, e.g. "http://ssds.shore.mbari.org:8080/"

=item *

baseServlet - set to the MetadataAccessServlet, e.g. 'access/MetadataAccessServlet?replyInHTML=false&';

=item *

baseUrl - the concatenation of ssdsServer and baseServlet

=item *

delim - the Delimiter to use for field separation of HTTP responses from SSDS, e.g. '|';

=item *

debug - set it to 1 for ERROR, 2 for WARN, 3 for INFO, 0 for no output (the default).
It's a good idea to run with debug set to 2, then set it to 0 when put into production.


=head2 Base class with private methods

=cut


sub new {
	my ($class) = @_;
	my $self = {};
	bless $self, $class;
}

sub ssdsServer { shift; @_ ? $_ssdsServer = shift : $_ssdsServer;}

sub baseServlet { shift; @_ ? $_baseServlet = shift : $_baseServlet;}

sub baseUrl { 
	shift; 
	@_ ? $_baseServlet = shift : $_ssdsServer . $_baseServlet . "delimiter=" . $_delim . "&";
}

sub delim { shift; @_ ? $_delim = shift : $_delim;}

sub debug { shift; @_ ? $_debug = shift : $_debug;}



#====================================================================
#                             ObjectCreator
#====================================================================
package SSDS::ObjectCreator;
use LWP::Simple;
use strict;
use Carp;
our @ISA = qw(SSDS);

=back

=head3 ObjectCreator

ObjectCreator - Contains private support methods to create SSDS objects.  The methods below are inherited by all the SSDS objects declared in the child packages.


=over 4

=item B<new()>

=cut

sub new {
	my $pkg = shift;
	bless {}, $pkg;
}

#--------------------------------------------------------------------
#

=item B<_buildHash()>
	 
Function to parse return of MetadataAccessServlet into a single (key, value) hash list.

Example:

	my ($returnedObject, $oHash) = $obj->_buildHash($obj->delim, $data);

=cut

sub _buildHash {
	
	my $obj = shift;
	my $delim = shift;
	my $servletResponse = shift;
	
	my %hash = ();
	
	my ($k, $v, $ssdsObjectName);
	
	$delim = '\|' if $delim eq '|';
	print "_buildHash(): delimiter set to $delim" if $obj->debug == $_debugLevel{'INFO'};
	
	foreach my $line (split('\n', $servletResponse)) {
		print "_buildHash(): line = $line\n" if $obj->debug == $_debugLevel{'INFO'};
		if ( $line =~ /Fault/ ) {
			if ( $obj->debug == $_debugLevel{'INFO'} || $obj->debug == $_debugLevel{'ERROR'} || $obj->debug == $_debugLevel{'WARN'} ) {
				print "$servletResponse\n";
			}
			last;
		}
		foreach my $prop ( split($delim, $line) ) {
			if ( $prop =~ /(.+?)=(.+)/ ) {
				$k = $1;
				$v = $2;
				print "_buildHash(): k = $k, v = $v\n" if $obj->debug == $_debugLevel{'INFO'};
				if ($v && $v !~ /null/) {
					$hash{${k}} = $v;
				}	
			}
			elsif ( $prop =~ /(^[A-Z].+[^=])/ ) {
				$ssdsObjectName = $prop;
				print "_buildHash(): ssdsObjectName = $ssdsObjectName\n" if $obj->debug == $_debugLevel{'INFO'};
			}
		}	
	}

	return ($ssdsObjectName, \%hash);
	
} # End _buildHash()


#--------------------------------------------------------------------
#

=item B<_buildHashes()>

Function to parse return of MetadataAccessServlet into a list of (key, value) hashes list.

Example:

	my ($returnedObject, $oHashes) = $obj->_buildHashes($obj->delim, $data);

=cut

sub _buildHashes {
	
	my $obj = shift;
	my $delim = shift;
	my $servletResponse = shift;
	
	my $rhash = {};		# Need to use a reference to the hash in order to push onto @hashes list
	my @hashes = ();
	my @foo = ();
	
	my ($k, $v, $ssdsObjectName);
	
	$delim = '\|' if $delim eq '|';
	
	foreach my $line (split('\n', $servletResponse)) {
		print "_buildHashes(): line = $line\n" if $obj->debug == $_debugLevel{'INFO'};
		if ( $line =~ /Fault/ ) {
			if ( $obj->debug == $_debugLevel{'INFO'} || $obj->debug == $_debugLevel{'ERROR'} || $obj->debug == $_debugLevel{'WARN'} ) {
				print "$servletResponse\n";
			}
			last;
		}
		$rhash = {};
		foreach my $prop ( split($delim, $line) ) {
			if ( $prop =~ /(.+?)=(.*)/ ) {
				$k = $1;
				$v = $2;
				print "_buildHashes(): k = $k, v = $v\n" if $obj->debug == $_debugLevel{'INFO'};
				if ( $v && $v !~ /null/ ) {
					##print "_buildHashes(): adding to hash\n" if $obj->debug == $_debugLevel{'INFO'};
					${$rhash}{$k} = $v;
				}
			}
			elsif ( $prop =~ /(^[A-Z].+[^=])/ ) {
				$ssdsObjectName = $prop;
				print "_buildHashes(): ssdsObjectName = $ssdsObjectName\n" if $obj->debug == $_debugLevel{'INFO'};
			}	
		}
		push @hashes, $rhash;	

	}

	print "\n\nChecking final hash list before returning it...\n" if $obj->debug == $_debugLevel{'INFO'};
	foreach my $rh (@hashes) {
		print "_buildHashes(): id = ${$rh}{'id'}\n" if $obj->debug == $_debugLevel{'INFO'};
		foreach my $key (keys %{$rh}) {
			print "_buildHashes(): key = $key, value = ${$rh}{$key}\n" if $obj->debug == $_debugLevel{'INFO'};
		}
		print "\n" if $obj->debug == $_debugLevel{'INFO'};
	}
	return ($ssdsObjectName, \@hashes);
	
} # End _buildHashes()


#--------------------------------------------------------------------
# Function that returns a reference to a hash 

=item B<_createSSDSobject()>

Build and return an SSDS object given a URL to a MetadataAccessServlet call.
If a second argument is given then that SSDS object will be instantiated rather than the object returned from
the MetadataAccessServlet.  If the types do not match a warning message will be printed.

Examples:

	# Request and return a Device object
	my $url = $obj->baseUrl . "className=DeviceTypeAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
	
	# Request a DataStream and return a DataFile object
	my $url = $obj->baseUrl . "className=DataStreamAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url, 'DataFile');

=cut

sub _createSSDSobject {
	
	my $obj = shift;
	my $url = shift;
	my $ssdsObject = shift;
	
	print "Getting metadata from url = $url\n" if $obj->debug == $_debugLevel{'INFO'};
	my $data = get($url);
	if  ( !$data ) {
		if ( $obj->debug == $_debugLevel{'INFO'} || $obj->debug == $_debugLevel{'ERROR'} || $obj->debug == $_debugLevel{'WARN'} ) {
			print "\nERROR _createSSDSobject(): No data returned from url:\n$url\n\n";
		}
		return;
	}
	
	#
	# Pull off attributes and return an SSDS object
	# Return a $returnedObject unless $ssdsObject specified
	#
	my ($returnedObject, $oHash) = $obj->_buildHash($obj->delim, $data);
	return unless $returnedObject;
	if ( $ssdsObject && $ssdsObject ne $returnedObject ) {
		if ( $obj->debug == $_debugLevel{'INFO'} || $obj->debug == $_debugLevel{'WARN'} ) {
			print "WARNING _createSSDSobject(): returned object is $returnedObject, requested to cast to $ssdsObject\n";
		}
	}
	$ssdsObject = ($ssdsObject) ? $ssdsObject : $returnedObject;
	
	my $o = "SSDS::$ssdsObject"->new();
	print "_createSSDSobject(): $ssdsObject object:\n" if $obj->debug == $_debugLevel{'INFO'};
	print "_createSSDSobject(): --------------\n" if $obj->debug == $_debugLevel{'INFO'};
	foreach my $k (keys %{$oHash}) {
		print "_createSSDSobject(): Assigning values to object attributes: $k = " . ${$oHash}{$k} . "\n" if $obj->debug == $_debugLevel{'INFO'};
		$o->$k(${$oHash}{$k});
	}
	
	return $o;
	
} # end _createSSDSobject


#--------------------------------------------------------------------
# Function that returns a reference to a list of hash references

=item B<_createSSDSobjects()>

Build and return a list of SSDS objects given a URL to a MetadataAccessServlet call.
If a second argument is given then that SSDS object will be instantiated rather than the object returned from
the MetadataAccessServlet.  If the types do not match a warning message will be printed.

Examples:

	# Request and return a Deployment 
	my $url = $obj->baseUrl . "className=Deployment&methodName=listChildDeployments&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobjects($url);
	
	# Request a DataStream and return a DataFile
	my $url = $obj->baseUrl . "className=Deployment&methodName=listOutputs&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobjects($url, 'DataFile');

=cut

sub _createSSDSobjects {
	
	my $obj = shift;
	my $url = shift;
	my $ssdsObject = shift;
	
	my @os = ();	# List of SSDS objects to be returned
	
	print "Getting metadata from url = $url\n" if $obj->debug == $_debugLevel{'INFO'};
	my $data = get($url);
	if  ( !$data ) {
		if ( $obj->debug == $_debugLevel{'INFO'} || $obj->debug == $_debugLevel{'ERROR'} || $obj->debug == $_debugLevel{'WARN'} ) {
			print "\nERROR _createSSDSobject(): No data returned from url:\n$url\n\n";
		}
		return;
	}
	
	#
	# Pull off attributes and return a list of SSDS objects
	# Return $returnedObjects unless $ssdsObject specified
	#
	my ($returnedObject, $oHashes) = $obj->_buildHashes($obj->delim, $data);
	return unless $returnedObject;
	if ( $ssdsObject && $ssdsObject ne $returnedObject ) {
		if ( $obj->debug == $_debugLevel{'INFO'} || $obj->debug == $_debugLevel{'WARN'} ) {
			print "WARNING _createSSDSobject(): returned object is $returnedObject, requested to cast to $ssdsObject\n";
		}
	}
	$ssdsObject = ($ssdsObject) ? $ssdsObject : $returnedObject;
	
	
	print "_createSSDSobjects(): $ssdsObject objects:\n" if $obj->debug == $_debugLevel{'INFO'};
	print "_createSSDSobjects(): -------------------\n" if $obj->debug == $_debugLevel{'INFO'};
	foreach my $oh (@{$oHashes}) {
		my $o = "SSDS::$ssdsObject"->new();
		foreach my $k (keys %{$oh}) {
			print "_createSSDSobjects():   $k = " . ${$oh}{$k} . "\n" if $obj->debug == $_debugLevel{'INFO'};
			$o->$k(${$oh}{$k});
		}
		push @os, $o;
		
	}
	return \@os;
	
} # end _createSSDSobjects



#--------------------------------------------------------------------
# Function that returns a reference to a hash 

=item B<_execSSDSmethod()>

Execute an SSDS access method that does not return an object.  Typically this is 
used for calling set(s), insert()s, update()s, and delete()s.  Returns 1 if successful, 0
if a Fault is detected or any data is returned from the call.

Example:

	my $url = $obj->baseUrl . "className=DeviceTypeAccess&methodName=insert&param1Type=DeviceType&param1Value=DeviceType";
	$url .= $obj->delim() . "name=" . $dt->name() if $dt->name();
	$url .= $obj->delim() . "description=" . $dt->description() if $dt->description();
	$url .= $obj->delim() . "defaultDeploymentRole=" . $dt->defaultDeploymentRole() if $dt->defaultDeploymentRole(); 
	$url .= $obj->delim() . "displayInDeviceTypesPicklist=" . $dt->displayInDeviceTypesPicklist() if $dt->displayInDeviceTypesPicklist();
	return $obj->_execSSDSmethod($url);

=cut

sub _execSSDSmethod {
	
	my $obj = shift;
	my $url = shift;
	my $ssdsObject = shift;
	
	print "Executing method in url = $url\n" if $obj->debug == $_debugLevel{'INFO'};
	my $data = get($url);
	
	if  ( !$data ) {
		return 1;
	}
	else {
		if ( $data =~ /Fault/ ) {
			if ( $obj->debug == $_debugLevel{'INFO'} || $obj->debug == $_debugLevel{'ERROR'} || $obj->debug == $_debugLevel{'WARN'} ) {
				print "$data\n";
			}
		}
		return 0;
	}
	
} # end execSSDSmethod



#====================================================================
#                           DataFile
#====================================================================
package SSDS::DataFile;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);

attributes qw(id name description startDate endDate original url contentLength mimeType webAccessible dodsAccessible fileName); 

=back 

=head2 Value and DataAccessObject classes

=head3 DataFile

DataFile - An SSDS DataFile object


=over 4

=item B<new()>

Instantiate an SSDS::DataFile object

The following methods may be used to get or set any of the attributes of a 
Deployment object (each attribute is both a setter and getter):

    id name description startDate endDate original url contentLength mimeType
    webAccessible dodsAccessible fileName

Note: the get_attribute_names() method will return this list of attributes

Example:

	my $dcAccess = new SSDS::DataFileAccess();
	my $dc = $dcAccess->findByPK(5777);

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}


#--------------------------------------------------------------------
# 

=item B<getRecordDescription()>

Return the RecordDescription object for this DataFile

Example:

	my $rd = $dc->getRecordDescription();
	foreach my $a ($rd->get_attribute_names()) {
		next unless $rd->$a;
		print "  $a = " . $rd->$a . "\n";
	}

=cut

sub getRecordDescription {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=DataFile&methodName=getRecordDescription&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobject($url);
}


#====================================================================
#                           DataFileAccess
#====================================================================
package SSDS::DataFileAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 DataFileAccess

DataFileAccess - Contains methods to query for and get DataFile objects


=over 4

=item B<new()>

Instantiate an SSDS::DataFileAccess object

Example:

	my $dcAccess = new SSDS::DataFileAccess();
	my $dc = $dcAccess->findByPK(5777);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}


#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the DataFile with PK = pk

Example:

	my $dcAccess = new SSDS::DataFileAccess();
	my $dc = $dcAccess->findByPK(5777);

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	
	my $url = $obj->baseUrl . "className=DataFileAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
}


#--------------------------------------------------------------------
# 

=item B<findByUrl(string)>

Return a list of DataFiles where the url is string.

Example:

	my $dfAccess = new SSDS::DataFileAccess();
	my $df = $dfAccess->findByUrl("http://ssds-test.shore.mbari.org/ssds/rawpackets/1313_0_1_1327");
	print "isWebAccessible = " . ${$df}[0]->webAccessible() . "\n";

=cut

sub findByUrl {
	my $obj = shift;
	my $string = shift;
	
	my $url = $obj->baseUrl . "className=DataFileAccess&methodName=findByUrl&param1Type=String&param1Value=" . $string;	
	return $obj->_createSSDSobjects($url);
}


#--------------------------------------------------------------------
# 

=item B<insert($df)>

Insert into the database the DataFile $df

Example:


=cut

sub insert {
	my $obj = shift;
	my $df = shift;
	my $url = $obj->baseUrl . "className=DataFileAccess&methodName=insert&param1Type=DeviceType&param1Value=DataFile";
	foreach my $a ($df->get_attribute_names()) {
		$url .= $obj->delim() . $a . "=" . $df->$a if $df->$a;
	}
	return $obj->_execSSDSmethod($url);
}

#--------------------------------------------------------------------
# 

=item B<delete($dt)>

Delete DataFile $dt from the database

Example:

	my $dfAccess = new SSDS::DataFileAccess();
	my $rdfs = $dfAccess->findByUrl($dfURL);;
	$dfAccess->delete(${$rdfs}[0]);

=cut

sub delete {
	my $obj = shift;
	my $dt = shift;
	my $url = $obj->baseUrl . "className=DataFileAccess&methodName=delete&param1Type=DataFile&param1Value=DataFile";
	$url .= $obj->delim() . "id=" . $dt->id() if $dt->id();
	return $obj->_execSSDSmethod($url);
}


#====================================================================
#                           DataStream
#====================================================================
package SSDS::DataStream;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);

attributes qw(id name description startDate endDate original url contentLength mimeType webAccessible); 

=back 

=head3 DataStream

DataStream - An SSDS DataStream object


=over 4

=item B<new()>

Instantiate an SSDS::DataStream object

The following methods may be used to get or set any of the attributes of a 
Deployment object (each attribute is both a setter and getter):

    id name description startDate endDate original url contentLength mimeType
    webAccessible

Note: the get_attribute_names() method will return this list of attributes

Example:

	my $dcAccess = new SSDS::DataStreamAccess();
	my $dc = $dcAccess->findByPK(5777);

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}


#--------------------------------------------------------------------
# 

=item B<getRecordDescription()>

Return the RecordDescription object for this DataStream

Example:

	my $rd = $dc->getRecordDescription();
	foreach my $a ($rd->get_attribute_names()) {
		next unless $rd->$a;
		print "  $a = " . $rd->$a . "\n";
	}

=cut

sub getRecordDescription {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=DataStream&methodName=getRecordDescription&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobject($url);
}


#====================================================================
#                           DataStreamAccess
#====================================================================
package SSDS::DataStreamAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 DataStreamAccess

DataStreamAccess - Contains methods to query for and get DataStream objects


=over 4

=item B<new()>

Instantiate an SSDS::DataStreamAccess object

Example:

	my $dcAccess = new SSDS::DataStreamAccess();
	my $dc = $dcAccess->findByPK(5777);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}

#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the DataStream with PK = pk

Example:

	my $dcAccess = new SSDS::DataStreamAccess();
	my $dc = $dcAccess->findByPK(5777);

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	
	my $url = $obj->baseUrl . "className=DataStreamAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
}


#--------------------------------------------------------------------
# 

=item B<findByUrl(string)>

return a list of DataStreams where the url is string.

Example:

	my $dfAccess = new SSDS::DataStreamAccess();
	my $df = $dfAccess->findByUrl("http://ssds-test.shore.mbari.org/ssds/rawpackets/1313_0_1_1327");
	print "isWebAccessible = " . ${$df}[0]->webAccessible() . "\n";

=cut

sub findByUrl {
	my $obj = shift;
	my $string = shift;
	
	my $url = $obj->baseUrl . "className=DataStreamAccess&methodName=findByUrl&param1Type=String&param1Value=" . $string;	
	return $obj->_createSSDSobjects($url);
}


#====================================================================
#                             Deployment
#====================================================================
package SSDS::Deployment;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;
use Time::Local;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);
		
attributes qw(id name description startDate endDate role nominalLatitude nominalLongitude nominalDepth xOffset yOffset zOffset x3DOrientationText);


=back 

=head3 Deployment

Deployment - An SSDS Deployment object

=over 4

=item B<new()>

Instantiate an SSDS::Deployment object

The following methods may be used to get or set any of the attributes of a 
Deployment object (each attribute is both a setter and getter):

	id name description startDate endDate role nominalLatitude nominalLongitude 
	nominalDepth xOffset yOffset zOffset x3DOrientationText

Note: the get_attribute_names() method will return this list of attributes

Example:

	my $depl = $deplAccess->findByPK(4322);
	print "NominalLongitude = " . $depl->nominalLongitude . "\n";
	$depl->nominalDepth(20.0);	# Set nominal depth of Deployment to 20 m
	
For setting startDate and endDate use ISO 8601 date/time format, e.g.:

	$depl->startDate("2005-08-25T14:00:00");
	$depl->endDate("2005-08-25T18:00:00");
	
As of August 2005 dates are transmitted without conversion to the database
and all the dates are local time in the SQL database as of this date.

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}


#--------------------------------------------------------------------
# 

=item B<childDeployments()>

return list of child deployments

Example:

	my $deplAccess = new SSDS::DeploymentAccess();
	my $depl = $deplAccess->findByPK(4309);
	my @deplAttr = ($depl->get_attribute_names());
	my $childDepls = $depl->childDeployments();
	foreach my $cd (@{$childDepls}) {
		foreach my $a (@deplAttr) {
			next unless $cd->$a;
			print "  $a = " . $cd->$a . "\n";
		}
		print"\n";
	}

=cut

sub childDeployments {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=Deployment&methodName=listChildDeployments&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobjects($url);
}

#--------------------------------------------------------------------
# 

=item B<getDevice()>

return the Device object for this Deployment

Example:

	my $dev = $id->getDevice($id->id());
	print "Instrument Deployment " . $id->name() . "\n";
	print "  of Device   " . $dev->name() . "\n";

=cut

sub getDevice {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=Deployment&methodName=getDevice&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobject($url);
}

#--------------------------------------------------------------------
# 

=item B<listOutputs()>

return list of output DataStreams or DataFiles

Example:

	$outputs = $depl->listOutputs();
	@dfAttr = ${$outputs}[0]->get_attribute_names() if $outputs;
	foreach my $o (@{$outputs}) {
		foreach my $a (@dfAttr) {
			next unless $o->$a;
			print "  $a = " . $o->$a . "\n";
		}
		print"\n";
	}

=cut

sub listOutputs {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=Deployment&methodName=listOutputs&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobjects($url);
}


#--------------------------------------------------------------------
# 

=item B<getParentDeployment()>

Return parent Deployment

Example:

	my $depl = $deplAccess->findByPK(4322);
	my $parentDepl = $depl->getParentDeployment();
	foreach my $a (@deplAttr) {
		next unless $parentDepl->$a;
		print "  $a = " . $parentDepl->$a . "\n";
	}

=cut

sub getParentDeployment {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=Deployment&methodName=getParentDeployment&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobject($url);
}

#--------------------------------------------------------------------
# 

=item B<getStartEsecs()>

Return the Unix Epoch seconds of the start time for this Deployment.
If there is no start time defined then nothing is returned.

Example:

	print "    startDate = " . $depl->startDate() . "\n";
	print "    start epoch seconds = " . $depl->getStartEsecs() . "\n";

=cut

sub getStartEsecs {
	my $obj = shift;
	
	my %monNumbers = ( Jan => 0, Feb => 1, Mar => 2, Apr => 3, May => 4, Jun => 5,
			   Jul => 6, Aug => 7, Sep => 8, Oct => 9, Nov => 10, Dec => 11 );
			
	my ($se, $mn, $hr, $da, $mo, $yr);
	
	# Sample date format:  Wed Oct 06 11:04:32 PDT 2004
	my $dateStr = $obj->startDate();
	if ( $obj->debug == $_debugLevel{'INFO'} ) {
		print "getStartEsecs(): dateStr = $dateStr\n";
	}
	return unless $dateStr;
	
	# First try stardard DB return format
	$dateStr =~ /.+\s+(.+)\s+(\d\d)\s+(\d\d):(\d\d):(\d\d)\s+.+\s+(\d\d\d\d)/;
	$mo = $monNumbers{$1};
	$da = $2;
	$hr = $3;
	$mn = $4;
	$se = $5;
	$yr = $6;
	
	# $da and $yr must be > 0; if not then parse assuming ISO 8601 format
	unless ( $da && $yr ) {
		if ( $obj->debug == $_debugLevel{'INFO'} ) {
			print "Parsing start date string with ISO 8601 pattern\n";
		}
		$dateStr =~ /(\d\d\d\d)-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d).+/
		|| $dateStr =~ /(\d\d\d\d)-(\d\d)-(\d\d)T(\d\d):(\d\d).+/ ;
		$yr = $1;
		$mo = (defined $2) ? $2 - 1 : '';
		$da = $3;
		$hr = $4;
		$mn = $5;
		$se = (defined $6) ? $6 : 0;
	}
	
	if ( $obj->debug == $_debugLevel{'INFO'} ) {
		print "$se, $mn, $hr, $da, $mo, $yr\n";
	}
	my $esecs = timelocal($se, $mn, $hr, $da, $mo, $yr);
	
	return $esecs;
}


#--------------------------------------------------------------------
# 

=item B<getEndEsecs()>

Return the Unix Epoch seconds of the end time for this Deployment.
If there is no start time defined then nothing is returned.

Example:

	print "    endDate = " . $depl->endDate() . "\n" if $depl->endDate();
	print "    end epoch seconds = " . $depl->getEndEsecs() . "\n" if $depl->getEndEsecs();

=cut

sub getEndEsecs {
	my $obj = shift;
	
	my %monNumbers = ( Jan => 0, Feb => 1, Mar => 2, Apr => 3, May => 4, Jun => 5,
			   Jul => 6, Aug => 7, Sep => 8, Oct => 9, Nov => 10, Dec => 11 );
			
	my ($se, $mn, $hr, $da, $mo, $yr);
	
	# Sample date format:  Wed Oct 06 11:04:32 PDT 2004
	my $dateStr = $obj->endDate();
	if ( $obj->debug == $_debugLevel{'INFO'} ) {
		print "getStartEsecs(): dateStr = $dateStr\n";
	}
	return unless $dateStr;
	
	$dateStr =~ /.+\s+(.+)\s+(\d\d)\s+(\d\d):(\d\d):(\d\d)\s+.+\s+(\d\d\d\d)/;
	$mo = $monNumbers{$1};
	$da = $2;
	$hr = $3;
	$mn = $4;
	$se = $5;
	$yr = $6;
	
	# $da and $yr must be > 0; if not then parse assuming ISO 8601 format
	unless ( $da && $yr ) {
		if ( $obj->debug == $_debugLevel{'INFO'} ) {
			print "Parsing end date string with ISO 8601 pattern\n";
		}
		$dateStr =~ /(\d\d\d\d)-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d).+/
		|| $dateStr =~ /(\d\d\d\d)-(\d\d)-(\d\d)T(\d\d):(\d\d).+/ ;
		$yr = $1;
		$mo = (defined $2) ? $2 - 1 : '';
		$da = $3;
		$hr = $4;
		$mn = $5;
		$se = (defined $6) ? $6 : 0;
	}
	
	if ( $obj->debug == $_debugLevel{'INFO'} ) {
		print "$se, $mn, $hr, $da, $mo, $yr\n";
	}
	my $esecs = timelocal($se, $mn, $hr, $da, $mo, $yr);
	return $esecs;
}


#--------------------------------------------------------------------
# 

=item B<setDevice()>

Set the Device object for this Deployment

Example:
	
	# Insert a test Device for this test Deployment
	my $devAccess = new SSDS::DeviceAccess();
	my $dev = new SSDS::Device();
	my $uniqueDevName = "SSDS.pm test Device";
	$dev->name($uniqueDevName);
	$dev->description("Test Device - to be deleted");
	$dev->preferredDeploymentRole("platform");
	$devAccess->insert($dev);
	
	# Find the device just inserted so that we have an ID
	my $foundDevs = $devAccess->findByLikeName($uniqueDevName);
	my $foundDev = $$foundDevs[0];

	# Insert a Deployment
	my $deplAccess = new SSDS::DeploymentAccess();
	my $depl = new SSDS::Deployment();
	my $uniqueDeplName = "SSDS.pm test Deployment";
	$depl->name($uniqueDeplName);
	$deplAccess->insert($depl);
	
	# Find the Deployment just inserted and set the Device 
	my $foundDepls = $deplAccess->findByName($uniqueDeplName);
	my $foundDepl = $$foundDepls[0];	
	$foundDepl->setDevice($foundDev);
	
=cut

sub setDevice {
	my $obj = shift;
	my $dev = shift;
	my $url = $obj->baseUrl . "className=Deployment&methodName=setDevice&findBy=PK&findByType=String&findByValue=" . $obj->id;
	$url .= "&param1Type=Device&param1Value=Device";
	foreach my $a ($dev->get_attribute_names()) {
		$url .= $obj->delim() . $a . "=" . $dev->$a if $dev->$a;
	}
	return $obj->_execSSDSmethod($url);
}


#====================================================================
#                          DeploymentAccess
#====================================================================
package SSDS::DeploymentAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 DeploymentAccess

DeploymentAccess - Contains methods to query for and get Deployment objects


=over 4

=item B<new()>

Instantiate an SSDS::DeploymentAccess object

Example:

	my $deplAccess = new SSDS::DeploymentAccess();
	$depl = $deplAccess->findByPK(4322);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}


#--------------------------------------------------------------------
# 

=item B<insert($depl)>

Insert into the database the Deployment $depl

Example:

	# This is pathological as no device is associated with this Deployment
	# (See Deployment->setDevice() for how to set a device.)
	my $deplAccess = new SSDS::DeploymentAccess();
	my $depl = new SSDS::Deployment();
	$depl->name("SSDS.pm test Deployment");
	$depl->description("Test Deployment - to be deleted");
	$depl->startDate("2005-08-23T07:00:00");
	$depl->endDate("2005-08-23T11:00:00");
	$depl->role("platform");
	$deplAccess->insert($depl);

=cut

sub insert {
	my $obj = shift;
	my $depl = shift;
	my $url = $obj->baseUrl . "className=DeploymentAccess&methodName=insert&param1Type=Deployment&param1Value=Deployment";
	foreach my $a ($depl->get_attribute_names()) {
		$url .= $obj->delim() . $a . "=" . $depl->$a if $depl->$a;
	}
	return $obj->_execSSDSmethod($url);
}


#--------------------------------------------------------------------
# 

=item B<delete($depl)>

Delete Deployment $depl from the database

Example:

	my $deplAccess = new SSDS::DeploymentAccess();
	my $depl = $deplAccess->findByPK(1);
	$deplAccess->delete($depl);

=cut

sub delete {
	my $obj = shift;
	my $depl = shift;
	my $url = $obj->baseUrl . "className=DeploymentAccess&methodName=delete&param1Type=Deployment&param1Value=Deployment";
	$url .= $obj->delim() . "id=" . $depl->id() if $depl->id();
	return $obj->_execSSDSmethod($url);
}


#--------------------------------------------------------------------
# 

=item B<update($depl)>

Update Deployment $depl in database

Example:

	my $deplAccess = new SSDS::DeploymentAccess();
	my $depl = $deplAccess->findByPK(4322);
	$deplAccess->update($depl);

=cut

sub update {
	my $obj = shift;
	my $depl = shift;
	my $url = $obj->baseUrl . "className=DeploymentAccess&methodName=update&param1Type=Deployment&param1Value=Deployment";
	foreach my $a ($depl->get_attribute_names()) {
		$url .= $obj->delim() . $a . "=" . $depl->$a if $depl->$a;
	}
	return $obj->_execSSDSmethod($url);
}



#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the Deployment that has ID = pk.

Example:

	my $deplAccess = new SSDS::DeploymentAccess();
	my $depl = $deplAccess->findByPK(4322);
	print "NominalLongitude = " . $depl->nominalLongitude . "\n";

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	my $url = $obj->baseUrl . "className=DeploymentAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;
	return $obj->_createSSDSobject($url);
}


#--------------------------------------------------------------------
# 

=item B<findByName(name)>

return a list of Deployments that match the name.

Example:

	my $deplAccess = new SSDS::DeploymentAccess();
	my @depl = $deplAccess->findByName("CIMT Mooring Deployment");

=cut

sub findByName {
	my $obj = shift;
	my $name = shift;
	my $url = $obj->baseUrl . "className=DeploymentAccess&methodName=findByName&param1Type=String&param1Value=" . $name;
	
	return $obj->_createSSDSobjects($url);
}


#--------------------------------------------------------------------
# 

=item B<findBySQL(sql)>

return a list of Deployments resulting from the SQL statement in sql.

Use of this method requires knowledge of the SSDS database schema.  Very selective queries can be constructed to 
return just the Deployments that satisfy a particular set of conatraints.  For example, to find the AUV deployments
during a Zephyr expedition this would be the query:


This level of query expression is not possible with the OJB Criteria methods.

Example:

	my $deplAccess = new SSDS::DeploymentAccess();
	my @rdepls = $deplAccess->findBySQL("SELECT * FROM DataProducer WHERE class_name = 'moos.ssds.model.Deployment' AND StartDateTime BETWEEN '7/1/2003' AND '7/17/2003'");

=cut

sub findBySQL {
	my $obj = shift;
	my $name = shift;
	my $url = $obj->baseUrl . "className=DeploymentAccess&methodName=findBySQL&param1Type=String&param1Value=" . $name;
	
	return $obj->_createSSDSobjects($url);
}


#====================================================================
#                              Device
#====================================================================
package SSDS::Device;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);

attributes qw(id name description mfgName mfgModel mfgSerialNumber preferredDeploymentRole uuid); 

=back 

=head3 Device

Device - An SSDS Device object


=over 4

=item B<new()>

Instantiate an SSDS::Device object

The following methods may be used to get or set any of the attributes of a 
Device object (each attribute is both a setter and getter):

   id name description mfgName mfgModel mfgSerialNumber preferredDeploymentRole uuid

Note: the get_attribute_names() method will return this list of attributes

Example:
	
	my $devAccess = new SSDS::DeviceAccess();
	$dev = $devAccess->findByPK(1319);
	print "mfgModel = " . $dev->mfgModel . "\n";
	$dev->mfgSerialNumber(1234);	# Set manufacturer serial number to 1234

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}


#--------------------------------------------------------------------
# 

=item B<getDeviceType()>

DeviceType is a special attribute.  It comes from a database lookup table via a DeviceType object.

Example:

	my $devAccess = new SSDS::DeviceAccess();
	$dev = $devAccess->findByPK(1319);
	print "DeviceType name = " . $dev->getDeviceType()->name() . "\n"

=cut

sub getDeviceType {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=Device&methodName=getDeviceType&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobject($url);
}


#====================================================================
#                           DeviceAccess
#====================================================================
package SSDS::DeviceAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 DeviceAccess

DeviceAccess - Contains methods to query for and get Device objects


=over 4

=item B<new()>

Instantiate an SSDS::DeviceAccess object

Example:
	
	my $devAccess = new SSDS::DeviceAccess();
	$dev = $devAccess->findByPK(1319);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}


#--------------------------------------------------------------------
# 

=item B<insert($dev)>

Insert into the database the Device $dev

Example:

	my $devAccess = new SSDS::DeviceAccess();
	my $dev = new SSDS::Device();
	$dev->name("SSDS.pm test Device");
	$dev->description("Test Device - to be deleted");
	$dev->preferredDeploymentRole("platform");
	$devAccess->insert($dev);

=cut

sub insert {
	my $obj = shift;
	my $dev = shift;
	my $url = $obj->baseUrl . "className=DeviceAccess&methodName=insert&param1Type=Device&param1Value=Device";
	foreach my $a ($dev->get_attribute_names()) {
		$url .= $obj->delim() . $a . "=" . $dev->$a if $dev->$a;
	}
	return $obj->_execSSDSmethod($url);
}


#--------------------------------------------------------------------
# 

=item B<delete($dev)>

Delete Device $dev from the database

Example:

	my $devAccess = new SSDS::DeviceAccess();
	my $dev = $devAccess->findByPK(1);
	$devAccess->delete($dev);

=cut

sub delete {
	my $obj = shift;
	my $dev = shift;
	my $url = $obj->baseUrl . "className=DeviceAccess&methodName=delete&param1Type=Device&param1Value=Device";
	$url .= $obj->delim() . "id=" . $dev->id() if $dev->id();
	return $obj->_execSSDSmethod($url);
}


#--------------------------------------------------------------------
# 

=item B<update($dev)>

Update Device $dev in database

Example:

	my $devAccess = new SSDS::DeviceAccess();
	my $dev = $devAccess->findByPK(4322);
	$devAccess->update($dev);

=cut

sub update {
	my $obj = shift;
	my $dev = shift;
	my $url = $obj->baseUrl . "className=DeviceAccess&methodName=update&param1Type=Device&param1Value=Device";
	foreach my $a ($dev->get_attribute_names()) {
		$url .= $obj->delim() . $a . "=" . $dev->$a if $dev->$a;
	}
	return $obj->_execSSDSmethod($url);
}



#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the Device with PK = pk

Example:

	my $devAccess = new SSDS::DeviceAccess();
	$dev = $devAccess->findByPK(1319);

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	
	my $url = $obj->baseUrl . "className=DeviceAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
}


#--------------------------------------------------------------------
# 

=item B<findByLikeName(string)>

return a list of Devicess where string matches any part of the Device name.

Example:

	$devs = $devAccess->findByLikeName("CTD");

=cut

sub findByLikeName {
	my $obj = shift;
	my $string = shift;
	
	my $url = $obj->baseUrl . "className=DeviceAccess&methodName=findByLikeName&param1Type=String&param1Value=" . $string;	
	return $obj->_createSSDSobjects($url);
}


#====================================================================
#                           DeviceType
#====================================================================
package SSDS::DeviceType;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);

attributes qw(id name description defaultDeploymentRole displayInDeviceTypesPicklist); 

=back 

=head3 DeviceType

DeviceType - An SSDS DeviceType object


=over 4

=item B<new()>

Instantiate an SSDS::DeviceType object

The following methods may be used to get or set any of the attributes of a 
DeviceType object (each attribute is both a setter and getter):

    id name description defaultDeploymentRole displayInDeviceTypesPicklist

Note: the get_attribute_names() method will return this list of attributes

Example:

	my $dcAccess = new SSDS::DeviceTypeAccess();
	my $dc = $dcAccess->findByPK(5777);

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}


#====================================================================
#                           DeviceTypeAccess
#====================================================================
package SSDS::DeviceTypeAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 DeviceTypeAccess

DeviceTypeAccess - Contains methods to query for and get DeviceType objects


=over 4

=item B<new()>

Instantiate an SSDS::DeviceTypeAccess object

Example:

	my $dtAccess = new SSDS::DeviceTypeAccess();
	my $dt = $dtAccess->findByPK(102);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}

#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the DeviceType with PK = pk

Example:

	my $dtAccess = new SSDS::DeviceTypeAccess();
	my $dt = $dtAccess->findByPK(110);

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	my $url = $obj->baseUrl . "className=DeviceTypeAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
}


#--------------------------------------------------------------------
# 

=item B<findByName(string)>

return the DeviceType with name = string

Example:

	my $dtAccess = new SSDS::DeviceTypeAccess();
	my $dt = $dtAccess->findByName("CTD");

=cut

sub findByName {
	my $obj = shift;
	my $string = shift;
	my $url = $obj->baseUrl . "className=DeviceTypeAccess&methodName=findByName&param1Type=String&param1Value=" . $string;	
	return $obj->_createSSDSobject($url);
}


#--------------------------------------------------------------------
# 

=item B<insert($dt)>

Insert into the database the DeviceType $dt

Example:

	my $dtAccess = new SSDS::DeviceTypeAccess();
	my $dt = new SSDS::DeviceType();
	$dt->name("IMCTD");
	$dt->description("Inductive Modem CTD");
	$dt->defaultDeploymentRole("instrument");
	$dt->displayInDeviceTypesPicklist("true");
	$dtAccess->insert($dt);

=cut

sub insert {
	my $obj = shift;
	my $dt = shift;
	my $url = $obj->baseUrl . "className=DeviceTypeAccess&methodName=insert&param1Type=DeviceType&param1Value=DeviceType";
	foreach my $a ($dt->get_attribute_names()) {
		$url .= $obj->delim() . $a . "=" . $dt->$a if $dt->$a;
	}
	return $obj->_execSSDSmethod($url);
}


#--------------------------------------------------------------------
# 

=item B<delete($dt)>

Delete DeviceType $dt from the database

Example:

	my $dtAccess = new SSDS::DeviceTypeAccess();
	my $dt = $dtAccess->findByName("IMCTD");
	$dtAccess->delete($dt);

=cut

sub delete {
	my $obj = shift;
	my $dt = shift;
	my $url = $obj->baseUrl . "className=DeviceTypeAccess&methodName=delete&param1Type=DeviceType&param1Value=DeviceType";
	$url .= $obj->delim() . "id=" . $dt->id() if $dt->id();
	return $obj->_execSSDSmethod($url);
}


#====================================================================
#                           RecordDescription
#====================================================================
package SSDS::RecordDescription;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);

attributes qw(id recordType bufferStyle bufferLengthType bufferItemSeparator recordTerminator endian parseable); 

=back 

=head3 RecordDescription

RecordDescription - An SSDS RecordDescription object


=over 4

=item B<new()>

Instantiate an SSDS::RecordDescription object

The following methods may be used to get or set any of the attributes of a 
RecordDescription object (each attribute is both a setter and getter):

    id recordType bufferStyle bufferLengthType bufferItemSeparator recordTerminator endian parseable

Note: the get_attribute_names() method will return this list of attributes

Example:

	my $rdAccess = new SSDS::RecordDescriptionAccess();
	my $rd = $rdAccess->findByPK(101);

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}


#--------------------------------------------------------------------
# 

=item B<listRecordVariables()>

Return list of RecordVariables from this RecordDescription

Example:

	$rvs = $rd->listRecordVariables();
	@rvAttr = ${$rvs}[0]->get_attribute_names() if $outputs;
	foreach my $o (@{$rvs}) {
		foreach my $a (@rvAttr) {
			next unless $o->$a;
			print "  $a = " . $o->$a . "\n";
		}
		print"\n";
	}

=cut

sub listRecordVariables {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=RecordDescription&methodName=listRecordVariables&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobjects($url);
}


#====================================================================
#                           RecordDescriptionAccess
#====================================================================
package SSDS::RecordDescriptionAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 RecordDescriptionAccess

RecordDescriptionAccess - Contains methods to query for and get RecordDescription objects


=over 4

=item B<new()>

Instantiate an SSDS::RecordDescriptionAccess object

Example:

	my $rdAccess = new SSDS::RecordDescriptionAccess();
	my $rd = $rdAccess->findByPK(102);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}


#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the RecordDescription with PK = pk

Example:

	my $rdAccess = new SSDS::RecordDescriptionAccess();
	my $rd = $rdAccess->findByPK(110);

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	my $url = $obj->baseUrl . "className=RecordDescriptionAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
}


#====================================================================
#                           RecordVariable
#====================================================================
package SSDS::RecordVariable;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);

attributes qw(id name longName description format units columnIndex validMin validMax missingValue parseRegExp); 

=back 

=head3 RecordVariable

RecordVariable - An SSDS RecordVariable object


=over 4

=item B<new()>

Instantiate an SSDS::RecordVariable object

The following methods may be used to get or set any of the attributes of a 
RecordVariable object (each attribute is both a setter and getter):

    id name longName description format units columnIndex validMin validMax missingValue parseRegExp

Note: the get_attribute_names() method will return this list of attributes

Example:

	my $rvAccess = new SSDS::RecordVariableAccess();
	my $rv = $rvAccess->findByPK(101);

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}


#--------------------------------------------------------------------
# 

=item B<getStandardVariable()>

Return the StandardVariable object for this RecordVariable

Example:

	my $standard_name = $rv->getStandardVariable->name();


=cut

sub getStandardVariable {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=RecordVariable&methodName=getStandardVariable&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobject($url);
}


#====================================================================
#                           RecordVariableAccess
#====================================================================
package SSDS::RecordVariableAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 RecordVariableAccess

RecordVariableAccess - Contains methods to query for and get RecordVariable objects


=over 4

=item B<new()>

Instantiate an SSDS::RecordVariableAccess object

Example:

	my $rvAccess = new SSDS::RecordVariableAccess();
	my $rv = $rvAccess->findByPK(102);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}

#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the RecordVariable with PK = pk

Example:

	my $rvAccess = new SSDS::RecordVariableAccess();
	my $rv = $rvAccess->findByPK(110);

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	my $url = $obj->baseUrl . "className=RecordVariableAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
}

#--------------------------------------------------------------------
# 

=item B<findAll()>

Return all RecordVariables in the database.

Example:

	$rvs = $drvAccess->findAll();

=cut

sub findAll {
	my $obj = shift;
	my $string = shift;
	
	my $url = $obj->baseUrl . "className=RecordVariableAccess&methodName=findAll";	
	return $obj->_createSSDSobjects($url);
}






#====================================================================
#                           Resource
#====================================================================
package SSDS::Resource;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);

attributes qw(id name url description resourceType contentLength MIMEType webAccessible startDate endDate); 

=back 

=head3 Resource

Resource - An SSDS Resource object


=over 4

=item B<new()>

Instantiate an SSDS::Resource object

The following methods may be used to get or set any of the attributes of a 
Resource object (each attribute is both a setter and getter):

    id name url description resourceType contentLength MIMEType webAccessible startDate endDate

Note: the get_attribute_names() method will return this list of attributes

Example:

	my $resAccess = new SSDS::ResourceAccess();
	my $res = $resAccess->findByPK(101);

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}


#====================================================================
#                           ResourceAccess
#====================================================================
package SSDS::ResourceAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 ResourceAccess

ResourceAccess - Contains methods to query for and get Resource objects


=over 4

=item B<new()>

Instantiate an SSDS::ResourceAccess object

Example:

	my $resAccess = new SSDS::ResourceAccess();
	my $res = $resAccess->findByPK(102);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}

#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the Resource with PK = pk

Example:

	my $resAccess = new SSDS::ResourceAccess();
	my $res = $resAccess->findByPK(110);

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	my $url = $obj->baseUrl . "className=ResourceAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
}


#--------------------------------------------------------------------
# 

=item B<findByLikeName(string)>

return a list of Resources where string matches any part of the Resource name.

Example:

	$ress = $resAccess->findByLikeName("CTD");

=cut

sub findByLikeName {
	my $obj = shift;
	my $string = shift;
	
	my $url = $obj->baseUrl . "className=ResourceAccess&methodName=findByLikeName&param1Type=String&param1Value=" . $string;	
	return $obj->_createSSDSobjects($url);
}



#====================================================================
#                           StandardVariable
#====================================================================
package SSDS::StandardVariable;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);

attributes qw(id name referenceScale displayInPickList description); 

=back 

=head3 StandardVariable

StandardVariable - An SSDS StandardVariable object


=over 4

=item B<new()>

Instantiate an SSDS::StandardVariable object

The following methods may be used to get or set any of the attributes of a 
StandardVariable object (each attribute is both a setter and getter):

    id name referenceScale displayInPickList description

Note: the get_attribute_names() method will return this list of attributes

Example:

	my $svAccess = new SSDS::StandardVariableAccess();
	my $sv = $svAccess->findByPK(101);

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}


#====================================================================
#                           StandardVariableAccess
#====================================================================
package SSDS::StandardVariableAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 StandardVariableAccess

StandardVariableAccess - Contains methods to query for and get StandardVariable objects


=over 4

=item B<new()>

Instantiate an SSDS::StandardVariableAccess object

Example:

	my $svAccess = new SSDS::StandardVariableAccess();
	my $sv = $svAccess->findByPK(102);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}

#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the StandardVariable with PK = pk

Example:

	my $svAccess = new SSDS::StandardVariableAccess();
	my $sv = $svAccess->findByPK(110);

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	my $url = $obj->baseUrl . "className=StandardVariableAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
}


#====================================================================
#                           StandardUnit
#====================================================================
package SSDS::StandardUnit;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);

attributes qw(id name longName symbol displayInPickList Description); 

=back 

=head3 StandardUnit

StandardUnit - An SSDS StandardUnit object


=over 4

=item B<new()>

Instantiate an SSDS::StandardUnit object

The following methods may be used to get or set any of the attributes of a 
StandardUnit object (each attribute is both a setter and getter):

    id name longName symbol displayInPickList Description

Note: the get_attribute_names() method will return this list of attributes

Example:

	my $suAccess = new SSDS::StandardUnitAccess();
	my $su = $suAccess->findByPK(101);

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}


#====================================================================
#                           StandardUnitAccess
#====================================================================
package SSDS::StandardUnitAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 StandardUnitAccess

StandardUnitAccess - Contains methods to query for and get StandardUnit objects


=over 4

=item B<new()>

Instantiate an SSDS::StandardUnitAccess object

Example:

	my $suAccess = new SSDS::StandardUnitAccess();
	my $su = $suAccess->findByPK(102);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}

#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the StandardUnit with PK = pk

Example:

	my $suAccess = new SSDS::StandardUnitAccess();
	my $su = $suAccess->findByPK(110);

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	my $url = $obj->baseUrl . "className=StandardUnitAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
}




#====================================================================
#                           ProcessRun
#====================================================================
package SSDS::ProcessRun;
use LWP::Simple;
use strict;
use Carp;
use Class::ObjectTemplate;

our @ISA = qw(SSDS::ObjectCreator Class::ObjectTemplate);

attributes qw(id name description startDate endDate hostName);

=back 

=head3 ProcessRun

ProcessRun - An SSDS ProcessRun object


=over 4

=item B<new()>

Instantiate an SSDS::ProcessRun object

The following methods may be used to get or set any of the attributes of a 
ProcessRun object (each attribute is both a setter and getter):

    id name description startDate endDate hostName

Note: the get_attribute_names() method will return this list of attributes

Example:

	my $prAccess = new SSDS::ProcessRunAccess();
	my $pr = $prAccess->findByPK(4389);

=cut

sub new {
	my $class = shift;
	my $obj = {};
	bless $obj, $class;
}

#--------------------------------------------------------------------
# 

=item B<addInput()>

Add an existing Input DataStream or DataFile to an existing ProcessRun.  
Both objects must already exist in the database for this to work as it simply associates IDs.

Example:

	# Get a DataStream and add as input
	$dcAccess = new SSDS::DataStreamAccess();
	$dc = $dcAccess->findByPK(4283);
	$p->addInput($dc);

=cut

sub addInput {
	my $obj = shift;
	my $dc = shift;			# A DataContainer
	return unless $dc;
	
	$dc =~ /SSDS::(.+)=.+/;		# Handle both: DataStream and DataFile
	my $dcType = $1;
	
	my $url = $obj->baseUrl . "className=ProcessRun&methodName=addInput&findBy=PK&findByType=String&findByValue=" . $obj->id;
	$url .=  "&param1Type=" . $dcType . "&param1Value=" . $dcType;
	$url .= $obj->delim() . "id=" . $dc->id();
	
	return $obj->_execSSDSmethod($url);
}


#--------------------------------------------------------------------
# 

=item B<addOutput()>

Add an Output DataStream or DataFile to a ProcessRun.  


Example:

	# Get a DataStream and add as input
	$dcAccess = new SSDS::DataStreamAccess();
	$dc = $dcAccess->findByPK(4283);
	$p->addOutput($dc);

=cut

sub addOutput {
	my $obj = shift;
	my $dc = shift;			# A DataContainer
	return unless $dc;
	
	$dc =~ /SSDS::(.+)=.+/;		# Handle both: DataStream and DataFile
	my $dcType = $1;
	
	my $url = $obj->baseUrl . "className=ProcessRun&methodName=addOutput&findBy=PK&findByType=String&findByValue=" . $obj->id;
	$url .=  "&param1Type=" . $dcType . "&param1Value=" . $dcType;
	$url .= $obj->delim() . "id=" . $dc->id();
	
	return $obj->_execSSDSmethod($url);
}


#--------------------------------------------------------------------
# 

=item B<listOutputs()>

Return list of output DataStreams or DataFiles from this ProcessRun

Example:

	$outputs = $pr->listOutputs();
	@dfAttr = ${$outputs}[0]->get_attribute_names() if $outputs;
	foreach my $o (@{$outputs}) {
		foreach my $a (@dfAttr) {
			next unless $o->$a;
			print "  $a = " . $o->$a . "\n";
		}
		print"\n";
	}

=cut

sub listOutputs {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=ProcessRun&methodName=listOutputs&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobjects($url);
}


#--------------------------------------------------------------------
# 

=item B<listInputs()>

Return list of output DataStreams or DataFiles from this ProcessRun

Example:

	$inputs = $pr->listInputs();
	@dfAttr = ${$Inputs}[0]->get_attribute_names() if $inputs;
	foreach my $o (@{$inputs}) {
		foreach my $a (@dfAttr) {
			next unless $o->$a;
			print "  $a = " . $o->$a . "\n";
		}
		print"\n";
	}

=cut

sub listInputs {
	my $obj = shift;
	my $url = $obj->baseUrl . "className=ProcessRun&methodName=listInputs&findBy=PK&findByType=String&findByValue=" . $obj->id;
	return $obj->_createSSDSobjects($url);
}


#====================================================================
#                           ProcessRunAccess
#====================================================================
package SSDS::ProcessRunAccess;

use LWP::Simple;
use strict;
use Carp;

our @ISA = qw(SSDS::ObjectCreator);

=back

=head3 ProcessRunAccess

ProcessRunAccess - Contains methods to query for and get ProcessRun objects


=over 4

=item B<new()>

Instantiate an SSDS::ProcessRunAccess object

Example:

	my $prAccess = new SSDS::ProcessRunAccess();
	my $pr = $prAccess->findByPK(102);

=cut

sub new {
	my $pkg = shift;
	my $obj = $pkg->SUPER::new();
	$obj;
}

#--------------------------------------------------------------------
# 

=item B<findByPK(pk)>

return the ProcessRun with PK = pk

Example:

	my $prAccess = new SSDS::ProcessRunAccess();
	my $pr = $prAccess->findByPK(110);

=cut

sub findByPK {
	my $obj = shift;
	my $pk = shift;
	my $url = $obj->baseUrl . "className=ProcessRunAccess&methodName=findByPK&param1Type=Long&param1Value=" . $pk;	
	return $obj->_createSSDSobject($url);
}


#--------------------------------------------------------------------
# 

=item B<findByName(string)>

return the ProcessRun with name  = string

Example:

	my $prAccess = new SSDS::ProcessRunAccess();
	my $pr = $prAccess->findByName("Test ProcessRun");

=cut

sub findByName {
	my $obj = shift;
	my $string = shift;
	my $url = $obj->baseUrl . "className=ProcessRunAccess&methodName=findByName&param1Type=String&param1Value=" . $string;	
	return $obj->_createSSDSobjects($url);
}


#--------------------------------------------------------------------
# 

=item B<insert($pr)>

Insert into the database the ProcessRun $pr

Example:

	my $prAccess = new SSDS::ProcessRunAccess();
	my $pr = new SSDS::ProcessRun();
	$pr->
	$prAccess->insert($pr);

=cut

sub insert {
	my $obj = shift;
	my $pr = shift;
	my $url = $obj->baseUrl . "className=ProcessRunAccess&methodName=insert&param1Type=ProcessRun&param1Value=ProcessRun";
	foreach my $a ($pr->get_attribute_names()) {
		$url .= $obj->delim() . $a . "=" . $pr->$a if $pr->$a;
	}
	return $obj->_execSSDSmethod($url);
}


#--------------------------------------------------------------------
# 

=item B<delete($pr)>

Delete ProcessRun $pr from the database

Example:

	my $prAccess = new SSDS::ProcessRunAccess();
	my $pr = $prAccess->findByPK(1);
	$prAccess->delete($pr);

=cut

sub delete {
	my $obj = shift;
	my $pr = shift;
	my $url = $obj->baseUrl . "className=ProcessRunAccess&methodName=delete&param1Type=ProcessRun&param1Value=ProcessRun";
	$url .= $obj->delim() . "id=" . $pr->id() if $pr->id();
	return $obj->_execSSDSmethod($url);
}


#--------------------------------------------------------------------
# 

=item B<update($pr)>

Update ProcessRun $pr in database

Example:

	my $prAccess = new SSDS::ProcessRunAccess();
	my $pr = $prAccess->findByPK(1);
	$prAccess->update($pr);

=cut

sub update {
	my $obj = shift;
	my $pr = shift;
	my $url = $obj->baseUrl . "className=ProcessRunAccess&methodName=update&param1Type=ProcessRun&param1Value=ProcessRun";
	foreach my $a ($pr->get_attribute_names()) {
		$url .= $obj->delim() . $a . "=" . $pr->$a if $pr->$a;
	}
	return $obj->_execSSDSmethod($url);
}



1;

__END__


=back

=head1 EXAMPLES

 #
 # Get top level mooring deployment and work down from there to get
 # URLs of output DataStreams
 #
 my $deplAccess = new SSDS::DeploymentAccess();
 my $devAccess = new SSDS::DeviceAccess();
 my $moorDepl = $deplAccess->findByName("CIMT Mooring Deployment");
 foreach my $md ( @{$moorDepl} ) {	
	my $canDepls = $md->childDeployments();
	foreach my $cd ( @{$canDepls} ) {
		my $instDepls = $cd->childDeployments();
		foreach my $id (@{$instDepls}) {
			print "Instrument Deployment " . $id->name() . "\n";
			my $dev = $id->getDevice($id->id());
			print "of Device " . $dev->name() . "\n";
			my $outputs = $id->listOutputs();
			foreach my $o ( @{$outputs} ) {
				my $name = $o->name();
				my $url = $o->url();
				print "  Output file name = $name\n";
				print "  url = $url\n";
			}
			print "\n";
		}
	}
 } 
 
 #
 # Get the variable delimiter for the output from an Instrument Deploymnet
 #
 $delim = ${$id->listOutputs}[0]->getRecordDescription->bufferItemSeparator();
 
 #
 # Get parsing details from the SSDS metadata
 # (ds: DataStream, rd: RecordRescription, rv:RecordVariable)
 #
 my $ds = ${$id->listOutputs}[0];
 print "Input DataStream name: " . $ds->name() . "\n" if $debug;
 foreach my $a ($ds->get_attribute_names()) {
        next unless $ds->$a;
        print "  $a = " . $ds->$a . "\n" if $debug;
 }

 my $rd = $ds->getRecordDescription();
 print "RecordDescription:\n" if $debug;
 foreach my $a ($rd->get_attribute_names()) {
        next unless $rd->$a;
        print "  $a = " . $rd->$a . "\n" if $debug;
 }

 # 
 # Read Record variable attributes into named arrays for use later
 #
 my $rrvs = $rd->listRecordVariables();
 foreach my $a (${$rrvs}[0]->get_attribute_names()) {
        @{$a} = ();
 }
 my $i = 0;
 foreach my $rv ( @{$rrvs} ) {
        print "RecordVariable name: " . $rv->name() . "\n" if $debug;
        foreach my $a ($rv->get_attribute_names()) {
                next unless  $rv->$a;
                print "  $a = " . $rv->$a . "\n" if $debug;
                $$a[$i] = $rv->$a;    # Save variable Attributes in named arrays
        }
        $i++;
 }

=head1 SEE ALSO

Java classes in the moos.ssds.model and moos.ssds.services packages of MBARI's SSDS project.  
Diagrams of the object model and the database are in the poster at 
L<http://www.mbari.org/staff/mccann/papers/AGUPoster2004_AUV_data_management.pdf>.

The Perl package ObjectTemplate (This package is required for SSDS.pm.  for ActiveState Perl's ppm: C<install Class-ObjectTemplate>.)
The version used as of this writing is 0.7.

=head1 BUGS 

Many of the SSDS object attributes are null and these values are simply not set when
SSDS.pm materializes the objects.  This causes many of these type of warnings from the
ObjectTemplate module:

   Use of uninitialized value in numeric eq (==) at (eval 3) line 14.

If you like using the -w option on the perl command then you may want to add this line
to your installation of ObjectTemplate.pm:

    no warnings 'uninitialized';

Not all characters work well as delimiters.  Ones that have been tested include '|' (the default)
and '~'.

=head1 AUTHOR

Mike McCann F<E<lt>mccann@mbari.orgE<gt>>

This module was developed as part of the MBARI 2003-2005 Shore Side Data System project.


http://www.mbari.org/ssds

=cut