package Win32::File::Summary;

use 5.006;
use strict;
use warnings;
use Carp;
use Archive::Zip;
use Archive::Zip::MemberRead;
use XML::Simple;
use Text::Wrap;

#use vars qw($VERSION @ISA @EXPORT);

#require Exporter;
#require DynaLoader;
#use AutoLoader;

use base qw/ DynaLoader /;
use vars qw/ $VERSION /;


#our @ISA = qw(Exporter DynaLoader);

# Items to export into callers namespace by default. Note: do not export
# names by default without a very good reason. Use EXPORT_OK instead.
# Do not simply export all your public functions/methods/constants.

# This allows declaration	use Win32::File::Summary ':all';
# If you do not need this, moving things directly into @EXPORT or @EXPORT_OK
# will save memory.
#our %EXPORT_TAGS = ( 'all' => [ qw(
	
#) ] );

#our @EXPORT_OK = ( @{ $EXPORT_TAGS{'all'} } );

#our @EXPORT = qw(
	
#);
our $VERSION = '0.02';

#require XSLoader;
#XSLoader::load('Win32::File::Summary', $VERSION);

bootstrap Win32::File::Summary $VERSION;

# Preloaded methods go here.



# Autoload methods go after =cut, and are processed by the autosplit program.

1;
__END__
# Below is stub documentation for your module. You better edit it!

=head1 NAME

Win32::File::Summary - Perl extension read property informations from MS compound files.
It reads also property informations from \"normal\" files.
(The Properties in OpenOffice documents are directly read from the meta.xml file using
the explanations from the following 
	http://books.evc-cit.info/ch02.php#document-description-fig
	http://books.evc-cit.info/pr01.php#who-should-read-section)


=head1 SYNOPSIS

  use Win32::File::Summary;
  my $Prop = Win32::File::Summary->new($file);
  my $iscorOS = $Prop->IsWin2000OrNT();
  print "This OS is the correct one\n";
  my $isStgfile = $Prop->IsStgFile();
  print "The file contains a storage object.\n";
  my $result = $Prop->Read();
  if(ref($result) eq "SCALAR")
  {
	my $err = $Prop->GetError();
	print "The Error: " . $$err  . "\n";
	exit;
  }

  my %hash = %{ $result };

  foreach my $key (keys %hash)
  {
	print "$key=" . $hash{$key} . "\n";
  }


=head1 DESCRIPTION

The modul Win32::File::Summary can be used to get the summary informations from a MS compound file or normal (text) files.
What are the summary information: 
For compound documents, e.g. Word, you can add Title, Author, Description and some other informations to the document.
The same, but not all of them you can add also to normal (text) files.
This informationes can be read and add in the Property Dialog under the Summary Tab. The module reads these informations.


=head1 FUNCTIONS

=over 4

=item new(file)
 
  This method is the constructor. The only parameter is the filename of the document which informations you want to get.
  
=item IsWin2000OrNT()

   This method returns 1 if the operating system currently used is Windows NT/2000/XP otherwise  0.
   
=item IsStgFile()

  This method returns 1 if that the file contains a storage object, otherwise 0.
  
=item Read()

  This method reads the property set and returns a refernce to a hash which contain the informations.
  If the method fail a scalar reference with the value \"0\" will be returned.
  To check use the following code:
  if(ref($result) eq "SCALAR")
  {
	my $err = $STR->GetError();
	print "The Error: " . $$err  . "\n";
	exit;
  } else
  {
  	my %hash = %{ $result };
  	(Do something with the hash.)
  }

  
=item GetError()

  The GetError method returns the error message (scalar reference).
  The method shall only called if the result from the Read() methode is a scalar reference.

=back

=head1 AUTHOR

Reinhard Pagitsch, E<lt>rpirpag@gmx.atE<gt>

=head1 SEE ALSO

L<perl>.

=TODO

  Adding support for OpenOffice and Star Office documents.
  Adding suport to write summary informations.

=cut
