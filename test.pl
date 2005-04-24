# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

#########################

# change 'tests => 1' to 'tests => last_test_to_print';

use Test;
BEGIN { plan tests => 1 };
use Win32::File::Summary;
ok(1); # If we made it this far, we're ok.

#########################

# Insert your test code below, the Test module is use()ed here so read
# its man page ( perldoc Test ) for help writing this test script.
my $file = $ARGV[0] || 'data\w.doc';
my $STR = Win32::File::Summary->new($file);
my $iscorOS = $STR->IsWin2000OrNT();
print "This OS is the correct one\n";
my $isStgfile = $STR->IsStgFile();
print "that the file contains a storage object.\n";
my $result = $STR->Read();
if(!$result)
{
	print $STR->GetError() . "\n";
	exit;
}

my %hash = %{ $result };

foreach my $key (keys %hash)
{
	print "$key=" . $hash{$key} . "\n";
}

