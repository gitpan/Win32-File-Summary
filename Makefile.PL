use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    'NAME'		=> 'Win32::File::Summary',
    'VERSION_FROM'	=> 'Summary.pm', # finds $VERSION
    'PREREQ_PM'		=> { Archive::Zip => 0, Archive::Zip::MemberRead => 0, XML::Simple => 0, Text::Wrap => 0 }, # e.g., Module::Name => 1.1
    ($] >= 5.005 ?    ## Add these new keywords supported since 5.005
      (ABSTRACT_FROM => 'Summary.pm', # retrieve abstract from module
       AUTHOR     => 'Reinhard Pagitsch <rpirpag@gmx.at>') : ()),
    'LIBS'		=> ['Uuid.lib'], # e.g., '-lm'
    'DEFINE'		=> '', # e.g., '-DHAVE_SOMETHING'
	# Insert -I. if you add *.h files later:
    'INC'		=> '', # e.g., '-I/usr/include/other'
	# Un-comment this if you add C files to link with later:
    #'OBJECT'		=> '$(O_FILES)', # link all the C files too
    'XSOPT'             => '-C++',
    #'XSPROTOARG'	=> '-noprototypes',
    'TYPEMAPS'	=> ['perlobject.map' ],
    'CC'	=> 'cl -TP -EHsc',
);
