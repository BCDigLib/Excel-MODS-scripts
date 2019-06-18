use strict;
use IO::File;
use utf8;
use Cwd;
use XML::Simple;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';


my $excelfile;

#Get the name of the excel workbook and worksheet you want to process
print "\n\nEnter the name of the file containing \nthe data you wish to convert to MODS: ";
my $file = <STDIN>; 
chomp $file; 
exit 0 if (!$file);

my $data = read_faculty_names_xml(); 

if (($file =~ m/\.xls/i) and $^O eq "MSWin32" )
	{
		$excelfile=$file;
		process_excel();
	}
elsif (($file =~ m/\.xls/i) and $^O ne "MSWin32")
	{
		die "Can only process excel on a PC, use a text file $!";
	}

else {process_text()}

sub process_text
{

open(my $input_file, '<:encoding(UTF-8)', $file)
  or die "Could not open file '$file' $!";


 
while (my $row = <$input_file>) {
	next if $. < 2;
  	chomp $row;
	my @row = split /\t/, $row;
	foreach(@row)
	{

		if ($_ =~ m /^"(.)*"$/)
		{
			$_ =~ s/^"//;
			$_ =~ s/"$//;
		}
		$_ =~ s/^\s//;
		$_ =~ s/\s$//;


	}
		
	$row = \@row;	
	create_mods($row);}
 


}

sub process_excel 
{

	my($worksheet_name, $Sheet, $excel_object) = setup_EXCEL_object(shift);

	##read and process each row in the EXCEL file
	my $usedRange = $Sheet->UsedRange()->{Value};
			
		shift(@$usedRange);

		my $CurrentRow=2;

		while (my $row=shift @$usedRange)
		{
			create_mods($row);
			
		};
};



sub create_mods
{
my $row=shift;
my ($title, $authors, $genre, $series, $number, $date, $abstract, $keywords, $fileName)=@$row;

		$fileName =~ s/\.pdf//;
		my $fh=open_ouput_file($fileName);
			

			
		mods_title($fh, $title);
		mods_name_element($fh, $authors, $data);
		mods_type_of_resource($fh);
		mods_genre($fh, $genre);
		mods_origin_info($fh, $date);
		mods_language($fh);
		mods_physical_description($fh);
		mods_abstract($fh, $abstract);
		mods_related_item($fh, $genre, $series, $number);
		mods_subject($fh, $keywords);
		mods_access_condition($fh);
		mods_extension($fh, $fileName);
		mods_record_info($fh);

		close_output_file ($fh);
}


### ### LIST OF MODS ELEMENTS


### MODS TitleInfo Element

sub mods_title
{
#Read a tab-delimited line of metadata and assign each element to an appropriately named variable
#
my $fh=shift;
my $title=shift;
my $subtitle;

if ($title =~ m/\&/i )
	{$title =~ s/\&/\&amp;/g;};

if ($title =~ m/\:/i )
	{	
	my ($title, $subtitle) = split (/:\s/, $title, 2);
	my $nonsort;
if ($title =~ m/^The (.*)/) 
	{$nonsort = "The "; 
	$title=$1} 
elsif ($title =~ m/^A (.*)/) 
	{$nonsort = "A ";
	$title=$1} 
elsif ($title =~ m /^An (.*)/) 
	{$nonsort = "An ";
	$title=$1}; 

$fh->print("<mods:titleInfo usage=\"primary\">\n");

if ($nonsort) {$fh->print ("\t<mods:nonSort>$nonsort<\/mods:nonSort>\n")};

$fh->print ("\t<mods:title>$title<\/mods:title>\n");

if ($subtitle) 
	{$fh->print ("\t<mods:subTitle>$subtitle<\/mods:subTitle>\n");}
$fh->print("<\/mods:titleInfo>\n\n");

	}

else	{
##Deal with initial articles
my $nonsort;
if ($title =~ m/^The (.*)/) 
	{$nonsort = "The "; 
	$title=$1} 
elsif ($title =~ m/^A (.*)/) 
	{$nonsort = "A ";
	$title=$1} 
elsif ($title =~ m /^An (.*)/) 
	{$nonsort = "An ";
	$title=$1}; 

$fh->print("<mods:titleInfo usage=\"primary\">\n");

if ($nonsort) {$fh->print ("\t<mods:nonSort>$nonsort<\/mods:nonSort>\n")};

$fh->print ("\t<mods:title>$title<\/mods:title>\n");

if ($subtitle) 
	{$fh->print ("\t<mods:subTitle>$subtitle<\/mods:subTitle>\n");}
$fh->print("<\/mods:titleInfo>\n\n");
	}


};



### See End of Document for MODS Author Element 



### MODS TypeOfResource Element

sub mods_type_of_resource
{
my $fh = shift;
$fh->print("<mods:typeOfResource>text<\/mods:typeOfResource>\n\n");

}


### MODS Genre Element

sub mods_genre
{
my $fh = shift;
my $genre = shift;

$fh->print("<mods:genre authority=\"local\" type=\"work type\" usage=\"primary\">$genre<\/mods:genre>\n\n");

}

### MODS OriginInfo Element

sub mods_origin_info
{
	
my ($fh, $date) = @_;

$fh->print("<mods:originInfo>\n");
	$fh->print("\t<mods:place>\n\t\t<mods:placeTerm type=\"text\">Chestnut Hill, Mass.<\/mods:placeTerm>\n\t<\/mods:place>\n");
	$fh->print("\t<mods:publisher>Center for Retirement Research at Boston College<\/mods:publisher>\n");
	
	if ($date) {
		
		my $month = substr($date, 5, 2);
		
		if ($month =~ "01") 
			{$month = "January"}
		elsif ($month =~ "02") 
			{$month = "February"}
		elsif ($month =~ "03") 
			{$month = "March"}
		elsif ($month =~ "04") 
			{$month = "April"}
		elsif ($month =~ "05") 
			{$month = "May"}
		elsif ($month =~ "06") 
			{$month = "June"} 
		elsif ($month =~ "07") 
			{$month = "July"}
		elsif ($month =~ "08") 
			{$month = "August"}
		elsif ($month =~ "09") 
			{$month = "September"}
		elsif ($month =~ "10") 
			{$month = "October"}
		elsif ($month =~ "11") 
			{$month = "November"}
		elsif ($month =~ "12") 
			{$month = "December"}
			
		my $year = substr($date, 0, 4);
		
		$fh->print("\t<mods:dateIssued>$month $year<\/mods:dateIssued>\n");}
	
	
	if ($date) {$fh->print("\t<mods:dateIssued encoding=\"w3cdtf\" keyDate=\"yes\">$date<\/mods:dateIssued>\n");}
	$fh->print("\t<mods:issuance>monographic<\/mods:issuance>\n");

$fh->print("<\/mods:originInfo>\n\n");
}



### MODS Language Element

sub mods_language
{
my $fh = shift;

$fh->print("<mods:language>\n\t<mods:languageTerm authority=\"iso639-2b\" type=\"text\">English<\/mods:languageTerm>\n\t<mods:languageTerm authority=\"iso639-2b\" type=\"code\">eng<\/mods:languageTerm>\n<\/mods:language>\n\n");

}



### MODS Physical Description

sub mods_physical_description
{
my $fh = shift;

$fh->print("<mods:physicalDescription>\n");
	$fh->print("\t<mods:form authority=\"marcform\">electronic<\/mods:form>\n");
	$fh->print("\t<mods:internetMediaType>application/pdf<\/mods:internetMediaType>\n");
	$fh->print("\t<mods:digitalOrigin>born digital<\/mods:digitalOrigin>\n");
$fh->print("<\/mods:physicalDescription>\n\n");

};


### MODS Physical Description

sub mods_abstract
{
my $fh = shift;
my $abstract = shift;

if ($abstract){$fh->print("<mods:abstract>$abstract<\/mods:abstract>\n\n");}

};


### MODS RelatedItem element

sub mods_related_item
{

my $fh=shift;
my $genre=shift;
my $series=shift;
my $number=shift;

$fh->print("<mods:relatedItem type=\"series\">\n\t<mods:titleInfo usage=\"primary\">");

if ($genre eq "working paper") {
	$fh->print ("\n\t\t<mods:title>CRR WP<\/mods:title>\n");
	if ($number)  {$fh->print("\t\t<mods:partNumber>$number<\/mods:partNumber>\n");}
}

else {
	if ($series) {$fh->print ("\n\t\t<mods:title>$series<\/mods:title>\n");}	  
	if ($number)  {$fh->print("\t\t<mods:partNumber>$number<\/mods:partNumber>\n");}
}

$fh->print ("\t<\/mods:titleInfo>\n");
$fh->print("<\/mods:relatedItem>\n\n");	

};


### MODS Subject

sub mods_subject

{
my $fh = shift;
my $keywords = shift;

if ($keywords){$fh->print("<mods:subject>\n\t<mods:topic>$keywords<\/mods:topic>\n<\/mods:subject>\n\n");}

};



### MODS Access Condition

sub mods_access_condition
{

my ($fh) = @_;

$fh->print("<mods:accessCondition type=\"use and reproduction\">These materials are made available for use in research, teaching and private study, pursuant to U.S. Copyright Law. The user must assume full responsibility for any use of the materials, including but not limited to, infringement of copyright and publication rights of reproduced materials. Any materials used for academic research or otherwise should be fully credited with the source. The publisher or original authors may retain copyright to the materials.<\/mods:accessCondition>\n\n");

}

### MODS Extension Element

sub mods_extension
{
my ($fh, $fileName) = @_;

	$fh->print("<mods:extension>\n\t");
	$fh->print("<localCollectionName>crr<\/localCollectionName>\n\t");
	$fh->print("<ingestFile>ir-crr-$fileName.pdf<\/ingestFile>\n");
	$fh->print("<\/mods:extension>\n\n");
}



### MODS RecordInfo Element

sub mods_record_info
{
my $fh = shift;

$fh->print("<mods:recordInfo>\n");	
	$fh->print("\t<mods:recordContentSource>MChB<\/mods:recordContentSource>\n");
	$fh->print("\t<mods:languageOfCataloging>\n\t\t<mods:languageTerm type=\"text\">English<\/mods:languageTerm>\n\t\t<mods:languageTerm type=\"code\" authority=\"iso639-2b\">eng<\/mods:languageTerm>\n\t<\/mods:languageOfCataloging>\n");
$fh->print("<\/mods:recordInfo>\n\n");


}



### MODS Name Element


sub mods_name_element
{
#Read a tab-delimited line of metadata and assign each element to an appropriately named variable
#
my $fh=shift;
my $authors = shift;
my $data = shift;
my $family;
my $given; 
my $given2;
my $dept;
my $school;

my @authors = split(/\s*;\s*/, $authors);


foreach (@authors) {


my $display_form = $_;
my ($family_name, $given_name) = split(/\s*,\s*/, $display_form);
my $isBC='false';

###### attempt to use username

	foreach my $e (@{$data->{'facultyNames'}})  {

	if ($e->{'shortname'} && $e->{'shortname'} eq $display_form) {
		$isBC='true';

		if ($e->{'naf'} && $e->{'naf'}=~m/\d+/)
			{$fh->print ("<mods:name type=\"personal\" authority=\"naf\" usage=\"primary\">\n\t");}
		else {$fh->print ("<mods:name type=\"personal\" usage=\"primary\">\n\t");}

		$fh->print ("<mods:namePart type=\"family\">$e->{'family'}<\/mods:namePart>\n\t");
		$fh->print ("<mods:namePart type=\"given\">$e->{'given'}<\/mods:namePart>\n\t");
		if ($e->{'year'}) {$fh->print ("<mods:namePart type=\"date\">$e->{'year'}<\/mods:namePart>\n\t");}
		if ($e->{'year'}) {$fh->print ("<mods:displayForm>$e->{'calc'}, $e->{'year'}<\/mods:displayForm>\n\t");}
		else {$fh->print ("<mods:displayForm>$e->{'calc'}<\/mods:displayForm>\n\t");}
		if ($e->{'DEPT'}) {$fh->print ("<mods:affiliation>$e->{'DEPT'}, $e->{'SCHL_CD'}<\/mods:affiliation>\n\t");}
		else {$fh->print ("<mods:affiliation>$e->{'SCHL_CD'}<\/mods:affiliation>\n\t");}
		$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Author<\/mods:roleTerm>\n\t\t");
		$fh->print ("<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>\n\t<\/mods:role>\n\t");
		$fh->print ("<mods:description>$e->{'shortname'}<\/mods:description>\n");
		$fh->print ("<\/mods:name>\n\n");

		}	

}

if ( $display_form =~ m/\(BC\)/i )  {
	$isBC='true';

	$given_name =~ s/ \(BC\)//;
	$display_form =~ s/ \(BC\)//;
	
	$fh->print ("<mods:name type=\"personal\" usage=\"primary\">\n\t");
	$fh->print ("<mods:namePart type=\"family\">$family_name<\/mods:namePart>\n\t");
	$fh->print ("<mods:namePart type=\"given\">$given_name<\/mods:namePart>\n\t");
	$fh->print ("<mods:displayForm>$family_name, $given_name<\/mods:displayForm>\n\t");
	$fh->print ("<mods:affiliation>Boston College<\/mods:affiliation>\n\t");
	$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Author<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>\n\t<\/mods:role>\n\t");
	$fh->print ("<mods:description>nonfaculty<\/mods:description>\n");
	$fh->print ("<\/mods:name>\n\n");

	}

if ($isBC eq 'false')  {

$fh->print ("<mods:name type=\"personal\" usage=\"primary\">\n\t<mods:namePart type=\"family\">$family_name<\/mods:namePart>\n\t<mods:namePart type=\"given\">$given_name<\/mods:namePart>\n\t<mods:displayForm>$family_name, $given_name<\/mods:displayForm>\n\t<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Author<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>\n\t<\/mods:role>\n<\/mods:name>\n\n");};

	} 

$fh->print ("<mods:name type=\"corporate\" authority=\"naf\">\n\t<mods:namePart>Boston College. Center for Retirement Research<\/mods:namePart>\n\t<mods:displayForm>Boston College. Center for Retirement Research<\/mods:displayForm>\n\t<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Issuing body<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">isb<\/mods:roleTerm>\n\t<\/mods:role>\n<\/mods:name>\n\n");


};

### ### OTHER TASKS


###  Open and Setup Excel


sub setup_EXCEL_object {



print "\n\nName of the worksheet containing the \ndata you wish to convert to MODS: ";
my $worksheet_name = <STDIN>; 
chomp $worksheet_name; 
exit 0 if (!$worksheet_name);

my $dir = getcwd;
$dir=~s/\//\\/g;
#print "dir is $dir\n";
$excelfile=$dir."\\".$excelfile;

#Get Ready to use $Win32::OLE

$Win32::OLE::Warn = 3; # Die on Errors.

# ::Warn = 2; throws the errors, but #
# expects that the programmer deals  #

#Create an EXCEL object to work with and define how the object is going to exit

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
        || Win32::OLE->new('Excel.Application', 'Quit');

#Turn off all the alert boxes, such as the SaveAs response "This file already exists", etc. using the DisplayAlerts property.

$Excel->{DisplayAlerts}=0;   

#Open an existing file to work with 
                                                 
my $book_object = $Excel->Workbooks->Open($excelfile);   

#Create a reference to a worksheet object and activate the sheet to give it focus so that actions taken on the workbook or application objects occur on this sheet unless otherwise specified.

my $sheet_object = $book_object->Worksheets($worksheet_name);
$sheet_object->Activate();  

return ($worksheet_name, $sheet_object, $Excel);
}



### Open Output File and Print XML declaration and root node

sub open_ouput_file {

my $fh=shift;

$fh = IO::File->new('ir-crr-'.$fh.'.xml', 'w')
	or die "unable to open output file for writing: $!";
binmode($fh, ':utf8');
$fh->print("<?xml version='1.0' encoding='UTF-8' ?>\n");
$fh->print("<mods:mods xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:mods=\"http://www.loc.gov/mods/v3\" xsi:schemaLocation=\"http://www.loc.gov/mods/v3 http://www.loc.gov/standards/mods/v3/mods-3-4.xsd\">\n");

return($fh);

};


### Read facultyNames.xml

sub read_faculty_names_xml
{

# create object
my $xml = new XML::Simple;

# read XML file
my $data = $xml->XMLin("facultyNames.xml");

#commenting this block out, cause we've already proved PERL is reading the xml file from ACCESS
#use Data Dumper to confirm xml file was read into perl
#print Dumper($data);  

return($data);

};



### Close Output File

sub close_output_file{
my $fh=shift;
$fh->print("<\/mods:mods>\n");
$fh->close();

};

=pod

Usage on a PC
 -- Requires an external data file: facultyNames.xml (This file is exported from the eScholarship workflow database and must be refreshed whenever the eScholarship database is updated with new names)
 -- execute the script:  crr.pl
 -- user will be promted to put in a file name, either a tab delimited text file or an excel spreadsheet is expected
 -- if an Excel file name is input, the user will be prompted to enter the worksheet name (tab)
 -- output is a MODS record for each row of the input file

Usage on  macOS
 -- the Perl module WIN32::OLE is not supported so lines 6 and 7 (use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';) must be deleted before executing the script
  -- Requires an external data file: facultyNames.xml (This file is exported from the eScholarship workflow database and must be refreshed whenever the eScholarship database is updated with new names)
 -- execute the script:  crr.pl
 -- user will be promted to put in a file name, only tab delimited files are expected (Excel files won't work)

betsy.post@bc.edu 20190618
=cut