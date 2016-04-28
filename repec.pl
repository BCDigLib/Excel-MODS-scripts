#!C:/Perl/bin/perl -w
use strict;
use IO::File;
use File::Basename qw(basename);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use utf8;
use Cwd;
use XML::Simple;
use Data::Dumper;

main();

sub main 
{

	my ($title, $authors, $creationdate, $revisiondate, $abstract, $number, $keywords, $jel, $handle, $url);

	my($worksheet_name, $Sheet, $excel_object) = setup_EXCEL_object(shift); 
	
	##read and process each row in the EXCEL file
	my $usedRange = $Sheet->UsedRange()->{Value};
			
		shift(@$usedRange);

		my $CurrentRow=2;

		while (my $row=shift @$usedRange)
		{
			
			($title, $authors, $creationdate, $revisiondate, $abstract, $number, $keywords, $jel, $handle, $url) = @$row;
			
			my $fh=open_ouput_file($number);
			my $faculty_data = read_faculty_names_xml(); 
			my $jel_data = read_JELCodeLookup_xml();
			
			mods_title($fh, $title);
			mods_name_author($fh, $authors, $faculty_data);
			mods_type_of_resource($fh);
			mods_genre($fh);
			mods_origin_info($fh, $creationdate, $revisiondate);
			mods_language($fh);
			mods_physical_description($fh);
			mods_abstract($fh, $abstract);
			mods_note($fh, $number, $handle, $creationdate, $revisiondate);
			mods_subject_jel($fh, $jel, $jel_data);
			mods_subject_keywords($fh, $keywords);
			mods_related_item($fh, $number);
			mods_identifier($fh, $handle, $url);
			mods_extension($fh); 
			mods_record_info($fh);
			
			close_output_file ($fh);
		};

};


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
$fh->print("<\/mods:titleInfo>\n");

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
$fh->print("<\/mods:titleInfo>\n");
	}


};



### See End of Document for MODS Author Element 



### MODS TypeOfResource Element

sub mods_type_of_resource
{
my $fh = shift;

$fh->print("<mods:typeOfResource>text<\/mods:typeOfResource>\n");

}


### MODS Genre Element

sub mods_genre
{
my $fh = shift;

$fh->print("<mods:genre authority=\"local\" type=\"work type\" usage=\"primary\">working paper<\/mods:genre>\n");

}

### MODS OriginInfo Element

sub mods_origin_info
{
	
my $fh=shift;
my $creationdate = shift;
my $revisiondate = shift;

$fh->print("<mods:originInfo>\n");

	if ($revisiondate) {
		
		my $revisionyear = substr( $revisiondate, 0, 4 ); 
		$fh->print("\t<mods:dateIssued>$revisionyear<\/mods:dateIssued>\n");
		$fh->print("\t<mods:dateIssued encoding=\"w3cdtf\" keyDate=\"yes\">$revisionyear<\/mods:dateIssued>\n");
		$fh->print("\t<mods:edition supplied=\"yes\">Revised edition<\/mods:edition>\n");}
	
	elsif ($creationdate) {
		
		my $creationyear = substr( $creationdate, 0, 4 );
		$fh->print("\t<mods:dateIssued>$creationyear<\/mods:dateIssued>\n");
		$fh->print("\t<mods:dateIssued encoding=\"w3cdtf\" keyDate=\"yes\">$creationyear<\/mods:dateIssued>\n");}
	
	$fh->print("\t<mods:issuance>monographic<\/mods:issuance>\n");

$fh->print("<\/mods:originInfo>\n");
}



### MODS Language Element

sub mods_language
{
my $fh = shift;

$fh->print("<mods:language>\n\t<mods:languageTerm authority=\"iso639-2b\" type=\"text\">English<\/mods:languageTerm>\n\t<mods:languageTerm authority=\"iso639-2b\" type=\"code\">eng<\/mods:languageTerm>\n<\/mods:language>\n");

}



### MODS Physical Description

sub mods_physical_description
{
my $fh = shift;

$fh->print("<mods:physicalDescription>\n");
	$fh->print("\t<mods:form authority=\"marcform\">electronic<\/mods:form>\n");
	$fh->print("\t<mods:internetMediaType>application/pdf<\/mods:internetMediaType>\n");
	$fh->print("\t<mods:digitalOrigin>born digital<\/mods:digitalOrigin>\n");
$fh->print("<\/mods:physicalDescription>\n");

};

### MODS Abstract

sub mods_abstract
{
my $fh = shift;
my $abstract = shift;

$abstract =~ s/\n/\ /g;
$abstract =~ s/\r/\ /g;

$fh->print("<mods:abstract>$abstract<\/mods:abstract>\n");
	


};

### MODS Note

sub mods_note
{
my $fh = shift;
my $number = shift;
my $handle = shift;
my $creationdate = shift;
my $revisiondate = shift;


if ($handle =~ m/RePEc:boc:bocoec:/i )
	{$fh->print("<mods:note>Originally posted on: http:\/\/ideas.repec.org\/p\/boc\/bocoec\/$number.html<\/mods:note>\n");}

if ($creationdate && $revisiondate) {
		my $creationmonth = substr($creationdate, 4, 2);
		
		if ($creationmonth =~ "01") 
			{$creationmonth = "January"}
		elsif ($creationmonth =~ "02") 
			{$creationmonth = "February"}
		elsif ($creationmonth =~ "03") 
			{$creationmonth = "March"}
		elsif ($creationmonth =~ "04") 
			{$creationmonth = "April"}
		elsif ($creationmonth =~ "05") 
			{$creationmonth = "May"}
		elsif ($creationmonth =~ "06") 
			{$creationmonth = "June"} 
		elsif ($creationmonth =~ "07") 
			{$creationmonth = "July"}
		elsif ($creationmonth =~ "08") 
			{$creationmonth = "August"}
		elsif ($creationmonth =~ "09") 
			{$creationmonth = "September"}
		elsif ($creationmonth =~ "10") 
			{$creationmonth = "October"}
		elsif ($creationmonth =~ "11") 
			{$creationmonth = "November"}
		elsif ($creationmonth =~ "12") 
			{$creationmonth = "December"}
		
		my $creationyear = substr($creationdate, 0, 4);
		
		$fh->print("<mods:note>Revised version of working paper originally released in $creationmonth $creationyear.<\/mods:note>\n");}
	
};


### MODS Subject

sub mods_subject_jel

{
my $fh = shift;
my $jel = shift;
my $jel_data = shift;

if ($jel) {

my @jel = split(/\s*,\s*/, $jel);

foreach (@jel)
	{
		
	my $jel_code = $_;
	
	foreach my $e (@{$jel_data->{'JELCodeToValue'}})
		{	
		if ($e->{'code'} eq $jel_code)
			{
			$fh->print("<mods:subject>\n\t<mods:topic>$e->{'value'}<\/mods:topic>\n<\/mods:subject>\n");
			}

		}
	}	
}
	
};


sub mods_subject_keywords

{
my $fh = shift;
my $keywords = shift;

if ($keywords) {

my @keywords = split(/\s*,\s*/, $keywords);

foreach (@keywords){
	
	my $keyword = $_;
	
	$fh->print("<mods:subject>\n\t<mods:topic>$keyword<\/mods:topic>\n<\/mods:subject>\n");}

}	
	
};



### MODS Related Item

sub mods_related_item
{

my ($fh, $number) = @_;

$fh->print("<mods:relatedItem type=\"series\">\n\t<mods:titleInfo usage=\"primary\">");

$fh->print ("\n\t\t<mods:title>Boston College Working Papers in Economics<\/mods:title>\n");	  
if ($number)  {$fh->print("\t\t<mods:partNumber>$number<\/mods:partNumber>\n");}
$fh->print ("\t<\/mods:titleInfo>\n");

$fh->print("<\/mods:relatedItem>\n");	

};

	
### MODS Identifier

sub mods_identifier
{
my $fh = shift;
my $handle = shift;
my $url = shift;

$fh->print("<mods:identifier type=\"repec\">$handle<\/mods:identifier>\n");
$fh->print("<mods:identifier type=\"uri\">$url<\/mods:identifier>\n");

}

### MODS Extension Element

sub mods_extension
{
my $fh = shift;

	$fh->print("<mods:extension>\n\t");
	$fh->print("<localCollectionName>repec<\/localCollectionName>\n");
	$fh->print("<\/mods:extension>\n");
}


### MODS RecordInfo Element

sub mods_record_info
{
my $fh = shift;

$fh->print("<mods:recordInfo>\n");	
	$fh->print("\t<mods:recordContentSource authority=\"marcorg\">MChB<\/mods:recordContentSource>\n");
	$fh->print("\t<mods:languageOfCataloging>\n\t\t<mods:languageTerm type=\"text\" authority=\"iso639-2b\">English<\/mods:languageTerm>\n\t\t<mods:languageTerm type=\"code\" authority=\"iso639-2b\">eng<\/mods:languageTerm>\n\t<\/mods:languageOfCataloging>\n");
$fh->print("<\/mods:recordInfo>\n");


}



### MODS Name Element


sub mods_name_author
{
#Read a tab-delimited line of metadata and assign each element to an appropriately named variable
#
my $fh=shift;
my $authors = shift;
my $faculty_data = shift;
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

	foreach my $e (@{$faculty_data->{'facultyNames'}})  {

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
		$fh->print ("<\/mods:name>\n");

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
	$fh->print ("<\/mods:name>\n");

	}

if ($isBC eq 'false')  {

$fh->print ("<mods:name type=\"personal\" usage=\"primary\">\n\t<mods:namePart type=\"family\">$family_name<\/mods:namePart>\n\t<mods:namePart type=\"given\">$given_name<\/mods:namePart>\n\t<mods:displayForm>$family_name, $given_name<\/mods:displayForm>\n\t<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Author<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>\n\t<\/mods:role>\n<\/mods:name>\n");};

	} 

};


### ### OTHER TASKS


###  Open and Setup Excel


sub setup_EXCEL_object {

#Get the name of the excel workbook and worksheet you want to process
print "\n\nEnter the name of the Excel file containing \nthe data you wish to convert to MODS: ";
my $excelfile = <STDIN>; 
chomp $excelfile; 
exit 0 if (!$excelfile);

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

$fh = IO::File->new('ir-repec-'.$fh.'.xml', 'w')
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
my $faculty_data = $xml->XMLin("facultyNames.xml");

#commenting this block out, cause we've already proved PERL is reading the xml file from ACCESS
#use Data Dumper to confirm xml file was read into perl
#print Dumper($faculty_data);  

return($faculty_data);

};

### Read JELCodeLookup.xml

sub read_JELCodeLookup_xml
{

# create object
my $xml = new XML::Simple;

# read XML file
my $jel_data = $xml->XMLin("JELCodeLookup.xml");

#commenting this block out, cause we've already proved PERL is reading the xml file from ACCESS
#use Data Dumper to confirm xml file was read into perl
#print Dumper($faculty_data);  

return($jel_data);

};

### Close Output File

sub close_output_file{
my $fh=shift;
$fh->print("<\/mods:mods>");
$fh->close();

};
