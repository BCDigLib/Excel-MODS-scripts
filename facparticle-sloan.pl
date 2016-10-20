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

	my ($wfID, $marcRelatorCode, $authorOrder, $family, $given, $given2, $shortname, $dept, $school, $title, $host, $enum1, $enum2, $chron2, $chron1, $startPage, $endPage, $pageList, $issn, $type, $url ,$doi, $ready, $version, $authors, $digitalOrigin, $accessCondition, $abstract, $titleNew, $class, $fileName);

	my($worksheet_name, $Sheet, $excel_object) = setup_EXCEL_object(shift);

	my $fh=open_ouput_file($worksheet_name);

	my $data = read_faculty_names_xml();

	##read and process each row in the EXCEL file
	my $usedRange = $Sheet->UsedRange()->{Value};
			
		shift(@$usedRange);

		my $CurrentRow=2;


		while (my $row=shift @$usedRange)
		{
			$fh->print("<mods:mods>\n\n");
			($wfID, $authors, $title, $enum1, $chron1, $chron2, $startPage, $endPage, $abstract, $class, $fileName, $accessCondition) = @$row;
			mods_title($fh, $title);
			mods_name_element($fh, $authors, $data);
			mods_type_of_resource($fh);
			mods_genre($fh);
			mods_origin_info($fh, $chron1);
			mods_language($fh);
			mods_physical_description($fh);
			mods_note($fh);
			mods_abstract($fh, $abstract);
			mods_related_item($fh, '1', $class, $enum1, $chron1, $chron2, $startPage, $endPage);
			mods_access_condition($fh, $accessCondition);
			mods_extension($fh, $fileName);
			mods_record_info($fh);

			$fh->print("<\/mods:mods>\n\n");
		};

	
	close_output_file ($fh);


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
	{$nonsort = "The"; 
	$title=$1} 
elsif ($title =~ m/^A (.*)/) 
	{$nonsort = "A";
	$title=$1} 
elsif ($title =~ m /^An (.*)/) 
	{$nonsort = "An";
	$title=$1}; 

$fh->print("<mods:titleInfo>\n");

if ($nonsort) {$fh->print ("\t<mods:nonSort>$nonsort <\/mods:nonSort>\n")};

$fh->print ("\t<mods:title>$title<\/mods:title>\n");

if ($subtitle) 
	{$fh->print ("\t<mods:subTitle>$subtitle<\/mods:subTitle>\n");}
$fh->print("<\/mods:titleInfo>\n\n");

	}

else	{
##Deal with initial articles
my $nonsort;
if ($title =~ m/^The (.*)/) 
	{$nonsort = "The"; 
	$title=$1} 
elsif ($title =~ m/^A (.*)/) 
	{$nonsort = "A";
	$title=$1} 
elsif ($title =~ m /^An (.*)/) 
	{$nonsort = "An";
	$title=$1}; 

$fh->print("<mods:titleInfo>\n");

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

$fh->print("<mods:genre authority=\"marcgt\" type=\"workType\">report<\/mods:genre>\n\n");

}

### MODS OriginInfo Element

sub mods_origin_info
{
my $fh = shift;
my $chron1 = shift;

$fh->print("<mods:originInfo>\n");
	$fh->print("\t<mods:place>\n\t\t<mods:placeTerm type=\"text\">Chestnut Hill, Mass.<\/mods:placeTerm>\n\t<\/mods:place>\n");
	$fh->print("\t<mods:publisher>Sloan Center on Aging &amp; Work at Boston College<\/mods:publisher>\n");
	if ($chron1) {$fh->print("\t<mods:dateIssued>$chron1<\/mods:dateIssued>\n");}
	if ($chron1) {$fh->print("\t<mods:dateIssued encoding=\"w3cdtf\" keyDate=\"yes\">$chron1<\/mods:dateIssued>\n");}
	$fh->print("\t<mods:issuance>monographic<\/mods:issuance>\n");
$fh->print("<\/mods:originInfo>\n\n");
}



### MODS Language Element

sub mods_language
{
my $fh = shift;

$fh->print("<mods:language>\n\t<mods:languageTerm type=\"text\">English<\/mods:languageTerm>\n\t<mods:languageTerm type=\"code\" authority=\"iso639-2b\">eng<\/mods:languageTerm>\n<\/mods:language>\n\n");

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



### MODS Note Element

sub mods_note
{
my $fh = shift;

$fh->print("<mods:note type=\"version identification\">Version of record.<\/mods:note>\n\n");			


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

my ($fh, $version, $class, $enum1, $chron1, $chron2, $startPage, $endPage) = @_;

if ($class)
{

$fh->print("<mods:relatedItem type=\"series\">\n\t<mods:titleInfo>");

$fh->print ("\n\t\t<mods:title>$class<\/mods:title>\n");
$fh->print ("\t<\/mods:titleInfo>\n\t<mods:part>\n");
	  
if ($enum1)  {$fh->print("\t\t<mods:detail level=\"1\" type=\"volume\">\n\t\t\t<mods:number>$enum1<\/mods:number>\n\t\t<\/mods:detail>\n");}

if ($chron2){$fh->print("\t\t<mods:date>$chron2 $chron1<\/mods:date>\n\t<\/mods:part>\n");}
else {$fh->print("\t\t<mods:date>$chron1<\/mods:date>\n\t<\/mods:part>\n");};


$fh->print("<\/mods:relatedItem>\n\n");
};	

};



### MODS Access Condition

sub mods_access_condition
{

my ($fh, $accessCondition) = @_;

	if ($accessCondition) {
		my $fh=shift;
		$fh->print("<mods:accessCondition type=\"useAndReproduction\">$accessCondition<\/mods:accessCondition>\n");
		}
	
	else	{
		my $fh=shift;
		$fh->print("<mods:accessCondition type=\"useAndReproduction\">This work is licensed under the Creative Commons Attribution-NonCommercial 3.0 Unported License (http://creativecommons.org/licenses/by-nc/3.0/).<\/mods:accessCondition>\n\n");
	}

}

### MODS Extension Element

sub mods_extension
{
my ($fh, $fileName) = @_;

	$fh->print("<mods:extension>\n\t");
	$fh->print("<localCollectionName>sloanagingwork<\/localCollectionName>\n\t");
	$fh->print("<ingestFile>$fileName<\/ingestFile>\n");
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

		if ($e->{'naf'} && $e->{'naf'}=~m/\d*/)
			{$fh->print ("<mods:name type=\"personal\">\n\t");}
		else {$fh->print ("<mods:name type=\"personal\">\n\t");}

		$fh->print ("<mods:namePart type=\"family\">$e->{'family'}<\/mods:namePart>\n\t");
		$fh->print ("<mods:namePart type=\"given\">$e->{'given'}<\/mods:namePart>\n\t");
		$fh->print ("<mods:displayForm>$e->{'calc'}<\/mods:displayForm>\n\t");
		$fh->print ("<mods:affiliation>$e->{'DEPT'}, $e->{'SCHL_CD'}<\/mods:affiliation>\n\t");
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
	
	$fh->print ("<mods:name type=\"personal\">\n\t");
	$fh->print ("<mods:namePart type=\"family\">$family_name<\/mods:namePart>\n\t");
	$fh->print ("<mods:namePart type=\"given\">$given_name<\/mods:namePart>\n\t");
	$fh->print ("<mods:displayForm>$family_name, $given_name<\/mods:displayForm>\n\t");
	$fh->print ("<mods:affiliation>Boston College<\/mods:affiliation>\n\t");
	$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Author<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>\n\t<\/mods:role>\n\t");
	$fh->print ("<mods:description>nonfaculty<\/mods:description>\n");
	$fh->print ("<\/mods:name>\n\n");

	}

if ($isBC eq 'false')  {

$fh->print ("<mods:name type=\"personal\">\n\t<mods:namePart type=\"family\">$family_name<\/mods:namePart>\n\t<mods:namePart type=\"given\">$given_name<\/mods:namePart>\n\t<mods:displayForm>$family_name, $given_name<\/mods:displayForm>\n\t<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Author<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>\n\t<\/mods:role>\n<\/mods:name>\n\n");};

	} 

$fh->print ("<mods:name type=\"corporate\">\n\t<mods:namePart>Sloan Center on Aging &amp; Work at Boston College<\/mods:namePart>\n\t<mods:displayForm>Sloan Center on Aging &amp; Work at Boston College<\/mods:displayForm>\n\t<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Issuing body<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">isb<\/mods:roleTerm>\n\t<\/mods:role>\n<\/mods:name>\n\n");


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

$fh = IO::File->new($fh.'.xml', 'w')
	or die "unable to open output file for writing: $!";
binmode($fh, ':utf8');
$fh->print("<?xml version='1.0' encoding='UTF-8' ?>\n");
$fh->print("<mods:modsCollection xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:mods=\"http://www.loc.gov/mods/v3\" xsi:schemaLocation=\"http://www.loc.gov/mods/v3 http://www.loc.gov/standards/mods/v3/mods-3-4.xsd\">\n");

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
$fh->print("<\/mods:modsCollection>\n");
$fh->close();

};
