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

	my ($title, $authors, $advisors, $year, $filename);

	my($worksheet_name, $Sheet, $excel_object) = setup_EXCEL_object(shift);
	
	##read and process each row in the EXCEL file
	my $usedRange = $Sheet->UsedRange()->{Value};
			
		shift(@$usedRange);

		my $CurrentRow=2;

		while (my $row=shift @$usedRange)
		{
			
			($title, $authors, $advisors, $year, $filename) = @$row;
			
			$filename =~ s/.pdf//i;
			my $fh=open_ouput_file($filename);
			
			mods_title($fh, $title);
			mods_name_element_author($fh, $authors);
			mods_name_element_advisor($fh, $advisors);
			mods_type_of_resource($fh);
			mods_genre($fh);
			mods_origin_info($fh, $year);
			mods_language($fh);
			mods_physical_description($fh);
			mods_subject($fh);
			mods_access_condition($fh);
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

$fh->print("<mods:genre authority=\"ndltd\" type=\"work type\">Electronic Thesis or Dissertation<\/mods:genre>\n");
$fh->print("<mods:genre authority=\"dct\" type=\"work type\">Text<\/mods:genre>\n");
$fh->print("<mods:genre authority=\"marcgt\" type=\"work type\" usage=\"primary\">thesis<\/mods:genre>\n\n");
}

### MODS OriginInfo Element

sub mods_origin_info
{
	
my ($fh, $year) = @_;

$fh->print("<mods:originInfo>\n");
	$fh->print("\t<mods:publisher>Boston College<\/mods:publisher>\n");
	if ($year) {$fh->print("\t<mods:dateIssued>$year<\/mods:dateIssued>\n");}
	if ($year) {$fh->print("\t<mods:dateIssued encoding=\"w3cdtf\" keyDate=\"yes\">$year<\/mods:dateIssued>\n");}
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


### MODS Subject

sub mods_subject

{
my $fh = shift;

$fh->print("<mods:subject>\n\t<mods:topic><\/mods:topic>\n<\/mods:subject>\n\n");

};



### MODS Access Condition

sub mods_access_condition
{

my $fh=shift;

$fh->print("<mods:accessCondition type=\"use and reproduction\">Copyright is held by the author, with all rights reserved, unless otherwise noted.<\/mods:accessCondition>\n\n");

}

### MODS Extension Element

sub mods_extension
{
my ($fh, $fileName) = @_;

	$fh->print("<mods:extension>\n\t");
	$fh->print("<etdms:degree>\n\t\t");
	$fh->print("<etdms:name>STL<\/etdms:name>\n\t\t");
	$fh->print("<etdms:level>Licentiate<\/etdms:level>\n\t\t");
	$fh->print("<etdms:discipline>Sacred Theology<\/etdms:discipline>\n\t\t");
	$fh->print("<etdms:grantor>Boston College. School of Theology and Ministry<\/etdms:grantor>\n\t");
	$fh->print("<\/etdms:degree>\n");
	$fh->print("<\/mods:extension>\n\n");
}



### MODS RecordInfo Element

sub mods_record_info
{
my $fh = shift;

$fh->print("<mods:recordInfo>\n");	
	$fh->print("\t<mods:recordContentSource authority=\"marcorg\">MChB<\/mods:recordContentSource>\n");
	$fh->print("\t<mods:languageOfCataloging>\n\t\t<mods:languageTerm type=\"text\" authority=\"iso639-2b\">English<\/mods:languageTerm>\n\t\t<mods:languageTerm type=\"code\" authority=\"iso639-2b\">eng<\/mods:languageTerm>\n\t<\/mods:languageOfCataloging>\n");
$fh->print("<\/mods:recordInfo>\n\n");


}



### MODS Name Element


sub mods_name_element_author
{
#Read a tab-delimited line of metadata and assign each element to an appropriately named variable
#
my $fh=shift;
my $authors = shift;
my $family;
my $given; 


my @authors = split(/\s*;\s*/, $authors);


foreach (@authors) {


my $display_form = $_;
my ($family_name, $given_name) = split(/\s*,\s*/, $display_form);


$fh->print ("<mods:name type=\"personal\" usage=\"primary\">\n\t<mods:namePart type=\"family\">$family_name<\/mods:namePart>\n\t<mods:namePart type=\"given\">$given_name<\/mods:namePart>\n\t<mods:displayForm>$family_name, $given_name<\/mods:displayForm>\n\t<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Author<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>\n\t<\/mods:role>\n<\/mods:name>\n\n");

	} 

}

sub mods_name_element_advisor
{
#Read a tab-delimited line of metadata and assign each element to an appropriately named variable
#
my $fh=shift;
my $advisors = shift;
my $family;
my $given; 


my @advisors = split(/\s*;\s*/, $advisors);


foreach (@advisors) {


my $display_form = $_;
my ($family_name, $given_name) = split(/\s*,\s*/, $display_form);


$fh->print ("<mods:name type=\"personal\">\n\t<mods:namePart type=\"family\">$family_name<\/mods:namePart>\n\t<mods:namePart type=\"given\">$given_name<\/mods:namePart>\n\t<mods:displayForm>$family_name, $given_name<\/mods:displayForm>\n\t<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Thesis advisor<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">ths<\/mods:roleTerm>\n\t<\/mods:role>\n<\/mods:name>\n\n");

	} 

}

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
$fh->print("<mods:mods xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:mods=\"http://www.loc.gov/mods/v3\" xmlns:etdms=\"http://www.ndltd.org/standards/metadata/etdms/1.0/\" xsi:schemaLocation=\"http://www.loc.gov/mods/v3 http://www.loc.gov/standards/mods/v3/mods-3-4.xsd http://www.ndltd.org/standards/metadata/etdms/1.0/ http://www.ndltd.org/standards/metadata/etdms/1.0/etdms.xsd\">\n\n");

return($fh);

};

### Close Output File

sub close_output_file{
my $fh=shift;
$fh->print("<\/mods:mods>\n");
$fh->close();

};
