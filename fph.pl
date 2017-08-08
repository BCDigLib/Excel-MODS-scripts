#!C:/Perl/bin/perl -w
use strict;
use warnings;
use XML::Simple;
use XML::Writer;
use IO::File;
use utf8;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use Cwd;

my $data;

main();

sub main
	
	{
	#call subroutine to set up Excel Object
	my($worksheet_name, $Sheet, $excel_object) = setup_EXCEL_object(shift);

	##read and process each row in the EXCEL file
	my $usedRange = $Sheet->UsedRange()->{Value};

	#skip header row of excel file			
	shift(@$usedRange);

	#read in faculty names file
	$data=read_faculty_names_xml();

	#make a MODS record for each row in Excel file
	while (my $row=shift @$usedRange) 
		{
  		createMODS($row);
		}
	}


sub read_faculty_names_xml
	{
	# create object
	my $xml = new XML::Simple;

	# read XML file
	$data = $xml->XMLin("facultyNames.xml");
	};


#Get data from faculty names file for a mods name element
sub mods_name_element
	{
      my $shortname = shift;
      my $family;
      my $given;
      my $calc;
      my $affiliation;
	my $orcid;

	foreach my $e (@{$data->{'facultyNames'}})  
	{
		if ($e->{'shortname'} && $e->{'shortname'} eq $shortname) 
		{
			$family = $e->{'family'};
			$given = $e->{'given'};
			$calc = $e->{'calc'};
			$affiliation = $e->{'DEPT'};
			$orcid = $e->{'ORCID'};
		}	
	}
	return ($family, $given, $calc, $affiliation, $orcid);
	};



#the main event
sub createMODS
	{
	my $row=shift;
	my ($fph_web, $title, $statement, $abstract, $interviewee, $interviewer, $extent, $date, $streaming_url)= @$row; 
	my $multi = $interviewee;
	$multi =~s/;/-/g;
	$multi =~s/\s+//g;

	my $output = IO::File->new('>ir-fph-'.$multi.'-'.$date.'.xml');

	my $writer = XML::Writer->new(OUTPUT => $output, DATA_MODE => 1, DATA_INDENT => 2, );
	$writer->xmlDecl('UTF-8');

	##Opening tag
	print "$title\n"; #for debug
	$writer->startTag('mods:mods', 'xmlns:xsi'=>'http://www.w3.org/2001/XMLSchema-instance', 'xmlns:mods'=>'http://www.loc.gov/mods/v3','xmlns:xlink'=>'http://www.w3.org/1999/xlink','xsi:schemaLocation'=>'http://www.loc.gov/mods/v3 http://www.loc.gov/standards/mods/v3/mods-3-6.xsd');


	#Get interviewee info
	my $interview_with='';
	my @interviewees = split/;/,$interviewee;  
	my $i=0;

	foreach (@interviewees)
		{
		$_=~s/^\s+|\s+$//g;
		my ($family, $given, $calc, $affiliation, $orcid)= mods_name_element($_);
		
		if ($i eq 0) {$interview_with =$given.' '.$family}
			else {$interview_with =$interview_with.' and '.$given.' '.$family}
		$i++;
		}

	##MODS Title

	$writer->startTag('mods:titleInfo', usage=>'primary');
	$writer->startTag('mods:title');
	$title=~s/^\"|\"$//g;
	$statement=~s/^\"|\"$//g;
	$writer->characters('Interview with '.$interview_with.' on '.$title.', '.$statement);
	$writer->endTag();
	$writer->endTag();


#Mods Name for interview - account for multiples

    $i=0;
	foreach (@interviewees)
		{
		my ($family, $given, $calc, $affiliation, $orcid)= mods_name_element($_);
		if ($i eq 0)
			{$writer->startTag('mods:name', type=>'personal', usage=>'primary')}
			else {$writer->startTag('mods:name', type=>'personal')}
      	$writer->startTag('mods:namePart', type=>'family');
      	$writer->characters($family);
      	$writer->endTag();
      	$writer->startTag('mods:namePart', type=>'given');
      	$writer->characters($given);
      	$writer->endTag();
		if ($orcid)
		{
			$writer->startTag('mods:identifier', type=>'orcid');
      		$writer->characters($orcid);
      		$writer->endTag();
			print "$interviewee has orcid $orcid\n";
		}
    		$writer->startTag('mods:displayForm');
      	$writer->characters($calc);
      	$writer->endTag();
      	$writer->startTag('mods:affiliation');
      	$writer->characters($affiliation);
      	$writer->endTag();
      	$writer->startTag('mods:role');
     		$writer->startTag('mods:roleTerm', type=>'text', authority=>'marcrelator');
      	$writer->characters("Interviewee");
      	$writer->endTag();
      	$writer->startTag('mods:roleTerm', type=>'code', authority=>'marcrelator');
   	      $writer->characters("ive");
	      $writer->endTag();
	      $writer->endTag();
	      $writer->startTag('mods:description');
	      $writer->characters($_);
	      $writer->endTag();
	      $writer->endTag();
		$i++;
		}



	#Mods Name for Interviewer

	my ($family, $given, $calc, $affiliation, $orcid)= mods_name_element($interviewer);
	$writer->startTag('mods:name',type=>'personal');
	$writer->startTag('mods:namePart', type=>'family');
	$writer->characters($family);
	$writer->endTag();
	$writer->startTag('mods:namePart', type=>'given');
	$writer->characters($given);
	$writer->endTag();
	if ($orcid)
		{
			$writer->startTag('mods:identifier', type=>'orcid');
      		$writer->characters($orcid);
      		$writer->endTag();
			print "$interviewer has orcid $orcid\n";
		}
	$writer->startTag('mods:displayForm');
	$writer->characters($calc);
	$writer->endTag();
	$writer->startTag('mods:affiliation');
	$writer->characters('University Libraries, Boston College');
	$writer->endTag();
	$writer->startTag('mods:role');
	$writer->startTag('mods:roleTerm', type=>'text', authority=>'marcrelator');
	$writer->characters("Interviewer");
	$writer->endTag();
	$writer->startTag('mods:roleTerm', type=>'code', authority=>'marcrelator');
	$writer->characters("ivr");
	$writer->endTag();
	$writer->endTag();
	$writer->startTag('mods:description');
	$writer->characters('nonfaculty');
	$writer->endTag();
	$writer->endTag();

	# MODS name corporate

	$writer->startTag('mods:name',type=>'corporate', authority=>'naf');
	$writer->startTag('mods:namePart');
	$writer->characters('Boston College Libraries');
	$writer->endTag();
	$writer->startTag('mods:displayForm');
	$writer->characters('Boston College Libraries');
	$writer->endTag();
	$writer->startTag('mods:role');
	$writer->startTag('mods:roleTerm', type=>'text', authority=>'marcrelator');
	$writer->characters("Publisher");
	$writer->endTag();
	$writer->startTag('mods:roleTerm', type=>'code', authority=>'marcrelator');
	$writer->characters("pbl");
	$writer->endTag();
	$writer->endTag();
	$writer->endTag();

	# MODS:typeOfResource

	$writer->startTag('mods:typeOfResource');
	$writer->characters('moving image');
	$writer->endTag();

	#MODS:genre

	$writer->startTag('mods:genre', authority=>'marcgt', type=>'work type', usage=>'primary');
	$writer->characters('interview');
	$writer->endTag();

	#MODS:OriginInfo

	$writer->startTag('mods:originInfo');
	$writer->startTag('mods:place');
	$writer->startTag('mods:placeTerm', type=>'code', authority=>'marccountry');
	$writer->characters('mau');
	$writer->endTag();
	$writer->startTag('mods:placeTerm', type=>'text');
	$writer->characters('Chestnut Hill, Mass.');
	$writer->endTag();
	$writer->endTag();
	$writer->startTag('mods:dateCreated');
	$writer->characters($date);
	$writer->endTag();
	$writer->startTag('mods:dateCreated', encoding=>'w3cdtf', keyDate=>'yes');
	$writer->characters($date);
	$writer->endTag();
	$writer->startTag('mods:issuance');
	$writer->characters('monographic');
	$writer->endTag();
	$writer->endTag();

	#MODS Language

	$writer->startTag('mods:language');
	$writer->startTag('mods:languageTerm', authority=>'iso639-2b', type=>'code');
	$writer->characters('eng');
	$writer->endTag();
	$writer->startTag('mods:languageTerm', authority=>'iso639-2b', type=>'text');
	$writer->characters('English');
	$writer->endTag();
	$writer->endTag();
	
	#mods:physicalDescription

	$writer->startTag('mods:physicalDescription');
	$writer->startTag('mods:form', authority=>'marcform');
	$writer->characters('electronic');
	$writer->endTag();
	$writer->startTag('mods:internetMediaType');
	$writer->characters('video/mp4');
	$writer->endTag();
	$writer->startTag('mods:extent');
	$extent=~s/^\"|\"$//g;
	$writer->characters($extent);
	$writer->endTag();
	$writer->startTag('mods:digitalOrigin');
	$writer->characters('born digital');
	$writer->endTag();
	$writer->endTag();

	#mods:abstract

	$writer->startTag('mods:abstract');
	$abstract=~s/^\"|\"$|//g;
	$abstract=~s/\x97|\x96/-/g;
	$writer->characters($abstract);
	$writer->endTag();

	#mods:note

	$writer->startTag('mods:note');
	$writer->characters('Title supplied by cataloger.');
	$writer->endTag();

	#mods:relatedItem for the FPH Series 

	$writer->startTag('mods:relatedItem', type=>'series');
	$writer->startTag('mods:titleInfo', usage=>'primary');
	$writer->startTag('mods:title');
	$writer->characters('Faculty publication highlights');
	$writer->endTag();
	$writer->endTag();
	$writer->endTag();

	#mods:location

	$writer->startTag('mods:location');
	$writer->startTag('mods:url');
	$streaming_url=~s/^\"|\"$//g;
	$writer->characters($streaming_url);
	$writer->endTag();
	$writer->endTag();

	#mods:accessCondition

	$writer->startTag('mods:accessCondition', type=>'use and reproduction');
	$writer->characters('This work is licensed under the Creative Commons Attribution-NonCommercial 4.0 International License (http://creativecommons.org/licenses/by-nc/4.0/).');
	$writer->endTag();

	#mods:recordInfo

	$writer->startTag('mods:recordInfo');
	$writer->startTag('mods:recordContentSource', authority=>'marcorg');
	$writer->characters('MChB');
	$writer->endTag();
	$writer->startTag('mods:languageOfCataloging');
	$writer->startTag('mods:languageTerm', authority=>'iso639-2b', type=>'code');
	$writer->characters('eng');
	$writer->endTag();
	$writer->startTag('mods:languageTerm', authority=>'iso639-2b', type=>'text');
	$writer->characters('English');
	$writer->endTag();
	$writer->endTag();
	$writer->endTag();
	$writer->endTag();
	$writer->end();

	$output->close();
	};

	###  Open and Setup Excel
	sub setup_EXCEL_object 
	{
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

=POD

=head1 Usage
fph.pl

Requires in the same directory:
	Recent export of the faculty names table from the ACCESS database named as facultyNames.xml
	EXCEL spreadsheet described below (should be closed)

=head1 Description
This script is used by the Boston College Libraries eScholarship program to batch generate MODS records for the Faculty Publication Highlights
Interviews series that is preserved in Islandora.

It requests as input a Microsoft Excel spreadhsheet with the following columns:  
	Link to web page for the highlight
	Highlighted Work Title
	Highlighted Work Statement of Responsibility
	Abstract
	Interviewee(s)' username -- can handle multiple
	Interviewer's username -- currently handling one
	Extent of video in minutes and seconds
	Publication year of Highlight
	Streaming URL
Special guidelines for recording each data element in the spreadsheet and should be consulted (to be posted on wiki-- add link here)
The script assumes that the spreadsheet will have a header row containing column labels.

ORCID ids for the interviewee and interviewer will be included if present in the ACCESS database.




Betsy Post
betsy.post@bc.edu
Last revision 20170808

=cut