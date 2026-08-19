#!/usr/bin/perl

use strict;
use warnings;

use PDF::API2;

use constant mm => 25.4 / 72;
use constant in => 1 / 72;
use constant pt => 1;

my $rovName = 'docr';
my $diveNumber = 219;
my $strDiveDate = '3/9/2011';
my $strJDay = '2011-068';
my $strDiveChiefSci = 'DeVogelaere';
my $strMission = 'Container Survey';
my $strCamera = 'Ikegami HDL40';
my $tapeNumberHD = 2;
my $tapeTotalHD = 11;
my $yrNow = 2011;



my $pdf = PDF::API2->new( -file => "${rovName}${diveNumber}_labels.pdf" );
my $now = scalar localtime;

$pdf->info(
        'Author'       => "Mike McCann",
        'CreationDate' => "$now",
        'Creator'      => "$0",
        'Producer'     => "PDF::API2",
        'Title'        => "Tape labels for $rovName $diveNumber",
        'Subject'      => "$rovName $diveNumber",
        'Keywords'     => "video tapes HD ROV dive"
    );

my $page = $pdf->page;
$page->mediabox( 215.9 / mm, 279.4 / mm );			# 8.5 x 11 "
##$page->cropbox( 13 / mm, 43 / mm, 177 / mm, 139 / mm );		# Lower label
##$page->cropbox( 13 / mm, 139 / mm, 177 / mm, 235 / mm );		# Upper label
##$page->artbox  ( 10/mm,  10/mm,   95/mm,  138/mm);

my %font = (
    Helvetica => {
        Bold   => $pdf->corefont( 'Helvetica-Bold',    -encoding => 'latin1' ),
        Roman  => $pdf->corefont( 'Helvetica',         -encoding => 'latin1' ),
        Italic => $pdf->corefont( 'Helvetica-Oblique', -encoding => 'latin1' ),
        BoldItalic => $pdf->corefont( 'Helvetica-BoldOblique', -encoding => 'latin1' ),
    },
);

my $logoFile = 'MBARI_Logo.pnm';
die("Unable to find image file $logoFile: $!") unless -e $logoFile;


my %tapeInfo = ( 	'rovName' => "Doc Ricketts",
			'shipName' => "Western Flyer",
			'diveDate' => "3/9/2011",
			
			);
			
##%tapeInfo = ( 	'rovName' => "Ventana",
##			'shipName' => "Point Lobos",
##			'diveDate' => "3/9/2011",
##			
##			);


##makeLabel('top', $page, \%font, $logoFile, \%tapeInfo);
makeLabel('bot', $page, \%font, $logoFile, \%tapeInfo);

$pdf->save;
$pdf->end();

#--------------------------------------------------------------------
#

=head3 makeLabel()

print a label on the page.  Assumes that the pdf page has been created 

Author: Mike McCann

Date Created: 8/8/2011

=cut


sub makeLabel {
	
	my ( $topOrBotFlag, $page, $rfont, $logoFile, $rtapeHash ) = @_;
	
	print "topOrBotFlag = $topOrBotFlag\n";
	
	my $font = %$rfont;
	

	my $logo = $page->gfx;
	my $logoObject = $pdf->image_pnm($logoFile);
	my $logoSpine = $page->gfx;
	my $titleText = $page->text;
	
	my ($x, $y);
	
	if ( $topOrBotFlag eq 'bot' ) {
		$logo->image( $logoObject, 48 / mm, 112 / mm, 38 / mm, 23.75 / mm );	# MBARI_Logo has 1.6 aspect ratio (376 x 235 pixels)
		$logoSpine->image( $logoObject, 17 / mm, 122 / mm, 23 / mm, 14.375 / mm );	
		if ( $$rtapeHash{'rovName'} eq 'Doc Ricketts' && $$rtapeHash{'shipName'} eq 'Western Flyer' ) {
			$titleText->translate( 102/mm, 128/mm );
		}
		elsif ( $$rtapeHash{'rovName'} eq 'Ventana' && $$rtapeHash{'shipName'} eq 'Point Lobos' ) {
			$titleText->translate( 102/mm, 128/mm );
		}
	}
	else {
	}
	
	$titleText->fillcolor('black');
	$titleText->font( $font{'Helvetica'}{'Bold'}, 16 / pt );
	$titleText->text('ROV ');
	$titleText->font( $font{'Helvetica'}{'BoldItalic'}, 16 / pt );
	$titleText->text($$rtapeHash{'rovName'});
	$titleText->font( $font{'Helvetica'}{'Bold'}, 16 / pt );
	$titleText->text(' Video');
	$titleText->distance( 11/mm, -6/mm );
	$titleText->font( $font{'Helvetica'}{'Bold'}, 14 / pt );
	$titleText->text('R/V ');
	$titleText->font( $font{'Helvetica'}{'BoldItalic'}, 14 / pt );
	$titleText->text($$rtapeHash{'shipName'});
	
	my $frontText = $page->text;
	$frontText->font( $font{'Helvetica'}{'Roman'}, 14 / pt );
	
	my $lineSpacing = 12;
	
	if ( $topOrBotFlag eq 'bot' ) {		
		$x = 48;
		$y = 104;
	} 
	else {
	}
	
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Date: $strDiveDate");
	
	$y -= $lineSpacing;
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Julian Day: $strJDay");
	
	$y -= $lineSpacing;
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Chief Scientist: $strDiveChiefSci");
	
	$y -= $lineSpacing;
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Mission: $strMission");
	
	$y -= $lineSpacing;
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Camera: $strCamera");
	
	if ( $topOrBotFlag eq 'bot' ) {	
		$x = 125;
		$y = 104;
	}
	else {
	}

	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Tape # $tapeNumberHD HD of $tapeTotalHD HD");
	
	$y -= $lineSpacing;
	$frontText->translate( $x/mm, $y/mm );
	$frontText->text("Dive Number: " . uc(substr($rovName,0,1)) . $diveNumber);
	
	
	my $spineText = $page->text;
	$spineText->font( $font{'Helvetica'}{'Roman'}, 12 / pt );
	$lineSpacing = 7;
	
	if ( $topOrBotFlag eq 'bot' ) {	
		$x = 39;
		$y = 121;
	}
	else {
	}
	
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text("$strDiveChiefSci");
	
	$x -= $lineSpacing;
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text("$strMission");
	
	$x -= $lineSpacing;
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text("$strDiveDate");
	
	$x -= $lineSpacing;
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text("$strJDay");
	
	if ( $topOrBotFlag eq 'bot' ) {	
		$x = 39;
		$y = 60;
	}
	else {
	}
	
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text_right("Dive " . uc(substr($rovName,0,1)) . $diveNumber);
	
	$x = 27;
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text_right("Tape # $tapeNumberHD");
	
	$x = 18;
	$spineText->transform(
		-translate => [ $x/mm, $y/mm ],
		-rotate => -90,
		);
	$spineText->text_right("© MBARI $yrNow");

} # End makeLabel()



