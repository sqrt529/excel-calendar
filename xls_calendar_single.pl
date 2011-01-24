#!/usr/bin/perl
# xls_calendar_single.pl - create a calendar as M$ Excel file 
#
# Copyright (C) 2010 Joachim "Joe" Stiegler <blablabla@trullowitsch.de>
# 
# This program is free software; you can redistribute it and/or modify it under the terms
# of the GNU General Public License as published by the Free Software Foundation; either
# version 3 of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
# without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
# See the GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License along with this program;
# if not, see <http://www.gnu.org/licenses/>.
#
# --
# 
# Version: 1.0 - 2010-12-14
#
# Known Bugs: 
# 1: Only the column of February has the right width.

use warnings;
use strict;
use Getopt::Std;
use Spreadsheet::WriteExcel;
use Date::Calc qw( Day_of_Week Easter_Sunday Add_Delta_Days leap_year );

our ($opt_h, $opt_o, $opt_y, $opt_l);

sub usage {
	die "usage: $0 -o <outfile.xls> [-y <year>] [-l <de> | <en> ]\n";
}

if ( (!getopts("ho:y:l:")) or (!defined($opt_o)) ) {
	usage();
}

my @month_names_en = qw( January February March April May June July August September October November December );
my @month_names_de = qw( Januar Februar Maerz April Mai Juni Juli August September Oktober November Dezember );
my @month_days = qw( 31 28 31 30 31 30 31 31 30 31 30 31);

my @month_names;

if (defined($opt_l)) {
	if ($opt_l eq "de") {
		push @month_names, @month_names_de;
	}
	elsif ($opt_l eq "en") {
		push @month_names, @month_names_en;
	}
	else {
		usage();
	}
}
else {
	push @month_names, @month_names_de;
}

my $year = undef;

if ( (defined($opt_y)) and ($opt_y =~ /^[\d]*$/) ) {
	$year = $opt_y;
}
else {
	my @now = localtime(time);
	$year = $now[5] + 1900;
}

if (leap_year($year)) {
	$month_days[1] += 1;
}

my @holidays;

# Holidays which depends on Easter Sunday
foreach my $offset (qw( -2 1 39 49 50 60 0 )) {
	my ($y, $m, $d) = (Add_Delta_Days(Easter_Sunday($year), $offset));

	$m = "0$m" if (length($m) < 2);
	$d = "0$d" if (length($d) < 2);

	push @holidays, "$y$m$d";
}

# Add your fixed holidays here (these are the days for germany). The  format is "mmdd"
foreach my $fixholidays (qw( 0101 0106 0501 1003 1101 1225 1226 )) {
	push @holidays, "$year$fixholidays";
}

my $excel = Spreadsheet::WriteExcel->new($opt_o);
die "Can't create $opt_o: $!" unless defined $excel;

my $sheet = $excel->add_worksheet($year);

my $left_bold = $excel->add_format();
$left_bold->set_bold();
$left_bold->set_color('black');
$left_bold->set_align('left');
$left_bold->set_size(8);

my $left_border = $excel->add_format();
$left_border->set_border();
$left_border->set_align('left');
$left_border->set_size(8);

my $center_border = $excel->add_format();
$center_border->set_border();
$center_border->set_align('center');
$center_border->set_size(8);

my $center_border_bg = $excel->add_format();
$center_border_bg->set_border();
$center_border_bg->set_align('center');
$center_border_bg->set_bg_color('silver');
$center_border_bg->set_size(8);

my $center_border_bg_holiday = $excel->add_format();
$center_border_bg_holiday->set_border();
$center_border_bg_holiday->set_align('center');
$center_border_bg_holiday->set_bg_color('gray');
$center_border_bg_holiday->set_size(8);

my $col = my $row = 0;

$sheet->write($row, $col, $year, $left_bold);

$row++;

my $mon = 0;

foreach my $month (@month_names) {
	$row = 2;

	$sheet->write($row, $col, $month, $left_bold);
	$sheet->set_column($row, $col, 2);

	$row++;

	my $tmpmon = $mon + 1;

	for (my $i=1; $i<=$month_days[$mon]; $i++) {
		$tmpmon = "0$tmpmon" if (length($tmpmon) < 2);
		my $tmpi = $i;
		$tmpi = "0$tmpi" if (length($tmpi) < 2);

		if ((grep /^$year$tmpmon$tmpi$/, @holidays) == 1) {
			$sheet->write($row, $col, $i, $center_border_bg_holiday);
		}
		elsif (Day_of_Week($year, $tmpmon, $i) < 6) {
			$sheet->write($row, $col, $i, $center_border);
		}
		else {
			$sheet->write($row, $col, $i, $center_border_bg);
		}

		$sheet->set_column($row, $col, 2);
		$row++;
	}

	$col++;
	$row = 3;

	for (my $j=1; $j<=$month_days[$mon]; $j++) {		
		$tmpmon = "0$tmpmon" if (length($tmpmon) < 2);
		my $tmpj = $j;
		$tmpj = "0$tmpj" if (length($tmpj) < 2);

		if ((grep /^$year$tmpmon$tmpj$/, @holidays) == 1) {
			$sheet->write_blank($row, $col, $center_border_bg_holiday);
		}
		elsif (Day_of_Week($year, $tmpmon, $j) < 6) {
			$sheet->write_blank($row, $col, $center_border);
		}
		else {
			$sheet->write_blank($row, $col, $center_border_bg);
		}

		$sheet->set_column($row, $col, 15);
		$row++
	}

	$col++;
	$mon++;
}

$excel->close() or die "Can't close $opt_o: $!";
