#!/usr/bin/perl -w

use URI;
use Web::Scraper;
use Excel::Writer::XLSX;
use Time::Piece;
use Data::Dumper;

# Your CSDirect login credentials. These need to be provided in order to pass the HTTP
#  authentication on the CSDirect pages.
my $csDirectUsername = 'your_csdirect_username';
my $csDirectPass = 'your_csdirect_password';

# Set up the Excel file.
my $ts = Time::Piece->new->strftime('%Y%m%d-%H%M');
my $filepath = 'c:\\path\\to\\your\\sierraFixes-'.$ts.'.xlsx'; # Make sure to escape the slashes in the file path.
my $workbook = Excel::Writer::XLSX->new("$filepath");
# I usually store my formats in a hash, however in this script I'm only using one format for the header row.
my %format = ();
$format{'header'} = $workbook->add_format();
$format{'header'}->set_bold();
# Set up the header row for the spreadsheet.
my @headers = ('Category', 'Title', 'Description', 'Status');


# Unfortunately the semantic markup on the CSDirect pages is a bit lacking,
#  so until I get time to munge the xPath to the div
#  that contains the URLs for the fix pages, they'll need to be manually placed here.
# The syntax is 
#  $urls{'useful.id'} = 'http://$csDirectUsername:$csDirectPass\@csdirect.iii.com/sierra/kb/index.php?cat_id=9998&tag_id=XYZ',
#  where XYZ is the individual tag id for "useful.id"'s fixes page;
#  also note that "useful.id" is the key that name the worksheets
#  in the created workbook.
my %urls;
$urls{'1.1.0'} = "http://$csDirectUsername:$csDirectPass\@csdirect.iii.com/sierra/kb/index.php?cat_id=9998&tag_id=60";
$urls{'1.1.1'} = "http://$csDirectUsername:$csDirectPass\@csdirect.iii.com/sierra/kb/index.php?cat_id=9998&tag_id=63";
$urls{'1.1.2'} = "http://$csDirectUsername:$csDirectPass\@csdirect.iii.com/sierra/kb/index.php?cat_id=9998&tag_id=66";
$urls{'1.1.3_preview'} = "http://$csDirectUsername:$csDirectPass\@csdirect.iii.com/sierra/kb/index.php?cat_id=9998&tag_id=67";
$urls{'future'} = "http://$csDirectUsername:$csDirectPass\@csdirect.iii.com/sierra/kb/index.php?cat_id=9998&tag_id=52";
$urls{'review'} = "http://$csDirectUsername:$csDirectPass\@csdirect.iii.com/sierra/kb/index.php?cat_id=9998&tag_id=2";

# Scrape the pages in the %urls hash.
foreach my $version ( keys %urls ) {
  my $fixes = scraper {
    process "ul.index", "list[]" => scraper {
      process "li", "fixes[]" => scraper {
        process "a", 'title' => 'TEXT';
        process 'span', "spans[]" => {'data' => 'TEXT'};
      };
    };
  };
  my $headings = scraper {
    process "h2.category", "headings[]" => 'TEXT'; 
  };

  print Dumper $fixes;
  my $uri = URI->new("$urls{$version}");
  my $res = $fixes->scrape ($uri);
  my $res2 = $headings->scrape ($uri);
  my @headings = ();
  my $hdgCount;
  for my $heading (@{$res2->{headings}}) {
    if ($heading =~ m/(\w+\s?\w+) \[(\d+)\]/) {
      for (my $i=0;$i<$2;$i++) {
        push (@headings, $1);
      }
    }
  }

  print Dumper $res;
  # Set up a worksheet for each version/url in the %urls hash,
  #  parse the obtained data and write it to the pertinent sheet.
  my $worksheet = $workbook->add_worksheet($version);
  $worksheet->activate();
  $worksheet->select();
  my ($col, $row) = (0, 0);
  $worksheet->write_row($row, $col, \@headers, $format{'header'});
  $row++;

  for my $list (@{$res->{list}}) {
    for my $fix (@{$list->{fixes}}) {
      my @dataRow = ();
      my $heading = shift(@headings);
      push (@dataRow, $heading);
      push (@dataRow, $fix->{title});
      my $i = 0;    
      for my $span (@{$fix->{spans}}) {
        if ($i == 0) { #Description
          push(@dataRow, substr($span->{data}, 13));
        } elsif ($i == 1) { #Status
          push(@dataRow, substr($span->{data}, 8));
        }
        $i++;
      }
      $worksheet->write_row($row, $col, \@dataRow);
      $row++;  
    }
  }
}
