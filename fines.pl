#!/usr/bin/perl

# load modules
use DBI;
use Excel::Writer::XLSX;
use Time::Piece;
use DateTime;
use DateTime::Format::Pg;

# the hostname of your Sierra database server
my $dbhost = "sierradb.myschool.edu";
# the name of your Sierra db, probably "iii"
my $dbname = "iii";
# the port you're accessing the database via, most likely "1032"
my $dbport = "1032";
# your db username (also your Sierra username -- user must have the "Direct SQL Access" application and appropriate perms)
my $dbuser = "username";
# your db pass
my $dbpass = "password";
my %sql = ();

# connect to the database server
my $dbh = DBI->connect("DBI:Pg:dbname=$dbname;host=$dbhost;port=$dbport", "$dbuser", "$dbpass", {'RaiseError' => 1});

# This part of the script may need to be tweaked depending on how you've structured your data.
# I would recommend testing it out in pgAdmin if you need to alter it.
$sql{'patron'} = "SELECT
 pr.id,
 (SELECT vv.field_content FROM sierra_view.varfield_view vv WHERE vv.record_id = pr.id AND vv.record_type_code = 'p' AND vv.varfield_type_code = 'b') AS barcode,
 pfn.last_name,
 pfn.first_name,
 pfn.middle_name,
 (SELECT pra.addr1 FROM sierra_view.patron_record_address pra WHERE pra.patron_record_id = pr.id AND pra.patron_record_address_type_id = 1) AS addr1,
 (SELECT pra.addr2 FROM sierra_view.patron_record_address pra WHERE pra.patron_record_id = pr.id AND pra.patron_record_address_type_id = 1) AS addr2, 
 (SELECT pra.city FROM sierra_view.patron_record_address pra WHERE pra.patron_record_id = pr.id AND pra.patron_record_address_type_id = 1) AS city,
 (SELECT pra.region FROM sierra_view.patron_record_address pra WHERE pra.patron_record_id = pr.id AND pra.patron_record_address_type_id = 1) AS state,
 (SELECT pra.postal_code FROM sierra_view.patron_record_address pra WHERE pra.patron_record_id = pr.id AND pra.patron_record_address_type_id = 1) AS zip,
 (SELECT pra.country FROM sierra_view.patron_record_address pra WHERE pra.patron_record_id = pr.id AND pra.patron_record_address_type_id = 1) AS country,
 (SELECT prp.phone_number FROM sierra_view.patron_record_phone prp WHERE prp.patron_record_id = pr.id AND prp.patron_record_phone_type_id = 1) AS phone,
 (SELECT ppn.description FROM sierra_view.ptype_property_name ppn WHERE ppn.ptype_id = (pr.ptype_code + 1)) AS patron_type,
 pr.owed_amt AS amt_owed,
 (SELECT array(SELECT f.id FROM sierra_view.fine f WHERE f.patron_record_id = pr.id AND f.paid_amt <= 0)) AS fine_ids
FROM sierra_view.patron_record pr
 JOIN sierra_view.patron_record_fullname pfn ON pfn.patron_record_id = pr.id
 JOIN sierra_view.patron_view pv ON pv.id = pr.id
WHERE
 pr.owed_amt > 0
ORDER BY
 pr.owed_amt DESC";

# execute patron query
my $sth = $dbh->prepare("$sql{'patron'}");
$sth->execute();

# set up the Excel file
my $ts = Time::Piece->new->strftime('%Y%m%d-%H%M');
my $filepath = 'c:\\path\\to\\your\\new\\file-'.$ts.'.xlsx'; # make sure to escape the slashes in the file path
my $workbook = Excel::Writer::XLSX->new("$filepath");
my $worksheet = $workbook->add_worksheet();
my ($col, $row) = (0, 0);
# I usually store my formats in a hash, however in this script I'm only using one format for the header row
my %format = ();
$format{'header'} = $workbook->add_format();
$format{'header'}->set_bold();
my @headers = ('Barcode', 'Type', 'Last Name', 'MI', 'First Name', 'Addr1', 'Addr2', 'City', 'State', 'ZIP', 'Phone', 'Patron: Total Owed', 'Fine: Item', 'Fine: Amt', 'Fine: Type', 'Fine: Date Assessed', 'Fine: Invoice #');#, 'DB ID');
$worksheet->write_row($row, $col, \@headers, $format{'header'});
$row++;

# iterate through returned records; on each patron iteration, patron's fines are iterated in the inner foreach/while loops
while(my $res = $sth->fetchrow_hashref()) {
  # $res->{'fine_ids'} returns an array reference, which we store into a scalar
  my $finesRef = $res->{'fine_ids'};
  # dereference the scalar array reference
  my @fines = @$finesRef;
  # and finally iterate through the array of fines, using the ID for each one to query its info
  foreach (@fines) {
    $sql{'fines'} = "SELECT
       f.invoice_num,
       f.item_charge_amt,
       f.processing_fee_amt,
       f.billing_fee_amt,
       f.loanrule_code_num,
       f.title,
       f.assessed_gmt
      FROM sierra_view.fine f
      WHERE f.id = $_";
    my $sth2 = $dbh->prepare("$sql{'fines'}");
    $sth2->execute();
    while (my $res2 = $sth2->fetchrow_hashref()) {
      my $itemFine = $res2->{'item_charge_amt'} + $res2->{'processing_fee_amt'} + $res2->{'billing_fee_amt'};
      my $ts = DateTime::Format::Pg->new();
      my $dateAssessed = $ts->format_date(DateTime::Format::Pg->parse_datetime($res2->{'assessed_gmt'}));
      my @dataRow = ($res->{'barcode'}, $res->{'patron_type'}, $res->{'last_name'}, $res->{'middle_name'}, $res->{'first_name'}, $res->{'addr1'}, $res->{'addr2'}, $res->{'city'}, $res->{'state'}, $res->{'zip'}, $res->{'phone'}, $res->{'amt_owed'}, $res2->{'title'}, $itemFine, $res2->{'loanrule_code_num'}, $dateAssessed, $res2->{'invoice_num'});#, $res->{'id'});
      $worksheet->write_row($row, $col, \@dataRow);
      $row++;  
    }
  }
}

# clean up
$dbh->disconnect();
