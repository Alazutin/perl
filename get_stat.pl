#####################################
#  Author:                          #
#   Lazutin Aleksei                 #
#   lazutin.aleksei@gmail.com       #
#   2015                            #
#####################################
  use strict;
  use Excel::Writer::XLSX;
  use DBI;
  use Unicode::Map
  
  my $user = "login in DB";
  
# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new( 'stat_app.xlsx' );
my $border = $workbook->add_format(border => 1);

my $header = $workbook->add_format();
$header->set_border(1);
$header->set_bold();
$header->set_align('center');

# Define Sheets in Excel workbook
my $worksheet_stat = $workbook->add_worksheet('stat_app');
my $worksheet_lost = $workbook->add_worksheet('stat_lost');
my $worksheet_mq_queue = $workbook->add_worksheet('stat_mq_queue');
my $worksheet_err = $workbook->add_worksheet('stat_err');
my $worksheet_current = $workbook->add_worksheet('app_current');

# DEFINE DURATION TIME
print "Enter START Date [DD.MM.YYYY HH24:MI:SS]: ";
my $start_d = <STDIN>;
#my $start_d = '28.11.2014 12:00:00';
print "Enter STOP Date [DD.MM.YYYY HH24:MI:SS]: ";
my $stop_d = <STDIN>;
#my $stop_d = '28.11.2014 15:00:00';

# DEFINE SQL REQUEST
my $req_add = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_add_wcm = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_add_crm = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_add_erib = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_add_fsb = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_proc_all = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_proc_ug = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_proc_otkaz = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_lost = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_und = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_lost_session = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_ckpit_queue = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_erib_queue = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_crm_queue = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_fsb_queue = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_app_err = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";
my $req_app_current = "select * from table_name where timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS') and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')";

$ENV{NLS_LANG}="AMERICAN_AMERICA.CL8MSWIN1251";
# Connect to DB
my $dbh = DBI-> connect('dbi:Oracle:host=server_name;sid=DB_SID;port=1529;','user_name','password') or die "CONNECT ERROR! :: $DBI::err $DBI::errstr $DBI::state $!\n"; 	
$dbh->do("ALTER SESSION SET NLS_DATE_FORMAT = 'DD.MM.YYYY HH24:MI:SS'");

#parameters list: string for sql statement, row number, column number, name worksheet
sub proc_req
	{	my $fields;
		my $sth = $dbh->prepare($_[0]);
		my $row = $_[1];
		my $cols = $_[2];
		my $worksheet = $_[3];
		if($sth->execute()) {
			my $map = Unicode::Map->new("WINDOWS-1251");
			$fields = $sth->{NUM_OF_FIELDS};
			my $name = $sth->{NAME};
			for (my $i =0; $i < $fields; $i++)
				{
					$worksheet->write_string( $row, $cols, $name->[$i], $header);
					$cols++;
				}
			while (my @rows = $sth->fetchrow_array)
			{	$row++;
				$cols = $_[2];
			foreach my $qw (@rows)
				{
					my $wq = $map->to_unicode($qw);
					$worksheet->write_utf16be_string( $row, $cols++, $wq, $border);
				}
			}
		}
	   $sth->finish();
	   return $fields;
	}
	
my $shift = 0;

$shift += proc_req($req_add, 0, $shift,$worksheet_stat);
print "ADD COMPLITE\n";

$shift += proc_req($req_add_wcm, 0, $shift,$worksheet_stat);
print "ADD_WCM COMPLITE\n";

$shift += proc_req($req_add_crm, 0, $shift,$worksheet_stat);
print "ADD_CRM COMPLITE\n";

$shift += proc_req($req_add_erib, 0, $shift,$worksheet_stat);
print "ADD_ERIB COMPLITE\n";

$shift += proc_req($req_add_fsb, 0, $shift,$worksheet_stat);
print "ADD_FSB COMPLITE\n";

$shift += proc_req($req_proc_all, 0, $shift,$worksheet_stat);
print "PROC_ALL COMPLITE\n";

$shift += proc_req($req_proc_ug, 0, $shift,$worksheet_stat);
print "PROC_UG COMPLITE\n";

$shift += proc_req($req_proc_otkaz, 0, $shift,$worksheet_stat);
print "PROC_OTKAZ COMPLITE\n";

$shift += proc_req($req_und, 0, $shift,$worksheet_stat);	
print "UND COMPLITE\n";

proc_req($req_app_err, 0, 0,$worksheet_err);	
print "ERR COMPLITE\n";

#Reset for new worksheet
$shift = 0;
$shift += proc_req($req_lost_session, 0, $shift, $worksheet_lost);	
print "LOST SESSION COMPLITE\n";
$shift++;
$shift += proc_req($req_lost, 0, $shift, $worksheet_lost);	
print "LOST APP COMPLITE\n";

#Reset for new worksheet
$shift = 0;
$shift += proc_req($req_erib_queue, 0, $shift, $worksheet_mq_queue);	
print "ERIB QUEUE COMPLITE\n";
$shift++;
$shift += proc_req($req_crm_queue, 0, $shift, $worksheet_mq_queue);	
print "CRM QUEUE COMPLITE\n";
$shift++;
$shift += proc_req($req_fsb_queue, 0, $shift, $worksheet_mq_queue);	
print "FSB QUEUE COMPLITE\n";

proc_req($req_app_current, 0, 0, $worksheet_current);	
print "CURRENT APP COMPLITE\n";
	
$dbh->disconnect;