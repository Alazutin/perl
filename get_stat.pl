#!/usr/bin/perl
    
    ################################################################################
    #
    # Damping data from Oracle database into MS XLSX file. 
    #
    # Author: Lazutin Aleksey
	# E-mail: Alexey.Lazutin@aplana.com
    # October 2017
	#
	# To use the command line parameters, you mast define first and second parameter
	# as the START and STOP Date in format DD.MM.YYYYTHH24:MI:SS
	#
	
use strict;
use Retail qw(:Both);
use DBI;
use Excel::Writer::XLSX;
use Config::Simple;
use Class::Struct;
  
  struct (Sqls =>{
		comment => '$',
		result_set => '$'
  });
  
# READ CONFIG FILE
  my %Config;
  Config::Simple->import_from('connection.ini', \%Config);
	my $host=$Config{"DB.host"};
	my $sid =$Config{"DB.sid"};
	my $port=$Config{"DB.port"};
	my $user=$Config{"AUTH.user"};
	my $pass=$Config{"AUTH.password"};

# DEFINE SQL REQUEST FOlDER
my $dir = "SQL\\test_statistics";	
my @filename = glob("$dir\\*.sql");

# DEFINE DURATION TIME
my $start_d;
my $stop_d;
my $cor_dt = 1;
if (scalar @ARGV > 0){ 
	foreach my $p (@ARGV) {
		if ($p !~ /^(\d{1,2})\.(\d{1,2})\.(\d{4})T(\d{1,2}):(\d{1,2}):(\d{1,2})/) {
			$cor_dt = 0;
		}
	}
	if ($cor_dt){
		$start_d = $ARGV[0];
		$start_d =~ s/T/ /;
		$stop_d  = $ARGV[1];
		$stop_d =~ s/T/ /;
		print " START and STOP Date given from command line parameters\n";
		print " Damping duration: $start_d - $stop_d\n";
	}
	else {
		print "Incorrect date format in command line parameters!!!\nEnter START and STOP Date from console\n\n";
		print "Enter START Date [DD.MM.YYYY HH24:MI:SS]: ";
		$start_d = <STDIN>;
		print "Enter STOP Date [DD.MM.YYYY HH24:MI:SS]: ";
		$stop_d = <STDIN>;
	}
}
else {
	print "Enter START Date [DD.MM.YYYY HH24:MI:SS]: ";
	$start_d = <STDIN>;
	print "Enter STOP Date [DD.MM.YYYY HH24:MI:SS]: ";
	$stop_d = <STDIN>;
}

$ENV{NLS_LANG}="AMERICAN_AMERICA.CL8MSWIN1251";

# CONNNECT TO DB
my $dbh = DBI-> connect('dbi:Oracle:host='.$host.';sid='.$sid.';port='.$port.';',''.$user.'',''.$pass.'') or die "CONNECT ERROR! :: $DBI::err $DBI::errstr $DBI::state $!\n"; 	
$dbh->do("ALTER SESSION SET NLS_DATE_FORMAT = 'DD.MM.YYYY HH24:MI:SS'");

# CREATE NEW XLSX WORKBOOK
$dir =~ s/^(\w+)\\(\w+)$/$2/;
my $workbook = Excel::Writer::XLSX->new( $dir.'.xlsx' );

foreach my $file(@filename)
{
# Define Sheets in Excel workbook
 my $user_stat = $workbook->add_worksheet(get_sheet_name($file));
 my @dataset = ();
 foreach  my $st (get_sql($file))
	{
		$st =~ s/\&start_date/$start_d/;
		$st =~ s/\&end_date/$stop_d/;
		
		my $holder = Sqls->new();
		
		if ($st =~ /--/){
			print "SQL contains single-line comment!!!\n Please, DELETE it from $file\n";
			my $str = "ERROR! SQL contains single-line comment!!! Please, DELETE it from ".$file;
			$holder->comment($str);
		}
		else{
			my	$sth = $dbh->prepare($st); 

			if ($sth) {
				$sth->execute();
				$holder->result_set($sth);
				if ($st =~ /\/\*(.*)\*\//) {
					$holder->comment($1);
					}
				}
		}	
		push (@dataset, $holder);
	}

	if (scalar @dataset > 0) {
		print "Start filing\n";
		my $shift = proc_req($workbook,$user_stat,\@dataset,'horizontally');
		if (length($shift) > 0){
			print "DAMPING DATA COMPLITE SUCCESSFULLY\n\n";
		}
	}

}
$dbh->disconnect;
