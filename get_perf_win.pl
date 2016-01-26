#####################################
#  Author:                          #
#   Lazutin Aleksei                 #
#   lazutin.aleksei@gmail.com       #
#   2015                            #
#####################################
  use DBI;
  use strict;
  use Excel::Writer::XLSX;

#Input data 
  my $dir = 'Total';
#Flag first line in .csv file  
  my $first_line = 1;
  my $tablename;
  my @cols;
#Number of records in a DB  
  my $cnt = 0;
  my $size = 0;
#Number of line in Sheets in Excel workbook 
  my $row_net = 1;
  my $row_ug = 1;
  my $row_ug2 = 1;
  my $row_total = 1;
  my $row_ses = 1;
  my $row_trs = 1;
  my $row_w3wp = 1;
  my $row_iis = 1;
#Flag for first line in .xlsx file
  my $row_flag_net = 0;
  my $row_flag_iis = 0;
  my $row_flag_ug = 0;
  my $row_flag_ug2 = 0;
  my $row_flag_total = 0;
  my $row_flag_ses = 0;
  my $row_flag_trs = 0;
  my $row_flag_w3wp = 0;

#Connection string for MS SQL  
my $dbh = DBI-> connect('dbi:ODBC:DSN=DSN;UID=user;PWD=password;{ RaiseError => 1, AutoCommit => 1 }') or die "CONNECT ERROR! :: $DBI::err $DBI::errstr $DBI::state $!\n"; 

#Main
opendir(CURR_DIR, ".\\".$dir) or die "Cannot open '$dir': $!";
	foreach my $item (readdir CURR_DIR)
	{
		if( $item !~/^\.{1,2}$/)
		{
			proc_dir($item);
			#Reset of lines for the new Excel workbook
			  $row_net = 1;
			  $row_ug = 1;
			  $row_ug2 = 1;
			  $row_total = 1;
			  $row_ses = 1;
			  $row_trs = 1;
			  $row_w3wp = 1;
			  $row_iis = 1;
			#Reset flag for the new Excel workbook
			  $row_flag_net = 0;
			  $row_flag_iis = 0;
			  $row_flag_ug = 0;
			  $row_flag_ug2 = 0;
			  $row_flag_total = 0;
			  $row_flag_ses = 0;
			  $row_flag_trs = 0;
			  $row_flag_w3wp = 0;
		}
	}
closedir CURR_DIR;
$dbh->disconnect;

#Pricessing for each dir with .csv file
#parameters list: directory with server name
sub proc_dir
{	
  my $item = shift;
  print "+++++++++++++++++++\nOpen folder: $item\n";
  #Out file name
  my $file = $item.".xlsx";
  my @sheet_prefix = split(/-/,$item);
  print " Create $file for $item\n";
	#Create .xlsx out file
	my $workbook = Excel::Writer::XLSX->new( $file );
	my $border = $workbook->add_format(border => 1);
	#Set properties of a cell
	my $header = $workbook->add_format();
	$header->set_border(1);
	$header->set_bold();
	$header->set_align('center');

	#Define Sheets in Excel workbook
	my $worksheet_iis = $workbook->add_worksheet('IIS-'.$sheet_prefix[3]);
	my $worksheet_net = $workbook->add_worksheet('Network-'.$sheet_prefix[3]);
	my $worksheet_w3wp = $workbook->add_worksheet('W3WP-'.$sheet_prefix[3]);
	my $worksheet_total = $workbook->add_worksheet('TotalSystem-'.$sheet_prefix[3]);
	my $worksheet_trs = $workbook->add_worksheet('Transact52-'.$sheet_prefix[3]);
	my $worksheet_ses = $workbook->add_worksheet('Sessions-'.$sheet_prefix[3]);
	my $worksheet_ug = $workbook->add_worksheet('UG-'.$sheet_prefix[3]);
	my $worksheet_ug2 = $workbook->add_worksheet('UG2-'.$sheet_prefix[3]);
	
	$item = ".\\".$dir."\\".$item;
  	if(opendir CSV_DIR, $item)
	{
		foreach my $script_item (readdir CSV_DIR)
		{
			if(	$script_item=~/\.csv$/)
			{
				proc_file("$item/$script_item", $item, $worksheet_iis, $worksheet_net, $worksheet_w3wp, $worksheet_total, $worksheet_trs, $worksheet_ses, $worksheet_ug, $worksheet_ug2);
			}
		}
		closedir CSV_DIR;
	}
	else
	{
		print "Can't open script dir '$item'\n";
	}
}  

#Pricessing for each .csv file
#parameters list: path to .csv file, path to dir with .csv files, list of name worksheet
sub proc_file {
my $source_file = shift;

if(open CSV, $source_file)
{
	print " Open file: $source_file\n";

	if ($dbh)
	{
	  my @dubl=();
	  my $name;
	  #String for SELECT statement
	  my $sqlmax = "select ";
	  my $sqlavg = "select ";
	  
	  my $thead;
	  my $z_flag = 0;

	  while(<CSV>)	
	  {
		  chomp;
		  my @tbl = split(/\//,$source_file);
		  my @tbl2 = split(/-/,$tbl[1]);
		  $name=$tbl2[0];
		  if ($first_line){
		  @cols = split (/","/,$_);
		  $size = scalar @cols;
		  my $table_sql;
		  my $table = "[TIME] [varchar](50) NULL,";
		  for(my $i=1;$i<$size;$i++)
		  {	
			#print "+$cols[$i]\n";
			my @column = split (/\\/,$cols[$i]);
			my @column_2 = split (/\\/,$cols[$i+1]);
			#Stub for long name network interface
			if ($column[3] =~ /Intel\[R\]/)
				{
					$column[3] = 'Intel_PRO_1000_MT';
				}
			
			my $columns = $column[3].'_'.$column[4];
			my $columns_2 = $column_2[3].'_'.$column_2[4];
			
				$tablename = $column[2];
				$tablename =~ s/-/_/g; 
				$tablename = $tablename.'_'.$name;
			
			if($i == 1){
				$thead = $thead.$tablename.";";
			}
			#Skip isatap network interface
			if ($i==1 && $column[3] =~ /isatap/)
			{
				$z_flag = 1;
				next;
			}
			
			if($i == $size-1){
				$table = $table."[".$columns."] [float] NULL)";
				$sqlmax = $sqlmax."max([".$columns."]) from ".$tablename;
				$sqlavg = $sqlavg."avg([".$columns."]) from ".$tablename;
				$thead = $thead.$columns;
			}
			else {
				$table = $table."[".$columns."] [float] NULL,";	
				$sqlmax = $sqlmax."max([".$columns."]),";
				$sqlavg = $sqlavg."avg([".$columns."]),";
				$thead = $thead.$columns.";";
			}
			#Check on duplicates	
			if ($name eq 'Network' || $columns eq $columns_2)
				{
					#print "+dubl\n";
					if ($columns eq $columns_2)
					{
						push(@dubl,$i);
					}
					$i++;
					if ($i == $size-1) {
						$table = substr ($table, 0, length($table)-1);
						$sqlmax = substr ($sqlmax, 0, length($sqlmax)-1);
						$sqlavg = substr ($sqlavg, 0, length($sqlavg)-1);
						$table = $table.")";
						$sqlmax = $sqlmax." from ".$tablename;
						$sqlavg= $sqlavg." from ".$tablename;
						#$thead = $thead."AVG ".$column[4].";MAX ".$column[4];
					}
					
				}			
		  }
		  #String for CREATE TABLE sql statement
		  $table_sql = "CREATE TABLE [dbo].".$tablename."(".$table;
			my $sth = $dbh->prepare($table_sql);
			  $sth->execute();
			  if($sth) {
				print " Successful create table [dbo].".$tablename."\n";
				$first_line=0;
			  }
			#Set a table head in .xlsx for each class metrics
			if ($name eq 'Network'){
				if ($row_flag_net == 0){	
					proc_head($thead, 0, 0,$_[2]);
					$row_flag_net = 1;
				}
			}
			
			if ($name eq 'IIS'){
				if ($row_flag_iis == 0)	{
					proc_head($thead, 0, 0,$_[1]);
					$row_flag_iis = 1;
				}
			}
			
			if ($name eq 'UG'){
				if ($row_flag_ug == 0){
					proc_head($thead, 0, 0,$_[7]);
					$row_flag_ug = 1;
				}
			}
			
			if ($name eq 'UG_UG2'){
				if ($row_flag_ug2 == 0)	{
					proc_head($thead, 0, 0,$_[8]);
					$row_flag_ug2 = 1;
				}
			}
			
			if ($name eq 'W3WP'){
				if ($row_flag_w3wp == 0){
					proc_head($thead, 0, 0,$_[3]);
					$row_flag_w3wp = 1;
				}
			}
			
			if ($name eq 'TotalSystem'){
				if ($row_flag_total == 0){
					proc_head($thead, 0, 0,$_[4]);
					$row_flag_total = 1;
				}	
			}
			
			if ($name eq 'Transact52'){
				if ($row_flag_trs == 0)	{
					proc_head($thead, 0, 0,$_[5]);
					$row_flag_trs = 1;
				}		
			}
			
			if ($name eq 'sessions' || $name eq 'Sessions'){
				if ($row_flag_ses == 0)	{
					proc_head($thead, 0, 0,$_[6]);
					$row_flag_ses = 1;
				}	
			}
		  }
		#Filling of the table in DB
		else {
		my @cols_value = split (/","/,$_);
		#String for INSERT sql statement
		my $value_sql= "INSERT INTO [dbo].".$tablename." VALUES('".$cols_value[0]."',";
		#print "==$dubl==\n";
		for(my $i=1;$i<$size;$i++)
		  {	
			if( grep( /^$i$/, @dubl))
			{
				#print"skip - $i\n";
				next;
			}
			#print "z_flag: $z_flag \n";
			if ($i == 1 && $name eq 'Network' && $z_flag == 1)
			{	
				#print "skip z_flag\n";
				next;
			}
			
			if($i == $size-1){
				$value_sql=$value_sql."'".$cols_value[$i]."')";
			}
			else{
				$value_sql=$value_sql."'".$cols_value[$i]."',";
				}
				
			if ($name eq 'Network')
				{
					#print "dubl\n";
					$i++;
					if ($i == $size-1) {
						$value_sql = substr ($value_sql, 0, length($value_sql)-1);
						$value_sql = $value_sql.")";
					}
					
				}
				
		  }
		  #Final string for INSERT sql statement
		  $value_sql=~ s/"//g;	  
		  my $sth = $dbh->prepare($value_sql);
			$sth->execute();
			if($sth){
			  $cnt++;
			}		
	  }

	  }
	 
	print " Inserting $cnt records in [dbo].".$tablename."\n";  
	$first_line=1;
	$cnt=0;
	
	print " Exec max and avg counters...\n";
	my @row_max = $dbh->selectrow_array($sqlmax);
		my $outstring_max="MAX";
		foreach my $r (@row_max)
			{
				$outstring_max=$outstring_max.";$r";
			}
	my @row_avg = $dbh->selectrow_array($sqlavg);
		my $outstring_avg="AVG";
		foreach my $r (@row_avg)
			{
				$outstring_avg=$outstring_avg.";$r";
			}
		
	#Set for change decimal delimeter		
	#$string =~ s/\./,/g;	
		
	#Filling of the table in .xlsx file
	print " Write max and avg counters to .xlsx...\n";
	if ($name eq 'Network'){
		proc_head($outstring_avg, 0, $row_net, $_[2]);
		$row_net++;
		proc_head($outstring_max, 0, $row_net, $_[2]);
		$row_net++;
		}
	
	if ($name eq 'IIS'){
		proc_head($outstring_avg, 0, $row_iis, $_[1]);
		$row_iis++;
		proc_head($outstring_max, 0, $row_iis, $_[1]);
		$row_iis++;
		}
	
	if ($name eq 'UG'){
		proc_head($outstring_avg, 0, $row_ug, $_[7]);
		$row_ug++;
		proc_head($outstring_max, 0, $row_ug, $_[7]);
		$row_ug++;
		}
	
	if ($name eq 'UG_UG2'){
		proc_head($outstring_avg, 0, $row_ug2, $_[8]);
		$row_ug2++;	
		proc_head($outstring_max, 0, $row_ug2, $_[8]);
		$row_ug2++;
		}
	
	if ($name eq 'W3WP'){
		proc_head($outstring_avg, 0, $row_w3wp, $_[3]);
		$row_w3wp++;
		proc_head($outstring_max, 0, $row_w3wp, $_[3]);
		$row_w3wp++;
		}

	if ($name eq 'TotalSystem'){
		proc_head($outstring_avg, 0, $row_total, $_[4]);
		$row_total++;
		proc_head($outstring_max, 0, $row_total, $_[4]);
		$row_total++;
		}
	
	if ($name eq 'Transact52'){
		proc_head($outstring_avg, 0, $row_trs, $_[5]);
		$row_trs++;	
		proc_head($outstring_max, 0, $row_trs, $_[5]);
		$row_trs++;
		}	
	
	if ($name eq 'sessions' || $name eq 'Sessions'){
		proc_head($outstring_avg, 0, $row_ses, $_[6]);
		$row_ses++;	
		proc_head($outstring_max, 0, $row_ses, $_[6]);
		$row_ses++;
		}
		
#Delete table in DB
	my $kill= "DROP TABLE ".$tablename;	
	my $sth = $dbh->prepare($kill);
		$sth->execute();
	if($sth){
		print " Drop Table $tablename Successful\n";
	}
  } 
} 

}

#Write to .xlsx file
#parameters list: string for write, row number, column number, name worksheet
sub proc_head {
		my @str = split (/;/,$_[0]);
		my $row = $_[1];
		my $cols = $_[2];
		my $worksheet = $_[3];			
		for (my $i =0; $i < scalar @str; $i++)
				{
					$worksheet->write( $row, $cols, $str[$i]);
					$row++;
				}
		
}