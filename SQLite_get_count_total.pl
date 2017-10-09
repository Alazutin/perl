#####################################
#  Author:                          #
#   Lazutin Aleksei                 #
#   lazutin.aleksei@gmail.com       #
#   2017                            #
##################################### 
 use DBI;
 use strict;

  my $tablename;
  my $first_line = 1;
  my @column;
  my $cnt = 0;

# Parse count name in CSV
sub proc_head
  {	my @cols = split (/","/,$_);
	my @str = ();
    shift @cols;
	while (@cols) {
		my @column = split (/\\/,$cols[0]);
		shift @column;
		shift @column;
		shift @column;
		my $st=join '_',@column;
		$st=~ s/\"//g;
		push @str, $st;
		shift @cols;
	}
		return @str;
  }
 
# Parse server name in CSV 
sub get_server
  {	my @cols = split (/","/,$_);
    shift @cols;
		my @column = split (/\\/,$cols[0]);
		shift @column;
		shift @column;
		$column[0] =~ s/-/_/g;
		return $column[0];
  }  
  
# Input data folder
  my $db_dir = '.\DB'; 
  my $dir = '.\Total';
  my @filename;

unless (-e $db_dir)
  {
	print "Folder $db_dir not exsist! Created it now.\n";
	mkdir $db_dir;
  }
  
my $dbh = DBI->connect('dbi:SQLite:uri=file:.\\DB\\counters_retail.db?mode=rwc',undef,undef,{ AutoCommit => 0, RaiseError => 1 })  or die $DBI::errstr;
 
  if ($dbh) {
	print "Opened database successfully\n"; 
  
  chdir ($dir) or die "Cannot CHDIR to '$dir': $!";
# Mask for you .csv file  
  @filename = glob('tv*.csv');
  
# Output data  
  open OUT_FILE, ">> !Total_Counters.csv";
  print OUT_FILE "Server;Avg % Processor Time;Max % Processor Time;Avg % User Time;Max % User Time;Avg Processor Queue Length;Max Processor Queue Length;Avg Available Bytes;Min Available Bytes;Avg Page Faults/sec;Max Page Faults/sec;Avg Pages/sec;Max Pages/sec;Avg % Disk Time;Max % Disk Time;Avg Avg. Disk sec/Read;Max Avg. Disk sec/Read;Avg Avg. Disk sec/Write;Max Avg. Disk sec/Write;Avg Avg. Disk Queue Length;Max Avg. Disk Queue Length\n";
	
foreach my $file (@filename){
	print "\n==Open $file\n";
	my @name = split(/-/,$file);

	open CSV, $file;
	  while(<CSV>)	
	  { chomp;
		  if ($first_line){
		  $tablename = get_server($_);
		  my $sql = "CREATE TABLE ".$tablename."(SNAP TEXT";
		  @column = proc_head($_);
		  foreach my $col (@column){
			$sql = $sql.",[".$col."] REAL";
		  }	
			$sql = $sql.")";
		    my $sth = $dbh->prepare($sql) or warn "PREPARE ERROR! $DBI::err | $DBI::errstr | $DBI::state | $!\n";
		    if($sth->execute()) {
			 print "===Create table ".$tablename." Successful\n";
			 $dbh->commit;
		    }
		  $first_line = 0;
		  }
		  else{
			my $val = $_;
			$val =~ s/\"/\'/g; 
			$val =~ s/;/','/g;
			my $sql = "INSERT INTO ".$tablename." VALUES (".$val.")";
		    my $sth = $dbh->prepare($sql) or warn "INSERT ERROR! $DBI::err | $DBI::errstr | $DBI::state | $!\n";
		  
		    if($sth->execute()) {
			  $cnt++;
		    }		
		  }
	  }
	  $dbh->commit;
	  $first_line = 1;
	  print "====Inserting $cnt records in ".$tablename."\n";
	  $cnt = 0;
	  
	  my $sqlmax = "select 
			avg([Processor(_Total)_% Processor Time]) as [Avg % Processor Time],
			max([Processor(_Total)_% Processor Time]) as [Max % Processor Time],
			avg([Processor(_Total)_% User Time]) as [Avg % User Time],
			max([Processor(_Total)_% User Time]) as [Max % User Time],
			avg([System_Processor Queue Length]) as [Avg Processor Queue Length],
			max([System_Processor Queue Length]) as [Max Processor Queue Length],
			avg([Memory_Available Bytes]) as [Avg Available Bytes],
			min([Memory_Available Bytes]) as [Min Available Bytes],
			avg([Memory_Page Faults/sec]) as [Avg Page Faults/sec],
			max([Memory_Page Faults/sec]) as [Max Page Faults/sec],
			avg([Memory_Pages/sec]) as [Avg Pages/sec],
			max([Memory_Pages/sec]) as [Max Pages/sec],
			avg([LogicalDisk(_Total)_% Disk Time]) as [Avg % Disk Time],
			max([LogicalDisk(_Total)_% Disk Time]) as [Max % Disk Time],
			avg([LogicalDisk(_Total)_Avg. Disk sec/Read]) as [Avg Avg. Disk sec/Read],
			max([LogicalDisk(_Total)_Avg. Disk sec/Read]) as [Max Avg. Disk sec/Read],
			avg([LogicalDisk(_Total)_Avg. Disk sec/Write]) as [Avg Avg. Disk sec/Write],
			max([LogicalDisk(_Total)_Avg. Disk sec/Write]) as [Max Avg. Disk sec/Write],
			avg([LogicalDisk(_Total)_Avg. Disk Queue Length]) as [Avg Avg. Disk Queue Length],
			max([LogicalDisk(_Total)_Avg. Disk Queue Length]) as [Max Avg. Disk Queue Length]
		from ".$tablename.";";
		
	  my @row_mas = $dbh->selectrow_array($sqlmax);
	  if(@row_mas){
		print "====Get agregate data for counters\n"
	  }
	  my $kill = "DROP TABLE ".$tablename."\;";
	  my $sth = $dbh->prepare($kill) or warn "DROP TABLE ERROR! $DBI::err | $DBI::errstr | $DBI::state | $!\n";
	  if($sth->execute()) {
			print "===Drop Table $tablename Successful\n";
			$dbh->commit;
		}	 
	 $tablename =~ s/_/-/g;
	  my $string=$tablename;
		foreach my $r (@row_mas){
			$string=$string.";$r";
		}
	  $string =~ s/\./,/g;	
#	  Add string in out file	  
	  print OUT_FILE "$string\n";
} 

$dbh->disconnect;
close OUT_FILE; 
}