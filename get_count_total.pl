  use DBI;
  use strict;

  my $tablename;
  my $first_line = 1;
  my @column;
  my $cnt = 0;

#Parse count name in CSV
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
  
#Parse server name in CSV  
sub get_server
  {	my @cols = split (/","/,$_);
    shift @cols;
		my @column = split (/\\/,$cols[0]);
		shift @column;
		shift @column;
		$column[0] =~ s/-/_/g;
		return $column[0];
  }  
  
#Input data 
  my $dir = '.\Total';
  my @filename;
  
  chdir ($dir) or die "Cannot CHDIR to '$dir': $!";
  @filename = glob('*.csv');
  
#Connection string for MS SQL
  my $dbh = DBI-> connect('dbi:ODBC:DSN=DSN;UID=user;PWD=password;{ RaiseError => 1, AutoCommit => 1 }') or die "CONNECT ERROR! :: $DBI::err $DBI::errstr $DBI::state $!\n"; 

#Output data  
  open OUT_FILE, ">> !Total_Counters.csv";
  print OUT_FILE "Server;Avg % Processor Time;Max % Processor Time;Avg % User Time;Max % User Time;Avg Processor Queue Length;Max Processor Queue Length;Avg Available Bytes;Min Available Bytes;Avg Page Faults/sec;Max Page Faults/sec;Avg Pages/sec;Max Pages/sec;Avg % Disk Time;Max % Disk Time;Avg Avg. Disk sec/Read;Max Avg. Disk sec/Read;Avg Avg. Disk sec/Write;Max Avg. Disk sec/Write;Avg Avg. Disk Queue Length;Max Avg. Disk Queue Length\n";

#Main  
foreach my $file (@filename){
	print "\n==Open $file\n";
	my @name = split(/-/,$file);

	open CSV, $file;
	  while(<CSV>)	
	  { chomp;
		  if ($first_line){
		  $tablename = get_server($_);
		  #String for CREATE TABLE sql statement
		  my $sql = "CREATE TABLE [dbo].".$tablename."([SNAP] [varchar](50) NULL";
		  @column = proc_head($_);
		  foreach my $col (@column){
			$sql = $sql.",[".$col."] [float] NULL";
		  }	
			$sql = $sql.")";
		  my $sth = $dbh->prepare($sql) or warn "PREPARE ERROR! $DBI::err | $DBI::errstr | $DBI::state | $!\n";
		  if($sth->execute()) {
			print "===Create table [dbo].".$tablename." Successful\n";
		  }
		  $first_line = 0;
		  }
		  else{
			my $val = $_;
			$val =~ s/\"/\'/g; 
			$val =~ s/;/','/g;
			#String for INSERT sql statement
			my $sql = "INSERT INTO [dbo].".$tablename." VALUES (".$val.")";
		    my $sth = $dbh->prepare($sql) or warn "INSERT ERROR! $DBI::err | $DBI::errstr | $DBI::state | $!\n";
		    if($sth->execute()) {
			$cnt++;
		    }			
		 }
	  }
	  $first_line = 1;
	  print "====Inserting $cnt records in [dbo].".$tablename."\n";
	  $cnt = 0;
	  #String for EXEC stored procedure
	  my $sqlmax = "EXEC getavg \@tbl = N'".$tablename."'";
	  my @row_mas = $dbh->selectrow_array($sqlmax);
	  if(@row_mas){
		print "====Get agregate data for counters\n"
	  }
	  #String for DROP TABLE sql statement
	  my $kill = "DROP TABLE ".$tablename."\;";
	  my $sth = $dbh->prepare($kill) or warn "DROP TABLE ERROR! $DBI::err | $DBI::errstr | $DBI::state | $!\n";
	  if($sth->execute()) {
			print "===Drop Table $tablename Successful\n";
		}
	  $tablename =~ s/_/-/g;
	  my $string=$tablename;
		foreach my $r (@row_mas){
			$string=$string.";$r";
		}
	  $string =~ s/\./,/g;
	  #Add string in out file
	  print OUT_FILE "$string\n";
} 

$dbh->disconnect;
close OUT_FILE; 