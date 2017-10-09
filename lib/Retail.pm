    ################################################################################
    #
    # Package for Damping data from Oracle database into MS XLSX file. 
    #
    # Author: Lazutin Aleksey
	# E-mail: Alexey.Lazutin@aplana.com
    # October 2017
	#
	
package Retail;

use strict;
use Exporter;
use Excel::Writer::XLSX;
use Unicode::Map;

use vars qw($VERSION @ISA @EXPORT @EXPORT_OK %EXPORT_TAGS);

$VERSION     = 1.10;
@ISA         = qw(Exporter);
@EXPORT      = ();
@EXPORT_OK   = qw(get_sql proc_req get_sheet_name);
%EXPORT_TAGS = ( DEFAULT => [qw(&get_sql)],
                 Both    => [qw(&get_sql &proc_req &get_sheet_name)]);

sub get_sql
	{
		open (SQL, $_[0]);
		my $st;
		while(<SQL>)
		{
			chop;
			$st=$st.' '.$_;
		}
		close(SQL);
		$st =~ s/\s+$//;
		my @sql = split(/\;/,$st);
		return @sql;
	}

sub get_sheet_name
  {	
	my @column = split (/\\/,$_[0]);
		shift @column;
		shift @column;
		$column[0] =~ s/^(\d+)-(\w+)\.[^\.]*$/$2/;
		return $column[0];
  } 
  
sub proc_req
	{	
		my $row = 0;
		my $cols = 0;
		my $save_point = 0;
		
		my $workbook = shift;
		my $worksheet = shift;
		my $refarr = shift;
		my $fill_type = shift;	
		
		my $border = $workbook->add_format(border => 1);
		
        my $numformat = $workbook->add_format(
            num_format => '0.00',
			border     => 1,
            align      => 'right'
        );		
		# Create a format for the date or time.
        my $dateformat = $workbook->add_format(
            num_format => 'dd/mm/yyyy',
			border     => 1,
            align      => 'right'
        );
		my $timeformat = $workbook->add_format(
            num_format => 'hh:mm',
			border     => 1,			
            align      => 'right'
        );
		
		my $header = $workbook->add_format();
		   $header->set_border(1);
		   $header->set_bold();
		   $header->set_align('center');
		
		my $title = $workbook->add_format(border => 1);
		   $title->set_pattern();
		   $title->set_bg_color('#FFFF00');
		   
		my $map = Unicode::Map->new("WINDOWS-1251");

		foreach my $content (@{$refarr})
		{	
			my $sth = $content->result_set;
			my $comment = $content->comment;

			if (defined $sth){
				my $fields = $sth->{NUM_OF_FIELDS};
				my $name = $sth->{NAME};
				my $merge = 0;
				
				if($fields <= 5) {
					$merge = $fields - 1;
				}
				else {
					$merge = int($fields / 2)-1;
				}	
				
				if (defined $comment){
					print "++$comment\n";
					$worksheet->merge_range($row,$save_point,$row,$save_point+$merge,$comment,$title);
					$row++;
				}
#				Filling Table Header
				for (my $i = 0; $i < $fields; $i++)
				{
					if($name->[$i] =~ /DATE_TIME/){
					 $worksheet->set_column($cols,$cols,10);
					 my @dt = split (/_/,$name->[$i]);
					 $worksheet->write_string( $row, $cols++, $dt[0], $header);
					 $worksheet->set_column($cols,$cols,6);
					 $worksheet->write_string( $row, $cols, $dt[1], $header);
					}
					else {
					 $worksheet->set_column($cols,$cols,length($name->[$i])+2);
					 $worksheet->write_string( $row, $cols, $name->[$i], $header);
					}
					$cols++
				}	
#				Filling Table Data
				while (my @rows = $sth->fetchrow_array)
				{	$row++;
					$cols = $save_point;
				  foreach my $qw (@rows)
					{
						if ($qw =~ /^(\d{1,2})\.(\d{1,2})\.(\d{4}) (\d{1,2}):(\d{1,2}):(\d{1,2})/) {
						# this is column with datetime
							my $date = sprintf "%4d-%02d-%02dT%02d:%02d:%02d", $3, $2, $1, $4, $5, $6;
							$worksheet->write_date_time( $row, $cols++, $date, $dateformat );
							$worksheet->write_date_time( $row, $cols, $date, $timeformat );
						}
						else {
							 if($qw =~ /^\d+$/){
							 # this is column with number	
								$worksheet->write_number( $row, $cols, $qw, $border);
							 }							 
							  else {
							    my $wq = '';
							    if (length($qw)>0){
							     # this is column with cyrillic string
								 $wq = $map->to_unicode($qw);
								} 
							    $worksheet->write_utf16be_string( $row, $cols, $wq, $border);
							   } 
							}
						$cols++;					
					}
				}
				$sth->finish();	
			
			}
			else {
			if (defined $comment){
				$worksheet->write_string( $row, $cols++, $comment, $title);
				}
			}
			if($fill_type =~ 'horizontally'){
				$cols++;
				$save_point = $cols;
				$row = 0;
			}
			if($fill_type =~ 'vertically') {
				$cols = 0;
				$save_point = 0;
				$row += 2;
			}			
		}
	   return $cols;
	}
	