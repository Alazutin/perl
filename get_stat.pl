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
my $worksheet_sms = $workbook->add_worksheet('stat_sms-email');

# DEFINE DURATION TIME
print "Enter START Date [DD.MM.YYYY HH24:MI:SS]: ";
my $start_d = <STDIN>;
#my $start_d = '28.11.2014 12:00:00';
print "Enter STOP Date [DD.MM.YYYY HH24:MI:SS]: ";
my $stop_d = <STDIN>;
#my $stop_d = '28.11.2014 15:00:00';

# DEFINE SQL REQUEST
my $req_add = "select trunc(a.sys_creationdate+3/24, 'HH24') as hour_,count(*) as added from transact.applicants a where a.sys_creationdate between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24  group by trunc(a.sys_creationdate+3/24, 'HH24') order by 1";
my $req_add_wcm = "select trunc(a.sys_creationdate+3/24, 'HH24') as hour_, count(*) as ADD_WCM from transact.applicants a where a.sys_creationdate between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 and a.sys_status = power(2,4) and a.t_31025_src_system_id=2 group by trunc(a.sys_creationdate+3/24, 'HH24') order by 1";
my $req_add_crm = "select trunc(a.sys_creationdate+3/24, 'HH24') as hour_, count(*) as added_crm from transact.applicants a where a.sys_creationdate between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 and a.t_31025_src_system_id=3 group by trunc(a.sys_creationdate+3/24, 'HH24') order by 1";
my $req_add_erib = "select trunc(a.sys_creationdate+3/24, 'HH24') as hour_, count(*) as added_erib from transact.applicants a where a.sys_creationdate between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 and a.t_31025_src_system_id=2 group by trunc(a.sys_creationdate+3/24, 'HH24') order by 1";
my $req_add_fsb = "select trunc(a.sys_creationdate+3/24, 'HH24') as hour_, count(*) as added_fsb from transact.applicants a where a.sys_creationdate between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 and to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 and a.t_31025_src_system_id=4 group by trunc(a.sys_creationdate+3/24, 'HH24') order by 1";
my $req_proc_all = "select trunc(h.sys_timestamp+3/24, 'HH24') as hour_, count(*) as processed_all from transact.historic h join transact.applicants t on h.sys_recordkey=t.sys_recordkey where h.sys_statusaf in(6, 3,15, 36, 45,55) and h.sys_statusbf in(2, 35, 0,33) and h.sys_timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 AND to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 group by  trunc(h.sys_timestamp+3/24, 'HH24') order by 1";
my $req_proc_ug = "select trunc(h.sys_timestamp+3/24, 'HH24') as hour_, count(*) as processed_UG from transact.historic h join transact.applicants t on h.sys_recordkey=t.sys_recordkey where h.sys_statusaf in(6, 3, 15, 36, 45, 55) and h.sys_statusbf in(2, 35) and h.sys_timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 AND to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 group by  trunc(h.sys_timestamp+3/24, 'HH24') order by 1";
my $req_proc_otkaz = "select trunc(h.sys_timestamp+3/24, 'HH24') as hour_, count(*) as processed_otkaz from transact.historic h join transact.applicants t on h.sys_recordkey=t.sys_recordkey where h.sys_statusaf in (3,15,45,55) and h.sys_statusbf in(0,33) and h.sys_timestamp between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 AND to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 group by  trunc(h.sys_timestamp+3/24, 'HH24') order by 1";
my $req_lost = "select trunc(a.sys_creationdate+3/24, 'HH24') as hour_, Count(*) as LOST_APP from transact.applicants a where a.sys_status in (2, power(2, 31), power(2, 33)) and a.sys_creationdate between to_date('".$start_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24 AND (to_date('".$stop_d."', 'DD.MM.YYYY HH24:MI:SS')-3/24) group by trunc(a.sys_creationdate+3/24, 'HH24') order by 1";
my $req_und = "select hour_, sum(async) as async, sum(async_ipo) as async_ipo, sum(under0) as under0, sum(under3) as under3, sum(vydan) as vydan, sum (error_) as error_ from (select h.sys_recordkey, trunc(h.sys_timestamp + 3/24, 'HH24') as hour_, case when h.sys_statusaf in (2) /*and h.sys_statusbf in (2, 35, 0, 33)*/ then 1 else 0 end as async, case when h.sys_statusaf in (35) /*and h.sys_statusbf in ( 35)*/ then 1 else 0 end as async_ipo, case when h.sys_statusaf in (3,45, 4, 32, 7, 38, 8, 40, 9, 42, 10, 44) and h.sys_statusbf in (6,36, 11, 37) then 1 else 0 end as under0 , case when h.sys_statusaf in (3, 45, 4, 32, 8, 40, 9, 42, 10, 44) and h.sys_statusbf in (7, 38, 8, 40, 9, 42, 12, 39, 18, 42, 19, 43) then 1 else 0 end as under3, case when h.sys_statusaf in (17) and h.sys_statusbf in (21, 54, 16) then 1 else 0 end as vydan, case when h.sys_statusaf in (5, 56) then 1 else 0 end as error_ from transact.historic h, transact.applicants a where  h.sys_recordkey = a.sys_recordkey and h.sys_timestamp between to_date('".$start_d."', 'dd.mm.yyyy hh24:mi:ss') - 3/24       and to_date('".$stop_d."', 'dd.mm.yyyy hh24:mi:ss') - 3/24) group by hour_ order by hour_";
my $req_lost_session = "SELECT COUNT(*)AS CNT_ALL, REASONCODE, HOUR_, SUM(UW) AS UW, SUM (CI) AS CI FROM (SELECT SYS_REASONCODE AS REASONCODE,TRUNC(SYS_DATETIME, 'HH') + 3 / 24 as hour_,DECODE (SERVICECENTRE, NULL, 0, 1) AS UW,DECODE (SERVICECENTRE, NULL, 1, 0) AS CI FROM TRANSACT.USERAUDIT UA, SBERBANK_HOOKS.COMPUTEL_ALLUSERS AU WHERE SYS_USER = 'SYSTEM' AND UA.SYS_STRINGPARAM = AU.Login) GROUP BY REASONCODE, HOUR_ ORDER BY HOUR_, REASONCODE";
my $req_ckpit_queue = "select trunc(q.queuedt, 'HH24') HOUR_, q.errmsg, q.status, Count(*) from UG_SBB.Tbl_Ckpit_Queue q group by trunc(q.queuedt, 'HH24'), q.errmsg, q.status order by 1, 3 desc";
my $req_erib_queue = "select hour_, sum(busy_0) as busy_0, sum(busy_1) as busy_1, sum(busy_NULL) as busy_NULL from (select trunc(mq.creation_time, 'hh24') as hour_, case when mq.busy = 0 then 1 else 0 end as busy_0, case when mq.busy = 1 then 1 else 0 end as busy_1, case when mq.busy is NULL then 1 else 0 end as busy_NULL from UG_SBB.MESSAGE_QUEUE mq where mq.ext_system_interface_id in (20)) group by hour_ order by hour_";
my $req_crm_queue = "select hour_, sum(busy_0) as busy_0, sum(busy_1) as busy_1, sum(busy_NULL) as busy_NULL from (select trunc(mq.creation_time, 'hh24') as hour_, case when mq.busy = 0 then 1 else 0 end as busy_0, case when mq.busy = 1 then 1 else 0 end as busy_1, case when mq.busy is NULL then 1 else 0 end as busy_NULL from UG_SBB.MESSAGE_QUEUE mq where mq.ext_system_interface_id in (21)) group by hour_ order by hour_";
my $req_fsb_queue = "select hour_, sum(busy_0) as busy_0, sum(busy_1) as busy_1, sum(busy_NULL) as busy_NULL from (select trunc(mq.creation_time, 'hh24') as hour_, case when mq.busy = 0 then 1 else 0 end as busy_0, case when mq.busy = 1 then 1 else 0 end as busy_1, case when mq.busy is NULL then 1 else 0 end as busy_NULL from UG_SBB.MESSAGE_QUEUE mq where mq.ext_system_interface_id in (26)) group by hour_ order by hour_";
my $req_app_err = "select ts.t_16462_ext_calls_last_intf_co as LAST_INTERFACE, case when ts.t_16462_ext_calls_last_intf_co = 01 then 'Ñòîï-ëèñò' when ts.t_16462_ext_calls_last_intf_co = 02 then 'ÖÎÄ' when ts.t_16462_ext_calls_last_intf_co = 03 then 'ÀÑÑÄ–ÃÑÇ' when ts.t_16462_ext_calls_last_intf_co = 04 then 'ÁÊÈ–ÀÑÑÄ' when ts.t_16462_ext_calls_last_intf_co = 05 then 'ÁÊÈ–Ýêâèôàêñ' when ts.t_16462_ext_calls_last_intf_co = 06 then 'ÁÊÈ–ÍÁÊÈ' when ts.t_16462_ext_calls_last_intf_co = 07 then 'ÁÊÈ-Experian-Interfax' when ts.t_16462_ext_calls_last_intf_co = 08 then 'ÏÔÐ' when ts.t_16462_ext_calls_last_intf_co = 09 then 'ÀÑ Êðåäèòîâàíèÿ' when ts.t_16462_ext_calls_last_intf_co = 10 then 'ÌÁÊÈ' when ts.t_16462_ext_calls_last_intf_co = 11 then 'ÁÐÑ' when ts.t_16462_ext_calls_last_intf_co = 12 then 'Hunter-Ïîèñê íåãàòèâà' when ts.t_16462_ext_calls_last_intf_co = 13 then 'Hunter–Update' when ts.t_16462_ext_calls_last_intf_co = 14 then 'Íàö Hunter-Ïîèñê íåãàòèâà' when ts.t_16462_ext_calls_last_intf_co = 15 then 'Íàö Hunter–Update' when ts.t_16462_ext_calls_last_intf_co = 16 then 'FPS Equifax-Ïîèñê íåãàòèâà' when ts.t_16462_ext_calls_last_intf_co = 17 then 'FPS Equifax–Update' when ts.t_16462_ext_calls_last_intf_co = 18 then 'FPS Equifax-Update Fraud' when ts.t_16462_ext_calls_last_intf_co = 19 then 'SRG' when ts.t_16462_ext_calls_last_intf_co = 20 then 'ÑÁÎË' when ts.t_16462_ext_calls_last_intf_co = 21 then 'SAP HCM' when ts.t_16462_ext_calls_last_intf_co = 22 then 'ÑÏÎÎÁÊ' when ts.t_16462_ext_calls_last_intf_co = 23 then 'ÔÌÑ-Çàïðîñ ñòàòóñà' when ts.t_16462_ext_calls_last_intf_co = 24 then 'ÔÌÑ-Ïîëó÷åíèå ñòàòóñà' when ts.t_16462_ext_calls_last_intf_co = 25 then 'SAP HCM' when ts.t_16462_ext_calls_last_intf_co = 26 then 'ÀÑ ÔÑÁ' when ts.t_16462_ext_calls_last_intf_co = 50 then 'ÁÊÈ-Îöåíêà ÊÈ' else NULL end AS INTF_NAME, count(ta.t_554_app_no) as CNT, ta.t_16453_ext_calls_async_timeou as TIMEOUT_FLAG from transact.applicants ta join transact.sb_sys ts on ta.sys_recordkey = ts.sys_recordkey where ta.sys_status in (power(2, 5), power(2, 56)) and ta.sys_creationdate between to_date('".$start_d."', 'dd.mm.yyyy hh24:mi:ss')-3/24 and to_date('".$stop_d."', 'dd.mm.yyyy hh24:mi:ss')-3/24 group by ts.t_16462_ext_calls_last_intf_co,ta.t_16453_ext_calls_async_timeou";
my $req_app_current = "select a.sys_statusinfo, count(*), s.status_info from transact.applicants a left join sberbank_hooks.app_status s on a.sys_statusinfo = s.status_id where a.sys_creationdate between to_date('".$start_d."', 'dd.mm.yyyy hh24:mi:ss') - 3/24 and to_date('".$stop_d."', 'dd.mm.yyyy hh24:mi:ss') - 3/24 group by a.sys_status, a.sys_statusinfo, s.status_info order by 1 asc";
my $req_sms ="select trunc(sl.received_time, 'HH24') as TIME_, count(*) from sbb_monitor.sms_log sl group by trunc(sl.received_time, 'HH24') order by 1";
my $req_email ="select trunc(sl.received_time, 'HH24') as TIME_, count(*) from sbb_monitor.email_log sl group by trunc(sl.received_time, 'HH24') order by 1";

$ENV{NLS_LANG}="AMERICAN_AMERICA.CL8MSWIN1251";
# Connect to DB
my $dbh = DBI-> connect('dbi:Oracle:host=server_name;sid=DB_SID;port=1529;','user_name','password') or die "CONNECT ERROR! :: $DBI::err $DBI::errstr $DBI::state $!\n"; 	
$dbh->do("ALTER SESSION SET NLS_DATE_FORMAT = 'DD.MM.YYYY HH24:MI:SS'");

#parameters list: string for sql statement, row number, column number, name worksheet
sub proc_req
	{	
		my $sth = $dbh->prepare($_[0]);
		my $row = $_[1];
		my $cols = $_[2];
		my $worksheet = $_[3];
		if($sth->execute()) {
			my $map = Unicode::Map->new("WINDOWS-1251");
			my $fields = $sth->{NUM_OF_FIELDS};
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
	}
	
proc_req($req_add, 0, 0,$worksheet_stat);
print "ADD COMPLITE\n";
proc_req($req_add_crm, 0, 4,$worksheet_stat);
print "ADD_CRM COMPLITE\n";
proc_req($req_add_erib, 0, 6,$worksheet_stat);
print "ADD_ERIB COMPLITE\n";
proc_req($req_add_fsb, 0, 8,$worksheet_stat);
print "ADD_FSB COMPLITE\n";
proc_req($req_add_wcm, 0, 2,$worksheet_stat);
print "ADD_WCM COMPLITE\n";
proc_req($req_proc_all, 0, 10,$worksheet_stat);
print "PROC_ALL COMPLITE\n";
proc_req($req_proc_ug, 0, 12,$worksheet_stat);
print "PROC_UG COMPLITE\n";
proc_req($req_proc_otkaz, 0, 14,$worksheet_stat);
print "PROC_OTKAZ COMPLITE\n";

proc_req($req_und, 0, 16,$worksheet_stat);	
print "UND COMPLITE\n";

proc_req($req_app_err, 0, 0,$worksheet_err);	
print "ERR COMPLITE\n";

proc_req($req_lost, 0, 6, $worksheet_lost);	
print "LOST APP COMPLITE\n";
proc_req($req_lost_session, 0, 0, $worksheet_lost);	
print "LOST SESSION COMPLITE\n";

proc_req($req_ckpit_queue, 0, 0, $worksheet_mq_queue);	
print "CKPIT QUEUE COMPLITE\n";
proc_req($req_erib_queue, 0, 5, $worksheet_mq_queue);	
print "ERIB QUEUE COMPLITE\n";
proc_req($req_crm_queue, 0, 10, $worksheet_mq_queue);	
print "CRM QUEUE COMPLITE\n";
proc_req($req_fsb_queue, 0, 15, $worksheet_mq_queue);	
print "FSB QUEUE COMPLITE\n";

proc_req($req_app_current, 0, 0, $worksheet_current);	
print "CURRENT APP COMPLITE\n";

proc_req($req_sms, 0, 0, $worksheet_sms);	
print "SMS COMPLITE\n";
proc_req($req_email, 0, 4, $worksheet_sms);	
print "EMAIL COMPLITE\n";

$dbh->disconnect;