use strict;
my @testNmap = `nmap 2>nul`;
my @testPsExec = `psexec 2>nul`;
unless(@testNmap) {die "error: could not find nmap";}
unless(@testPsExec) {die "error: could not find psexec";}

#clear any results from the previous run
system('del /q *.log >nul 2>&1');

my $network = "";
my $user = "";
my $adSite = "";

#verify we have the arguments we need
#if(@ARGV != 2) {die "usage:  logoffTermedUsers.pl [username] [AD_site]\n";}
$user = shift;
$adSite = shift;



if($adSite =~ /NY/i) {$network = '172.27.8,16,29.1-254';}
elsif($adSite =~ /HOU/i 
	  or $adSite =~ /TX/i) {$network = '172.18.8,29,31.1-254';}
elsif($adSite =~ /UK/i 
	  or $adSite =~ /EU/i) {$network = '172.28.8,16,24,29,31-32.1-254';} 
elsif($adSite =~ /GEN/i 
	  or $adSite =~ /FF/i) {$network = '172.21.7.1-254';}
elsif($adSite =~ /SIN/i
	  or $adSite =~ /SI/i) {$network = '172.16.8,16,24,29,30.1-254';}
elsif($adSite =~ /AU/i) {$network = '172.19.8,16,29,31.1-254';}
elsif($adSite =~ /HK/i) {$network = '172.21.16.1-254';}
elsif($adSite =~ /TO/i) {$network = '172.24.10,30,31.1-254';}
else{$network = '172.17.4,8,9,12-14,17-20,23,24,72.1-254';}


#if($network =~ /^$/) {die "error: no network defined for AD site $adSite\n";}

#scanned the defined network for Windows machines
print "Looking for $user in $network.\n";
system("nmap -p 3389 $network -oG nmap1.txt >nul 2>&1");
open(RESULTS,"nmap1.txt") or die "error: could not find nmap1.txt";
my @results = <RESULTS>;
close RESULTS;

#parse the results
my @machines = ();
foreach my $line (@results)
{
	chomp $line;
	if($line =~ /^Host:/i)
	{
		if($line =~ /\/open\//i)
		
		{
			my @tokens = split(/\s+/, $line);
			push @machines, $tokens[1];
		}
	}
}
#|| $line =~ /\/filtered\//i

my $machineCount = @machines;
print "Found $machineCount machines total in $network.\n";


foreach my $machine (@machines)
{
	system("start checkMachine.pl $machine $user");
}

