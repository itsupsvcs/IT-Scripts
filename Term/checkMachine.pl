use strict;

if(@ARGV != 2) {die "usage:  checkMachine.pl [machine] [username]\n";}

#grab the options from the arguments passed
my $machine = shift;
my $user = shift;

#psexec timeout is in seconds
my $psexecTimeout = 8;
print "Checking $machine for $user\n";
my $cmdSucceeded = 0;
my @results = `psexec -n $psexecTimeout \\\\$machine qwinsta 2>nul`;
#verify the command succeeded
foreach my $line (@results) 
{
	if($line =~ /SESSIONNAME/i) {$cmdSucceeded = 1;}
}

if($cmdSucceeded)
{
	#look for a console session
	my $consoleUser = "";
	foreach my $line (@results)
	{
		chomp $line;
		if($line =~ /console/i)
		{
			my @tokens = split(/\s+/, $line);
			#the position of the token we need depends on whether or not there's a > in the output
			#argh
			if($line =~ />console/i) {$consoleUser = $tokens[1];}
			else {$consoleUser = $tokens[2];}
		}
	}
	
	my $needToLogoff = 0;
	if($user =~ /$consoleUser/i) {$needToLogoff = 1;}
	if($needToLogoff)
	{
		print "Logging off console session on $machine\n";
		system("psexec -n $psexecTimeout \\\\" . $machine . ' logoff console >nul 2>&1');
		system("type nul > $consoleUser.$machine.log");
	}
}
else{print "warning: failed to check $machine\n";}
print "Done!\n";