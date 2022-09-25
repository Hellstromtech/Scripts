#Find ips

$ips = 1..254 | % {"172.16.0.$_"}
Find-LANHosts $IPs