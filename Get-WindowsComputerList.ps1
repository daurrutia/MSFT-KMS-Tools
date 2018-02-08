#Get list of all AD Computer Objects
#Get-ADComputer -filter * -Properties dnshostname,operatingsystem | sort Name | select Name,DNSHostName,OperatingSystem

#Export list of all AD Computer Objects inlcuding Ipv4Address property
#Get-ADComputer -filter * -Properties dnshostname,operatingsystem,ipv4address | sort Name | select Name,DNSHostName,Ipv4Address,OperatingSystem | out-file C:\Users\david.urrutia\Desktop\computerIpList.txt

#Get list of all Windows Server computers
#Get-ADComputer -filter {operatingsystem -like "*server*"} -Properties dnshostname,operatingsystem | sort Name | select Name,DNSHostName,OperatingSystem

#Get list of all Windows Server computers 
#Get-ADComputer -filter {operatingsystem -like "*server*"} -Properties dnshostname | sort Name | select -ExpandProperty DNSHostName

#Get list of all enabled Windows "Server" computer object DNS hostnames and write list to current working directory
$path=Get-Location 
Get-ADComputer -filter {(operatingsystem -like "*server*") -and (enabled -eq "true")} -Properties dnshostname | sort Name | select -ExpandProperty DNSHostName | Out-File $path\ComputerList.txt
Write-Host "List written to $path\ComputerList.txt"  