<#
.Synopsis
   Returns the installed applications of all available AD computers with "server" in the operating system attribute.
.DESCRIPTION
   Requires WinRM enabled on target computers.
   Requires AD module
   Error handling: Failed attempts return error message in OS Name property and Test-Connection (ping) result
   in OS Version property/
   *Conditional added to workaround winrm hardened DC (hardcoded dns hostname).

   Returns the installed applications of all available AD computers with "server" in the operating system attribute.
   Uses Invoke-Command (remoting) to return installed applications output as a new object.

   Author: David U.

   ## Future Tinkering:
   ## 1. remove if/else
   ## 2. add enabled property to $servers filtering
   ## 3. add test-connection to $servers filtering
   ## 4. invoke-command against $servers

.EXAMPLE
   Get-InstalledPrograms -Verbose | Export-Csv -NoTypeInformation -Path "~\Desktop\InstalledPrograms.csv"
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-InstalledPrograms
{
    [CmdletBinding()]
    #[Alias()]
    #[OutputType([int])]
    Param
    (
        ## Param1 help description
        #[Parameter(Mandatory=$true,
        #           ValueFromPipelineByPropertyName=$true,
        #           Position=0)]
        #$Param1,
        #
        ## Param2 help description
        #[int]
        #$Param2
    )

    Begin
    {
        $servers = Get-ADComputer -Filter {(operatingsystem -like "*server*") -and (enabled -eq "true")} | sort dnshostname | select -ExpandProperty dnshostname
        $na = 'N/A'
    }
    Process
    {
        foreach($s in $servers){
            Write-Verbose "Collecting info $s"
            try{
                ## If/Else Conditional required for WinRM hardened DC on which script is ran on.
                if ($s -eq "DC1.contoso.com"){
                    $x86apps = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | sort DisplayName
                    $x64apps = Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | sort DisplayName

                    foreach($a in $x86apps){
                        $properties=@{
                            'DisplayName' = $a.DisplayName;
                            'DisplayVersion' = $a.DisplayVersion;
                            'Publisher' = $a.Publisher;
                            'InstallDate' = $a.InstallDate;
                            'Default' = $a.'(Default)';
                            'InstallLocation' = $a.InstallLocation;
                            'HelpLink' = $a.HelpLink
                        }#endproperties
                        $object = New-Object -TypeName psobject -Prop $properties
                        $object | select DisplayName,DisplayVersion,Publisher,InstallDate,Default,InstallLocation,HelpLink
                    }#endforeach

                    foreach($a in $x64apps){
                        $properties=@{
                            'DisplayName' = $a.DisplayName;
                            'DisplayVersion' = $a.DisplayVersion;
                            'Publisher' = $a.Publisher;
                            'InstallDate' = $a.InstallDate;
                            'Default' = $a.'(Default)';
                            'InstallLocation' = $a.InstallLocation;
                            'HelpLink' = $a.HelpLink
                        }#endproperties
                        $object = New-Object -TypeName psobject -Prop $properties
                        $object | select DisplayName,DisplayVersion,Publisher,InstallDate,Default,InstallLocation,HelpLink
                    }#endforeach
                }else{
                    Invoke-Command -ComputerName $s -ScriptBlock{
                        $x86apps = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | sort DisplayName
                        $x64apps = Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | sort DisplayName
                        
                        foreach($a in $x86apps){
                            $properties=@{
                                'DisplayName' = $a.DisplayName;
                                'DisplayVersion' = $a.DisplayVersion;
                                'Publisher' = $a.Publisher;
                                'InstallDate' = $a.InstallDate;
                                'Default' = $a.'(Default)';
                                'InstallLocation' = $a.InstallLocation;
                                'HelpLink' = $a.HelpLink
                            }#endproperties
                            $object = New-Object -TypeName psobject -Prop $properties
                            $object | select DisplayName,DisplayVersion,Publisher,InstallDate,Default,InstallLocation,HelpLink
                        }#endforeach
                        
                        foreach($a in $x64apps){
                            $properties=@{
                                'DisplayName' = $a.DisplayName;
                                'DisplayVersion' = $a.DisplayVersion;
                                'Publisher' = $a.Publisher;
                                'InstallDate' = $a.InstallDate;
                                'Default' = $a.'(Default)';
                                'InstallLocation' = $a.InstallLocation;
                                'HelpLink' = $a.HelpLink
                            }#endproperties
                            $object = New-Object -TypeName psobject -Prop $properties
                            $object | select DisplayName,DisplayVersion,Publisher,InstallDate,Default,InstallLocation,HelpLink
                        }#endforeach
                    } -ErrorAction Stop | select -Property * -ExcludeProperty RunspaceId
                }#endelse
            }catch{
                $properties=@{
                    'PSComputerName'= $s;
                    'DisplayName' = $Error[0].ExceptionMessage;
                    'DisplayVersion' = (Test-Connection $s -Count 2 -Quiet);
                    'Publisher' = $na;
                    'InstallDate' = $na;
                    'Default' = $na;
                    'InstallLocation' = $na;
                    'HelpLink' = $na
                }#endproperties
                $object = New-Object -TypeName psobject -Prop $properties
                $object | select DisplayName,DisplayVersion,Publisher,InstallDate,Default,InstallLocation,HelpLink,PSComputerName 
            }#endtrycatch
        }#endforeach
    }
    End
    {
    }
}