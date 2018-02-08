<#
.Synopsis
   Gets the license status for Office 2010 32-bit installed on a 64-bit Windows OS. 
.DESCRIPTION
   Gets the license status for Office 2010 32-bit installed on a 64-bit Windows OS.
   Utilizes ospp.vbs and cscript. 
   Creates a temp txt file of the ospp.vbs output.
   Parses the text file and assigns the string data to an object.
   Outputs each object.
   Removes the temp txt file.

   Thanks:
   http://bit.ly/2E59J50
.EXAMPLE
   Get-OfficeLicenseStatus | Export-Csv -notypeinformation ~\Desktop\OfficeLicenseStatus.csv
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-OfficeLicenseStatus
{
    [CmdletBinding()]
    #[OutputType([int])]
    Param
    (
    #    # Param1 help description
    #    [Parameter(Mandatory=$true,
    #               ValueFromPipelineByPropertyName=$true,
    #               Position=0)]
    #    $Param1,
    #
    #    # Param2 help description
    #    [int]
    #    $Param2
    )

    Begin
    {
        if("${env:ProgramFiles(x86)}\Microsoft Office\Office14"){
            C:\Windows\System32\cscript.exe 'C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS' /dstatus| Out-File $env:temp\ol_stat.txt -Force -Append
        }#endif
        if("${env:ProgramFiles(x86)}\Microsoft Office\Office16"){
            C:\Windows\System32\cscript.exe 'C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS' /dstatus| Out-File $env:temp\ol_stat.txt -Force -Append
        }#endif
        if("$env:ProgramFiles\Microsoft Office\Office16"){
            C:\Windows\System32\cscript.exe "$env:ProgramFiles\Microsoft Office\Office16\OSPP.VBS" /dstatus| Out-File $env:temp\ol_stat.txt -Force -Append
        }#endif
    }
    Process
    {
        if("$env:temp\ol_stat.txt"){
            $dStatus = (Get-Content $env:temp\ol_stat.txt -raw) `
                            -replace ":"," =" `
                            -split "---------------------------------------" `
                            -notmatch "---Processing--------------------------" `
                            -notmatch "---Exiting-----------------------------"

            $dStatus | 
                ForEach-Object {
                    $Props = ConvertFrom-StringData -StringData ($_ -replace '\n-\s+')
                    $obj = New-Object psobject -Property $Props
                    $obj | Add-Member -NotePropertyName ComputerName -NotePropertyValue $env:COMPUTERNAME
                    if(($obj.'SKU ID' -ne $null) -and (($obj.'LICENSE NAME' -ne $null))){
                        Write-Output $obj | 
                            select ComputerName,
                                'SKU ID',
                                'LICENSE NAME',
                                'LICENSE DESCRIPTION',
                                'LICENSE STATUS',
                                'ERROR CODE',
                                'ERROR DESCRIPTION',
                                'Last 5 characters of installed product key'
                    }#endif
                }#endforeach
        }else{
            Write-Error "Failed to generate temp OSPP.VBS output. (ol_stat.txt in $env:temp)"
        }#endelse
    }
    End
    {
        if("$env:temp\ol_stat.txt"){Remove-Item "$env:temp\ol_stat.txt" -Force}    
    }
}