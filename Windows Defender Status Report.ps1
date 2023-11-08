function Invoke-wcParallel {
    [CmdletBinding()]
    param (
        # Mandatory array of servers to run commands against
        [Parameter(Mandatory=$true,
        HelpMessage="This is the array of machines you want to run the scriptblock against.")]
        [Array]$servers,

        # Mandatory scriptblock param with the work to perform per server
        [Parameter(Mandatory=$true,
        HelpMessage="This is the scriptblock to run against the servers.")]
        [scriptblock]$scriptblock,

        # Mandatory credential parameter
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$credential,

        # Mandatory throttlelimit parameter with default value
        [int]$throttlelimit = 100
    )     
    Begin{
        $results = New-Object -TypeName System.Collections.ArrayList
    }
    Process{
        $servers | ForEach-Object -Parallel $scriptblock -ThrottleLimit $throttlelimit | ForEach-Object {
            $results.add($_) | Out-Null
        }
    }
    End{
        $results
    }
}

function Get-RegionalOSservers {
    [CmdletBinding()]
    param (
        # Mandatory region parameter with allowed values
        [Parameter(Mandatory=$true)]
        [ValidateSet("AMERICAS", "EMEA", "ASIAPAC")]
        [string]$region,

        # Mandatory OS parameter with allowed values
        [Parameter(Mandatory=$true)]
        [ValidateSet('2016', '2019', '2022')]
        [string]$os
    )
    # Define the searchbase string for the AD query
    $searchbase = "OU=$os-Servers,OU=Servers,OU=$region,OU=DATACENTERS,DC=<subdomain>,DC=<domain>,DC=com"
    Try{
        # Fetch server names from AD based on searchbase
        $regional_os_servers = Get-ADComputer -filter * -SearchBase $searchbase -ErrorAction Stop | Select-Object -ExpandProperty name
        Write-Host "`t`tThere are $($regional_os_servers.count) $os servers in $region"
        return $regional_os_servers
    }
    Catch{
        Write-Host "Unable to retrieve list of servers from : $searchbase"
        Write-Error $_
    }
}

Function Export-DefenderDatatoExcel{
    [CmdletBinding()]
    param (
        # Mandatory results parameter
        [Parameter(Mandatory=$true)]
        $results,

        # Mandatory region parameter with allowed values
        [Parameter(Mandatory=$true)]
        [ValidateSet("AMERICAS", "EMEA", "ASIAPAC")]
        [string]$region,

        # Non-Mandatory path parameter with default value
        [string]$path = 'C:\temp\defender_status_report.xlsx',

        # Mandatory OS parameter with allowed values
        [Parameter(Mandatory=$true)]
        [ValidateSet('2016', '2019', '2022')]
        [string]$os
    )       

    $worksheet_name = $region + '_' + $os
    Export-Excel -Path $path -WorksheetName $worksheet_name -InputObject $results -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
}

$regions = @('AMERICAS','EMEA','ASIAPAC')
$operating_systems = @('2016','2019','2022')
$credential = (Get-Credential)

# Loop through regions and OS types
Foreach($region in $regions){
    Write-Host "Processing Region: $Region"
    Foreach($os in $operating_systems){
        Write-Host "`tProcessing Operating System: $os"
        # Fetch server list
        $servers = Get-RegionalOSservers -region $region -os $os
        $scriptblock = {
            Write-Host "Processing Server: $_"
            function Get-ServerDefenderInstallStatus {
                [CmdletBinding()]
                param (
                    # Mandatory server parameter
                    [Parameter(Mandatory=$true)]
                    [string]$server,
            
                    # Mandatory credential parameter
                    [Parameter(Mandatory=$true)]
                    [System.Management.Automation.PSCredential]$credential,
            
                    # Mandatory region parameter with allowed values
                    [Parameter(Mandatory=$true)]
                    [ValidateSet("AMERICAS", "EMEA", "ASIAPAC")]
                    [string]$region,
            
                    # Mandatory OS parameter with allowed values
                    [Parameter(Mandatory=$true)]
                    [ValidateSet('2016', '2019', '2022')]
                    [string]$os

                )   
                #Write-host "Processing Server : $server"
                Try{
                    $result = Invoke-Command -ComputerName $server -Credential $credential -ScriptBlock {
                        try{
                            # Define the base object to be used as a result.
                            $return_obj = New-Object -TypeName psobject -Property @{
                                antivirus_enabled = ''
                                error_message = ''
                                region = $USING:region
                                os = $USING:os
                            }
                            <# using get-mpcomputerstatus because it will exist on any machine with defender installed.
                               this allows for US to catch for the error that the command is missing and know that defender
                               is not installed on the server.
                             #>
                            $defender_query = Get-MpComputerStatus -ErrorAction stop | Select-Object AntivirusEnabled
                            $return_obj.antivirus_enabled = $defender_query.AntivirusEnabled
                            $return_obj.error_message = 'N/A'
                        }
                        Catch{
                            <# 
                                Both of the below errors exist only if defender is not installed.
                                I believe the first error is what 2019 servers throw and the second
                                is what 2016 servers throw. leaving the last block as a catch-all for errors i didn't witness in testing.
                            #>
                            If($_ -like '*Invalid class*'){
                                $return_obj.antivirus_enabled = 'Not Installed'
                                $return_obj.error_message = $_
                            }
                            Elseif($_ -like '*Get-MpComputerStatus*'){
                                $return_obj.antivirus_enabled = 'Not Installed'
                                $return_obj.error_message = $_
                            }
                            Else{
                                $return_obj.antivirus_enabled = 'Unknown'
                                $return_obj.error_message = $_
                            }
                        }
                        Write-Output $return_obj
                    } -ErrorAction Stop
                    return $result
                }
                Catch{
                    #If the remoting connection fails we want to instantiate the same type of object so the resulting data is in the same format.
                    $result = New-Object -TypeName psobject -Property @{
                        antivirus_enabled = 'Unknown'
                        error_message = 'Unable to Connect Remotely With Powershell'
                        PSComputerName = $server
                        region = $region
                        os = $os
                    }
                    return $result
            
                }
            }

            $result = Get-ServerDefenderInstallStatus -Server $_ -Credential $using:credential -region $using:region -os $using:os
            [PSCustomObject]@{
                PSComputerName    = $result.PSComputerName
                antivirus_enabled = $result.antivirus_enabled
                error_message     = $result.error_message
                region            = $result.region
                os                = $result.os
            }
        }
        $results = Invoke-wcParallel -servers $servers -scriptblock $scriptblock -credential $credential -throttlelimit 100

        Export-DefenderDatatoExcel -Results $results -region $region -os $os
    }
}