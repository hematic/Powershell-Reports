#region functions
Function Export-HtmlReport{

    <#
    .SYNOPSIS
        Creates a HTML report
    .DESCRIPTION
        Creates an eye-friendly HTML report with an inline cascading style sheet from given PowerShell objects and saves it to a file.
    .PARAMETER InputObject
        Hashtable containing data to be converted into a HTML report.
        
        HashTable Proberties:
        
        [array]  .Object       Any PowerShell object to be converted into a HTML table.
    
        [string] .Property     Select a set of properties from the object. Default is "*".
        
        [sting]  .As           Use "List" to create a vertical table instead of horizontal alignment.
                
        [string] .Title        Add a table headline.
        
        [string] .Description  Add a table description.
    
    .PARAMETER Title
        Title of the HTML report.
    .PARAMETER OutputFileName
        Full path of the output file to create, e.g. "C:\temp\output.html".
    .EXAMPLE
        @{Object = Get-ChildItem "Env:"} | Export-HtmlReport -OutputFile "HtmlReport.html" | Invoke-Item
    
        Creates a directory item HTML report and opens it with the default browser.
    .EXAMPLE
    
        $ReportTitle = "HTML-Report"
        $OutputFileName = "HtmlReport.html"
        
        $InputObject =  @{ 
                           Title  = "Directory listing for C:\";
                           Object = Get-Childitem "C:\" | Select -Property FullName,CreationTime,LastWriteTime,Attributes
                        },
                        @{
                           Title  = "PowerShell host details";
                           Object = Get-Host;
                           As     = 'List'
                        },
                        @{
                           Title       = "Running processes";
                           Description = 'Information about the first 2 running processes.'
                           Object      = Get-Process | Select -First 2
                        },
                        @{
                           Object = Get-ChildItem "Env:"
                        }
                        
        Export-HtmlReport -InputObject $InputObject -ReportTitle $ReportTitle -OutputFile $OutputFileName
    
        Creates a HTML report with separated tables for each given object.
    .INPUTS
        Data object, title and alignment parameter
    .OUTPUTS
        File object for the created HTML report file.
    #>
    
        [CmdletBinding()]
        Param(
            [Parameter(ValueFromPipeline=$True, Mandatory=$True)]
            [ValidateNotNullOrEmpty()]
            [Array]$InputObject,
    
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [String]$ReportTitle = 'Generic HTML-Report',
    
            [Parameter(Mandatory=$True)]
            [ValidateNotNullOrEmpty()]
            [String]$OutputFileName
        )
    
        BEGIN
        {
            $HtmlTable		= ''
        }
    
        PROCESS
        {
            ForEach ($InputElement in $InputObject)
            {
                If ($InputElement.ContainsKey('Title') -eq $False)
                {
                    $InputElement.Title = ''
                }
    
                If ($InputElement.ContainsKey('As') -eq $False)
                {
                    $InputElement.As = 'Table'
                }
    
                If ($InputElement.ContainsKey('Property') -eq $False)
                {
                    $InputElement.Property = '*'
                }
    
                If ($InputElement.ContainsKey('Description') -eq $False)
                {
                    $InputElement.Description = ''
                }
    
                $HtmlTable += $InputElement.Object | ConvertTo-HtmlTable -Title $InputElement.Title -Description $InputElement.Description -Property $InputElement.Property -As $InputElement.As
                $HtmlTable += '<br>'
            }
        }
    
        END
        {
            $HtmlTable | New-HtmlReport -Title $ReportTitle | Set-Content $OutputFileName
            Get-Childitem $OutputFileName | Write-Output
        }
    }
    
Function ConvertTo-HtmlTable{

<#
.SYNOPSIS
    Converts a PowerShell object into a HTML table
.DESCRIPTION
    Converts a PowerShell object into a HTML table.	Then use "New-HtmlReport" to create an eye-friendly HTML report with an inline cascading style sheet.
.PARAMETER InputObject
    Any PowerShell object to be converted into a HTML table.
.PARAMETER Property
    Select object properties to be used for table creation. Default is "*"
.PARAMETER As
    Use "List" to create a vertical table. All other values will create a horizontal table.
.PARAMETER Title
    Adds an additional table with a title. Very useful for multi-table-reports!
.PARAMETER Description
    Adds an additional table with a description. Very useful for multi-table-reports!
EXAMPLE
    Get-Process | ConvertTo-HtmlTable
    
    Returns a HTML table as a string.
.EXAMPLE
    Get-Process | ConvertTo-HtmlTable | New-HtmlReport | Set-Content "HtmlReport.html"

    Returns a HTML report and saves it as a file.
.EXAMPLE
    $body =	ConvertTo-HtmlTable -InputObject $(Get-Process) -Property Name,ID,Path -As "List" -Title "Process list" -Description "Shows running processes as a list"
    New-HtmlReport -Body $body | Set-Content "HtmlReport.html"

    Returns a HTML report and saves it as a file.
.INPUTS
    Any PowerShell object
.OUTPUTS
    HTML table as String
#>

    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$True, Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$InputObject,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [Object[]]$Property = '*',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [String]$As = 'TABLE',
        
        [Parameter()]
        [String]$Title,

        [Parameter()]
        [String]$Description
    )
    
    BEGIN
    {
        $InputObjectList = @()
        
        If ($As -ne 'LIST')
        {
            $As = 'TABLE'
        }
    }

    PROCESS
    {
        $InputObjectList += $InputObject
    }

    END
    {
        $ofs = "`r`n"	# Set separator for string-convertion to carrige return
        [String]$HtmlTable = $InputObjectList | ConvertTo-HTML -Property $Property -As $As -Fragment
        Remove-Variable ofs -force

        # Add description table
        If ($Description)
        {
            $Html		= '<table id="TableDescription"><tr><td>' + "$Description</td></tr></table>`n"
            $Html 		+= '<table id="TableSpacer"></table>' + "`n"
            $HtmlTable	= $Html + $HtmlTable
        }
        Else
        {
            $Html 		= '<table id="TableSpacer"></table>' + "`n"
            $HtmlTable	= $Html + $HtmlTable
        }
        
        # Add title table
        If ($Title)
        {
            $Html		= '<table id="TableHeader"><tr><td>' + "$Title</td></tr></table>`n"
            $HtmlTable	= $Html + $HtmlTable
        }

        # Add missing data separator tag <hr> to second column (on list-tables only)
        $HtmlTable = $HtmlTable -Replace '<hr>', '<hr><td><hr></td>'
                
        Write-Output $HtmlTable	
    }
}
    
Function New-HtmlReport{
    
    <#
    .SYNOPSIS
        Creates a HTML report
    .DESCRIPTION
        Creates an eye-friendly HTML report with an inline cascading style sheet for a given HTML body.
        Usage of "ConvertTo-HtmlTable" is recommended to create the HTML body.
    .PARAMETER Body
        Any HTML body, e.g. a table. Usage of "ConvertTo-HtmlTable" is recommended
        to create an according table from any PowerShell object.
    .PARAMETER Title
        Title of the HTML report.
    .PARAMETER Head
        Any HTML code to be inserted into the head-tag, e.g. scripts or meta-information.
    .PARAMETER CssUri
        Path to a CSS-File to be included as an inline css.
        If CssUri is invalid or not provided, a default css is used instead.
    .EXAMPLE
        Get-Process | ConvertTo-HtmlTable | New-HtmlReport
        
        Returns a HTML report as a string.
    .EXAMPLE
        Get-Process | ConvertTo-HtmlTable | New-HtmlReport -Title "HTML Report with CSS" | Set-Content "HtmlReport.html"
    
        Returns a HTML report and saves it as a file.
    .EXAMPLE
        $body =	Get-Process | ConvertTo-HtmlTable
        New-HtmlReport -Body $body -Title "HTML Report with CSS" -Head '<meta name="author" content="Thomas Franke">' -CssUri "stylesheet.css" | Set-Content "HtmlReport.html"
    
        Returns a HTML report with an alternative CSS and saves it as a file.
    .INPUTS
        HTML body as String
    .OUTPUTS
        HTML page as String
    #>
    
    [CmdletBinding()]
    Param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [String]$CssUri,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [String]$Title = 'HTML Report',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [String]$Head = '',

        [Parameter(ValueFromPipeline=$True, Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Array]$Body
    )

    # Add title to head because -Title parameter is ignored if -Head parameter is used
    If ($Title){
        $Head = "<title>$Title</title>`n$Head`n"
    }

    # Add inline stylesheet
    If (($CssUri) -And (Test-Path $CssUri)){
        $Head += "<style>`n" + $(Get-Content $CssUri | ForEach {"`t$_`n"}) + "</style>`n"
    }
    Else{
        $Head += @'
<style>
table
    {
        Margin: 0px 0px 0px 4px;
        Border: 1px solid rgb(190, 190, 190);
        Font-Family: Tahoma;
        Font-Size: 8pt;
        Background-Color: rgb(252, 252, 252);
    }

tr:hover td
    {
        Background-Color: rgb(150, 150, 220);
        Color: rgb(255, 255, 255);
    }

tr:nth-child(even)
    {
        Background-Color: rgb(242, 242, 242);
    }
    
th
    {
        Text-Align: Left;
        Color: rgb(150, 150, 220);
        Padding: 1px 4px 1px 4px;
    }


td
    {
        Vertical-Align: Top;
        Padding: 1px 4px 1px 4px;
    }

#TableHeader
    {
        Margin-Bottom: -1px;
        Background-Color: rgb(255, 255, 225);
        Width: 30%;
    }
    
#TableDescription
    {
        Background-Color: rgb(252, 252, 252);
        Width: 30%;
    }

#TableSpacer
    {
        Height: 6px;
        Border: 0px;
    }
</style>
'@
    }

    $HtmlReport = ConvertTo-Html -Head $Head
    
    # Delete empty table that is created because no input object was given
    $HtmlReport = $($HtmlReport -Replace '<table>', '') -Replace '</table>', $Body

    Write-Output $HtmlReport
}
#endregion

#region Variable Declarations
$Credential = Import-SavedCredential glo-sa-powershell
$servers = $env:servers.split(',').trim()
If(!(Test-Path "$ENV:Workspace\attachments")){
    Try{
        New-item -Path $ENV:Workspace -Name 'Attachments' -ItemType Directory -ErrorAction Stop
    } 
    Catch{
        Write-Host $_
        exit 1;
    }
}
$emails = $ENV:Email.split(',').trim()
#endregion

#region Main Script
Foreach($Server in $Servers){
    Write-Host "Running report on $Server"
    #Process each Server
    $report_info = Invoke-command -ComputerName $Server -Credential $Credential -ScriptBlock {
        
        #region Disk info
        $Diskinfo = 
        Get-WmiObject Win32_DiskDrive | % {
            $disk = $_
            $partitions = "ASSOCIATORS OF " + "{Win32_DiskDrive.DeviceID='$($disk.DeviceID)'} " + "WHERE AssocClass = Win32_DiskDriveToDiskPartition"
            Get-WmiObject -Query $partitions | % {
                $partition = $_
                $drives = "ASSOCIATORS OF " + "{Win32_DiskPartition.DeviceID='$($partition.DeviceID)'} " + "WHERE AssocClass = Win32_LogicalDiskToPartition"
                    Get-WmiObject -Query $drives | % {
                        $DiskSize  = '{0:d} GB' -f [int]($Disk.Size / 1GB)
                        $RawSize   = '{0:d} GB' -f [int]($partition.Size / 1GB)
                        $Freespace = '{0:d} GB' -f [int]($_.FreeSpace / 1GB)
                        $DriveSize = '{0:d} GB' -f [int]($_.Size / 1GB)
                            New-Object -Type PSCustomObject -Property @{
                                Disk        = $disk.DeviceID
                                DiskModel   = $disk.Model
                                DiskSize    = $disksize
                                DriveLetter = $_.DeviceID
                                DriveName   = $_.VolumeName
                                DriveSize   = $Drivesize
                                Partition   = $partition.Name
                                RawSize     = $rawsize
                                FreeSpace   = $FreeSpace
                            }
                    }
            }
        }
        #endregion

        #region Agent Info
        $solarwinds_status = Get-Service SolarWindsAgent64 -ErrorAction SilentlyContinue
        If(!$solarwinds_status){
            $solarwinds_status_obj = New-object -TypeName PSobject -Property @{
                Status = "MISSING!"
                Name = 'SolarWindsAgent64'
                Displayname = 'SolarWinds Agent'
            }
        }
        Else{
            $solarwinds_status_obj = New-object -TypeName PSobject -Property @{
                Status = $solarwinds_status.status
                Name = $solarwinds_status.name
                Displayname = $solarwinds_status.displayname
            }
        }

        $cisco_amp_status = get-service ciscoamp -ErrorAction SilentlyContinue
        If(!$cisco_amp_status){
            $cisco_amp_status_obj = New-object -TypeName PSobject -Property @{
                Status = 'MISSING!'
                Name = 'ciscoamp'
                Displayname = 'Cisco Secure Endpoint'
            }
        }
        Else{
            $cisco_amp_status_obj = New-object -TypeName PSobject -Property @{
                Status = $cisco_amp_status.status
                Name = $cisco_amp_status.name
                Displayname = $cisco_amp_status.displayname
            }
        }

        $mecp_status = Get-Service ccmexec  -ErrorAction SilentlyContinue
        If(!$mecp_status){
            $mecp_status_obj = New-object -TypeName PSobject -Property @{
                Status = "MISSING!"
                Name = 'ccmexec'
                Displayname = 'SMS Agent Host'
            }
        }
        Else{
            $mecp_status_obj = New-object -TypeName PSobject -Property @{
                Status = $mecp_status.status
                Name = $mecp_status.name
                Displayname = $mecp_status.displayname
            }
        }

        $splunk_status = get-service SplunkForwarder  -ErrorAction SilentlyContinue
        If(!$splunk_status){
            $splunk_status_obj = New-object -TypeName PSobject -Property @{
                Status = "MISSING!"
                Name = 'SplunkForwarder'
                Displayname = 'Splunk Agent'
            }
        }
        Else{
            $splunk_status_obj = New-object -TypeName PSobject -Property @{
                Status = $splunk_status.status
                Name = $splunk_status.name
                Displayname = $splunk_status.displayname
            }
        }

        #endregion

        #region NTP Info
        $ntp_status = w32tm /query /status
        $ntp_last_sync = ($ntp_status | where-object {$_ -like '*Last Successful*'}).trim('Last Successful Sync Time: ')
        $ntp_source = ($ntp_status | where-object {$_ -like '*Source:*'}).trim('Source: ')
        $ntp_status_obj = New-object -TypeName PSobject -Property @{
            Source = $ntp_source
            LastSync = $ntp_last_sync
        }
        #endregion

        #region remote desktop members
        $Groupname = 'Remote Desktop Users'
        $Group = [ADSI]"WinNT://$env:COMPUTERNAME/$groupname" 
        $group_members = @($group.Invoke('Members') | % {([adsi]$_).path})
        $remote_desktop_users = @()
        Foreach($member in $group_members){
            $name = $member.replace('WinNT://','')

            $remote_desktop_users += New-Object -TypeName psobject -Property @{
                Name = $Name
            }
        }
        #endregion

        #region administrator members
        $Groupname = 'Administrators'
        $Group = [ADSI]"WinNT://$env:COMPUTERNAME/$groupname" 
        $group_members = @($group.Invoke('Members') | % {([adsi]$_).path})
        $administrators_users = @()
        Foreach($member in $group_members){
            $name = $member.replace('WinNT://','')

            $administrators_users += New-Object -TypeName psobject -Property @{
                Name = $Name
            }
        }
        #endregion

        #region Result Object Creation
        $log_list = @('Application','Security','System')
        
        $results = New-Object -TypeName psobject -Property @{
            local_admins = $administrators_users
            local_users = Get-LocalUser -ErrorAction stop
            remote_desktop_members = $remote_desktop_users
            telnet_status = Get-WindowsFeature -Name 'telnet-server' -ErrorAction stop
            ftp_status = Get-WindowsFeature | Where-Object Name -like '*Ftp*' -ErrorAction stop
            solarwinds_status = $solarwinds_status_obj
            ntp_status = $ntp_status_obj
            mecp_status = $mecp_status_obj
            max_eventlog_sizes = get-eventlog -list | where-object {$_.log -in $log_list} |  Select-Object Log, MaximumKilobytes
            cisco_amp_status = $cisco_amp_status_obj
            splunk_status = $splunk_status_obj
            net_adapters = Get-WmiObject win32_networkadapterconfiguration | where-object {$_.ipenabled -eq $true}
            disk_info = $diskinfo
            cpu_info = Get-WmiObject Win32_Processor
            system_data = Get-WmiObject Win32_ComputerSystem
            os_data = Get-WmiObject Win32_OperatingSystem
        }
        #endregion
        $results
    } -ErrorAction Stop

    Write-Host "Info Gathered. Formatting Output."
    #Region format Output
    $GeneralInfo_OBJ = New-Object -TypeName PSObject
        $GeneralInfo_OBJ | Add-Member -MemberType NoteProperty -Name "Server Name"  -Value $report_info.system_data.Name
        $GeneralInfo_OBJ | Add-Member -MemberType NoteProperty -Name "Domain Name"  -Value $report_info.system_data.Domain
        $GeneralInfo_OBJ | Add-Member -MemberType NoteProperty -Name "OS"           -Value $report_info.os_data.Caption
        $GeneralInfo_OBJ | Add-Member -MemberType NoteProperty -Name "Service Pack" -Value $report_info.os_data.ServicePackMajorVersion

    $HardwareInfo_OBJ = New-Object -TypeName PSObject
        $HardwareInfo_OBJ | Add-Member -MemberType NoteProperty -Name "Total CPUs"   -Value $report_info.cpu_info.count 
        $HardwareInfo_OBJ | Add-Member -MemberType NoteProperty -Name "Total RAM"    -Value $(($report_info.os_data.totalvisiblememorysize / 1MB).tostring("F00") + "GB")

    $DiskInfo_OBJ = New-Object -TypeName PSObject
        $DiskInfo_OBJ | Add-Member -MemberType NoteProperty -Name "Disks"   -Value $($report_info.disk_info | select DiskSize, Disk, DiskModel, DriveName, DriveLetter, DriveSize, FreeSpace)

    $NetInfo_OBJ = New-Object -TypeName PSObject
        $NetInfo_OBJ | Add-Member -MemberType NoteProperty -Name "Adapters"   -Value $($report_info.net_adapters | select Description, IPAddress, DefaultIPGateway, IPSubnet, DNSServerSearchOrder)

    $Agentsarray = @($report_info.solarwinds_status, $report_info.mecp_status,$report_info.cisco_amp_status,$report_info.splunk_status)
    #endregion

    #region OutputThe results to HTML.
        
    $OutputFileName = "$ENV:Workspace\attachments\$Server" + " - Audit Report - " + "$(get-date -Format "dd-MM-yyyy")" + '.html'
    $ReportTitle = "Server Audit Report for $Server"
    $InputObject =  @{
                        Title  = "General Server Information";
                        Description = 'The following table contains General Server Information'
                        Object = $GeneralInfo_OBJ
                    },
                    @{
                        Title  = "Hardware Information";
                        Description = 'The following table contains Hardware Information.'
                        Object = $HardwareInfo_OBJ
                    },
                    @{
                        Title  = "Disk Information";
                        Description = 'The following table contains Disk Information.'
                        Object = $DiskInfo_OBJ.disks | sort-object -Property DriveLetter
                    },
                    @{
                        Title  = "Network Information";
                        Description = 'The following table contains Network Information.'
                        Object = $NetInfo_OBJ.Adapters
                    },
                    @{
                        Title  = "Agents Information";
                        Description = 'The following table contains Agent Information.'
                        as = 'List'
                        Object = $Agentsarray
                    },
                    @{
                        Title  = "Local Admins";
                        Description = 'The following table lists all users and groups in Local Admins.'
                        Object = $report_info.local_admins
                    },
                    @{
                        Title  = "Local Users";
                        Description = 'The following table lists all Local Users.'
                        Object = $report_info.local_users | Select-Object -Property Name, Enabled, Description,PasswordExpires,PasswordLastSet
                    },
                    @{
                        Title  = "Remote Desktop Members";
                        Description = 'The following table lists all members explicitly in the Remote Desktop Group.'
                        Object = $report_info.remote_desktop_members | select -Property Name
                    },
                    @{
                        Title  = "Telnet";
                        Description = 'The following table shows the status of the Telnet Windows Feature.'
                        Object = $report_info.telnet_status | Select-Object -Property DisplayName, Description, Installed
                    },
                    @{
                        Title  = "FTP";
                        Description = 'The following table shows the status of the FTP Related Windows Features'
                        Object = $report_info.ftp_status | Select-Object -Property DisplayName, Description, Installed
                    },
                    @{
                        Title  = "NTP";
                        Description = 'The following table shows the NTP Settings and Information'
                        Object = $report_info.ntp_status
                    },
                    @{
                        Title  = "Event Logs";
                        Description = 'The following table shows theMaximum Sizes for the Event Logs'
                        Object = $report_info.max_eventlog_sizes
                    }
                    
        Export-HtmlReport -InputObject $InputObject -ReportTitle $ReportTitle -OutputFile $OutputFileName -Verbose
        #endregion
    #endregion
    }
    
    Send-MailMessage -From <email_address> -to $emails -cc 'someone@domain.com' -Subject "Server Audit Report Results" -Attachments $((get-childitem "$ENV:Workspace\attachments").fullname) -SmtpServer <smtp_server>