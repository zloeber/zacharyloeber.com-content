#Set-StrictMode -Version 2

#region System Report General Options
$Option_EventLogPeriod = 24                 # in hours
$Option_EventLogResults = 5                 # Number of event logs per log type returned
$Option_TotalProcessesByMemory = 5          # Number of top memory using processes to return
$Option_TotalProcessesByMemoryWarn = 100    # Warning highlight on processes over MB amount
$Option_TotalProcessesByMemoryAlert = 300   # Alert highlight on processes over MB amount
$Option_DriveUsageWarningThreshold = 80     # Warning at this percentage of drive space used
$Option_DriveUsageAlertThreshold = 90       # Alert at this percentage of drive space used
$Option_DriveUsageWarningColor = 'Orange'
$Option_DriveUsageAlertColor = 'Red'
$Option_DriveUsageColor = 'Green'
$Option_DriveFreeSpaceColor = 'Transparent'
# Try to keep this as an even number for the best results. The larger you make this
# number the less flexible your columns will be in html reports.
$DiskGraphSize = 26
#endregion System Report General Options

#region Global Options
# Change this to allow for more or less result properties to span horizontally
#  anything equal to or above this threshold will get displayed vertically instead.
#  (NOTE: This only applies to sections set to be dynamic in html reports)
$HorizontalThreshold = 10
#endregion Global Options

#region System Report Section Postprocessing Definitions
# If you are going to do some post-processing love then be cognizent of the following:
#  - The only variable which goes through post-processing is the section table as html.
#    This variable is aptly called $Table and will contain a string with a full html table.
#  - When you are done doing whatever processing you are aiming to do please return the fully formated 
#    html.
#  - I don't know whether to be proud or ashamed of this code. I think probably ashamed....
#  - These are assigned later on in the report structure as hash key entries 'PostProcessing'
# For this example I've performed two colorize table checks on the memory utilization.
$ProcessesByMemory_Postprocessing = 
@'
    [scriptblock]$scriptblock = {[float]$($args[0]|ConvertTo-MB) -gt [int]$args[1]}
    $temp = Colorize-Table $Table -Scriptblock $scriptblock -Column 'Memory Usage (WS)' -ColumnValue $Option_TotalProcessesByMemoryWarn -Attr 'class' -AttrValue 'warn' -WholeRow $true
    Colorize-Table $temp -Scriptblock $scriptblock -Column 'Memory Usage (WS)' -ColumnValue $Option_TotalProcessesByMemoryAlert -Attr 'class' -AttrValue 'alert' -WholeRow $true
'@

$RouteTable_Postprocessing = 
@'
    $temp = Colorize-Table $Table -Column 'Persistent' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
    Colorize-Table $temp -Column 'Persistent' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
'@

# For this example I've performed two colorize table checks on the event log entries.
$EventLogs_Postprocessing =
@'
    $temp = Colorize-Table $Table -Column 'Type' -ColumnValue 'Warning' -Attr 'class' -AttrValue 'warn' -WholeRow $true
    $temp = Colorize-Table $temp  -Column 'Type' -ColumnValue 'Error' -Attr 'class' -AttrValue 'alert' -WholeRow $true
    Colorize-Table $temp -Column 'Log' -ColumnValue 'Security' -Attr 'class' -AttrValue 'security' -WholeRow $true
'@

$HPServerHealth_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -ne [string]$args[1]}
    $temp = Colorize-Table $Table -Column 'Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
    Colorize-Table $temp -Scriptblock $scriptblock -Column 'Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'    
'@

$HPServerHealthArrayController_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -ne [string]$args[1]}
    $temp = Colorize-Table $Table  -Scriptblock $scriptblock -Column 'Battery Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'    
    $temp = Colorize-Table $temp -Column 'Battery Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
    $temp = Colorize-Table $temp  -Scriptblock $scriptblock -Column 'Controller Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'    
    $temp = Colorize-Table $temp -Column 'Controller Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
    $temp = Colorize-Table $temp  -Scriptblock $scriptblock -Column 'Array Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'    
    Colorize-Table $temp -Column 'Array Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
'@

# Printer
$Printer_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -ne [string]$args[1]}
    $temp = Colorize-Table $Table -Scriptblock $scriptblock  -Column 'Status' -ColumnValue 'Idle' -Attr 'class' -AttrValue 'warn'
    $temp = Colorize-Table $temp -Scriptblock $scriptblock  -Column 'Job Errors' -ColumnValue '0' -Attr 'class' -AttrValue 'warn'
    $temp = Colorize-Table $temp -Column 'Status' -ColumnValue 'Idle' -Attr 'class' -AttrValue 'healthy'
    $temp = Colorize-Table $temp -Column 'Shared' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
    $temp = Colorize-Table $temp -Column 'Shared' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
    $temp = Colorize-Table $temp -Column 'Published' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
    Colorize-Table $temp -Column 'Published' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
'@

$ComputerReportPreProcessing =
@'
    Gather-ReportInformation @credsplat @VerboseDebug `
               -ComputerName $AssetNames `
               -ReportContainer $ReportContainer `
               -SortedRpts $SortedReports
'@
#endregion Report Section Postprocessing Definitions

#region System Report Structure
<#
 This hash deserves some explanation. Each report section is the first referenced key.
 For each section there are several subkeys:
  Enabled - ($true/$false)
    Determines if the section is enabled. Use this to disable/enable a section
    for ALL report types.
  AllowEmptyReport - ($true/$false)
    Determines if the section will still process when there is no data. Use this 
    to force report layouts into a specific patterns if data is variable.
    
  Order - ([int])
    Hash tables in powershell v2 have no easy way to maintain a specific order. This is used to 
    workaround that limitation is a hackish way. You can have duplicates but then section order
    will become unpredictable.
    
  AllData - ([hashtable])
    This holds a hashtable with all data which is being reported upon. You will load this up
    in Gather-ReportInformation. It is up to you to fill the data appropriately if a new type
    of report is being templated out for your poject. AllData expects a hash of names with their
    value being an array of values.
    
  Title - ([string])
    The section title for the top of the table. This spans across all columns and looks nice.

  PostProcessing - ([string])
    Used to colorize table elements before putting them into an html report
    
  ReportTypes - [hashtable]
    This is the meat and potatoes of each section. For each report type you have defined there will 
    generally be several properties which are selected directly from the AllData hashes which
    make up your report. Several advanced report type definitions have been included for the
    system report as examples. Generally each section contains the same report types as top
    level hash keys. There are some special keys which can be defined with each report type
    that allow you to manually manipulate how reports are generated per section. These are:
    
      SectionOverride - ($true/$false)
        When a section break is determined this will ignore the break and keep this section part 
        of the prior section group. This is an advanced layout option. This is almost always
        going to be $false.
        
      ContainerType - (See below)
        Use ths to force a particular report element to use a specific section container. This 
        affects how the element gets laid out on the page. So far the following values have
        been defined.
           Half    - The report section consumes half of the row. Even report sections end up on 
                     the left side, odd report sections end up on the right side.
           Full    - The report section consumes the entire width of the row.
           Third   - The report section consumes approximately a third of the row.
           TwoThirds - The report section consumes approximately 2/3rds of the row.
           Fourth  - The section consumes a fourth of the row.
           ThreeFourths - Ths section condumes 3/4ths of the row.
           
        You can end up with some wonky looking reports if you don't plan the sections appropriately
        to match your sectional data. So a section with 20 properties and a horizontal layout will
        look like crap taking up only a 4th of the page.
        
      TableType - (See below) 
        Use this to force a particular report element to use a specific table layout. Thus far
        the following vales have been defined.
           Vertical   - Data headers are the first row of the table
           Horizontal - Data headers are the first column of the table
           Dynamic    - If the number of data properties equals or surpasses the HorizontalThreshold 
                        the table is presented vertically. Otherwise it displays horizontally
#>
$SystemReport = @{
    'Configuration' = @{
        'TOC' = $true           #Possibly used in the future to create a table of contents
        'PreProcessing' = $ComputerReportPreProcessing
    }
    'Sections' = @{
        'Break_Summary' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 0
            'AllData' = @{}
            'Title' = 'System Information'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'Summary' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 1
            'AllData' = @{}
            'Title' = 'System Summary'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Uptime';e={$_.Uptime}},
                        @{n='OS';e={$_.OperatingSystem}},
                        @{n='Total Physical RAM';e={$_.PhysicalMemoryTotal}},
                        @{n='Free Physical RAM';e={$_.PhysicalMemoryFree}},
                        @{n='Total RAM Utilization';e={"$($_.PercentPhysicalMemoryUsed)%"}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='OS';e={$_.OperatingSystem}},
                        @{n='OS Architecture';e={$_.OSArchitecture}},
                        @{n='OS Service Pack';e={$_.OSServicePack}},
                        @{n='OS SKU';e={$_.OSSKU}},
                        @{n='OS Version';e={$_.OSVersion}},
                        @{n='Server Chassis Type';e={$_.ChassisModel}},
                        @{n='Server Model';e={$_.Model}},
                        @{n='Serial Number';e={$_.SerialNumber}},
                        @{n='CPU Architecture';e={$_.SystemArchitecture}},
                        @{n='CPU Sockets';e={$_.CPUSockets}},
                        @{n='Total CPU Cores';e={$_.CPUCores}},
                        @{n='Virtual';e={$_.IsVirtual}},
                        @{n='Virtual Type';e={$_.VirtualType}},
                        @{n='Total Physical RAM';e={$_.PhysicalMemoryTotal}},
                        @{n='Free Physical RAM';e={$_.PhysicalMemoryFree}},
                        @{n='Total Virtual RAM';e={$_.VirtualMemoryTotal}},
                        @{n='Free Virtual RAM';e={$_.VirtualMemoryFree}},
                        @{n='Total Memory Slots';e={$_.MemorySlotsTotal}},
                        @{n='Memory Slots Utilized';e={$_.MemorySlotsUsed}},
                        @{n='Uptime';e={$_.Uptime}},
                        @{n='Install Date';e={$_.InstallDate}},
                        @{n='Last Boot';e={$_.LastBootTime}},
                        @{n='System Time';e={$_.SystemTime}}
                }
            }
        }
        'ExtendedSummary' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 2
            'AllData' = @{}
            'Title' = 'Extended Summary'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $true
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Registered Owner';e={$_.RegisteredOwner}},
                        @{n='Registered Organization';e={$_.RegisteredOrganization}},
                        @{n='System Root';e={$_.SystemRoot}},
                        @{n='Product Key';e={ConvertTo-ProductKey $_.DigitalProductId}},
                        @{n='Product Key (64 bit)';e={ConvertTo-ProductKey $_.DigitalProductId4 -x64}},
                        @{n='NTP Type';e={$_.NTPType}},
                        @{n='NTP Servers';e={$_.NTPServers}}
                }
            }
        }
        'DellWarrantyInformation' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 3
            'AllData' = @{}
            'Title' = 'Dell Warranty Information'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =                    
                        @{n='Type';e={$_.Type}},
                        @{n='Model';e={$_.Model}},                    
                        @{n='Service Tag';e={$_.ServiceTag}},
                        @{n='Ship Date';e={$_.ShipDate}},
                        @{n='Start Date';e={$_.StartDate}},
                        @{n='End Date';e={$_.EndDate}},
                        @{n='Days Left';e={$_.DaysLeft}},
                        @{n='Service Level';e={$_.ServiceLevel}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $true
                    'TableType' = 'Vertical'
                    'Properties' =                    
                        @{n='Type';e={$_.Type}},
                        @{n='Model';e={$_.Model}},                    
                        @{n='Service Tag';e={$_.ServiceTag}},
                        @{n='Ship Date';e={$_.ShipDate}},
                        @{n='Start Date';e={$_.StartDate}},
                        @{n='End Date';e={$_.EndDate}},
                        @{n='Days Left';e={$_.DaysLeft}},
                        @{n='Service Level';e={$_.ServiceLevel}}
                }
            }
        }
        'Disk' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 4
            'AllData' = @{}
            'Title' = 'Disk Report'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Drive';e={$_.Drive}},
                        @{n='Type';e={$_.DiskType}},
                        @{n='Size';e={$_.DiskSize}},
                        @{n='Free Space';e={$_.FreeSpace}},
                        @{n='Disk Usage';
                          e={$color = $Option_DriveUsageColor
                            if ((100 - $_.PercentageFree) -ge $Option_DriveUsageWarningThreshold)
                            {
                                if ((100 - $_.PercentageFree) -ge $Option_DriveUsageAlertThreshold)
                                {
                                    $color = $Option_DriveUsageAlertColor
                                }
                                else
                                {
                                    $color = $Option_DriveUsageWarningColor
                                }
                            }
                            New-HTMLBarGraph -GraphSize $DiskGraphSize -PercentageUsed (100 - $_.PercentageFree) `
                                             -LeftColor $color -RightColor $Option_DriveFreeSpaceColor
                            }}
                }    
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Drive';e={$_.Drive}},
                        @{n='Type';e={$_.DiskType}},
                        @{n='Disk';e={$_.Disk}},
                        @{n='Serial Number';e={$_.SerialNumber}},
                        @{n='Model';e={$_.Model}},
                        @{n='Partition ';e={$_.Partition}},
                        @{n='Size';e={$_.DiskSize}},
                        @{n='Free Space';e={$_.FreeSpace}},
                        @{n='Disk Usage';
                          e={$color = $Option_DriveUsageColor
                            if ((100 - $_.PercentageFree) -ge $Option_DriveUsageWarningThreshold)
                            {
                                if ((100 - $_.PercentageFree) -ge $Option_DriveUsageAlertThreshold)
                                {
                                    $color = $Option_DriveUsageAlertColor
                                }
                                else
                                {
                                    $color = $Option_DriveUsageWarningColor
                                }
                            }
                            New-HTMLBarGraph -GraphSize $DiskGraphSize -PercentageUsed (100 - $_.PercentageFree) `
                                             -LeftColor $color -RightColor $Option_DriveFreeSpaceColor
                            }}
                }
            }
        }
        'Memory' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 5
            'AllData' = @{}
            'Title' = 'Memory Banks'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Bank';e={$_.Bank}},
                        @{n='Label';e={$_.Label}},
                        @{n='Capacity';e={$_.Capacity}},
                        @{n='Speed';e={$_.Speed}},
                        @{n='Detail';e={$_.Detail}},
                        @{n='Form Factor';e={$_.FormFactor}}
                }
            }
        }
        'ProcessesByMemory' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 6
            'AllData' = @{}
            'Title' = 'Top Processes by Memory'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='ID';e={$_.ProcessID}},
                        @{n='Memory Usage (WS)';e={$_.WS | ConvertTo-KMG}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='ID';e={$_.ProcessID}},
                        @{n='Memory Usage (WS)';e={$_.WS | ConvertTo-KMG}}
                }
            }
            'PostProcessing' = $ProcessesByMemory_Postprocessing
        }
        'StoppedServices' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 7
            'AllData' = @{}
            'Title' = 'Stopped Services'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Service Name';e={$_.Name}},
                        @{n='State';e={$_.State}},
                        @{n='Start Mode';e={$_.StartMode}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Service Name';e={$_.Name}},
                        @{n='State';e={$_.State}},
                        @{n='Start Mode';e={$_.StartMode}}
                }
            }
        }
        'NonStandardServices' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 8
            'AllData' = @{}
            'Title' = 'NonStandard Service Accounts'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Service Name';e={$_.Name}},
                        @{n='State';e={$_.State}},
                        @{n='Start Mode';e={$_.StartMode}},
                        @{n='Start As';e={$_.StartName}}
                }
            }
        }
 
        'Break_EventLogs' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 10
            'AllData' = @{}
            'Title' = 'Event Log Information'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'EventLogSettings' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 11
            'AllData' = @{}
            'Title' = 'Event Log Settings'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.LogfileName}},
                        @{n='Status';e={$_.Status}},
                        @{n='OverWrite';e={$_.OverWritePolicy}},
                        @{n='Entries';e={$_.NumberOfRecords}},
                        #@{n='Archive';e={$_.Archive}},
                        #@{n='Compressed';e={$_.Compressed}},
                        @{n='Max File Size';e={$_.MaxFileSize}}
                }
            }
        }
        'EventLogs' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 12
            'AllData' = @{}
            'Title' = 'Event Log Settings'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Log';e={$_.LogFile}},
                        @{n='Type';e={$_.Type}},
                        @{n='Source';e={$_.SourceName}},
                        @{n='Event';e={$_.EventCode}},
                        @{n='Message';e={$_.Message}},
                        @{n='Time';e={([wmi]'').ConvertToDateTime($_.TimeGenerated)}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Log';e={$_.LogFile}},
                        @{n='Type';e={$_.Type}},
                        @{n='Source';e={$_.SourceName}},
                        @{n='Event';e={$_.EventCode}},
                        @{n='Message';e={$_.Message}},
                        @{n='Time';e={([wmi]'').ConvertToDateTime($_.TimeGenerated)}}
                }
            }
            'PostProcessing' = $EventLogs_Postprocessing
        }
        
        'Break_Network' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 20
            'AllData' = @{}
            'Title' = 'Networking Information'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'Network' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 21
            'AllData' = @{}
            'Title' = 'Network Adapters'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Network Name';e={$_.NetworkName}},
                        @{n='Adapter Name';e={$_.AdapterName}},
                        @{n='Index';e={$_.Index}},
                        @{n='Ip Address';e={$_.IpAddress -join ', '}},
                        @{n='Ip Subnet';e={$_.IpSubnet -join ', '}},
                        @{n='MAC Address';e={$_.MACAddress}},
                        @{n='Gateway';e={$_.DefaultIPGateway}},
                        @{n='Description';e={$_.Description}},
                       # @{n='Interface Index';e={$_.InterfaceIndex}},
                        @{n='DHCP Enabled';e={$_.DHCPEnabled}},
                        @{n='Connection Status';e={$_.ConnectionStatus}},
                        @{n='Promiscuous Mode';e={$_.PromiscuousMode}}
                }         
            }

        }
        'RouteTable' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 22
            'AllData' = @{}
            'Title' = 'Route Table'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Destination';e={$_.Destination}},
                        @{n='Mask';e={$_.Mask}},
                        @{n='Next Hop';e={$_.NextHop}},
                        @{n='Persistent';e={$_.Persistent}},
                        @{n='Metric';e={$_.Metric}},
                        @{n='Interface Index';e={$_.InterfaceIndex}},
                        @{n='Type';e={$_.Type}}
                }         
            }
            'PostProcessing' = $RouteTable_Postprocessing
        }

        'Break_SoftwareAudit' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 30
            'AllData' = @{}
            'Title' = 'Software Audit'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'InstalledUpdates' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 31
            'AllData' = @{}
            'Title' = 'Installed Windows Updates'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'ThreeFourths'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='HotFixID';e={$_.HotFixID}},
                        @{n='Description';e={$_.Description}},
                        @{n='Installed By';e={$_.InstalledBy}},
                        @{n='Installed On';e={$_.InstalledOn}},
                        @{n='Link';e={$_.Caption}}
                }
            }
        } 
        'Applications' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 32
            'AllData' = @{}
            'Title' = 'Installed Applications'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'ThreeFourths'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Display Name';e={$_.DisplayName}},
                        @{n='Publisher';e={$_.Publisher}}
                }
            }
        }
        'WSUSSettings' = @{
            'Enabled' = $false
            'AllowEmptyReport' = $false
            'Order' = 33
            'AllData' = @{}
            'Title' = 'WSUS Settings'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $true
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='WSUS Setting';e={$_.Key}},
                        @{n='Value';e={$_.KeyValue}}
                }
            }
        }
        
        'Break_FilePrint' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 40
            'AllData' = @{}
            'Title' = 'File Print'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'Shares' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 41
            'AllData' = @{}
            'Title' = 'Shares'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Share Name';e={$_.Name}},
                        @{n='Status';e={$_.Status}},
                        @{n='Type';e={$ShareType[[string]$_.Type]}},
                        @{n='Allow Maximum';e={$_.AllowMaximum}},
                        @{n='Maximum Allowed';e={$_.MaximumAllowed}}
                }
            }
        }    
        'ShareSessionInfo' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 42
            'AllData' = @{}
            'Title' = 'Share Sessions'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Share Name';e={$_.Name}},
                        @{n='Sessions';e={$_.Count}}
                }
            }
        }
        'Printers' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 43
            'AllData' = @{}
            'Title' = 'Printers'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Status';e={$_.Status}},
                        @{n='Location';e={$_.Location}},
                        @{n='Shared';e={$_.Shared}},
                        @{n='Share Name';e={$_.ShareName}},
                        @{n='Published';e={$_.Published}},
          #              @{n='Local';e={$_.Local}},
          #              @{n='Network';e={$_.Network}},
          #              @{n='Keep Printed Jobs';e={$_.KeepPrintedJobs}},
          #              @{n='Driver Name';e={$_.DriverName}},
                        @{n='Port Name';e={$_.PortName}},
          #              @{n='Default';e={$_.Default}},
                        @{n='Current Jobs';e={$_.CurrentJobs}},
                        @{n='Jobs Printed';e={$_.TotalJobsPrinted}},
                        @{n='Pages Printed';e={$_.TotalPagesPrinted}},
                        @{n='Job Errors';e={$_.JobErrors}}
                }
            }
            'PostProcessing' = $Printer_Postprocessing
        }    
        
        'Break_LocalSecurity' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $true
            'Order' = 50
            'AllData' = @{}
            'Title' = 'Local Security'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'LocalGroupMembership' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 51
            'AllData' = @{}
            'Title' = 'Local Group Membership'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Group Name';e={$_.Group}},
                        @{n='Member';e={$_.GroupMember}},
                        @{n='Member Type';e={$_.MemberType}}
                }
            }
        }
        'AppliedGPOs' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 52
            'AllData' = @{}
            'Title' = 'Applied Group Policies'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Enabled';e={$_.Enabled}},
                        @{n='Source OU';e={$_.SourceOU}},
                        @{n='Link Order';e={$_.linkOrder}},
                        @{n='Applied Order';e={$_.appliedOrder}},
                        @{n='No Override';e={$_.noOverride}}
                }
            }
        }
        
        'Break_HardwareHealth' = @{
            'Enabled' = $false
            'AllowEmptyReport' = $true
            'Order' = 80
            'AllData' = @{}
            'Title' = 'Hardware Health'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'HP_GeneralHardwareHealth' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 90
            'AllData' = @{}
            'Title' = 'HP Overall Hardware Health'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Manufacturer';e={$_.Manufacturer}},
                        @{n='Status';e={$_.HealthState}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Manufacturer';e={$_.Manufacturer}},
                        @{n='Status';e={$_.HealthState}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_EthernetTeamHealth' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 91
            'AllData' = @{}
            'Title' = 'HP Ethernet Team Health'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Description';e={$_.Description}},
                        @{n='Status';e={$_.RedundancyStatus}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Description';e={$_.Description}},
                        @{n='Status';e={$_.RedundancyStatus}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_ArrayControllerHealth' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 92
            'AllData' = @{}
            'Title' = 'HP Array Controller Health'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' =  @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.ArrayName}},
                        @{n='Array Status';e={$_.ArrayStatus}},
                        @{n='Battery Status';e={$_.BatteryStatus}},
                        @{n='Controller Status';e={$_.ControllerStatus}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.ArrayName}},
                        @{n='Array Status';e={$_.ArrayStatus}},
                        @{n='Battery Status';e={$_.BatteryStatus}},
                        @{n='Controller Status';e={$_.ControllerStatus}}
                }
            }
            'PostProcessing' = $HPServerHealthArrayController_Postprocessing
        }
        'HP_EthernetHealth' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 93
            'AllData' = @{}
            'Title' = 'HP Ethernet Adapter Health'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Port Type';e={$_.PortType}},
                        @{n='Port Number';e={$_.PortNumber}},
                        @{n='Status';e={$_.HealthState}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Port Type';e={$_.PortType}},
                        @{n='Port Number';e={$_.PortNumber}},
                        @{n='Status';e={$_.HealthState}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_FanHealth' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 94
            'AllData' = @{}
            'Title' = 'HP Fan Health'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Removal Conditions';e={$_.RemovalConditions}},
                        @{n='Status';e={$_.HealthState}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Removal Conditions';e={$_.RemovalConditions}},
                        @{n='Status';e={$_.HealthState}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_HBAHealth' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 95
            'AllData' = @{}
            'Title' = 'HP HBA Health'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Manufacturer';e={$_.Manufacturer}},
                        @{n='Model';e={$_.Model}},
                        #@{n='Location';e={$_.OtherIdentifyingInfo}},
                        @{n='Status';e={$_.OperationalStatus}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Manufacturer';e={$_.Manufacturer}},
                        @{n='Model';e={$_.Model}},
                        #@{n='Location';e={$_.OtherIdentifyingInfo}},
                        @{n='Status';e={$_.OperationalStatus}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_PSUHealth' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 96
            'AllData' = @{}
            'Title' = 'HP Power Supply Health'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Type';e={$_.Type}},
                        @{n='Status';e={$_.HealthState}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Type';e={$_.Type}},
                        @{n='Status';e={$_.HealthState}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_TempSensors' = @{
            'Enabled' = $true
            'AllowEmptyReport' = $false
            'Order' = 97
            'AllData' = @{}
            'Title' = 'HP Temperature Sensors'
            'Type' = 'Section'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Description';e={$_.Description}},
                        @{n='Percent To Critical';e={$_.PercentToCritical}}
                }
            }
        }
    }
}
#endregion System Report Structure

#region System Report Static Variables
# Generally you don't futz with these, they are mostly just registry locations anyway
#WSUS Settings
$reg_WSUSSettings = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
# Product key (and other) settings
# for 32-bit: DigitalProductId
# for 64-bit: DigitalProductId4
$reg_ExtendedInfo = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$reg_NTPSettings = 'SYSTEM\CurrentControlSet\Services\W32Time\Parameters'
$ShareType = @{
    '0' = 'Disk Drive'
    '1' = 'Print Queue'
    '2' = 'Device'
    '3' = 'IPC'
    '2147483648' = 'Disk Drive Admin'
    '2147483649' = 'Print Queue Admin'
    '2147483650' = 'Device Admin'
    '2147483651' = 'IPC Admin'
}
#endregion System Report Static Variables

#region HTML Template Variables
# This is the meat and potatoes of how the reports are spit out. Currently it is
# broken down by html component -> rendering style.
$HTMLRendering = @{
    # Markers: 
    #   <0> - Server Name
    'Header' = @{
        'DynamicGrid' = @'
<!DOCTYPE html>
<!-- HTML5 Mobile Boilerplate -->
<!--[if IEMobile 7]><html class="no-js iem7"><![endif]-->
<!--[if (gt IEMobile 7)|!(IEMobile)]><!--><html class="no-js" lang="en"><!--<![endif]-->

<!-- HTML5 Boilerplate -->
<!--[if lt IE 7]><html class="no-js lt-ie9 lt-ie8 lt-ie7" lang="en"> <![endif]-->
<!--[if (IE 7)&!(IEMobile)]><html class="no-js lt-ie9 lt-ie8" lang="en"><![endif]-->
<!--[if (IE 8)&!(IEMobile)]><html class="no-js lt-ie9" lang="en"><![endif]-->
<!--[if gt IE 8]><!--> <html class="no-js" lang="en"><!--<![endif]-->

<head>

    <meta charset="utf-8">
    <!-- Always force latest IE rendering engine (even in intranet) & Chrome Frame -->
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <title>System Report</title>
    <meta http-equiv="cleartype" content="on">
    <link rel="shortcut icon" href="/favicon.ico">

    <!-- Responsive and mobile friendly stuff -->
    <meta name="HandheldFriendly" content="True">
    <meta name="MobileOptimized" content="320">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">

    <!-- Stylesheets 
    <link rel="stylesheet" href="css/html5reset.css" media="all">
    <link rel="stylesheet" href="css/responsivegridsystem.css" media="all">
    <link rel="stylesheet" href="css/col.css" media="all">
    <link rel="stylesheet" href="css/2cols.css" media="all">
    <link rel="stylesheet" href="css/3cols.css" media="all">
    -->
    <!--<link rel="stylesheet" href="AllStyles.css" media="all">-->
        <!-- Responsive Stylesheets 
    <link rel="stylesheet" media="only screen and (max-width: 1024px) and (min-width: 769px)" href="/css/1024.css">
    <link rel="stylesheet" media="only screen and (max-width: 768px) and (min-width: 481px)" href="/css/768.css">
    <link rel="stylesheet" media="only screen and (max-width: 480px)" href="/css/480.css">
    -->
    <!-- All JavaScript at the bottom, except for Modernizr which enables HTML5 elements and feature detects -->
    <!-- <script src="js/modernizr-2.5.3-min.js"></script> -->

    <style type="text/css">
    <!--
        /* html5reset.css - 01/11/2011 */
        html, body, div, span, object, iframe,
        h1, h2, h3, h4, h5, h6, p, blockquote, pre,
        abbr, address, cite, code,
        del, dfn, em, img, ins, kbd, q, samp,
        small, strong, sub, sup, var,
        b, i,
        dl, dt, dd, ol, ul, li,
        fieldset, form, label, legend,
        table, caption, tbody, tfoot, thead, tr, th, td,
        article, aside, canvas, details, figcaption, figure, 
        footer, header, hgroup, menu, nav, section, summary,
        time, mark, audio, video {
            margin: 0;
            padding: 0;
            border: 0;
            outline: 0;
            font-size: 100%;
            vertical-align: baseline;
            background: transparent;
        }

        body {
            line-height: 1;
        }

        article,aside,details,figcaption,figure,
        footer,header,hgroup,menu,nav,section { 
            display: block;
        }

        nav ul {
            list-style: none;
        }

        blockquote, q {
            quotes: none;
        }

        blockquote:before, blockquote:after,
        q:before, q:after {
            content: '';
            content: none;
        }

        a {
            margin: 0;
            padding: 0;
            font-size: 100%;
            vertical-align: baseline;
            background: transparent;
        }

        /* change colours to suit your needs */
        ins {
            background-color: #ff9;
            color: #000;
            text-decoration: none;
        }

        /* change colours to suit your needs */
        mark {
            background-color: #ff9;
            color: #000; 
            font-style: italic;
            font-weight: bold;
        }

        del {
            text-decoration:  line-through;
        }

        abbr[title], dfn[title] {
            border-bottom: 1px dotted;
            cursor: help;
        }

        table {
            border-collapse: collapse;
            border-spacing: 0;
        }

        /* change border colour to suit your needs */
        hr {
            display: block;
            height: 1px;
            border: 0;   
            border-top: 1px solid #cccccc;
            margin: 1em 0;
            padding: 0;
        }

        input, select {
            vertical-align: middle;
        }

        /* RESPONSIVE GRID SYSTEM =============================================================================  */
        /* BASIC PAGE SETUP ============================================================================= */
        body { 
        margin : 0 auto;
        padding : 0;
        font : 100%/1.4 'lucida sans unicode', 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif;     
        color : #000; 
        text-align: center;
        background: #fff url(/images/bodyback.png) left top;
        }

        button, 
        input, 
        select, 
        textarea { 
        font-family : MuseoSlab100, lucida sans unicode, 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif; 
        color : #333; }

        /*  HEADINGS  ============================================================================= */
        h1, h2, h3, h4, h5, h6 {
        font-family:  MuseoSlab300, 'lucida sans unicode', 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif;
        font-weight : normal;
        margin-top: 0px;
        letter-spacing: -1px;
        }

        h1 { 
        font-family:  LeagueGothicRegular, 'lucida sans unicode', 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif;
        color: #000;
        margin-bottom : 0.0em;
        font-size : 4em; /* 40 / 16 */
        line-height : 1.0;
        }

        h2 { 
        color: #222;
        margin-bottom : .5em;
        margin-top : .5em;
        font-size : 2.75em; /* 40 / 16 */
        line-height : 1.2;
        }

        h3 { 
        color: #333;
        margin-bottom : 0.3em;
        letter-spacing: -1px;
        font-size : 1.75em; /* 28 / 16 */
        line-height : 1.3; }

        h4 { 
        color: #444;
        margin-bottom : 0.5em;
        font-size : 1.5em; /* 24 / 16  */
        line-height : 1.25; }

            footer h4 { 
                color: #ccc;
            }

        h5 { 
        color: #555;
        margin-bottom : 1.25em;
        font-size : 1em; /* 20 / 16 */ }

        h6 { 
        color: #666;
        font-size : 1em; /* 16 / 16  */ }

        /*  TYPOGRAPHY  ============================================================================= */
        p, ol, ul, dl, address { 
        margin-bottom : 1.5em; 
        font-size : 1em; /* 16 / 16 = 1 */ }

        p {
        hyphens : auto;  }

        p.introtext {
        font-family:  MuseoSlab100, 'lucida sans unicode', 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif;
        font-size : 2.5em; /* 40 / 16 */
        color: #333;
        line-height: 1.4em;
        letter-spacing: -1px;
        margin-bottom: 0.5em;
        }

        p.handwritten {
        font-family:  HandSean, 'lucida sans unicode', 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif; 
        font-size: 1.375em; /* 24 / 16 */
        line-height: 1.8em;
        margin-bottom: 0.3em;
        color: #666;
        }

        p.center {
        text-align: center;
        }

        .and {
        font-family: GoudyBookletter1911Regular, Georgia, Times New Roman, sans-serif;
        font-size: 1.5em; /* 24 / 16 */
        }

        .heart {
        font-family: Pictos;
        font-size: 1.5em; /* 24 / 16 */
        }

        ul, 
        ol { 
        margin : 0 0 1.5em 0; 
        padding : 0 0 0 24px; }

        li ul, 
        li ol { 
        margin : 0;
        font-size : 1em; /* 16 / 16 = 1 */ }

        dl, 
        dd { 
        margin-bottom : 1.5em; }

        dt { 
        font-weight : normal; }

        b, strong { 
        font-weight : bold; }

        hr { 
        display : block; 
        margin : 1em 0; 
        padding : 0;
        height : 1px; 
        border : 0; 
        border-top : 1px solid #ccc;
        }

        small { 
        font-size : 1em; /* 16 / 16 = 1 */ }

        sub, sup { 
        font-size : 75%; 
        line-height : 0; 
        position : relative; 
        vertical-align : baseline; }

        sup { 
        top : -.5em; }

        sub { 
        bottom : -.25em; }

        .subtext {
            color: #666;
            }

        /* LINKS =============================================================================  */
        a { 
        color : #cc1122;
        -webkit-transition: all 0.3s ease;
        -moz-transition: all 0.3s ease;
        -o-transition: all 0.3s ease;
        transition: all 0.3s ease;
        text-decoration: none;
        }

        a:visited { 
        color : #ee3344; }

        a:focus { 
        outline : thin dotted; 
        color : rgb(0,0,0); }

        a:hover, 
        a:active { 
        outline : 0;
        color : #dd2233;
        }

        footer a { 
        color : #ffffff;
        -webkit-transition: all 0.3s ease;
        -moz-transition: all 0.3s ease;
        -o-transition: all 0.3s ease;
        transition: all 0.3s ease;
        }

        footer a:visited { 
        color : #fff; }

        footer a:focus { 
        outline : thin dotted; 
        color : rgb(0,0,0); }

        footer a:hover, 
        footer a:active { 
        outline : 0;
        color : #fff;
        }

        /* IMAGES ============================================================================= */

        img {
        border : 0;
        max-width: 100%;}

        img.floatleft { float: left; margin: 0 10px 0 0; }
        img.floatright { float: right; margin: 0 0 0 10px; }

        /* TABLES ============================================================================= */

        table { 
        border-collapse : collapse;
        border-spacing : 0;
        margin-bottom : 1.4em; 
        width : 100%; }

        th, td, caption { 
        padding : .25em 10px .25em 5px; }

        tfoot { 
        font-style : italic; }

        caption { 
        background-color : transparent; }

        /*  MAIN LAYOUT    ============================================================================= */
        #skiptomain { display: none; }

        #wrapper {
            width: 100%;
            position: relative;
            text-align: left;
        }

            #headcontainer {
                width: 100%;
            }

                header {
                    clear: both;
                    width: 80%; /* 1000px / 1250px */
                    font-size: 0.8125em; /* 13 / 16 */
                    max-width: 92.3em; /* 1200px / 13 */
                    margin: 0 auto;
                    padding: 30px 0px 10px 0px;
                    position: relative;
                    color: #000;
                    text-align: center;
                }

            #maincontentcontainer {
                width: 100%;
            }

                .standardcontainer {
                    
                }
                
                .darkcontainer {
                    background: rgba(102, 102, 102, 0.05);
                }

                .lightcontainer {
                    background: rgba(255, 255, 255, 0.25);
                }
                
                    #maincontent{
                        clear: both;
                        width: 80%; /* 1000px / 1250px */
                        font-size: 0.8125em; /* 13 / 16 */
                        max-width: 92.3em; /* 1200px / 13 */
                        margin: 0 auto;
                        padding: 1em 0px;
                        color: #333;
                        line-height: 1.5em;
                        position: relative;
                    }
                
                    .maincontent{
                        clear: both;
                        width: 80%; /* 1000px / 1250px */
                        font-size: 0.8125em; /* 13 / 16 */
                        max-width: 92.3em; /* 1200px / 13 */
                        margin: 0 auto;
                        padding: 1em 0px;
                        color: #333;
                        line-height: 1.5em;
                        position: relative;
                    }

            #footercontainer {
                width: 100%;    
                border-top: 1px solid #000;
                background: #222 url(/images/footerback.png) left top;
            }
            
                footer {
                    clear: both;
                    width: 80%; /* 1000px / 1250px */
                    font-size: 0.8125em; /* 13 / 16 */
                    max-width: 92.3em; /* 1200px / 13 */
                    margin: 0 auto;
                    padding: 20px 0px 10px 0px;
                    color: #999;
                }

                footer strong {
                    font-size: 1.077em; /* 14 / 13 */
                    color: #aaa;
                }

                footer a:link, footer a:visited { color: #999; text-decoration: underline; }
                footer a:hover { color: #fff; text-decoration: underline; }

                ul.pagefooterlist, ul.pagefooterlistimages {
                    display: block;
                    float: left;
                    margin: 0px;
                    padding: 0px;
                    list-style: none;
                }

                ul.pagefooterlist li, ul.pagefooterlistimages li {
                    clear: left;
                    margin: 0px;
                    padding: 0px 0px 3px 0px;
                    display: block;
                    line-height: 1.5em;
                    font-weight: normal;
                    background: none;

                }

                ul.pagefooterlistimages li {
                    height: 34px;
                }

                ul.pagefooterlistimages li img {
                    padding: 5px 5px 5px 0px;
                    vertical-align: middle;
                    opacity: 0.75;
                    -ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=75)";
                    filter: alpha( opacity  = 75);
                    -webkit-transition: all 0.3s ease;
                    -moz-transition: all 0.3s ease;
                    -o-transition: all 0.3s ease;
                    transition: all 0.3s ease;
                }

                ul.pagefooterlistimages li a
                {
                    text-decoration: none;
                }

                ul.pagefooterlistimages li a:hover img {
                    opacity: 1.0;
                    -ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=100)";
                    filter: alpha( opacity  = 100);
                }

                    #smallprint {
                        margin-top: 20px;
                        line-height: 1.4em;
                        text-align: center;
                        color: #999;
                        font-size: 0.923em; /* 12 / 13 */
                    }

                    #smallprint p{
                        vertical-align: middle;
                    }

                    #smallprint .twitter-follow-button{
                        margin-left: 1em;
                        vertical-align: middle;
                    }

                    #smallprint img {
                        margin: 0px 10px 15px 0px;
                        vertical-align: middle;
                        opacity: 0.5;
                        -ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=50)";
                        filter: alpha( opacity  = 50);
                        -webkit-transition: all 0.3s ease;
                        -moz-transition: all 0.3s ease;
                        -o-transition: all 0.3s ease;
                        transition: all 0.3s ease;
                    }

                    #smallprint a:hover img {
                        opacity: 1.0;
                        -ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=100)";
                        filter: alpha( opacity  = 100);
                    }

                    #smallprint a:link, #smallprint a:visited { color: #999; text-decoration: none; }
                    #smallprint a:hover { color: #999; text-decoration: underline; }

        /*  SECTIONS  ============================================================================= */
        .section {
            clear: both;
            padding: 0px;
            margin: 0px;
        }

        /*  CODE  ============================================================================= */
        pre.code {
            padding: 0;
            margin: 0;
            font-family: monospace;
            white-space: pre-wrap;
            font-size: 1.1em;
        }

        strong.code {
            font-weight: normal;
            font-family: monospace;
            font-size: 1.2em;
        }

        /*  EXAMPLE  ============================================================================= */
        #example .col {
            background: #ccc;
            background: rgba(204, 204, 204, 0.85);

        }

        /*  NOTES  ============================================================================= */
        .note {
            position:relative;
            padding:1em 1.5em;
            margin: 0 0 1em 0;
            background: #fff;
            background: rgba(255, 255, 255, 0.5);
            overflow:hidden;
        }

        .note:before {
            content:"";
            position:absolute;
            top:0;
            right:0;
            border-width:0 16px 16px 0;
            border-style:solid;
            border-color:transparent transparent #cccccc #cccccc;
            background:#cccccc;
            -webkit-box-shadow:0 1px 1px rgba(0,0,0,0.3), -1px 1px 1px rgba(0,0,0,0.2);
            -moz-box-shadow:0 1px 1px rgba(0,0,0,0.3), -1px 1px 1px rgba(0,0,0,0.2);
            box-shadow:0 1px 1px rgba(0,0,0,0.3), -1px 1px 1px rgba(0,0,0,0.2);
            display:block; width:0; /* Firefox 3.0 damage limitation */
        }

        .note.rounded {
            -webkit-border-radius:5px 0 5px 5px;
            -moz-border-radius:5px 0 5px 5px;
            border-radius:5px 0 5px 5px;
        }

        .note.rounded:before {
            border-width:8px;
            border-color:#ff #ff transparent transparent;
            background: url(/images/bodyback.png);
            -webkit-border-bottom-left-radius:5px;
            -moz-border-radius:0 0 0 5px;
            border-radius:0 0 0 5px;
        }

        /*  SCREENS  ============================================================================= */
        .siteimage {
            max-width: 90%;
            padding: 5%;
            margin: 0 0 1em 0;
            background: transparent url(/images/stripe-bg.png);
            -webkit-transition: background 0.3s ease;
            -moz-transition: background 0.3s ease;
            -o-transition: background 0.3s ease;
            transition: background 0.3s ease;
        }

        .siteimage:hover {
            background: #bbb url(/images/stripe-bg.png);
            position: relative;
            top: -2px;
            
        }

        /*  COLUMNS  ============================================================================= */
        .twocolumns{
            -moz-column-count: 2;
            -moz-column-gap: 2em;
            -webkit-column-count: 2;
            -webkit-column-gap: 2em;
            column-count: 2;
            column-gap: 2em;
          }

        /*  GLOBAL OBJECTS ============================================================================= */
        .breaker { clear: both; }

        .group:before,
        .group:after {
            content:"";
            display:table;
        }
        .group:after {
            clear:both;
        }
        .group {
            zoom:1; /* For IE 6/7 (trigger hasLayout) */
        }

        .floatleft {
            float: left;
        }

        .floatright {
            float: right;
        }

        /* VENDOR-SPECIFIC ============================================================================= */
        html { 
        -webkit-overflow-scrolling : touch; 
        -webkit-tap-highlight-color : rgb(52,158,219); 
        -webkit-text-size-adjust : 100%; 
        -ms-text-size-adjust : 100%; }

        .clearfix { 
        zoom : 1; }

        ::-webkit-selection { 
        background : rgb(23,119,175); 
        color : rgb(250,250,250); 
        text-shadow : none; }

        ::-moz-selection { 
        background : rgb(23,119,175); 
        color : rgb(250,250,250); 
        text-shadow : none; }

        ::selection { 
        background : rgb(23,119,175); 
        color : rgb(250,250,250); 
        text-shadow : none; }

        button, 
        input[type="button"], 
        input[type="reset"], 
        input[type="submit"] { 
        -webkit-appearance : button; }

        ::-webkit-input-placeholder {
        font-size : .875em; 
        line-height : 1.4; }

        input:-moz-placeholder { 
        font-size : .875em; 
        line-height : 1.4; }

        .ie7 img,
        .iem7 img { 
        -ms-interpolation-mode : bicubic; }

        input[type="checkbox"], 
        input[type="radio"] { 
        box-sizing : border-box; }

        input[type="search"] { 
        -webkit-box-sizing : content-box;
        -moz-box-sizing : content-box; }

        button::-moz-focus-inner, 
        input::-moz-focus-inner { 
        padding : 0;
        border : 0; }

        p {
        /* http://www.w3.org/TR/css3-text/#hyphenation */
        -webkit-hyphens : auto;
        -webkit-hyphenate-character : "\2010";
        -webkit-hyphenate-limit-after : 1;
        -webkit-hyphenate-limit-before : 3;
        -moz-hyphens : auto; }

        /*  SECTIONS  ============================================================================= */
        .section {
            clear: both;
            padding: 0px;
            margin: 0px;
        }

        /*  GROUPING  ============================================================================= */
        .group:before,
        .group:after {
            content:"";
            display:table;
        }
        .group:after {
            clear:both;
        }
        .group {
            zoom:1; /* For IE 6/7 (trigger hasLayout) */
        }

        /*  GRID COLUMN SETUP   ==================================================================== */
        .col {
            display: block;
            float:left;
            margin: 1% 0 1% 1.6%;
        }

        .col:first-child { margin-left: 0; } /* all browsers except IE6 and lower */

        /*  REMOVE MARGINS AS ALL GO FULL WIDTH AT 480 PIXELS */
        @media only screen and (max-width: 480px) {
            .col { 
                margin: 1% 0 1% 0%;
            }
        }

        /*  GRID OF TWO   ============================================================================= */
        .span_2_of_2 {
            width: 100%;
        }

        .span_1_of_2 {
            width: 49.2%;
        }

        /*  GO FULL WIDTH AT LESS THAN 480 PIXELS */
        @media only screen and (max-width: 480px) {
            .span_2_of_2 {
                width: 100%; 
            }
            .span_1_of_2 {
                width: 100%; 
            }
        }
        /*  GRID OF THREE   ============================================================================= */
        .span_3_of_3 {
            width: 100%; 
        }

        .span_2_of_3 {
            width: 66.1%; 
        }

        .span_1_of_3 {
            width: 32.2%; 
        }
        /*  GO FULL WIDTH AT LESS THAN 480 PIXELS */
        @media only screen and (max-width: 480px) {
            .span_3_of_3 {
                width: 100%; 
            }
            .span_2_of_3 {
                width: 100%; 
            }
            .span_1_of_3 {
                width: 100%;
            }
        }
        
        /*  GRID OF FOUR   ============================================================================= */

            
        .span_4_of_4 {
            width: 100%; 
        }

        .span_3_of_4 {
            width: 74.6%; 
        }

        .span_2_of_4 {
            width: 49.2%; 
        }

        .span_1_of_4 {
            width: 23.8%; 
        }


        /*  GO FULL WIDTH AT LESS THAN 480 PIXELS */

        @media only screen and (max-width: 480px) {
            .span_4_of_4 {
                width: 100%; 
            }
            .span_3_of_4 {
                width: 100%; 
            }
            .span_2_of_4 {
                width: 100%; 
            }
            .span_1_of_4 {
                width: 100%; 
            }
        }
        
        body {
            font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
        }
        
        table{
            border-collapse: collapse;
            border: none;
            font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
            color: black;
            margin-bottom: 10px;
        }
        table td{
            font-size: 10px;
            padding-left: 0px;
            padding-right: 20px;
            text-align: left;
        }
        table td:last-child{
            padding-right: 5px;
        }
        table th {
            font-size: 12px;
            font-weight: bold;
            padding-left: 0px;
            padding-right: 20px;
            text-align: left;
            border-bottom: 1px  grey solid;
        }
        h2{ 
            clear: both;
            font-size: 130%; 
            margin-left: 20px;
        }
        h3{
            clear: both;
            font-size: 115%;
            margin-left: 20px;
            margin-top: 30px;
        }
        p{ 
            margin-left: 20px; font-size: 12px;
        }
        table.list{
            float: left;
        }
        table.list td:nth-child(1){
            font-weight: bold;
            border-right: 1px grey solid;
            text-align: right;
        }
        table.list td:nth-child(2){
            padding-left: 7px;
        }
        table tr:nth-child(even) td:nth-child(even){ background: #CCCCCC; }
        table tr:nth-child(odd) td:nth-child(odd){ background: #F2F2F2; }
        table tr:nth-child(even) td:nth-child(odd){ background: #DDDDDD; }
        table tr:nth-child(odd) td:nth-child(even){ background: #E5E5E5; }
        
        /*  Error and warning highlighting - Row*/
        table tr.warn:nth-child(even) td:nth-child(even){ background: #FFFF88; }
        table tr.warn:nth-child(odd) td:nth-child(odd){ background: #FFFFBB; }
        table tr.warn:nth-child(even) td:nth-child(odd){ background: #FFFFAA; }
        table tr.warn:nth-child(odd) td:nth-child(even){ background: #FFFF99; }
        
        table tr.alert:nth-child(even) td:nth-child(even){ background: #FF8888; }
        table tr.alert:nth-child(odd) td:nth-child(odd){ background: #FFBBBB; }
        table tr.alert:nth-child(even) td:nth-child(odd){ background: #FFAAAA; }
        table tr.alert:nth-child(odd) td:nth-child(even){ background: #FF9999; }
        
        table tr.healthy:nth-child(even) td:nth-child(even){ background: #88FF88; }
        table tr.healthy:nth-child(odd) td:nth-child(odd){ background: #BBFFBB; }
        table tr.healthy:nth-child(even) td:nth-child(odd){ background: #AAFFAA; }
        table tr.healthy:nth-child(odd) td:nth-child(even){ background: #99FF99; }
        
        /*  Error and warning highlighting - Cell*/
        table tr:nth-child(even) td.warn:nth-child(even){ background: #FFFF88; }
        table tr:nth-child(odd) td.warn:nth-child(odd){ background: #FFFFBB; }
        table tr:nth-child(even) td.warn:nth-child(odd){ background: #FFFFAA; }
        table tr:nth-child(odd) td.warn:nth-child(even){ background: #FFFF99; }
        
        table tr:nth-child(even) td.alert:nth-child(even){ background: #FF8888; }
        table tr:nth-child(odd) td.alert:nth-child(odd){ background: #FFBBBB; }
        table tr:nth-child(even) td.alert:nth-child(odd){ background: #FFAAAA; }
        table tr:nth-child(odd) td.alert:nth-child(even){ background: #FF9999; }
        
        table tr:nth-child(even) td.healthy:nth-child(even){ background: #88FF88; }
        table tr:nth-child(odd) td.healthy:nth-child(odd){ background: #BBFFBB; }
        table tr:nth-child(even) td.healthy:nth-child(odd){ background: #AAFFAA; }
        table tr:nth-child(odd) td.healthy:nth-child(even){ background: #99FF99; }
        
        /* security highlighting */
        table tr.security:nth-child(even) td:nth-child(even){ 
            border-color: #FF1111; 
            border: 1px #FF1111 solid;
        }
        table tr.security:nth-child(odd) td:nth-child(odd){ 
            border-color: #FF1111; 
            border: 1px #FF1111 solid;
        }
        table tr.security:nth-child(even) td:nth-child(odd){
            border-color: #FF1111; 
            border: 1px #FF1111 solid;
        }

        table tr.security:nth-child(odd) td:nth-child(even){
            border-color: #FF1111; 
            border: 1px #FF1111 solid;
        }

        
        table th.title{ 
            text-align: center;
            background: #848482;
            border-bottom: 1px  grey solid;
            font-weight: bold;
            color: white;
        }
        
        table th.sectionbreak{ 
            text-align: center;
            background: #848482;
            border-bottom: 1px  grey solid;
            font-weight: bold;
            color: white;
            font-size: 130%; 
        }
        table tr.divide{
            border-bottom: 1px  grey solid;
        }

    -->
    </style>
</head>

<body>
<hr noshade size="3" width='100%'>
<div id="wrapper">

'@
        'EmailFriendly' = @'
<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Frameset//EN' 'http://www.w3.org/TR/html4/frameset.dtd'>
<html><head><title><0></title>
<style type='text/css'>
<!--
body {
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
table{
   border-collapse: collapse;
   border: none;
   font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
   color: black;
   margin-bottom: 10px;
   margin-left: 20px;
}
table td{
   font-size: 12px;
   padding-left: 0px;
   padding-right: 20px;
   text-align: left;
   border:1px solid black;
}
table th {
   font-size: 12px;
   font-weight: bold;
   padding-left: 0px;
   padding-right: 20px;
   text-align: left;
}

h1{ clear: both;
    font-size: 150%; 
    text-align: center;
  }
h2{ clear: both; font-size: 130%; }

h3{
   clear: both;
   font-size: 115%;
   margin-left: 20px;
   margin-top: 30px;
}

p{ margin-left: 20px; font-size: 12px; }

table.list{ float: left; }
   table.list td:nth-child(1){
   font-weight: bold;
   border: 1px grey solid;
   text-align: right;
}

table th.title{ 
    text-align: center;
    background: #848482;
    border: 2px  grey solid;
    font-weight: bold;
    color: white;
}
table tr.divide{
    border-bottom: 5px  grey solid;
}
.odd { background-color:#ffffff; }
.even { background-color:#dddddd; }
.warn { background-color:yellow; }
.alert { background-color:red; }
-->
</style>
</head>
<body>
'@
    }
    'Footer' = @{
        'DynamicGrid' = @'
</div>
</body>
</html>        
'@
        'EmailFriendly' = @'
</div>
</body>
</html>       
'@
    }

    # Markers: 
    #   <0> - Server Name
    'ServerBegin' = @{
        'DynamicGrid' = @'

    <div id="headcontainer">
        <header>
        <h1><0></h1>
        </header>
    </div>
    <div id="maincontentcontainer">
        <div id="maincontent">
            <div class="section group">
                <hr noshade size="3" width='100%'>
            </div>
            <div>

       
'@
        'EmailFriendly' = @'
    <div id='report'>
    <hr noshade size=3 width='100%'>
    <h1><0></h1>

    <div id="maincontentcontainer">
    <div id="maincontent">
      <div class="section group">
        <hr noshade="noshade" size="3" width="100%" style=
        "display:block;height:1px;border:0;border-top:1px solid #ccc;margin:1em 0;padding:0;" />
      </div>
      <div>

'@    
    }
    'ServerEnd' = @{
        'DynamicGrid' = @'

            </div>
        </div>
    </div>
</div>

'@
        'EmailFriendly' = @'

            </div>
        </div>
    </div>
</div>

'@
    }
    
    # Markers: 
    #   <0> - columns to span title
    #   <1> - Table header title
    'TableTitle' = @{
        'DynamicGrid' = @'
        
            <tr>
                <th class="title" colspan=<0>><1></th>
            </tr>
'@
        'EmailFriendly' = @'
            
            <tr>
              <th class="title" colspan="<0>"><1></th>
            </tr>
              
'@
    }

    'SectionContainers' = @{
        'DynamicGrid'  = @{
            'Half' = @{
                'Head' = @'
        
        <div class="col span_2_of_4">
'@
                'Tail' = @'
        </div>
'@
            }
            'Full' = @{
                'Head' = @'
        
        <div class="col span_4_of_4">
'@
                'Tail' = @'
        </div>
'@
            }
            'Third' = @{
                'Head' = @'
        
        <div class="col span_1_of_3">
'@
                'Tail' = @'
        </div>
'@
            }
            'TwoThirds' = @{
                'Head' = @'
        
        <div class="col span_2_of_3">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'Fourth'        = @{
                'Head' = @'
        
        <div class="col span_1_of_4">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'ThreeFourths'  = @{
                'Head' = @'
               
        <div class="col span_3_of_4">
'@
                'Tail'          = @'
        
        </div>
'@
            }
        }
        'EmailFriendly'  = @{
            'Half' = @{
                'Head' = @'
        
        <div class="col span_2_of_4">
        <table><tr WIDTH="50%">
'@
                'Tail' = @'
        </tr></table>       
        </div>
'@
            }
            'Full' = @{
                'Head' = @'
        
        <div class="col span_4_of_4">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'Third' = @{
                'Head' = @'
        
        <div class="col span_1_of_3">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'TwoThirds' = @{
                'Head' = @'
        
        <div class="col span_2_of_3">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'Fourth'        = @{
                'Head' = @'
        
        <div class="col span_1_of_4">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'ThreeFourths'  = @{
                'Head' = @'
               
        <div class="col span_3_of_4">
'@
                'Tail'          = @'
        
        </div>
'@
            }
        }
    }
    
    'SectionContainerGroup' = @{
        'DynamicGrid' = @{ 
            'Head' = @'
        
        <div class="section group">
'@
            'Tail' = @'
        </div>
'@
        }
        'EmailFriendly' = @{
            'Head' = @'
    
        <div class="section group">
'@
            'Tail' = @'
        </div>
'@
        }
    }
    
    'CustomSections' = @{
        # Markers: 
        #   <0> - Header
        'SectionBreak' = @'
    
    <div class="section group">        
        <div class="col span_4_of_4"><table>        
            <tr>
                <th class="sectionbreak"><0></th>
            </tr>
        </table>
        </div>
    </div>
'@
    }
}
#endregion HTML Template Variables
#endregion Globals

#region Functions
#region Functions - Multiple Runspace
Function Get-RemoteRouteTable
{
    <#
    .SYNOPSIS
       Gathers remote system route entries.
    .DESCRIPTION
       Gathers remote system route entries, including persistent routes. Utilizes multiple runspaces and
       alternate credentials if desired.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information

    .EXAMPLE
       PS > Get-RemoteRouteTable

       <output>
       
       Description
       -----------
       <Placeholder>

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 08/31/2013
        - Initial release
    #>
    [CmdletBinding()]
    PARAM
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
       
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    BEGIN
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Route Table: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Remote Route Table: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Route Table: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Route Table: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Route Table: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Route Table: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Remote Route Table: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                # General variables
                $ResultSet = @()
                $PSDateTime = Get-Date
                $RouteType = @('Unknown','Other','Invalid','Direct','Indirect')
                $Routes = @()
                
                #region ShareSessions
                Write-Verbose -Message ('Remote Route Table: Runspace {0}: Route table information' -f $ComputerName)

                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','Routes')
                                         
                # WMI data
                $wmi_routes = Get-WmiObject @WMIHast -Class win32_ip4RouteTable
                $wmi_persistedroutes = Get-WmiObject @WMIHast -Class win32_IP4PersistedRouteTable
                foreach ($iproute in $wmi_routes)
                {
                    $Persistant = $false
                    foreach ($piproute in $wmi_persistedroutes)
                    {
                        if (($iproute.Destination -eq $piproute.Destination) -and
                            ($iproute.Mask -eq $piproute.Mask) -and
                            ($iproute.NextHop -eq $piproute.NextHop))
                        {
                            $Persistant = $true
                        }
                    }
                    $RouteProperty = @{
                        'InterfaceIndex' = $iproute.InterfaceIndex
                        'Destination' = $iproute.Destination
                        'Mask' = $iproute.Mask
                        'NextHop' = $iproute.NextHop
                        'Metric' = $iproute.Metric1
                        'Persistent' = $Persistant
                        'Type' = $RouteType[[int]$iproute.Type]
                    }
                    $Routes += New-Object -TypeName PSObject -Property $RouteProperty
                }
                # Setup the default properties for output
                $ResultObject = New-Object PSObject -Property @{
                                                                'PSComputerName' = $ComputerName
                                                                'ComputerName' = $ComputerName
                                                                'PSDateTime' = $PSDateTime
                                                                'Routes' = $Routes
                                                               }
                $ResultObject.PSObject.TypeNames.Insert(0,'My.RouteTable.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                #endregion Routes

                Write-Output -InputObject $ResultObject
            }
            catch
            {
                Write-Warning -Message ('Remote Route Table: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Route Table: Runspace {0}: End' -f $ComputerName)
        }
 
        Function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Route Table: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Route Table: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Remote Route Table: Getting info'
                        Status = 'Remote Route Table: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Route Table: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
     END
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Route Table: Getting route table information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Route Table: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-RemoteGroupMembership
{
    <#
    .SYNOPSIS
       Gather list of all assigned users in all local groups on a computer.  
    .DESCRIPTION
       Gather list of all assigned users in all local groups on a computer.  
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER IncludeEmptyGroups
       Include local groups without any user membership.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information

    .EXAMPLE
       PS > (Get-RemoteGroupMembership -verbose).GroupMembership

       <output>
       
       Description
       -----------
       List all group membership of the local machine.

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 09/09/2013
        - Initial release
    #>
    [CmdletBinding()]
    PARAM
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter(HelpMessage="Include empty groups in results")]
        [switch]
        $IncludeEmptyGroups,
       
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    BEGIN
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Local Group Membership: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Local Group Membership: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Local Group Membership: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Local Group Membership: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Local Group Membership: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Local Group Membership: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID,
                
                [Parameter()]
                [switch]
                $IncludeEmptyGroups
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Local Group Membership: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                # General variables
                $GroupMembership = @()
                $PSDateTime = Get-Date
                
                #region Group Information
                Write-Verbose -Message ('Local Group Membership: Runspace {0}: Group memberhsip information' -f $ComputerName)

                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','GroupMembership')
                $wmi_groups = Get-WmiObject @WMIHast -Class win32_group -filter "Domain = '$ComputerName'"
                foreach ($group in $wmi_groups)
                {
                    $Query = "SELECT * FROM Win32_GroupUser WHERE GroupComponent = `"Win32_Group.Domain='$ComputerName',Name='$($group.name)'`""
                    $wmi_users = Get-WmiObject @WMIHast -query $Query
                    if (($wmi_users -eq $null) -and ($IncludeEmptyGroups))
                    {
                        $MembershipProperty = @{
                            'Group' = $group.Name
                            'GroupMember' = ''
                            'MemberType' = ''
                        }
                        $GroupMembership += New-Object PSObject -Property $MembershipProperty
                    }
                    else
                    {
                        foreach ($user in $wmi_users.partcomponent)
                        {
                            if ($user -match 'Win32_UserAccount')
                            {
                                $Type = 'User Account'
                            }
                            elseif ($user -match 'Win32_Group')
                            {
                                $Type = 'Group'
                            }
                            elseif ($user -match 'Win32_SystemAccount')
                            {
                                $Type = 'System Account'
                            }
                            else
                            {
                                $Type = 'Other'
                            }
                            $MembershipProperty = @{
                                'Group' = $group.Name
                                'GroupMember' = ($user.replace("Domain="," , ").replace(",Name=","\").replace("\\",",").replace('"','').split(","))[2]
                                'MemberType' = $Type
                            }
                            $GroupMembership += New-Object PSObject -Property $MembershipProperty
                        }
                    }
                }
                
                $ResultProperty = @{
                    'PSComputerName' = $ComputerName
                    'PSDateTime' = $PSDateTime
                    'ComputerName' = $ComputerName
                    'GroupMembership' = $GroupMembership
                }
                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty
                
                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.GroupMembership.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                Write-Output -InputObject $ResultObject
                #endregion Group Information
            }
            catch
            {
                Write-Warning -Message ('Local Group Membership: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Local Group Membership: Runspace {0}: End' -f $ComputerName)
        }
 
        Function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Local Group Membership: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Local Group Membership: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Local Group Membership: Getting info'
                        Status = 'Local Group Membership: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('IncludeEmptyGroups',$IncludeEmptyGroups)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Local Group Membership: Starting {0}' -f $Computer)
            
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
     END
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Local Group Membership: Getting local group information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Local Group Membership: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-ComputerAssetInformation
{
    <#
    .SYNOPSIS
       Get inventory data for specified computer systems.
    .DESCRIPTION
       Gather inventory data for one or more systems using wmi. Data proccessing utilizes multiple runspaces
       and supports custom timeout parameters in case of wmi problems. You can optionally include 
       drive, memory, and network information in the results. You can view verbose information on each 
       runspace thread in realtime with the -Verbose option.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER IncludeMemoryInfo
       Include information about the memory arrays and the installed memory within them. (_Memory and _MemoryArray)
    .PARAMETER IncludeDiskInfo
       Include disk partition and mount point information. (_Disks)
    .PARAMETER IncludeNetworkInfo
       Include general network configuration for enabled interfaces. (_Network)
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information
    .PARAMETER PromptForCredential
       Prompt for remote system credential prior to processing request.
    .PARAMETER Credential
       Accept alternate credential (ignored if the localhost is processed)
    .EXAMPLE
       PS > Get-ComputerAssetInformation -ComputerName test1
     
            ComputerName        : TEST1
            IsVirtual           : False
            Model               : ProLiant DL380 G7
            ChassisModel        : Rack Mount Unit
            OperatingSystem     : Microsoft Windows Server 2008 R2 Enterprise 
            OSServicePack       : 1
            OSVersion           : 6.1.7601
            OSSKU               : Enterprise Server Edition
            OSArchitecture      : x64
            SystemArchitecture  : x64
            PhysicalMemoryTotal : 12.0 GB
            PhysicalMemoryFree  : 621.7 MB
            VirtualMemoryTotal  : 24.0 GB
            VirtualMemoryFree   : 5.7 GB
            CPUCores            : 24
            CPUSockets          : 2
            SystemTime          : 08/04/2013 20:33:47
            LastBootTime        : 07/16/2013 07:42:01
            InstallDate         : 07/02/2011 17:52:34
            Uptime              : 19 days 12 hours 51 minutes
     
       Description
       -----------
       Query and display basic information ablout computer test1

    .EXAMPLE
       PS > $cred = Get-Credential
       PS > $b = Get-ComputerAssetInformation -ComputerName Test1 -Credential $cred -IncludeMemoryInfo 
       PS > $b | Select MemorySlotsTotal,MemorySlotsUsed | fl

       MemorySlotsTotal : 18
       MemorySlotsUsed  : 6
       
       PS > $b._Memory | Select DeviceLocator,@{n='MemorySize'; e={$_.Capacity/1Gb}}

       DeviceLocator                                                                   MemorySize
       -------------                                                                   ----------
       PROC 1 DIMM 3A                                                                           2
       PROC 1 DIMM 6B                                                                           2
       PROC 1 DIMM 9C                                                                           2
       PROC 2 DIMM 3A                                                                           2
       PROC 2 DIMM 6B                                                                           2
       PROC 2 DIMM 9C                                                                           2
       
       Description
       -----------
       Query information about computer test1 using alternate credentials, including detailed memory information. Return
       physical memory slots available and in use. Then display the memory location and size.
    .EXAMPLE
        PS > $a = Get-ComputerAssetInformation -IncludeDiskInfo -IncludeMemoryInfo -IncludeNetworkInfo
        PS > $a._MemorySlots | ft

        Label      Bank        Detail           FormFactor      Capacity
        -----      ----        ------           ----------      --------
        BANK 0     Bank 1      Synchronous      SODIMM              4096
        BANK 2     Bank 2      Synchronous      SODIMM              4096
        
       Description
       -----------
       Query local computer for all information, store the results in $a, then show the memory slot utilization in further
       detail in tabular form.
    .EXAMPLE
        PS > (Get-ComputerAssetInformation -IncludeDiskInfo)._Disks

        Drive            : C:
        DiskType         : Partition
        Description      : Installable File System
        VolumeName       : 
        PercentageFree   : 10.64
        Disk             : \\.\PHYSICALDRIVE0
        SerialNumber     :       0SGDENZA091227
        FreeSpace        : 6.3 GB
        PrimaryPartition : True
        DiskSize         : 59.6 GB
        Model            : SAMSUNG SSD PM800 TH 64G
        Partition        : Disk #0, Partition #0

        Description
        -----------
        Query information about computer, include disk information, and immediately display it.

    .NOTES
       Originally posted at: http://learn-powershell.net/2013/05/08/scripting-games-2013-event-2-favorite-and-not-so-favorite/
       Author: Zachary Loeber
       Props To: David Lee (www.linkedin.com/pub/david-lee/2/686/482/) - Helped to troubleshoot and resolve numerous aspects of this script
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0
       Info: WMI prefered over CIM as there no speed advantage using cimsessions in multithreading against old systems. Starting
             around line 263 you can modify the WMI_<Property>Props arrays to include extra wmi data for each element should you
             require information I may have missed. You can also change the default display properties by modifying $defaultProperties.
             Keep in mind that including extra elements like the drive space and network information will increase the processing time per
             system. You may need to increase the timeout parameter accordingly.
       
       Version History
       1.2.1 - 9/09/2013
        - Fixed a regression bug for os and system architecture detection (based on processor addresslength and width)       
       1.2.0 - 9/05/2013
        - Got rid of the embedded add-member scriptblock for converting to kb/mb/gb format in
          favor of a filter. This results in string object proeprties being returned instead
          of uint64 which can cause bizzare issues with excel when importing data.
       1.1.8 - 8/30/2013
        - Included the system serial number in the default general results
       1.1.7 - 8/29/2013
        - Fixed incorrect installdate in general information
        - Prefixed all warnings and verbose messages with function specific verbage
        - Forced STA apartement state before opening a runspace
        - Added memory speed to Memory section
       1.1.6 - 8/19/2013
        - Refactored the date/time calculations to be less region specific.
        - Added PercentPhysicalMemoryUsed to general info section
       1.1.5 - 8/16/2013
        - Fixed minor powershell 2.0 compatibility issue with empty array detection in the mountpoint calculation area.
       1.1.4 - 8/15/2013
        - Fixed cpu architecture determination logic (again).
        - Included _MemorySlots in the memory results option. This includes an array of objects describing the memory
          array, which slots are utilized, and what type of ram is utilizing them.
        - Added RAM lookup tables for memory model and  details.
       1.1.3 - 8/13/2013
        - Fixed improper variable assignment for virtual platform detection
        - Changed network connection results to simply include all adapters, connected or not and include a new derived property called 
          'ConnectionStatus'. This fixes a backwards compatibility issue with pre-2008 servers  and network detection.
        - Added nic promiscuous mode detection for adapters. 
            (http://praetorianprefect.com/archives/2009/09/whos-being-promiscuous-in-your-active-directory/)
       1.1.2 - 8/12/2013
        - Fixed a backward compatibility bug with SystemArchitecture and OSArchitecture properties
        - Added the actual network adapter display name to the _Network results.
        - Added another example in the comment based help   
       1.1.1 - 8/7/2013
        - Added wmi BIOS information to results (as _BIOS)
        - Added IsVirtual and VirtualType to default result properties
       1.1.0 - 8/3/2013
        - Added several parameters
        - Removed parameter sets in favor of arrays of custom object as note properties
        - Removed ICMP response requirements
        - Included more verbose runspace logging    
       1.0.2 - 8/2/2013
        - Split out system and OS architecture (changing how each it determined)
       1.0.1 - 8/1/2013
        - Updated to include several more bits of information and customization variables
       1.0.0 - ???
        - Discovered original script on the internet and totally was blown away at how awesome it is.
    #>
    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter()]
        [switch]
        $IncludeMemoryInfo,
 
        [Parameter()]
        [switch]
        $IncludeDiskInfo,
 
        [Parameter()]
        [switch]
        $IncludeNetworkInfo,
       
        [Parameter()]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter()]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter()]
        [switch]
        $ShowProgress,
        
        [Parameter()]
        [switch]
        $PromptForCredential,
        
        [Parameter()]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    BEGIN
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Asset Information: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Remote Asset Information: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Asset Information: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Asset Information: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Asset Information: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Asset Information: Defining background runspaces scriptblock'
        $ScriptBlock = 
        {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID,
          
                [Parameter()]
                [switch]
                $IncludeMemoryInfo,
         
                [Parameter()]
                [switch]
                $IncludeDiskInfo,
         
                [Parameter()]
                [switch]
                $IncludeNetworkInfo
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Remote Asset Information: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                Filter ConvertTo-KMG 
                {
                     <#
                     .Synopsis
                      Converts byte counts to Byte\KB\MB\GB\TB\PB format
                     .DESCRIPTION
                      Accepts an [int64] byte count, and converts to Byte\KB\MB\GB\TB\PB format
                      with decimal precision of 2
                     .EXAMPLE
                     3000 | convertto-kmg
                     #>

                     $bytecount = $_
                        switch ([math]::truncate([math]::log($bytecount,1024))) 
                        {
                            0 {"$bytecount Bytes"}
                            1 {"{0:n2} KB" -f ($bytecount / 1kb)}
                            2 {"{0:n2} MB" -f ($bytecount / 1mb)}
                            3 {"{0:n2} GB" -f ($bytecount / 1gb)}
                            4 {"{0:n2} TB" -f ($bytecount / 1tb)}
                            Default {"{0:n2} PB" -f ($bytecount / 1pb)}
                        }
                }

                #region GeneralInfo
                Write-Verbose -Message ('Remote Asset Information: Runspace {0}: General asset information' -f $ComputerName)
                ## Lookup arrays
                $SKUs   = @("Undefined","Ultimate Edition","Home Basic Edition","Home Basic Premium Edition","Enterprise Edition",`
                            "Home Basic N Edition","Business Edition","Standard Server Edition","DatacenterServer Edition","Small Business Server Edition",`
                            "Enterprise Server Edition","Starter Edition","Datacenter Server Core Edition","Standard Server Core Edition",`
                            "Enterprise ServerCoreEdition","Enterprise Server Edition for Itanium-Based Systems","Business N Edition","Web Server Edition",`
                            "Cluster Server Edition","Home Server Edition","Storage Express Server Edition","Storage Standard Server Edition",`
                            "Storage Workgroup Server Edition","Storage Enterprise Server Edition","Server For Small Business Edition","Small Business Server Premium Edition")
                $ChassisModels = @("PlaceHolder","Maybe Virtual Machine","Unknown","Desktop","Thin Desktop","Pizza Box","Mini Tower","Full Tower","Portable",`
                                   "Laptop","Notebook","Hand Held","Docking Station","All in One","Sub Notebook","Space-Saving","Lunch Box","Main System Chassis",`
                                   "Lunch Box","SubChassis","Bus Expansion Chassis","Peripheral Chassis","Storage Chassis" ,"Rack Mount Unit","Sealed-Case PC")
                $NetConnectionStatus = @('Disconnected','Connecting','Connected','Disconnecting','Hardware not present','Hardware disabled','Hardware malfunction',`
                                         'Media disconnected','Authenticating','Authentication succeeded','Authentication failed','Invalid address','Credentials required')
                $MemoryModels = @("Unknown","Other","SIP","DIP","ZIP","SOJ","Proprietary","SIMM","DIMM","TSOP","PGA","RIMM",`
                                  "SODIMM","SRIMM","SMD","SSMP","QFP","TQFP","SOIC","LCC","PLCC","BGA","FPBGA","LGA")
                $MemoryDetail = @{
                    '1' = 'Reserved'
                    '2' = 'Other'
                    '4' = 'Unknown'
                    '8' = 'Fast-paged'
                    '16' = 'Static column'
                    '32' = 'Pseudo-static'
                    '64' = 'RAMBUS'
                    '128' = 'Synchronous'
                    '256' = 'CMOS'
                    '512' = 'EDO'
                    '1024' = 'Window DRAM'
                    '2048' = 'Cache DRAM'
                    '4096' = 'Nonvolatile'
                }

                # Modify this variable to change your default set of display properties
                $defaultProperties = @('ComputerName','IsVirtual','Model','ChassisModel','SerialNumber','OperatingSystem','OSServicePack','OSVersion','OSSKU', `
                                       'OSArchitecture','SystemArchitecture','PhysicalMemoryTotal','PhysicalMemoryFree','VirtualMemoryTotal', `
                                       'VirtualMemoryFree','CPUCores','CPUSockets','SystemTime','LastBootTime','InstallDate','Uptime')
                # WMI Properties
                $WMI_OSProps = @('BuildNumber','Version','SerialNumber','ServicePackMajorVersion','CSDVersion','SystemDrive',`
                                 'SystemDirectory','WindowsDirectory','Caption','TotalVisibleMemorySize','FreePhysicalMemory',`
                                 'TotalVirtualMemorySize','FreeVirtualMemory','OSArchitecture','Organization','LocalDateTime',`
                                 'RegisteredUser','OperatingSystemSKU','OSType','LastBootUpTime','InstallDate')
                $WMI_ProcProps = @('Name','Description','MaxClockSpeed','CurrentClockSpeed','AddressWidth','NumberOfCores','NumberOfLogicalProcessors', `
                                   'DataWidth')
                $WMI_CompProps = @('DNSHostName','Domain','Manufacturer','Model','NumberOfLogicalProcessors','NumberOfProcessors','PrimaryOwnerContact', `
                                   'PrimaryOwnerName','TotalPhysicalMemory','UserName')
                $WMI_ChassisProps = @('ChassisTypes','Manufacturer','SerialNumber','Tag','SKU')
                $WMI_BIOSProps = @('Version','SerialNumber')
                
                # WMI data
                $wmi_compsystem = Get-WmiObject @WMIHast -Class Win32_ComputerSystem | select $WMI_CompProps
                $wmi_os = Get-WmiObject @WMIHast -Class Win32_OperatingSystem | select $WMI_OSProps
                $wmi_proc = Get-WmiObject @WMIHast -Class Win32_Processor | select $WMI_ProcProps
                $wmi_chassis = Get-WmiObject @WMIHast -Class Win32_SystemEnclosure | select $WMI_ChassisProps
                $wmi_bios = Get-WmiObject @WMIHast -Class Win32_BIOS | select $WMI_BIOSProps

                ## Calculated properties
                # CPU count
                if (@($wmi_proc)[0].NumberOfCores) #Modern OS
                {
                    $Sockets = @($wmi_proc).Count
                    $Cores = ($wmi_proc | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
                    $OSArchitecture = "x" + $(@($wmi_proc)[0]).AddressWidth
                    $SystemArchitecture = "x" + $(@($wmi_proc)[0]).DataWidth
                }
                else #Legacy OS
                {
                    $Sockets = @($wmi_proc | Select-Object -Property SocketDesignation -Unique).Count
                    $Cores = @($wmi_proc).Count
                    $OSArchitecture = "x" + ($wmi_proc | Select-Object -Property AddressWidth -Unique)
                    $SystemArchitecture = "x" + ($wmi_proc | Select-Object -Property DataWidth -Unique)
                }
                
                # OperatingSystemSKU is not availble in 2003 and XP
                if ($wmi_os.OperatingSystemSKU -ne $null)
                {
                    $OS_SKU = $SKUs[$wmi_os.OperatingSystemSKU]
                }
                else
                {
                    $OS_SKU = 'Not Available'
                }
               
                $temptime = ([wmi]'').ConvertToDateTime($wmi_os.LocalDateTime)
                $System_Time = "$($temptime.ToShortDateString()) $($temptime.ToShortTimeString())"

                $temptime = ([wmi]'').ConvertToDateTime($wmi_os.LastBootUptime)
                $OS_LastBoot = "$($temptime.ToShortDateString()) $($temptime.ToShortTimeString())"
                
                $temptime = ([wmi]'').ConvertToDateTime($wmi_os.InstallDate)
                $OS_InstallDate = "$($temptime.ToShortDateString()) $($temptime.ToShortTimeString())"
                
                $Uptime = New-TimeSpan -Start $OS_LastBoot -End $System_Time
                $IsVirtual = $false
                $VirtualType = ''
                if ($wmi_bios.Version -match "VIRTUAL") 
                {
                    $IsVirtual = $true
                    $VirtualType = "Virtual - Hyper-V"
                }
                elseif ($wmi_bios.Version -match "A M I") 
                {
                    $IsVirtual = $true
                    $VirtualType = "Virtual - Virtual PC"
                }
                elseif ($wmi_bios.Version -like "*Xen*") 
                {
                    $IsVirtual = $true
                    $VirtualType = "Virtual - Xen"
                }
                elseif ($wmi_bios.SerialNumber -like "*VMware*")
                {
                    $IsVirtual = $true
                    $VirtualType = "Virtual - VMWare"
                }
                elseif ($wmi_compsystem.manufacturer -like "*Microsoft*")
                {
                    $IsVirtual = $true
                    $VirtualType = "Virtual - Hyper-V"
                }
                elseif ($wmi_compsystem.manufacturer -like "*VMWare*")
                {
                    $IsVirtual = $true
                    $VirtualType = "Virtual - VMWare"
                }
                elseif ($wmi_compsystem.model -like "*Virtual*")
                {
                    $IsVirtual = $true
                    $VirtualType = "Unknown Virtual Machine"
                }
                $ResultProperty = @{
                    ### Defaults
                    'PSComputerName' = $ComputerName
                    'IsVirtual' = $IsVirtual
                    'VirtualType' = $VirtualType 
                    'Model' = $wmi_compsystem.Model
                    'ChassisModel' = $ChassisModels[$wmi_chassis.ChassisTypes[0]]
                    'SerialNumber' = $wmi_bios.SerialNumber
                    'ComputerName' = $wmi_compsystem.DNSHostName                        
                    'OperatingSystem' = $wmi_os.Caption
                    'OSServicePack' = $wmi_os.ServicePackMajorVersion
                    'OSVersion' = $wmi_os.Version
                    'OSSKU' = $OS_SKU
                    'OSArchitecture' = $OSArchitecture
                    'SystemArchitecture' = $SystemArchitecture
                    'PercentPhysicalMemoryUsed' = [math]::round(((($wmi_os.TotalVisibleMemorySize - $wmi_os.FreePhysicalMemory)/$wmi_os.TotalVisibleMemorySize) * 100),2)
                    'PhysicalMemoryTotal' = ($wmi_os.TotalVisibleMemorySize * 1024) | ConvertTo-KMG
                    'PhysicalMemoryFree' = ($wmi_os.FreePhysicalMemory * 1024) | ConvertTo-KMG
                    'VirtualMemoryTotal' = ($wmi_os.TotalVirtualMemorySize * 1024) | ConvertTo-KMG
                    'VirtualMemoryFree' = ($wmi_os.FreeVirtualMemory * 1024) | ConvertTo-KMG
                    'CPUCores' = $Cores
                    'CPUSockets' = $Sockets
                    'LastBootTime' = $OS_LastBoot
                    'InstallDate' = $OS_InstallDate
                    'SystemTime' = $System_Time
                    'Uptime' = "$($Uptime.days) days $($Uptime.hours) hours $($Uptime.minutes) minutes"
                    '_BIOS' = $wmi_bios
                    '_OS' = $wmi_os
                    '_System' = $wmi_compsystem
                    '_Processor' = $wmi_proc
                    '_Chassis' = $wmi_chassis
                }
                #endregion GeneralInfo
                
                #region Memory
                if ($IncludeMemoryInfo)
                {
                    Write-Verbose -Message ('Remote Asset Information: Runspace {0}: Memory information' -f $ComputerName)
                    $WMI_MemProps = @('BankLabel','DeviceLocator','Capacity','PartNumber','Speed','Tag','FormFactor','TypeDetail')
                    $WMI_MemArrayProps = @('Tag','MemoryDevices','MaxCapacity')
                    $wmi_memory = Get-WmiObject @WMIHast -Class Win32_PhysicalMemory | select $WMI_MemProps
                    $wmi_memoryarray = Get-WmiObject @WMIHast -Class Win32_PhysicalMemoryArray | select $WMI_MemArrayProps
                    
                    # Memory Calcs
                    $Memory_Slotstotal = 0
                    $Memory_SlotsUsed = (@($wmi_memory)).Count                
                    @($wmi_memoryarray) | % {$Memory_Slotstotal = $Memory_Slotstotal + $_.MemoryDevices}
                    
                    # Add to the existing property set
                    $ResultProperty.MemorySlotsTotal = $Memory_Slotstotal
                    $ResultProperty.MemorySlotsUsed = $Memory_SlotsUsed
                    $ResultProperty._MemoryArray = $wmi_memoryarray
                    $ResultProperty._Memory = $wmi_memory
                    
                    # Add a few of these properties to our default property set
                    $defaultProperties += 'MemorySlotsTotal'
                    $defaultProperties += 'MemorySlotsUsed'
                    
                    # Add a more detailed memory slot utilization object array (cause I'm nice)
                    $membankcounter = 1
                    $MemorySlotOutput = @()
                    foreach ($obj1 in $wmi_memoryarray)
                    {
                        $slots = $obj1.MemoryDevices + 1
                            
                        foreach ($obj2 in $wmi_memory)
                        {
                            if($obj2.BankLabel -eq "")
                            {
                                $MemLabel = $obj2.DeviceLocator
                            }
                            else
                            {
                                $MemLabel = $obj2.BankLabel
                            }       
                            $slotprops = @{
                                'Bank' = "Bank " + $membankcounter
                                'Label' = $MemLabel
                                'Capacity' = $obj2.Capacity/1024/1024
                                'Speed' = $obj2.Speed
                                'FormFactor' = $MemoryModels[$obj2.FormFactor]
                                'Detail' = $MemoryDetail[[string]$obj2.TypeDetail]
                            }
                            $MemorySlotOutput += New-Object PSObject -Property $slotprops
                            $membankcounter = $membankcounter + 1
                        }
                        while($membankcounter -lt $slots)
                        {
                            $slotprops = @{
                                'Bank' = "Bank " + $membankcounter
                                'Label' = "EMPTY"
                                'Capacity' = ''
                                'Speed' = ''
                                'FormFactor' = "EMPTY"
                                'Detail' = "EMPTY"
                            }
                            $MemorySlotOutput += New-Object PSObject -Property $slotprops
                            $membankcounter = $membankcounter + 1
                        }
                    }
                    $ResultProperty._MemorySlots = $MemorySlotOutput
                }
                #endregion Memory
                
                #region Network
                if ($IncludeNetworkInfo)
                {
                    Write-Verbose -Message ('Remote Asset Information: Runspace {0}: Network information' -f $ComputerName)
                    $wmi_netadapters = Get-WmiObject @WMIHast -Class Win32_NetworkAdapter
                    $alladapters = @()
                    ForEach ($adapter in $wmi_netadapters)
                    {  
                        $wmi_netconfig = Get-WmiObject @WMIHast -Class Win32_NetworkAdapterConfiguration `
                                                                -Filter "Index = '$($Adapter.Index)'"
                        $wmi_promisc = Get-WmiObject @WMIHast -Class MSNdis_CurrentPacketFilter `
                                                              -Namespace 'root\WMI' `
                                                              -Filter "InstanceName = '$($Adapter.Name)'"
                        $promisc = $False
                        if ($wmi_promisc.NdisCurrentPacketFilter -band 0x00000020)
                        {
                            $promisc = $True
                        }
                        
                        $NetConStat = ''
                        if ($adapter.NetConnectionStatus -ne $null)
                        {
                            $NetConStat = $NetConnectionStatus[$adapter.NetConnectionStatus]
                        }
                        $alladapters += New-Object PSObject -Property @{
                              NetworkName = $adapter.NetConnectionID
                              AdapterName = $adapter.Name
                              ConnectionStatus = $NetConStat
                              Index = $wmi_netconfig.Index
                              IpAddress = $wmi_netconfig.IpAddress
                              IpSubnet = $wmi_netconfig.IpSubnet
                              MACAddress = $wmi_netconfig.MACAddress
                              DefaultIPGateway = $wmi_netconfig.DefaultIPGateway
                              Description = $wmi_netconfig.Description
                              InterfaceIndex = $wmi_netconfig.InterfaceIndex
                              DHCPEnabled = $wmi_netconfig.DHCPEnabled
                              DHCPServer = $wmi_netconfig.DHCPServer
                              DNSDomain = $wmi_netconfig.DNSDomain
                              DNSDomainSuffixSearchOrder = $wmi_netconfig.DNSDomainSuffixSearchOrder
                              DomainDNSRegistrationEnabled = $wmi_netconfig.DomainDNSRegistrationEnabled
                              WinsPrimaryServer = $wmi_netconfig.WinsPrimaryServer
                              WinsSecondaryServer = $wmi_netconfig.WinsSecondaryServer
                              PromiscuousMode = $promisc
                       }
                    }
                    $ResultProperty._Network = $alladapters
                }                    
                #endregion Network
                
                #region Disk
                if ($IncludeDiskInfo)
                {
                    Write-Verbose -Message ('Remote Asset Information: Runspace {0}: Disk information' -f $ComputerName)
                    $WMI_DiskPartProps    = @('DiskIndex','Index','Name','DriveLetter','Caption','Capacity','FreeSpace','SerialNumber')
                    $WMI_DiskVolProps     = @('Name','DriveLetter','Caption','Capacity','FreeSpace','SerialNumber')
                    $WMI_DiskMountProps   = @('Name','Label','Caption','Capacity','FreeSpace','Compressed','PageFilePresent','SerialNumber')
                    
                    # WMI data
                    $wmi_diskdrives = Get-WmiObject @WMIHast -Class Win32_DiskDrive | select $WMI_DiskDriveProps
                    $wmi_mountpoints = Get-WmiObject @WMIHast -Class Win32_Volume -Filter "DriveType=3 AND DriveLetter IS NULL" | select $WMI_DiskMountProps
                    
                    $AllDisks = @()
                    foreach ($diskdrive in $wmi_diskdrives) 
                    {
                        $partitionquery = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($diskdrive.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
                        $partitions = @(Get-WmiObject @WMIHast -Query $partitionquery)
                        foreach ($partition in $partitions)
                        {
                            $logicaldiskquery = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($partition.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition"
                            $logicaldisks = @(Get-WmiObject @WMIHast -Query $logicaldiskquery)
                            foreach ($logicaldisk in $logicaldisks)
                            {
                               $diskprops = @{
                                               Disk = $diskdrive.Name
                                               Model = $diskdrive.Model
                                               Partition = $partition.Name
                                               Description = $partition.Description
                                               PrimaryPartition = $partition.PrimaryPartition
                                               VolumeName = $logicaldisk.VolumeName
                                               Drive = $logicaldisk.Name
                                               DiskSize = $logicaldisk.Size | ConvertTo-KMG
                                               FreeSpace = $logicaldisk.FreeSpace | ConvertTo-KMG
                                               PercentageFree = [math]::round((($logicaldisk.FreeSpace/$logicaldisk.Size)*100), 2)
                                               DiskType = 'Partition'
                                               SerialNumber = $diskdrive.SerialNumber
                                             }
                                $AllDisks += New-Object psobject -Property $diskprops
                            }
                        }
                    }
                    # Mountpoints are wierd so we do them seperate.
                    if ($wmi_mountpoints)
                    {
                        foreach ($mountpoint in $wmi_mountpoints)
                        {                    
                            $diskprops = @{
                                   Disk = $mountpoint.Name
                                   Model = ''
                                   Partition = ''
                                   Description = $mountpoint.Caption
                                   PrimaryPartition = ''
                                   VolumeName = ''
                                   VolumeSerialNumber = ''
                                   Drive = [Regex]::Match($mountpoint.Caption, "^.:\\").Value
                                   DiskSize = $mountpoint.Capacity  | ConvertTo-KMG
                                   FreeSpace = $mountpoint.FreeSpace  | ConvertTo-KMG
                                   PercentageFree = [math]::round((($mountpoint.FreeSpace/$mountpoint.Capacity)*100), 2)
                                   DiskType = 'MountPoint'
                                   SerialNumber = $mountpoint.SerialNumber
                                 }
                            $AllDisks += New-Object psobject -Property $diskprops
                        }
                    }
                    $ResultProperty._Disks = $AllDisks
                }
                #endregion Disk
                
                # Final output
                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty

                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.Asset.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers

                Write-Output -InputObject $ResultObject
            }
            catch
            {
                Write-Warning -Message ('Remote Asset Information: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Asset Information: Runspace {0}: End' -f $ComputerName)
        }
 
        function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Asset Information: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Asset Information: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Remote Asset Information: Getting asset info'
                        Status = '{0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('IncludeMemoryInfo',$IncludeMemoryInfo)
            $null = $psCMD.AddParameter('IncludeDiskInfo',$IncludeDiskInfo)
            $null = $psCMD.AddParameter('IncludeNetworkInfo',$IncludeNetworkInfo)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)               # Passthrough the hidden verbose option so write-verbose works within the runspaces
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Asset Information: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
                })
           Get-Result
        }
    }
 
    End
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Asset Information: Getting asset info' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Asset Information: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-HPServerHealth
{
    <#
    .SYNOPSIS
       Get HP server hardware health status information from WBEM wmi providers.
    .DESCRIPTION
       Get HP server hardware health status information from WBEM wmi providers. Results returned are the overall
       health status of the server by default. Optionally further health information about several individual
       components can be queries and returned as well.
    .PARAMETER ComputerName
       Specifies the target computer or computers for data query.
    .PARAMETER IncludePSUHealth
        Include power supply health results
    .PARAMETER IncludeTempSensors
        Include temperature sensor results
    .PARAMETER IncludeEthernetTeamHealth
        Include ethernet team health results
    .PARAMETER IncludeFanHealth
        Include fan health results
    .PARAMETER IncludeEthernetHealth
       Include ethernet adapter health results
    .PARAMETER IncludeHBAHealth
        Include HBA health results
    .PARAMETER IncludeArrayControllerHealth
       Include array controller health results
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information
    .PARAMETER PromptForCredential
       Prompt for remote system credential prior to processing request.
    .PARAMETER Credential
       Accept alternate credential (ignored if the localhost is processed)
    .EXAMPLE
       PS> $cred = get-credential
       PS> Get-HPServerHealth -ComputerName 'TestServer' -Credential $cred
       
            ComputerName                       Manufacturer           HealthState
            ------------                       ------------           -----------
            TestServer                         HP                     OK       
            
       Description
       -----------
       Attempts to retrieve overall health status of TestServer.

    .EXAMPLE
       PS> $cred = get-credential
       PS> $c = Get-HPServerhealth -ComputerName 'TestServer' -Credential $cred 
                                   -IncludeEthernetTeamHealth 
                                   -IncludeArrayControllerHealth 
                                   -IncludeEthernetHealth 
                                   -IncludeFanHealth 
                                   -IncludeHBAHealth 
                                   -IncludePSUHealth 
                                   -IncludeTempSensors
        
       PS> $c._TempSensors | select Name, Description, PercentToCritical

        Name                         Description                                                   PercentToCritical
        ----                         -----------                                                   -----------------
        Temperature Sensor 1         Temperature Sensor 1 detects for Ambient / External /...                  41.46
        Temperature Sensor 2         Temperature Sensor 2 detects for CPU board                                48.78
        Temperature Sensor 3         Temperature Sensor 3 detects for CPU board                                48.78
        Temperature Sensor 4         Temperature Sensor 4 detects for Memory board                             29.89
        Temperature Sensor 5         Temperature Sensor 5 detects for Memory board                             28.74
        Temperature Sensor 6         Temperature Sensor 6 detects for Memory board                             32.18
        Temperature Sensor 7         Temperature Sensor 7 detects for Memory board                             33.33
        Temperature Sensor 8         Temperature Sensor 8 detects for Power Supply Bays                        42.22
        Temperature Sensor 9         Temperature Sensor 9 detects for Power Supply Bays                        47.69
        Temperature Sensor 10        Temperature Sensor 10 detects for System board                            46.67
        Temperature Sensor 11        Temperature Sensor 11 detects for System board                            38.57
        Temperature Sensor 12        Temperature Sensor 12 detects for System board                            41.11
        Temperature Sensor 13        Temperature Sensor 13 detects for I/O board                               41.43
        Temperature Sensor 14        Temperature Sensor 14 detects for I/O board                               48.57
        Temperature Sensor 15        Temperature Sensor 15 detects for I/O board                               44.29
        Temperature Sensor 19        Temperature Sensor 19 detects for System board                               30
        Temperature Sensor 20        Temperature Sensor 20 detects for System board                            38.57
        Temperature Sensor 21        Temperature Sensor 21 detects for System board                            33.75
        Temperature Sensor 22        Temperature Sensor 22 detects for System board                            33.75
        Temperature Sensor 23        Temperature Sensor 23 detects for System board                            45.45
        Temperature Sensor 24        Temperature Sensor 24 detects for System board                               40
        Temperature Sensor 25        Temperature Sensor 25 detects for System board                               40
        Temperature Sensor 26        Temperature Sensor 26 detects for System board                               40
        Temperature Sensor 29        Temperature Sensor 29 detects for Storage bays                            58.33
        Temperature Sensor 30        Temperature Sensor 30 detects for System board                            60.91
       Description
       -----------
       Gathers all HP health information about TestServer using an alternate credential. Displays the temperature
       sensor information.

    .NOTES
       For obvious reasons, you will need to have the HP WBEM software installed on the server.
       
       WBEM Provider Download:
       http://h18004.www1.hp.com/products/servers/management/wbem/providerdownloads.html
       
       If you are troubleshooting this function your best bet is to use the hidden verbose option 
       when calling the function. This will display information within each runspace at appropriate intervals.
       
       Version History
       1.0.1 - 8/28/2013
        - Added differentiation between array system and controller health
       1.0.0 - 8/22/2013
        - Initial Release
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0,
                   HelpMessage="Computer or array of computer names to process")]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter(HelpMessage="Include power supply health results")]
        [switch]
        $IncludePSUHealth,
  
        [Parameter(HelpMessage="Include temperature sensor results")]
        [switch]
        $IncludeTempSensors,
 
        [Parameter(HelpMessage="Include ethernet team health results")]
        [switch]
        $IncludeEthernetTeamHealth,

        [Parameter(HelpMessage="Include fan health results")]
        [switch]
        $IncludeFanHealth,

        [Parameter(HelpMessage="Include ethernet adapter health results")]
        [switch]
        $IncludeEthernetHealth,
        
        [Parameter(HelpMessage="Include HBA health results")]
        [switch]
        $IncludeHBAHealth,
        
        [Parameter(HelpMessage="Include array controller health results")]
        [switch]
        $IncludeArrayControllerHealth,
       
        [Parameter(HelpMessage="Maximum amount of runspaces")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout in seconds for each runspace before it gives up")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display visual progress bar")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
        
        [Parameter()]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )
    BEGIN
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'HP Server Health: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'HP Server Health: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'HP Server Health: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "HP Server Health: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'HP Server Health: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'HP Server Health: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID,
                
                [Parameter()]
                [switch]
                $IncludePSUHealth,
          
                [Parameter()]
                [switch]
                $IncludeTempSensors,
         
                [Parameter()]
                [switch]
                $IncludeEthernetTeamHealth,

                [Parameter()]
                [switch]
                $IncludeFanHealth,
 
                [Parameter()]
                [switch]
                $IncludeEthernetHealth,
                
                [Parameter()]
                [switch]
                $IncludeHBAHealth,
                
                [Parameter()]
                [switch]
                $IncludeArrayControllerHealth
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('HP Server Health: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                #region Lookup arrays
                $BatteryStatus=@{
                    '1'="OK"
                    '2'="Degraded"
                    '3'="Not Fully Charged"
                    '4'="Not Present"
                    }
                $OperationalStatus=@{
                    '0'="Unknown"
                    '2'="OK"
                    '3'="Degraded"
                    '6'="Error"
                    } 
                $HealthStatus=@{
                    '0'="Unknown"
                    '5'="OK"
                    '10'="Degraded"
                    '15'="Minor"
                    '20'="Major"
                    '25'="Critical"
                    '30'="Non-Recoverable"
                    }
                $TeamStatus=@{
                    '0'="Unknown"
                    '2'="OK"
                    '3'="Degraded"
                    '4'="Redundancy Lost"
                    '5'="Overall Failure"
                    }
                $FanRemovalConditions=@{
                    '3'="Removable when off"
                    '4'="Removable when on or off"
                    }
                $EthernetPortType = @{
                    "0" =   "Unknown"
                    "1" =   "Other"
                    "50" =  "10BaseT"
                    "51" =  "10-100BaseT"
                    "52" =  "100BaseT"
                    "53" =  "1000BaseT"
                    "54" =  "2500BaseT"
                    "55" =  "10GBaseT"
                    "56" =  "10GBase-CX4"
                    "100" = "100Base-FX"
                    "101" = "100Base-SX"
                    "102" = "1000Base-SX"
                    "103" = "1000Base-LX"
                    "104" = "1000Base-CX"
                    "105" = "10GBase-SR"
                    "106" = "10GBase-SW"
                    "107" = "10GBase-LX4"
                    "108" = "10GBase-LR"
                    "109" = "10GBase-LW"
                    "110" = "10GBase-ER"
                    "111" = "10GBase-EW"
                }
                # Get this definition from hp_sensor.mof
                $PSUType = @("Unknown","Other","System board","Host System board","I/O board","CPU board", `
                             "Memory board","Storage bays","Removable Media Bays","Power Supply Bays", `
                             "Ambient / External / Room","Chassis","Bridge Card","Management board",`
                             "Remote Management Card","Generic Backplane","Infrastructure Network", `
                             "Blade Slot in Chassis/Infrastructure","Compute Cabinet Bulk Power Supply",`
                             "Compute Cabinet System Backplane Power Supply",`
                             "Compute Cabinet I/O chassis enclosure Power Supply",`
                             "Compute Cabinet AC Input Line","I/O Expansion Cabinet Bulk Power Supply",`
                             "I/O Expansion Cabinet System Backplane Power Supply",`
                             "I/O Expansion Cabinet I/O chassis enclosure Power Supply",
                             "I/O Expansion Cabinet AC Input Line","Peripheral Bay","Device Bay","Switch")
                $SensorType = @("Unknown","Other","System board","Host System board","I/O board","CPU board",`
                                "Memory board","Storage bays","Removable Media Bays","Power Supply Bays",`
                                "Ambient / External / Room","Chassis","Bridge Card","Management board",`
                                "Remote Management Card","Generic Backplane","Infrastructure Network",`
                                "Blade Slot in Chassis/Infrastructure","Front Panel","Back Panel","IO Bus",`
                                "Peripheral Bay","Device Bay","Switch","Software-defined")
                #endregion Lookup arrays
                
                # Change the default output properties here
                $defaultProperties = @('ComputerName','Manufacturer','HealthState')
          
                Write-Verbose -Message ('HP Server Health: Runspace {0}: Server general information' -f $ComputerName)                
                # Modify this variable to change your default set of display properties
                $WMI_CompProps = @('DNSHostName','Manufacturer')
                $wmi_compsystem = Get-WmiObject @WMIHast -Class Win32_ComputerSystem | select $WMI_CompProps
                if (($wmi_compsystem.Manufacturer -eq "HP") -or ($wmi_compsystem.Manufacturer -like "Hewlett*"))
                {
                    if (Get-WmiObject @WMIHast -Namespace 'root' -Class __NAMESPACE -filter "name='hpq'") 
                    {
                        #region HP General
                        Write-Verbose -Message ('HP Server Health: Runspace {0}: HP general health information' -f $ComputerName)
                        $WMI_HPHealthProps = @('HealthState')
                        $wmi_hphealth = Get-WmiObject @WMIHast -Namespace 'root\hpq' -Class  HP_WinComputerSystem | 
                                        Select $WMI_HPHealthProps
                        
                        $ResultProperty = @{
                            ### Defaults
                            'PSComputerName' = $ComputerName
                            'ComputerName' = $wmi_compsystem.DNSHostName
                            'Manufacturer' = $wmi_compsystem.Manufacturer
                            'HealthState' = $Healthstatus[[string]$wmi_hphealth.HealthState]
                        }
                        #endregion HP General
                        
                        #region HP PSU
                        if ($IncludePSUHealth)
                        {
                            Write-Verbose -Message ('HP Server Health: Runspace {0}: HP PSU health information' -f $ComputerName)
                            $WMI_HPPowerProps = @('ElementName','PowerSupplyType','HealthState')                
                            $wmi_hppower = @(Get-WmiObject @WMIHast -Namespace 'root\hpq' -Class HP_WinPowerSupply | 
                                             Select $WMI_HPPowerProps)
                            $_PSUHealth = @()
                            foreach ($psu in $wmi_hppower)
                            {
                                $psuprop = @{
                                    'Name' = $psu.ElementName
                                    'Type' = $PSUType[[int]$psu.PowerSupplyType]
                                    'HealthState' = $HealthStatus[[string]$psu.HealthState]
                                }
                                $_PSUHealth += New-Object PSObject -Property $psuprop
                            }
                            $ResultProperty._PSUHealth = @($_PSUHealth)
                        }
                        #endregion HP PSU
                        
                        #region HP Temperature Sensors
                        if ($IncludeTempSensors)
                        {                
                            Write-Verbose -Message ('HP Server Health: Runspace {0}: HP sensor information' -f $ComputerName)
                            $WMI_HPTempSensorProps = @('ElementName','SensorType','Description','CurrentReading',`
                                                   'UpperThresholdCritical')
                            $wmi_hptempsensor = @(Get-WmiObject @WMIHast -Namespace 'root\hpq' -Class HP_WinNumericSensor |
                                                Select $WMI_HPTempSensorProps)
                            $_TempSensors = @()
                            foreach ($sensor in $wmi_hptempsensor)
                            {
                                $PercentCrit = 0
                                if (($sensor.CurrentReading) -and ($sensor.UpperThresholdCritical))
                                {
                                    $PercentCrit = [math]::round((($sensor.CurrentReading/$sensor.UpperThresholdCritical)*100), 2)
                                }
                                $sensorprop = @{
                                    'Name' = $sensor.ElementName
                                    'Type' = $SensorType[[int]$sensor.SensorType]
                                    'Description' = [regex]::Match($sensor.Description,"(.+)(?=\..+$)").Value
                                    'CurrentReading' = $sensor.CurrentReading
                                    'UpperThresholdCritical' = $sensor.UpperThresholdCritical
                                    'PercentToCritical' = $PercentCrit
                                }
                                $_TempSensors += New-Object PSObject -Property $sensorprop
                            }
                            $ResultProperty._TempSensors = @($_TempSensors)
                        }
                        #endregion HP Temperature Sensors              
                        
                        #region HP Ethernet Team
                        if ($IncludeEthernetTeamHealth)
                        {
                            Write-Verbose -Message ('HP Server Health: Runspace {0}: HP ethernet team information' -f $ComputerName)
                            $WMI_HPEthTeamsProps = @('ElementName','Description','RedundancyStatus')
                            $wmi_ethernetteam = @(Get-WmiObject @WMIHast -Namespace 'root\hpq' -Class HP_EthernetTeam |
                                                Select $WMI_HPEthTeamsProps)
                            $_EthernetTeamHealth = @()
                            foreach ($ethteam in $wmi_ethernetteam)
                            {
                                $ethteamprop = @{
                                    'Name' = $ethteam.ElementName
                                    'Description' = $ethteam.Description
                                    'RedundancyStatus' = $TeamStatus[[string]$ethteam.RedundancyStatus]
                                }
                                $_EthernetTeamHealth += New-Object PSObject -Property $ethteamprop
                            }
                            $ResultProperty._EthernetTeamHealth = @($_EthernetTeamHealth)
                        }
                        #endregion HP Ethernet Team
                        
                        #region HP Fans
                        if ($IncludeFanHealth)
                        {
                            Write-Verbose -Message ('HP Server Health: Runspace {0}: HP fan information' -f $ComputerName)
                            $WMI_HPFanProps = @('ElementName','HealthState','RemovalConditions')
                            $wmi_fans = @(Get-WmiObject @WMIHast -Namespace 'root\hpq' -Class HP_FanModule |
                                          Select $WMI_HPFanProps)
                            $_FanHealth = @()
                            foreach ($fan in $wmi_fans)
                            {
                                $fanprop = @{
                                    'Name' = $fan.ElementName
                                    'HealthState' = $HealthStatus[[string]$fan.HealthState]
                                    'RemovalConditions' = $FanRemovalConditions[[string]$fan.RemovalConditions]
                                }
                                $_FanHealth += New-Object PSObject -Property $fanprop
                            }
                            $ResultProperty._FanHealth = @($_FanHealth)
                        }
                        #endregion HP Fans
                        
                        #region HP Ethernet
                        if ($IncludeEthernetHealth)
                        {
                            Write-Verbose -Message ('HP Server Health: Runspace {0}: HP ethernet information' -f $ComputerName)
                            $WMI_HPEthernetPortProps = @('ElementName','PortNumber','PortType','HealthState')
                            $wmi_ethernet = @(Get-WmiObject @WMIHast -Namespace 'root\hpq' -Class HP_EthernetPort |
                                            Select $WMI_HPEthernetPortProps)
                            $_EthernetHealth = @()
                            foreach ($eth in $wmi_ethernet)
                            {
                                $ethprop = @{
                                    'Name' = $eth.ElementName
                                    'HealthState' = $HealthStatus[[string]$eth.HealthState]
                                    'PortType' = $EthernetPortType[[string]$eth.PortType]
                                    'PortNumber' = $eth.PortNumber
                                }
                                $_EthernetHealth += New-Object PSObject -Property $ethprop
                            }
                            $ResultProperty._EthernetHealth = @($_EthernetHealth)
                        }
                        #endregion HP Ethernet
                                        
                        #region HBA
                        if ($IncludeHBAHealth)
                        {
                            Write-Verbose -Message ('HP Server Health: Runspace {0}: HP HBA information' -f $ComputerName)
                            $WMI_HPFCPortProps = @('ElementName','Manufacturer','Model','OtherIdentifyingInfo','OperationalStatus')
                            $wmi_hba = @(Get-WmiObject @WMIHast -Namespace 'root\hpq' -Class HPFCHBA_PhysicalPackage |
                                       Select $WMI_HPFCPortProps)
                            $_HBAHealth = @()
                            foreach ($hba in $wmi_hba)
                            {
                                $hbaprop = @{
                                    'Name' = $hba.ElementName
                                    'Manufacturer' = $hba.Manufacturer
                                    'Model' = $hba.Model
                                    'OtherIdentifyingInfo' = $hba.OtherIdentifyingInfo
                                    'OperationalStatus' = $OperationalStatus[[string]$hba.OperationalStatus]
                                }
                                $_HBAHealth += New-Object PSObject -Property $hbaprop
                            }
                            $ResultProperty._HBAHealth = @($_HBAHealth)
                        }
                        #endregion HBA
                        
                        #region ArrayControllers
                        if ($IncludeArrayControllerHealth)
                        {
                            Write-Verbose -Message ('HP Server Health: Runspace {0}: HP array controller information' -f $ComputerName)
                            #$WMI_ArraySystemProps = @()
                            $WMI_ArrayCtrlProps = @('ElementName','BatteryStatus','ControllerStatus')
                            
                            $wmi_arraysystem = @(Get-WMIObject @WMIHast -Namespace 'root\hpq' -class HPSA_ArraySystem | 
                                                 Select Name,@{n='ArrayStatus';e={$_.StatusDescriptions}})
                            $_ArrayControllers = @()
                            Foreach ($array in $wmi_arraysystem)
                            {
                                $wmi_arrayctrl = @(Get-WMIObject @WMIHast `
                                                                       -Namespace 'root\hpq' `
                                                                       -class HPSA_ArrayController `
                                                                       -filter "name='$($array.Name)'")
                                $BatteryStat = ''
                                if ($wmi_arrayctrl.batterystatus)
                                {
                                    $BatteryStat = $BatteryStatus[[string]$wmi_arrayctrl.batterystatus]
                                }
                                $arrayprop = @{
                                    'ArrayName' = $wmi_arrayctrl.ElementName
                                    'BatteryStatus' = $BatteryStat
                                    'ArrayStatus' = $array.ArrayStatus
                                    'ControllerStatus' = $OperationalStatus[[string]$wmi_arrayctrl.ControllerStatus]
                                }
                                $_ArrayControllers += New-Object PSObject -Property $arrayprop
                            }
                            $ResultProperty._ArrayControllers = $_ArrayControllers
                        }
                        #endregion ArrayControllers
                    }
                    else
                    {
                        Write-Warning -Message ('HP Server Health: {0}: {1}' -f $ComputerName, 'WBEM Provider software needs to be installed')
                    }
                }
                else
                {
                    Write-Warning -Message ('HP Server Health: {0}: {1}' -f $ComputerName, 'Not determined to be HP hardware')
                }

                    # Final output
                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty

                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.HPServerHealth.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers

                Write-Output -InputObject $ResultObject

            }
            catch
            {
                Write-Warning -Message ('HP Server Health: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('HP Server Health: Runspace {0}: End' -f $ComputerName)
        }
 
        function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('HP Server Health: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('HP Server Health: Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('HP Server Health: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Getting asset info'
                        Status = '{0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('IncludePSUHealth',$IncludePSUHealth)       
            $null = $psCMD.AddParameter('IncludeTempSensors',$IncludeTempSensors)
            $null = $psCMD.AddParameter('IncludeEthernetTeamHealth',$IncludeEthernetTeamHealth)
            $null = $psCMD.AddParameter('IncludeFanHealth',$IncludeFanHealth)
            $null = $psCMD.AddParameter('IncludeEthernetHealth',$IncludeEthernetHealth)
            $null = $psCMD.AddParameter('IncludeHBAHealth',$IncludeHBAHealth)
            $null = $psCMD.AddParameter('IncludeArrayControllerHealth',$IncludeArrayControllerHealth)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference) # Passthrough the hidden verbose option so write-verbose works within the runspaces
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('HP Server Health: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
                })
           Get-Result
        }
    }
    END
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'HP Server Health: Getting HP health info' -Status 'Done' -Completed
        }
        Write-Verbose -Message "HP Server Health: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-MultiRunspaceWMIObject
{
    <#
    .SYNOPSIS
       Get generic wmi object data from a remote or local system.
    .DESCRIPTION
       Get wmi object data from a remote or local system. Multiple runspaces are utilized and 
       alternate credentials can be provided.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER Namespace
       Namespace to query
    .PARAMETER Class
       Class to query
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information

    .EXAMPLE
       PS > (Get-MultiRunspaceWMIObject -Class win32_printer).WMIObjects

       <output is all your local printers>
       
       Description
       -----------
       Queries the local machine for all installed printer information and spits out what is found.

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 08/31/2013
        - Initial release
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter(HelpMessage="WMI class to query",
                   Position=1)]
        [string]
        $Class,
        
        [Parameter(HelpMessage="WMI namespace to query")]
        [string]
        $NameSpace = 'root\cimv2',
        
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message ('WMI Query {0}: Creating local hostname list' -f $Class)
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message ('WMI Query {0}: Creating initial variables' -f $Class)
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message ('WMI Query {0}: Creating Initial Session State' -f $Class)
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message ("WMI Query {0}: Adding variable $ExternalVariable to initial session state" -f $Class)
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message ('WMI Query {0}: Creating runspace pool' -f $Class)
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message ('WMI Query {0}: Defining background runspaces scriptblock' -f $Class)
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter()]
                [string]
                $ComputerName,
                
                [Parameter()]
                [int]
                $bgRunspaceID,
                
                [Parameter()]
                [string]
                $Class,
                
                [Parameter()]
                [string]
                $NameSpace = 'root\cimv2'
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('WMI Query {0}: Runspace {1}: Start' -f $Class,$ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                $PSDateTime = Get-Date
                
                #region WMI Data
                Write-Verbose -Message ('WMI Query {0}: Runspace {1}: WMI information' -f $Class,$ComputerName)

                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','WMIObjects')
                                         
                # WMI data
                $wmi_data = Get-WmiObject @WMIHast -Namespace $Namespace -Class $Class

                $ResultProperty = @{
                    'PSComputerName' = $ComputerName
                    'PSDateTime' = $PSDateTime
                    'ComputerName' = $ComputerName
                    'WMIObjects' = $wmi_data
                }
                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty
                
                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.WMIObject.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                Write-Output -InputObject $ResultObject
                #endregion WMI Data
            }
            catch
            {
                Write-Warning -Message ('WMI Query {0}: {1}: {2}' -f $Class, $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('WMI Query {0}: Runspace {1}: End' -f $Class,$ComputerName)
        }
 
        function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('WMI Query {0}: Thread done for {1}' -f $Class,$runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('WMI Query {0}: Timeout {1}' -f $Class,$runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('WMI Query {0}: Removing {1} from runspaces' -f $Class,$threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = ('WMI Query {0}: Getting info' -f $Class)
                        Status = 'WMI Query {0}: {1} of {2} total threads done' -f $Class,($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    Process
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Class',$Class)
            $null = $psCMD.AddParameter('Namespace',$Namespace)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('WMI Query {0}: Starting {1}' -f $Class,$Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
 
    End
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity ('WMI Query {0}: Getting wmi information' -f $Class) -Status 'Done' -Completed
        }
        Write-Verbose -Message ("WMI Query {0}: Closing runspace pool" -f $Class)
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-RemoteInstalledPrograms
{
    <#
    .SYNOPSIS
       Retrieves installed programs from remote systems via the registry.
    .DESCRIPTION
       Retrieves installed programs from remote systems via the registry.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information
    .EXAMPLE
       PS > Get-RemoteInstalledPrograms
       
       Description
       -----------
       Lists all of the programs found in the registry of the localhost.

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 08/28/2013
        - Initial release
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter(HelpMessage="Maximum number of concurrent runspaces.")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a runspaces stops trying to gather the information.")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function.")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials.")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials.")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Installed Programs: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Remote Installed Programs: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Installed Programs: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Installed Programs: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Installed Programs: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Installed Programs: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,

                [Parameter()]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Remote Installed Programs: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }
                
                #region Installed Programs
                Write-Verbose -Message ('Remote Installed Programs: Runspace {0}: Gathering registry information' -f $ComputerName)
                $hklm = '2147483650'
                $basekey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
                  
                # WMI data
                $wmi_data = Get-WmiObject @WMIHast -Class StdRegProv -Namespace 'root\default' -List:$true
                $allsubkeys = $wmi_data.EnumKey($hklm,$basekey)
                $Programs = @()
                
                foreach ($subkey in $allsubkeys.sNames) 
                {
                   # $keydata = $wmi_data.EnumValues($hklm,"$basekey\$subkey")
                    $displayname = $wmi_data.GetStringValue($hklm,"$basekey\$subkey",'DisplayName').sValue
                    if ($DisplayName)
                    {
                        $publisher = $wmi_data.GetStringValue($hklm,"$basekey\$subkey",'Publisher').sValue
                        $uninstallstring = $wmi_data.GetExpandedStringValue($hklm,"$basekey\$subkey",'UninstallString').sValue
                        
                        $ProgramProperty = @{
                            'DisplayName' = $displayname
                            'Publisher' = $publisher
                            'UninstallString' = $uninstallstring
                        }
                        $Programs += New-Object PSObject -Property $ProgramProperty
                    }
                }
                If ($Programs.Count -gt 0)
                {
                    $ResultProperty = @{
                        'PSComputerName' = $ComputerName
                        'ComputerName' = $ComputerName
                        'Programs' = $Programs
                    }
                    $ResultObject = New-Object PSObject -Property $ResultProperty
                    Write-Output -InputObject $ResultObject
                }
            }
            catch
            {
                Write-Warning -Message ('Remote Installed Programs: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Installed Programs: Runspace {0}: End' -f $ComputerName)
        }
 
        function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Installed Programs: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Remote Installed Programs: Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Installed Programs: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Getting installed programs'
                        Status = 'Remote Installed Programs: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    Process
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Installed Programs: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
    End
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Installed Programs: Getting program listing' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Installed Programs: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-RemoteInstalledPrinters
{
    <#
    .SYNOPSIS
       Gather remote printer information with multiple runspaces and wmi.
    .DESCRIPTION
       Gather remote printer information with multiple runspaces and wmi. Can provide alternate credentials if
       required.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information

    .EXAMPLE
       PS > (Get-RemoteInstalledPrinters).Printers | Select Name,Status,CurrentJobs

        Name                                      Status                   CurrentJobs
        ----                                      ------                   -----------
        Send To OneNote 2010                      Idle                               0
        PDFCreator                                Idle                               0
        Microsoft XPS Document Writer             Idle                               0
        Foxit Reader PDF Printer                  Idle                               0
        Fax                                       Idle                               0

       
       Description
       -----------
       Get a list of locally installed printers (both network and locally attached) and show the status
       and current number of jobs in its queue.

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 09/01/2013
        - Initial release
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
       
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Printers: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Printers: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Printers: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Printers: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Printers: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Remote Printers: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                # General variables
                $PrinterObjects = @()
                $PSDateTime = Get-Date
                
                #region Printers
                $lookup_printerstatus = @('PlaceHolder','Other','Unknown','Idle','Printing', `
                          'Warming Up','Stopped printing','Offline')
                          
                Write-Verbose -Message ('Remote Printers: Runspace {0}: Printer information' -f $ComputerName)

                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','Printers')

                # WMI data
                $wmi_printers = Get-WmiObject @WMIHast -Class Win32_Printer
                foreach ($printer in $wmi_printers)
                {
                    if (($printer.Name -ne '_Total') -and ($printer.Name -notlike '\\*'))
                    {
                        $Filter = "Name='$($printer.Name)'"
                        $wmi_printerqueues = Get-WMIObject @WMIHast `
                                                           -Class Win32_PerfFormattedData_Spooler_PrintQueue `
                                                           -Filter $Filter
                        $CurrJobs = $wmi_printerqueues.Jobs
                        $TotalJobs = $wmi_printerqueues.TotalJobsPrinted
                        $TotalPages = $wmi_printerqueues.TotalPagesPrinted
                        $JobErrors = $wmi_printerqueues.JobErrors 
                    }
                    else
                    {
                        $CurrJobs = 'NA'
                        $TotalJobs = 'NA'
                        $TotalPages = 'NA'
                        $JobErrors = 'NA'
                    }
                    $PrinterProperty = @{
                        'Name' = $printer.Name
                        'Status' = $lookup_printerstatus[[int]$printer.PrinterStatus]
                        'Location' = $printer.Location
                        'Shared' = $printer.Shared
                        'ShareName' = $printer.ShareName
                        'Published' = $printer.Published
                        'Local' = $printer.Local
                        'Network' = $printer.Network
                        'KeepPrintedJobs' = $printer.KeepPrintedJobs
                        'Driver Name' = $printer.DriverName
                        'PortName' = $printer.PortName
                        'Default' = $printer.Default
                        'CurrentJobs' = $CurrJobs
                        'TotalJobsPrinted' = $TotalJobs
                        'TotalPagesPrinted' = $TotalPages
                        'JobErrors' = $JobErrors
                    }
                    $PrinterObjects += New-Object PSObject -Property $PrinterProperty
                }

                $ResultProperty = @{
                    'PSComputerName' = $ComputerName
                    'PSDateTime' = $PSDateTime
                    'ComputerName' = $ComputerName
                    'Printers' = $PrinterObjects
                }
                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty
                
                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.Printer.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                #endregion Printers

                Write-Output -InputObject $ResultObject
            }
            catch
            {
                Write-Warning -Message ('Remote Printers: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Printers: Runspace {0}: End' -f $ComputerName)
        }
 
        function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Printers: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Printers: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Remote Printers: Getting info'
                        Status = 'Remote Printers: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    Process
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Printers: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
 
    End
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Printers: Getting printer information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Printers: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-RemoteEventLogs
{
    <#
    .SYNOPSIS
       Retrieves event logs via WMI in multiple runspaces.
    .DESCRIPTION
       Retrieves event logs via WMI and, if needed, alternate credentials. This function utilizes multiple runspaces.
    .PARAMETER ComputerName
       Specifies the target computer or comptuers for data query.
    .PARAMETER Hours
       Gather event logs from the last number of hourse specified here.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information
    .EXAMPLE
       PS > Get-RemoteEventLogs
       
       Description
       -----------
       Lists all of the event logs found on the localhost in the last 24 hours.

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 08/28/2013
        - Initial release
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter(HelpMessage="Gather logs for this many previous hours.")]
        [ValidateRange(1,65535)]
        [int32]
        $Hours = 24,
        
        [Parameter(HelpMessage="Maximum number of concurrent runspaces.")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a runspaces stops trying to gather the information.")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function.")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials.")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials.")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Event Logs: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Remote Event Logs: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Event Logs: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Event Logs: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Event Logs: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Event Logs: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
                
                [Parameter()]
                [int32]
                $Hours,

                [Parameter()]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Remote Event Logs: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                Write-Verbose -Message ('Remote Event Logs: Runspace {0}: Gathering logs in last {1} hours' -f $ComputerName,$Hours)

                #Statics
                $time = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime((Get-Date).AddHours(-$Hours))
                
                # WMI data
                $_EventLogSettings = Get-WmiObject @WMIHast -Class win32_NTEventlogFile | 
                    Where {$_.NumberOfRecords -gt 0}
                $EventLogResults = @()
                Foreach ($LogFile in $_EventLogSettings.LogfileName)
                {
                    Write-Verbose -Message ('Remote Event Logs: Runspace {0}: Processing {1} log file' -f $ComputerName,$Logfile)
                    $filter = "(Type <> 'information' AND Type <> 'audit success') and TimeGenerated>='$time' and LogFile='$LogFile'"
                    $EventLogResults += Get-WmiObject @WMIHast -Class Win32_NTLogEvent -filter  $filter |
                                              Sort-Object -Property TimeGenerated -Descending
                }
                
                If ($EventLogResults.Count -gt 0)
                {
                    $ResultProperty = @{
                        'PSComputerName' = $ComputerName
                        'ComputerName' = $ComputerName
                        'EventLogs' = $EventLogResults
                    }
                    $ResultObject = New-Object PSObject -Property $ResultProperty
                    Write-Output -InputObject $ResultObject
                }
            }
            catch
            {
                Write-Warning -Message ('Remote Event Logs: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Event Logs: Runspace {0}: End' -f $ComputerName)
        }
 
        function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Event Logs: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Remote Event Logs: Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Event Logs: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Getting installed programs'
                        Status = 'Remote Event Logs: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    Process
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Hours',$Hours)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Event Logs: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
    End
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Event Logs: Getting event logs' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Event Logs: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-RemoteAppliedGPOs
{
    <#
    .SYNOPSIS
       Gather applied GPO information from local or remote systems.
    .DESCRIPTION
       Gather applied GPO information from local or remote systems. Can utilize multiple runspaces and 
       alternate credentials.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information

    .EXAMPLE
       $a = Get-RemoteAppliedGPOs
       $a.AppliedGPOs | 
            Select Name,AppliedOrder |
            Sort-Object AppliedOrder
       
       Name                            appliedOrder
       ----                            ------------
       Local Group Policy                         1
       
       Description
       -----------
       Get all the locally applied GPO information then display them in their applied order.

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 09/01/2013
        - Initial release
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
       
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    BEGIN
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Applied GPOs: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Applied GPOs: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Applied GPOs: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Applied GPOs: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Applied GPOs: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Remote Applied GPOs: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                # General variables
                $GPOPolicies = @()
                $PSDateTime = Get-Date
                
                #region GPO Data

                $GPOQuery = Get-WmiObject @WMIHast `
                                          -Namespace "ROOT\RSOP\Computer" `
                                          -Class RSOP_GPLink `
                                          -Filter "AppliedOrder <> 0" |
                            Select @{n='linkOrder';e={$_.linkOrder}},
                                   @{n='appliedOrder';e={$_.appliedOrder}},
                                   @{n='GPO';e={$_.GPO.ToString().Replace("RSOP_GPO.","")}},
                                   @{n='Enabled';e={$_.Enabled}},
                                   @{n='noOverride';e={$_.noOverride}},
                                   @{n='SOM';e={[regex]::match( $_.SOM , '(?<=")(.+)(?=")' ).value}},
                                   @{n='somOrder';e={$_.somOrder}}
                foreach($GP in $GPOQuery)
                {
                    $AppliedPolicy = Get-WmiObject @WMIHast `
                                                   -Namespace 'ROOT\RSOP\Computer' `
                                                   -Class 'RSOP_GPO' -Filter $GP.GPO
                        $ObjectProp = @{
                            'Name' = $AppliedPolicy.Name
                            'GuidName' = $AppliedPolicy.GuidName
                            'ID' = $AppliedPolicy.ID
                            'linkOrder' = $GP.linkOrder
                            'appliedOrder' = $GP.appliedOrder
                            'Enabled' = $GP.Enabled
                            'noOverride' = $GP.noOverride
                            'SourceOU' = $GP.SOM
                            'somOrder' = $GP.somOrder
                        }
                        
                        $GPOPolicies += New-Object PSObject -Property $ObjectProp
                }
                          
                Write-Verbose -Message ('Remote Applied GPOs: Runspace {0}: Applied GPO information' -f $ComputerName)

                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','AppliedGPOs')
                $ResultProperty = @{
                    'PSComputerName' = $ComputerName
                    'PSDateTime' = $PSDateTime
                    'ComputerName' = $ComputerName
                    'AppliedGPOs' = $GPOPolicies
                }
                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty
                
                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.AppliedGPOs.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                #endregion GPO Data

                Write-Output -InputObject $ResultObject
            }
            catch
            {
                Write-Warning -Message ('Remote Applied GPOs: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Applied GPOs: Runspace {0}: End' -f $ComputerName)
        }
 
        function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Applied GPOs: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Applied GPOs: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Remote Applied GPOs: Getting info'
                        Status = 'Remote Applied GPOs: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Applied GPOs: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
    END
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Applied GPOs: Getting applied GPO information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Applied GPOs: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-RemoteServiceInformation
{
    <#
    .SYNOPSIS
       Gather remote service information.
    .DESCRIPTION
       Gather remote service information. Uses multiple runspaces and, if required, alternate credentials.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER ServiceName
       Specific service name to query.
    .PARAMETER IncludeDriverServices
       Include driver level services.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously.
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information.

    .EXAMPLE
       PS > Get-RemoteServiceInformation

       <output>
       
       Description
       -----------
       <Placeholder>

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 08/31/2013
        - Initial release
    #>
    [CmdletBinding()]
    PARAM
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter( ValueFromPipelineByPropertyName=$true,                    
                    ValueFromPipeline=$true,
                    HelpMessage="The service name to return." )]
        [Alias('Name')]
        [string[]]$ServiceName,
        
        [parameter( HelpMessage="Include the normally hidden driver services. Only applicable when not supplying a specific service name." )]
        [switch]
        $IncludeDriverServices,
        
        [parameter( HelpMessage="Optional WMI filter")]
        [string]
        $Filter,
        
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    BEGIN
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Service Information: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Remote Service Information: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Service Information: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Service Information: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Service Information: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Service Information: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter()]
                [string]
                $ComputerName,
                
                [Parameter()]                
                [string[]]
                $ServiceName,
                
                [parameter()]
                [switch]
                $IncludeDriverServices,
                
                [parameter()]
                [string]
                $Filter,
 
                [Parameter()]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Remote Service Information: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if ($ServiceName -ne $null)
                {
                    $WMIHast.Filter = "Name LIKE '$ServiceName'"
                }
                elseif ($Filter -ne $null)
                {
                    $WMIHast.Filter = $Filter
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                # General variables
                $ResultSet = @()
                $PSDateTime = Get-Date
                
                #region Services
                Write-Verbose -Message ('Remote Service Information: Runspace {0}: Service information' -f $ComputerName)

                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','Services')                                          
                                         
                $Services = @()
                $wmi_data = Get-WmiObject @WMIHast -Class Win32_Service
                foreach ($service in $wmi_data)
                {
                    $ServiceProperty = @{
                        'Name' = $service.Name
                        'DisplayName' = $service.DisplayName
                        'PathName' = $service.PathName
                        'Started' = $service.Started
                        'StartMode' = $service.StartMode
                        'State' = $service.State
                        'ServiceType' = $service.ServiceType
                        'StartName' = $service.StartName
                    }
                    $Services += New-Object PSObject -Property $ServiceProperty
                }
                if ($IncludeDriverServices)
                {
                    $wmi_data = Get-WmiObject @WMIHast -Class 'Win32_SystemDriver'
                    foreach ($service in $wmi_data)
                    {
                        $ServiceProperty = @{
                            'Name' = $service.Name
                            'DisplayName' = $service.DisplayName
                            'PathName' = $service.PathName
                            'Started' = $service.Started
                            'StartMode' = $service.StartMode
                            'State' = $service.State
                            'ServiceType' = $service.ServiceType
                            'StartName' = $service.StartName
                        }
                        $Services += New-Object PSObject -Property $ServiceProperty
                    }
                }
                $ResultProperty = @{
                    'PSComputerName' = $ComputerName
                    'PSDateTime' = $PSDateTime
                    'ComputerName' = $ComputerName
                    'Services' = $Services
                }
                
                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty
                    
                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.Services.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                
                $ResultSet += $ResultObject
            
                #endregion Services

                Write-Output -InputObject $ResultSet
            }
            catch
            {
                Write-Warning -Message ('Remote Service Information: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Service Information: Runspace {0}: End' -f $ComputerName)
        }
 
        Function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Service Information: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Service Information: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Remote Service Information: Getting info'
                        Status = 'Remote Service Information: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('ServiceName',$ServiceName)
            $null = $psCMD.AddParameter('IncludeDriverServices',$IncludeDriverServices)
            $null = $psCMD.AddParameter('Filter',$Filter)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Service Information: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
 
    END
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Service Information: Getting service information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Service Information: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-RemoteShareSessionInformation
{
    <#
    .SYNOPSIS
       Get share session information from remote or local host.
    .DESCRIPTION
       Get share session information from remote or local host. Uses multiple runspaces if 
       multiple hosts are processed.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information

    .EXAMPLE
       PS > Get-RemoteShareSessionInformation

       <output>
       
       Description
       -----------
       <Placeholder>

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 08/05/2013
        - Initial release
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
       
        [Parameter()]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter()]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter()]
        [switch]
        $ShowProgress,
        
        [Parameter()]
        [switch]
        $PromptForCredential,
        
        [Parameter()]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Share Information: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Share Session Information: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Share Session Information: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Share Session Information: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Share Session Information: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Share Session Information: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Share Session Information: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                # General variables
                $ResultSet = @()
                $PSDateTime = Get-Date
                
                #region ShareSessions
                Write-Verbose -Message ('Share Session Information: Runspace {0}: Share session information' -f $ComputerName)

                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','Sessions')                                          
                $WMI_ConnectionProps  = @('ShareName','UserName','RemoteComputerName')
                $SessionData = @()
                $wmi_connections = Get-WmiObject @WMIHast -Class Win32_ServerConnection | select $WMI_ConnectionProps
                foreach ($userSession in $wmi_connections)
                {
                    $SessionProperty = @{
                        'ShareName' = $userSession.ShareName
                        'UserName' = $userSession.UserName
                        'RemoteComputerName' = $userSession.ComputerName
                    }
                    $SessionData += New-Object -TypeName PSObject -Property $SessionProperty
                }
             
                $ResultProperty = @{
                    'PSComputerName' = $ComputerName
                    'PSDateTime' = $PSDateTime
                    'ComputerName' = $ComputerName
                    'Sessions' = $SessionData
                }

                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty
                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.ShareSession.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                
                $ResultSet += $ResultObject

                #endregion ShareSessions

                Write-Output -InputObject $ResultSet
            }
            catch
            {
                Write-Warning -Message ('Share Session Information: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Share Session Information: Runspace {0}: End' -f $ComputerName)
        }
 
        function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Share Session Information: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Share Session Information: Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Share Session Information: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Getting share session information'
                        Status = 'Share Session Information: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    Process
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            #$psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock).AddParameter('bgRunspaceID',$bgRunspaceCounter).AddParameter('ComputerName',$Computer).AddParameter('IncludeMemoryInfo',$IncludeMemoryInfo).AddParameter('IncludeDiskInfo',$IncludeDiskInfo).AddParameter('IncludeNetworkInfo',$IncludeNetworkInfo).AddParameter('Verbose',$Verbose)
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Share Session Information: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
 
    End
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Share Session Information: Getting share session information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Share Session Information: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-RemoteRegistryInformation
{
    <#
    .SYNOPSIS
       Retrieves registry subkey information.
    .DESCRIPTION
       Retrieves registry subkey information. All subkeys and their values are returned as a custom psobject. Optionally
       an array of psobjects can be returned which contain extra information like the registry key type,computer, and datetime.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER Hive
       Registry hive to retrieve from. By default this is 2147483650 (HKLM). Valid hives include:
          HKEY_CLASSES_ROOT = 2147483648
          HKEY_CURRENT_USER = 2147483649
          HKEY_LOCAL_MACHINE = 2147483650
          HKEY_USERS = 2147483651
          HKEY_CURRENT_CONFIG = 2147483653
          HKEY_DYN_DATA = 2147483654
    .PARAMETER Key
       Registry key to inspect (ie. SYSTEM\CurrentControlSet\Services\W32Time\Parameters)
    .PARAMETER AsHash
       Return a hash where the keys are the registry entries. This is only suitable for getting the regisrt
       values of one computer at a time.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information
    .EXAMPLE
       PS > $(Get-RemoteRegistryInformation -AsHash -Key "SYSTEM\CurrentControlSet\Services\W32Time\Parameters")['Type']

       NT5DS
       
       Description
       -----------
       Return the value of the 'Type' subkey within SYSTEM\CurrentControlSet\Services\W32Time\Parameters of
       HKLM.
       
    .EXAMPLE
       PS > $(Get-RemoteRegistryInformation -AsObject -Key "SYSTEM\CurrentControlSet\Services\W32Time\Parameters").Type

       NT5DS
       
       Description
       -----------
       Return the value of the 'Type' subkey within SYSTEM\CurrentControlSet\Services\W32Time\Parameters of
       HKLM from an object containing all registry keys in HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Parameters
       as individual object properties.
       
    .EXAMPLE
       PS > $b = Get-RemoteRegistryInformation -Key "SYSTEM\CurrentControlSet\Services\W32Time\Parameters"
       PS > $b.Registry | Select SubKey,SubKeyValue,SubKeyType
       
        SubKey                                         SubKeyValue                                    SubKeyType
        ------                                         -----------                                    ----------                                   
        ServiceDll                                     C:\Windows\system32\w32time.dll                REG_EXPAND_SZ
        ServiceMain                                    SvchostEntry_W32Time                           REG_SZ
        ServiceDllUnloadOnStop                         1                                              REG_DWORD
        Type                                           NT5DS                                          REG_SZ
        NtpServer                                                                                     REG_SZ
       
       Description
       -----------
       Return subkeys and their values as well as key types within SYSTEM\CurrentControlSet\Services\W32Time\Parameters of
       HKLM.

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.2 - 08/30/2013 
        - Changed AsArray option to be AsHash and restructured code to reflect this
        - Changed examples
        - Prefixed all warnings and verbose messages with function specific verbage
        - Forced STA apartement state before opening a runspace
       1.0.1 - 08/07/2013
        - Removed the explicit return of subkey values from output options
        - Fixed issue where only string values were returned
       1.0.0 - 08/06/2013
        - Initial release
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter( HelpMessage="Registry Hive (Default is HKLM)." )]
        [UInt32]
        $Hive = 2147483650,
        
        [Parameter( Mandatory=$true,
                    HelpMessage="Registry Key to inspect." )]
        [String]
        $Key,
        
        [Parameter(HelpMessage="Return a hash with key value pairs representing the registry being queried.")]
        [switch]
        $AsHash,
        
        [Parameter(HelpMessage="Return an object wherein the object properties are the registry keys and the property values are their value.")]
        [switch]
        $AsObject,
        
        [Parameter(HelpMessage="Maximum number of concurrent threads.")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information.")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function.")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials.")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials.")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    BEGIN
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Registry: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Remote Registry: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Registry: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Registry: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Registry: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Registry: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,

                [Parameter()]
                [UInt32]
                $Hive = 2147483650,
                
                [Parameter()]
                [String]
                $Key,
                
                [Parameter()]
                [switch]
                $AsHash,
                
                [Parameter()]
                [switch]
                $AsObject,
                
                [Parameter()]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            $regtype = @("Placeholder","REG_SZ","REG_EXPAND_SZ","REG_BINARY","REG_DWORD","Placeholder","Placeholder","REG_MULTI_SZ",`
                          "Placeholder","Placeholder","Placeholder","REG_QWORD")

            try
            {
                Write-Verbose -Message ('Remote Registry: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                # General variables
                $PSDateTime = Get-Date
                
                #region Registry
                Write-Verbose -Message ('Remote Registry: Runspace {0}: Gathering registry information' -f $ComputerName)

                # WMI data
                $wmi_data = Get-WmiObject @WMIHast -Class StdRegProv -Namespace 'root\default' -List:$true
                $allregkeys = $wmi_data.EnumValues($Hive,$Key)
                $allsubkeys = $wmi_data.EnumKey($Hive,$Key)

                $ResultHash = @{}
                $RegObjects = @() 
                $ResultObject = @{}
                       
                for ($i = 0; $i -lt $allregkeys.Types.Count; $i++) 
                {
                    switch ($allregkeys.Types[$i]) {
                            1 {
                                $keyvalue = ($wmi_data.GetStringValue($Hive,$Key,$allregkeys.sNames[$i])).sValue
                            }
                            2 {
                                $keyvalue = ($wmi_data.GetExpandedStringValue($Hive,$Key,$allregkeys.sNames[$i])).sValue
                            }
                            3 {
                                $keyvalue = ($wmi_data.GetBinaryValue($Hive,$Key,$allregkeys.sNames[$i])).uValue
                            }
                            4 {
                                $keyvalue = ($wmi_data.GetDWORDValue($Hive,$Key,$allregkeys.sNames[$i])).uValue
                            }
                               7 {
                                $keyvalue = ($wmi_data.GetMultiStringValue($Hive,$Key,$allregkeys.sNames[$i])).uValue
                            }
                            11 {
                                $keyvalue = ($wmi_data.GetQWORDValue($Hive,$Key,$allregkeys.sNames[$i])).sValue
                            }
                            default {
                                break
                            }
                    }
                    if ($AsHash -or $AsObject)
                    {
                        $ResultHash[$allregkeys.sNames[$i]] = $keyvalue
                    }
                    else
                    {
                        $RegProperties = @{
                            'Key' = $allregkeys.sNames[$i]
                            'KeyType' = $regtype[($allregkeys.Types[$i])]
                            'KeyValue' = $keyvalue
                        }
                        $RegObjects += New-Object PSObject -Property $RegProperties
                    }
                }
                foreach ($subkey in $allsubkeys.sNames) 
                {
                    if ($AsHash)
                    {
                        $ResultHash[$subkey] = ''
                    }
                    else
                    {
                        $RegProperties = @{
                            'Key' = $subkey
                            'KeyType' = 'SubKey'
                            'KeyValue' = ''
                        }
                        $RegObjects += New-Object PSObject -Property $RegProperties
                    }
                }
                if ($AsHash)
                {
                    $ResultHash
                }
                elseif ($AsObject)
                {
                    $ResultHash['PSComputerName'] = $ComputerName
                    $ResultObject = New-Object PSObject -Property $ResultHash
                    Write-Output -InputObject $ResultObject
                }
                else
                {
                    $ResultProperty = @{
                        'PSComputerName' = $ComputerName
                        'PSDateTime' = $PSDateTime
                        'ComputerName' = $ComputerName
                        'Registry' = $RegObjects
                    }
                    $Result = New-Object PSObject -Property $ResultProperty
                    Write-Output -InputObject $Result
                }
            }
            catch
            {
                Write-Warning -Message ('Remote Registry: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Registry: Runspace {0}: End' -f $ComputerName)
        }
 
        function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Registry: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Registry: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Getting asset info'
                        Status = '{0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Hive',$Hive)
            $null = $psCMD.AddParameter('Key',$Key)
            $null = $psCMD.AddParameter('AsHash',$AsHash)
            $null = $psCMD.AddParameter('AsObject',$AsObject)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
    END
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Registry: Getting registry information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Registry: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-RemoteProcessInformation
{
    <#
    .SYNOPSIS
       Get process information from remote machines.
    .DESCRIPTION
       Get process information from remote machines. use alternate credentials if desired. Filter by
       process name if desired.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER Process
       Optional process name to filter by
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information

    .EXAMPLE
       PS > (Get-RemoteProcessInformation -Process GoogleUpdate%).Processes | select Name

       name                                                                                                                     
       ----                                                                                                                     
       GoogleUpdate.exe 
       
       Description
       -----------
       Select all processes from the local machine with GoogleUpdate in the name then display the whole name.

    .NOTES
       Author: Zachary Loeber
       Site: http://zacharyloeber.com/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 09/01/2013
        - Initial release
    #>
    [CmdletBinding()]
    PARAM
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,

        [Parameter(HelpMessage="Process name")]
        [string]
        $Process='',
        
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )
    BEGIN
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Process Information: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Remote Process Information: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Process Information: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Process Information: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Process Information: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Process Information: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
                
                [Parameter(Position=1)]
                [string]
                $Process='',
 
                [Parameter(Position=2)]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Remote Process Information: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }
                if ($Process -ne '')
                {
                    $WMIHast.Filter = "Name LIKE '$Process'"
                }

                # General variables
                $PSDateTime = Get-Date
                
                #region Remote Process Information
                Write-Verbose -Message ('Remote Process Information: Runspace {0}: Process information' -f $ComputerName)

                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','Processes')

                # WMI data
                $wmi_processes = @(Get-WmiObject @WMIHast -Class Win32_Process)
                $ResultProperty = @{
                    'PSComputerName' = $ComputerName
                    'PSDateTime' = $PSDateTime
                    'ComputerName' = $ComputerName
                    'Processes' = $wmi_processes
                }
                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty
                    
                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.RemoteProcesses.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                
                $ResultSet += $ResultObject

                #endregion Remote Process Information

                Write-Output -InputObject $ResultSet
            }
            catch
            {
                Write-Warning -Message ('Remote Process Information: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Process Information: Runspace {0}: End' -f $ComputerName)
        }
 
        Function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Process Information: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Process Information: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Remote Process Information: Getting info'
                        Status = 'Remote Process Information: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Process',$Process)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Process Information: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
    END
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Process Information: Getting process information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Process Information: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}
#endregion Functions - Multiple Runspace

#region Functions - Serial or Utility
Function New-HTMLBarGraph
{
    <#
    .SYNOPSIS
       Creates an HTML fragment that looks like a horizontal bar graph when rendered.
    .DESCRIPTION
       Creates an HTML fragment that looks like a horizontal bar graph when rendered. Can be customized to use different
       characters for the left and right sides of the graph. Can also be customized to be a certain number of characters
       in size (Highly recommend sticking with even numbers)
    .PARAMETER LeftGraphChar
        HTML encoded character to use for the left part of the graph (the percentage used).
    .PARAMETER RightGraphChar
        HTML encoded character to use for the right part of the graph (the percentage unused).
    .PARAMETER GraphSize
        Overall character size of the graph
    .PARAMETER PercentageUsed
        The percentage of the graph which is "used" (the left side of the graph).
    .PARAMETER LeftColor
        The HTML color code for the left/used part of the graph.
    .PARAMETER RightColor
        The HTML color code for the right/unused part of the graph.
    .EXAMPLE
        PS> New-HTMLBarGraph -GraphSize 20 -PercentageUsed 10
        
        <Font Color=Red>&#9608;&#9608;</Font>
        <Font Color=Green>&#9608;&#9608;&#9608;&#9608;&#9608;&#9608;&#9608;
        &#9608;&#9608;&#9608;&#9608;&#9608;&#9608;&#9608;&#9608;&#9608;
        &#9608;&#9608;</Font>
        
    .NOTES
        Author: Zachary Loeber
        Site: http://zacharyloeber.com/
        Requires: Powershell 2.0
        Version History:
           1.0.0 - 08/10/2013
            - Initial release
            
        Some good characters to use for your graphs include:
        ▬  &#9644;
        ░  &#9617; {dither light}    
        ▒  &#9618; {dither medium}    
        ▓  &#9619; {dither heavy}    
        █  &#9608; {full box}
        
        Find more html character codes here: http://brucejohnson.ca/SpecialCharacters.html
        
        The default colors are not all that impressive. Used (left side of graph) is red and unused 
        (right side of the graph) is green. You use any colors which the font attribute will accept in html, 
        this includes transparent!
        
        If you are including this output in a larger table with other results remember that the 
        special characters will get converted and look all crappy after piping through convertto-html.
        To fix this issue, simply html decode the results like in this long winded example for memory
        utilization:
        
        $a = gwmi win32_operatingSystem | `
        select PSComputerName, @{n='Memory Usage';
                                 e={New-HTMLBarGraph -GraphSize 20 -PercentageUsed `
                                   (100 - [math]::Round((100 * ($_.FreePhysicalMemory)/`
                                                               ($_.TotalVisibleMemorySize))))}
                                }
        $Output = [System.Web.HttpUtility]::HtmlDecode(($a | ConvertTo-Html))
        $Output
        
        Props for original script from http://windowsmatters.com/index.php/category/powershell/page/2/
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="Character to use for left side of graph")]
        [string]$LeftGraphChar='&#9608;',
        
        [Parameter(HelpMessage="Character to use for right side of graph")]
        [string]$RightGraphChar='&#9608;',
        
        [Parameter(HelpMessage="Total size of the graph in character length. You really should stick with even numbers here.")]
        [int]$GraphSize=50,
        
        [Parameter(HelpMessage="Percentage for first part of graph")]
        [int]$PercentageUsed=50,
 
        [Parameter(HelpMessage="Color of left (used) graph segment.")]
        [string]$LeftColor='Red',

        [Parameter(HelpMessage="Color of left (unused) graph segment.")]
        [string]$RightColor='Green'
    )

    [int]$LeftSideCount = [Math]::Round((($PercentageUsed/100)*$GraphSize))
    [int]$RightSideCount = [Math]::Round((((100 - $PercentageUsed)/100)*$GraphSize))
    for ($index = 0; $index -lt $LeftSideCount; $index++) {
        $LeftSide = $LeftSide + $LeftGraphChar
    }
    for ($index = 0; $index -lt $RightSideCount; $index++) {
        $RightSide = $RightSide + $RightGraphChar
    }
    
    $Result = "<Font Color={0}>{1}</Font><Font Color={2}>{3}</Font>" `
               -f $LeftColor,$LeftSide,$RightColor,$RightSide

    Return $Result
}

Filter ConvertTo-KMG 
{
     <#
     .Synopsis
      Converts byte counts to Byte\KB\MB\GB\TB\PB format
     .DESCRIPTION
      Accepts an [int64] byte count, and converts to Byte\KB\MB\GB\TB\PB format
      with decimal precision of 2
     .EXAMPLE
     3000 | convertto-kmg
     #>

     $bytecount = $_
        switch ([math]::truncate([math]::log($bytecount,1024))) 
        {
                  0 {"$bytecount Bytes"}
                  1 {"{0:n2} KB" -f ($bytecount / 1kb)}
                  2 {"{0:n2} MB" -f ($bytecount / 1mb)}
                  3 {"{0:n2} GB" -f ($bytecount / 1gb)}
                  4 {"{0:n2} TB" -f ($bytecount / 1tb)}
            Default {"{0:n2} PB" -f ($bytecount / 1pb)}
          }
}

Function ConvertTo-PropertyValue 
{
    <#
    .SYNOPSIS
    Convert an object with various properties into an array of property, value pairs 
    
    .DESCRIPTION
    Convert an object with various properties into an array of property, value pairs

    If you output reports or other formats where a table with one long row is poorly formatted, this is a quick way to create a table of property value pairs.

    There are other ways you could do this.  For example, I could list all noteproperties from Get-Member results and return them.
    This function will keep properties in the same order they are provided, which can often be helpful for readability of results.

    .PARAMETER inputObject
    A single object to convert to an array of property value pairs.

    .PARAMETER leftheader
    Header for the left column.  Default:  Property

    .PARAMETER rightHeader
    Header for the right column.  Default:  Value

    .PARAMETER memberType
    Return only object members of this membertype.  Default:  Property, NoteProperty, ScriptProperty

    .EXAMPLE
    get-process powershell_ise | convertto-propertyvalue

    I want details on the powershell_ise process.
        With this command, if I output this to a table, a csv, etc. I will get a nice vertical listing of properties and their values
        Without this command, I get a long row with the same info

    .EXAMPLE
    #This example requires and demonstrates using the New-HTMLHead, New-HTMLTable, Add-HTMLTableColor, ConvertTo-PropertyValue and Close-HTML functions.
    
    #get processes to work with
        $processes = Get-Process
    
    #Build HTML header
        $HTML = New-HTMLHead -title "Process details"

    #Add CPU time section with top 10 PrivateMemorySize processes.  This example does not highlight any particular cells
        $HTML += "<h3>Process Private Memory Size</h3>"
        $HTML += New-HTMLTable -inputObject $($processes | sort PrivateMemorySize -Descending | select name, PrivateMemorySize -first 10)

    #Add Handles section with top 10 Handle usage.
    $handleHTML = New-HTMLTable -inputObject $($processes | sort handles -descending | select Name, Handles -first 10)

        #Add highlighted colors for Handle count
            
            #build hash table with parameters for Add-HTMLTableColor.  Argument and AttrValue will be modified each time we run this.
            $params = @{
                Column = "Handles" #I'm looking for cells in the Handles column
                ScriptBlock = {[double]$args[0] -gt [double]$args[1]} #I want to highlight if the cell (args 0) is greater than the argument parameter (arg 1)
                Attr = "Style" #This is the default, don't need to actually specify it here
            }

            #Add yellow, orange and red shading
            $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 1500 -attrValue "background-color:#FFFF99;" @params
            $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 2000 -attrValue "background-color:#FFCC66;" @params
            $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 3000 -attrValue "background-color:#FFCC99;" @params
      
        #Add title and table
        $HTML += "<h3>Process Handles</h3>"
        $HTML += $handleHTML

    #Add process list containing first 10 processes listed by get-process.  This example does not highlight any particular cells
        $HTML += New-HTMLTable -inputObject $($processes | select name -first 10 ) -listTableHead "Random Process Names"

    #Add property value table showing details for PowerShell ISE
        $HTML += "<h3>PowerShell Process Details PropertyValue table</h3>"
        $processDetails = Get-process powershell_ise | select name, id, cpu, handles, workingset, PrivateMemorySize, Path -first 1
        $HTML += New-HTMLTable -inputObject $(ConvertTo-PropertyValue -inputObject $processDetails)

    #Add same PowerShell ISE details but not in property value form.  Close the HTML
        $HTML += "<h3>PowerShell Process Details object</h3>"
        $HTML += New-HTMLTable -inputObject $processDetails | Close-HTML

    #write the HTML to a file and open it up for viewing
        set-content C:\test.htm $HTML
        & 'C:\Program Files\Internet Explorer\iexplore.exe' C:\test.htm

    .FUNCTIONALITY
    General Command
    #> 
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromRemainingArguments=$false)]
        [PSObject]$InputObject,
        
        [validateset("AliasProperty", "CodeProperty", "Property", "NoteProperty", "ScriptProperty",
            "Properties", "PropertySet", "Method", "CodeMethod", "ScriptMethod", "Methods",
            "ParameterizedProperty", "MemberSet", "Event", "Dynamic", "All")]
        [string[]]$memberType = @( "NoteProperty", "Property", "ScriptProperty" ),
            
        [string]$leftHeader = "Property",
            
        [string]$rightHeader = "Value"
    )

    begin{
        #init array to dump all objects into
        $allObjects = @()

    }
    process{
        #if we're taking from pipeline and get more than one object, this will build up an array
        $allObjects += $inputObject
    }

    end{
        #use only the first object provided
        $allObjects = $allObjects[0]

        #Get properties.  Filter by memberType.
        $properties = $allObjects.psobject.properties | ?{$memberType -contains $_.memberType} | select -ExpandProperty Name

        #loop through properties and display property value pairs
        foreach($property in $properties){

            #Create object with property and value
            $temp = "" | select $leftHeader, $rightHeader
            $temp.$leftHeader = $property.replace('"',"")
            $temp.$rightHeader = try { $allObjects | select -ExpandProperty $temp.$leftHeader -erroraction SilentlyContinue } catch { $null }
            $temp
        }
    }
}

Function ConvertTo-HashArray
{
    <#
    .SYNOPSIS
    Convert an array of objects to a hash table based on a single property of the array. 
    
    .DESCRIPTION
    Convert an array of objects to a hash table based on a single property of the array.
    
    .PARAMETER InputObject
    An array of objects to convert to a hash table array.

    .PARAMETER PivotProperty
    The property to use as the key value in the resulting hash.

    .EXAMPLE
    <Placeholder>

    Description
    -----------
    <Placeholder>
    
    .NOTES

    #> 
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [AllowEmptyCollection()]
        [PSObject[]]
        $InputObject,
        
        [Parameter(Mandatory=$true)]
        [string]$PivotProperty
    )

    BEGIN
    {
        #init array to dump all objects into
        $allObjects = @()
        $Results = @{}
    }
    PROCESS
    {
        #if we're taking from pipeline and get more than one object, this will build up an array
        $allObjects += $inputObject
    }

    END
    {
        ForEach ($object in $allObjects)
        {
            if ($object -ne $null)
            {
                #$object
                if ($object.PSObject.Properties.Match($PivotProperty).Count) 
                {
                    $Results[$object.$PivotProperty] = $object
                }
            }
        }
        $Results
    }
}

Function ConvertTo-ProductKey
{
    <#   
    .SYNOPSIS   
        Converts registry key value to windows product key.
         
    .DESCRIPTION   
        Converts registry key value to windows product key. Specifically the following keys:
            SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId
            SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId4
        
    .PARAMETER Registry
        Either DigitalProductId or DigitalProductId4 (as described in the description)
         
    .NOTES   
        Author: Zachary Loeber
        Original Author: Boe Prox
        Version: 1.0
         - Took the registry setting retrieval portion from Boe's original script and converted it
           to this basic conversion function. This is to be used in conjunction with my other
           function, get-remoteregistryinformation
     
    .EXAMPLE 
     PS > $reg_ProductKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
     PS > $a = Get-RemoteRegistryInformation -Key $reg_ProductKey -AsObject
     PS > ConvertTo-ProductKey $a.DigitalProductId
     
            XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
            
     PS > ConvertTo-ProductKey $a.DigitalProductId4 -x64
     
            XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
         
        Description 
        ----------- 
        Retrieves the product key information from the local machine and converts it to a readible format.
    #>      
    [cmdletbinding()]
    Param (
        [parameter(Mandatory=$True,Position=0)]
        $Registry,
        [parameter()]
        [Switch]$x64
    )
    Begin 
    {
        $map="BCDFGHJKMPQRTVWXY2346789" 
    }
    Process 
    {
        $ProductKey = ""

        $prodkey = $Registry[0x34..0x42]

        for ($i = 24; $i -ge 0; $i--) 
        { 
            $r = 0 
            for ($j = 14; $j -ge 0; $j--) 
            {
                $r = ($r * 256) -bxor $prodkey[$j] 
                $prodkey[$j] = [math]::Floor([double]($r/24)) 
                $r = $r % 24 
            } 
            $ProductKey = $map[$r] + $ProductKey 
            if (($i % 5) -eq 0 -and $i -ne 0)
            { 
                $ProductKey = "-" + $ProductKey
            }
        }
        $ProductKey
    }
}

Function ConvertTo-PSObject
{
    <# 
     Take an array of like psobject and convert it to a singular psobject based on two shared
     properties across all psobjects in the array.
     Example Input object: 
    $obj = @()
    $a = @{ 
        'PropName' = 'Property 1'
        'Val1' = 'Value 1'
        }
    $b = @{ 
        'PropName' = 'Property 2'
        'Val1' = 'Value 2'
        }
    $obj += new-object psobject -property $a
    $obj += new-object psobject -property $b

    $c = $obj | ConvertTo-PSObject -propname 'PropName' -valname 'Val1'
    $c.'Property 1'
    Value 1
    #>
    [cmdletbinding()]
    PARAM(
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [PSObject[]]$InputObject,
        [string]$propname,
        [string]$valname
    )

    BEGIN
    {
        #init array to dump all objects into
        $allObjects = @()
    }
    PROCESS
    {
        #if we're taking from pipeline and get more than one object, this will build up an array
        $allObjects += $inputObject
    }
    END
    {
        $returnobject = New-Object psobject
        foreach ($obj in $allObjects)
        {
            if ($obj.$propname -ne $null)
            {
                $returnobject | Add-Member -MemberType NoteProperty -Name $obj.$propname -Value $obj.$valname
            }
        }
        $returnobject
    }
}

Function ConvertTo-MultiArray 
{
 <#
 .Notes
 NAME: ConvertTo-MultiArray
 AUTHOR: Tome Tanasovski
 Website: http://powertoe.wordpress.com
 Twitter: http://twitter.com/toenuff
 Version: 1.0
 CREATED: 11/5/2010
 LASTEDIT:
 11/5/2010 1.0
 Initial Release
 11/5/2010 1.1
 Removed array parameter and passes a reference to the multi-dimensional array as output to the cmdlet
 11/5/2010 1.2
 Modified all rows to ensure they are entered as string values including $null values as a blank ("") string.

 .Synopsis
 Converts a collection of PowerShell objects into a multi-dimensional array

 .Description
 Converts a collection of PowerShell objects into a multi-dimensional array.  The first row of the array contains the property names.  Each additional row contains the values for each object.

 This cmdlet was created to act as an intermediary to importing PowerShell objects into a range of cells in Exchange.  By using a multi-dimensional array you can greatly speed up the process of adding data to Excel through the Excel COM objects.

 .Parameter InputObject
 Specifies the objects to export into the multi dimensional array.  Enter a variable that contains the objects or type a command or expression that gets the objects. You can also pipe objects to ConvertTo-MultiArray.

 .Inputs
 System.Management.Automation.PSObject
        You can pipe any .NET Framework object to ConvertTo-MultiArray

 .Outputs
 [ref]
        The cmdlet will return a reference to the multi-dimensional array.  To access the array itself you will need to use the Value property of the reference

 .Example
 $arrayref = get-process |Convertto-MultiArray

 .Example
 $dir = Get-ChildItem c:\
 $arrayref = Convertto-MultiArray -InputObject $dir

 .Example
 $range.value2 = (ConvertTo-MultiArray (get-process)).value

 .LINK

http://powertoe.wordpress.com

#>
    param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [PSObject[]]$InputObject
    )
    BEGIN {
        $objects = @()
        [ref]$array = [ref]$null
    }
    Process {
        $objects += $InputObject
    }
    END {
        $properties = $objects[0].psobject.properties |%{$_.name}
        $array.Value = New-Object 'object[,]' ($objects.Count+1),$properties.count
        # i = row and j = column
        $j = 0
        $properties |%{
            $array.Value[0,$j] = $_.tostring()
            $j++
        }
        $i = 1
        $objects |% {
            $item = $_
            $j = 0
            $properties | % {
                if ($item.($_) -eq $null) {
                    $array.value[$i,$j] = ""
                }
                else {
                    $array.value[$i,$j] = $item.($_).tostring()
                }
                $j++
            }
            $i++
        }
        $array
    }
}

Function Get-DellWarranty
{
    <# 
    .Synopsis 
       Get Warranty Info for Dell Computer 
    .DESCRIPTION 
       This takes a Computer Name, returns the ST of the computer, 
       connects to Dell's SOAP Service and returns warranty info and 
       related information. If computer is offline, no action performed. 
       ST is pulled via WMI. 
    .EXAMPLE 
       get-dellwarranty -Name bob, client1, client2 | ft -AutoSize 
        WARNING: bob is offline 
     
        ComputerName ServiceLevel  EndDate   StartDate DaysLeft ServiceTag Type                       Model ShipDate  
        ------------ ------------  -------   --------- -------- ---------- ----                       ----- --------  
        client1      C, NBD ONSITE 2/22/2017 2/23/2014     1095 7GH6SX1    Dell Precision WorkStation T1650 2/22/2013 
        client2      C, NBD ONSITE 7/16/2014 7/16/2011      334 74N5LV1    Dell Precision WorkStation T3500 7/15/2010 
    .EXAMPLE 
        Get-ADComputer -Filter * -SearchBase "OU=Exchange 2010,OU=Member Servers,DC=Contoso,DC=com" | get-dellwarranty | ft -AutoSize 
     
        ComputerName ServiceLevel            EndDate   StartDate DaysLeft ServiceTag Type      Model ShipDate  
        ------------ ------------            -------   --------- -------- ---------- ----      ----- --------  
        MAIL02       P, Gold or ProMCritical 4/26/2016 4/25/2011      984 CGWRNQ1    PowerEdge M905  4/25/2011 
        MAIL01       P, Gold or ProMCritical 4/26/2016 4/25/2011      984 DGWRNQ1    PowerEdge M905  4/25/2011 
        DAG          P, Gold or ProMCritical 4/26/2016 4/25/2011      984 CGWRNQ1    PowerEdge M905  4/25/2011 
        MAIL         P, Gold or ProMCritical 4/26/2016 4/25/2011      984 CGWRNQ1    PowerEdge M905  4/25/2011 
    .EXAMPLE 
        get-dellwarranty -ServiceTag CGABCQ1,DGEFGQ1 | ft  -AutoSize 
     
        ServiceLevel            EndDate   StartDate DaysLeft ServiceTag Type      Model ShipDate  
        ------------            -------   --------- -------- ---------- ----      ----- --------  
        P, Gold or ProMCritical 4/26/2016 4/25/2011      984 CGABCQ1    PowerEdge M905  4/25/2011 
        P, Gold or ProMCritical 4/26/2016 4/25/2011      984 DGEFGQ1    PowerEdge M905  4/25/201 
    .INPUTS 
       Name(ComputerName), ServiceTag 
    .OUTPUTS 
       System.Object 
    .NOTES 
       General notes 
    #> 
    [CmdletBinding()] 
    [OutputType([System.Object])] 
    Param( 
        # Name should be a valid computer name or IP address. 
        [Parameter(Mandatory=$False,  
                   ValueFromPipeline=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
         
        [Alias('HostName', 'Identity', 'DNSHostName', 'Name')] 
        [string[]]$ComputerName=$env:COMPUTERNAME, 
         
        [Parameter()] 
        [string[]]$ServiceTag = $null,
     
        [Parameter()]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    ) 
 
    BEGIN
    {
        $wmisplat = @{}
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try 
            {
                [net.dns]::GetHostByAddress($_)
            } catch 
            {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
        if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
        {
            $wmisplat.Credential = $Credential
        }
    } 
    PROCESS
    {
        if($ServiceTag -eq $null)
        { 
            foreach($C in $ComputerName)
            { 
                $test = Test-Connection -ComputerName $c -Count 1 -Quiet 
                if($test -eq $true)
                {
                    try
                    {
                        $system = Get-WmiObject @wmisplat -ComputerName $C win32_bios -ErrorAction SilentlyContinue                

                        if ($system.Manufacturer -like "Dell*")
                        {
                            try
                            {
                                $service = New-WebServiceProxy -Uri http://143.166.84.118/services/assetservice.asmx?WSDL -ErrorAction Stop
                                $serial =  $system.serialnumber 
                                $guid = [guid]::NewGuid() 
                                $info = $service.GetAssetInformation($guid,'check_warranty.ps1',$serial) 
                                if ($info -ne $null)
                                {
                                    $Result=@{ 
                                        'ComputerName'=$c 
                                        'ServiceLevel'=$info[0].Entitlements[0].ServiceLevelDescription.ToString() 
                                        'EndDate'=$info[0].Entitlements[0].EndDate.ToShortDateString() 
                                        'StartDate'=$info[0].Entitlements[0].StartDate.ToShortDateString() 
                                        'DaysLeft'=$info[0].Entitlements[0].DaysLeft 
                                        'ServiceTag'=$info[0].AssetHeaderData.ServiceTag 
                                        'Type'=$info[0].AssetHeaderData.SystemType 
                                        'Model'=$info[0].AssetHeaderData.SystemModel 
                                        'ShipDate'=$info[0].AssetHeaderData.SystemShipDate.ToShortDateString() 
                                    } 
                                 
                                    $obj = New-Object -TypeName psobject -Property $result 
                                    Write-Output $obj 
                                }
                                else
                                {
                                    Write-Warning -Message ('{0}: No warranty information returned' -f $C)
                                }
                            }
                            catch
                            {
                                Write-Warning -Message ('{0}: Unable to connect to web service' -f $C)
                            }
                        }
                        else
                        {
                            Write-Warning -Message ('{0}: Not a Dell computer' -f $C)
                        }
                    }
                    catch
                    {
                        Write-Warning -Message ('{0}: Not able to gather service tag' -f $C)
                    }
                }  
                else
                { 
                    Write-Warning -Message ('{0}: System is offline' -f $C)
                }         
 
            } 
        } 
        else
        { 
            foreach($s in $ServiceTag)
            {
                try
                {
                    $service = New-WebServiceProxy -Uri http://143.166.84.118/services/assetservice.asmx?WSDL -ErrorAction Stop
                    $guid = [guid]::NewGuid() 
                    $info = $service.GetAssetInformation($guid,'check_warranty.ps1',$S) 
                     
                    if($info)
                    { 
                        $Result=@{ 
                            'ServiceLevel'=$info[0].Entitlements[0].ServiceLevelDescription.ToString() 
                            'EndDate'=$info[0].Entitlements[0].EndDate.ToShortDateString() 
                            'StartDate'=$info[0].Entitlements[0].StartDate.ToShortDateString() 
                            'DaysLeft'=$info[0].Entitlements[0].DaysLeft 
                            'ServiceTag'=$info[0].AssetHeaderData.ServiceTag 
                            'Type'=$info[0].AssetHeaderData.SystemType 
                            'Model'=$info[0].AssetHeaderData.SystemModel 
                            'ShipDate'=$info[0].AssetHeaderData.SystemShipDate.ToShortDateString() 
                        } 
                    } 
                    else
                    { 
                        Write-Warning "$S is not a valid Dell Service Tag." 
                    } 

                    $obj = New-Object -TypeName psobject -Property $result 
                    Write-Output $obj 
                }
               catch
               {
                    Write-Warning -Message ('{0}: Unable to connect to web service' -f $C)
               }
            }
        } 
    } 
    END
    { 
    } 
}

# I use this to normalize for comparisons
#Function ConvertTo-MB 
#{
#    [CmdletBinding()]
#    param(
#        [parameter()]
#        [string]$val
#    )
#    switch ([Regex]::Match($val, "(?s)(?<=(.+\ ))(.+)").Value) {
#        'B' {
#            [float](0)
#        }
#        'KB' {
#            ([float]([Regex]::Match($val, "(?s)(.+)(?=(\ (.+)))").Value) / 11024)
#        }
#        'MB' {
#            [float]([Regex]::Match($val, "(?s)(.+)(?=(\ (.+)))").Value)
#        }
#        'GB' {
#            ([float]([Regex]::Match($val, "(?s)(.+)(?=(\ (.+)))").Value) * 1024)
#        }
#        'TB' {
#            ([float]([Regex]::Match($val, "(?s)(.+)(?=(\ (.+)))").Value) * 1048576)
#        }
#        default {
#            [float]([Regex]::Match($val, "(?s)(.+)(?=(\ (.+)))").Value)
#        }
#    }
#}

Filter ConvertTo-MB 
{
    <#
    .SYNOPSIS
    Converts KB\MB\GB\TB\PB format to MB float
    .DESCRIPTION
    Accepts an [int64] byte count, and converts to Byte\KB\MB\GB\TB\PB format
    with decimal precision of 2
    .EXAMPLE
    '3000 MB' | ConvertTo-MB
    .EXAMPLE
    '123 Tb' | ConvertTo-MB
    #>
    $val = $_
    switch ([Regex]::Match($val, "(?s)(?<=(.+\ ))(.+)").Value) {
        'B'  {([float](0))}
        'KB' {([float]([Regex]::Match($val, "(?s)(.+)(?=(\ (.+)))").Value) / 11024)}
        'MB' {([float]([Regex]::Match($val, "(?s)(.+)(?=(\ (.+)))").Value))}
        'GB' {([float]([Regex]::Match($val, "(?s)(.+)(?=(\ (.+)))").Value) * 1024)}
        'TB' {([float]([Regex]::Match($val, "(?s)(.+)(?=(\ (.+)))").Value) * 1048576)}
     default {([float]([Regex]::Match($val, "(?s)(.+)(?=(\ (.+)))").Value))}
    }
}

Function Colorize-Table 
{
<# 
.SYNOPSIS 
Colorize-Table 
 
.DESCRIPTION 
Create an html table and colorize individual cells or rows of an array of objects based on row header and value. Optionally, you can also
modify an existing html document or change only the styles of even or odd rows.
 
.PARAMETER  InputObject 
An array of objects (ie. (Get-process | select Name,Company) 
 
.PARAMETER  Column 
The column you want to modify. (Note: If the parameter ColorizeMethod is not set to ByValue the 
Column parameter is ignored)

.PARAMETER ScriptBlock
Used to perform custom cell evaluations such as -gt -lt or anything else you need to check for in a
table cell element. The scriptblock must return either $true or $false and is, by default, just
a basic -eq comparisson. You must use the variables as they are used in the following example.
(Note: If the parameter ColorizeMethod is not set to ByValue the ScriptBlock parameter is ignored)

[scriptblock]$scriptblock = {[int]$args[0] -gt [int]$args[1]}

$args[0] will be the cell value in the table
$args[1] will be the value to compare it to

Strong typesetting is encouraged for accuracy.

.PARAMETER  ColumnValue 
The column value you will modify if ScriptBlock returns a true result. (Note: If the parameter 
ColorizeMethod is not set to ByValue the ColumnValue parameter is ignored)
 
.PARAMETER  Attr 
The attribute to change should ColumnValue be found in the Column specified. 
- A good example is using "style" 
 
.PARAMETER  AttrValue 
The attribute value to set when the ColumnValue is found in the Column specified 
- A good example is using "background: red;" 
 
.EXAMPLE 
This will highlight the process name of Dropbox with a red background. 

$TableStyle = @'
<title>Process Report</title> 
    <style>             
    BODY{font-family: Arial; font-size: 8pt;} 
    H1{font-size: 16px;} 
    H2{font-size: 14px;} 
    H3{font-size: 12px;} 
    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;} 
    TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;} 
    TD{border: 1px solid black; padding: 5px;} 
    </style>
'@

$tabletocolorize = $(Get-Process | ConvertTo-Html -Head $TableStyle) 
$colorizedtable = Colorize-Table $tabletocolorize -Column "Name" -ColumnValue "Dropbox" -Attr "style" -AttrValue "background: red;"
$colorizedtable | Out-File "$pwd/testreport.html" 
ii "$pwd/testreport.html"

You can also strip out just the table at the end if you are working with multiple tables in your report:
if ($colorizedtable -match '(?s)<table>(.*)</table>')
{
    $result = $matches[0]
}

.EXAMPLE 
Using the same $TableStyle variable above this will create a table of top 5 processes by memory usage,
color the background of a whole row yellow for any process using over 150Mb and red if over 400Mb.

$tabletocolorize = $(get-process | select -Property ProcessName,Company,@{Name="Memory";Expression={[math]::truncate($_.WS/ 1Mb)}} | Sort-Object Memory -Descending | Select -First 5 ) 

[scriptblock]$scriptblock = {[int]$args[0] -gt [int]$args[1]}
$testreport = Colorize-Table $tabletocolorize -Column "Memory" -ColumnValue 150 -Attr "style" -AttrValue "background:yellow;" -ScriptBlock $ScriptBlock -HTMLHead $TableStyle -WholeRow $true
$testreport = Colorize-Table $testreport -Column "Memory" -ColumnValue 400 -Attr "style" -AttrValue "background:red;" -ScriptBlock $ScriptBlock -WholeRow $true
$testreport | Out-File "$pwd/testreport.html" 
ii "$pwd/testreport.html"

.NOTES 
If you are going to convert something to html with convertto-html in powershell v2 there is a bug where the  
header will show up as an asterick if you only are converting one object property. 

This script is a modification of something I found by some rockstar named Jaykul at this site
http://stackoverflow.com/questions/4559233/technique-for-selectively-formatting-data-in-a-powershell-pipeline-and-output-as

I believe that .Net 3.5 is a requirement for using the Linq libraries

.LINK 
http://zacharyloeber.com 
#> 
[CmdletBinding(DefaultParameterSetName = "ObjectSet")] 
param ( 
    [Parameter( Position=0,
                Mandatory=$true, 
                ValueFromPipeline=$true, 
                ParameterSetName="ObjectSet")]
    [PSObject[]]$InputObject, 
    [Parameter( Position=0, 
                Mandatory=$true, 
                ValueFromPipeline=$true, 
                ParameterSetName="StringSet")] 
    [String[]]$InputString='', 
    [Parameter( Mandatory=$false, 
                ValueFromPipeline=$false)]
    [String]$Column="Name", 
    [Parameter( Mandatory=$false, 
                ValueFromPipeline=$false)]
    $ColumnValue=0,
    [Parameter( Mandatory=$false, 
                ValueFromPipeline=$false)]
    [ScriptBlock]$ScriptBlock = {[string]$args[0] -eq [string]$args[1]}, 
    [Parameter( Mandatory=$true, 
                ValueFromPipeline=$false)] 
    [String]$Attr, 
    [Parameter( Mandatory=$true, 
                ValueFromPipeline=$false)] 
    [String]$AttrValue, 
    [Parameter( Mandatory=$false, 
                ValueFromPipeline=$false)] 
    [Bool]$WholeRow=$false, 
    [Parameter( Mandatory=$false, 
                ValueFromPipeline=$false, 
                ParameterSetName="ObjectSet")]
    [String]$HTMLHead='<title>HTML Table</title>',
    [Parameter( Mandatory=$false, 
                ValueFromPipeline=$false)]
    [ValidateSet('ByValue','ByEvenRows','ByOddRows')]
    [String]$ColorizeMethod='ByValue'
    )
    
BEGIN 
{ 
    # A little note on Add-Type, this adds in the assemblies for linq with some custom code. The first time this 
    # is run in your powershell session it is compiled and loaded into your session. If you run it again in the same
    # session and the code was not changed at all powershell skips the command (otherwise recompiling code each time
    # the function is called in a session would be pretty ineffective so this is by design). If you make any changes
    # to the code, even changing one space or tab, it is detected as new code and will try to reload the same namespace
    # which is not allowed and will cause an error. So if you are debugging this or changing it up, either change the
    # namespace as well or exit and restart your powershell session.
    #
    # And some notes on the actual code. It is my first jump into linq (or C# for that matter) so if it looks not so 
    # elegant or there is a better way to do this I'm all ears. I define four methods which names are self-explanitory:
    # - GetElementByIndex
    # - GetElementByValue
    # - GetOddElements
    # - GetEvenElements
    $LinqCode = @"
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByIndex(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, int index)
    {
        return doc.Descendants(element)
                .Where  (e => e.NodesBeforeSelf().Count() == index)
                .Select (e => e);
    }
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByValue(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, string value)
    {
        return  doc.Descendants(element) 
                .Where  (e => e.Value == value)
                .Select (e => e);
    }
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetOddElements(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element)
    {
        return doc.Descendants(element)
                .Where  ((e,i) => i % 2 != 0)
                .Select (e => e);
    }
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetEvenElements(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element)
    {
        return doc.Descendants(element)
                .Where  ((e,i) => i % 2 == 0)
                .Select (e => e);
    }
"@

    Add-Type -ErrorAction SilentlyContinue -Language CSharpVersion3 `
    -ReferencedAssemblies System.Xml, System.Xml.Linq `
    -UsingNamespace System.Linq `
    -Name XUtilities `
    -Namespace Huddled `
    -MemberDefinition $LinqCode
    
    $Objects = @() 
} 
 
PROCESS 
{ 
    # Handle passing object via pipe 
    If ($PSBoundParameters.ContainsKey('InputObject')) {
        $Objects += $InputObject 
    }
} 
 
END 
{ 
    # Convert our data to x(ht)ml 
    if ($InputString)    # If a string was passed just parse it 
    {   
        $xml = [System.Xml.Linq.XDocument]::Parse("$InputString")  
    } 
    else    # Otherwise we have to convert it to html first 
    { 
        $xml = [System.Xml.Linq.XDocument]::Parse("$($Objects | ConvertTo-Html -Head $HTMLHead)")     
    } 
    
    switch ($ColorizeMethod) {
        "ByEvenRows" {
            $evenrows = [Huddled.XUtilities]::GetEvenElements($xml, "{http://www.w3.org/1999/xhtml}tr")    
            foreach ($row in $evenrows)
            {
                $row.SetAttributeValue($Attr, $AttrValue)
            }            
        
        }
        "ByOddRows" {
            $oddrows = [Huddled.XUtilities]::GetOddElements($xml, "{http://www.w3.org/1999/xhtml}tr")    
            foreach ($row in $oddrows)
            {
                $row.SetAttributeValue($Attr, $AttrValue)
            }
        }
        "ByValue" {
            # Find the index of the column you want to format 
            $ColumnLoc = [Huddled.XUtilities]::GetElementByValue($xml, "{http://www.w3.org/1999/xhtml}th",$Column) 
            $ColumnIndex = $ColumnLoc | Foreach-Object{($_.NodesBeforeSelf() | Measure-Object).Count} 
    
            # Process each xml element based on the index for the column we are highlighting 
            switch([Huddled.XUtilities]::GetElementByIndex($xml, "{http://www.w3.org/1999/xhtml}td", $ColumnIndex)) 
            { 
                {$(Invoke-Command $ScriptBlock -ArgumentList @($_.Value, $ColumnValue))} {
                    if ($WholeRow)
                    {
                        $_.Parent.SetAttributeValue($Attr, $AttrValue)
                    }
                    else
                    {
                        $_.SetAttributeValue($Attr, $AttrValue)
                    }
                }
            }
        }
    }

    Return $xml.Document.ToString()
}
}

Function Get-OUResults
{
    #----------------------------------------------
    #region Import Assemblies
    #----------------------------------------------
    [void][Reflection.Assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][Reflection.Assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][Reflection.Assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    [void][Reflection.Assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][Reflection.Assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][Reflection.Assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][Reflection.Assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    #endregion Import Assemblies

    #Define a Param block to use custom parameters in the project
    #Param ($CustomParameter)

    function Main {
        Param ([String]$Commandline)
        if((Call-MainForm_pff) -eq "OK")
        {
            
        }
        
        $global:ExitCode = 0 #Set the exit code for the Packager
    }
    #endregion Source: Startup.pfs

    #region Source: MainForm.pff
    function Call-MainForm_pff
    {
        #----------------------------------------------
        #region Import the Assemblies
        #----------------------------------------------
        [void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
        [void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
        [void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
        [void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
        [void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
        [void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
        [void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
        [void][reflection.assembly]::Load("System.Windows.Forms.DataVisualization, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
        #endregion Import Assemblies

        #----------------------------------------------
        #region Generated Form Objects
        #----------------------------------------------
        [System.Windows.Forms.Application]::EnableVisualStyles()
        $MainForm = New-Object 'System.Windows.Forms.Form'
        $buttonOK = New-Object 'System.Windows.Forms.Button'
        $buttonLoadOU = New-Object 'System.Windows.Forms.Button'
        $groupbox1 = New-Object 'System.Windows.Forms.GroupBox'
        $radiobuttonDomainControllers = New-Object 'System.Windows.Forms.RadioButton'
        $radiobuttonWorkstations = New-Object 'System.Windows.Forms.RadioButton'
        $radiobuttonServers = New-Object 'System.Windows.Forms.RadioButton'
        $radiobuttonAll = New-Object 'System.Windows.Forms.RadioButton'
        $listboxComputers = New-Object 'System.Windows.Forms.ListBox'
        $labelOrganizationalUnit = New-Object 'System.Windows.Forms.Label'
        $txtOU = New-Object 'System.Windows.Forms.TextBox'
        $btnSelectOU = New-Object 'System.Windows.Forms.Button'
        $timerFadeIn = New-Object 'System.Windows.Forms.Timer'
        $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
        #endregion Generated Form Objects

        #----------------------------------------------
        # User Generated Script
        #----------------------------------------------
        
        $OnLoadFormEvent={
            $Results = @()
        }
        
        $form1_FadeInLoad={
            #Start the Timer to Fade In
            $timerFadeIn.Start()
            $MainForm.Opacity = 0
        }
        
        $timerFadeIn_Tick={
            #Can you see me now?
            if($MainForm.Opacity -lt 1)
            {
                $MainForm.Opacity += 0.1
                
                if($MainForm.Opacity -ge 1)
                {
                    #Stop the timer once we are 100% visible
                    $timerFadeIn.Stop()
                }
            }
        }
        
        #region Control Helper Functions
        function Load-ListBox 
        {
        <#
            .SYNOPSIS
                This functions helps you load items into a ListBox or CheckedListBox.
        
            .DESCRIPTION
                Use this function to dynamically load items into the ListBox control.
        
            .PARAMETER  ListBox
                The ListBox control you want to add items to.
        
            .PARAMETER  Items
                The object or objects you wish to load into the ListBox's Items collection.
        
            .PARAMETER  DisplayMember
                Indicates the property to display for the items in this control.
            
            .PARAMETER  Append
                Adds the item(s) to the ListBox without clearing the Items collection.
            
            .EXAMPLE
                Load-ListBox $ListBox1 "Red", "White", "Blue"
            
            .EXAMPLE
                Load-ListBox $listBox1 "Red" -Append
                Load-ListBox $listBox1 "White" -Append
                Load-ListBox $listBox1 "Blue" -Append
            
            .EXAMPLE
                Load-ListBox $listBox1 (Get-Process) "ProcessName"
        #>
            Param (
                [ValidateNotNull()]
                [Parameter(Mandatory=$true)]
                [System.Windows.Forms.ListBox]$ListBox,
                [ValidateNotNull()]
                [Parameter(Mandatory=$true)]
                $Items,
                [Parameter(Mandatory=$false)]
                [string]$DisplayMember,
                [switch]$Append
            )
            
            if(-not $Append)
            {
                $listBox.Items.Clear()    
            }
            
            if($Items -is [System.Windows.Forms.ListBox+ObjectCollection])
            {
                $listBox.Items.AddRange($Items)
            }
            elseif ($Items -is [Array])
            {
                $listBox.BeginUpdate()
                foreach($obj in $Items)
                {
                    $listBox.Items.Add($obj)
                }
                $listBox.EndUpdate()
            }
            else
            {
                $listBox.Items.Add($Items)    
            }
        
            $listBox.DisplayMember = $DisplayMember    
        }
        
        #endregion
        
        $btnSelectOU_Click={
            $SelectedOU = Select-OU
            $txtOU.Text = $SelectedOU.OUDN
        }
        
        $buttonLoadOU_Click={
            if ($txtOU.Text -ne '')
            {
                $root = [ADSI]"LDAP://$($txtOU.Text)"
                $search = [adsisearcher]$root
                if ($radiobuttonAll.Checked)
                {
                    $Search.Filter = '(&(objectClass=computer))'
                }
                if ($radiobuttonServers.Checked)
                {
                    $Search.Filter = '(&(objectClass=computer)(OperatingSystem=Windows*Server*))'
                }
                if ($radiobuttonWorkstations.Checked)
                {
                    $Search.Filter = '(&(objectClass=computer)(!OperatingSystem=Windows*Server*))'
                }
                if ($radiobuttonDomainControllers.Checked)
                {
                    $search.Filter = '(&(&(objectCategory=computer)(objectClass=computer))(UserAccountControl:1.2.840.113556.1.4.803:=8192))'
                }
                
                $colResults = $Search.FindAll()
                $OUResults = @()
                foreach ($i in $colResults)
                {
                    $OUResults += [string]$i.Properties.Item('Name')
                }
        
               # $OUResults | Measure-Object
                Load-ListBox $listBoxComputers $OUResults
            }
        }
        
        $buttonOK_Click={
            if ($listboxComputers.Items.Count -eq 0)
            {
                #[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
                [void][System.Windows.Forms.MessageBox]::Show('No computers listed. If you selected an OU already then please click the Load button.',"Nothing to do")
            }
            else
            {
                $MainForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            }
        
        }
            # --End User Generated Script--
        #----------------------------------------------
        #region Generated Events
        #----------------------------------------------
        
        $Form_StateCorrection_Load=
        {
            #Correct the initial state of the form to prevent the .Net maximized form issue
            $MainForm.WindowState = $InitialFormWindowState
        }
        
        $Form_StoreValues_Closing=
        {
            #Store the control values
            $script:MainForm_radiobuttonDomainControllers = $radiobuttonDomainControllers.Checked
            $script:MainForm_radiobuttonWorkstations = $radiobuttonWorkstations.Checked
            $script:MainForm_radiobuttonServers = $radiobuttonServers.Checked
            $script:MainForm_radiobuttonAll = $radiobuttonAll.Checked
            $script:MainForm_listboxComputersSelected = $listboxComputers.SelectedItems
            $script:MainForm_listboxComputersAll = $listboxComputers.Items
            $script:MainForm_txtOU = $txtOU.Text
        }

        $Form_Cleanup_FormClosed=
        {
            #Remove all event handlers from the controls
            try
            {
                $buttonOK.remove_Click($buttonOK_Click)
                $buttonLoadOU.remove_Click($buttonLoadOU_Click)
                $btnSelectOU.remove_Click($btnSelectOU_Click)
                $MainForm.remove_Load($form1_FadeInLoad)
                $timerFadeIn.remove_Tick($timerFadeIn_Tick)
                $MainForm.remove_Load($Form_StateCorrection_Load)
                $MainForm.remove_Closing($Form_StoreValues_Closing)
                $MainForm.remove_FormClosed($Form_Cleanup_FormClosed)
            }
            catch [Exception]
            { }
        }
        #endregion Generated Events

        #----------------------------------------------
        #region Generated Form Code
        #----------------------------------------------
        #
        # MainForm
        #
        $MainForm.Controls.Add($buttonOK)
        $MainForm.Controls.Add($buttonLoadOU)
        $MainForm.Controls.Add($groupbox1)
        $MainForm.Controls.Add($listboxComputers)
        $MainForm.Controls.Add($labelOrganizationalUnit)
        $MainForm.Controls.Add($txtOU)
        $MainForm.Controls.Add($btnSelectOU)
        $MainForm.ClientSize = '627, 255'
        $MainForm.FormBorderStyle = 'FixedDialog'
        $MainForm.MaximizeBox = $False
        $MainForm.MinimizeBox = $False
        $MainForm.Name = "MainForm"
        $MainForm.StartPosition = 'CenterScreen'
        $MainForm.Tag = ""
        $MainForm.Text = "System Selection"
        $MainForm.add_Load($form1_FadeInLoad)
        #
        # buttonOK
        #
        $buttonOK.Location = '547, 230'
        $buttonOK.Name = "buttonOK"
        $buttonOK.Size = '75, 23'
        $buttonOK.TabIndex = 8
        $buttonOK.Text = "OK"
        $buttonOK.UseVisualStyleBackColor = $True
        $buttonOK.add_Click($buttonOK_Click)
        #
        # buttonLoadOU
        #
        $buttonLoadOU.Location = '288, 52'
        $buttonLoadOU.Name = "buttonLoadOU"
        $buttonLoadOU.Size = '58, 20'
        $buttonLoadOU.TabIndex = 7
        $buttonLoadOU.Text = "Load -->"
        $buttonLoadOU.UseVisualStyleBackColor = $True
        $buttonLoadOU.add_Click($buttonLoadOU_Click)
        #
        # groupbox1
        #
        $groupbox1.Controls.Add($radiobuttonDomainControllers)
        $groupbox1.Controls.Add($radiobuttonWorkstations)
        $groupbox1.Controls.Add($radiobuttonServers)
        $groupbox1.Controls.Add($radiobuttonAll)
        $groupbox1.Location = '13, 52'
        $groupbox1.Name = "groupbox1"
        $groupbox1.Size = '136, 111'
        $groupbox1.TabIndex = 6
        $groupbox1.TabStop = $False
        $groupbox1.Text = "Computer Type"
        #
        # radiobuttonDomainControllers
        #
        $radiobuttonDomainControllers.Location = '7, 79'
        $radiobuttonDomainControllers.Name = "radiobuttonDomainControllers"
        $radiobuttonDomainControllers.Size = '117, 25'
        $radiobuttonDomainControllers.TabIndex = 3
        $radiobuttonDomainControllers.Text = "Domain Controllers"
        $radiobuttonDomainControllers.UseVisualStyleBackColor = $True
        #
        # radiobuttonWorkstations
        #
        $radiobuttonWorkstations.Location = '7, 59'
        $radiobuttonWorkstations.Name = "radiobuttonWorkstations"
        $radiobuttonWorkstations.Size = '104, 25'
        $radiobuttonWorkstations.TabIndex = 2
        $radiobuttonWorkstations.Text = "Workstations"
        $radiobuttonWorkstations.UseVisualStyleBackColor = $True
        #
        # radiobuttonServers
        #
        $radiobuttonServers.Location = '7, 40'
        $radiobuttonServers.Name = "radiobuttonServers"
        $radiobuttonServers.Size = '104, 24'
        $radiobuttonServers.TabIndex = 1
        $radiobuttonServers.Text = "Servers"
        $radiobuttonServers.UseVisualStyleBackColor = $True
        #
        # radiobuttonAll
        #
        $radiobuttonAll.Checked = $True
        $radiobuttonAll.Location = '7, 20'
        $radiobuttonAll.Name = "radiobuttonAll"
        $radiobuttonAll.Size = '104, 24'
        $radiobuttonAll.TabIndex = 0
        $radiobuttonAll.TabStop = $True
        $radiobuttonAll.Text = "All"
        $radiobuttonAll.UseVisualStyleBackColor = $True
        #
        # listboxComputers
        #
        $listboxComputers.FormattingEnabled = $True
        $listboxComputers.Location = '352, 25'
        $listboxComputers.Name = "listboxComputers"
        $listboxComputers.SelectionMode = 'MultiSimple'
        $listboxComputers.Size = '270, 199'
        $listboxComputers.Sorted = $True
        $listboxComputers.TabIndex = 5
        #
        # labelOrganizationalUnit
        #
        $labelOrganizationalUnit.Location = '76, 5'
        $labelOrganizationalUnit.Name = "labelOrganizationalUnit"
        $labelOrganizationalUnit.Size = '125, 17'
        $labelOrganizationalUnit.TabIndex = 4
        $labelOrganizationalUnit.Text = "Organizational Unit"
        #
        # txtOU
        #
        $txtOU.Location = '76, 25'
        $txtOU.Name = "txtOU"
        $txtOU.ReadOnly = $True
        $txtOU.Size = '270, 20'
        $txtOU.TabIndex = 3
        #
        # btnSelectOU
        #
        $btnSelectOU.Location = '13, 25'
        $btnSelectOU.Name = "btnSelectOU"
        $btnSelectOU.Size = '58, 20'
        $btnSelectOU.TabIndex = 2
        $btnSelectOU.Text = "Select"
        $btnSelectOU.UseVisualStyleBackColor = $True
        $btnSelectOU.add_Click($btnSelectOU_Click)
        #
        # timerFadeIn
        #
        $timerFadeIn.add_Tick($timerFadeIn_Tick)
        #endregion Generated Form Code

        #----------------------------------------------

        #Save the initial state of the form
        $InitialFormWindowState = $MainForm.WindowState
        #Init the OnLoad event to correct the initial state of the form
        $MainForm.add_Load($Form_StateCorrection_Load)
        #Clean up the control events
        $MainForm.add_FormClosed($Form_Cleanup_FormClosed)
        #Store the control values when form is closing
        $MainForm.add_Closing($Form_StoreValues_Closing)
        #Show the Form
        return $MainForm.ShowDialog()

    }
    #endregion Source: MainForm.pff

    #region Source: Globals.ps1
        #--------------------------------------------
        # Declare Global Variables and Functions here
        #--------------------------------------------
        
        #Sample function that provides the location of the script
        function Get-ScriptDirectory
        { 
            if($hostinvocation -ne $null)
            {
                Split-Path $hostinvocation.MyCommand.path
            }
            else
            {
                Split-Path $script:MyInvocation.MyCommand.Path
            }
        }
        
        #Sample variable that provides the location of the script
        [string]$ScriptDirectory = Get-ScriptDirectory
        
        function Select-OU
        {
            <#
              .SYNOPSIS
              .DESCRIPTION
              .PARAMETER <Parameter-Name>
              .EXAMPLE
              .INPUTS
              .OUTPUTS
              .NOTES
                My Script Name.ps1 Version 1.0 by Thanatos on 7/13/2013
              .LINK
            #>
            [CmdletBinding()]
            param(
            )
            #$ErrorActionPreference = 'Stop'
            #Set-StrictMode -Version Latest
        
            #region Show / Hide PowerShell Window
            $WindowDisplay = @"
        using System;
        using System.Runtime.InteropServices;

        namespace Window
        {
          public class Display
          {
            [DllImport("Kernel32.dll")]
            private static extern IntPtr GetConsoleWindow();

            [DllImport("user32.dll")]
            private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

            public static bool Hide()
            {
              return ShowWindowAsync(GetConsoleWindow(), 0);
            }

            public static bool Show()
            {
              return ShowWindowAsync(GetConsoleWindow(), 5);
            }
          }
        }
"@
            Add-Type -TypeDefinition $WindowDisplay
            #endregion
            #[Void][Window.Display]::Hide()
        
            [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
            [void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')
        
            #region ******** Select OrganizationalUnit Dialog ******** 
        
            <#
              $SelectOU_Form.Tag.SearchRoot
                Current: Current Domain Ony, Default Value
                Forest: All Domains in the Forest
                Specific OU: DN of a Spedific OU
              
              $SelectOU_Form.Tag.IncludeContainers
                $False: Only Display OU's, Default Value
                $True: Display OU's and Contrainers
              
              $SelectOU_Form.Tag.SelectedOUName
                The Returned Name of the Selected OU
                
              $SelectOU_Form.Tag.SelectedOUDName
                The Returned Name of the Selected OU
        
              if ($SelectOU_Form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
              {
                Write-Host -Object "Selected OU Name = $($SelectOU_Form.Tag.SelectedOUName)"
                Write-Host -Object "Selected OU DName = $($SelectOU_Form.Tag.SelectedOUDName)"
              }
              else
              {
                Write-Host -Object "Selected OU Name = None"
                Write-Host -Object "Selected OU DName = None"
              }
            #>
        
            $SelectOUSpacer = 8
        
            #region $SelectOU_Form = System.Windows.Forms.Form
            $SelectOU_Form = New-Object -TypeName System.Windows.Forms.Form
            $SelectOU_Form.ControlBox = $False
            $SelectOU_Form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $SelectOU_Form.Font = New-Object -TypeName System.Drawing.Font ("Verdana",9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point)
            $SelectOU_Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
            $SelectOU_Form.MaximizeBox = $False
            $SelectOU_Form.MinimizeBox = $False
            $SelectOU_Form.Name = "SelectOU_Form"
            $SelectOU_Form.ShowInTaskbar = $False
            $SelectOU_Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
            $SelectOU_Form.Tag = New-Object -TypeName PSObject -Property @{ 
                                                                    'SearchRoot' = "Current"
                                                                    'IncludeContainers' = $False
                                                                    'SelectedOUName' = "None"
                                                                    'SelectedOUDName' = "None"
                                                                   }
            $SelectOU_Form.Text = "Select OrganizationalUnit"
            #endregion
        
            #region function Load-SelectOU_Form
            function Load-SelectOU_Form ()
            {
                <#
                .SYNOPSIS
                  Load event for the SelectOU_Form Control
                .DESCRIPTION
                  Load event for the SelectOU_Form Control
                .PARAMETER Sender
                   The Form Control that fired the Event
                .PARAMETER EventArg
                   The Event Arguments for the Event
                .EXAMPLE
                   Load-SelectOU_Form -Sender $SelectOU_Form -EventArg $_
                .INPUTS
                .OUTPUTS
                .NOTES
                .LINK
              #>
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory = $True)]
                    [object]$Sender,
                    [Parameter(Mandatory = $True)]
                    [object]$EventArg
                )
                try
                {
                    $SelectOU_Domain_ComboBox.Items.Clear()
                    $SelectOU_OrgUnit_TreeView.Nodes.Clear()
                    switch ($SelectOU_Form.Tag.SearchRoot)
                    {
                        "Current"
                        {
                            $SelectOU_Domain_GroupBox.Visible = $False
                            $SelectOU_OrgUnit_GroupBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,$SelectOUSpacer)
                            $SelectOU_Domain_ComboBox.Items.AddRange($([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain() | Select-Object -Property @{ "Name" = "Text"; "Expression" = { $_.Name } },@{ "Name" = "Value"; "Expression" = { $_.GetDirectoryEntry().distinguishedName } },@{ "Name" = "Domain"; "Expression" = { $Null } }))
                            $SelectOU_Domain_ComboBox.SelectedIndex = 0
                            break
                        }
                        "Forest"
                        {
                            $SelectOU_Domain_GroupBox.Visible = $True
                            $SelectOU_OrgUnit_GroupBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOU_Domain_GroupBox.Bottom + $SelectOUSpacer))
                            $SelectOU_Domain_ComboBox.Items.AddRange($($([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()).Domains | Select-Object -Property @{ "Name" = "Text"; "Expression" = { $_.Name } },@{ "Name" = "Value"; "Expression" = { $_.GetDirectoryEntry().distinguishedName } },@{ "Name" = "Domain"; "Expression" = { $Null } }))
                            $SelectOU_Domain_ComboBox.SelectedItem = $SelectOU_Domain_ComboBox.Items | Where-Object -FilterScript { $_.Value -eq $([adsi]"").distinguishedName }
                            break
                        }
                        Default
                        {
                            $SelectOU_Domain_GroupBox.Visible = $False
                            $SelectOU_OrgUnit_GroupBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,$SelectOUSpacer)
                            $SelectOU_Domain_ComboBox.Items.AddRange($([adsi]"LDAP://$($SelectOU_Form.Tag.SearchRoot)" | Select-Object -Property @{ "Name" = "Text"; "Expression" = { $_.Name } },@{ "Name" = "Value"; "Expression" = { $_.distinguishedName } },@{ "Name" = "Domain"; "Expression" = { $Null } }))
                            $SelectOU_Domain_ComboBox.SelectedIndex = 0
                            break
                        }
                    }
                    $SelectOU_OK_Button.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOU_OrgUnit_GroupBox.Bottom + $SelectOUSpacer))
                    $SelectOU_Cancel_Button.Location = New-Object -TypeName System.Drawing.Point (($SelectOU_OK_Button.Right + $SelectOUSpacer),($SelectOU_OrgUnit_GroupBox.Bottom + $SelectOUSpacer))
                    $SelectOU_Form.ClientSize = New-Object -TypeName System.Drawing.Size (($($SelectOU_Form.Controls[$SelectOU_Form.Controls.Count - 1]).Right + $SelectOUSpacer),($($SelectOU_Form.Controls[$SelectOU_Form.Controls.Count - 1]).Bottom + $SelectOUSpacer))
                }
                catch
                {
                    Write-Warning ('Load-SelectOU_Form Error: {0}' -f $_.Exception.Message)
                    $SelectOU_OK_Button.Enabled = $false
                }
            }
            #endregion
        
            $SelectOU_Form.add_Load({ Load-SelectOU_Form -Sender $SelectOU_Form -EventArg $_ })
        
            #region $SelectOU_Domain_GroupBox = System.Windows.Forms.GroupBox
            $SelectOU_Domain_GroupBox = New-Object -TypeName System.Windows.Forms.GroupBox
            $SelectOU_Form.Controls.Add($SelectOU_Domain_GroupBox)
            $SelectOU_Domain_GroupBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,$SelectOUSpacer)
            $SelectOU_Domain_GroupBox.Name = "SelectOU_Domain_GroupBox"
            $SelectOU_Domain_GroupBox.Text = "Select Domain"
            #endregion
        
            #region $SelectOU_Domain_ComboBox = System.Windows.Forms.ComboBox
            $SelectOU_Domain_ComboBox = New-Object -TypeName System.Windows.Forms.ComboBox
            $SelectOU_Domain_GroupBox.Controls.Add($SelectOU_Domain_ComboBox)
            $SelectOU_Domain_ComboBox.AutoSize = $True
            $SelectOU_Domain_ComboBox.DisplayMember = "Text"
            $SelectOU_Domain_ComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            $SelectOU_Domain_ComboBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOUSpacer + (($SelectOU_Domain_GroupBox.Font.Size * ($SelectOU_Domain_GroupBox.CreateGraphics().DpiY)) / 72)))
            $SelectOU_Domain_ComboBox.Name = "SelectOU_Domain_ComboBox"
            $SelectOU_Domain_ComboBox.ValueMember = "Value"
            $SelectOU_Domain_ComboBox.Width = 400
            #endregion
        
            #region function SelectedIndexChanged-SelectOU_Domain_ComboBox
            function SelectedIndexChanged-SelectOU_Domain_ComboBox ()
            {
                <#
                .SYNOPSIS
                  SelectedIndexChanged event for the SelectOU_Domain_ComboBox Control
                .DESCRIPTION
                  SelectedIndexChanged event for the SelectOU_Domain_ComboBox Control
                .PARAMETER Sender
                   The Form Control that fired the Event
                .PARAMETER EventArg
                   The Event Arguments for the Event
                .EXAMPLE
                   SelectedIndexChanged-SelectOU_Domain_ComboBox -Sender $SelectOU_Domain_ComboBox -EventArg $_
                .INPUTS
                .OUTPUTS
                .NOTES
                .LINK
              #>
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory = $True)]
                    [object]$Sender,
                    [Parameter(Mandatory = $True)]
                    [object]$EventArg
                )
                try
                {
                    if ($SelectOU_Domain_ComboBox.SelectedIndex -gt -1)
                    {
                        $SelectOU_OrgUnit_TreeView.Nodes.Clear()
                        if ([string]::IsNullOrEmpty($SelectOU_Domain_ComboBox.SelectedItem.Domain))
                        {
                            $TempNode = New-Object System.Windows.Forms.TreeNode ($SelectOU_Domain_ComboBox.SelectedItem.Text,[System.Windows.Forms.TreeNode[]](@( "$*$")))
                            $TempNode.Tag = $SelectOU_Domain_ComboBox.SelectedItem.Value
                            $TempNode.Checked = $True
                            $SelectOU_OrgUnit_TreeView.Nodes.Add($TempNode)
                            $SelectOU_OrgUnit_TreeView.Nodes.Item(0).Expand()
                            $SelectOU_Domain_ComboBox.SelectedItem.Domain = $SelectOU_OrgUnit_TreeView.Nodes.Item(0)
                        }
                        else
                        {
                            $SelectOU_OrgUnit_TreeView.Nodes.Add($SelectOU_Domain_ComboBox.SelectedItem.Domain)
                        }
                    }
                }
                catch
                {
                }
            }
            #endregion
        
            $SelectOU_Domain_ComboBox.add_SelectedIndexChanged({ SelectedIndexChanged-SelectOU_Domain_ComboBox -Sender $SelectOU_Domain_ComboBox -EventArg $_ })
        
            $SelectOU_Domain_GroupBox.ClientSize = New-Object -TypeName System.Drawing.Size (($($SelectOU_Domain_GroupBox.Controls[$SelectOU_Domain_GroupBox.Controls.Count - 1]).Right + $SelectOUSpacer),($($SelectOU_Domain_GroupBox.Controls[$SelectOU_Domain_GroupBox.Controls.Count - 1]).Bottom + $SelectOUSpacer))
        
            #region $SelectOU_OrgUnit_GroupBox = System.Windows.Forms.GroupBox
            $SelectOU_OrgUnit_GroupBox = New-Object -TypeName System.Windows.Forms.GroupBox
            $SelectOU_Form.Controls.Add($SelectOU_OrgUnit_GroupBox)
            $SelectOU_OrgUnit_GroupBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOU_Domain_GroupBox.Bottom + $SelectOUSpacer))
            $SelectOU_OrgUnit_GroupBox.Name = "SelectOU_OrgUnit_GroupBox"
            $SelectOU_OrgUnit_GroupBox.Text = "Select OrganizationalUnit"
            $SelectOU_OrgUnit_GroupBox.Width = $SelectOU_Domain_GroupBox.Width
            #endregion
        
            #region $SelectOU_OrgUnit_TreeView = System.Windows.Forms.TreeView
            $SelectOU_OrgUnit_TreeView = New-Object -TypeName System.Windows.Forms.TreeView
            $SelectOU_OrgUnit_GroupBox.Controls.Add($SelectOU_OrgUnit_TreeView)
            $SelectOU_OrgUnit_TreeView.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOUSpacer + (($SelectOU_OrgUnit_GroupBox.Font.Size * ($SelectOU_OrgUnit_GroupBox.CreateGraphics().DpiY)) / 72)))
            $SelectOU_OrgUnit_TreeView.Name = "SelectOU_OrgUnit_TreeView"
            $SelectOU_OrgUnit_TreeView.Size = New-Object -TypeName System.Drawing.Size (($SelectOU_OrgUnit_GroupBox.ClientSize.Width - ($SelectOUSpacer * 2)),300)
            #endregion
        
            #region function BeforeExpand-SelectOU_OrgUnit_TreeView
            function BeforeExpand-SelectOU_OrgUnit_TreeView ()
            {
                <#
                .SYNOPSIS
                  BeforeExpand event for the SelectOU_OrgUnit_TreeView Control
                .DESCRIPTION
                  BeforeExpand event for the SelectOU_OrgUnit_TreeView Control
                .PARAMETER Sender
                   The Form Control that fired the Event
                .PARAMETER EventArg
                   The Event Arguments for the Event
                .EXAMPLE
                   BeforeExpand-SelectOU_OrgUnit_TreeView -Sender $SelectOU_OrgUnit_TreeView -EventArg $_
                .INPUTS
                .OUTPUTS
                .NOTES
                .LINK
              #>
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory = $True)]
                    [object]$Sender,
                    [Parameter(Mandatory = $True)]
                    [object]$EventArg
                )
                try
                {
                    if ($EventArg.Node.Checked)
                    {
                        $EventArg.Node.Checked = $False
                        $EventArg.Node.Nodes.Clear()
                        if ($SelectOU_Form.Tag.IncludeContainers)
                        {
                            $MySearcher = [adsisearcher]"(|((&(objectClass=organizationalunit)(objectCategory=organizationalUnit))(&(objectClass=container)(objectCategory=container))(&(objectClass=builtindomain)(objectCategory=builtindomain))))"
                        }
                        else
                        {
                            $MySearcher = [adsisearcher]"(&(objectClass=organizationalunit)(objectCategory=organizationalUnit))"
                        }
                        $MySearcher.SearchRoot = [adsi]"LDAP://$($EventArg.Node.Tag)"
                        $MySearcher.SearchScope = "OneLevel"
                        $MySearcher.Sort = New-Object -TypeName System.DirectoryServices.SortOption ("Name","Ascending")
                        $MySearcher.SizeLimit = 0
                        [void]$MySearcher.PropertiesToLoad.Add("name")
                        [void]$MySearcher.PropertiesToLoad.Add("distinguishedname")
                        foreach ($Item in $MySearcher.FindAll())
                        {
                            $TempNode = New-Object System.Windows.Forms.TreeNode ($Item.Properties["name"][0],[System.Windows.Forms.TreeNode[]](@( "$*$")))
                            $TempNode.Tag = $Item.Properties["distinguishedname"][0]
                            $TempNode.Checked = $True
                            $EventArg.Node.Nodes.Add($TempNode)
                        }
                    }
                }
                catch
                {
                    Write-Host $Error[0]
                }
            }
            #endregion
            $SelectOU_OrgUnit_TreeView.add_BeforeExpand({ BeforeExpand-SelectOU_OrgUnit_TreeView -Sender $SelectOU_OrgUnit_TreeView -EventArg $_ })
        
            $SelectOU_OrgUnit_GroupBox.ClientSize = New-Object -TypeName System.Drawing.Size (($($SelectOU_OrgUnit_GroupBox.Controls[$SelectOU_OrgUnit_GroupBox.Controls.Count - 1]).Right + $SelectOUSpacer),($($SelectOU_OrgUnit_GroupBox.Controls[$SelectOU_OrgUnit_GroupBox.Controls.Count - 1]).Bottom + $SelectOUSpacer))
        
            #region $SelectOU_OK_Button = System.Windows.Forms.Button
            $SelectOU_OK_Button = New-Object -TypeName System.Windows.Forms.Button
            $SelectOU_Form.Controls.Add($SelectOU_OK_Button)
            $SelectOU_OK_Button.AutoSize = $True
            $SelectOU_OK_Button.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOU_OrgUnit_GroupBox.Bottom + $SelectOUSpacer))
            $SelectOU_OK_Button.Name = "SelectOU_OK_Button"
            $SelectOU_OK_Button.Text = "OK"
            $SelectOU_OK_Button.Width = ($SelectOU_OrgUnit_GroupBox.Width - $SelectOUSpacer) / 2
            #endregion
        
            #region function Click-SelectOU_OK_Button
            function Click-SelectOU_OK_Button
            {
                <#
                .SYNOPSIS
                  Click event for the SelectOU_OK_Button Control
                .DESCRIPTION
                  Click event for the SelectOU_OK_Button Control
                .PARAMETER Sender
                   The Form Control that fired the Event
                .PARAMETER EventArg
                   The Event Arguments for the Event
                .EXAMPLE
                   Click-SelectOU_OK_Button -Sender $SelectOU_OK_Button -EventArg $_
                .INPUTS
                .OUTPUTS
                .NOTES
                .LINK
              #>
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory = $True)]
                    [object]$Sender,
                    [Parameter(Mandatory = $True)]
                    [object]$EventArg
                )
                try
                {
                    if (-not [string]::IsNullOrEmpty($SelectOU_OrgUnit_TreeView.SelectedNode))
                    {
                        $SelectOU_Form.Tag.SelectedOUName = $SelectOU_OrgUnit_TreeView.SelectedNode.Text
                        $SelectOU_Form.Tag.SelectedOUDName = $SelectOU_OrgUnit_TreeView.SelectedNode.Tag
                        $SelectOU_Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    }
                }
                catch
                {
                    Write-Host $Error[0]
                }
            }
            #endregion
            $SelectOU_OK_Button.add_Click({ Click-SelectOU_OK_Button -Sender $SelectOU_OK_Button -EventArg $_ })
        
            #region $SelectOU_Cancel_Button = System.Windows.Forms.Button
            $SelectOU_Cancel_Button = New-Object -TypeName System.Windows.Forms.Button
            $SelectOU_Form.Controls.Add($SelectOU_Cancel_Button)
            $SelectOU_Cancel_Button.AutoSize = $True
            $SelectOU_Cancel_Button.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $SelectOU_Cancel_Button.Location = New-Object -TypeName System.Drawing.Point (($SelectOU_OK_Button.Right + $SelectOUSpacer),($SelectOU_OrgUnit_GroupBox.Bottom + $SelectOUSpacer))
            $SelectOU_Cancel_Button.Name = "SelectOU_Cancel_Button"
            $SelectOU_Cancel_Button.Text = "Cancel"
            $SelectOU_Cancel_Button.Width = ($SelectOU_OrgUnit_GroupBox.Width - $SelectOUSpacer) / 2
            #endregion
        
            $SelectOU_Form.ClientSize = New-Object -TypeName System.Drawing.Size (($($SelectOU_Form.Controls[$SelectOU_Form.Controls.Count - 1]).Right + $SelectOUSpacer),($($SelectOU_Form.Controls[$SelectOU_Form.Controls.Count - 1]).Bottom + $SelectOUSpacer))
            #endregion
        
            $ReturnedOU = @{
                'OUName' = $null;
                'OUDN' = $null
            }
            if ($SelectOU_Form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $ReturnedOU.OUName = $($SelectOU_Form.Tag.SelectedOUName)
                $ReturnedOU.OUDN = $($SelectOU_Form.Tag.SelectedOUDName)
            }
            New-Object PSobject -Property $ReturnedOU
        }
        
        
    #endregion Source: Globals.ps1

    #Start the application
    Main ($CommandLine)
        New-Object PSObject -Property @{
                    'AllResults' = $MainForm_listboxComputersAll
                    'SelectedResults' = $MainForm_listboxComputersSelected
                }
}
#endregion Functions - Serial or Utility

#region Functions - Asset Report Project
Function Create-ReportSection
{
    #** This function is specific to this script and does all kinds of bad practice
    #   stuff. Use this function neither to learn from or judge me please. **
    #
    #   That being said, this function pretty much does all the report output
    #   options and layout magic. It depends upon the report layout hash and
    #   $HTMLRendering global variable hash.
    #
    #   This function generally shouldn't need to get changed in any way to customize your
    #   reports.
    #
    # .EXAMPLE
    #    Create-ReportSection -Rpt $ReportSection -Asset $Asset 
    #                         -Section 'Summary' -TableTitle 'System Summary'
    
    [CmdletBinding()]
    param(
        [parameter()]
        $Rpt,
        
        [parameter()]
        [string]$Asset,

        [parameter()]
        [string]$Section,
        
        [parameter()]
        [string]$TableTitle        
    )
    BEGIN
    {
    }
    PROCESS
    {
    }
    END
    {
        # Get our section type
        $SectionType = $Rpt[$Section]['Type']
        switch ($SectionType)
        {
            'Section'     # default to a data section
            {
                Write-Verbose -Message ('Create-ReportSection: {0}:{1}' -f $Asset,$Section)
                $ReportElementSource = @($Rpt[$Section]['AllData'][$Asset])
                if ((($ReportElementSource.Count -gt 0) -and 
                     ($ReportElementSource[0] -ne $null)) -or 
                     ($Rpt[$Section]['AllowEmptyReport']))
                {
                    $SourceProperties = $Rpt[$Section]['ReportTypes'][$ReportType]['Properties']
                    
                    #region report section type and layout
                    $TableType = $Rpt[$Section]['ReportTypes'][$ReportType]['TableType']
                    $ContainerType = $Rpt[$Section]['ReportTypes'][$ReportType]['ContainerType']

                    switch ($TableType)
                    {
                        'Horizontal' 
                        {
                            $PropertyCount = $SourceProperties.Count
                            $Vertical = $false
                        }
                        'Vertical' {
                            $PropertyCount = 2
                            $Vertical = $true
                        }
                        default {
                            if ((($SourceProperties.Count) -ge $HorizontalThreshold))
                            {
                                $PropertyCount = 2
                                $Vertical = $true
                            }
                            else
                            {
                                $PropertyCount = $SourceProperties.Count
                                $Vertical = $false
                            }
                        }
                    }
                    #endregion report section type and layout
                    
                    $Table = ''
                    If ($PropertyCount -ne 0)
                    {
                        # Create our future HTML table header
                        $TableHeader = $HTMLRendering['TableTitle'][$HTMLMode] -replace '<0>',$PropertyCount
                        $TableHeader = $TableHeader -replace '<1>',$TableTitle
                        
                        $AllTableElements = @()
                        Foreach ($TableElement in $ReportElementSource)
                        {
                            $AllTableElements += $TableElement | Select $SourceProperties
                        }

                        # If we are creating a vertical table it takes a bit of transformational work
                        if ($Vertical)
                        {
                            $Count = 0
                            foreach ($Element in $AllTableElements)
                            {
                                $Count++
                                $SingleElement = ($Element | ConvertTo-PropertyValue | ConvertTo-Html)
                                # Add class elements for even/odd rows
                                $SingleElement = Colorize-Table $SingleElement -ColorizeMethod 'ByEvenRows' -Attr 'class' -AttrValue 'even' -WholeRow:$true
                                $SingleElement = Colorize-Table $SingleElement -ColorizeMethod 'ByOddRows' -Attr 'class' -AttrValue 'odd' -WholeRow:$true
                                if (($Rpt[$Section].ContainsKey('PostProcessing')) -and 
                                   (($Rpt[$Section]['PostProcessing'].Value -ne $false)))
                                {
                                    $Table = $(Invoke-Command ([scriptblock]::Create($Rpt[$Section]['PostProcessing'])))
                                }
                                $SingleElement = [Regex]::Match($SingleElement, "(?s)(?<=</tr>)(.+)(?=</table>)").Value
                                $Table += $SingleElement 
                                if ($Count -ne $AllTableElements.Count)
                                {
                                    $Table += '<tr class="divide"><td></td><td></td></tr>'
                                }
                            }
                            $Table = '<table class="list">' + $TableHeader + $Table + '</table>'
                            $Table = [System.Web.HttpUtility]::HtmlDecode($Table)
                        }
                        # Otherwise it is a horizontal table
                        else
                        {
                            $Table = $AllTableElements | ConvertTo-Html
                            # Add class elements for even/odd rows
                            $Table = Colorize-Table $Table -ColorizeMethod 'ByEvenRows' -Attr 'class' -AttrValue 'even' -WholeRow:$true
                            $Table = Colorize-Table $Table -ColorizeMethod 'ByOddRows' -Attr 'class' -AttrValue 'odd' -WholeRow:$true
                            if ($Rpt[$Section].ContainsKey('PostProcessing')) 
                            {
                                if ($Rpt[$Section].ContainsKey('PostProcessing'))
                                {
                                    if ($Rpt[$Section]['PostProcessing'] -ne $false)
                                    {
                                        $Table = $(Invoke-Command ([scriptblock]::Create($Rpt[$Section]['PostProcessing'])))
                                    }
                                }
                            }
                            # This will gank out everything after the first colgroup so we can replace it with our own spanned header
                            $Table = [Regex]::Match($Table, "(?s)(?<=</colgroup>)(.+)(?=</table>)").Value
                            $Table = '<table>' + $TableHeader + $Table + '</table>'
                            $Table = [System.Web.HttpUtility]::HtmlDecode(($Table))
                        }
                    }
                    
                    $Output = $HTMLRendering['SectionContainers'][$HTMLMode][$ContainerType]['Head'] + 
                              $Table + $HTMLRendering['SectionContainers'][$HTMLMode][$ContainerType]['Tail']
                    $Output
                }
            }
            default
            {
                Write-Verbose -Message ('Create-ReportSection: {0}' -f $SectionType)
                $Output = $HTMLRendering['CustomSections'][$SectionType] -replace '<0>',$TableTitle
                $Output
            }
        }
    }
}

Function ReportProcessing
{
    [CmdletBinding()]
    param
    (
        [Parameter( HelpMessage="Report body, typically in HTML format",
                    ValueFromPipeline=$true,
                    Mandatory=$true )]
        [string[]]
        $Report = ".",
        
        [Parameter( HelpMessage="Email server to relay report through")]
        [string]
        $EmailRelay = ".",
        
        [Parameter( HelpMessage="Email sender")]
        [string]
        $EmailSender='systemreport@localhost',
        
        [Parameter( HelpMessage="Email recipient")]
        [string]
        $EmailRecipient='default@yourdomain.com',
        
        [Parameter( HelpMessage="Email subject")]
        [string]
        $EmailSubject='System Report',
        
        [Parameter( HelpMessage="Email body as html")]
        [switch]
        $EmailBodyAsHTML=$true,
        
        [Parameter( HelpMessage="Send email of resulting report?")]
        [switch]
        $SendMail,

        [Parameter( HelpMessage="Save the report?")]
        [switch]
        $SaveReport,
        
        [Parameter( HelpMessage="If saving the report, what do you want to call it?")]
        [string]
        $ReportName="Report.html"
    )
    BEGIN
    {
    }
    PROCESS
    {
        if ($SaveReport)
        {
            $Report | Out-File $ReportName
        }
        if ($Sendmail)
        {
            send-mailmessage -from $EmailSender -to $EmailRecipient -subject $EmailSubject `
            -BodyAsHTML:$EmailBodyAsHTML -Body $Report -priority Normal -smtpServer $EmailRelay
        }
    }
    END
    {
    }
}

Function Gather-ReportInformation
{
    [CmdletBinding()]
    param
    (
        [Parameter( HelpMessage="Computer or computers to return information about",
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter( HelpMessage="The custom report hash variable structure you plan to report upon")]
        $ReportContainer,
        
        [Parameter( HelpMessage="List of sorted report elements within ReportContainer")]
        $SortedRpts,
        
        [parameter( HelpMessage="Pass an alternate credential" )]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter( HelpMessage="View visual progress bar.")]
        [switch]
        $ShowProgress
    )
    BEGIN
    {
        $ComputerNames = @($ComputerName)
    }
    PROCESS
    {
        # I think I use a different variable name for credential splatting in every script i write...
        $_credsplat = @{
            'ComputerName' = $ComputerNames
            'Verbose' = ($PSBoundParameters['Verbose'] -eq $true)
        }
        $_summarysplat = @{
            'ComputerName' = $ComputerNames
            'Verbose' = ($PSBoundParameters['Verbose'] -eq $true)
        }
        $_hphardwaresplat = @{
            'ComputerName' = $ComputerNames
            'Verbose' = ($PSBoundParameters['Verbose'] -eq $true)
        }
        $_credsplatserial = @{            
            'Verbose' = ($PSBoundParameters['Verbose'] -eq $true)
        }
        
        if ($Credential -ne $null)
        {
            $_credsplat.Credential = $Credential
            $_summarysplat.Credential = $Credential
            $_hphardwaresplat.Credential = $Credential
            $_credsplatserial.Credential = $Credential
        }
        
        # Multithreaded Information Gathering
        $HPHardwareHealthTesting = $false   # Only run this test if at least one of the several HP health tests are enabled
        # Call multiple runspace supported info gathering functions where supported and create
        # splats for functions which gather multiple section data.
        
        $SortedRpts | %{ switch ($_.Section) {
            'ExtendedSummary' {
                $NTPInfo = @(Get-RemoteRegistryInformation @_credsplat `
                                    -Key $reg_NTPSettings)
                $NTPInfo = ConvertTo-HashArray $NTPInfo 'PSComputerName'
                $ExtendedInfo = @(Get-RemoteRegistryInformation @_credsplat `
                                    -Key $reg_ExtendedInfo)
                $ExtendedInfo = ConvertTo-HashArray $ExtendedInfo 'PSComputerName'
            }
            'LocalGroupMembership' {
                $LocalGroupMembership = @(Get-RemoteGroupMembership @_credsplat)
                $LocalGroupMembership = ConvertTo-HashArray $LocalGroupMembership 'PSComputerName'
            }
            'Memory' {
                $_summarysplat.IncludeMemoryInfo = $true
            }
            'Disk' {
                $_summarysplat.IncludeDiskInfo = $true
            }
            'Network' {
                $_summarysplat.IncludeNetworkInfo = $true
            }
            'RouteTable' {
                $RouteTables = @(Get-RemoteRouteTable @_credsplat)
                $RouteTables = ConvertTo-HashArray $RouteTables 'PSComputerName'
            }
            'ShareSessionInfo' {
                $ShareSessions = @(Get-RemoteShareSessionInformation @_credsplat)
                $ShareSessions = ConvertTo-HashArray $ShareSessions 'PSComputerName'
            }
            'ProcessesByMemory' {
                # Processes by memory
                $ProcsByMemory = @(Get-RemoteProcessInformation @_credsplat)
                $ProcsByMemory = ConvertTo-HashArray $ProcsByMemory 'PSComputerName'
            }
            'StoppedServices' {
                $Filter = "(StartMode='Auto') AND (State='Stopped')"
                $StoppedServices = @(Get-RemoteServiceInformation @_credsplat -Filter $Filter)
                $StoppedServices = ConvertTo-HashArray $StoppedServices 'PSComputerName'
            }
            'NonStandardServices' {
                $Filter = "NOT startName LIKE 'NT AUTHORITY%' AND NOT startName LIKE 'localsystem'"
                $NonStandardServices = @(Get-RemoteServiceInformation @_credsplat -Filter $Filter)
                $NonStandardServices = ConvertTo-HashArray $NonStandardServices 'PSComputerName'
            }
            'Applications' {
                $InstalledPrograms = @(Get-RemoteInstalledPrograms @_credsplat)
                $InstalledPrograms = ConvertTo-HashArray $InstalledPrograms 'PSComputerName'
            }
            'InstalledUpdates' {
                $InstalledUpdates = @(Get-MultiRunspaceWMIObject @_credsplat `
                                            -Class Win32_QuickFixEngineering)
                $InstalledUpdates = ConvertTo-HashArray $InstalledUpdates 'PSComputerName'
            }
            'Printers' {
                $Printers = @(Get-RemoteInstalledPrinters @_credsplat)
                $Printers = ConvertTo-HashArray $Printers 'PSComputerName'
            }
            'EventLogSettings' {
                $EventLogSettings = @(Get-MultiRunspaceWMIObject @_credsplat `
                                            -Class win32_NTEventlogFile)
                $EventLogSettings = ConvertTo-HashArray $EventLogSettings 'PSComputerName'
            }
            'Shares' {
                $Shares = @(Get-MultiRunspaceWMIObject @_credsplat `
                                            -Class win32_Share)
                $Shares = ConvertTo-HashArray $Shares 'PSComputerName'
            }
            'EventLogs' {
                # Event log errors/warnings/audit failures
                $EventLogs = @(Get-RemoteEventLogs @_credsplat -Hours $Option_EventLogPeriod)
                $EventLogs = ConvertTo-HashArray $EventLogs 'PSComputerName'
            }
            'AppliedGPOs' {
                $AppliedGPOs = @(Get-RemoteAppliedGPOs @_credsplat)
                $AppliedGPOs = ConvertTo-HashArray $AppliedGPOs 'PSComputerName'
            }
            'WSUSSettings' {
                # WSUS settings
                $WSUSSettings = @(Get-RemoteRegistryInformation @_credsplat -Key $reg_WSUSSettings)
                $WSUSSettings = ConvertTo-HashArray $WSUSSettings 'PSComputerName'
            }
            'HP_GeneralHardwareHealth' {
                $HPHardwareHealthTesting = $true
            }
            'HP_EthernetTeamHealth' {
                $HPHardwareHealthTesting = $true
                $_hphardwaresplat.IncludeEthernetTeamHealth = $true
            }
            'HP_ArrayControllerHealth' {
                $HPHardwareHealthTesting = $true
                $_hphardwaresplat.IncludeArrayControllerHealth = $true
            }
            'HP_EthernetHealth' {
                $HPHardwareHealthTesting = $true
                $_hphardwaresplat.IncludeEthernetHealth = $true
            }
            'HP_FanHealth' {
                $HPHardwareHealthTesting = $true
                $_hphardwaresplat.IncludeFanHealth = $true
            }
            'HP_HBAHealth' {
                $HPHardwareHealthTesting = $true
                $_hphardwaresplat.IncludeHBAHealth = $true
            }
            'HP_PSUHealth' {
                $HPHardwareHealthTesting = $true
                $_hphardwaresplat.IncludePSUHealth = $true
            }
            'HP_TempSensors' {
                $HPHardwareHealthTesting = $true
                $_hphardwaresplat.IncludeTempSensors = $true
            }
        } }
        $_summarysplat.ShowProgress = $ShowProgress
        $Assets = @(Get-ComputerAssetInformation @_summarysplat)

        # HP Server Health
        if ($HPHardwareHealthTesting)
        {
            $HPServerHealth = @(Get-HPServerhealth @_hphardwaresplat)
            $HPServerHealth = ConvertTo-HashArray $HPServerHealth 'PSComputerName'
        }
        
        # Serial Information Gathering
        # Seperate our data out to its appropriate report section under 'AllData' as a hash key with 
        #   an array of objects/data as the key value.
        # This is also where you can gather and store report section data with non-multithreaded
        #  functions.
        Foreach ($AssetInfo in $Assets)
        {
            $SortedRpts | %{ 
            switch ($_.Section) {
                'Summary' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($AssetInfo | select *)
                }
                'ExtendedSummary' {
                    # we have to mash up the results of a few different reg entries for this one
                    $tmpobj = ConvertTo-PSObject `
                                -InputObject $ExtendedInfo[$AssetInfo.PScomputername].Registry `
                                -propname 'Key' -valname 'KeyValue'
                    $tmpobj2 = ConvertTo-PSObject `
                                -InputObject $NTPInfo[$AssetInfo.PScomputername].Registry `
                                -propname 'Key' -valname 'KeyValue'
                    $tmpobj | Add-Member -MemberType NoteProperty -Name 'NTPType' -Value $tmpobj2.Type
                    $tmpobj | Add-Member -MemberType NoteProperty -Name 'NTPServer' -Value $tmpobj2.NtpServer
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = @($tmpobj)
                }
                'LocalGroupMembership' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($LocalGroupMembership[$AssetInfo.PScomputername].GroupMembership)
                }
                'Disk' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = @($AssetInfo._Disks)
                }                    
                'Network' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        $AssetInfo._Network | Where {$_.ConnectionStatus}
                }
                'RouteTable' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($RouteTables[$AssetInfo.PScomputername].Routes |
                            Sort-Object 'Metric1')
                }
                'Memory' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($AssetInfo._MemorySlots)
                }
                'StoppedServices' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($StoppedServices[$AssetInfo.PScomputername].Services)
                }
                'NonStandardServices' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($NonStandardServices[$AssetInfo.PScomputername].Services)
                }
                'ProcessesByMemory' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($ProcsByMemory[$AssetInfo.PScomputername].Processes |
                            Sort WS -Descending |
                            Select -First $Option_TotalProcessesByMemory)
                }
                'EventLogSettings' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($EventLogSettings[$AssetInfo.PScomputername].WMIObjects)
                }
                'Shares' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($Shares[$AssetInfo.PScomputername].WMIObjects)
                }
                'DellWarrantyInformation' {
                    $_DellWarrantyInformation = Get-DellWarranty @_credsplatserial -ComputerName $AssetInfo.PSComputerName
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = @($_DellWarrantyInformation)
                }
                
                'Applications' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] =
                        @($InstalledPrograms[$AssetInfo.PScomputername].Programs | 
                            Sort-Object DisplayName)
                }
                
                'InstalledUpdates' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($InstalledUpdates[$AssetInfo.PScomputername].WMIObjects)
                }
                'Printers' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] =
                        @($Printers[$AssetInfo.PScomputername].Printers)
                }
                
                'EventLogs' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($EventLogs[$AssetInfo.PScomputername].EventLogs |
                            Select -First $Option_EventLogResults)
                }
                'AppliedGPOs' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($AppliedGPOs[$AssetInfo.PScomputername].AppliedGPOs |
                            Sort-Object AppliedOrder)
                }
                'ShareSessionInfo' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($ShareSessions[$AssetInfo.PScomputername].Sessions | 
                            Group-Object -Property ShareName | Sort-Object Count -Descending)
                }
                'WSUSSettings' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($WSUSSettings[$AssetInfo.PScomputername].Registry)
                }
                'HP_GeneralHardwareHealth' {
                    if ($HPServerHealth -ne $null)
                    {
                        $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                            @($HPServerHealth[$AssetInfo.PScomputername])
                    }
                }
                'HP_EthernetTeamHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].ContainsKey('_EthernetTeamHealth'))
                        {
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._EthernetTeamHealth)
                        }
                    }
                }
                'HP_ArrayControllerHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].ContainsKey('_ArrayControllers'))
                        {
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._ArrayControllers)
                        }
                    }
                }
                'HP_EthernetHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].ContainsKey('_EthernetHealth'))
                        {  
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._EthernetHealth)
                        }
                    }
                }
                'HP_FanHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].ContainsKey('_FanHealth'))
                        {  
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._FanHealth)
                        }
                    }
                }
                'HP_HBAHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].ContainsKey('_HBAHealth'))
                        {  
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._HBAHealth)
                        }
                    }
                }
                'HP_PSUHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].ContainsKey('_PSUHealth'))
                        {  
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._PSUHealth)
                        }
                    }
                }
                'HP_TempSensors' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].ContainsKey('_TempSensors'))
                        {                
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._TempSensors)
                        }
                    }
                }
            }}
        }
    }
    END
    {
    }
}

Function New-AssetReport
{
    <#
    .SYNOPSIS
        Generates a new asset report from gathered data.
    .DESCRIPTION
        Generates a new asset report from gathered data. There are multiple input and output methods.
    .PARAMETER ComputerName
        Computer or computers to return information about.
    .PARAMETER ReportContainer
        The custom report hash variable structure you plan to report upon.
    .PARAMETER ReportType
        The report type.
    .PARAMETER HTMLMode
        The HTML rendering type (DynamicGrid or EmailFriendly).
    .PARAMETER ExportToExcel
        Export an excel document.
    .PARAMETER PromptForCredential
        Set this if you want the function to prompt for alternate credentials.
    .PARAMETER Credential
        Pass an alternate credential.
    .PARAMETER EmailRelay
        Email server to relay report through.
    .PARAMETER EmailSender
        Email sender.
    .PARAMETER EmailRecipient
        Email recipient.
    .PARAMETER EmailSubject
        Email subject.
    .PARAMETER SendMail
        Send email of resulting report?
    .PARAMETER SaveReport
        Save the report?
    .PARAMETER OutputMethod
        If saving the report, will it be one big report or individual reports?
    .PARAMETER ReportName
        If saving the report, what do you want to call it?
    .PARAMETER ReportLocation
        If saving multiple reports, where will they be saved?
        
    .EXAMPLE
        $Computers = @('Server1','Server2')
        $cred = get-credential
        New-AssetReport -ComputerName $Computers `
                -ReportContainer $SystemReport `
                -SaveReport `
                -OutputMethod 'OneBigReport' `
                -HTMLMode 'DynamicGrid' `
                -ReportType 'FullDocumentation' `
                -ReportName 'servers.html' `
                -ExporttoExcel `
                -Credential $cred `
                -Verbose

        Description:
        ------------------
        Prompt for an alternate credential then use it to generate one big html report called servers.html and
        create an excel file with all of the gathered data. The report type is a full documentation report of
        Server1 and Server2. The output format for html is dynamic grid. Verbose output is displayed (which shows
        different aspects of processing as it is occurring).

    .NOTES
        Version    : 1.0.0 Sept 7th 2013
                           - First release

        Author     : Zachary Loeber

        Disclaimer : This script is provided AS IS without warranty of any kind. I 
                     disclaim all implied warranties including, without limitation,
                     any implied warranties of merchantability or of fitness for a 
                     particular purpose. The entire risk arising out of the use or
                     performance of the sample scripts and documentation remains
                     with you. In no event shall I be liable for any damages 
                     whatsoever (including, without limitation, damages for loss of 
                     business profits, business interruption, loss of business 
                     information, or other pecuniary loss) arising out of the use of or 
                     inability to use the script or documentation. 

        Copyright  : I believe in sharing knowledge, so this script and its use is 
                     subject to : http://creativecommons.org/licenses/by-sa/3.0/


    .LINK
        http://zacharyloeber.com/

    .LINK
        http://nl.linkedin.com/in/zloeber

    #>

    #region Parameters
    [CmdletBinding()]
    PARAM
    (
        [Parameter( HelpMessage="Computer or computers to return information about",
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter( Mandatory=$true,
                    HelpMessage="The custom report hash variable structure you plan to report upon")]
        $ReportContainer,
        
        [Parameter( HelpMessage="The report type")]
        [ValidateSet('Troubleshooting','FullDocumentation')]
        [string]
        $ReportType = 'FullDocumentation',        
        
        [Parameter( HelpMessage="The HTML rendering type (DynamicGrid or EmailFriendly)")]
        [ValidateSet("DynamicGrid","EmailFriendly")]
        [string]
        $HTMLMode = 'DynamicGrid',
        
        [Parameter( HelpMessage='Export an excel document')]
        [switch]
        $ExportToExcel,
        
        [parameter( HelpMessage="Set this if you want the function to prompt for alternate credentials" )]
        [switch]
        $PromptForCredential,
        
        [parameter( HelpMessage="Pass an alternate credential" )]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,
        
        [Parameter( HelpMessage="Email server to relay report through")]
        [string]
        $EmailRelay = ".",
        
        [Parameter( HelpMessage="Email sender")]
        [string]
        $EmailSender='systemreport@localhost',
     
        [Parameter( HelpMessage="Email recipient")]
        [string]
        $EmailRecipient='default@yourdomain.com',
        
        [Parameter( HelpMessage="Email subject")]
        [string]
        $EmailSubject='System Report',
        
        [Parameter( HelpMessage="Send email of resulting report?")]
        [switch]
        $SendMail,

        [Parameter( HelpMessage="Save the report?")]
        [switch]
        $SaveReport,
        
        [Parameter( HelpMessage="If saving the report, will it be one big report or individual reports?")]
        [ValidateSet('OneBigReport','IndividualReport')]
        [string]
        $OutputMethod='OneBigReport',
        
        [Parameter( HelpMessage="If saving the report, what do you want to call it?")]
        [string]
        $ReportName="Report.html",
        
        [Parameter( HelpMessage="If saving multiple reports, where will they be saved?")]
        [string]
        $ReportLocation="."
    )
    #endregion Parameters
    BEGIN
    {
        # Use this to keep a splat of our CmdletBinding options
        $VerboseDebug=@{}
        If ($PSBoundParameters.ContainsKey('Verbose')) {
            If ($PSBoundParameters.Verbose -eq $true) { $VerboseDebug.Verbose = $true } else { $VerboseDebug.Verbose = $false }
        }
        If ($PSBoundParameters.ContainsKey('Debug')) {
            If ($PSBoundParameters.Debug -eq $true) { $VerboseDebug.Debug = $true } else { $VerboseDebug.Debug = $false }
        }
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        $AssetNames = @()
        $credsplat = @{}
        $summarysplat = @{}
        if ($Credential -ne $null)
        {
            $credsplat.Credential = $Credential
            $summarysplat.Credential = $Credential
        }
        
        $ReportProcessingSplat = @{
            'EmailSender' = $EmailSender
            'EmailRecipient' = $EmailRecipient
            'EmailSubject' = $EmailSubject
            'EmailRelay' = $EmailRelay
            'SendMail' = $SendMail
            'SaveReport' = $SaveReport
        }
        
        # Some basic initialization
        $FinalReport = ''
        $ServerReports = ''
        
        # There must be a more elegant way to do this hash sorting but this also allows
        # us to pull a list of only the sections which are defined and need to be generated.
        $SortedReports = @()
        Foreach ($Key in $ReportContainer['Sections'].Keys) 
        {
            if ($ReportContainer['Sections'][$Key]['ReportTypes'].ContainsKey($ReportType))
            {
                if ($ReportContainer['Sections'][$Key]['Enabled'] -and 
                    ($ReportContainer['Sections'][$Key]['ReportTypes'][$ReportType] -ne $false))
                {
                    $_SortedReportProp = @{
                                            'Section' = $Key
                                            'Order' = $ReportContainer['Sections'][$Key]['Order']
                                          }
                    $SortedReports += New-Object -Type PSObject -Property $_SortedReportProp
                }
            }
        }
        $SortedReports = $SortedReports | Sort-Object Order
    }
    PROCESS
    {
        $AssetNames += $ComputerName
    }
    END 
    {
        if ($AssetNames -ne $null)
        {
            # Information Gathering
            Invoke-Command ([scriptblock]::Create($ReportContainer['Configuration']['PreProcessing']))

            # if we are to export all data to excel, then we do so per section then per computer
            if ($ExportToExcel)
            {
                # First make sure we have data to export, this shlould also weed out non-data sections meant for html
                #  (like section breaks and such)
                $ProcessExcelReport = $false
                foreach ($ReportSection in $SortedReports)
                {
                    if ($ReportContainer['Sections'][$ReportSection.Section]['AllData'].Count -gt 0)
                    {
                        $ProcessExcelReport = $true
                    }
                }

                #region Excel
                if ($ProcessExcelReport)
                {
                    # Create the excel workbook
                    try
                    {
                        #$Excel = New-Object -Com Excel.Application -ErrorAction Stop
                        $Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
                        $ExcelExists = $True
                        $Excel.visible = $True
                        #Start-Sleep -s 1
                        $Workbook = $Excel.Workbooks.Add()
                        $Excel.DisplayAlerts = $false
                    }
                    catch
                    {
                        Write-Warning ('Issues opening excel: {0}' -f $_.Exception.Message)
                        $ExcelExists = $False
                    }
                    if ($ExcelExists)
                    {
                        # going through every section, but in reverse so it shows up in the correct
                        #  sheet in excel. 
                        $SortedExcelReports = $SortedReports | Sort-Object Order -Descending
                        Foreach ($ReportSection in $SortedExcelReports)
                        {
                            $SectionData = $ReportContainer['Sections'][$ReportSection.Section]['AllData']
                            $SectionProperties = $ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType]['Properties']
                            
                            # Gather all the asset information in the section (remember that each asset may
                            #  be pointing to an array of psobjects)
                            $TransformedSectionData = @()                        
                            foreach ($asset in $SectionData.Keys)
                            {
                                # Get all of our calculated properties, then add in the asset name
                                $TempProperties = $SectionData[$asset] | Select $SectionProperties
                                $TransformedSectionData += ($TempProperties | Select @{n='PSComputerName';e={$asset}},*)
                            }
                            if (($TransformedSectionData.Count -gt 0) -and ($TransformedSectionData -ne $null))
                            {
                                $temparray1 = $TransformedSectionData | ConvertTo-MultiArray
                                if ($temparray1 -ne $null)
                                {    
                                    $temparray = $temparray1.Value
                                    $starta = [int][char]'a' - 1
                                    
                                    if ($temparray.GetLength(1) -gt 26) 
                                    {
                                        $col = [char]([int][math]::Floor($temparray.GetLength(1)/26) + $starta) + [char](($temparray.GetLength(1)%26) + $Starta)
                                    } 
                                    else 
                                    {
                                        $col = [char]($temparray.GetLength(1) + $starta)
                                    }
                                    
                                    Start-Sleep -s 1
                                    $xlCellValue = 1
                                    $xlEqual = 3
                                    $BadColor = 13551615    #Light Red
                                    $BadText = -16383844    #Dark Red
                                    $GoodColor = 13561798    #Light Green
                                    $GoodText = -16752384    #Dark Green
                                    $Worksheet = $Workbook.Sheets.Add()
                                    $Worksheet.Name = $ReportSection.Section
                                    $Range = $Worksheet.Range("a1","$col$($temparray.GetLength(0))")
                                    $Range.Value2 = $temparray

                                    #Format the end result (headers, autofit, et cetera)
                                    [void]$Range.EntireColumn.AutoFit()
                                    [void]$Range.FormatConditions.Add($xlCellValue,$xlEqual,'TRUE')
                                    $Range.FormatConditions.Item(1).Interior.Color = $GoodColor
                                    $Range.FormatConditions.Item(1).Font.Color = $GoodText
                                    [void]$Range.FormatConditions.Add($xlCellValue,$xlEqual,'OK')
                                    $Range.FormatConditions.Item(2).Interior.Color = $GoodColor
                                    $Range.FormatConditions.Item(2).Font.Color = $GoodText
                                    [void]$Range.FormatConditions.Add($xlCellValue,$xlEqual,'FALSE')
                                    $Range.FormatConditions.Item(3).Interior.Color = $BadColor
                                    $Range.FormatConditions.Item(3).Font.Color = $BadText
                                    
                                    # Header
                                    $range = $Workbook.ActiveSheet.Range("a1","$($col)1")
                                    $range.Interior.ColorIndex = 19
                                    $range.Font.ColorIndex = 11
                                    $range.Font.Bold = $True
                                    $range.HorizontalAlignment = -4108
                                }
                            }
                        }
                        # Get rid of the blank default worksheets
                        $Workbook.Worksheets.Item("Sheet1").Delete()
                        $Workbook.Worksheets.Item("Sheet2").Delete()
                        $Workbook.Worksheets.Item("Sheet3").Delete()
                    }
                }
                #endregion Excel
            }
            foreach ($Asset in $AssetNames)
            {
                # First check if there is any data to report upon for each asset
                $ContainsData = $false
                $SectionCount = 0
                Foreach ($ReportSection in $SortedReports)
                {
                    if ($ReportContainer['Sections'][$ReportSection.Section]['AllData'].ContainsKey($Asset))
                    {
                        $ContainsData = $true
                        #$SectionCount++             #Not currently used
                    }
                }
                
                # If we have any data then we have a report to create
                if ($ContainsData)
                {
                    $ServerReport = ''
                    $ServerReport += $HTMLRendering['ServerBegin'][$HTMLMode] -replace '<0>',$Asset
                    $UsedSections = 0
                    $TotalSectionsPerRow = 0
                    
                    Foreach ($ReportSection in $SortedReports)
                    {
                        if ($ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType])
                        {
                            #region Section Calculation
                            # Use this code to track where we are at in section usage
                            #  and create new section groupss as needed
                            
                            # Current section type
                            $CurrContainer = $ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType]['ContainerType']
                            
                            # Grab first two digits found in the section container div
                            $SectionTracking = ([Regex]'\d{1}').Matches($HTMLRendering['SectionContainers'][$HTMLMode][$CurrContainer]['Head'])
                            #Write-Verbose -Message ('Report {0}: Section calculation - {1}' -f $Asset,$ReportSection.Section)
                            #Write-Verbose -Message ('Section: {0}, HTML: {1}' -f $CurrContainer,$HTMLRendering['SectionContainers'][$HTMLMode][$CurrContainer]['Head'])
                            if (($SectionTracking[1].Value -ne $TotalSectionsPerRow) -or `
                                ($SectionTracking[0].Value -eq $SectionTracking[1].Value) -or `
                                (($UsedSections + [int]$SectionTracking[0].Value) -gt $TotalSectionsPerRow) -and `
                                (!$ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType]['SectionOverride']))
                            {
                                $NewGroup = $true
                            }
                            else
                            {
                                $NewGroup = $false
                                $UsedSections += [int]$SectionTracking[0].Value
                                Write-Verbose -Message ('Report {0}: NOT a new group, Sections used {1}' -f $Asset,$UsedSections)
                            }
                            
                            if ($NewGroup)
                            {
                                if ($UsedSections -ne 0)
                                {
                                    $ServerReport += $HTMLRendering['SectionContainerGroup'][$HTMLMode]['Tail']
                                }
                                #$ServerReport += $HTMLRendering['SectionContainerGroup'][$HTMLMode]['Tail']
                                $ServerReport += $HTMLRendering['SectionContainerGroup'][$HTMLMode]['Head']
                                $UsedSections = [int]$SectionTracking[0].Value
                                $TotalSectionsPerRow = [int]$SectionTracking[1].Value
                                Write-Verbose -Message ('Report {0}: {1}/{2} Sections Used' -f $Asset,$UsedSections,$TotalSectionsPerRow)
                            }
                            #endregion Section Calculation
                            
                            Write-Verbose -Message ('Report {0}: HTML Table creation - {1}' -f $Asset,$ReportSection.Section)
                            $ServerReport += Create-ReportSection -Rpt $ReportContainer['Sections'] `
                                                                  -Asset $Asset `
                                                                  -Section $ReportSection.Section `
                                                                  -TableTitle $ReportContainer['Sections'][$ReportSection.Section]['Title']
                        }
                    }
                    
                    $ServerReport += $HTMLRendering['SectionContainerGroup'][$HTMLMode]['Tail']
                    $ServerReport += $HTMLRendering['ServerEnd'][$HTMLMode]
                    $ServerReports += $ServerReport
                    
                }
                # If we are creating per-asset reports then create one now, otherwise keep going
                if ($OutputMethod -eq 'IndividualReport')
                {
                    $ReportProcessingSplat.Report = ($HTMLRendering['Header'][$HTMLMode] -replace '<0>','$Asset') + 
                                                        $ServerReports + 
                                                        $HTMLRendering['Footer'][$HTMLMode]
                    $ReportProcessingSplat.ReportName = $ReportLocation + '\' + $Asset + '.html'
                    ReportProcessing @ReportProcessingSplat
                    $ServerReports = ''
                }
            }
            
            # If one big report is getting sent/saved do so now
            if ($OutputMethod -eq 'OneBigReport')
            {
                $FullReport = ($HTMLRendering['Header'][$HTMLMode] -replace '<0>','Multiple Systems') + 
                               $ServerReports + 
                               $HTMLRendering['Footer'][$HTMLMode]
                $ReportProcessingSplat.ReportName = $ReportLocation + '\' + $ReportName
                $ReportProcessingSplat.Report = ($HTMLRendering['Header'][$HTMLMode] -replace '<0>','Multiple Systems') + 
                                                    $ServerReports + 
                                                    $HTMLRendering['Footer'][$HTMLMode]
                ReportProcessing @ReportProcessingSplat
            }
        }
    }
}
#endregion Functions - Asset Report Project
#endregion Functions

#region Main
# MODIFY THIS AREA TO SUIT YOUR NEED

# Enter in admin credentials for your remote systems (cause you don't simply login to
# your workstation as a domain admin right?... Right!?!).
#$Cred = Get-Credential
#
## Use our nifty AD OU selector computer filtering GUI to get a list of systems
## and feed them into the reporting engine
#(Get-OUResults).SelectedResults | New-AssetReport `
#									-ReportContainer $SystemReport `
#						            -SaveReport `
#						            -OutputMethod 'IndividualReport' `
#						            -HTMLMode 'DynamicGrid' `
#						            -ReportName 'report.html' `
#						            -ExporttoExcel `
#						            -Credential $Cred `
#						            -Verbose 
#endregion Main   