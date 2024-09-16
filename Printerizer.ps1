<#
This Script will gather the existing printer configuration info from WMI, export the data to .CSV files for later reference AND copy the queues and ports over to a new server as specified in the $NewPrintServer variable
Assumptions: 
  You are running this script in an elevated command prompt from the current print server
  Your current credentials have the rights to remote into the new print server by its name and install printers there.
#> 

# Set some constants
$NewPrintServer = ""
$CurrentPrintServer = (hostname)
$OutputPath = "C:\Users\$($env:USERNAME)\Downloads"

# Create Subdirectory for Printer Bin Files
$NULL = New-Item -Path "$OutputPath\PrinterBin\" -ItemType Directory

# Gather raw queue and port information
$WMI_Printers = Get-WmiObject -Class Win32_Printer | where Shared -eq $True
$WMI_PrinterPorts = Get-WmiObject -Class Win32_TCPIPPrinterPort | where Name -in $WMI_Printers.PortName

# Output Relevant Drivers to Console as those need to be migrated
write-host "Make sure to add all these drivers to the new server so that the printers can be added correctly!!!" -ForegroundColor Magenta
$WMI_Printers | group DriverName | select @{n="Driver";e={$_.Name}},Count | ft
#----------------------------------------------------------------------------------------------------#

# Re-serialize Port information into a simple tabular structure (and ping-test the Host Addresses)
$I=0
$PortData = foreach ($WMI_PrinterPort in $WMI_PrinterPorts)
{
    Write-Progress -Activity "Getting Port Data" -CurrentOperation "[$I/$($WMI_PrinterPorts.count)]" -PercentComplete ($I/$WMI_PrinterPorts.Count*100)

    $CanPing = Test-Connection -ComputerName $WMI_PrinterPort.HostAddress -Count 1 -Quiet
    [pscustomobject]([ordered]@{
        Name = $WMI_PrinterPort.Name
        HostAddress = $WMI_PrinterPort.HostAddress
        PortNumber = $WMI_PrinterPort.PortNumber
        Caption = $WMI_PrinterPort.Caption
        Protocol = $WMI_PrinterPort.Protocol
        Queue = $WMI_PrinterPort.Queue
        SNMPCommunity = $WMI_PrinterPort.SNMPCommunity
        SNMPDevIndex = $WMI_PrinterPort.SNMPDevIndex
        SNMPEnabled = $WMI_PrinterPort.SNMPEnabled
        Status = $WMI_PrinterPort.Status
        SystemName = $WMI_PrinterPort.SystemName
        Type = $WMI_PrinterPort.Type
        ComputerName = $WMI_PrinterPort.PSComputerName
        CanPing = $CanPing
    })
    $I++
}
# Export Port information to .csv file
$PortData | Export-Csv -Path "$OutputPath/$($CurrentPrintServer)-PrinterPortData.csv" -NoTypeInformation 

#----------------------------------------------------------------------------------------------------#

# Re-serialize Printer information into a simple tabular structure
$I=0
$PrinterData = foreach ($WMI_Printer in $WMI_Printers)
{
    Write-Progress -Activity "Getting Printer Data" -CurrentOperation "[$I/$($WMI_Printers.count)]" -PercentComplete ($I/$WMI_Printers.Count*100)
    $Port = $WMI_PrinterPorts | where Name -eq $WMI_Printer.PortName
    $StateText = switch ($WMI_Printer.PrinterState)
    {
        0 {"Idle"}
        1 {"Paused"}
        2 {"Error"}
        3 {"Pending Deletion"}
        4 {"Paper Jam"}
        5 {"Paper Out"}
        6 {"Manual Feed"}
        7 {"Paper Problem"}
        8 {"Offline"}
        9 {"I/O Active"}
        10 {"Busy"}
        11 {"Printing"}
        12 {"Output Bin Full"}
        13 {"Not Available"}
        14 {"Waiting"}
        15 {"Processing"}
        16 {"Initialization"}
        17 {"Warming Up"}
        18 {"Toner Low"}
        19 {"No Toner"}
        20 {"Page Punt"}
        21 {"User Intervention Required"}
        22 {"Out of Memory"}
        23 {"Door Open"}
        24 {"Server_Unknown"}
        25 {"Power Save"}
    }

    $StatusText = switch ($WMI_Printer.PrinterStatus)
    {
        1 {"Other"}
        2 {"Unknown"}
        3 {"Idle"}
        4 {"Printing"}
        5 {"Warmup"}
        6 {"Stopped Printing"}
        7 {"Offline"}
    }

    # Export Bin Files for this printer
    $ServerPrinterString = "\\$(hostname)\$($WMI_Printer.Name)"
    $BinFileString = "$OutputPath\PrinterBin\$(hostname)-$($WMI_Printer.Name).bin" 
    PrintUI.exe /Ss /n $ServerPrinterString /a $BinFileString

    # Output Flattened Printer Object
    [pscustomobject]([ordered]@{
        SystemName = $WMI_Printer.SystemName
        DriverName = $WMI_Printer.DriverName
        Name = $WMI_Printer.Name
        DeviceID = $WMI_Printer.DeviceID
        Caption = $WMI_Printer.Caption
        Comment = $WMI_Printer.Comment
        Location = $WMI_Printer.Location
        Description = $WMI_Printer.Description
        PrinterState = $WMI_Printer.PrinterState
        PrinterState_Text = $StateText
        PrinterStatus = $WMI_Printer.PrinterStatus
        PrinterStatus_Text = $StatusText
        Shared = $WMI_Printer.Shared
        ShareName = $WMI_Printer.ShareName
        RawOnly = $WMI_Printer.RawOnly
        EnableBIDI = $WMI_Printer.EnableBIDI
        DoCompleteFirst = $WMI_Printer.DoCompleteFirst
        PrinterPaperNames = $WMI_Printer.PrinterPaperNames
        PaperSizesSupported = $WMI_Printer.PaperSizesSupported
        Network = $WMI_Printer.Network
        CreationClassName = $WMI_Printer.CreationClassName
        PortName = $WMI_Printer.PortName
        Port_Description = $Port.Description
        Port_HostAddress = $Port.HostAddress
        Port_PortNumber = $Port.PortNumber
        Port_Protocol = $Port.Protocol
        Port_Caption = $Port.Caption
        Port_SNMPEnabled = $Port.SNMPEnabled
        Port_SNMPCommunity = $Port.SNMPCommunity
        Port_SNMPDevIndex = $Port.SNMPDevIndex
        Port_SystemCreationClassname = $Port.SystemCreationClassName
        Port_SystemName = $Port.SystemName
        BinFilePath = $BinFileString
    })
    $I++
}
# Export Printer information to .csv file
$PrinterData | Export-Csv -Path "$OutputPath/$($CurrentPrintServer)-PrinterData.csv" -NoTypeInformation 

#----------------------------------------------------------------------------------------------------#

# Create session to new Print Server and pre-build the folder that will contain Tray Defaults (.bin files)
$Session = New-PSSession -ComputerName $NewPrintServer
Invoke-Command -Session $Session -ScriptBlock {
    $OutputPath = $Using:OutputPath
    $NULL = New-Item -Path "$OutputPath\PrinterBin\" -ItemType Directory -Force
}

# Create new Ports and Queues based on Existing Printer Data
$I=0
Foreach ($P in $PrinterData)
{
    Write-Progress -Activity "Creating New Ports and Printers from Data" -CurrentOperation "[$I/$($PrinterData.count)] Printer: $($P.Name)" -PercentComplete ($I/$PrinterData.count*100)
    
    # Copy tray configuration (.bin file) to the new print server
    Copy-Item -Path $P.BinFilePath -Destination $P.BinFilePath -ToSession $Session

    # Create New TCPIPPrinterPort for this Printer
    Invoke-Command -Session $Session -ScriptBlock {
        $PrintServer = (hostname)
        $P = $Using:P

        # Create New TCPIPPrinterPort for this Printer
        $New_Port = ([WMICLASS]"\\$PrintServer\ROOT\cimv2:Win32_TCPIPPrinterPort").createInstance()
        $New_Port.Caption = $P.Caption
        $New_Port.CreationClassName = $P.CreationClassName
        $New_Port.Description = $P.Port_Description
        $New_Port.HostAddress = $P.Port_HostAddress
        $New_Port.InstallDate = (Get-Date)
        $New_Port.Name = $P.PortName
        $New_Port.PortNumber = $P.Port_PortNumber
        $New_Port.Protocol = $P.Port_Protocol
        $New_Port.SNMPCommunity = $P.Port_SNMPCommunity
        $New_Port.SNMPDevIndex = $P.Port_SNMPDevIndex
        $New_Port.SNMPEnabled = $P.Port_SNMPEnabled
        $New_Port.SystemCreationClassName = $P.Port_SystemCreationClassName
        $New_Port.SystemName = $P.Port_SystemName
        $New_Port.Put() | Out-Null
        $New_Port = $Null

        # Create Queue for this Printer
        $New_Queue = ([WMICLASS]"\\$PrintServer\ROOT\cimv2:Win32_Printer").createInstance()
	    $New_Queue.DriverName = $P.DriverName
        $New_Queue.Name = $P.Name
        $New_Queue.PortName = $P.PortName
        $New_Queue.Shared = $P.Shared
	    $New_Queue.Caption = $P.Caption
        $New_Queue.DeviceID = $P.DeviceID
        $New_Queue.Comment = $P.Comment
	    $New_Queue.Location = $P.Location
	    $New_Queue.RawOnly = $P.RawOnly
	    $New_Queue.EnableBIDI= $P.EnableBIDI
        $New_Queue.DoCompleteFirst = $P.DoCompleteFirst
        $New_Queue.PrinterPaperNames = $P.PrinterPaperNames
        $New_Queue.PaperSizesSupported = $P.PaperSizesSupported
        $New_Queue.Network = $P.Network
        $New_Queue.CreationClassName = $P.CreationClassName
	    $New_Queue.Put() | Out-Null
        $New_Queue = $Null

        # Set Tray Default from copied .Bin File
        $ServerPrinterString = "\\$(hostname)\$($P.Name)"
        PrintUI.exe /Sr /n $ServerPrinterString /q /a $P.BinFilePath r d u g
    }
    $I++
}