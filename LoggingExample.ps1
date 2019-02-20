##########################
# script imports printer information from csvs
# then installs on all computers listed in
# the second csv
# ers@swc.com - 1/21/15
##########################

Function CreatePrinterPort ($PrinterIP, $PrinterPort, $PrinterPortName, $ComputerName) {
    $wmi = [wmiclass]"\\$ComputerName\root\cimv2:win32_tcpipPrinterPort"
    $wmi.psbase.scope.options.enablePrivileges = $true
    $Port = $wmi.createInstance()
    $Port.name = $PrinterPortName
    $Port.hostAddress = $PrinterIP
    $Port.portNumber = $PrinterPort
    $Port.SNMPEnabled = $false
    $Port.Protocol = 1
    $Port.put()
}

Function InstallPrinterDriver ($DriverName, $DriverPath, $DriverInf, $ComputerName) {
    $wmi = [wmiclass]"\\$ComputerName\Root\cimv2:Win32_PrinterDriver"
    $wmi.psbase.scope.options.enablePrivileges = $true
    $wmi.psbase.Scope.Options.Impersonation = `
    [System.Management.ImpersonationLevel]::Impersonate
    $Driver = $wmi.CreateInstance()
    $Driver.Name = $DriverName
    $Driver.DriverPath = $DriverPath
    $Driver.InfName = $DriverInf
    $wmi.AddPrinterDriver($Driver)
    $wmi.Put()
}

Function CreatePrinter ($PrinterCaption, $PrinterPortName, $DriverName, $ComputerName) {
    $wmi = ([WMIClass]"\\$ComputerName\Root\cimv2:Win32_Printer")
    $Printer = $wmi.CreateInstance()
    $Printer.Caption = $PrinterCaption
    $Printer.DriverName = $DriverName
    $Printer.PortName = $PrinterPortName
    $Printer.DeviceID = $PrinterCaption
    $Printer.Default = $false
    $Printer.Put()
}

function Write-HostAndLog ( $logString ) {
	Write-Host $logString
	Write-Log $logString
}

function Write-Log( $logString ) {
	Out-File -FilePath $LogFile -Append -Encoding utf8 -NoClobber -InputObject ( [DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + $logString )
}

####################################################

## get printer list from csv (printer name, printer IP, driver, model)
$NewPrinterList = Import-Csv .\partstown.printers.csv
$ComputerList = Import-Csv .\partstown.computers.csv

# set log file
$LogFile = [DateTime]::Today.ToString("yyyy-MM-dd") + ".printer.install.log"

# get list of current printers
$CurrentPrinterList = gwmi win32_printer

# get list of current printer ports
$CurrentPortList = gwmi win32_TCPIPPrinterPort

# get list of current printer drivers
$CurrentDriverList = gwmi win32_printerdriver

## loop through computers
foreach ($computer in $ComputerList) {
    
    # get new computer name
    $ComputerHostName = $computer.name

    # log it
    Write-HostAndLog "Installing printers on $ComputerHostName"

    ## loop through printers
    foreach ($printer in $NewPrinterList) {

        ## assign variables
        $PrinterPort = "9100"
        $PrinterIP = $printer.ip
        $PrinterPortName = "IP_" + $PrinterIP
        $PrinterName = $printer.name
        $PrinterDriver = $printer.driver
        $PrinterModel = $printer.model
        $PrinterDriverPath = $printer.driverpath
        $PrinterDriverInf = $printer.driverinf

        ## check if printer already exists
        if ($CurrentPrinterList | ? {$_.ShareName -eq $PrinterName}) {
            Write-HostAndLog "Printer $PrinterName already exists on $ComputerHostName"
        } else {

            # check for the printer port
            if($CurrentPortList | ? {$_.HostAddress -eq $PrinterIP}) {
                Write-HostAndLog "Port $PrinterIP already exists on $ComputerHostName"
                
                # get the port name in case it is different
                $PrinterPortName = $($CurrentPortList | ? {$_.HostAddress -eq $PrinterIP}).Name

            } else {
                # create the printer port
                Write-HostAndLog "Creating port for $PrinterPortName on $ComputerHostName"
                CreatePrinterPort $PrinterIP $PrinterPort $PrinterPortName $ComputerHostName
            }

            # check for printer driver
            if(-not ($CurrentDriverList | ? {$_.Name -like "*$PrinterDriver*"})) {
                # install driver
                #InstallPrinterDriver $PrinterDriver $PrinterDriverPath $PrinterDriverInf $ComputerName
                Write-HostAndLog "Driver $PrinterDriver not found for $PrinterName"

            } else {

                # create printer
                CreatePrinter $PrinterName $PrinterPortName $PrinterDriver $ComputerHostName
        
                # log it
                Write-Log "Successfully installed $PrinterName at $PrinterIP with driver $PrinterDriver on $ComputerHostName"
            }
        }
    }
}