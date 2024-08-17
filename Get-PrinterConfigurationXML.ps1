<#
.SYNOPSIS
    Enumerates the printers on a device that do not match the exclusion list and extracts the print configuration from the PrintTicketXML property.
.DESCRIPTION
    Using Get-Printer the script will gather a list of all printers on the system.  Printers with names that match the exclusion rules will be omitted.  The PrintTicketXML property will be extracted for each and saved to a file in the path specified.
.NOTES
    
.EXAMPLE
    .\Get-PrinterConfigurationXML.ps1 -Exclusions ("Fax","OneNote (Desktop)") -OutputPath C:\tmp
    .\Get-PrinterConfigurationXML.ps1 -Exclusions (Get-Content List_of_NewLine_Seperated_Exclusions.txt) -OutputPath C:\tmp

    If you use a text file to list exclusions, the file should be one printer name per line, without any quotes or commas.

#>

param(
    [array]$Exclusions,
    $OutputPath = "C:\IT\PrinterScriptFiles"
)
Function Export-PrinterDriver{
    param(
        $InfPath,
        $DriverName,
        $OutputPath
    )
    $infFile = Split-Path $InfPath -Leaf
    $infDirectory = Split-Path $InfPath -Parent
    If(
        (
            $OutputPath[-1] -eq '\') -or (
            $OutputPath[-1] -eq '/')
        ){
            $OutputPath = $OutputPath[0..($OutputPath.Length -2)] -Join ""
        }
    Write-Host "Backing Up $infFile Driver Folder..."

    Try{
        Compress-Archive -Path $InfDirectory -DestinationPath $OutputPath\$DriverName -Force -ErrorAction STOP
        Write-Host "Success! Folder Backed Up to $OutputPath\$DriverName.zip"
    }
    Catch{
        Write-Host $error[0]
        Write-Error "Unable to create archive backup of $infDirectory"
        Exit 1
    }
    
}

$currentDate = $(Get-Date -Format MMddyyyy)
$devicePrinters = Get-Printer
$tcpipPrinterPorts = Get-PrinterPort | Where Description -like "*TCP/IP*"
$printerDrivers = Get-PrinterDriver

If ($devicePrinters.count -lt 1){
    Write-Host "No Printers Found! Exiting..."
    Exit 0
}

Write-Host "Found the Following Printers:"
Write-Host ($devicePrinters|Out-String)

$realPrinters = $devicePrinters |
    Where-Object Name -notin $exclusions |
    Where-Object Name -notlike "*PDF*" |
    Where-Object Name -notlike "*XPS Document Writer*" |
    Where-Object Name -notlike "*OneNote*" |
    Where-Object Name -ne "Fax" 

$excludedPrinters = $devicePrinters | Where-Object Name -notin $realPrinters.Name
Write-Host "The follwing devices were excluded:"
Write-Host ($excludedPrinters|Out-String)

$printConfigurations = $realPrinters|
    Select-Object Name,DriverName,PortName,Location,PrinterHostAddress,PrinterHostIP,PortNumber,
    @{
        N="PortNumber";
        E={
            Get-PrinterPort -Name $_.PortName |
            Select-Object -ExpandProperty PortNumber -ErrorAction SilentlyContinue
        }
    },
    @{
        N="PrinterHostIP";
        E={
            Get-PrinterPort -Name $_.PortName |
            Select-Object -ExpandProperty PrinterHostIP -ErrorAction SilentlyContinue
        }
    },
    @{
        N="PrinterHostAddress";
        E={
            Get-PrinterPort -Name $_.PortName |
            Select-Object -ExpandProperty PrinterHostAddress -ErrorAction SilentlyContinue
        }
    },
    @{
        N="DeviceURL";
        E={
            Get-PrinterPort -Name $_.PortName |
            Select-Object -ExpandProperty DeviceURL -ErrorAction SilentlyContinue
        }
    },
    @{
        N="InfName";
        E={
            Get-PrinterDriver -Name $_.DriverName|
            Select-Object -ExpandProperty InfPath|
            Split-Path -Leaf
        }
    },
    @{
        N="PrintTicketXML";
        E={
            Get-PrintConfiguration -PrinterName $_.Name|
            Select-Object -ExpandProperty PrintTicketXML
        }
    }

    
If(-Not(Test-Path $outputPath)){
    Try{New-Item -ItemType Directory -Path $outputPath -ErrorAction Stop}
    Catch{
        Write-Host $error[0]
        Write-Error "Unable to Create Output Path $outputPath"
        Exit 1
    }
}
Write-Host "`nAttempting to output the following printer configurations:"
Write-Host $($printConfigurations|Format-Table|out-string)
Try{
    $fileName = "$env:computerName--$currentDate.csv"
$printConfigurations |
    Export-CSV -Path "$outputPath\$fileName" -NoTypeInformation -ErrorAction Stop
    Write-Host "Printer information outputted to $outputPath\$fileName"
}
Catch{
    Write-Host $error[0]
    Write-Error "Unable to Write Printer Information to File!"
    Exit 1
}

Write-Host "Attempting to Backup the Following Drivers:"
Write-Host ($realPrinters.DriverName|Out-String)
Foreach($printer in $($realPrinters)){
    Write-Host "Backing up $($printer.DriverName)..."
    $infP = Get-PrinterDriver -Name $printer.DriverName | Select-Object -ExpandProperty InfPath
    Try{
        Export-PrinterDriver -InfPath $infP -DriverName $printer.DriverName -OutputPath $OutputPath
        Write-Host ""
    }
    Catch{
        Write-Host $error[0]
        Write-Error "Error while Exporting Printer Driver ($printer.DriverName)"
        Exit 1
    }
}
Exit 0

