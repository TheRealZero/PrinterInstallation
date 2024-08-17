<#
    .SYNOPSIS 
    This script is used for installing printer drivers and adding new printers from a configuration file or manually entering printer port, driver path, and driver name.

    .DESCRIPTION 
    This script is used for installing printer drivers and adding new printers from a configuration file or manually entering printer port, driver path, and driver name.
    The main functions are 


    .PARAMETER ConfigFile 
    A configuration file with printer details, used when importing printer settings.
    Use Get-PrinterConfigurationXML.ps1 to generate a configuration file, with the following format:
    "Name","DriverName","PortName","Location","PrinterHostAddress","PrinterHostIP","PortNumber"


    .PARAMETER ConfigFile
    The configuration file to use when importing printer settings.

    .PARAMETER PrinterName
    The name of the printer, used when manually adding a printer.

    .PARAMETER Port 
    The printer port, used when manually adding a printer.
    
    .PARAMETER PrintTicketXML
    The PrintTicketXML configuration string, used when manually adding a printer.

    .PARAMETER DriverPath 
    The path to the printer driver, used when manually adding a printer.

    .PARAMETER DriverName 
    The name of the printer driver, used when manually adding a printer.

    .EXAMPLE 
    Example of usage: .\PrinterScript.ps1 -ConfigFile "C:\IT\PrinterScriptFiles\config.csv" 
    Example of usage: .\PrinterScript.ps1 -PrinterName "PrinterName" -Port "PortName" -PrintTicketXML "C:\path\to\PrintTicketXML.xml" -DriverPath "DriverPath" -DriverName "DriverName"
#>
param(
    [Parameter(ParameterSetName = "ConfigFile")]
    [string]$ConfigFile = $(Get-ChildItem -Path $PSScriptRoot -Filter *.csv |Select-Object -First 1),
    [Parameter(ParameterSetName = "ConfigFile")]
    [switch]$RemoveExisting = $false,
    [Parameter(ParameterSetName = "Manual", Mandatory)]
    [string]$PrinterName,
    [Parameter(ParameterSetName = "Manual", Mandatory)]
    [string]$Port,
    [Parameter(ParameterSetName = "Manual", Mandatory)]
    [string]$PrintTicketXML,
    [Parameter(ParameterSetName = "Manual", Mandatory)]
    [string]$DriverPath,
    [Parameter(ParameterSetName = "Manual", Mandatory)]
    [string]$DriverName
)
$ErrorActionPreference = "STOP"
Start-Transcript -Path "$env:TEMP\PrinterScript.log" -Append

Function Assert-PrinterExists {
    param(
        $Name
    )
    $printerExists = Get-Printer -Name $Name -ErrorAction SilentlyContinue
    If ($printerExists) {
        
        Write-Output $true
    }
}

Function Assert-PrinterDriverExists {
    param(
        $DriverName
    )
    $driverExists = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
    If ($driverExists) {
        
        Write-Output $true
    }
}

Function Assert-PrinterPortExists {
    param(
        $PortName
    )
    $portExists = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
    If ($portExists) {
        
        Write-Output $true
    }
}

Function Add-NewPrinterPort {
    param(
        $PortName,
        $PrinterHostAddress
    )
    Write-Host "Adding '$($PortName)' Printer Port..."
    Try {
        Add-PrinterPort -Name $PortName -PrinterHostAddress $PrinterHostAddress
        Write-Host "Success!`n"
    }
    Catch {
        Write-Host $error[0]
        Write-Host "Unable to Add Printer Port!"
        Exit 1
    }
}

Function Add-NewPrinterConfiguration {
    param(
        $PrinterName,
        $PrintTicketXML
    )
    Write-Host "Setting '$($PrinterName)' Printer Configuration with XML..."
    Try {
        Set-PrintConfiguration -PrintTicketXml ($PrintTicketXml | Out-String) -PrinterName $PrinterName
    }
    Catch {
        Write-Host $error[0]
        Write-Host"Unable to Set Printer Configuration!"
        Exit 1
    }
}

Function Install-PrinterDriverFromBackup {
    param(
        $BackupFilePath,
        $InfName
    )
    $backupExpandedPath = $BackupFilePath.replace(".zip", "")
        
    Write-Host "Files in $BackupFilePath Will Be Expanded to $backupExpandedPath"
    Write-Host "Expanding..."
    Try {
        Expand-Archive -Path $BackupFilePath -DestinationPath $backupExpandedPath -Force
        Write-Host "Success!"
    }
    Catch {
        Write-Host "Unable To Expand Archive '$BackupFilePath'!"
        Exit 1
    }


    $infFullPath = Get-ChildItem -Path "$backupExpandedPath\*\$InfName" | Select-Object -ExpandProperty FullName
    Write-Host "Using the following Inf File to add Driver to Driver Store:`n$infFullPath"
    $scriptBlock = { C:\Windows\System32\pnputil.exe /a $infFullPath }
    Invoke-Command -ScriptBlock $scriptBlock
    If ($LASTEXITCODE -ne 0) {
        Write-Host "Error While Adding $InfName Driver to DriverStore!"
        Exit 1
    }
}

Function Add-NewPrinterDriver {
    param(
        $DriverName,
        $InfName
    )
    Try {
        Write-Host "Adding Printer Driver from Backup..."
        Install-PrinterDriverFromBackup -BackupFilePath "$PSSCriptRoot\$DriverName.zip" -InfName $InfName
        Write-Host "Adding Printer Driver from Driver Store..."
    }
    Catch {
        Write-Host $error[0]
        Write-Host "Unable to Add Printer Driver from Driver Store!"
        Exit 1
    }

    Try {
        $driverInfFolderName = Get-ChildItem $PSSCriptRoot\$DriverName -Directory |Select-Object -ExpandProperty Name
        Add-PrinterDriver -Name $DriverName -InfPath $(Get-ChildItem -Recurse C:\Windows\System32\DriverStore\FileRepository\$driverInfFolderName\*$InfName)
        Write-Host "Success!`n"
    }
    Catch {
        Write-Host $error[0]
        Write-Host "Unable to Add Printer Driver!"
        Exit 1
    }
}

Function Add-NewPrinter {
    param(
        $PrinterName,
        $PortName,
        $DriverName,
        $Location
    )
    Write-Host "Adding '$($PrinterName)' Printer..."
    Try {
        Add-Printer -Name $PrinterName -DriverName $DriverName -Location $Location -PortName $PortName
    }
    Catch {
        Write-Host $error[0]
        Write-Host "Unable to Add Printer!"
        Exit 1
    }
}

Function Add-NewPrinterFromConfig{
    param(
        $ConfigData
    )
    
    If($removeExisting -eq $true){
        Write-Host "Removing Existing Printers Not Found in Config File..."

            $devicePrinters = Get-Printer 
            $printersToRemove = $devicePrinters |
            Where-Object Name -notlike "*PDF*" |
            Where-Object Name -notlike "*XPS Document Writer*" |
            Where-Object Name -notlike "*OneNote*" |
            Where-Object Name -ne "Fax"|
            Where-Object Name -notin $($ConfigData|Select-Object -ExpandProperty Name)
            
            If($printersToRemove.Count -eq 0){
                
                Write-Host "No Printers to Remove!`n"
            }
            Else{
                
                Write-Host "Removing the following Printers:"
                Write-Host $printersToRemove
                Foreach($printer in $printersToRemove){
                    Try{
                        Write-Host "Removing $($printer.Name)..."
                        Remove-Printer -Name $printer.Name -ErrorAction Stop
                        Write-Host "Success!`n"
                    }
                    Catch {
                        Write-Host $error[0]
                        Write-Host "Unable to Remove Printer!"
                        Exit 1
                    }
                }
            }
            
        
    }
    Foreach($item in $ConfigData){
        Write-Host "Beginning Printer $($item.Name)..."
        $printerExists = Assert-PrinterExists -Name $item.Name
        $portExists = Assert-PrinterPortExists -PortName $item.PortName
        $driverExists = Assert-PrinterDriverExists -DriverName $item.DriverName

        Write-Host "Printer Exists: $printerExists"
        Write-Host "Port Exists: $portExists"
        Write-Host "Driver Exists: $driverExists"
        
        If($portExists -ne $true){
            Add-NewPrinterPort -PortName $item.PortName -PrinterHostAddress $item.PrinterHostAddress
        }
        If($driverExists -ne $true){
            Add-NewPrinterDriver -DriverName $item.DriverName -InfName $item.InfName
        }
        If($printerExists -ne $true){
            Add-NewPrinter -PrinterName $item.Name -PortName $item.PortName -DriverName $item.DriverName -Location $item.Location
        }
        Else{
            Try{
                Write-Host "Setting Printer '$($item.Name)' Location, Port and Driver..."
                Set-Printer -Name $item.Name -Location $item.Location -PortName $item.PortName -DriverName $item.DriverName
            }
            Catch{
                Write-Host $error[0]
                Write-Host "Unable to Set Printer '$($item.Name)' Location, Port and Driver!"
                Exit 1
            }

        }
        Add-NewPrinterConfiguration -PrinterName $item.Name -PrintTicketXml $item.PrintTicketXML
        Write-Host "Printer $($item.Name) Complete!`n"
    }

}

$configData = Import-CSV $ConfigFile
Add-NewPrinterFromConfig -ConfigData $configData
Stop-Transcript