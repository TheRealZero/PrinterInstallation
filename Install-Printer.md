## Install-Printer

### Summary

This script is used for installing printer drivers and adding new printers from a configuration file or manually entering printer port, driver path, and driver name.

### Parameters

- **ConfigFile**: A configuration file with printer details, used when importing printer settings. Use `Get-PrinterConfigurationXML.ps1` to generate a configuration file, with the following format:
  `"Name","DriverName","PortName","Location",PrinterHostAddress,"InfName","PrintTicketXML"`
  
- **RemoveExisting**: A switch to remove existing printers if not found in the ConfigFile. Use `-RemoveExisting` in the command-line to apply the switch.

**THE FOLLOWING PARAMETERS FOR MANUAL INSTALLATION ARE NOT YET IMPLEMENTED**

- **PrinterName**: The name of the printer, used when manually adding a printer.

- **Port**: The printer port, used when manually adding a printer.

- **PrintTicketXML**: The PrintTicketXML configuration string, used when manually adding a printer.

- **DriverPath**: The path to the printer driver, used when manually adding a printer.

- **DriverName**: The name of the printer driver, used when manually adding a printer.

### Usage Examples

1. When using a config file: `. Install-Printer.ps1 -ConfigFile "C:\IT\PrinterScriptFiles\config.csv"`
2. When using a config file and removing any printers not matching the exclusions for buil-in devices and not listed in the config: `. Install-Printer.ps1 -ConfigFile "C:\IT\PrinterScriptFiles\config.csv" -RemoveExisting`
2. When manually adding a printer (**Not Yet Implemented**): `. Install-Printer.ps1 -PrinterName "PrinterName" -Port "PortName" -PrintTicketXML "C:\path\to\PrintTicketXML.xml" -DriverPath "DriverPath" -DriverName "DriverName"`
