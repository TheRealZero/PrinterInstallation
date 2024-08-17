# Print Manager Scripts Documentation
Files: Install-Printer.md and Get-PrinterConfigurationXML.md
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


## Get-PrinterConfigurationXML

### Summary
This script is used for generating a printer list and configuration XML data by extracting the current printers and settings. It will also extract the printer drivers necessary for each printer so they can be installed on another device. The configuration file can be later used as input in the Install-Printer.ps1 script.

### Parameters

- **OutputPath**: The path where the generated configuration file will be saved. The default output path is the current working directory. Example: `-OutputPath "C:\\IT\\PrinterScriptFiles"`

### Usage Examples

- Generate a configuration file at the default location: `.\Get-PrinterConfigurationXML.ps1`
- Generate a configuration file at a specific location: `.\Get-PrinterConfigurationXML.ps1 -OutputPath "C:\\IT\\PrinterScriptFiles"`

### File Headers
The CSV file generated will have the following headers:
` "Name","DriverName","PortName","Location",PrinterHostAddress,"InfName","PrintTicketXML" `

Each driver will be extracted to a zip file with the driver's display name.

Example 

Driver Name: 'RICOH MP 501 PCL 6'

File Name: 'RICOH MP 501 PCL 6.zip'
