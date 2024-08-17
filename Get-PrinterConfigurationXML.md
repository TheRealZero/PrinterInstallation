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
