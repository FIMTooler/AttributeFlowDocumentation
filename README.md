# AttributeFlowDocumentation
Scripts for documenting attributes flows with MIIS/ILM/FIM/MIM, Now Azure AD Connnect - AADC

## Recent Changes
Using the base of the FIM2AFCSV script, I have now added one for AADC.  The new script is AADCAF2CSV.ps1

AFComparison will compare 2 server exports and attempt to compare the attribute flows between them.  Ideal for comparing Dev and Prod attribute flows.  Fairly new script, but will at least narrow down where possible difference exist so you can manually compare environments in the Sync UI or using FIMAF2CSV.

## About
These two scripts are useful for documenting attribute flows from a Server Export of the Synchronization Engine.  As noted by their names, one script outputs data into a CSV which can be used with Excel.  The other uses Visio 2016 to draw a visual representation of attribute flows in and out of the Metaverse based on each attribute in the Metaverse schema.

The critical methods of parsing the XML export files is based on methods from the Get-FimSyncConfiguration.psm1 that can be found here - https://archive.codeplex.com/?p=fimpowershellmodule.

Improvements have been made to the Get-ImportAttributeFlow and Get-ExportAttributeFlow methods to allow filtering based on the Metaverse objectType.  The Get-ImportToExportAttributeFlow method is loosely based on an older and different version of the Join-ImportToExportAttributeFlow method.
