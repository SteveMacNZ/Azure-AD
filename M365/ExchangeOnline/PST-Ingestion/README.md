# PST Ingestion to Exchange Online

 * Get-PSTSize.ps1 used to return PST files sizes in GB from local directory prior to uploading to Azure Blob storage using AZ Copy
 * Invoke-CustomPSTImport.ps1 used to loop through CSV file and import PSTs from Custom Azure Blob URI
 * ###-PstImportMappingFile.csv used as the mapping file for ingestion of the PST files via complaince centre