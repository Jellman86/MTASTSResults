@{
    RootModule           = 'MtaStsResults.psm1'
    ModuleVersion        = '1.0.0.0'
    GUID                 = 'd4c8f1e3-2b7a-4d6c-9e1f-5a3c2b8d7e4a'
    Author               = 'Scott Powdrill'
    Description          = 'PowerShell module for downloading and parsing MTA-STS/DMARC JSON reports from Microsoft Exchange using Microsoft Graph API'
    
    #RequiredModules = @(
        #@{ ModuleName = 'Microsoft.Graph'; ModuleVersion = '2.0.0' }
    #)
    
    FunctionsToExport    = @(
        'Invoke-DmarcAttachmentDownloader',
        'Invoke-JsonParse',
        'Invoke-CleanUp'
    )
    
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    
    PrivateData = @{
        PSData = @{
            Tags                       = @('DMARC', 'MTA-STS', 'Graph', 'Exchange', 'Reports')
            LicenseUri                 = 'https://github.com/Jellman86/MTASTSResults/blob/main/LICENSE'
            ProjectUri                 = 'https://github.com/Jellman86/MTASTSResults'
            ReleaseNotes               = 'Initial release - Convert script to module format with improved documentation'
        }
    }
}
