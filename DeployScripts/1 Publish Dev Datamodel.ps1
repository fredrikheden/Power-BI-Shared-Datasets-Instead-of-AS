cls
. "$PSScriptRoot\_pbi.functions.include.ps1"
. "$PSScriptRoot\_pbi.config.ps1"
    
# Upload the dev data model to the admin  workspace
PostImportInGroup   -pbixFile "$PSScriptRoot\..\Dev datamodel.pbix" `
                    -groupId $adminWorkspace -nameConflict "CreateOrOverwrite"
