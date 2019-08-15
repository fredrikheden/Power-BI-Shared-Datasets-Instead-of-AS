cls
. "$PSScriptRoot\_pbi.functions.include.ps1"
. "$PSScriptRoot\_pbi.config.ps1"

# Create/overwrite the test dataset, reconnect to the test db and refresh it.
PostImportInGroup   -pbixFile "$PSScriptRoot\Dev datamodel.pbix" `
                    -groupId $adminWorkspace `
                    -nameConflict "CreateOrOverwrite" `
                    -datasetDisplayName $testDatasetName 
$testDS = GetDatasetInGroup -groupId $adminWorkspace `
                            -datasetName $testDatasetName
UpdateDatasourcesInGroup_SqlServer  -groupId $adminWorkspace `
                                    -datasetId $testDS.id `
                                    -targetServer $testSqlServerName `
                                    -targetDatabase $testSqlServerDBName
RefreshDatasetInGroup  -groupId $adminWorkspace `
                        -datasetId $testDS.id `
                        -waitForRefreshToFinnish $True

