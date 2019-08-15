cls
. "$PSScriptRoot\_pbi.functions.include.ps1"
. "$PSScriptRoot\_pbi.config.ps1"

$reportName = "Report 1"
$testDataModelDS = GetDatasetInGroup    -groupId $adminWorkspace `
                                        -datasetName $testDatasetName
$devDataModelDS = GetDatasetInGroup     -groupId $adminWorkspace `
                                        -datasetName $devDatasetName

if ( ReportExistsInGroup -groupId $testWorkspace -reportName $reportName ) 
{
    # If the report already exists, rebind it to its original dataset (the dev dataset)
    $report = GetReportInGroup  -groupId $testWorkspace `
                                -reportName $reportName
    RebindReportInGroup     -groupId $adminWorkspace `
                            -reportId $report.id `
                            -targetDatasetId $devDataModelDS.id   
}
$reportId = PostImportInGroup   -pbixFile "$PSScriptRoot\..\$reportName.pbix" `
                                -groupId $testWorkspace `
                                -nameConflict "CreateOrOverwrite"
RebindReportInGroup -groupId $adminWorkspace `
                    -reportId $reportId `
                    -targetDatasetId $testDataModelDS.id
