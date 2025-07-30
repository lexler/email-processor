# Export emails from Outlook folder to .msg files

function Process-EmailBatch {
    param($sourceFolder, $processedFolder, $namespace)
    Write-Host "Starting batch"
    # Collect EntryIDs from current batch
    $entryIds = @()
    foreach($mail in $sourceFolder.Items) {
        $entryIds += $mail.EntryID
    }
    
    if ($entryIds.Count -eq 0) {
        return 0
    }
    
    Write-Host "Processing batch of $($entryIds.Count) emails"
    
    # Process each email by EntryID
    foreach($entryId in $entryIds) {
        $mail = $namespace.GetItemFromID($entryId)
        $filename = "C:\Users\wilsora3\Documents\PCIT Automated\$($mail.ReceivedTime.ToString('yyyyMMdd_HHmmss')).msg"
        $mail.SaveAs($filename, 3)  # 3 = MSG format
        $mail.Move($processedFolder)
    }
    
    return [int]$entryIds.Count
}

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inboxFolderId = 6
$sourceFolder = $namespace.GetDefaultFolder($inboxFolderId).Folders("PCIT Automatic")
$processedFolder = $namespace.GetDefaultFolder($inboxFolderId).Folders("PCIT-Processed")

$totalProcessed = 0
do {
    $batchCount = Process-EmailBatch -sourceFolder $sourceFolder -processedFolder $processedFolder -namespace $namespace
    $totalProcessed += $batchCount
    Write-Host "Total processed so far: $totalProcessed"
    Start-Sleep -Seconds 5
} while ($batchCount -gt 0)

Write-Host "Finished processing $totalProcessed emails total"

