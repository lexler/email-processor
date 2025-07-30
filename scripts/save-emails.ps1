# Export emails from Outlook folder to .msg files

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$sourceFolder = $namespace.GetDefaultFolder(6).Folders("PCIT Automatic")
$processedFolder = $namespace.GetDefaultFolder(6).Folders("PCIT-Processed")

# Collect EntryIDs first to avoid collection modification during iteration
$entryIds = @()
foreach($mail in $sourceFolder.Items) {
    $entryIds += $mail.EntryID
}

# Process each email by EntryID
foreach($entryId in $entryIds) {
    $mail = $namespace.GetItemFromID($entryId)
    $filename = "C:\Users\wilsora3\Documents\PCIT Automated\$($mail.ReceivedTime.ToString('yyyyMMdd_HHmmss')).msg"
    $mail.SaveAs($filename, 3)  # 3 = MSG format
    $mail.Move($processedFolder)
}

