# Export emails from Outlook folder to .msg files

$outlook = New-Object -ComObject Outlook.Application
$folder = $outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders("PCIT Automatic")

# Collect EntryIDs first to avoid collection modification during iteration
$entryIds = @()
foreach($mail in $folder.Items) {
    $entryIds += $mail.EntryID
}

# Now process each email by looking it up by EntryID
$namespace = $outlook.GetNamespace("MAPI")
$processedFolder = $namespace.GetDefaultFolder(6).Folders("PCIT-Processed")

foreach($entryId in $entryIds) {
    $mail = $namespace.GetItemFromID($entryId)
    $filename = "C:\Users\wilsora3\Documents\PCIT Automated\$($mail.ReceivedTime.ToString('yyyyMMdd_HHmmss')).msg"
    $mail.SaveAs($filename, 3)  # 3 = MSG format
    $mail.Move($processedFolder)
}

