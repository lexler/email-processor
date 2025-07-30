# Export emails from Outlook folder to .msg files

$outlook = New-Object -ComObject Outlook.Application
$folder = $outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders("PCIT Automatic")

# Copy items to array first to avoid collection modification during iteration
$mailItems = @()
foreach($mail in $folder.Items) {
    $mailItems += $mail
}

# Now process the copied items
foreach($mail in $mailItems) {
    $filename = "C:\Users\wilsora3\Documents\PCIT Automated\$($mail.ReceivedTime.ToString('yyyyMMdd_HHmmss')).msg"
    $mail.SaveAs($filename, 3)  # 3 = MSG format
    $mail.Move($outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders("PCIT-Processed"))
}

