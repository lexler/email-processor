# Export emails from Outlook folder to .msg files

$outlook = New-Object -ComObject Outlook.Application
$folder = $outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders("PCIT Automatic")

foreach($mail in $folder.Items) {
    $filename = "C:\Users\wilsora3\Documents\PCIT Automated\$($mail.ReceivedTime.ToString('yyyyMMdd_HHmmss')).msg"
    $mail.SaveAs($filename, 3)  # 3 = MSG format
    $mail.Move($outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders("PCIT-Processed"))
}

