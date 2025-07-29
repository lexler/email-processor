# Export emails from Outlook folder to .msg files
$outlook = New-Object -ComObject Outlook.Application
$folder = $outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders("PCIT-ToProcess")

foreach($mail in $folder.Items) {
    $filename = "C:\EmailProcessing\inbox\$($mail.ReceivedTime.ToString('yyyyMMdd_HHmmss')).msg"
    $mail.SaveAs($filename, 3)  # 3 = MSG format
    $mail.Move($outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders("PCIT-Processed"))
}