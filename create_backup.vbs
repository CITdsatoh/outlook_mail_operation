Dim FileName
FileName="outlook_mail_dest_list.csv"

Dim FSObj
Set FSObj=CreateObject("Scripting.FileSystemObject")

If FSObj.FileExists(FileName) Then
 FSObj.CopyFile FileName,"outlook_mail_dest_list_backup.csv",True
End If