<div align="center">

## List all contacts stored in all Outlook contactfolders


</div>

### Description

This peace of code shows you how to obtain all the contacts as been stored in your Outlook pst-file. No matter how mutch contact-folders you have and how deep the three might be. (tested with Outlook 2000)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joost Rongen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joost-rongen.md)
**Level**          |Intermediate
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VBA MS Excel
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joost-rongen-list-all-contacts-stored-in-all-outlook-contactfolders__1-31603/archive/master.zip)





### Source Code

```
Private objApp As Outlook.Application
Private objNS As Outlook.NameSpace
Private objFolder As Outlook.MAPIFolder
Private objItem As Outlook.ContactItem
Private colAdressFolders As Collection
Sub Main()
 Dim lngLoop As Long
 Set objApp = New Outlook.Application
 Set objNS = objApp.GetNamespace("MAPI")
 Set colAdressFolders = New Collection
 Set objFolder = objNS.Folders.GetFirst  ' get root-folder
 ' recursive loop thrue all folders to collect the references to Adressbooks
 For lngLoop = 1 To objFolder.Folders.Count
  If objFolder.Folders.Item(lngLoop).DefaultItemType = olContactItem Then
   RecursiveSearch objFolder.Folders.Item(lngLoop), colAdressFolders
  End If
 Next lngLoop
 ' open every contact-folder and loop all entries
 For Each objFolder In colAdressFolders
  For lngLoop = 1 To objFolder.Items.Count
   Set objItem = objFolder.Items(lngLoop)
   Debug.Print objFolder.Name, objItem.FileAs
  Next lngLoop
 Next
End Sub
Private Sub RecursiveSearch(objSubFolder As Outlook.MAPIFolder, colAdrFolders As Collection)
On Error GoTo Errorhandler
Dim lngLoop As Long
 ' check for entries in this subfolder
 If objSubFolder.Items.Count > 0 Then
  'add reference to collection
   colAdrFolders.Add objSubFolder
 End If
 ' check for subfolders
 If objSubFolder.Folders.Count > 0 Then
   For lngLoop = 1 To objSubFolder.Folders.Count
    RecursiveSearch objSubFolder.Folders.Item(lngLoop), colAdrFolders
   Next lngLoop
 End If
Exit Sub
Errorhandler:
  MsgBox "An unexpected error occured methode RECURSIVESEARCH", vbCritical + vbOKOnly, "Problem"
  Err.Clear
End Sub
```

