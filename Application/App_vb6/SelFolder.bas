Attribute VB_Name = "SelFolder"
Option Explicit
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
        "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
        "SHGetPathFromIDListA" (ByVal pidl As Long, _
        ByVal pszPath As String) As Long

Public Const BIF_RETURNONLYFSDIRS = &H1

Type BROWSEINFO
   hOwner As Long
   pidlRoot As Long
   pszDisplayName  As String
   lpszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type

Type SHITEMID
   cb As Long
   abID As Byte
End Type

Type ITEMIDLIST
   mkid As SHITEMID
End Type
'-- End --'



