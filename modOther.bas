Attribute VB_Name = "modOther"
Option Explicit

Private Type SHFILEOPTSTRUCT
  hWnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As Long
End Type

Private Declare Function SHFileOperation Lib "Shell32.dll" _
  Alias "SHFileOperationA" (lpFileOp As SHFILEOPTSTRUCT) As Long
  
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.

Public Sub DeleteFileToRecycleBin(Filename As String)

Dim fop As SHFILEOPTSTRUCT

With fop
  .wFunc = FO_DELETE
  .pFrom = Filename
  .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
End With

SHFileOperation fop

End Sub

