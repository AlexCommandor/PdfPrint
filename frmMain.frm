VERSION 5.00
Object = "{9C14D47E-B221-4FD0-AB86-67C30750874F}#5.0#0"; "STrayIco.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDF-автопринтер"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkA5 
      Caption         =   "Печатать всё на формат А5"
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Value           =   1  'Checked
      Width           =   4380
   End
   Begin VB.CheckBox chkDelNotBin 
      Caption         =   "Удалять PDF, минуя корзину"
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Value           =   1  'Checked
      Width           =   4380
   End
   Begin VB.ListBox lstJobs 
      Height          =   1815
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   4455
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   3960
      Top             =   1320
   End
   Begin VB.CheckBox chkHalftone 
      Caption         =   "Использовать полутона из документа"
      Height          =   285
      Left            =   117
      TabIndex        =   21
      Top             =   2925
      Value           =   1  'Checked
      Width           =   4380
   End
   Begin VB.Frame Frame2 
      Caption         =   "Формат выходного потока"
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   2895
      Begin VB.OptionButton optOUTPUT 
         Caption         =   "Binary"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optOUTPUT 
         Caption         =   "ASCII"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Уровень PostScript"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   2895
      Begin VB.OptionButton chkPS 
         Caption         =   "3"
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   17
         Top             =   260
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton chkPS 
         Caption         =   "2"
         Height          =   195
         Index           =   2
         Left            =   1260
         TabIndex        =   16
         Top             =   260
         Width           =   495
      End
      Begin VB.OptionButton chkPS 
         Caption         =   "1"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   260
         Width           =   495
      End
   End
   Begin VB.CommandButton BtnAbout 
      Caption         =   "О программе"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   720
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Left            =   4080
      Picture         =   "frmMain.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Выбрать каталог"
      Top             =   270
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.TextBox txtPathForSearch 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Остановить"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Запустить"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1520
      Width           =   1335
   End
   Begin VB.PictureBox picWait 
      Height          =   615
      Left            =   960
      Picture         =   "frmMain.frx":0544
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   3471
      Visible         =   0   'False
      Width           =   615
   End
   Begin SysTrayIcon.STrayIco STrayIco1 
      Left            =   0
      Top             =   3354
      _ExtentX        =   953
      _ExtentY        =   953
      TextToolTip     =   "PDF-автопринтер (остановлен)"
      Icon            =   "frmMain.frx":0986
      ErrorMessage    =   0   'False
   End
   Begin VB.PictureBox picStopped 
      Height          =   615
      Left            =   720
      Picture         =   "frmMain.frx":0DD8
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   3354
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picPrint 
      Height          =   495
      Index           =   0
      Left            =   -240
      Picture         =   "frmMain.frx":121A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picPrint 
      Height          =   495
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":165C
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   3471
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picPrint 
      Height          =   495
      Index           =   2
      Left            =   240
      Picture         =   "frmMain.frx":1A9E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   3354
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picPrint 
      Height          =   495
      Index           =   3
      Left            =   360
      Picture         =   "frmMain.frx":1EE0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   3237
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Папка, в которую сохраняются PDF-файлы"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Принтер, на который выводить PDF -файлы"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   3345
   End
   Begin VB.Menu MENU1 
      Caption         =   "MENU1"
      Visible         =   0   'False
      Begin VB.Menu mnStart 
         Caption         =   "Запустить"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnStop 
         Caption         =   "Остановить"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnAbout 
         Caption         =   "О программе"
      End
      Begin VB.Menu as 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Выход"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bProcessing As Boolean, oFS As Object, AF As New AFileLib
Dim wsNet As Object, wsShell As Object
Dim oFLD As Object, sFileName As String
Dim sPrn As String, nPSLevel As Long, nOUTData As Long, sHotFolder As String
Dim nUseDocHalftone As Integer, nDelNotBin As Integer
Dim oAApp As Object, oAVDoc As Object
Dim oAPDoc As Object, oAPPage As Object, oAPPoint As Object, i As Long
Dim tmpPrn As Printer, nAVDoc As Long, nPages As Long, lRes As Double
Dim sCurPrinter As String

Private Declare Function GetTickCount Lib "user32" () As Long

Private Sub BtnAbout_Click()
  frmAbout.Show
End Sub

Private Sub btnStart_Click()
    WriteParam
  Me.btnStart.Enabled = False: mnStart.Enabled = False
  Me.btnStop.Enabled = True: mnStop.Enabled = True
  Me.Combo1.Enabled = False
  Me.cmdBrowse.Enabled = False
  Me.txtPathForSearch.Enabled = False
  Me.Frame1.Enabled = False
  Me.Frame2.Enabled = False
  bProcessing = True
  Set oAApp = CreateObject("AcroExch.App")
  Set oAVDoc = CreateObject("AcroExch.AVDoc")
  For i = 1 To 3
    If chkPS(i).Value = True Then nPSLevel = i: Exit For ' PS Level?
  Next i
  nOUTData = Abs(optOUTPUT(1).Value) 'Binary data?
  Me.Timer1.Enabled = True
  Set Me.STrayIco1.Icon = picWait.Picture
  Me.STrayIco1.TextToolTip = "PDF-автопринтер (запущен)"
  Me.STrayIco1.Modify
  Set Me.Icon = picWait.Picture
  Me.WindowState = vbMinimized
End Sub

Private Sub btnStop_Click()
  Me.btnStart.Enabled = True: mnStart.Enabled = True
  Me.btnStop.Enabled = False: mnStop.Enabled = False
  Me.Combo1.Enabled = True
  Me.cmdBrowse.Enabled = True
  Me.txtPathForSearch.Enabled = True
  Me.Frame1.Enabled = True
  Me.Frame2.Enabled = True
  bProcessing = False
  Me.Timer1.Enabled = False
  Set oAPDoc = Nothing
  Set oAVDoc = Nothing
  Set oAApp = Nothing
  Set Me.STrayIco1.Icon = picStopped.Picture
  Me.STrayIco1.TextToolTip = "PDF-автопринтер (остановлен)"
  Me.STrayIco1.Modify
  Set Me.Icon = picStopped.Picture
End Sub

Private Sub cmdBrowse_Click()
  Dim sNewDir As String, sInitPath As String
  If oFS.FolderExists(Me.txtPathForSearch.Text) Then
    sInitPath = Me.txtPathForSearch.Text
  Else
    sInitPath = wsShell.SpecialFolders(0)
  End If
  sNewDir = AF.Z_BrowseForFolder(Me.hWnd, _
        "Укажите папку, в которую будут сохраняться PDF-файлы", sInitPath)
  If Len(sNewDir) > 0 Then
    txtPathForSearch.Text = sNewDir
    wsShell.RegWrite "HKLM\Software\PDF-autoprinter\PDF folder", sNewDir
    Me.btnStart.Enabled = (Me.Combo1.Text <> vbNullString)
    mnStart.Enabled = btnStart.Enabled
  End If
  WriteParam
End Sub

Private Sub Combo1_Click()
  Me.btnStart.Enabled = (Me.Combo1.ListIndex >= 0) And _
        (Len(Me.txtPathForSearch.Text) > 0)
  mnStart.Enabled = btnStart.Enabled
  If btnStart.Enabled Then _
      wsShell.RegWrite "HKLM\Software\PDF-autoprinter\Printer for output", Me.Combo1.Text
  sCurPrinter = Me.Combo1.Text
  WriteParam
End Sub

Private Sub Combo1_DropDown()
  Me.Combo1.Clear
  For Each tmpPrn In Printers
    Me.Combo1.AddItem tmpPrn.DeviceName
  Next
  Me.Combo1.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WriteParam
    End
End Sub

Private Sub mnAbout_Click()
  frmAbout.Show
End Sub

Private Sub Timer1_Timer()
  Dim bFitToPage As Boolean, dTime1 As Long, dTime2 As Long, sJobName As String
  On Error Resume Next
  If bProcessing = False Then Exit Sub
  If Not oFS.FolderExists(Me.txtPathForSearch.Text) Then Exit Sub
  If Len(sCurPrinter) = 0 Then Exit Sub
    Set oFLD = oFS.GetFolder(Me.txtPathForSearch.Text)
    sFileName = Dir(oFS.BuildPath(oFLD.Path, "*.pdf"), vbNormal Or _
        vbHidden Or vbSystem Or vbReadOnly)
    If (Len(sFileName) <> 0) Then
      dTime1 = GetTickCount / 1000
      Do While Not FileExists(oFLD.Path & "\" & sFileName)
        dTime2 = GetTickCount / 1000
        If Abs(dTime2 - dTime1) > 20 * 60 Then Exit Sub
        DoEvents
      Loop
      Set Me.STrayIco1.Icon = picPrint(0).Picture
      Me.STrayIco1.TextToolTip = "PDF-автопринтер (настройка принтера)"
      Me.STrayIco1.Modify
      Set Me.Icon = picPrint(0).Picture
      
      sPrn = Printer.DeviceName
      For Each tmpPrn In Printers
        If tmpPrn.DeviceName = sCurPrinter Then
          wsNet.SetDefaultPrinter tmpPrn.DeviceName
          Exit For
        End If
      Next
      sJobName = vbNullString
      Do While Len(sFileName) <> 0
        sJobName = sFileName & "..."
        Call Sleep(1000) 'This pause is required to avoid Acrobat error "File not found" after Distiller error
        Timer2.Interval = 20000
        Timer2.Enabled = True
        If Not FileExists(oFLD.Path & "\" & sFileName) Then
          Me.lstJobs.AddItem Format$(Me.lstJobs.ListCount - 1, "0000") & ". " & sJobName & "ERROR " & Now()
          sJobName = vbNullString
          Call AddHorScrollToList(Me.lstJobs)
          Exit Do
        End If
        oAVDoc.Open oFLD.Path & "\" & sFileName, ""
        Timer2.Enabled = False
        Set Me.STrayIco1.Icon = picPrint(1).Picture
        Me.STrayIco1.TextToolTip = "PDF-автопринтер (обработка " & _
              oFLD.Path & "\" & sFileName & ")"
        Me.STrayIco1.Modify
        Set Me.Icon = picPrint(1).Picture
        
        Set oAPDoc = oAVDoc.GetPDDoc
        nPages = oAPDoc.GetNumPages
        Set oAPPage = oAPDoc.AcquirePage(0)
        Set oAPPoint = oAPPage.GetSize
        If oAPPoint.x < oAPPoint.y Then 'orientation - Portraitt
         If Me.chkA5.Value = 0 Then
            If oAPPoint.x < 595 And oAPPoint.y < 842 Then
              Call SetOrientation(Me.hWnd, 1, 1, 9) '1st - duplex, 2nd - orient(1-Port., 2-Lands.), 3rd - Size (8-A3,9-A4, 11-A5)
              bFitToPage = False
            ElseIf oAPPoint.x < 842 And oAPPoint.y < 1191 Then
              Call SetOrientation(Me.hWnd, 1, 1, 8) '1st - duplex, 2nd - orient(1-Port., 2-Lands.), 3rd - Size (8-A3,9-A4, 11-A5)
              bFitToPage = False
            Else
              Call SetOrientation(Me.hWnd, 1, 1, 8) '1st - duplex, 2nd - orient(1-Port., 2-Lands.), 3rd - Size (8-A3,9-A4, 11-A5)
              bFitToPage = True
            End If
         Else 'A5!!!
            Call SetOrientation(Me.hWnd, 1, 1, 11) '1st - duplex, 2nd - orient(1-Port., 2-Lands.), 3rd - Size (8-A3,9-A4, 11-A5)
            bFitToPage = True
         End If
        Else 'Landscape
         If Me.chkA5.Value = 0 Then
            If oAPPoint.x < 595 And oAPPoint.y < 842 Then
              Call SetOrientation(Me.hWnd, 1, 2, 9) '1st - duplex, 2nd - orient(1-Port., 2-Lands.), 3rd - Size (8-A3,9-A4, 11-A5)
              bFitToPage = False
            ElseIf oAPPoint.x < 842 And oAPPoint.y < 1191 Then
              Call SetOrientation(Me.hWnd, 1, 2, 8) '1st - duplex, 2nd - orient(1-Port., 2-Lands.), 3rd - Size (8-A3,9-A4, 11-A5)
              bFitToPage = False
            Else
              Call SetOrientation(Me.hWnd, 1, 2, 8) '1st - duplex, 2nd - orient(1-Port., 2-Lands.), 3rd - Size (8-A3,9-A4, 11-A5)
              bFitToPage = True
            End If
         Else
            Call SetOrientation(Me.hWnd, 1, 2, 11) '1st - duplex, 2nd - orient(1-Port., 2-Lands.), 3rd - Size (8-A3,9-A4, 11-A5)
            bFitToPage = True
         End If
        End If
        
        'long PrintPagesSilentEx(long nFirstPage, long nLastPage,
              'long nPSLevel, long bBinaryOk,
              'long bShrinkToFit, long bReverse,
              'long bFarEastFontOpt,
              'long bEmitHalftones,
              'long iPageOption);
        'PDAllPages -3 All pages of a document.
        'PDOddPagesOnly -4 Odd pages of a document.
        'PDEvenPagesOnly -5 Even pages of a document.
        oAVDoc.PrintPagesSilentEx _
            0, nPages - 1, nPSLevel, nOUTData, _
            Abs(bFitToPage), 0, 0, chkHalftone.Value, -3
          
        Me.lstJobs.AddItem Format$(Me.lstJobs.ListCount + 1, "00000") & " " & sJobName & "OK " & Now()
        sJobName = vbNullString
        Call AddHorScrollToList(Me.lstJobs)
        
        Set Me.STrayIco1.Icon = picPrint(2).Picture
        Me.STrayIco1.TextToolTip = "PDF-автопринтер (удаление " & _
              oFLD.Path & "\" & sFileName & ")"
        Me.STrayIco1.Modify
        Set Me.Icon = picPrint(2).Picture
        Set oAPPoint = Nothing
        Set oAPPage = Nothing
        oAPDoc.Close
        Set oAPDoc = Nothing
        oAVDoc.Close 0
        
        If Me.chkDelNotBin.Value <> 0 Then
            oFS.DeleteFile oFLD.Path & "\" & sFileName, True
        Else
            DeleteFileToRecycleBin oFLD.Path & "\" & sFileName
        End If
        
        sFileName = Dir()  ' Get next file
        
        Set Me.STrayIco1.Icon = picPrint(3).Picture
        Me.STrayIco1.TextToolTip = "PDF-автопринтер (печать окончена)"
        Me.STrayIco1.Modify
        Set Me.Icon = picPrint(3).Picture
        DoEvents
      Loop
      oAApp.CloseAllDocs
            
      For Each tmpPrn In Printers
        If InStr(tmpPrn.DeviceName, sPrn) > 0 Then
          wsNet.SetDefaultPrinter tmpPrn.DeviceName
          Exit For
        End If
      Next
    End If
    Set Me.STrayIco1.Icon = picWait.Picture
    Me.STrayIco1.TextToolTip = "PDF-автопринтер (запущен)"
    Me.STrayIco1.Modify
    Set Me.Icon = picWait.Picture
    DoEvents
  On Error GoTo 0
End Sub

Private Sub Form_Load()
  Dim bStart As Boolean
  On Error Resume Next
  If App.PrevInstance Then End
  Screen.MousePointer = vbHourglass
  Me.STrayIco1.Create
  If Err.Number <> 0 Then
    MsgBox "Невозможно создать объект STrayIcon! Переустановите программу!", vbCritical, "PDF-автопринтер"
    End
  End If
  
  Set oAApp = CreateObject("AcroExch.App")
  Set oAVDoc = CreateObject("AcroExch.AVDoc")
  If Err.Number <> 0 Then
    MsgBox "Невозможно создать объект Adobe Acrobat! Проверьте правильность его установки!", vbCritical, "PDF-автопринтер"
    End
  End If
  Set oAApp = Nothing
  Set oAPDoc = Nothing
  Set oAVDoc = Nothing
  
  Set oFS = CreateObject("Scripting.FileSystemObject")
  Set wsNet = CreateObject("WScript.Network")
  Set wsShell = CreateObject("WScript.Shell")
  If Err.Number <> 0 Then
    MsgBox "Невозможно создать объект Windows Scripting Host! Проверьте его наличие в компонентах Windows!", vbCritical, "PDF-автопринтер"
    End
  End If

  'Me.txtPathForSearch.Text = wsShell.RegRead("HKLM\Software\PDF-autoprinter\PDF folder")
  'sCurPrinter = wsShell.RegRead("HKLM\Software\PDF-autoprinter\Printer for output")
  
  Call ReadParam
  
  bStart = False
  If Len(sCurPrinter) > 0 And Len(Me.txtPathForSearch.Text) > 0 Then
    Me.Combo1.Clear
    For Each tmpPrn In Printers
      Me.Combo1.AddItem tmpPrn.DeviceName
    Next
    For Each tmpPrn In Printers
      If tmpPrn.DeviceName = sCurPrinter Then
        Me.Combo1.Text = sCurPrinter
        bStart = True
        Exit For
      End If
    Next
  End If
  Err.Clear
  On Error GoTo 0
  Screen.MousePointer = vbNormal
  If bStart Then Call btnStart_Click
End Sub

Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then
    Me.Hide
    Me.WindowState = vbNormal
  End If
End Sub

Private Sub mnExit_Click()
  End
End Sub

Private Sub mnStart_Click()
  Call btnStart_Click
End Sub

Private Sub mnStop_Click()
  Call btnStop_Click
End Sub

Private Sub STrayIco1_LeftDblClick()
  Me.Show
End Sub

Private Sub STrayIco1_RightClick()
  PopupMenu MENU1
End Sub

Private Function FileExists(sFileName As String) As Boolean
  Dim iFNum As Integer
  On Error Resume Next
  Err.Clear: iFNum = FreeFile
  Open sFileName For Input As iFNum
  If Err.Number <> 0 Then 'File may be not exists or access denied!
    Err.Clear: FileExists = False
  Else
    Close #iFNum: FileExists = True
  End If
End Function

Private Sub Timer2_Timer()
  Me.Timer2.Enabled = False
  Call btnStop_Click
  Call btnStart_Click
End Sub

Private Sub ReadParam()
    Dim sPrn1 As String, nPSLevel1 As Long, nOUTData1 As Long, sHotFolder1 As String
    Dim nUseDocHalftone1 As Integer, nDelNotBin1 As Integer
    
    On Error Resume Next
    
    sHotFolder1 = wsShell.RegRead("HKLM\Software\PDF-autoprinter\PDF folder")
    sPrn1 = wsShell.RegRead("HKLM\Software\PDF-autoprinter\Printer for output")
    nPSLevel1 = wsShell.RegRead("HKLM\Software\PDF-autoprinter\PSLevel")
    nOUTData1 = wsShell.RegRead("HKLM\Software\PDF-autoprinter\Data format")
    nUseDocHalftone1 = wsShell.RegRead("HKLM\Software\PDF-autoprinter\UseDocHalftone")
    nDelNotBin1 = wsShell.RegRead("HKLM\Software\PDF-autoprinter\DeleteNotBin")
  
    Me.txtPathForSearch.Text = sHotFolder1
    sCurPrinter = sPrn1
  
    If nPSLevel1 < 1 Or nPSLevel1 > 3 Then
        chkPS(2).Value = True
        nPSLevel = 2
    Else
        chkPS(nPSLevel1).Value = True
        nPSLevel = nPSLevel1
    End If
    If nOUTData1 > 1 Or nOUTData1 < 0 Then
        nOUTData = 0
        optOUTPUT(0).Value = True
    Else
        nOUTData = nOUTData1
        optOUTPUT(nOUTData).Value = True
    End If
    Me.chkHalftone.Value = nUseDocHalftone1
    nUseDocHalftone = nUseDocHalftone1
    Me.chkDelNotBin = nDelNotBin1
    nDelNotBin = nDelNotBin1
    
    Err.Clear
End Sub

Private Sub WriteParam()
    Dim i As Integer
    
    On Error Resume Next
    
    wsShell.RegWrite "HKLM\Software\PDF-autoprinter\PDF folder", Me.txtPathForSearch.Text
    wsShell.RegWrite "HKLM\Software\PDF-autoprinter\Printer for output", Me.Combo1.Text
    sCurPrinter = Me.Combo1.Text

    For i = 1 To 3
      If chkPS(i).Value = True Then nPSLevel = i: Exit For ' PS Level?
    Next i
    nOUTData = Abs(optOUTPUT(1).Value) 'Binary data?

    wsShell.RegWrite "HKLM\Software\PDF-autoprinter\PSLevel", nPSLevel
    wsShell.RegWrite "HKLM\Software\PDF-autoprinter\Data format", nOUTData
    
    nUseDocHalftone = Me.chkHalftone.Value
    wsShell.RegWrite "HKLM\Software\PDF-autoprinter\UseDocHalftone", nUseDocHalftone
    nDelNotBin = Me.chkDelNotBin.Value
    wsShell.RegWrite "HKLM\Software\PDF-autoprinter\DeleteNotBin", nDelNotBin
      
    Err.Clear
End Sub

