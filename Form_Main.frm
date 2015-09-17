VERSION 5.00
Begin VB.Form Form_Main 
   Caption         =   "工资条发送器"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form_Main"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7245
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClose 
      Caption         =   "退出"
      Height          =   855
      Left            =   7080
      TabIndex        =   32
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton btn_ShowExcel 
      Caption         =   "显示/隐藏Excel"
      Height          =   615
      Left            =   6960
      TabIndex        =   31
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CheckBox IsTestSend 
      Caption         =   "测试"
      Height          =   375
      Left            =   6960
      TabIndex        =   30
      Top             =   2400
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "取消发送"
      Height          =   615
      Left            =   6840
      TabIndex        =   29
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Timer TaskCheckTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7320
      Top             =   1560
   End
   Begin VB.TextBox txMessage 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   6000
      Width           =   6375
   End
   Begin VB.TextBox txSubject 
      Height          =   405
      Left            =   2280
      TabIndex        =   27
      Text            =   "$NAME$-$YEAR$-$MONTH$工资明细"
      Top             =   5400
      Width           =   4215
   End
   Begin VB.TextBox txTempPath 
      BackColor       =   &H80000004&
      Height          =   405
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   4800
      Width           =   5295
   End
   Begin VB.TextBox txAttachRange 
      Height          =   405
      Left            =   1200
      TabIndex        =   25
      Text            =   "B2:D3,I2:Y3"
      Top             =   600
      Width           =   1950
   End
   Begin VB.TextBox txMonth 
      BackColor       =   &H80000004&
      Height          =   405
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   4200
      Width           =   4575
   End
   Begin VB.TextBox txYear 
      BackColor       =   &H80000004&
      Height          =   405
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox txCheck 
      BackColor       =   &H80000004&
      Height          =   405
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3000
      Width           =   4575
   End
   Begin VB.TextBox txMail 
      BackColor       =   &H80000004&
      Height          =   405
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2400
      Width           =   4575
   End
   Begin VB.TextBox txPass 
      BackColor       =   &H80000004&
      Height          =   405
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1800
      Width           =   4575
   End
   Begin VB.TextBox txMonthRange 
      Height          =   405
      Left            =   1200
      TabIndex        =   18
      Text            =   "C3"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txYearRange 
      Height          =   405
      Left            =   1200
      TabIndex        =   17
      Text            =   "B3"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txCheckRange 
      Height          =   405
      Left            =   1200
      TabIndex        =   16
      Text            =   "B3"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txMailRange 
      Height          =   405
      Left            =   1200
      TabIndex        =   15
      Text            =   "AB3"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txPassRange 
      Height          =   405
      Left            =   1200
      TabIndex        =   14
      Text            =   "AA3"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txName 
      BackColor       =   &H80000004&
      Height          =   405
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   4575
   End
   Begin VB.TextBox txNameRange 
      Height          =   405
      Left            =   1200
      TabIndex        =   4
      Text            =   "D3"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txExcelPath 
      BackColor       =   &H80000004&
      Height          =   405
      Left            =   1080
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   0
      Width           =   5415
   End
   Begin VB.CommandButton BtnSend 
      Caption         =   "发送"
      Height          =   735
      Left            =   6840
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "发送内容"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "邮件主题"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "邮件模板"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "临时文件夹"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4815
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "月"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4245
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "年"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3660
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "检测有效"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3075
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "邮箱"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2505
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "密码"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "姓名"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Excel文件"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetModuleFileName Lib "KERNEL32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "KERNEL32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    
Private ThisWorkbook As Object
Private CurrentSheet As Object
Private excelApp As Object    'Excel的应用程序
Private outlookApp As Object
Private StopSend As Boolean

Dim ProjectName As String
Dim ThisINIFile  As String

Private Sub btn_ShowExcel_Click()
     If excelApp = Null Then
          MsgBox "未安装Microsoft Excel !!!!!!!!!!!!"
    Else
        If excelApp.Visible Then
            excelApp.Visible = False
        Else
            excelApp.Visible = True
        End If
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub BtnSend_Click()
    start_timer
End Sub



Private Sub btnStop_Click()
    stop_timer
End Sub

Private Sub start_timer()
    If Not TaskCheckTimer.Enabled Then
        TaskCheckTimer.Enabled = True
    End If
    BtnSend.Enabled = False
    btnStop.Enabled = True
    StopSend = False
End Sub

Private Sub stop_timer()
    If TaskCheckTimer.Enabled Then
        TaskCheckTimer.Enabled = False
    End If
    BtnSend.Enabled = True
    btnStop.Enabled = False
    StopSend = True
End Sub


Private Sub Form_Initialize()
    ProjectName = "WagesSend"
    ThisINIFile = ProjectName + ".ini"

    LoadIni
    Set outlookApp = CreateObject("Outlook.Application")
    If outlookApp = Null Then
        MsgBox "未安装Microsoft Outlook!!!!!!!!!!!!"
    End If
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    If excelApp = Null Then
        MsgBox "未安装Microsoft Excel !!!!!!!!!!!!"
    End If
    StopSend = True
End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
     
    For Each file In Data.Files
        I = InStrRev(file, ".")
        If UCase(Mid(file, I + 1)) = UCase("xlsm") Or UCase(Mid(file, I + 1)) = UCase("xlsx") Then
            txExcelPath.Text = file
            LoadExcelFile (file)
            Exit Sub
        End If
    Next
End Sub


Private Sub LoadExcelFile(ExcelPath As String)
    Set ThisWorkbook = excelApp.Workbooks.Open(ExcelPath, , VARIANT_FALSE)
    Set CurrentSheet = ThisWorkbook.ActiveSheet
    FlushUI
    txTempPath.Text = ThisWorkbook.Path + "\~tmp\"
    dirs = Dir(txTempPath.Text, vbDirectory)
    If dirs = "" Then
        MkDir txTempPath.Text
    End If
 End Sub

Private Sub Form_Terminate()


    If (Not ThisWorkbook Is Nothing) Then
        ThisWorkbook.Save
    End If
    
    If (Not excelApp Is Nothing) And (excelApp <> Null) Then
        excelApp.Quit
    End If
    If (Not outlookApp Is Nothing) And (outlookApp <> Null) Then
        outlookApp.Quit
    End If
End Sub

Private Sub FlushUI()
    If IsLeftRowSend Then
        Dim strTmp As String
        Dim StringCount As Variant
        Dim Rng  As Object
        Dim DestName As String
        
        txPass.Text = CurrentSheet.Range(txPassRange.Text).Text
        txMail.Text = CurrentSheet.Range(txMailRange.Text).Text
        txCheck.Text = CurrentSheet.Range(txCheckRange.Text).Text
        txYear.Text = CurrentSheet.Range(txYearRange.Text).Text
        txMonth.Text = CurrentSheet.Range(txMonthRange.Text).Text
        txName.Text = CurrentSheet.Range(txNameRange.Text).Text
        
    End If

End Sub

Function IsLeftRowSend() As Boolean
    Dim rngTest As Object
    Set rngTest = CurrentSheet.Range(txCheckRange.Text)
   
    If VBA.IsNumeric(rngTest.Value) Then
        IsLeftRowSend = True
    Else
        IsLeftRowSend = False
    End If
   
End Function



Private Sub Form_Unload(Cancel As Integer)
    SaveIni
End Sub

Private Sub TaskCheckTimer_Timer()
    
    If (Not IsLeftRowSend) Or StopSend Then
         stop_timer
        Exit Sub
    End If
    
    sendOneMail
    
    If IsTestSend.Value = 1 Then
         stop_timer
    End If
    
    FlushUI
    
End Sub

Private Function getPostFixNumber(s As String) As Integer
    getPostFixNumber = 0
    L = Len(s)
    For I = L To 0 Step -1
        If Not IsNumeric(Mid(s, I, 1)) Then
            S1 = Mid(s, I + 1)
            getPostFixNumber = Int(Val(S1))
            If getPostFixNumber <> 0 Then
                Exit Function
            End If
                
        End If
    Next
End Function

Private Sub sendOneMail()
    Dim StringCount As Variant
    Dim iRanges() As Object
    Dim TmpFile As String
    If outlookApp = Null Then
        MsgBox "未安装Microsoft Outlook!!!!!!!!!!!!"
        Exit Sub
    End If
    
    excelApp.Application.DisplayAlerts = False
    
    If txMail.Text <> "" And txName.Text <> "" Then

        RangeStrings = Split(txAttachRange.Text, ",")
        ReDim Preserve iRanges(LBound(RangeStrings) To UBound(RangeStrings)) As Object
        For I = LBound(RangeStrings) To UBound(RangeStrings)
            Set iRanges(I) = CurrentSheet.Range(RangeStrings(I))
        Next
        
        TmpFile = txTempPath.Text + txName.Text + ".xlsx"
        MakeRangeExcel TmpFile, txPass.Text, txName.Text, iRanges
            
            
        
        SendExcel TmpFile, txYear.Text, txMonth.Text, txMail.Text, txName.Text
            
    End If
    
    Dim delRow  As Integer
    delRow = getPostFixNumber(txAttachRange.Text)
    If delRow <> 0 Then
       DeleteSentRow delRow
    End If
    
End Sub

Private Sub MakeRangeExcel(NewFile As String, strPass As String, strName As String, iRanges As Variant)
    Dim newBook As Object
    Dim newSheet As Object
    Dim rngBegin As Object
    Dim rngEnd As Object
    Dim rngDest As Object
    
    
    Set newBook = excelApp.Workbooks.Add
    Set newSheet = newBook.ActiveSheet
    
    Dim Col As Integer
    Col = 0
    For I = LBound(iRanges) To UBound(iRanges)
        Set ThisRange = iRanges(I)
        
        Set rngBegin = newSheet.cells(1, Col + 1)
        Set rngEnd = newSheet.cells(ThisRange.Rows.Count, Col + ThisRange.Columns.Count)
        Set rndDest = newSheet.Range(rngBegin, rngEnd)
    
        iRanges(I).Copy
        newSheet.Paste (rndDest)
        Col = Col + ThisRange.Columns.Count
    Next
    
    If Len(strPass) <> 0 Then
        newBook.Password = strPass
    End If
    newBook.SaveAs NewFile, , , , , , , xlLocalSessionChanges
    newBook.Close False
    
End Sub



Private Function SendExcel(strExcelFile As String, Year As String, Month As String, strEmail As String, strName As String) As Boolean

    Dim strSubject As String
    Dim strContent As String
    

    strSubject = Replace(txSubject.Text, "$NAME$", strName)
    strSubject = Replace(strSubject, "$YEAR$", Year)
    strSubject = Replace(strSubject, "$MONTH$", Month)
    
'    %USER_NAME% 您好！
'    现将 $YEAR$-$MONTH$工资明细  发送给您
'    如有问题您可以随时与财务部联系，我们会尽快处理您的问题并答复结果。
'                                 财务部 $NOW_TIME$
    
    strContent = Replace(txMessage.Text, "$NAME$", strName)
    
    strContent = Replace(strContent, "$YEAR$", Year)
    strContent = Replace(strContent, "$MONTH$", Month)
    strContent = Replace(strContent, "$NOW_TIME$", Now)
    
    Set oItemMail = outlookApp.CreateItem(olMailItem)
    With oItemMail
            .Subject = strSubject
            .Body = strContent
            .Attachments.Add (strExcelFile)
            .Importance = olImportanceHigh
            .Recipients.Add (strEmail)
            .Sensitivity = olPersonal
            .Send
    End With
    Set oItemMail = Nothing
    SendExcel = True
End Function


Private Sub DeleteSentRow(rowIndex As Integer)
    CurrentSheet.Rows(rowIndex).Delete
End Sub

Private Sub LoadIni()
        txAttachRange.Text = ReadINI("OPTIONS", "SEND_RANGE", ThisINIFile, "B2:D3,I2:Y3")
        txNameRange.Text = ReadINI("OPTIONS", "NAME_RANGE", ThisINIFile, "D3")
        txPassRange.Text = ReadINI("OPTIONS", "PASSR_ANGE", ThisINIFile, "AA3")
        txMailRange.Text = ReadINI("OPTIONS", "MAIL_RANGE", ThisINIFile, "AB3")
        txCheckRange.Text = ReadINI("OPTIONS", "CHECK_RANGE", ThisINIFile, "B3")
        txYearRange.Text = ReadINI("OPTIONS", "YEAR_RANGE", ThisINIFile, "B3")
        txMonthRange.Text = ReadINI("OPTIONS", "MONTH_RANGE", ThisINIFile, "C3")
        txSubject.Text = ReadINI("OPTIONS", "SUBJECT_RANGE", ThisINIFile, "$NAME$-$YEAR$-$MONTH$工资明细")
        msgStr = ReadINI("OPTIONS", "MESSAGE_RANGE", ThisINIFile, "$NAME$ 您好！\r\n现将 $YEAR$-$MONTH$工资明细  发送给您 \r\n如有问题您可以随时与财务部联系，我们会尽快处理您的问题并答复结果。\r\n财务部 $NOW_TIME$ ")
        txMessage.Text = Replace(msgStr, "\r\n", "" + Chr(13) + Chr(10))
End Sub

Private Sub SaveIni()
        WriteINI "OPTIONS", "SEND_RANGE", txAttachRange.Text, ThisINIFile
        WriteINI "OPTIONS", "NAME_RANGE", txNameRange.Text, ThisINIFile
        WriteINI "OPTIONS", "PASSR_ANGE", txPassRange.Text, ThisINIFile
        WriteINI "OPTIONS", "MAIL_RANGE", txMailRange.Text, ThisINIFile
        WriteINI "OPTIONS", "CHECK_RANGE", txCheckRange.Text, ThisINIFile
        WriteINI "OPTIONS", "YEAR_RANGE", txYearRange.Text, ThisINIFile
        WriteINI "OPTIONS", "MONTH_RANGE", txMonthRange.Text, ThisINIFile
        WriteINI "OPTIONS", "SUBJECT_RANGE", txSubject.Text, ThisINIFile
        WriteINI "OPTIONS", "MESSAGE_RANGE", Replace(txMessage.Text, "" + Chr(13) + Chr(10), "\r\n"), ThisINIFile
End Sub
Sub WriteINI(wiSection As String, wiKey As String, wiValue As String, wiFile As String)
    WritePrivateProfileString wiSection, wiKey, wiValue, App.Path & "\" & wiFile
End Sub

Function ReadINI(riSection As String, riKey As String, riFile As String, riDefault As String)

    Dim sRiBuffer As String
    Dim sRiValue As String
    Dim sRiLong As String
    Dim INIFile As String
    INIFile = App.Path & "\" & riFile
    If Dir(INIFile) <> "" Then
        sRiBuffer = String(255, vbNull)
        sRiLong = GetPrivateProfileString(riSection, riKey, Chr(1), sRiBuffer, 255, INIFile)
        If Left$(sRiBuffer, 1) <> Chr(1) Then
            sRiValue = Left$(sRiBuffer, sRiLong)
            If sRiValue <> "" Then
                ReadINI = sRiValue
            Else
                ReadINI = riDefault
            End If
        Else
            ReadINI = riDefault
        End If
    Else
        ReadINI = riDefault
    End If
End Function


