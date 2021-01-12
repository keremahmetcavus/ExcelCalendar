VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} kcls_calendar 
   Caption         =   "Tarih Seçim Ekraný"
   ClientHeight    =   1455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4545
   OleObjectBlob   =   "kcls_calendar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "kcls_calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' HOW TO USE ////GUÝDE BOOK//

'' /////////////////////////////////////
'' myvariable = kcls_calendar.Get_Date()
'' /////////////////////////////////////
'' Write this code to get date var

'' /////////////////////////////////////
'' myvariable = kcls_calendar.Get_Date(True)
'' /////////////////////////////////////
'' Write this code for just month year type

'' /////////////////////////////////////
'' mydayvariable = day(myvariable)
'' mymonthvariable = month(myvariable)
'' myyearvariable = year(myvariable)
'' /////////////////////////////////////
'' Simple exported Values TO DAY,MONTH,YEAR
'' EXAMPLER [01] [01] [2021]

'' /////////////////////////////////////
'' mydayvariable = kcls_calendar.date_fixer(day(myvariable))
'' mymonthvariable = kcls_calendar.date_fixer(month(myvariable))
'' myyearvariable = year(myvariable)
'' /////////////////////////////////////
'' Fixed exported Values TO DAY,MONTH,YEAR
'' EXAMPLER [01] [01] [2021]


'' HOW TO USE ////GUÝDE BOOK//

'' AFTER THÝS LÝNE CODES GOES UNDER..





''____________________________________________
''Global Control For Title Bar////////////////
''--------------------------------------------
Option Explicit
Public justmonth As Boolean
Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const SC_CLOSE = &HF060
#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
        ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
        ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar _
        Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function DeleteMenu _
        Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function GetSystemMenu _
        Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
#End If
Public Sub SystemButtonSettings(frm As Object, show As Boolean)
    Dim windowStyle As Long
    Dim windowHandle As Long
    windowHandle = FindWindowA(vbNullString, frm.Caption)
    windowStyle = GetWindowLong(windowHandle, GWL_STYLE)
    If show = False Then
        SetWindowLong windowHandle, GWL_STYLE, (windowStyle And Not WS_SYSMENU)
    Else
        SetWindowLong windowHandle, GWL_STYLE, (windowStyle + WS_SYSMENU)
    End If
    DrawMenuBar (windowHandle)
End Sub
''____________________________________________
''Global Control For Title Bar////////////////
''--------------------------------------------

''____________________________________________
''Control subs ///////////////////////////////
''--------------------------------------------
Private Sub mouse_control(a)
    Select Case a
    Case "event_move"
        Button1.BackColor = RGB(217, 83, 79)
        Label4.ForeColor = RGB(256, 256, 256)
    Case "event_down"
        Button1.Move Button1.Left + 1, Button1.Top + 1
        Label4.Move Label4.Left + 1, Label4.Top + 1
    Case "event_up"
        Button1.Move 18, 30
        Label4.Move 60, 42
    End Select
End Sub
Public Function date_fixer(Optional a As Variant)
 If Len(a) = 1 Then a = "0" & a
 date_fixer = a
End Function
''____________________________________________
''Control subs ///////////////////////////////
''--------------------------------------------
Public Sub Button1_Click()
Me.Hide
End Sub

''____________________________________________
''Makeup Events ///////////////////////////////
''--------------------------------------------
Private Sub Button1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call mouse_control("event_down")
End Sub
Private Sub Button1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call mouse_control("event_up")
End Sub
Private Sub Button1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call mouse_control("event_move")
End Sub
Private Sub Label4_Click()
Call Button1_Click
End Sub
Private Sub Label4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call mouse_control("event_down")
End Sub
Private Sub Label4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call mouse_control("event_up")
End Sub
Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call mouse_control("event_move")
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Button1.BackColor = RGB(256, 256, 256)
Label4.ForeColor = &H80000012
End Sub
''____________________________________________
''Makeup Events///////////////////////////////
''--------------------------------------------

''____________________________________________
''Label Control///////////////////////////////
''--------------------------------------------
Private Sub ComboBox1_Change()
 Label1.Caption = ComboBox1.Value
End Sub
Private Sub ComboBox2_Change()
 Label2.Caption = ComboBox2.Value
End Sub
Private Sub ComboBox3_Change()
 Label3.Caption = ComboBox3.Value
End Sub
Private Sub Label1_Click()
ComboBox1.DropDown
End Sub
Private Sub Label2_Click()
ComboBox2.DropDown
End Sub
Private Sub Label3_Click()
ComboBox3.DropDown
End Sub
''____________________________________________
''Label Control///////////////////////////////
''--------------------------------------------


''____________________________________________
''Userform Declare code///////////////////////
''--------------------------------------------
Private Sub UserFormInitialize()
Dim a As String
Dim j As Integer
Dim i As Integer
Call SystemButtonSettings(Me, False)
If justmonth = True Then Call justmonth_type

 Label1.Font.Size = 20
 Label2.Font.Size = 20
 Label3.Font.Size = 20
 Label4.Font.Bold = True
 Label4.Font.Size = 10
 Button1.Font.Size = 10
 
 ComboBox1.Font.Weight = 500
 ComboBox1.Font.Size = 15
 ComboBox2.Font.Size = 15
 ComboBox3.Font.Size = 15

'' Add Date var
For i = 1 To 31
    a = i
    Call date_fixer(a)
    ComboBox1.AddItem a
Next i
'' Add Month var
ComboBox2.AddItem "OCAK"
ComboBox2.AddItem "ÞUBAT"
ComboBox2.AddItem "MART"
ComboBox2.AddItem "NÝSAN"
ComboBox2.AddItem "MAYIS"
ComboBox2.AddItem "HAZÝRAN"
ComboBox2.AddItem "TEMMUZ"
ComboBox2.AddItem "AÐUSTOS"
ComboBox2.AddItem "EYLÜL"
ComboBox2.AddItem "EKÝM"
ComboBox2.AddItem "KASIM"
ComboBox2.AddItem "ARALIK"
'' Add Year Var
For i = 1 To 30
    j = 2015 + i
    ComboBox3.AddItem j
Next i
'
a = Day(Date)
Call date_fixer(Day(Date))
Label1.Caption = a
ComboBox1.Value = a
Label3.Caption = Year(Date)
ComboBox3.Value = Year(Date)
''Casing date
Call Convert_Month_To_Text(Month(Date))
End Sub
''____________________________________________
''Userform Declare code///////////////////////
''--------------------------------------------

Public Function Get_Date(Optional justmonth As Boolean = False)
startoffunction:
    Me.justmonth = justmonth
    Call UserFormInitialize
    Dim a As Variant
    Me.show 1
    a = Convert_Label_To_Date()
    If a = "err" Then
    MsgBox "Lütfen Geçerli Bir Tarih Giriniz."
    GoTo startoffunction
    Else
    Get_Date = a
    End If
    Unload Me
End Function
Private Function Convert_Label_To_Date()
Dim a As Variant
Dim b As Variant
Dim c As Variant
If Me.justmonth = True Then
    a = 1
    b = Convert_Text_To_Month(Label2.Caption)
    c = Label3.Caption
    a = a & "." & b & "." & c
    Convert_Label_To_Date = CDate(a)
Else
    a = Label1.Caption
    b = Convert_Text_To_Month(Label2.Caption)
    c = Label3.Caption
    a = a & "." & b & "." & c
    On Error GoTo errhand
    Convert_Label_To_Date = CDate(a)
End If
Exit Function
errhand:
Convert_Label_To_Date = "err"
End Function
Private Function Convert_Text_To_Month(a As Variant)
Select Case a
    Case "OCAK"
     a = "01"
    Case "ÞUBAT"
     a = "02"
    Case "MART"
    a = "03"
    Case "NÝSAN"
    a = "04"
    Case "MAYIS"
    a = "05"
    Case "HAZÝRAN"
    a = "06"
    Case "TEMMUZ"
    a = "07"
    Case "AÐUSTOS"
    a = "08"
    Case "EYLÜL"
    a = "09"
    Case "EKÝM"
    a = "10"
    Case "KASIM"
    a = "11"
    Case "ARALIK"
    a = "12"
End Select
Convert_Text_To_Month = a
End Function
Private Function Convert_Month_To_Text(a As Variant)
Select Case a
Case 1
Label2.Caption = "OCAK"
ComboBox2.Value = "OCAK"
Case 2
Label2.Caption = "ÞUBAT"
ComboBox2.Value = "ÞUBAT"
Case 3
Label2.Caption = "MART"
ComboBox2.Value = "MART"
Case 4
Label2.Caption = "NÝSAN"
ComboBox2.Value = "NÝSAN"
Case 5
Label2.Caption = "MAYIS"
ComboBox2.Value = "MAYIS"
Case 6
Label2.Caption = "HAZÝRAN"
ComboBox2.Value = "HAZÝRAN"
Case 7
Label2.Caption = "TEMMUZ"
ComboBox2.Value = "TEMMUZ"
Case 8
Label2.Caption = "AÐUSTOS"
ComboBox2.Value = "AÐUSTOS"
Case 9
Label2.Caption = "EYLÜL"
ComboBox2.Value = "EYLÜL"
Case 10
Label2.Caption = "EKÝM"
ComboBox2.Value = "EKÝM"
Case 11
Label2.Caption = "KASIM"
ComboBox2.Value = "KASIM"
Case 12
Label2.Caption = "ARALIK"
ComboBox2.Value = "ARALIK"
End Select
End Function

Private Sub justmonth_type()
Label1.Enabled = False
ComboBox1.Enabled = False
Label1.Visible = False
ComboBox1.Visible = False
Label2.Left = Label2.Left - 30
ComboBox2.Left = ComboBox2.Left - 30
Label3.Left = Label3.Left - 30
ComboBox3.Left = ComboBox3.Left - 30
End Sub
