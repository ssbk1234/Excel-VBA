VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm11 
   Caption         =   "Automator For Data Extraction"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9795
   OleObjectBlob   =   "AutomatorForDataExtraction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim nok As Integer
Dim dd As String
Dim WS As Worksheet


Private Sub CommandButton45_Click()

End Sub

Private Sub CommandButton55_Click()
Dim f As Integer
f = 100

If (TextBox2.Value = "" Or TextBox3.Value = "") Then
    MsgBox "Please fill the empty fields !"
    f = 0
End If
 If f = 100 Then
    Dim uname As String
    Dim pass As String
    uname = TextBox2.Value
    pass = TextBox3.Value
   
    
    Dim sdate As Date
    Dim edate As Date
    sdate = TextBox4.Value
    edate = TextBox5.Value
  '  MsgBox sdate & " " & edate
    Call GetData(uname, pass, sdate, edate)
    
End If
End Sub

Private Sub CommandButton60_Click()

End Sub

Private Sub CommandButton56_Click()
Call UserForm_Initialize
End Sub

Private Sub CommandButton57_Click()
Unload Me
End Sub

Private Sub Frame1_Click()
End Sub





Private Sub Label2_Click()

End Sub

'~~> Prepare the Calendar on Userform
Private Sub UserForm_Initialize()
    '~~> Create a temp sheet for `GenerateCal` to work upon
    Set WS = Sheets.Add
    WS.Visible = xlSheetVeryHidden
    GenerateCal Format(Date, "mm/yyyy")
    TextBox2.Value = ""
     TextBox3.Value = ""
   
End Sub

'~~> Next Month
Private Sub CommandButton43_Click()
    Dim dat As Date
    dat = DateAdd("m", -1, DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), 1))
    GenerateCal Format(dat, "mm/yyyy")
    CommandButton45.Caption = Format(dat, "mmm - yyyy")
    
End Sub

'~~> Previous Month
Private Sub CommandButton44_Click()

    Dim dat As Date
    dat = DateAdd("m", 1, DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), Val(Format(CommandButton45.Caption, "MM")), 1))
    GenerateCal Format(dat, "mm/yyyy")
    CommandButton45.Caption = Format(dat, "mmm - yyyy")
    
End Sub

'~~> Ok Button
Private Sub CommandButton53_Click()
     Dim dt As Date
    dt = TextBox1.Text
   Dim dtt As String
   dtt = Format(dt, "dd-mm-yyyy")
  ' MsgBox dtt
   nok = nok + 1
   If (nok = 1) Then
     TextBox4.Value = dtt
   End If
   If (nok = 2) Then
     TextBox5.Value = dtt
   End If
  ' nok = 0

End Sub

'~~> Cancel Button
Private Sub CommandButton54_Click()
    Unload Me
End Sub

'~~> Generate Sheet
'~~> Code based on http://support.microsoft.com/kb/150774
Private Sub GenerateCal(dt As String)
    With WS
        .Cells.Clear
        StartDay = DateValue(dt)
        ' Check if valid date but not the first of the month
        ' -- if so, reset StartDay to first day of month.
        If Day(StartDay) <> 1 Then
            StartDay = DateValue(Month(StartDay) & "/1/" & _
                Year(StartDay))
        End If
        ' Prepare cell for Month and Year as fully spelled out.
        .Range("a1").NumberFormat = "mmmm yyyy"
        ' Center the Month and Year label across a1:g1 with appropriate
        ' size, height and bolding.
        With .Range("a1:g1")
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlCenter
            .Font.Size = 18
            .Font.Bold = True
            .RowHeight = 35
        End With
        ' Prepare a2:g2 for day of week labels with centering, size,
        ' height and bolding.
        With .Range("a2:g2")
            .ColumnWidth = 11
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = xlHorizontal
            .Font.Size = 12
            .Font.Bold = True
            .RowHeight = 20
        End With
        ' Put days of week in a2:g2.
        .Range("a2") = "Sunday"
        .Range("b2") = "Monday"
        .Range("c2") = "Tuesday"
        .Range("d2") = "Wednesday"
        .Range("e2") = "Thursday"
        .Range("f2") = "Friday"
        .Range("g2") = "Saturday"
        ' Prepare a3:g7 for dates with left/top alignment, size, height
        ' and bolding.
        With .Range("a3:g8")
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlTop
            .Font.Size = 18
            .Font.Bold = True
            .RowHeight = 21
        End With
        ' Put inputted month and year fully spelling out into "a1".
        .Range("a1").Value = Application.Text(dt, "mmmm yyyy")
        ' Set variable and get which day of the week the month starts.
        DayofWeek = Weekday(StartDay)
        ' Set variables to identify the year and month as separate
        ' variables.
        CurYear = Year(StartDay)
        CurMonth = Month(StartDay)
        ' Set variable and calculate the first day of the next month.
        FinalDay = DateSerial(CurYear, CurMonth + 1, 1)
        ' Place a "1" in cell position of the first day of the chosen
        ' month based on DayofWeek.
        Select Case DayofWeek
            Case 1
                .Range("a3").Value = 1
            Case 2
                .Range("b3").Value = 1
            Case 3
                .Range("c3").Value = 1
            Case 4
                .Range("d3").Value = 1
            Case 5
                .Range("e3").Value = 1
            Case 6
                .Range("f3").Value = 1
            Case 7
                .Range("g3").Value = 1
        End Select
        ' Loop through .Range a3:g8 incrementing each cell after the "1"
        ' cell.
        For Each cell In .Range("a3:g8")
            RowCell = cell.Row
            ColCell = cell.Column
            ' Do if "1" is in first column.
            If cell.Column = 1 And cell.Row = 3 Then
            ' Do if current cell is not in 1st column.
            ElseIf cell.Column <> 1 Then
                If cell.Offset(0, -1).Value >= 1 Then
                    cell.Value = cell.Offset(0, -1).Value + 1
                    ' Stop when the last day of the month has been
                    ' entered.
                    If cell.Value > (FinalDay - StartDay) Then
                        cell.Value = ""
                        ' Exit loop when calendar has correct number of
                        ' days shown.
                        Exit For
                    End If
                End If
            ' Do only if current cell is not in Row 3 and is in Column 1.
            ElseIf cell.Row > 3 And cell.Column = 1 Then
                cell.Value = cell.Offset(-1, 6).Value + 1
                ' Stop when the last day of the month has been entered.
                If cell.Value > (FinalDay - StartDay) Then
                    cell.Value = ""
                    ' Exit loop when calendar has correct number of days
                    ' shown.
                    Exit For
                End If
            End If
        Next
    
        ' Create Entry cells, format them centered, wrap text, and border
        ' around days.
        For x = 0 To 5
            .Range("A4").Offset(x * 2, 0).EntireRow.Insert
            With .Range("A4:G4").Offset(x * 2, 0)
                .RowHeight = 65
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlTop
                .WrapText = True
                .Font.Size = 10
                .Font.Bold = False
                ' Unlock these cells to be able to enter text later after
                ' sheet is protected.
                .Locked = False
            End With
            ' Put border around the block of dates.
            With .Range("A3").Offset(x * 2, 0).Resize(2, _
            7).Borders(xlLeft)
                .Weight = xlThick
                .ColorIndex = xlAutomatic
            End With
    
            With .Range("A3").Offset(x * 2, 0).Resize(2, _
            7).Borders(xlRight)
                .Weight = xlThick
                .ColorIndex = xlAutomatic
            End With
            .Range("A3").Offset(x * 2, 0).Resize(2, 7).BorderAround _
               Weight:=xlThick, ColorIndex:=xlAutomatic
        Next
        If .Range("A13").Value = "" Then .Range("A13").Offset(0, 0) _
           .Resize(2, 8).EntireRow.Delete
    
        ' Resize window to show all of calendar (may have to be adjusted
        ' Allow screen to redraw with calendar showing.
        Application.ScreenUpdating = True
        
        '~~> Update Dates on command button
        CommandButton1.Caption = .Range("A3").Text
        CommandButton2.Caption = .Range("B3").Text
        CommandButton3.Caption = .Range("C3").Text
        CommandButton4.Caption = .Range("D3").Text
        CommandButton5.Caption = .Range("E3").Text
        CommandButton6.Caption = .Range("F3").Text
        CommandButton7.Caption = .Range("G3").Text
        
        CommandButton8.Caption = .Range("A5").Text
        CommandButton9.Caption = .Range("B5").Text
        CommandButton10.Caption = .Range("C5").Text
        CommandButton11.Caption = .Range("D5").Text
        CommandButton12.Caption = .Range("E5").Text
        CommandButton13.Caption = .Range("F5").Text
        CommandButton14.Caption = .Range("G5").Text
        
        CommandButton15.Caption = .Range("A7").Text
        CommandButton16.Caption = .Range("B7").Text
        CommandButton17.Caption = .Range("C7").Text
        CommandButton18.Caption = .Range("D7").Text
        CommandButton19.Caption = .Range("E7").Text
        CommandButton20.Caption = .Range("F7").Text
        CommandButton21.Caption = .Range("G7").Text
        
        CommandButton22.Caption = .Range("A9").Text
        CommandButton23.Caption = .Range("B9").Text
        CommandButton24.Caption = .Range("C9").Text
        CommandButton25.Caption = .Range("D9").Text
        CommandButton26.Caption = .Range("E9").Text
        CommandButton27.Caption = .Range("F9").Text
        CommandButton28.Caption = .Range("G9").Text
        
        CommandButton29.Caption = .Range("A11").Text
        CommandButton30.Caption = .Range("B11").Text
        CommandButton31.Caption = .Range("C11").Text
        CommandButton32.Caption = .Range("D11").Text
        CommandButton33.Caption = .Range("E11").Text
        CommandButton34.Caption = .Range("F11").Text
        CommandButton35.Caption = .Range("G11").Text
        
        CommandButton46.Caption = .Range("A13").Text
        CommandButton47.Caption = .Range("B13").Text
        CommandButton48.Caption = .Range("C13").Text
        CommandButton49.Caption = .Range("D13").Text
        CommandButton50.Caption = .Range("E13").Text
        CommandButton51.Caption = .Range("F13").Text
        CommandButton52.Caption = .Range("G13").Text
    End With
End Sub

'~~> Delete the Temp Sheet that was created
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    Application.DisplayAlerts = False
    WS.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

'~~> This section simply updates the date in the text box when a button is pressed
Private Sub CommandButton1_Click()
    If CommandButton1.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton1.Caption)
End Sub
Private Sub CommandButton2_Click()
    If CommandButton2.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton2.Caption)
End Sub
Private Sub CommandButton3_Click()
    If CommandButton3.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton3.Caption)
End Sub
Private Sub CommandButton4_Click()
    If CommandButton4.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton4.Caption)
End Sub
Private Sub CommandButton5_Click()
    If CommandButton5.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton5.Caption)
End Sub
Private Sub CommandButton6_Click()
    If CommandButton6.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton6.Caption)
End Sub
Private Sub CommandButton7_Click()
    If CommandButton7.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton7.Caption)
End Sub
Private Sub CommandButton8_Click()
    If CommandButton8.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton8.Caption)
End Sub
Private Sub CommandButton9_Click()
    If CommandButton9.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton9.Caption)
End Sub
Private Sub CommandButton10_Click()
    If CommandButton10.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton10.Caption)
End Sub
Private Sub CommandButton11_Click()
    If CommandButton11.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton11.Caption)
End Sub
Private Sub CommandButton12_Click()
    If CommandButton12.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton12.Caption)
End Sub
Private Sub CommandButton13_Click()
    If CommandButton13.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton13.Caption)
End Sub
Private Sub CommandButton14_Click()
    If CommandButton14.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton14.Caption)
End Sub
Private Sub CommandButton15_Click()
    If CommandButton15.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton15.Caption)
End Sub
Private Sub CommandButton16_Click()
    If CommandButton16.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton16.Caption)
End Sub
Private Sub CommandButton17_Click()
    If CommandButton17.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton17.Caption)
End Sub
Private Sub CommandButton18_Click()
    If CommandButton18.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton18.Caption)
End Sub
Private Sub CommandButton19_Click()
    If CommandButton19.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton19.Caption)
End Sub
Private Sub CommandButton20_Click()
    If CommandButton20.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton20.Caption)
End Sub
Private Sub CommandButton21_Click()
    If CommandButton21.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton21.Caption)
End Sub
Private Sub CommandButton22_Click()
    If CommandButton22.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton22.Caption)
End Sub
Private Sub CommandButton23_Click()
    If CommandButton23.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton23.Caption)
End Sub
Private Sub CommandButton24_Click()
    If CommandButton24.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton24.Caption)
End Sub
Private Sub CommandButton25_Click()
    If CommandButton25.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton25.Caption)
End Sub
Private Sub CommandButton26_Click()
    If CommandButton26.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton26.Caption)
End Sub
Private Sub CommandButton27_Click()
    If CommandButton27.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton27.Caption)
End Sub
Private Sub CommandButton28_Click()
    If CommandButton28.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton28.Caption)
End Sub
Private Sub CommandButton29_Click()
    If CommandButton29.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton29.Caption)
End Sub
Private Sub CommandButton30_Click()
    If CommandButton30.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton30.Caption)
End Sub
Private Sub CommandButton31_Click()
    If CommandButton31.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton31.Caption)
End Sub
Private Sub CommandButton32_Click()
    If CommandButton32.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton32.Caption)
End Sub

Private Sub CommandButton33_Click()
    If CommandButton33.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton33.Caption)
End Sub

Private Sub CommandButton34_Click()
    If CommandButton34.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton34.Caption)
End Sub
Private Sub CommandButton35_Click()
    If CommandButton35.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton35.Caption)
End Sub
Private Sub CommandButton46_Click()
    If CommandButton46.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton46.Caption)
End Sub

Private Sub CommandButton47_Click()
    If CommandButton47.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton47.Caption)
End Sub

Private Sub CommandButton48_Click()
    If CommandButton48.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton48.Caption)
End Sub
Private Sub CommandButton49_Click()
    If CommandButton49.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton49.Caption)
End Sub
Private Sub CommandButton50_Click()
    If CommandButton50.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton50.Caption)
End Sub
Private Sub CommandButton51_Click()
    If CommandButton51.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton51.Caption)
End Sub

Private Sub CommandButton52_Click()
    If CommandButton52.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton52.Caption)
End Sub

Sub GetData(u As String, p As String, sd As Date, ed As Date)
    Dim oIE As Object
   ' MsgBox sd & " " & ed
    Set oIE = CreateObject("InternetExplorer.Application")
    Dim url_string As String
    url_string = "http://www.facebook.com/login"
    Application.ScreenUpdating = False
    'url_string = "http://abhaykptiitk.comoj.com/"
    oIE.Visible = True
    oIE.navigate url_string
        Do While oIE.Busy
            Application.Wait DateAdd("s", 1, Now)
        Loop
    Dim aTag As Object
    Set aTag = oIE.document.getelementsByTagName("a")
    
    Dim it As Integer
    it = 0
    Dim fk As Boolean
    fk = 0
    Do Until it > aTag.Length - 2
           '  MsgBox aTag(it).InnerText
             it = it + 1
             If InStr(aTag(it).innerText, "Logout") = 1 Then
               fk = 1
               Exit Do
         End If
    Loop
    

    If fk = 0 Then
        oIE.document.All.Item("email").Value = u
        oIE.document.All.Item("pass").Value = p
        oIE.document.All.Item("login").Click
    End If
        Do While oIE.Busy
            Application.Wait DateAdd("s", 1, Now)
        Loop
    Dim oDoc As Object, oElem As Object
    Set oDoc = oIE.document
    Set pDiv = oIE.document.getelementsByTagName("div")
    Dim pp As Integer
    pp = -1
  
    Count = 0
    Do Until pp > pDiv.Length - 2
        pp = pp + 1
        If (pDiv(pp).getAttribute("class") = "linkWrap noCount") Then
        Count = Count + 1
         
            If (InStr(pDiv(pp).innerText, "age") > 0) Then
                pDiv(pp).Click
                 Application.Wait Now + TimeValue("00:00:02")
                
                 Exit Do
            End If
        End If
    Loop
  
    Set aiTg = oIE.document.getelementsByTagName("a")
    Dim ppart As String
    ppart = InputBox(Prompt:="please enter a part of your page name according to view insight url", _
          Title:="Part page name")
    itt = -1
       Do Until itt > aiTg.Length - 2
             itt = itt + 1
             If aiTg(itt).innerText = "View Insights" And InStr(aiTg(itt), ppart) > 0 Then
                aiTg(itt).Click
                Application.Wait Now + TimeValue("00:00:02")
                Exit Do
             End If
    Loop
    
  '  Dim ie As Object
'    Set ie = CreateObject("InternetExplorer.Application")
   ' With ie
    
    '    .navigate my_url
     '   .Visible = True
       
     '   Do While ie.Busy
    '        Application.Wait DateAdd("s", 1, Now)
    '    Loop
        
    '    Set objHTML = .document
    '    DoEvents
  '  End With
   '  Set e1Div = objHTML.getelementsByTagName("div")
   ' Set a1Tg = objHTML.getelementsByTagName("a")
    
  
  '  MsgBox a1Tg.Length
    
  '  MsgBox "Number of div class elements is " & count
    'Annual Data tab:
   ' Set oElem = GetElementsByClassNameAndInnerText(oDoc, "adsManagerReportsButton", True, "View Report", False)
  '  oElem.Click
    'this works
  
 ' Dim purl As String
 ' purl = "https://www.facebook.com/" & pg & "?sk=insights"
       
   '     oIE.navigate purl
   '     oIE.Visible = True
     '   Do While oIE.Busy
   '         Application.Wait DateAdd("s", 1, Now)
  '      Loop
    
   ' nurl = "https://www.facebook.com/" & pg & "?sk=insights&section=navPosts"
    
  '  oIE.navigate nurl
  '      oIE.Visible = False
  '      Do While oIE.Busy
    '        Application.Wait DateAdd("s", 1, Now)
   '     Loop
  '
    'Set objShell = CreateObject("Shell.Application")
 '   IE_count = objShell.Windows.count
 ''   For x = 0 To (IE_count - 1)
  '  '    On Error Resume Next
     '   my_url = objShell.Windows(x).document.Location
  '  '    my_title = objShell.Windows(x).document.Title
        
    '    If my_url = nurl Then
    
   '         Set iee = objShell.Windows(x)
         '   MsgBox "my title is" & my_title
   '         Exit For
  ''      End If
  '  Next

    
    Set aiiTg = oIE.document.getelementsByTagName("a")
    MsgBox "Clicking on See All Post"
    iitt = -1
       Do Until iitt > aiiTg.Length - 2
           
             iitt = iitt + 1
            
             If aiiTg(iitt).innerText = "See All Posts" Then
             
                aiiTg(iitt).Click
                Application.Wait Now + TimeValue("00:00:03")
                Exit Do
             End If
    Loop
    
    
    
    Set objShell = CreateObject("Shell.Application")
    IE_count = objShell.Windows.Count
    
    For x = 0 To (IE_count - 1)
        On Error Resume Next
        my_url = objShell.Windows(x).document.Location
        my_title = objShell.Windows(x).document.Title
    
        If InStr(my_url, "navPosts") > 0 Then
    
           Set iee = objShell.Windows(x)
      
            Exit For
           End If
    Next

    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    With ie
        .navigate my_url
        .Visible = True
        Do While ie.Busy
            Application.Wait DateAdd("s", 1, Now)
        Loop
         Application.Wait Now + TimeValue("00:00:02")
        Set objHTML = .document
        DoEvents
    End With
  
    Set eSpan = objHTML.getelementsByTagName("span")
    Set eDiv = objHTML.getelementsByTagName("div")
    Set aTg = objHTML.getelementsByTagName("a")
  
   
Dim nc As Integer
Dim t As Integer
t = -1
Dim i As Integer
Dim j As Integer
i = 1
nc = 0
Dim c As Integer
c = 1
Dim strName As String
strName = InputBox(Prompt:="Enter the name of the sheet !", _
          Title:="Sheet name")
          
ActiveSheet.Name = strName

MsgBox ActiveSheet.Name & " Sheet is being cleared!"
ActiveSheet.Cells.Clear

Do Until t > eSpan.Length - 2
    t = t + 1
    If InStr(eSpan(t).getAttribute("class"), "_5k43") = 1 Then
    
            If (eSpan(t).innerText = "Published") Then
                Cells(1, c) = "Date"
                c = c + 1
                Cells(1, c) = "Time"
                c = c + 1
                
            ElseIf (eSpan(t).innerText = "Post") Then
                 Cells(1, c) = "Post"
                 c = c + 1
                
            ElseIf (eSpan(t).innerText = "Reach") Then
                 Cells(1, c) = "Reach"
                 c = c + 1
                 Cells(1, c) = "Post Clicks"
                 c = c + 1
                 Cells(1, c) = "Likes , Comments & Share"
                 c = c + 1
            End If
            
    End If
Loop

 t = -1
'Dim y As Integer
Dim nr As Integer
nr = 1
y = -1
Dim w As String
Dim nrr As Integer
nrr = 1
Dim r As Integer
r = 1
Dim fmDate As String
Dim ddtt As String
Dim tm As String



    Dim yy As Integer
       yy = -1
      Count = 0
      Dim l As Integer
      l = eDiv.Length
   
        Do Until yy > l - 2
            yy = yy + 1
        
            If (eDiv(yy).innerText = "See More") Then
          
                eDiv(yy).Click
               Application.Wait Now + TimeValue("00:00:02")
               
            End If
           l = eDiv.Length
        Loop
    Dim f As Integer
    Application.Wait Now + TimeValue("00:00:02")
    l = eDiv.Length
    
      Dim de As String
      Dim dc As Date
      Dim md As String
      Dim st As String
      Dim dsr As String
      Dim lt As String
      Dim cdt As Date
      Dim dttt As String
      Dim fn As String
      Dim strR As String
      
      Dim cellstr As String
      
    Do Until y > eDiv.Length - 2
       y = y + 1
       
                If (Cells(1, nr) = "Date") And (Right(eDiv(y).getAttribute("data-reactid"), 9) = "$postTime") Then
                 '  MsgBox eDiv(y).InnerText
                    de = Left(eDiv(y).innerText, 10)
                    st = Left(de, 2)
                    md = Mid(de, 4, 2)
                    lt = Right(de, 4)
                    
                    dsr = md & "-" & st & "-" & lt
                    
                    cdt = dsr
                        
                If (cdt >= sd And cdt <= ed) Then
    
                        Cells(nrr + 1, nr) = cdt
                        nr = nr + 1
                End If
             
            ElseIf (Cells(1, nr) = "Time") And (InStr(eDiv(y).getAttribute("data-reactid"), ".$postTime.1") > 0) Then
                 
                    Cells(nrr + 1, nr) = eDiv(y).innerText
                    nr = nr + 1
             
            ElseIf (Cells(1, nr) = "Post") And (InStr(eDiv(y).getAttribute("data-reactid"), "$postContent.0.$right") > 0) Then
                    Cells(nrr + 1, nr) = eDiv(y).innerText
                    nr = nr + 1
                  
            ElseIf (Cells(1, nr) = "Reach") And (InStr(eDiv(y).getAttribute("class"), "_5kn3 ellipsis") > 0) Then
                    Cells(nrr + 1, nr) = eDiv(y).innerText
                    nr = nr + 1
            
            ElseIf (Cells(1, nr) = "Post Clicks") And (InStr(eDiv(y).getAttribute("data-reactid"), "$engagement.0.$clicks.0") > 0) Then
                    Cells(nrr + 1, nr) = eDiv(y).innerText
                    nr = nr + 1
                  
            ElseIf (Cells(1, nr) = "Likes , Comments & Share") And (InStr(eDiv(y).getAttribute("data-reactid"), "$engagement.0.$lcs.0") > 0) Then
                    Cells(nrr + 1, nr) = eDiv(y).innerText
                    nr = nr + 1
                End If
    
            If nr = 7 Then
                    nrr = nrr + 1
                    nr = 1
            End If
                
Loop
       it = -1
       Do Until it > aTg.Length - 2
           '  MsgBox aTag(it).InnerText
             it = it + 1
             If InStr(aTg(it).innerText, "Logout") = 1 Then
                aTg(it).Click
               Exit Do
         End If
    Loop
    
Application.ScreenUpdating = True
ie.Quit
oIE.Quit
oIE = Nothing
ie = Nothing
Unload Me
End Sub

