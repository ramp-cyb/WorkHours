Dim AppStartedFlag As Boolean
Dim RefreshTimerFlag As Boolean
'''''''''''''''''''''''''
Dim currWS As Worksheet
'''''''''''''''''''''''''
Public empID As String
Public misPath As String
Public reportPage As String
'''''''''''''''''''''''''Gates
Dim enterMainGate() As String
Dim enterPlayGate() As String
Dim enterWorkGate() As String
''''''''''''''''
Dim swipeRowNum As Integer
Dim swipeRowNum1 As Integer
Dim swipeColNum_yesterday As Integer
Dim swipeColNum_today As Integer
'''''''''''''''''''''''''
Dim shellWins As New ShellWindows
Dim IE As Object ''should be  --- As InternetExplorer
Dim objHTMLDocument As HTMLDocument
Dim totRow As Integer

Sub test111()
    'refreshToStartApp
    
    If (getUserInputForEmpId = "") Then
        dummyReply = getUserInputForEmpId
    End If
    MsgBox dummyReply
    
End Sub







 Public Sub StartCat_Click()
 Dim r As Integer
 Dim nb_days As Integer
 r = 5
 
 nb_days = Day(DateSerial(Year(Date), Month(Date) + 1, 1) - 1)
 r = r + nb_days - 1
    refreshToStartApp
    Set objHTMLDocument = getDocument("iframe1")
    clear_ReportGrid
    getFortnightDates
    getEmployeeId
    '
    navigateTo reportPage
    'Attendance Log Report
    dummyReturn = followLink("Attendance Log Report", "text")
    'For Emp -> dropdown ID on "Actual Hour" page
    dummyReturn = selectDropdown("title", "EmployeeID", empID)
    'From Date
    IE.Document.getElementById("DMNDateDateRangeControl4392_FromDateCalender_DTB").Value = Format(Cells(5, 8).Value, "dd-mmm-yyyy")
    'To Date
    IE.Document.getElementById("DMNDateDateRangeControl4392_ToDateCalender_DTB").Value = Format(Cells(r, 8).Value, "dd-mmm-yyyy")
    'Generate report
    dummyReturn = followLink("ViewReportImageButton", "id")
    '
    r = Range("H5").Row
     For i = 1 To CInt(IE.Document.getElementById("ReportViewer1_ctl01_ctl01_ctl04").innerText)
        'MsgBox "kk" & i
        
        getDataToCells_tillDayBeforeYesterday (r) ' readSwipeLogTable colNum_pm
        If (i <> CInt(IE.Document.getElementById("ReportViewer1_ctl01_ctl01_ctl04").innerText)) Then
            dummyReturn = followLink("ReportViewer1$ctl01$ctl01$ctl05$ctl00$ctl00", "name")  'Next Page
        End If
        
        r = r + totRow
    Next
    
    
    ''''''''''''
    Yesterday_ClickFn
    ''''''''''''
    Today_ClickFn
   ' today_yesterday
    ''''''''''''
    success
    'klog ("str_pm")
End Sub

Private Function refreshToStartApp()
    gSetVars 'set all the global variables
    lSetVars 'set all the (local to sheet) variables
    'clears 'clear the NOT required data
    getIE misPath
End Function

'Sub tt()
'Set els = objHTMLDocument.getElementsByTagName("div")
'    For Each el In els
'        If (el.innerText = empID) Then
'
'            dateTemp = ""
'
'            klog ("=======================")
'            'MsgBox el.innerText
'            klog ("el.innerText=" & el.innerText)
'            empIdTemp = el.innerText
'            '
'            Set elTemp = el.ParentNode.NextSibling.FirstChild
'            'MsgBox elTemp.nodeName & ", " & elTemp.ParentNode.nodeName & ", " & elTemp.ParentNode.NextSibling.nodeName & ", " & elTemp.ParentNode.NextSibling.FirstChild.nodeName
'            klog ("elTemp.innerText=" & elTemp.innerText)
'            dateTemp = elTemp.innerText
'End Sub

Private Function lSetVars()
    AppStartedFlag = True
    '
    Set currWS = gGetActiveSheet
    '
    empID = Cells(gFindStr(currWS, "empID").Row, gFindStr(currWS, "empID").Column + 1).Value
    misPath = Cells(gFindStr(currWS, "misPath").Row, gFindStr(currWS, "misPath").Column + 1).Value
    reportPage = Cells(gFindStr(currWS, "reportPage").Row, gFindStr(currWS, "reportPage").Column + 1).Value
    '
    enterMainGate = Split(Cells(gFindStr(currWS, "enterMainGate").Row, gFindStr(currWS, "enterMainGate").Column + 1).Value, ",")
    enterPlayGate = Split(Cells(gFindStr(currWS, "enterPlayGate").Row, gFindStr(currWS, "enterPlayGate").Column + 1).Value, ",")
    enterWorkGate = Split(Cells(gFindStr(currWS, "enterWorkGate").Row, gFindStr(currWS, "enterWorkGate").Column + 1).Value, ",")
    '
    'swipeRowNum1 = gFindStr(currWS, ).Row + 1
    swipeRowNum = gFindStr(currWS, "Swipe - YesterDay").Row + 1
    swipeColNum_yesterday = gFindStr(currWS, "Swipe - YesterDay").Column
    swipeColNum_today = gFindStr(currWS, "Swipe - Today").Column
    '
    
    'MsgBox swipeColNum_today
End Function
Private Function setLogVars()
    'logRowNumFront = getValueOfSettingsVar("logRowNumFront") '35
    'logRowNumBU = getValueOfSettingsVar("logRowNumBU") '37
    'logColNum = getValueOfSettingsVar("logColNum") '1
    'klog "set Log variables"
End Function

Sub getIE__1(url_pm)
    Dim objShell As Object
    Dim marker As Integer: marker = 0
    Dim IE_count As Integer
    Dim my_url As String
    
    Set objShell = CreateObject("Shell.Application")
    IE_count = objShell.Windows.Count
    For x = 0 To (IE_count - 1)
        On Error Resume Next    ' sometimes more web pages are counted than are open
        my_url = objShell.Windows(x).Document.Location
        If my_url Like "http://cybagemis.cybage.com" & "*" Then 'compare to find if the desired web page is already open
           Set IE = objShell.Windows(x)
           marker = 1
           Exit For
        Else
            '
        End If
    Next
    If marker = 0 Then
        MsgBox ("Please open CybageMIS in IE")
    Else
        'MsgBox ("A matching webpage was found:" & IE.Document.Location)
    End If
    
ExitSub:
    Set shellWins = Nothing
End Sub

Private Function getIE(url_pm)
    Dim MIS_flagTemp As Boolean
    Dim IE_flagTemp As Boolean
    Dim shellWins As ShellWindows
    'Dim IE As InternetExplorer
    Dim i As Integer
    
    Set shellWins = New ShellWindows
    
    flagTemp = False
    If shellWins.Count > 0 Then
        'For i = 0 To shellWins.Count - 1
         '   If (shellWins.Item(i) = "Windows Internet Explorer") Then
           '     Set IE = shellWins.Item(i)
            'End If
        'Next
        'finding MIS URL
        For Each win In shellWins
            klog "________" & win
            If (win = "Windows Internet Explorer") Then
                Set IE = win
                klog "Found IE"
                IE_flagTemp = True
                If win.LocationURL <> vbNullString Then
                    If win.LocationURL = url_pm Then
                        Set IE = win
                        klog "Found IE with a blank tab"
                        MIS_flagTemp = True
                        klog "Found IE with MIS page."
                        Exit For
                    End If
                End If
            End If
        Next
    End If
    If (Not IE_flagTemp) Then
        createIE
        navigateTo url_pm
    Else
        If (Not MIS_flagTemp) Then
            navigateTo url_pm
        End If
    End If
    
ExitSub:
    Set shellWins = Nothing
End Function

Private Function createIE()
    Set IE = New InternetExplorerMedium
    IE.Visible = True
    klog "Created IE"
End Function

Function getIE__2(url_pm)
    Dim dummyRet As Boolean
    dummyRet = isIEOpen
    If (dummyRet) Then
        'do nothing
    Else
        openIE
    End If
    navigateTo "http://cybagemis.cybage.com/Framework/Iframe.aspx"
    'MsgBox IE.Document.Location
    'MsgBox "abcde=" & dummyRet
End Function

Function isIEOpen()
    Dim i, IECount As Integer
    Dim shellWins As ShellWindows
    Set shellWins = New ShellWindows
    MsgBox shellWins.Count
    MsgBox shellWins.Item(0)
    MsgBox shellWins.Item(1)
    MsgBox shellWins.Item(0).Visible
    If shellWins.Count = 0 Then
        isIEOpen = False
        Exit Function
    Else
        IECount = 0
        For i = 1 To shellWins.Count
            MsgBox shellWins.Item(i - 1)
            If shellWins.Item(i - 1) = "Windows Internet Explorer" Then
                Set IE = shellWins.Item(i - 1)
                isIEOpen = True
                Exit Function
                'Exit For
            Else
                isIEOpen = False
                Exit Function
            End If
        Next
    End If
End Function

Function openIE()
    Dim URL As String
    URL = "http://www.google.com"
    Dim ieApp As InternetExplorer
    'Dim prodID As Object
    Set ieApp = CreateObject("InternetExplorer.Application")
    With ieApp
        .Navigate URL
        .Visible = True
        Do While ieApp.Busy
            DoEvents
        Loop
        'MsgBox ieApp.Document.all.tags("span").Length
Label1:
On Error GoTo errorHandler:
        'Set prodID = .Document.getElementById("ProductID")
        'Range("A20").Value = prodID.Value
        '.Quit
    End With
    Set IE = ieApp
Exit Function
errorHandler:
    If Err.Number <> 462 Then
        GoTo Label1:
    End If
End Function

Private Function navigateTo(url_pm)
    IE.Navigate url_pm
    waitToLoad
    klog "Navigated to :" & url_pm
End Function

'Private Function loadURL(url_pm)
'    With IE
'        .Navigate url_pm
'        'While .Busy Or .readyState <> 4
'        While .Busy
'            klog "Loading..." & url_pm
'            DoEvents
'        Wend
'        'DoEvents
'    End With
'    'klog "with end..."
'End Function

Private Function followLink(str_pm, attr_pm)
    Dim i As Integer
    If (attr_pm = "id") Then
        Set objHTMLDocument = IE.Document.getElementById(str_pm)
    ElseIf (attr_pm = "name") Then
        Set objHTMLDocument = IE.Document.getElementsByName(str_pm)(0)
    ElseIf (attr_pm = "text") Then
        For i = 0 To IE.Document.getElementsByTagName("a").Length
            klog (str_pm & "---" & IE.Document.getElementsByTagName("a")(i).innerText)
            If IE.Document.getElementsByTagName("a")(i).innerText = str_pm Then
                Set objHTMLDocument = IE.Document.getElementsByTagName("a")(i)
                Exit For
            End If
        Next
    End If
    'klog objHTMLDocument
    'klog objHTMLDocument.innerText
    objHTMLDocument.Click
    waitToLoad
End Function

Function waitToLoad()
    Dim maxTry, try As Integer
    maxTry = 50
    try = 1
    Application.Wait (Now + TimeValue("0:00:1"))
    Do While (try < maxTry And IE.Busy)
        'WScript.sleep 100
        Application.Wait (Now + TimeValue("0:00:3"))
        try = try + 1
    Loop
    klog "waitToLoad()...IE.Busy=" & IE.Busy
End Function

Private Function getDocument(str_pm)
    Dim objHTMLDocumentTemp As HTMLDocument
    Dim objIframeTemp As Object
    Dim objCWDocumentTemp As HTMLDocument 'holding reference of ContentWindow of iFrame
    
    Set objHTMLDocumentTemp = IE.Document
    
    If (str_pm <> "") Then
        Set objIframeTemp = objHTMLDocumentTemp.getElementById(str_pm)
        
        Set objCWDocumentTemp = objIframeTemp.contentWindow.Document
        Set objHTMLDocumentTemp = objCWDocumentTemp
        'klog objCWDocumentTemp.getElementById("FirstHeadingLabel")
        'klog objCWDocumentTemp.getElementById("FirstHeadingLabel").innerText
        'objCWDocumentTemp.getElementById("FirstHeadingLabel").Style.display = "none" 'working fine
    End If
    Set getDocument = objHTMLDocumentTemp
    klog "stored reference of HTML Document"
End Function

Private Function selectDropdownById(dropdownID_pm, str_pm)
    klog "selectDropdownById()"
    Set objEl = IE.Document.getElementById(dropdownID_pm)
    Dim i As Integer
    i = 0
    For Each el In objEl.Children
        klog el.innerText
        If (InStr(el.innerText, str_pm)) Then
            klog "k-----------" & el.innerText
            objEl.selectedIndex = i
        End If
        i = i + 1
    Next
End Function


Private Function selectDropdown(attrName_pm, attrVal_pm, selectedVal_pm)
    'How To : selectDropdown("title", attrVal_pm, selectedVal_pm)
    klog "selectDropdown() " & attrName_pm & ", " & attrVal_pm & ", " & selectedVal_pm
    Dim objEl As IHTMLElement
    Dim strTemp As String
    Dim i, j As Integer
    i = 0
    For i = 0 To IE.Document.getElementsByTagName("select").Length
        If (attrName_pm = "title") Then
            strTemp = IE.Document.getElementsByTagName("select")(i).Title
            'MsgBox "attrVal_pm=" & attrVal_pm & ", strTemp=" & strTemp
            If (attrVal_pm = strTemp) Then
                Set objEl = IE.Document.getElementsByTagName("select")(i)
                Exit For
            End If
        End If
    Next
    If (i > IE.Document.getElementsByTagName("select").Length) Then
        'do nothing
    Else
        j = 0
        For Each el In objEl.Children
            klog el.innerText
            If (InStr(el.innerText, selectedVal_pm)) Then
                klog "k-----------" & el.innerText
                objEl.selectedIndex = j
            End If
            j = j + 1
        Next
    End If
End Function


Private Function getFortnightDates()
    Dim strTemp As String
    Dim date_test As Date
    Dim nb_days As Integer
    Dim first_Date As Date
    
    Dim r, c As Integer
    r = gFindStr(currWS, "WEEK 1").Row
    c = gFindStr(currWS, "WEEK 1").Column + 2
    'r = 5
    'c = 7
    'klog "kkkk:" & objHTMLDocument.getElementById("NoticeList").innerText
    'Set objEl = objHTMLDocument.getElementById("Notice")
    'klog objEl.getElementsByTagName("li")
    'strTemp = objEl.FirstChild.innerText
    'strTemp = Mid(strTemp, InStr(strTemp, ": ") + 2, 9)
    
    date_test = Date
    first_Date = DateSerial(Year(date_test), Month(date_test), 1)
    nb_days = Day(DateSerial(Year(date_test), Month(date_test) + 1, 1) - 1)
    
    Cells(r, c).Value = first_Date
    
    For i = r + 1 To r + nb_days - 1
        Cells(i, c).Value = Cells(i - 1, c).Value + 1
    Next i
End Function

Private Function getEmployeeId()
    On Error GoTo ERROR_HANDLE:
    '
    Dim r, c As Integer
    r = gFindStr(currWS, "EMP ID").Row
    c = gFindStr(currWS, "EMP ID").Column + 1
    'CybageMIS Profile
    dummyReturn = followLink("CybageMIS Profile", "text")
    Dim objHTMLDocTemp As HTMLDocument
    Set objHTMLDocTemp = getDocument("CYB_Employeeinfo")
    Dim objEl As IHTMLElement
    Set objEl = objHTMLDocTemp.getElementById("EmpDetailDatagrid") 'Getting TABLE
    Set objEl = objEl.Children(0) 'getting BODY
    Set objEl = objEl.Children(1) 'getting TR
    Set objEl = objEl.Children(0) 'getting TD
    'klog "Employee ID found on MIS=" & objEl.innerText 'It contains Employee ID
    Cells(r, c).Value = objEl.innerText
    '
    Exit Function
ERROR_HANDLE:
    MsgBox "error...getEmployeeId...Could not get employee ID from MIS"
    dummyReply = InputBox(prompt:="Employee ID please?", Title:="Important")
    Cells(r, c).Value = dummyReply
End Function

Public Function getDataToCells_tillDayBeforeYesterday(ByVal r As Integer)
    Dim t As IHTMLElement
    'Dim els As IHTMLElement
    Dim c As Integer
    Dim st As String
    'r = swipeRowNum1  '5
    c = 8
    Set objHTMLDocument = IE.Document
    Set els = objHTMLDocument.getElementsByTagName("td")
    'dateTemp = els.innerText
    totRow = getTotalRowOnPage(els)
    
    For i = Cells(r, c).Value To Cells(r + (totRow - 1), c).Value
    
        If (i > Date) Then
            Exit For
        End If
        klog "i=" & i
        'Cells(r - 1, c).Value = readActualHour(i, els)
        Cells(r, c + 1).Value = readActualHour(i, els)
        Cells(r, c + 1).Font.Bold = False
        If IsEmpty(Cells(r, c + 1).Value) Then
            Cells(r, c + 2).Value = readLeaveHolidayHour(i, els)
            Cells(r, c + 2).Font.Bold = False
        End If
        r = r + 1
    Next
    '
End Function
Private Function readActualHour(date_pm, els_pm)
    Dim strTemp As String
    Dim j As Integer
    strTemp = Format(date_pm, "dd-mmm-yyyy")
    For Each el In els_pm
        'klog "k" & el.innerText
        If (el.innerText = strTemp) Then
'        st = el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText
'            If (InStr(1, st, "Leave")) Then
'             result = j
'             Else
'             result = False
'             j = j + 1
'             End If
            'klog el.NextSibling.innerText
            'klog el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText
            'Cells(r - 1, c).Value = el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText
            'c = c + 1
            If (el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText <> " ") Then
                readActualHour = el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText
            End If
            Exit For
        End If
    Next
End Function

Private Function getTotalRowOnPage(els_pm)
    Dim strTemp As String
    Dim j As Integer
    j = 0
    strTemp = empID
    For Each el In els_pm
        If (el.innerText = strTemp) Then
        j = j + 1
        End If
    Next
    getTotalRowOnPage = j
End Function

Private Function readLeaveHolidayHour(date_pm, els_pm)
    Dim strTemp As String
    Dim j As Integer
    strTemp = Format(date_pm, "dd-mmm-yyyy")
    For Each el In els_pm
        'klog "k" & el.innerText
        If (el.innerText = strTemp) Then
'        st = el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText
'            If (InStr(1, st, "Leave")) Then
'             result = j
'             Else
'             result = False
'             j = j + 1
'             End If
            'klog el.NextSibling.innerText
            'klog el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText
            'Cells(r - 1, c).Value = el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText
            'c = c + 1
            If (el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText = " ") Then
             If (el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText Like "Holiday" Or el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText Like "Holiday" Or el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText Like "* Leave" Or el.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.innerText Like "* Leave") Then
                readLeaveHolidayHour = "8:00"
                End If
            End If
            Exit For
        End If
    Next
End Function

Private Function success()
    Set IE = Nothing
    MsgBox "SUCCESSfully DONE!!!"
End Function




'''''''''''''''''''''''''''''' Yesterday and Today
Private Sub Today_Click()
    Today_ClickFn
    success
End Sub

Sub today_yesterday()
Dim c As Range
For Each c In Range("H5:H39")

        If c.Value = Date - 1 Then
        Range("I" & c.Row).Value = Range("L64").Value
        End If
        
        If c.Value = Date Then
        Range("I" & c.Row).Value = Range("U64").Value
        End If
    Next c
End Sub
Private Sub Today_ClickFn()
    refreshToStartApp
    '
    navigateTo reportPage
    'Today's and Yesterday's Swipe Log
    dummyReturn = followLink("Today's and Yesterday's Swipe Log", "text")
    'For Emp -> dropdown ID on "Actual Hour" page
    dummyReturn = selectDropdown("title", "EmployeeID", empID)
    'dropdown ID on "Today's and Yesterday's Swipe Log " page
    dummyReturn = selectDropdown("title", "Day", "Today")
    'Generate report
    dummyReturn = followLink("ViewReportImageButton", "id")
    clear_SwipeData swipeColNum_today
    getDataToCells_SwipeLogTable swipeColNum_today
    'calculateSwipeLogs swipeColNum_today
    '
End Sub

Private Sub Yesterday_Click()
    Yesterday_ClickFn
    success
End Sub

Private Sub Yesterday_ClickFn()
    refreshToStartApp
    clear_SwipeData swipeColNum_yesterday
    '
    navigateTo reportPage
    'Today's and Yesterday's Swipe Log
    dummyReturn = followLink("Today's and Yesterday's Swipe Log", "text")
    'For Emp -> dropdown ID on "Actual Hour" page
    dummyReturn = selectDropdown("title", "EmployeeID", empID)
    'dropdown ID on "Today's and Yesterday's Swipe Log " page
    dummyReturn = selectDropdown("title", "Day", "Yesterday")
    'Generate report
    dummyReturn = followLink("ViewReportImageButton", "id")
    getDataToCells_SwipeLogTable swipeColNum_yesterday
    'calculateSwipeLogs swipeColNum_yesterday
    '
End Sub



Private Function getDataToCells_SwipeLogTable(colNum_pm)
    On Error GoTo ERROR_HANDLE:
    klog "Do not delete it please..." & CInt(IE.Document.getElementById("ReportViewer1_ctl01_ctl01_ctl04").innerText)
    For i = 1 To CInt(IE.Document.getElementById("ReportViewer1_ctl01_ctl01_ctl04").innerText)
        'MsgBox "kk" & i
        readSwipeLogTable colNum_pm
        If (i <> CInt(IE.Document.getElementById("ReportViewer1_ctl01_ctl01_ctl04").innerText)) Then
            dummyReturn = followLink("ReportViewer1$ctl01$ctl01$ctl05$ctl00$ctl00", "name")  'Next Page
        End If
    Next
    '
    calculateSwipeLogs colNum_pm
    '
    Exit Function
ERROR_HANDLE:
    MsgBox "[getDataToCells_SwipeLogTable] : " & vbCrLf & vbCrLf & _
            "Error...No Swipe Log Data Available for Yesterday or Today!"
End Function
Private Function readSwipeLogTable(colNum_pm)
    Dim t As IHTMLElement
    'Dim els As IHTMLElement
    Dim empIdTemp, dateTemp, gateTemp, entryExitTemp, SwipeTimeTemp As String
    Dim i As Integer
    i = 0
    Set objHTMLDocument = IE.Document
    Set els = objHTMLDocument.getElementsByTagName("div")
    For Each el In els
        If (el.innerText = empID) Then
            empIdTemp = ""
            dateTemp = ""
            gateTemp = ""
            entryExitTemp = ""
            SwipeTimeTemp = ""
            klog ("=======================")
            'MsgBox el.innerText
            klog ("el.innerText=" & el.innerText)
            empIdTemp = el.innerText
            '
            Set elTemp = el.ParentNode.NextSibling.FirstChild
            'MsgBox elTemp.nodeName & ", " & elTemp.ParentNode.nodeName & ", " & elTemp.ParentNode.NextSibling.nodeName & ", " & elTemp.ParentNode.NextSibling.FirstChild.nodeName
            klog ("elTemp.innerText=" & elTemp.innerText)
            dateTemp = elTemp.innerText
            '
            Set elTemp = elTemp.ParentNode.NextSibling.FirstChild
            'MsgBox elTemp.nodeName & ", " & elTemp.ParentNode.nodeName & ", " & elTemp.ParentNode.NextSibling.nodeName & ", " & elTemp.ParentNode.NextSibling.FirstChild.nodeName
            klog ("elTemp.innerText=" & elTemp.innerText)
            gateTemp = elTemp.innerText
            '
            Set elTemp = elTemp.ParentNode.NextSibling.FirstChild
            'MsgBox elTemp.nodeName & ", " & elTemp.ParentNode.nodeName & ", " & elTemp.ParentNode.NextSibling.nodeName & ", " & elTemp.ParentNode.NextSibling.FirstChild.nodeName
            klog ("elTemp.innerText=" & elTemp.innerText)
            entryExitTemp = elTemp.innerText
            '
            Set elTemp = elTemp.ParentNode.NextSibling.FirstChild
            'MsgBox elTemp.nodeName & ", " & elTemp.ParentNode.nodeName & ", " & elTemp.ParentNode.NextSibling.nodeName & ", " & elTemp.ParentNode.NextSibling.FirstChild.nodeName
            klog ("elTemp.innerText=" & elTemp.innerText)
            SwipeTimeTemp = elTemp.innerText
            '
            Cells(swipeRowNum, colNum_pm).Value = empIdTemp 'el.innerText
            Cells(swipeRowNum, colNum_pm + 1).Value = dateTemp 'elTemp.innerText
            Cells(swipeRowNum, colNum_pm + 2).Value = gateTemp 'elTemp.innerText
            Cells(swipeRowNum, colNum_pm + 3).Value = entryExitTemp 'elTemp.innerText
            Cells(swipeRowNum, colNum_pm + 4).Value = SwipeTimeTemp 'elTemp.innerText
            '
            swipeRowNum = swipeRowNum + 1
        End If
    Next
End Function

Private Sub calculateSwipeLogs(col_pm)
    Dim r, c, i As Integer
    r = gFindStr(currWS, "Swipe - YesterDay").Row + 1 'getValueOfSettingsVar("swipeRowNum") + 1
    c = col_pm
    Dim entryExit As String
    Dim gate As String
    i = r
    '
    While Cells(i, c).Value <> ""
    'For i = r To 72 'swipeRowNum
        Cells(i, c + 6).Value = ""
        entryExit = Cells(i, c + 3).Value
        gate = Cells(i, c + 2).Value
        gate = identifyGate(gate)
        Cells(i, c + 5).Value = gate
        If (entryExit = "Exit" And gate = "MainGate") Then
            'campus area : if you do sustraction as -- this time minus minus just previous swipe
            
        ElseIf (entryExit = "Exit" And gate = "PlayGate") Then
            'play area : if you do sustraction as -- this time minus minus just previous swipe
            Cells(i, c + 6).Value = "PLAY"
        ElseIf (entryExit = "Exit" And gate = "WorkGate") Then
            'work area : if you do sustraction as -- this time minus minus just previous swipe
            Cells(i, c + 6).Value = "WORK"
        ElseIf (entryExit = "Entry" And gate = "WorkGate") Then
            'play area : if you do sustraction as -- this time minus minus just previous swipe
            Cells(i, c + 6).Value = "PLAY"
        End If
    'Next
    i = i + 1
    Wend
    If (col_pm = swipeColNum_today) Then
        Cells(i, c + 3).Value = "Exit" 'put  swipeRowNum for row
        Cells(i, c + 4).Value = Time 'put  swipeRowNum for row
        Cells(i, c + 5).Value = "WorkGate" 'put  swipeRowNum for row
        Cells(i, c + 6).Value = "WORK" 'put  swipeRowNum for row
    End If
    'take the addition of hours to report grid
    Dim strTemp As String
    strTemp = Format(Cells(r, c + 1).Value, "dd-mmm-yy")
    Dim cellTemp As Range
    Set cellTemp = gFindStr(currWS, strTemp)
    If (cellTemp Is Nothing) Then
        Set cellTemp = Cells(10, 16)
         Cells(10, 15).Value = "Yesterday Work Hours:="
    End If
    Cells(cellTemp.Row, cellTemp.Column + 1).Value = Cells(r - 1, c + 8).Value
    Cells(cellTemp.Row, cellTemp.Column + 1).Font.Bold = True
    'Cells(gFindStr(currWS, Format(Cells(r, c + 1).Value, "dd-mmm-yy")).Row, gFindStr(currWS, Format(Cells(r, c + 1).Value, "dd-mmm-yy")).Column + 1).Value = "=" & Cells(r - 1, c + 8).Address
End Sub
Private Function identifyGate(gate_pm)
    Dim strTemp As String
    
    For i = LBound(enterMainGate) To UBound(enterMainGate)
        If (InStr(gate_pm, enterMainGate(i))) Then
            identifyGate = "MainGate"
            Exit Function
        End If
    Next
    '
    For i = LBound(enterPlayGate) To UBound(enterPlayGate)
        If (InStr(gate_pm, enterPlayGate(i))) Then
            identifyGate = "PlayGate"
            Exit Function
        End If
    Next
    '
    For i = LBound(enterWorkGate) To UBound(enterWorkGate)
        If (InStr(gate_pm, enterWorkGate(i))) Then
            identifyGate = "WorkGate"
            Exit Function
        End If
    Next
    '
    identifyGate = ""
End Function


''''''''''''''''''''''''''''''''''''''''''''''
 Sub RefreshReport_Click()
    RefreshReportFn
End Sub

Public Function RefreshReportFn()
    Dim lastCell As Range
    Dim lr As Long
    'MsgBox "asd" & Cells(Rows.Count, "P").End(xlUp).Column
    'MsgBox (Cells(Rows.Count, "R").End(xlUp))
    lr = Range("Q" & Cells.Rows.Count).End(xlUp).Row
    Set lastCell = Cells(lr, "Q")
    'MsgBox (lastCell.Address)
    lastCell = Time
    '
    Dim cellTemp1 As Range
    Dim strTemp As String
    strTemp = Format(Cells(lastCell.Row - 1, lastCell.Column - 3).Value, "dd-mmm-yy")
    Set cellTemp1 = gFindStr(gGetSheetByName("Sheet2"), strTemp)
    Set cellTemp2 = gFindStr(gGetSheetByName("Sheet2"), "Swipe - Today")
    Set cellTemp2 = Cells(cellTemp2.Row, cellTemp2.Column + 8)
    Cells(cellTemp1.Row, cellTemp1.Column + 1).Value = cellTemp2.Value
    '
'    RefreshReportTimer
End Function

'Private Function RefreshReportTimer()
'    If (RefreshTimerFlag <> True) Then
'        RefreshTimerFlag = True
'        alertTime = Now + TimeValue("00:00:05")
'        Application.OnTime alertTime, "kk"
'    End If
    
'End Function
'Public Sub kk()
'    MsgBox "kkkkkk"
'End Sub



''''''''''''''''''''''''''''''''''''''''''''''
Private Function clear_ReportGrid()
    'refreshToStartApp 'just for unit testing
    '
    Dim r, c As Integer
    r = gFindStr(currWS, "WEEK 1").Row
    c = gFindStr(currWS, "WEEK 1").Column + 2
    While (Cells(r, c).Value <> "")
        Cells(r, c).Value = ""
        Cells(r, c + 1).Value = ""
        Cells(r, c + 2).Value = ""
        r = r + 1
    Wend
End Function

Private Sub clear_SwipeData(col_pm)
'Private Sub clear_SwipeData() 'just for unit testing
    'refreshToStartApp 'just for unit testing
    '
    Dim r, c As Integer
    r = swipeRowNum
    c = col_pm 'swipeColNum_today
    While (Cells(r, c).Value <> "")
        Cells(r, c).Value = ""
        Cells(r, c + 1).Value = ""
        Cells(r, c + 2).Value = ""
        Cells(r, c + 3).Value = ""
        Cells(r, c + 4).Value = ""
        Cells(r, c + 5).Value = ""
        Cells(r, c + 6).Value = ""
        r = r + 1
    Wend
    Cells(r, c).Value = ""
    Cells(r, c + 1).Value = ""
    Cells(r, c + 2).Value = ""
    Cells(r, c + 3).Value = ""
    Cells(r, c + 4).Value = ""
    Cells(r, c + 5).Value = ""
    Cells(r, c + 6).Value = ""
End Sub

