Function save*sap_file(ByVal sess As Object, ByVal fPath As String, ByVal fName As String, ByVal sheet_name As String, *
ByVal description As String, Optional ByVal bCheckIfEmpty As Boolean = False, \_
Optional ByVal bDeleteBlankHeader As Boolean = True) As Variant
Dim err_wnd As ctrl_check, run_check, msg

    Application.StatusBar = "Exporting data from SAP...."

    ' Make sure wnd[1] exists
    err_wnd = Exist_Ctrl(sess, 1, "", True)
    ' ZMDESNR
    ' session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
    If err_wnd.cband Then
        ' First part of export (Menu and button will be done by tCode-specific script
        ' Below fills in export information
        dbLog.log "Found window title: (" & err_wnd.ctext & ")." & vbCrLf & "Extracting to filename: (" _
                & wPath & "\" & fName & ")" & vbCrLf

        ' check if cband = msg window
        err_wnd = Exist_Ctrl(sess, 1, "/usr/txtMESSTXT1", True)
        If err_wnd.cband Then
            msg = get_sap_text_errors(sess, 1, "/usr/txtMESSTXT1", 10)
            If Len(msg) > 0 Then GoTo errlv
        End If

        With sess
            .FindById("wnd[1]/usr/ctxtDY_PATH").text = (fPath) ' Fill in $PATH
            .FindById("wnd[1]/usr/ctxtDY_FILENAME").text = fName ' Fill in $FILENAME
            .FindById("wnd[1]/tbar[0]/btn[0]").press
            '.findById("wnd[1]").sendVKey 0
        End With
        run_check = True
    Else
        dbLog.log "Error Exporting report: (" & description & ")"
        If msgSwitch Then MsgBox "Error Exporting report: (" & description & ")"
        run_check = False
    End If

    Time_Event

    Application.StatusBar = "Saving file..."

    ' Workbooks(nmExportBook).Close
    Dim sBook: Set sBook = wait_for_file_open(fName)
    If sBook Is Nothing Then Set sBook = Workbooks.Open(fPath & "\" & fName)
    ' Copy exported WB (Sheet1) to sheet on ThisWorkbook with same name
    Dim oWS, ext: ext = dbLog.getfileext(sBook)
    If run_check Then
        Select Case True
        Case Contains(ext, "xl")
            Set oWS = CopyPaste(tWB, sBook, UCase(rename_sheet_if_exists(tWB, sheet_name, "_")), _
                                bDeleteBlankHeader:=bDeleteBlankHeader)
        Case Contains(ext, "txt")
            Set oWS = CopyPaste_alt(tWB, sBook, UCase(rename_sheet_if_exists(tWB, sheet_name, "_")), _
                                bDeleteBlankHeader:=bDeleteBlankHeader)
        Case Else
        End Select
    End If
    If run_check Then Set save_sap_file = oWS
    ' default for blank headers
    If found_blank_header(oWS) Then
        Call FillBlankHeader(oWS.Parent, oWS.name, DEFAULT_BLANK_FILL_HEADER)
    End If

    ' bCheckIfEmpty
    If bCheckIfEmpty Then
        If sheet_is_empty(oWS.Parent, oWS.name) Then
            msg = "bCheckIfEmpty: Sheet (" & oWS.name & ") is empty."
            GoTo errlv  ' return nothing
        Else
            msg = "bCheckIfEmpty: Sheet (" & oWS.name & ") has useable data."
        End If
    End If

    ' delete blank rows
    Evaluate DeleteEmptyRows(oWS)


    'MsgBox "Report " & tCode & " copied"
    'Application.Workbooks.Open (wPath & "\" & nmExportBook)
    Time_Event

    Application.DisplayAlerts = True
    Application.StatusBar = False
    Exit Function

errlv:
dbLog.log "Encountered msg: " & msg & " in save_sap_file for file: " & fName, msgPopup:=rMsgSwitch.value, msgType:=vbInformation
Set save_sap_file = Nothing
End Function
