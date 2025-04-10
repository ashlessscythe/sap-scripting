Function VT11*Run_Export(ByVal xlb As Workbook, sDate, eDate, tCode, SAPVariantName, LayoutRow, sess, *
Optional Limiter As String, Optional sheetin, Optional HeaderIn, \_
Optional ByVal byDate As Boolean = True, Optional ByVal obj As Variant) As Variant
DoEvents
Dim err_ctl As ctrl_check
Dim i
Dim listRows As New Collection
Dim ctrl_msg As String
Dim run_check As Boolean
Dim j, maxtries, aux_str

    ' Prevent Inf Loop with retry
    maxtries = 3
    j = 0

retry:
' Check if tCode is active
run_check = check_tcode(sess, tCode)

    With sess
        If Not IsEmpty(SAPVariantName) Then
            .FindById("wnd[0]/tbar[1]/btn[17]").press ' Choose variant
            .FindById("wnd[1]/usr/txtV-LOW").text = SAPVariantName ' Enter variant name
            .FindById("wnd[1]/usr/txtENAME-LOW").text = "" ' Blank username
            .FindById("wnd[1]/tbar[0]/btn[8]").press ' Close variant select window
        Else
            .FindById("wnd[0]/usr/ctxtK_STTRG-LOW").text = "1" ' In case of empty variantname enter 1 to 7
            .FindById("wnd[0]/usr/ctxtK_STTRG-HIGH").text = "7" ' for Overall Transport Status
        End If

        Select Case byDate
        Case False
            ' .findById("wnd[0]/usr/txtK_TPBEZ-LOW").text = ((Format(sDate, "mm/dd/yyyy")) & "*") ' Date in Description
            ' .findById("wnd[0]/usr/txtK_TPBEZ-HIGH").text = ((Format(eDate, "mm/dd/yyyy")) & "*") ' Date in Description
            ' Updated 190226 remove high date
            ' If sDate = eDate Then                ' Leave end date blank of dates match
            '     .findById("wnd[0]/usr/txtK_TPBEZ-HIGH").text = ""
            ' End If
        Case True
            .FindById("wnd[0]/usr/ctxtK_DATEN-LOW").text = (Format(sDate, "mm/dd/yyyy")) ' Start Date
            .FindById("wnd[0]/usr/ctxtK_DATEN-HIGH").text = (Format(eDate, "mm/dd/yyyy")) ' End Date
            If sDate = eDate Then
                .FindById("wnd[0]/usr/ctxtK_DATEN-HIGH").text = ""
            End If
        Case Else
            MsgBox "Invalid WHSE in VT11"
            End
        End Select
        If Len(Limiter) > 0 Then
            Select Case LCase(Limiter)
            Case "delivery"
                Dim sWS: Set sWS = xlb.Sheets(sheetin)
                If j < maxtries Then
                    Call set_clipboard(xlb, sheetin, FindColumn(sWS, HeaderIn))
                Else
                    Call CopyColumn(xlb, sheetin, FindColumn(sWS, HeaderIn))
                End If
                .FindById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press ' Multi Outbound Deliv Button
                .FindById("wnd[1]").SendVKey 16  ' Clear Previous entries
                .FindById("wnd[1]").SendVKey 24  ' Grab from Clipboard


                run_check = check_multi_paste(sess, tCode, 1, 0) ' Check if items were pasted
                Do While Not run_check
                    'Call close_popups(Sess)
                    dbLog.log "Paste not successful, retrying..."
                    ' Run tCode again, killing popups
                    Call check_tcode(sess, tCode, 1, 1)
                    j = j + 1
                    GoTo retry
                Loop

                .FindById("wnd[1]").SendVKey 8   ' Close Popup
            Case "date_range"
                ' Blank 2nd description to prevent
                .FindById("wnd[0]/usr/txtK_TPBEZ-HIGH").text = ""
                Dim obj2 As New BetterArray
                Set obj2 = obj
                Call ClipBoard_SetData(obj2.ToString(0, vbCrLf, "", "", 0))
                .FindById("wnd[0]/usr/btn%_K_TPBEZ_%_APP_%-VALU_PUSH").press ' Multi description of shipment button
                ' check window exists
                err_ctl = Exist_Ctrl(sess, 1, "", True)
                If err_ctl.cband Then
                    .FindById("wnd[1]").SendVKey 24 ' clipboard
                    .FindById("wnd[1]").SendVKey 8 ' f8 close popup
                End If
            Case Else
            End Select
        End If


        .FindById("wnd[0]/tbar[1]/btn[8]").press ' Execute

        ' Check for error (No Shipments Found)
        err_ctl = Exist_Ctrl(sess, 1, "/usr/txtMESSTXT1", "")
        If err_ctl.cband = True Then
            Select Case .FindById("wnd[1]/usr/txtMESSTXT1").text
            Case "No shipments were found for the selection criteria"
                If msgSwitch Then
                    MsgBox "No shipments found from dates (" & sDate & " to " & eDate & ") ending execution..."
                End If
                ' Close window then goto err
                .FindById("wnd[1]").Close
                GoTo errlv
            Case Else
            End Select
        End If

        ' Layout Selection
        .FindById("wnd[0]/mbar/menu[3]/menu[0]/menu[1]").Select ' Choose Layout Button

wndCheck:
Dim bNoLayout As Boolean
' Check if statusbar says "No layouts found"
aux_str = Hit_Ctrl(sess, 0, "/sbar", "Text", "Get", "")
If Contains(LCase(aux_str), "no layouts found") Then
bNoLayout = True
End If
' Check if window exists
err_ctl = Exist_Ctrl(sess, 1, "", True)
' Check if layout exists
Select Case True
Case IsEmpty(LayoutRow) Or Len(LayoutRow) = 0
' If layout is empty or zero-length, close popup window and export as-is
If err_ctl.cband Then
.FindById("wnd[1]").Close
End If
dbLog.log "Layout (" & LayoutRow & ") is empty or zero-length. " & vbCrLf & "Exporting as-is."

        Case IsNumeric(LayoutRow)
            err_ctl = Exist_Ctrl(sess, 1, "/usr/lbl[1," & LayoutRow & "]", True)
            If Not err_ctl.cband Then
                ' Current - if numeric not found, decrease until found.
                dbLog.log "Layout number (" & LayoutRow & ") not found, decreasing"
                LayoutRow = LayoutRow - 1
                GoTo wndCheck
            Else
                ctrl_msg = err_ctl.ctext
                .FindById("wnd[1]/usr/lbl[1," & LayoutRow & "]").SetFocus ' Highlight LayoutRow
                .FindById("wnd[1]").SendVKey 2   ' Select
                dbLog.log "Layout number (" & LayoutRow & "), (" & ctrl_msg & ") selected."
            End If

        Case Else
            ' Check if window exists
            err_ctl = Exist_Ctrl(sess, 1, "", True)
            ' Loop through available saved layouts (Still needs work)
            If err_ctl.cband Then
                For i = 1 To 30                  ' layouts start at 3
                    err_ctl = Exist_Ctrl(sess, 1, "/usr/lbl[1," & i & "]", True)
                    If err_ctl.cband Then
                        ctrl_msg = .FindById("wnd[1]/usr/lbl[1," & i & "]").text
                        If UCase(ctrl_msg) = UCase(LayoutRow) Then
                            .FindById("wnd[1]/usr/lbl[1," & i & "]").SetFocus
                            .FindById("wnd[1]").SendVKey 2 ' Select
                            dbLog.log "Layout number (" & i & "), (" & LayoutRow & ") selected."
                            GoTo layoutFound     ' Exit loop if (string) layout found
                        End If
                    End If
                Next
                ' // If we get here, layout wasn't found above (1 to 30)
                bNoLayout = True
            End If
        End Select

        If bNoLayout Then
            ' Current: If layout not found, close popup window and Setup_Layout_li()
            ' Future: if string layout not found, set it up.
            err_ctl = Exist_Ctrl(sess, 1, "", True)
            If err_ctl.cband Then
                .FindById("wnd[1]").Close
            End If

            dbLog.log "Layout (" & LayoutRow & ") not found." & vbCrLf & "Setting up layout " & Format(Now, time_form)
            Call Populate_Collection_Vertical(xlb, "Layouts", LCase(LayoutRow & "_" & tCode & "_layoutrows"), listRows, 0)
            Debug.Print listRows.Count
            .FindById("wnd[0]/mbar/menu[3]/menu[0]/menu[0]").Select ' Open Current Layout
            Call SetupLayout_li _
                 (sess, tCode, 1, "/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810", LayoutRow, listRows, 200)

        End If

layoutFound:

        .FindById("wnd[0]/mbar/menu[0]/menu[10]/menu[0]").Select ' Export to Excel
        ''''''''''''Check Export Window Name ''''''''''''''''''''''''''
        run_check = Check_Export_Window(sess, tCode, "SHIPMENT LIST: PLANNING")
    End With

    VT11_Run_Export = tCode
    Exit Function

errlv:
VT11_Run_Export = False
End Function

Function VL06O*DeliveryList_Run_Export(ByVal xlb As Workbook, sDate, eDate, SAPVariantName, *
sess, tCode, LayoutRow, sheetName, header) As Variant
Dim SrcSheet, searchString As String
Dim run_check As Boolean
Dim run_msg, i
Dim listRows As New Collection
Dim err_msg As String, ctrl_msg As String
Dim err_ctl As ctrl_check
Dim j, maxtries
Dim ws: Set ws = xlb.Sheets(sheetName)
' Prevent Inf Loop with retry
maxtries = 3
j = 0
retry:
' Check if tCode is active
run_check = check_tcode(sess, tCode)

    If j < maxtries Then
        Call set_clipboard(xlb, sheetName, FindColumn(ws, header))
    Else
        Call CopyColumn(xlb, sheetName, FindColumn(ws, header))
    End If

    On Error GoTo errlv:

    With sess
        .FindById("wnd[0]/usr/btnBUTTON6").press ' Press -List Outbound Deliveries-
        .FindById("wnd[0]/usr/ctxtIT_WADAT-LOW").text = "" ' Clear sDate and eDate
        .FindById("wnd[0]/usr/ctxtIT_WADAT-HIGH").text = ""
        .FindById("wnd[0]/usr/btn%_IT_TKNUM_%_APP_%-VALU_PUSH").press ' Multi Shipment Number

        .FindById("wnd[1]").SendVKey 16          ' Clear Previous entries
        .FindById("wnd[1]").SendVKey 24          ' Grab from Clipboard


        run_check = check_multi_paste(sess, tCode, 1, 0) ' Check if items were pasted
        Do While Not run_check
            'Call close_popups(Sess)
            dbLog.log "Paste not successful, retrying..."
            ' Run tCode again, killing popups
            Call check_tcode(sess, tCode, 1, 1)
            j = j + 1
            GoTo retry
        Loop

        .FindById("wnd[1]").SendVKey 8           ' Close Multi-Window
        .FindById("wnd[0]").SendVKey 8           ' Execute
        '.findById("wnd[0]/tbar[1]/btn[33]").press                       ' Choose layout - HeaderView
        '.findById("wnd[1]").sendVKey 12                                 ' Escape
        .FindById("wnd[0]/tbar[1]/btn[18]").press ' Item View Button
        '.findById("wnd[0]/tbar[1]/btn[33]").press
        .FindById("wnd[0]/mbar/menu[3]/menu[2]/menu[1]").Select ' Choose Layout

wndCheck:
' Check if layout exists
Select Case True
Case IsEmpty(LayoutRow) Or Len(LayoutRow) = 0
' If layout is empty or zero-length, close popup window and export as-is
.FindById("wnd[1]").Close
dbLog.log "Layout (" & LayoutRow & ") is empty or zero-length. " & vbCrLf & "Exporting as-is."

        Case IsNumeric(LayoutRow)
            err_ctl = Exist_Ctrl(sess, 1, "/usr/lbl[1," & LayoutRow & "]", True)
            If Not err_ctl.cband Then
                dbLog.log "Layout number (" & LayoutRow & ") not found, decreasing"
                LayoutRow = LayoutRow - 1
                GoTo wndCheck
            Else
                err_msg = err_ctl.ctext
                .FindById("wnd[1]/usr/lbl[1," & LayoutRow & "]").SetFocus ' Highlight LayoutRow
                .FindById("wnd[1]").SendVKey 2   ' Select
                dbLog.log "Layout number (" & LayoutRow & "), (" & err_msg & ") selected."
            End If

        Case Else
            ' Check if window exists
            Dim scrl_pos
            err_ctl = Exist_Ctrl(sess, 1, "", True)
            If err_ctl.cband = True Then
                ' Get Scrollbar Position
                scrl_pos = .FindById("wnd[1]/usr").verticalScrollbar.Position
                Debug.Print "Vertical Scrollbar Position is " & scrl_pos
                ' Scroll Up to top (0)
                ctrl_msg = Hit_Ctrl(sess, 1, "/usr", "Position", "SetV", (0))
                j = 0
            End If
            ' Loop through available saved layouts (Still needs work)
            If err_ctl.cband Then
                For i = 1 To 30                  ' layouts start at 3
                    err_ctl = Exist_Ctrl(sess, 1, "/usr/lbl[1," & i & "]", True)
                    If err_ctl.cband Then
                        ctrl_msg = .FindById("wnd[1]/usr/lbl[1," & i & "]").text
                        If UCase(ctrl_msg) = UCase(LayoutRow) Then
                            .FindById("wnd[1]/usr/lbl[1," & i & "]").SetFocus
                            .FindById("wnd[1]").SendVKey 2 ' Select
                            dbLog.log "Layout number (" & i & "), (" & LayoutRow & ") selected."
                            Exit For
                        End If
                    End If
                Next
                ' Current: If layout not found, close popup window and Setup_Layout_li()
                ' Future: if string layout not found, set it up.
                err_ctl = Exist_Ctrl(sess, 1, "", True)
                If err_ctl.cband Then
                    .FindById("wnd[1]").Close
                    dbLog.log "Layout (" & LayoutRow & ") not found." & vbCrLf & "Setting up layout " & Format(Now, time_form)
                    Call Populate_Collection_Vertical(xlb, "Layouts", LCase(LayoutRow & "_" & tCode & "_layoutrows"), listRows, 0)
                    Debug.Print listRows.Count
                    .FindById("wnd[0]/mbar/menu[3]/menu[2]/menu[0]").Select ' Open Current Layout
                    Call SetupLayout_li _
                         (sess, tCode, 1, "/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810", LayoutRow, listRows, 200)
                End If
            End If
        End Select

        err_msg = Hit_Ctrl(sess, 0, "/sbar", "Text", "Get", "")
        dbLog.log "Statusbar message: (" & err_msg & ")"

        .FindById("wnd[0]/mbar/menu[0]/menu[5]/menu[1]").Select ' Export as Excel
        ''''''''''''Check Export Window''''''''''''''''''
        run_check = Check_Export_Window(sess, tCode, "LIST OF OUTBOUND DELIVERIES")
    End With

    VL06O_DeliveryList_Run_Export = tCode
    Exit Function

errlv:

    run_msg = err.description & vbCrLf & run_msg
    VL06O_DeliveryList_Run_Export = False

End Function

Function ZMDESNR*With_Exclude_Export *
(ByVal xlb As Workbook, tCodeDescription, tCode, SAPVariantName, sess, LayoutRow, \_
SrcSheet, strFind, exclSheet, exclStr) As params
Dim DeliveryCount
Dim run_check As Boolean
Dim err_wnd As ctrl_check
Dim bar_msg As String
Dim listRows As New Collection
Dim j, maxtries
Dim sWS: Set sWS = xlb.Sheets(SrcSheet)
Dim local_rVal As params

    DoEvents

    ' Prevent inf loop with retry
    j = 0
    maxtries = 2

retry:
' Check if tCode is active
Evaluate assert_tcode(sess, tCode)

    If j < maxtries Then
        Call set_clipboard(xlb, SrcSheet, FindColumn(sWS, strFind))
    Else
        Call CopyColumn(xlb, SrcSheet, FindColumn(sWS, strFind))
    End If

    ' Fill dictionary

    With sess

        If Not IsEmpty(SAPVariantName) Then
            .FindById("wnd[0]/tbar[1]/btn[17]").press ' Open Variant Select window
            .FindById("wnd[1]/usr/txtV-LOW").text = SAPVariantName ' Variant Name
            .FindById("wnd[1]/usr/txtENAME-LOW").text = "" ' Created by blank
            .FindById("wnd[1]/tbar[0]/btn[8]").press ' Close Variant Select window
        End If
        .FindById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2").Select ' General Selection tab
        '.findById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2").Select
        .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9002/ctxtS_PALLTD-LOW") _
        .text = ""                               ' Remove 'X' from Palletized
        .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9002/txtS_VBELN-LOW") _
        .text = ""                               ' Low Delivery Number
        .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9002/txtS_VBELN-HIGH") _
        .text = ""                               ' High Delivery Number
        .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9002/btn%_S_VBELN_%_APP_%-VALU_PUSH") _
        .press                                   ' Press Multi Delivery Entry button
        '.findById("wnd[0]").sendVKey 0                                  ' Enter

        .FindById("wnd[1]").SendVKey 16          ' Clear Previous entries
        .FindById("wnd[1]").SendVKey 24          ' Grab from Clipboard


        run_check = check_multi_paste(sess, tCode, 1, 0) ' Check if items were pasted
        Do While Not run_check
            'Call close_popups(Sess)
            dbLog.log "Paste not successful, retrying..."
            ' Run tCode again, killing popups
            Call check_tcode(sess, tCode, 1, 1)
            j = j + 1
            GoTo retry
        Loop

        .FindById("wnd[1]").SendVKey 8           ' Close (Save) Popup

        ''''''''''''
        j = 0
        maxtries = 2

retry2:
Dim exWS: Set exWS = xlb.Sheets(exclSheet)

        ' exclude highlights?
        Evaluate Repack_Filter_Exclude(xlb, exclSheet)
        Dim cfc As Integer
        cfc = Count_Filtered_Cells(xlb, exclSheet)
        If cfc < 2 Then GoTo execute

        If j < maxtries Then
            Call set_clipboard(xlb, exclSheet, FindColumn(exWS, exclStr))
        Else
            Call CopyColumn(xlb, exclSheet, FindColumn(exWS, exclStr))
        End If
        ''''''''''''

        '.findById _
        ("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9002/btn%_S_SERILS_%_APP_%-VALU_PUSH") _
        .press                                   ' Open Multi Serial Number Popup
        .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9002/btn%_S_PARENT_%_APP_%-VALU_PUSH") _
        .press                                   ' Open Multi Parent SN Popup
        .FindById _
        ("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV") _
        .Select                                  ' Select Exclude Tab

        .FindById("wnd[1]").SendVKey 16          ' Clear Previous entries
        .FindById("wnd[1]").SendVKey 24          ' Grab from Clipboard


        run_check = check_multi_paste(sess, tCode, 1, 0) ' Check if items were pasted
        Do While Not run_check
            'Call close_popups(Sess)
            dbLog.log "Paste not successful, retrying..."
            ' Run tCode again, killing popups
            Call check_tcode(sess, tCode, 1, 1)
            j = j + 1
            GoTo retry2
        Loop

        .FindById("wnd[1]").SendVKey 8           ' Close Popup

execute:
.FindById("wnd[0]/tbar[1]/btn[8]").press ' Execute

        If Not IsEmpty(LayoutRow) Then
            .FindById("wnd[0]/tbar[1]/btn[33]") _
        .press
            run_check = SelectLayout _
                        (sess, 1, "/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell", LayoutRow)
            ' If above fails
            If Not run_check Then
                ' Check if wnd1 exists
                err_wnd = Exist_Ctrl(sess, 1, "", True)
                ' Close if true
                If err_wnd.cband = True Then .FindById("wnd[1]").Close
                ' Open Current Layout window
                .FindById("wnd[0]/tbar[1]/btn[32]").press
                ' Begin Setup
                dbLog.log "Setting up layout " & Format(Now, time_form)
                Call Populate_Collection_Vertical(xlb, "Layouts", LayoutRow & "_" & tCode & "_layoutrows", listRows)
                Debug.Print listRows.Count
                run_check = SetupLayout _
                            (sess, 1, "/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620", LayoutRow, listRows, 200)
            End If
        End If

        ''''''''''''''Check Status bar''''''''''''''''''
        bar_msg = Hit_Ctrl(sess, 0, "/sbar", "Text", "Get", "")
        Select Case bar_msg
        Case ""
        Case "No layouts found"
            dbLog.log "Statusbar msg (" & bar_msg & ") for layout " & LayoutRow
            GoTo errlv
        Case Else
            dbLog.log "Statusbar msg (" & bar_msg & ") for layout " & LayoutRow
        End Select
        ''''''''''''''''''''''''''''''''''''''''''''''''

        .FindById _
        ("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select ' Export as Excel

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''Check Export Window Name ''''''''''''''''''''''''''
        local_rVal.run_check = Check_Export_Window(sess, tCode, "ZMDEMAIN SERIAL NUMBER HISTORY CONTENTS")
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    End With

out:
Call ClearClipboard

    If local_rVal.run_check Then
        local_rVal.name = sanitize(tCodeDescription)
        ZMDESNR_With_Exclude_Export = local_rVal
    End If
    Exit Function

errlv:
If Len(local_rVal.err) < 1 Then local_rVal.err = "Error: " & err.description
GoTo out
End Function

Function ZMDETPC*PalletCount_Run_Export *
(ByVal xlb As Workbook, sDate, eDate, SAPVariantName, tCode, \_
sess, LayoutRow, Radial, SrcSheet, searchString) As Variant
Dim run_check As Boolean
Dim err_wnd As ctrl_check
Dim err_msg As String
Dim j, maxtries
Dim sWS: Set sWS = xlb.Sheets(SrcSheet)
' assert date formats
sDate = replace_delim(Format(sDate, SAP_DATE_FORMAT), get_first_delim(SAP_DATE_FORMAT))
eDate = replace_delim(Format(eDate, SAP_DATE_FORMAT), get_first_delim(SAP_DATE_FORMAT))
' Prevent Inf Loop with retry
maxtries = 3
j = 0
retry:
' Check if tCode is active
run_check = check_tcode(sess, tCode)

    ' Copy Delivery numbers from VL06O sheet
    If j < maxtries Then
        Call set_clipboard(xlb, SrcSheet, FindColumn(sWS, searchString))
    Else
        Call CopyColumn(xlb, SrcSheet, FindColumn(sWS, searchString), 0)
    End If

    On Error GoTo errlv:

    With sess
        Evaluate assert_tcode(sess, tCode)
        Select Case Radial
        Case "Summarized"
            .FindById("wnd[0]/tbar[1]/btn[17]").press ' Open Variant Window
            .FindById("wnd[1]/usr/txtV-LOW").text = SAPVariantName ' Enter variant name
            .FindById("wnd[1]/usr/txtENAME-LOW").text = "" ' Blank username
            .FindById("wnd[1]/tbar[0]/btn[8]").press ' Close variant select window
            ' Call SelectLayout(Sess, "wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell", LayoutRow)
            '.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentcellrow = LayoutRow  ' Variant Row
            '.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedrows = LayoutRow
            '.findById("wnd[1]/tbar[0]/btn[2]").press                                    ' Select Variant
            ' low date
            .FindById("wnd[0]/usr/ctxtS_DATE-LOW").text = sDate
            ' high date
            .FindById("wnd[0]/usr/ctxtS_DATE-HIGH").text = eDate

            .FindById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press ' Multi Delivery Entry Window
            .FindById("wnd[1]").SendVKey 16      ' Clear Previous entries
            .FindById("wnd[1]").SendVKey 24      ' Grab from Clipboard

winClose:
' Delete first item if not numeric
err*wnd = Exist_Ctrl(sess, 1, "", True)
If err_wnd.cband Then
err_msg = *
.FindById _
("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]") _
.text
If Not IsNumeric(err_msg) Then
.FindById("wnd[1]/tbar[0]/btn[14]").press
End If

                ' Check paste after 2nd winclose, before 1st window closes
                run_check = check_multi_paste(sess, tCode, 1, 0) ' Check if items were pasted
                Do While Not run_check
                    'Call close_popups(Sess)
                    dbLog.log "Paste not successful, retrying..."
                    ' Run tCode again, killing popups
                    Call check_tcode(sess, tCode, 1, 1)
                    j = j + 1
                    GoTo retry
                Loop
                .FindById("wnd[1]").SendVKey 8   ' Close Multi Window
            End If

            ' Check if wnd2 exists
            err_wnd = Exist_Ctrl(sess, 2, "", True)
            ' Close if true
            If err_wnd.cband = True Then
                err_msg = .FindById("wnd[2]/usr/txtMESSTXT1").text
                .FindById("wnd[2]").SendVKey 0
                GoTo winClose
            End If


            .FindById("wnd[0]/usr/radR_SUMM").Select
            ' .findById("wnd[0]/usr/ctxtS_DATE-LOW").Text = (sDate - 7) ' Date range 7 by 7
            If .FindById("wnd[0]/usr/ctxtS_DATE-HIGH").text = "" Then
                .FindById("wnd[0]/usr/ctxtS_DATE-HIGH").text = eDate
            End If
            '.findById("wnd[0]/usr/radR_OPEN").Select
            .FindById("wnd[0]/usr/radR_COMPL").Select
            .FindById("wnd[0]/usr/radR_SUMM").Select

            '.findById("wnd[0]/usr/ctxtS_VSTEL-LOW").Text = shp0                        ' Ship Point
            '.findById("wnd[0]/usr/btn%_S_VSTEL_%_APP_%-VALU_PUSH").press                ' Multi Ship Point Entry
            '.findById _
            ("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = shp1
            '.findById("wnd[1]").sendVKey 8                                              ' Close Multi Window

            ' execute
            .FindById("wnd[0]/tbar[1]/btn[8]").press
            ' export
            .FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select

            ' Text with tabs radial
            run_check = Check_Export_Window(sess, tCode, "SAVE LIST IN FILE...")


        Case "Detail"
            '.findById("wnd[0]/usr/radR_OPEN").Select
            .FindById("wnd[0]/usr/radR_COMPL").Select
            '.findById("wnd[0]/usr/radR_DETA").Select
            '.findById("wnd[0]/usr/radR_SUMM").Select
            .FindById("wnd[0]/usr/ctxtS_DATE-LOW").text = date0
            .FindById("wnd[0]/usr/ctxtS_DATE-HIGH").text = date1
            '.findById("wnd[0]/usr/ctxtS_VSTEL-LOW").Text = shp0
            .FindById("wnd[0]/usr/btn%_S_VSTEL_%_APP_%-VALU_PUSH").press
            '.findById _
            ("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = shp1
            .FindById("wnd[1]").SendVKey 8
            .FindById("wnd[0]/usr/radR_DETA").Select
            .FindById("wnd[0]/tbar[1]/btn[8]").press
            .FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select

            ' Text with tabs radial
            run_check = Check_Export_Window(sess, tCode, "SAVE LIST IN FILE...")

        End Select

        '''''''''''''Check Export Window'''''''''''
        run_check = Check_Export_Window(sess, tCode, "TRACK PALLET COUNT")
    End With

    ZMDETPC_PalletCount_Run_Export = tCode
    Exit Function

errlv:
If msgSwitch Then MsgBox ("Error " & err.description)
ZMDETPC_PalletCount_Run_Export = False
End Function

Function ZVT11*From_Deliv(ByVal xlb As Workbook, sess, tCode, LayoutRow, ZVT_Type, sheetName, header, Optional ShipPt, *
Optional ByVal sDate, Optional ByVal eDate, Optional ByVal sapvar As String = "", Optional dateObj As Variant) As Variant
Dim i, local_rVal As params
Dim run_check As Boolean
Dim err_msg As String
Dim err_ctl As ctrl_check
Dim err_wnd As ctrl_check
Dim ctrl_msg As String
Dim bar_msg As String
Dim listRows As New Collection
DoEvents

    ' get dates from dateObj if provided
    Dim dt As dateObj
    If Not IsMissing(dateObj) Then
        dt = get_json_date(dateObj)
        sDate = dt.date1
        If Len(dt.date2) < 1 Then
            eDate = DateAdd("d", 6, sDate)  ' probably don't need to extend
        Else                                ' they're extended below
            eDate = dt.date2
        End If
    End If

    ' Extend dates
    If Not IsMissing(sDate) Or Not IsMissing(eDate) Then
        sDate = DateAdd("d", -100, sDate)
        eDate = DateAdd("d", 100, eDate)
    End If

    Dim j, maxtries

    On Error GoTo errlv

    ' prevent inf loop with retry
    maxtries = 2
    j = 0
    Dim ws: Set ws = xlb.Sheets(sheetName)

retry:
' Check if tCode is active
run_check = check_tcode(sess, tCode)

    If j < maxtries Then
        Call set_clipboard(xlb, sheetName, FindColumn(ws, header))
    Else
        Call CopyColumn(xlb, sheetName, FindColumn(xlb, header, sheetName))
    End If

    With sess
        Select Case ZVT_Type
        Case "Closed"
            ' Completed Shipment Tab
            .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB2").Select
            If Not IsMissing(ShipPt) Then
                If Not Len(ShipPt) = 0 Then
                    ' Ship Point
                    .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB2/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0102/ctxtS_VSTEL2-LOW").text = ShipPt
                End If
            End If
            If Not Len(sapvar) = 0 Then
                ' Open Variant Select Screen
                .FindById("wnd[0]/tbar[1]/btn[17]").press
                ' Enter Variant Name
                .FindById("wnd[1]/usr/txtV-LOW").text = sapvar
                ' Username Empty
                .FindById("wnd[1]/usr/txtENAME-LOW").text = ""
                ' Close Variant Select Screen
                .FindById("wnd[1]/tbar[0]/btn[8]").press
            End If
            ' Start Date
            If Not IsEmpty(sDate) Then
            .FindById _
                    ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB2/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0102/ctxtS_DATE_2-LOW") _
                    .text = sDate
            End If
            ' End Date
            If Not IsEmpty(eDate) Then
             .FindById _
                     ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB2/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0102/ctxtS_DATE_2-HIGH") _
                    .text = eDate
            End If
            ' Multi Delivery
            .FindById _
                ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB2/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0102/btn%_S_VBELN2_%_APP_%-VALU_PUSH") _
                .press


            .FindById("wnd[1]").SendVKey 16      ' Clear Previous entries
            .FindById("wnd[1]").SendVKey 24      ' Grab from Clipboard


            run_check = check_multi_paste(sess, tCode, 1, 0) ' Check if items were pasted
            Do While Not run_check
                'Call close_popups(Sess)
                dbLog.log "Paste not successful, retrying..."
                ' Run tCode again, killing popups
                Call check_tcode(sess, tCode, 1, 1)
                j = j + 1
                GoTo retry
            Loop

            ' Close Multi Window
            .FindById("wnd[1]").SendVKey 8
            ' Display All Deliveries Radial
            .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB2/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0102/radR1") _
        .Select
            ' Non post good issued delivery Radial
            '.findById("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB2/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0102/radR2").Select
            ' Display ALV true/false
            .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB2/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0102/chkP_ALV_2") _
        .Selected = True


        Case "Open"
            ' Open Shipments Tab
            .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB1").Select
            ' Open Variant Select Screen
            .FindById("wnd[0]/tbar[1]/btn[17]").press
            ' Enter Variant Name
            .FindById("wnd[1]/usr/txtV-LOW").text = sapvar
            ' Username Empty
            .FindById("wnd[1]/usr/txtENAME-LOW").text = ""
            ' Close Variant Select Screen
            .FindById("wnd[1]/tbar[0]/btn[8]").press
            ' Start Date
            '.findById _
            ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB1/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0101/ctxtS_DATE_1-LOW").text = sDate
            ' End Date
            '.findById _
            ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB2/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0101/ctxtS_DATE_1-HIGH").text = eDate
            ' Low Delivery
            '.findById _
            ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB1/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0101/ctxtS_VBELN1-LOW").Text = ""
            ' High Delivery
            '.findById _
            ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB1/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0101/ctxtS_VBELN1-HIGH").Text = ""
            ' Multi Delivery Number
            .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB1/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0101/btn%_S_VBELN1_%_APP_%-VALU_PUSH").press
            ' Clear Previous entries
            .FindById("wnd[1]").SendVKey 16
            ' Grab from clipboard
            .FindById("wnd[1]").SendVKey 24

            ' Check if items were pasted
            run_check = check_multi_paste(sess, tCode, 1, 0)
            Do While Not run_check
                'Call close_popups(Sess)
                dbLog.log "Paste not successful, retrying..."
                ' Run tCode again, killing popups
                Call check_tcode(sess, tCode, 1, 1)
                j = j + 1
                GoTo retry
            Loop

            ' Close Multi Window
            .FindById("wnd[1]").SendVKey 8
            ' Display ALV checkmark
            .FindById _
        ("wnd[0]/usr/tabsTABSTRIP_BLK1/tabpTAB1/ssub%_SUBSCREEN_BLK1:ZSDR_SHIPMENT_REPORT:0101/chkP_ALV_1").Selected = True
        Case Else
            dbLog.log ("Select Either Open or Closed Type for ZVT")
            If msgSwitch Then
                MsgBox ("Select Either Open or Closed Type for ZVT")
            End If
        End Select

        ' Execute
        .FindById("wnd[0]/tbar[1]/btn[8]").press

        ' Check if No Data window appears
        err_wnd = Exist_Ctrl(sess, 1, "/usr/txtMESSTXT1", True)
        If err_wnd.cband Then
            If err_wnd.ctext = "No data exists in the database for this selection" Then
                dbLog.log "Error (" & err_wnd.ctext & ") " & Format(Now, time_form)
                .FindById("wnd[1]").Close
                GoTo errlv
            End If
        End If

        If Not IsEmpty(LayoutRow) Then
            .FindById("wnd[0]/tbar[1]/btn[33]").press
            ' Check if layout exists
            local_rVal = check_select_layout(sess, tCode, LayoutRow, , 1)
        End If
    End With

    Call ClearClipboard
    ZVT11_From_Deliv = tCode
    Exit Function

errlv:
ZVT11_From_Deliv = False
End Function
