Function VL03N_Run_Export(ByVal xlb As Workbook, DeliveryNumber, SAPVariantName, sess, tCode) As Variant
Dim j, last As Double, run_check
j = 0
last = 100

    ' Check if tCode is active
    run_check = check_tcode(sess, tCode)


    With sess
        .FindById("wnd[0]/usr/ctxtLIKP-VBELN").text = DeliveryNumber
        .FindById("wnd[0]/tbar[1]/btn[18]").press ' Pack Button
        '.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS").Select        ' Pack Material Tab
        '.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6HUS").Select        ' Pack HUs Tab
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6INH").Select ' Ttl Content Tab

        dbLog.log .FindById _
                  ("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6INH/ssubTAB:SAPLV51G:6040/tblSAPLV51GTC_HU_005").Columns.ElementAt(1) _
                  .Count
        .FindById _
        ("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6INH/ssubTAB:SAPLV51G:6040/tblSAPLV51GTC_HU_005/txtHUMV4-IDENT[1," & j & "]") _
        .text = ""

        dbLog.log Trim(Right(.FindById _
                             ("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6INH/ssubTAB:SAPLV51G:6040/tblSAPLV51GTC_HU_005/txtHUMV4-IDENT[1," & j & "]") _
                             .text, 9))
        '.findById _
        ("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6INH/ssubTAB:SAPLV51G:6040/tblSAPLV51GTC_HU_005").Columns.elementAt(1) _
        .Copy
        '.findById("wnd[0]").sendVKey 26                                 ' Double-click
    End With

    VL03N_Run_Export = tCode

End Function

Sub vl06o_date_update()

    ' multi deliv
    ' Session.findById("wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH").press
    ' Session.findById("wnd[1]").sendVKey 24
    ' Session.findById("wnd[1]").sendVKey 8
    ' Session.findById("wnd[0]").sendVKey 8

    Dim err_ctrl As ctrl_check, counter As Double, rng As Range
    Dim sess As Variant: Set sess = S_(0)

    Set rng = Application.InputBox("Select delivery range", "Obtain Range Object", type:=8)

    If rng Is Nothing Then GoTo errlv
    ' ignore any blanks
    Set rng = rng.SpecialCells(xlCellTypeConstants)

    If rng Is Nothing Then GoTo errlv
    rng.Copy

    ' get Deliveries from selection into vl06o
    Dim tCode As String, sap_var As String
    tCode = "vl06o"
    sap_var = "blank_"
    With sess
        Evaluate assert_tcode(sess, tCode)
        .FindById("wnd[0]/usr/btnBUTTON6").press
            ' variant select window
            .FindById("wnd[0]").SendVKey 17
            ' traditional variant select
            .FindById("wnd[1]/usr/txtV-LOW").text = sap_var
            ' clear name
            .FindById("wnd[1]/usr/txtENAME-LOW").text = ""
            ' enter
            .FindById("wnd[1]").SendVKey 0
            ' close
            .FindById("wnd[1]").SendVKey 8
        ' multi deliv
        .FindById("wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH").press

        .FindById("wnd[1]").SendVKey 24
        .FindById("wnd[1]").SendVKey 8
        .FindById("wnd[0]").SendVKey 8
        .FindById("wnd[0]").SendVKey 5
        .FindById("wnd[0]").SendVKey 13
    End With


    dbLog.log "Starting vl06o date update"

    ' select all deliv
    Dim ans: ans = MsgBox("Start for " & (rng.Rows.Count - 1) & " deliveries?", vbYesNo + vbQuestion, "Start VL06O?")

    If ans <> vbYes Then GoTo errlv

    Dim sDate As String: sDate = InputBox("Enter target date: ", "Enter Date", _
                                          replace_delim(Format(DateAdd("d", 1, TODAYS_DATE), SAP_DATE_FORMAT), _
                                          get_first_delim(SAP_DATE_FORMAT)))

    If Not IsDate(sDate) Then GoTo errlv

    With sess
        counter = 0
        Dim msg As String
        If Exist_Ctrl(sess, 1, "", 1).cband Then
            msg = get_sap_text_errors(sess, 1, "/usr/txtMESSTXT1", 10)
            If Contains(msg, "loading") Then .FindById("wnd[0]").SendVKey 0
        End If
        ' select item overview tab (1st)
        .FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
        err_ctrl = Exist_Ctrl(sess, 0, "/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/ctxtLIKP-WADAT", 1)
        Do While err_ctrl.cband
        Dim deliv_number As String
        deliv_number = .FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV50A:1502/ctxtLIKP-VBELN").text
        dbLog.log str_form & "Working with delivery (" & deliv_number & ")"

            ' check if changeable
            If Not .FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/ctxtLIKP-WADAT").changeable Then
                dbLog.log "Delivery date not changeable", msgPopup:=True, msgType:=vbInformation
                ' yellow arrow back
                .FindById("wnd[0]/tbar[0]/btn[15]").press
                GoTo skip_loop
            Else
                ' Change date
                Dim original_date As String
                original_date = .FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/ctxtLIKP-WADAT").text
                .FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/ctxtLIKP-WADAT").text = sDate
                dbLog.log "Changing date from (" & original_date & ") to (" & sDate & ")"
            End If

            ' Enter loop
            Do
                dbLog.log Hit_Ctrl(sess, 0, "/sbar", "Text", "Get", "")
                ' press enter
                .FindById("wnd[0]").SendVKey 0
            Loop While Len(Hit_Ctrl(sess, 0, "/sbar", "Text", "Get", "")) > 1

            ' format deliv on excel
            Dim c As Range
            If Not Contains(original_date, sDate) Then
                Set c = rng.Find(deliv_number)
                c.Font.Bold = True
                c.Font.Italic = True
                c.Interior.Color = vbCyan
                ' note change on adjacent cell
                c.offset(0, 1).value = original_date & "=>" & sDate
            End If

            ' save
            .FindById("wnd[0]").SendVKey 11

            ' Yes, continue

retry_loop:
If Exist_Ctrl(sess, 1, "/usr/btnSPOP-OPTION1", 1).cband Then
.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
End If
If Exist_Ctrl(sess, 1, "", 1).cband Then
msg = get_sap_text_errors(sess, 1, "/usr/txtMESSTXT1", 10)
If Contains(msg, "loading") Then .FindById("wnd[0]").SendVKey 0
End If
err_ctrl = Exist_Ctrl(sess, 0, "/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/ctxtLIKP-WADAT", 1)
skip_loop:
counter = counter + 1
Loop
' check if done or err_msg
Dim bar_msg As String
bar_msg = Hit_Ctrl(sess, 0, "/sbar", "Text", "Get", "")
If Contains(bar_msg, "currently being") Then
' F3
dbLog.log "Error: (" & bar_msg & ")"
.FindById("wnd[0]").SendVKey 3
GoTo retry_loop
End If
dbLog.log "Done... with (" & counter & ") items.", msgPopup:=1
End With
Exit Sub
errlv:

End Sub

Sub delivery_qty_update()

    Dim err_ctrl As ctrl_check, counter As Double
    Dim sess As Variant: Set sess = S_(0)

    ' select all deliv
    Dim ans: ans = MsgBox("Start?", vbYesNo + vbQuestion, "Start qty update?")

    If ans <> vbYes Then GoTo errlv

    Dim deliv, pn, curr_val, set_val
    Dim rng As Range, r As Range, val As String, sap_val
    Set rng = Application.InputBox("Select a range", "Obtain Range Object", type:=8)

    Debug.Print rng.address
    Debug.Print rng.Parent.Parent.name
    counter = 0
    For Each r In rng.Rows
        deliv = rng.Cells(r.row, 1).value
        pn = rng.Cells(r.row, 6).value
        val = rng.Cells(r.row, 7)
        If Not IsNumeric(deliv) Or Not IsNumeric(pn) Then GoTo skip
        If Not Contains(val, "=>") Then GoTo skip
        curr_val = CInt(Split(val, "=>")(0))
        set_val = CInt(Split(val, "=>")(1))
        With sess
            Evaluate assert_tcode(sess, "vl02n")
            ' delivery
            .FindById("wnd[0]/usr/ctxtLIKP-VBELN").text = deliv
            .FindById("wnd[0]").SendVKey 0

            ' handle info window if loading commenced
            err_ctrl = Exist_Ctrl(sess, 1, "", True)
            Do While err_ctrl.cband
                ' enter to close
                If Contains(err_ctrl.ctext, "information") Then .FindById("wnd[1]").SendVKey 0
                err_ctrl = Exist_Ctrl(sess, 1, "", True)
            Loop

            Dim bar_msg As String
            bar_msg = Hit_Ctrl(sess, 0, "/sbar", "Text", "Get", "")
            If Contains(bar_msg, "currently being") Then
                ' in use
                Debug.Print
            End If

            ' select item overview tab (1st)
            .FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
            err_ctrl = Exist_Ctrl(sess, 0, "/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/ctxtLIKP-WADAT", 1)

            ' find
            .FindById("wnd[0]").SendVKey 71
            ' pn in 2nd field
            .FindById("wnd[1]/usr/ctxtRV50A-PO_MATNR").text = pn
            ' enter
            .FindById("wnd[1]").SendVKey 0
            ' where it lands
            ' get text
            sap_val = CInt(.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPSD-G_LFIMG[2,0]").text)
            If curr_val = sap_val Then
                ' check if changeable
                If Not .FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPSD-G_LFIMG[2,0]").changeable Then
                    dbLog.log "PN (" & pn & ") not changeable."
                    ' yellow arrow back
                    .FindById("wnd[0]/tbar[0]/btn[15]").press
                    GoTo skip
                Else
                    .FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPSD-G_LFIMG[2,0]").text = set_val
                    rng.Cells(r.row, 1).Font.Bold = True
                    ' increment counter
                    counter = counter + 1
                End If
            Else
                ' no match
                dbLog.log "current sap value (" & sap_val & ") doesn't match with sheet value (" & curr_val & ") .... skipping", msgPopup:=True, msgType:=vbCritical
                rng.Cells(r.row, 1).Font.Italic = True
                GoTo skip
            End If
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]").SendVKey 0
            ' shift+f6
            .FindById("wnd[0]").SendVKey 18
            ' ctrl+f3
            .FindById("wnd[0]").SendVKey 27
            ' enter for window popup
            If Exist_Ctrl(sess, 1, "", True).cband Then
                .FindById("wnd[1]").SendVKey 0
            End If
            ' save
            .FindById("wnd[0]").SendVKey 11
        End With

skip:
Next r
dbLog.log "Finished updating (" & counter & ") delivery qtys.", msgPopup:=1, msgType:=vbInformation
Exit Sub
errlv:

End Sub
