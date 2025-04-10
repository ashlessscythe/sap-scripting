Function run*tcode_n_variant(ByVal sess, ByVal rName As String, ByVal tCode, ByVal vars_in, Optional ByVal arg, *
Optional ByVal sExportAction As String = "") As params
Dim temp_ws, rVal As New BetterArray, report_name As String
Dim err_ctl As ctrl_check, bar_msg As String, ctrl_msg As String, err_msg As String, run_check As Boolean
Dim run_vars, sheet_name, column_name, msg

    ' for return values
    Dim params_out As params
    params_out.name = rName

    Application.StatusBar = "Running SAP tCode (" & tCode & ")"

    ' parse opt
    Dim opt
    opt = get_json_opt(arg)

    Dim bSkipEmptyCheck As Boolean
    bSkipEmptyCheck = False
    If Contains(opt, "skip_empty_check") Then bSkipEmptyCheck = True

    On Error GoTo errlv
    With sess
        If LCase(Left(vars_in, "3")) = "use" Then
            run_vars = Split(vars_in, " ")
            sheet_name = run_vars(1)
            column_name = run_vars(2)
        End If

        ' get working date
        If IsMissing(arg) Then GoTo noLayout

        Select Case True
        Case arg.Exists("date")
            Dim datevar As dateObj
            ' datevar = get_arg_date(arg)
            datevar = get_json_date(arg("date"))
        Case Contains(vars_in, "date:")
            datevar = get_arg_date(vars_in)
        End Select

noLayout:
Dim rSheet As String
If Len(sheet_name) > 1 Then
Select Case True
Case SheetExists(tWB, recent_sheet(tWB, sheet_name))
rSheet = recent_sheet(tWB, sheet_name)
Case Else
Dim hdr As New BetterArray
hdr.Push column_name
Evaluate create_or_get_sheet(tWB, sheet_name, hdr)
tWB.Sheets(sheet_name).Activate
params_out.err = "no_sheet"
GoTo out_no_ws
End Select
End If
Select Case LCase(tCode)
Case "vl06o"
run_check = VL06O_Delivery_Check(tWB, "", "", "", sess, tCode, "", rSheet, column_name, arg:=arg)

            If run_check Then
                Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                                rName & "_" & tCode, ""))
            Else
                GoTo errlv
            End If
        Case "zleqty"
            If arg.Exists("filter") Then
                Dim f_header As String, f_value As String
                f_header = "": f_value = ""
                f_header = arg("filter")
                f_value = arg("filter")
            End If

            Dim p As New BetterArray: p.FromExcelRange get_rng(shLookup, arg("layout"), 1, 0, toLastRow:=True)
            Dim x As Variant, oWS As Worksheet, wsName As String
            ' Loop plants in shLookup sheet
            If f_header <> "" Then p.Clear: p.Push f_value
            For Each x In p.items
                ' check if already exported
                If (SheetExists(tWB, recent_sheet(tWB, x))) Then GoTo skip_x
                wsName = zleqty_by_plant(tWB, tCode, tCode, sess, "E01", x, "")
                If Len(wsName) > 0 And LCase(wsName) <> "false" Then
                    Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                                    wsName, ""))
                    If Not temp_ws Is Nothing Then
                        ' check if no data
                        If sheet_is_empty(tWB, temp_ws.name, , 5) Then GoTo skip_x
                        Call ConValue(temp_ws, "value")
                        Call Clear_AutoFilter(temp_ws, True)
                        Call FormatSheetColumnsDateTime(tWB, temp_ws.name)
                        ' wsExports.Push temp_ws
                    Else
                    End If
                Else
                End If

skip_x:
Next x

        Case "lt23"
            ' default today's date
            If Len(datevar.date1) < 1 Then datevar.date1 = TODAYS_DATE
            Dim radial_val As String
            If arg.Exists("radial") Then
                radial_val = arg("radial")
            Else
                radial_val = "confirmed"    ' default
            End If

            ' if 2nd date not provided, make equal to 1
            If Len(datevar.date2) < 1 Then datevar.date2 = datevar.date1
            If Len(sheet_name) > 0 And Len(column_name) > 0 Then
                params_out = LT23_Run_Export(tWB, "E01", sess, tCode, arg("layout"), rSheet, column_name, radial_val)
            Else
                params_out = LT23_date(tWB, "E01", datevar.date1, datevar.date2, _
                                   sess, tCode, vars_in, arg("layout:"), radial_val)
            End If
            ' Save
            If params_out.run_check Then Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                                                         rName & "_" & tCode, "lt23_by_to"))
            Debug.Print
        Case "zmdesnr_exclude"
             tCode = "zmdesnr"
            Dim ex_sheet, ex_col, s_args
            Set s_args = JsonConverter.ParseJson(sExportAction)
            ex_sheet = s_args("exclude_sheet")
            ex_col = s_args("exclude_col")

            params_out = ZMDESNR_With_Exclude_Export(tWB, "", tCode, arg("var"), sess, _
                                                     arg("layout"), rSheet, column_name, ex_sheet, ex_col)

             If Not params_out.run_check Then params_out.name = rName: GoTo out_no_ws
            ' Save
             Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            rName & "_sn_excl", "sn_excl"))
            Debug.Print
        Case "zmdesnr_hist"
            tCode = "zmdesnr"
            params_out = ZMDESNR_Run_Export(tWB, tCode, "", rSheet, column_name, parse_arg(arg, "var:", 0, , 1), sess, _
                                            parse_arg(arg, "layout:", 0, , 1), "history", True)

            If Not params_out.run_check Then params_out.name = rName: GoTo out_no_ws
            ' Save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            rName & "_sn_hist", "sn_hist"))
            Debug.Print
        Case "zmdesnr_to"
            tCode = "zmdesnr"
            If Len(rSheet) < 1 Then
                params_out.err = "no sheet"
                GoTo out_no_ws
            Else
                Dim dataWS As Worksheet: Set dataWS = tWB.Sheets(rSheet)
                If Contains(column_name, "multi") Then
                    Dim del_col As String, to_col As String
                    del_col = arg("del_col")
                    to_col = arg("to_col")
                End If
                params_out = sn_from_to_run(sess, rName, tCode, vars_in, arg, dataWS, del_col, to_col)
            End If
            If Len(params_out.err) > 1 Then
                msg = params_out.err
                GoTo out_no_ws
            End If
            ' check layout
            ' params_out = check_select_layout(sess, tCode, parse_arg(arg, "layout:", 0, , 1), arg, 1)

            If Not params_out.run_check Then params_out.name = rName: GoTo out_no_ws
            ' Save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            rName & "_sn_to", "sn_from_to"))
            ' filters
            params_out.run_check = Filter_Copy_MG(tWB, temp_ws.name, rName & "_filtered", arg)
            Dim bDupSheet As Boolean: If Contains(arg("option"), "dupsheet") Then bDupSheet = True
            ' duplicate for backup
            If bDupSheet Then Evaluate duplicatesheet(temp_ws, "to_orig")

        Case "zmdesnr_del"
            tCode = "zmdesnr"
            If Len(rSheet) < 1 Then
                params_out.err = "no sheet"
                params_out = sn_from_del_run(sess, rName, tCode, vars_in, arg, _
                                             bSkipEmptyCheck:=bSkipEmptyCheck)
            Else
                params_out = sn_from_del_run(sess, rName, tCode, vars_in, arg, _
                                             tWB.Sheets(rSheet), column_name, bSkipEmptyCheck)
            End If
            If Len(params_out.err) > 1 Then
                msg = params_out.err
                GoTo out_no_ws
            End If
            ' check layout
            params_out = check_select_layout(sess, tCode, arg("layout"), arg, 1)

            If Not params_out.run_check Then params_out.name = rName: GoTo out_no_ws
            ' Save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            rName & "_sn_del", "sn_from_del", True))
            ' duplicate for backup
            If Not temp_ws Is Nothing Then Evaluate duplicatesheet(temp_ws, "snr_orig") Else GoTo out_no_ws

        Case "zmdesnr_bin"
            tCode = "zmdesnr"
            If Len(rSheet) < 1 Then
                params_out.err = "no sheet"
                params_out = sn_from_bin_run(sess, rName, tCode, vars_in, arg)
            Else
                params_out = sn_from_bin_run(sess, rName, tCode, vars_in, arg, tWB.Sheets(rSheet), column_name)
            End If
            If Len(params_out.err) > 1 Then
                msg = params_out.err
                GoTo out_no_ws
            End If
            ' check layout
            params_out = check_select_layout(sess, tCode, arg("layout"), arg, 1)

            If Not params_out.run_check Then params_out.name = rName: GoTo out_no_ws
            ' Save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            rName & "_sn_bin", "sn_from_bin"))
            Debug.Print
        Case "zmdesnr_sn"
            tCode = "zmdesnr"
            If Len(rSheet) < 1 Then params_out.err = "no sheet": GoTo out_no_ws
            params_out = sn_from_pn_run(sess, tCode, tWB.Sheets(rSheet), column_name, arg)
            If Len(params_out.err) > 1 Then
                msg = params_out.err
                GoTo out_no_ws
            End If
            ' check layout
            params_out = check_select_layout(sess, tCode, arg("layout"), arg, 1)

            If Not params_out.run_check Then params_out.name = rName: GoTo out_no_ws
            ' Save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            rName & "_sn", "sn_from_pn"))
            Debug.Print
        Case "zmdesnr_pn"
            tCode = "zmdesnr"
            If Len(rSheet) < 1 Then params_out.err = "no sheet": GoTo out_no_ws
            Evaluate pn_from_sn_run(sess, tCode, tWB.Sheets(rSheet), column_name, arg)

            ' check layout
            params_out = check_select_layout(sess, tCode, parse_arg(arg, "layout:", 0, , 1), arg, 1)

            ' Save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            "doh_pn", "pn_from_sn"))
        Case "zmdetpc"
            ' wk in arg
            Select Case datevar.date_type
            Case dt_multiple, dt_range, dt_diff
                Call ZMDETPC_PalletCount_Run_Export(tWB, DateAdd("d", -3, datevar.date1), DateAdd("d", 3, datevar.date2), _
                                                    "EPDC_TOM", tCode, sess, "", "summarized", _
                                                    rSheet, column_name)
            Case dt_single, dt_default
                ' date provided or date default
                Call ZMDETPC_PalletCount_Run_Export(tWB, DateAdd("d", -4, datevar.date1), _
                                                    DateAdd("d", 3, datevar.date1), _
                                                    "EPDC_TOM", tCode, sess, "", "summarized", _
                                                    rSheet, column_name)

            Case Else
                Call ZMDETPC_PalletCount_Run_Export(tWB, DateAdd("d", -4, TODAYS_DATE), _
                                                    DateAdd("d", 3, TODAYS_DATE), _
                                                    "EPDC_TOM", tCode, sess, "", "summarized", _
                                                    rSheet, column_name)
            End Select
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            tCode & "_" & Format(datevar.date1, EXTERNAL_DATE_FORMAT), ""))
        Case "vt03n"
            Select Case LCase(rName)
                Case "add_del"
                    Set temp_ws = vt03n_get_info(tWB, sess, tCode, _
                                    recent_sheet(tWB, sheet_name), column_name, "pallets")
                Case "loop_print"
                    Set temp_ws = vt03n_print(tWB, sess, tCode, arg("sheet"), arg("column"), arg)
            Case Else
                ' check arg passed in
                If arg.Exists("option") Then
                Dim out_col: out_col = "pallets"
                ' set out col
                If arg.Exists("out_col") And Len(arg("out_col")) > 0 Then out_col = arg("out_col")
                    Select Case arg("option")
                        Case "get_info"
                            Set temp_ws = vt03n_get_info(tWB, sess, tCode, _
                                    recent_sheet(tWB, sheet_name), column_name, out_col)
                        Case "loop"
                            Set temp_ws = vt03n_print(tWB, sess, tCode, arg("sheet"), arg("column"), arg)
                    End Select
                End If
            End Select
        Case "vt22"
            Call vt22_changes(tWB, sess, tCode, arg, recent_sheet(tWB, sheet_name), column_name)

            If arg.Exists("sheet_out") Then
                Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".txt"), _
                                         arg("sheet_out") & "_" & Format(datevar.date1, EXTERNAL_DATE_FORMAT), ""))
                ' Cleanup text file
                Evaluate format_vt22(tWB.Sheets(recent_sheet(tWB, arg("sheet_out"))), rName, arg)
            Else
                Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".txt"), _
                                         tCode & "_" & Format(datevar.date1, EXTERNAL_DATE_FORMAT), ""))
                ' Cleanup text file
                Evaluate format_vt22(tWB.Sheets(recent_sheet(tWB, "vt22")), rName, arg)
            End If

            Debug.Print
        Case "vt11"
            Call VT11_Run_Export(tWB, datevar.date1, datevar.date2, tCode, arg("var"), _
                                 arg("layout"), sess, _
                                 CStr(arg("field")), rSheet, column_name, False)
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                                rName & "_" & tCode, ""))
            Case Else
        End Select

        If IsEmpty(temp_ws) Or temp_ws Is Nothing Then
            Set temp_ws = Nothing
            msg = "no export for tcode: (" & tCode & ")": GoTo errlv
        End If
        Call ConValue(temp_ws)
        Evaluate Clear_AutoFilter(temp_ws, 1)
        Evaluate sanitize_headers(temp_ws.Parent, temp_ws.name)
        Evaluate FormatSheetColumnsDateTime(temp_ws.Parent, temp_ws.name)
    End With

    ' Check for export action
    If sExportAction <> "" Then
        Dim ex: ex = Split(sExportAction, " ")
        Dim col_name As String, col_1 As String, col_2 As String, sh_name
        Dim bCheck As Boolean
        Select Case True
        Case Contains(ex(0), "mustgo")
            ' remove dups
            Evaluate RemoveDuplicates(temp_ws, "transfer_order_number")
            ' insert datetime
            Evaluate insert_column(temp_ws.Parent, temp_ws.name, "datetime", FindColumn(temp_ws, "creation_time", 1, 0) + 1)
            Call morph_datetime(temp_ws.Parent, temp_ws, FindColumn(temp_ws, "creation_date", 1, 0), _
                                FindColumn(temp_ws, "creation_time", 1, 0), FindColumn(temp_ws, "datetime", 1, 0))
            ' sort
            Call sort_multi_fields(temp_ws, Col_Letter(FindColumn(temp_ws, "delivery", 1, 0)), xlDescending, _
                                   Col_Letter(FindColumn(temp_ws, "datetime", 1, 0)), xlDescending, _
                                   Col_Letter(FindColumn(temp_ws, "source_storage_bin", 1, 0)), xlDescending)

            ' insert formula
            Call insert_column(temp_ws.Parent, temp_ws.name, "bDup", FindColumn(temp_ws, "delivery", 1, 0))
            Dim fml As String, coldiff As Double
            coldiff = FindColumn(temp_ws, "delivery", 1, 0) - FindColumn(temp_ws, "bDup", 1, 0)
            fml = "=IF(COUNTIF(C[" & coldiff & "],RC[" & coldiff & "])>1,COUNTIF(C[" & coldiff & _
                  "],RC[" & coldiff & "])," & Chr(34) & "unique" & Chr(34) & ")"
            Call Insert_Formula(temp_ws.Parent, temp_ws.name, "bDup", fml)

            Call clear_formats(temp_ws)
            Call delete_rows_with_criteria(temp_ws.Parent, temp_ws.name, "unique", FindColumn(temp_ws, "bDup", 1, 0))

            ' remove one instance from dups
            Evaluate loop_filter_col_action(temp_ws, "delivery", xDelete)

            ' new book
            ' Set temp_ws = save_sheet_new_book(temp_ws, rName_folder(rname), rname, tCode, get_arg_date(vars_in))

            ' pivot
            ' Dim pivSh As Worksheet
            ' Set pivSh = CreateClearSheet(temp_ws.Parent, "pivot")

            ' Call createPivot(temp_ws, pivSh, "A2", "piv_", "", "transfer_order_number", , , , , _
                             "creation_date", "plant")

            Debug.Print
        Case Contains(ex(0), "multi_duplicates")
            sh_name = temp_ws.name
            col_1 = parse_arg(sExportAction, "dup_col_1:", 0, , 1)
            col_2 = parse_arg(sExportAction, "dup_col_2:", 0, , 1)
            ' sort qty largest on top
            ' sort pn ascending
            If Not sort_multi_fields(temp_ws, Col_Letter(FindColumn(temp_ws, col_1, 1, 0)), xlAscending, _
                                     Col_Letter(FindColumn(temp_ws, col_2, 1, 0)), xlDescending) Then
                params_out.err = "sort error"
            End If
            If Not RemoveDuplicates_Mult_2(temp_ws.Parent, temp_ws.name, col_1, col_2) Then
                params_out.err = "remove_dups_error"
            End If
            Debug.Print
        Case Contains(ex(0), "remove_dups")
            sh_name = recent_sheet(tWB, tCode)
            col_name = parse_arg(ex(0), "remove_dups:", 0)
            Evaluate RemoveDuplicates(tWB.Sheets(sh_name), col_name)
            Debug.Print
        Case Contains(ex(0), "filter")
            Dim filter_header, filter_value, bDeleteHidden As Boolean
            ' Very limited, currently only supports 1 filter_value, delim by "|"
            filter_header = parse_arg(ex(0), "filter:", 0)
            filter_value = parse_arg(ex(0), "filter:", 1)
            ' Filter
            Call FilterColumn(temp_ws.Parent, temp_ws.name, FindColumn(temp_ws, filter_header, 1, 0), xlFilterValues, filter_value)
            If sheet_is_empty(temp_ws.Parent, temp_ws.name, 1) Then
                dbLog.log "Sheet for tCode (" & tCode & ") has no results for in column (" & filter_header & _
                          ") for item (" & filter_value & ")", msgPopup:=True, msgType:=vbExclamation
                GoTo out
            End If
            ' delete hidden???
            If UBound(ex) >= 1 Then
                If Contains(ex(1), "delete") Then
                    Call deleteHidden(temp_ws.Parent, temp_ws.name)
                End If
            End If
            Debug.Print
        Case Contains(ex(0), "merge")
            '  Stuff
            Dim sWS As Worksheet, dWS As Worksheet
            Set sWS = temp_ws
            Select Case LCase(tCode)
            Case "zmdetpc"
                ' Change header before merge
                sWS.usedRange.Find("delivery").value = "delivery_"
                Set dWS = tWB.Sheets(recent_sheet(tWB, ex(1)))
                Call remove_dupes(dWS, "delivery", bPartialMatch:=True)
                ' Sort
                Call SortCol(dWS, "delivery", bPartialMatch:=True)
                Call SortCol(sWS, "delivery_", bPartialMatch:=True)
                Dim s As boundsObj: s = Bounds(sWS)
                Dim d As boundsObj: d = Bounds(dWS)
                ' copy/paste
                s.rng(1, 1).CurrentRegion.Copy Destination:=d.rng(d.fRow, d.lCol + 1)
                dWS.Activate
                ' formula
                Evaluate Insert_Formula(tWB, dWS.name, "total", "=SUM(RC[-2]:RC[-1])")
                Evaluate insert_spacers(dWS, "delivery", "delivery_", 4)

                ' If rName xdock, move num_of_packages to blanks
                If Contains(rName, "xdock") Or Contains(rName, "toyota") Then
                    Evaluate copy_val_col_filter(dWS, "number_of_packages", "total", "", "=")
                    Evaluate copy_val_col_filter(dWS, "delivery", "delivery_", "", "=")
                    ' replace 0 with 1
                    dWS.Columns(FindColumn(dWS, "total")).Replace "0", "1"
                ElseIf Contains(rName, "add_del") Then
                    ' replace 0 with 1
                    Dim r As Range, c As Range
                    d = Bounds(dWS)
                    Set r = d.rng.Columns(d.headers("pallets"))
                    Debug.Print r.address
                    With r
                        For Each c In .Cells
                            If Not IsNumeric(c.value) And c.value <> "pallets" Or IsEmpty(c.value) Then
                                c.value = "1": c.Interior.ColorIndex = 8
                            End If
                        Next c
                    End With
                End If

                ' modified 230410-
                ' moved to tcode save with data_sheet and pivot args
                ' New book
                ' Dim xl2 As Workbook, ws2 As Worksheet, shPiv As Worksheet
                ' Set xl2 = Workbooks.add
                ' dWS.Copy After:=xl2.Sheets(1)
                ' Set ws2 = xl2.Sheets(dWS.name)
                ' Call Force_Del_Sheet(xl2, Sheets(1).name)
                ' Pivot
                ' Set shPiv = CreateClearSheet(xl2, "pivot")
                ' Select Case True
                    ' default is pallets, rma, etc...
                ' Case Contains(rName, "add_del")
                '     Call createPivot(ws2, shPiv, "A1", "pivot1", "pallets", "", rowItem1:="date")
                ' Case Contains(rName, "k2xx")
                '     Call createPivot(ws2, shPiv, "A1", "pivot1", "total", "", rowItem1:="name_of_the_shipto_party", _
                                     colItem1:="plant")
                ' Case Contains(rName, "misc")
                '     Call createPivot(ws2, shPiv, "A1", "pivot1", "number_of_packages", "", rowItem1:="plant")
                ' Case Mult_Contains(rName, "xdock", "toyota")
                '     Call createPivot(ws2, shPiv, "A1", "pivot1", "total", "", rowItem1:="plant")
                ' Case Contains(rName, "tesla")
                '     Call createPivot(ws2, shPiv, "A1", "piv1", "total", "", rowItem1:="date")
                ' Case Contains(rName, "rma")
                '     Call createPivot(ws2, shPiv, "A1", "pivot1", "total", "", _
                                     rowItem1:="date", rowItem2:="name_of_the_shipto_party", showTabularRow:=1)
                ' Case Contains(rName, "addon")
                '     Call createPivot(ws2, shPiv, "A1", "pivot1", "total", "", _
                                     rowItem1:="date", colItem1:="plant")
                ' Case Else
                '     Call createPivot(ws2, shPiv, "A1", "pivot1", "total", "", _
                                     rowItem1:="plant", rowItem2:="name_of_the_shipto_party", showTabularRow:=1)
                    ' msg = "Check report name for this pivot creation, rName =  (" & rName & "), tcode = (" & tCode & ")"
                    ' GoTo errlv
                ' End Select
                ' Save
                ' xl2.SaveAs get_save_name(rName_folder(rName), rName, tCode, datevar)
                ' Set returnObj
                Evaluate LetOrSetElement(temp_ws, dWS)
            Case "zleqty"
                ' Merge exports
                Dim vN, wsExports As New BetterArray
                wsExports.Clear
                For Each vN In tWB.Sheets
                    If Contains(vN.name, tCode) Then wsExports.Push vN
                Next vN
                Set dWS = better_merge_sheets(tWB, "item_master", wsExports, 1)
                Call ConValue(dWS, "value")
                Call Clear_AutoFilter(dWS, True)
                Call FormatSheetColumnsDateTime(tWB, dWS.name)
                Evaluate remove_dups_mult_cols(dWS)
                Set temp_ws = save_sheet_new_book(dWS, rName_folder(rName), rName, tCode, get_arg_date("date:diff+0"))
                Debug.Print
            Case Else                            ' tcode
            End Select
        Case Else

        End Select                               ' true ex()
    End If

out:
Application.StatusBar = False
If temp*ws Is Nothing Then GoTo out_no_ws
' Use rName if provided, else keep temp_ws.name
If Len(rName) > 0 And Not temp_ws Is Nothing And Not Contains(temp_ws.name, rName) Then temp_ws.name = UCase(rName) & "*" & temp_ws.name
Set params_out.obj = temp_ws
params_out.type = "sheet"
out_no_ws:
run_tcode_n_variant = params_out
Exit Function
errlv:

    msg = "Error: " & msg & " in run_tcode_n_variant"

    dbLog.log msg, msgPopup:=rMsgSwitch.value, msgType:=vbInformation
    params_out.err = msg

    GoTo out

End Function

Function run*tcode_w_variant(ByVal sess, ByVal rName As String, ByVal tCode, ByVal sap_var, Optional ByVal arg, *
Optional ByVal action As String = "", Optional ByVal rObj As Variant) As params
Dim temp*ws, report_name As String
Dim err_ctl As ctrl_check, err_wnd As ctrl_check, bar_msg As String, ctrl_msg As String, *
msg As String, err_msg As String, run_check As Boolean

    Application.StatusBar = "Running SAP tCode (" & tCode & ")"

    ' for return values
    Dim params_out As params
    params_out.name = rName
    If Not IsMissing(rObj) And Not Contains(TypeName(rObj), "string") Then Set params_out.obj = rObj


    ' get date from arg
    Dim datevar As dateObj
    ' datevar = get_arg_date(arg("date"))
    If arg.Exists("date") Then
        datevar = get_json_date(arg("date"))
    End If

    ' check if arg:monday_include_sunday provided in action
    If Contains(action, "arg:monday_include_sunday") Then
        ' check if monday
        If Contains(Format(datevar.date1, "dddd"), "monday") Then
            ' extend back 1 day
            datevar = get_arg_date("date:diff-1")
        End If
    End If

    With sess
        Evaluate assert_tcode(sess, tCode)
        If Contains(sap_var, ", ") Then
            ' loop variant, call self
            Dim arr As New BetterArray, v, vN: vN = Split(sap_var, ",")
            Dim export_count As Integer: export_count = 0
            For v = 0 To UBound(vN)
                If arg.Exists("multi") Then arg = "multi (" & v & ") of (" & UBound(vN) & ")"
                params_out = run_tcode_w_variant(sess, rName, tCode, Trim(vN(v)), arg, action)
                If Len(params_out.err) < 1 Then
                    export_count = export_count + 1
                    arr.Push params_out.obj
                End If
            Next v
            If export_count > 0 Then
                ' clear error
                params_out.err = ""
                params_out.type = "multi_export"
                params_out.value = export_count
                Set temp_ws = better_merge_sheets(tWB, tCode & "_mult", arr, 1)
                Debug.Print
                GoTo check
            Else
                ' no exports, exit gracefully? (try to, anyway)
                dbLog.log "No Negatives found for variants (" & sap_var & ") tCode (" & tCode & ")", _
                          msgPopup:=True, msgType:=vbInformation
            End If
            GoTo out
        ElseIf Contains(sap_var, "use ") Then
            params_out = run_tcode_n_variant(sess, rName, tCode, sap_var, arg, action)
            If Not Len(params_out.err) > 1 Then Debug.Print params_out.obj.name
            ' return sheet
            If IsEmpty(params_out.obj) Then _
                                       Set temp_ws = tWB.Sheets(recent_sheet(tWB, tCode)) Else _
                                       Set temp_ws = params_out.obj
                                       params_out.type = "sheet"
            GoTo out
        Else
            Evaluate choose_variant(sess, tCode, sap_var)
        End If

        ' save name, if provided
        Dim out_sheet_name As String
        out_sheet_name = get_if_arg(arg, "out_sheet_name", tCode)

        Select Case LCase(tCode)
        Case "zforjit"
            params_out = zforjit_export(tWB, sess, tCode, arg)
            If Len(params_out.err) > 1 And Contains(params_out.err, "no data") Then
                dbLog.log "No data found for report (" & rName & ") tCode (" & tCode & ") using variant (" & sap_var & ")", _
                          msgPopup:=rMsgSwitch.value, msgType:=vbInformation
                GoTo out
            End If

            ' Save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            out_sheet_name, sap_var, bDeleteBlankHeader:=False))

            ' false delete blank header from CopyPaste() to prevent err with number cols
            Debug.Print
        Case "mb51"
            params_out = mb51_with_variant(sess, tCode, arg)
            If Len(params_out.err) > 1 And Contains(params_out.err, "no data") Then
                dbLog.log "No data found for report (" & rName & ") tCode (" & tCode & ") using variant (" & sap_var & ")", _
                          msgPopup:=rMsgSwitch.value, msgType:=vbInformation
                GoTo out
            End If
            ' Save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            tCode, sap_var))
        Case "mb52"
            params_out = mb52_w_variant(sess, tCode, arg)
            If Len(params_out.err) > 1 And Contains(params_out.err, "no data") Then
                dbLog.log "No data found for report (" & rName & ") tCode (" & tCode & ") using variant (" & sap_var & ")", _
                          msgPopup:=rMsgSwitch.value, msgType:=vbInformation
                GoTo out
            End If
            ' Save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            lValue & "_" & tCode, sap_var))
            Debug.Print
        Case "y_dn3_47000149"
            Select Case LCase(rName)
            Case "mg_snapshot"
                params_out = y_149_by_limiter(sess, tCode, arg, SAPVariantName:=sap_var)
                ' Save
                Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".txt"), _
                                                                tCode & "_" & Format(datevar.date1, EXTERNAL_DATE_FORMAT), ""))
            Case Else
                If IsMissing(rObj) Then
                    params_out = y_149_by_limiter(sess, tCode, arg)
                Else
                    params_out = y_149_by_limiter(sess, tCode, arg, obj:=rObj)
                End If

                ' Save
                If params_out.run_check Then
                    ' get out_sheet_name if provided
                    Dim out_sheet_prefix As String
                    out_sheet_prefix = get_if_arg(arg, "bool_out_sheet_rename")

                    If out_sheet_prefix Then
                        Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".txt"), _
                                                 Right(arg("sheet"), 4) & "_" & tCode, ""))
                    Else
                        Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".txt"), _
                                                                tCode & "_" & Format(datevar.date1, EXTERNAL_DATE_FORMAT), ""))

                    End If
                Else
                    Set temp_ws = Nothing
                    If params_out.err = "" Then
                        params_out.err = "no_ws"
                    End If
                End If

                ' dbLog.log "rName (" & rName & ") not implemented in run_tcode_w_variant() for tCode (" & tCode & ")", _
                msgPopup:=True, msgType:=vbExclamation
            End Select

        Case "zmdesnr"
            Select Case True
            Case Contains(rName, "redtag") Or Contains(rName, "uactivity") Or Contains(rName, "knock-offs") Or Contains(rName, "nc_master")
                ' User Activity Tab
                .FindById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM7").Select
                ' Multi status button
                ' .findById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM7/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9007/btn%_S_STAT1_%_APP_%-VALU_PUSH").press
                ' Clear
                ' .findById("wnd[1]").sendVKey 16
                ' Status
                ' .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "ash"
                ' Close popup
                ' .findById("wnd[1]").sendVKey 8
                ' Multi SLOC
                ' .findById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM7/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9007/btn%_S_LGORT1_%_APP_%-VALU_PUSH").press
                ' Clear
                ' .findById("wnd[1]").sendVKey 16
                ' Enter Multiple
                ' .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").Text = "0003"
                ' .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").Text = "0010"
                ' Close
                ' .findById("wnd[1]/tbar[0]/btn[8]").press

                ' Display report checkmark
                .FindById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM7/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9007/chkP_ALV").Selected = True
                .FindById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM7/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9007/ctxtS_ERDAT1-LOW").text = datevar.date1
                .FindById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM7/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9007/ctxtS_ERDAT1-HIGH").text = datevar.date2
                ' User Activity Information
                ' .findById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM7/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9007/radP_USER").Select
                ' User Activity report
                .FindById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM7/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9007/radP_PENSKE").Select

                ' Execute
                .FindById("wnd[0]/tbar[1]/btn[8]").press
                ' Layout setup, if can't load,
                ' no layout should load in case pre_export_back is passed to check_select_layout via arg
                ' .findById("wnd[0]/tbar[1]/btn[32]").press
            Case Contains(rName, "busbar")
                .FindById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2").Select
                .FindById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9002/ctxtS_ERDAT-LOW").text = datevar.date1
                .FindById("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2/ssub%_SUBSCREEN_TABB1:ZMDE_SERIALNUMBER_HISTORY:9002/ctxtS_ERDAT-HIGH").text = datevar.date2
                ' Execute
                .FindById("wnd[0]/tbar[1]/btn[8]").press
            Case Else
            End Select


            ' check layout
            params_out = check_select_layout(sess, tCode, arg("layout"), arg, 1)

            ' Save
            If Not Len(params_out.err) = 0 Then
                Evaluate LetOrSetElement(temp_ws, Nothing)
                GoTo check
            Else
                Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            tCode & "_", sap_var))
            End If
        Case "zvt11"
            Dim sheet_name As String, column_name As String, run_vars
            If arg.Exists("sheet") And arg.Exists("column") Then
                sheet_name = arg("sheet")
                column_name = arg("column")
            End If
            Dim zvtType As String
            zvtType = "closed"
            If Contains(rName, "mg_") Then zvtType = "open"

            ' Run
            Evaluate ZVT11_From_Deliv(tWB, sess, tCode, arg("layout"), _
                                      zvtType, recent_sheet(tWB, sheet_name), _
                                      column_name, sapvar:=sap_var, dateObj:=arg("date"))

            ' Save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            rName & "_" & tCode, sap_var))

            Debug.Print
        Case "vt11"
            ' high button
            ' .findById("wnd[0]/usr/ctxtK_STTRG-HIGH").Text = "7"
            ' date
            Select Case datevar.date_type
            Case dt_single, dt_default
                ' If not visible, then variant is hiding
                If Exist_Ctrl(sess, 0, "/usr/ctxtK_DATEN-LOW", 1).cband Then
                    ' Date actual end
                    .FindById("wnd[0]/usr/ctxtK_DATEN-LOW").text = datevar.date1
                End If
            Case dt_multiple, dt_range, dt_diff
                Select Case True
                Case Contains(datevar.date_str, "created")
                    .FindById("wnd[0]/usr/ctxtK_ERDAT-LOW").text = datevar.date1
                    .FindById("wnd[0]/usr/ctxtK_ERDAT-HIGH").text = datevar.date2
                Case Else
                    .FindById("wnd[0]/usr/ctxtK_DATEN-LOW").text = datevar.date1
                    .FindById("wnd[0]/usr/ctxtK_DATEN-HIGH").text = datevar.date2
                End Select
            Case dt_none
                ' none....
                dbLog.log "No date provided"
            End Select
            ' execute
            .FindById("wnd[0]/tbar[1]/btn[8]").press
            ' check for err in sBar
            Dim sBar As String: sBar = Hit_Ctrl(sess, 0, "/sbar", "Text", "Get", "")
            If Contains(sBar, "E:") Then
                msg = sBar
                params_out.err = sBar
                GoTo out
            End If
            ' check for error popup
            err_wnd = Exist_Ctrl(sess, 1, "", True)
            If err_wnd.cband Then
                If Contains(err_wnd.ctext, "information") Then
                    ' No shipments found?
                    msg = get_sap_text_errors(sess, 1, "/usr/txtMESSTXT1", 10)
                    If Contains(msg, "no shipments") Then
                        params_out.err = "no results"
                    End If
                    GoTo out
                End If
            End If
            ' check for err msg
            err_msg = Hit_Ctrl(sess, 0, "/sbar", "Text", "Get", 1)
            Debug.Print err_msg
            If Len(err_msg) > 0 Then
                dbLog.log err_msg, msgPopup:=rMsgSwitch.value, msgType:=vbCritical
                params_out.err = err_msg
                GoTo out
            End If

            ' select layout
            params_out = check_select_layout(sess, tCode, arg("layout"), arg, 1)


            ' if results save
            If params_out.run_check Then
                If IsNull(datevar.date1) Or IsEmpty(datevar.date1) Or datevar.date_type = dt_none Then
                    Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                                    tCode & "_" & datevar.date_str, sap_var))
                Else
                    Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                                    tCode & "_" & Format(datevar.date1, EXTERNAL_DATE_FORMAT), sap_var))
                End If
            Else
                Set temp_ws = Nothing
            End If
        Case "lx03", "lx02"
            ' execute
            .FindById("wnd[0]/tbar[1]/btn[8]").press
            ' select layout
            params_out = check_select_layout(sess, tCode, arg("layout"), arg, 1)

            ' save
            Evaluate LetOrSetElement(temp_ws, save_sap_file(sess, wPath, CStr(lValue & "_" & tCode & ".xlsx"), _
                                                            out_sheet_name & "_" & Format(Now, EXTERNAL_DATE_FORMAT), sap_var))

        Case "lx04"
            ' Execute
            .FindById("wnd[0]/tbar[1]/btn[8]").press

            ' Get Data
            Dim sap_arr As New BetterArray
            Set temp_ws = CreateClearSheet(tWB, sap_var)
            Evaluate LetOrSetElement(sap_arr, get_info_lbl(sess, 0, "/usr/lbl", 1, 6))
            sap_arr.ToExcelRange temp_ws.Cells(1, 1)
            temp_ws.Columns.AutoFit
        Case Else
        End Select

check:
If temp_ws Is Nothing Then GoTo out
' TODO: prevent blank sheets or sheets with only headers in col 1 from being processed
If sheet_is_empty(temp_ws.Parent, temp_ws.name, min_cols:=2) Then
params_out.err = "empty_sheet"
GoTo out
End If

        Call ConValue(temp_ws)
        Evaluate Clear_AutoFilter(temp_ws, 1)
        Evaluate sanitize_headers(temp_ws.Parent, temp_ws.name)
        Evaluate FormatSheetColumnsDateTime(temp_ws.Parent, temp_ws.name)

        ' Check for additional action
        If action = "" And arg.Exists("option") Then
            action = arg("option")  ' only 1, fix for mult
        End If
        If action = "" Then GoTo out

        Application.StatusBar = "Performing actions on sheet..."
        ' If action, use out_ws
        Dim out_ws As Worksheet
        Set out_ws = temp_ws
        Dim sType As String, header As String, criteria As Range
        Dim nWB As Workbook, nWS As Worksheet, vals As New Collection
        Select Case True
        Case Contains(action, "combine")
            Set out_ws = Nothing
            Set params_out.obj = temp_ws
            params_out.type = "sheet"
            params_out.name = temp_ws.name
        Case Contains(action, "merge")
            Dim merge_sheet_name As String
            merge_sheet_name = parse_arg(action, "sheet:", 0, , 1)
            Dim wsExports As New BetterArray
            wsExports.Clear
            For Each vN In tWB.Sheets
                If Contains(vN.name, merge_sheet_name) Then wsExports.Push vN
            Next vN
            Set out_ws = better_merge_sheets(tWB, merge_sheet_name, wsExports, 1)
            Evaluate save_sheet_new_book(out_ws, rName_folder(rName), rName, tCode, get_arg_date("", 1))
        Case Contains(action, "clean")
            ' in-transit
            If Contains(tCode, "47000149") Then
                Evaluate DeleteBlankRows(out_ws.usedRange)
                Set params_out.obj = out_ws
                params_out.type = "sheet"
            End If
        Case Contains(action, "split")
            sType = Split(action, " ")(1)
            header = Split(action, " ")(2)
            ' Move ws to new wb
            Set nWB = Workbooks.add
            out_ws.Copy nWB.Sheets(1)
            Set nWS = nWB.Sheets(out_ws.name)
            nWS.name = "master"
            Call Force_Del_Sheet(nWB, "Sheet1")
            Call Populate_Collection_Vertical(nWB, nWS.name, header, vals)
            Debug.Print vals.Count
            Call split_filter_sheet(nWS, header, vals)
            nWS.Activate
            ' if arg, try to copy sheet named arg
            If SheetExists(tWB, arg) Then tWB.Sheets(arg).Copy nWB.Sheets(1)
        Case Contains(action, "filter")
            Select Case LCase(tCode)
            Case "vt11"
                Dim e_rng As Range
                Dim fcol As Integer, xArg As String
                Dim chWS As Worksheet, exWS As Worksheet, adws As Worksheet
                Select Case LCase(rName)
                Case "addons"
                    ' pre-filter and delete scac codes
                    Dim fRng, f As Range
                    Set fRng = get_rng(shFilters, "addon_filter_scac", 1, 1, toLastRow:=1).SpecialCells(xlCellTypeConstants)
                    For Each f In fRng
                        Call delete_rows_with_criteria(out_ws.Parent, out_ws.name, f.value, _
                                                        FindColumn(out_ws, "service_agent", 1, 0))
                    Next f

                    ' make copies
                    Set adws = duplicatesheet(out_ws, "vt11_addons")
                    Set chWS = duplicatesheet(out_ws, "vt11_charters")
                    Set exWS = duplicatesheet(out_ws, "vt11_expedites")

                    With chWS
                        ' remove dups svc/trailer
                        'Call remove_dups_mult_cols(chWS, "service_agent", "trailer_number", "seal_number_2")
                        Set e_rng = adv_filter_fill(chWS, shFilters, "dispatch_filter_2", "seal_number_2", "planning_end", 7, _
                                                 bDeleteHidden:=1, bCheckVisibleOnly:=True)

                        Set e_rng = Nothing
                    End With

                    With exWS
                        ' remove dups svc/trailer
                        ' Call remove_dups_mult_cols(exWS, "service_agent", "trailer_number", "seal_number_2")
                        ' remove charters prior
                        Evaluate dbFilter_Exclude(exWS, chWS, _
                                                  FindColumn(chWS, "shipment_number", 1, 0), _
                                                  FindColumn(exWS, "shipment_number", 1, 0), _
                                                  bDeleteFiltered:=1)
                        ' Get expedites
                        Set e_rng = adv_filter_fill(exWS, shFilters, "dispatch_filter_1", "seal_number_2", "planning_end", 6, _
                                                 bDeleteHidden:=1)
                        Set e_rng = Nothing
                    End With

                    With adws
                        ' remove ex and ch
                        For Each vN In Array(exWS, chWS)
                            Evaluate dbFilter_Exclude(adws, vN.name, _
                                                      FindColumn(adws, "shipment_number", 1, 0), _
                                                      FindColumn(vN, "shipment_number", 1, 0), _
                                                      bDeleteFiltered:=1)
                        Next vN

                        ' get dups (addons)
                        ' sort
                        Call sort_multi_fields(adws, Col_Letter(FindColumn(adws, "seal_number_2", 1, 0)), xlDescending, _
                                               Col_Letter(FindColumn(adws, "trailer_number", 1, 0)), xlDescending, _
                                               Col_Letter(FindColumn(adws, "planning_end", 1, 0)), xlDescending, _
                                               Col_Letter(FindColumn(adws, "time", 1, 0)), xlDescending)

                        ' insert formula
                        Call insert_column(adws.Parent, adws.name, "concat", FindColumn(adws, "trailer_number", 1, 0) + 1)
                        Call insert_column(adws.Parent, adws.name, "bDup", FindColumn(adws, "concat", 1, 0) + 1)
                        Dim fml As String
                        fml = "=RC[-2]&RC[-1]"
                        Call Insert_Formula(adws.Parent, adws.name, "concat", fml): fml = ""

                        Dim coldiff As Double
                        coldiff = FindColumn(adws, "concat", 1, 0) - FindColumn(adws, "bDup", 1, 0)
                        fml = "=IF(COUNTIF(C[" & coldiff & "],RC[" & coldiff & "])>1,COUNTIF(C[" & coldiff & _
                              "],RC[" & coldiff & "])," & Chr(34) & "unique" & Chr(34) & ")"
                        Call Insert_Formula(adws.Parent, adws.name, "bDup", fml): fml = ""

                        Call clear_formats(adws)
                        Call delete_rows_with_criteria(adws.Parent, adws.name, "unique", FindColumn(adws, "bDup", 1, 0))

                        ' remove first instance
                        Dim rRng As Range, trl_col As Integer
                        Set rRng = get_rng(adws, "seal_number_2", 1, 0, toLastRow:=1)
                        trl_col = FindColumn(adws, "trailer_number", 1, 0)

                        ' delete bDup col
                        Evaluate delete_column_by_header(adws.Parent, adws.name, "concat")
                        Evaluate delete_column_by_header(adws.Parent, adws.name, "bDup")

                        ' loop_filter each row, delete 1 per pass
                        Evaluate loop_filter_col_action(adws, "seal_number_2", xDelete)


                        ' merge sheets
                        ' dims
                        Dim b As boundsObj
                        b = Bounds(adws)
                        Call copy_sheet_data(exWS, adws, "A2", "A" & b.lRow + 1, 0, 0)
                        b = Bounds(adws)
                        Call copy_sheet_data(chWS, adws, "A2", "A" & b.lRow + 1, 0, 0)

                        Set out_ws = adws
                    End With

                Case "dispatches"
                    Set out_ws = duplicatesheet(out_ws, "VT11_DISPATCHES")
                    Set chWS = duplicatesheet(out_ws, "vt11_charters")
                    Set exWS = duplicatesheet(out_ws, "vt11_expedites")

                    With chWS
                        ' remove dups svc/trailer
                        Call remove_dups_mult_cols(chWS, "service_agent", "trailer_number", "seal_number_2")
                        Set e_rng = adv_filter_fill(chWS, shFilters, "dispatch_filter_2", "seal_number_2", "shipping_type", 7, _
                                                 bDeleteHidden:=1)
                        Set e_rng = Nothing
                    End With

                    With exWS
                        ' remove dups svc/trailer
                        Call remove_dups_mult_cols(exWS, "service_agent", "trailer_number", "seal_number_2")
                        ' remove charters prior
                        Evaluate dbFilter_Exclude(exWS, chWS, _
                                                  FindColumn(chWS, "shipment_number", 1, 0), _
                                                  FindColumn(exWS, "shipment_number", 1, 0), _
                                                  bDeleteFiltered:=1)
                        ' Get expedites
                        Set e_rng = adv_filter_fill(exWS, shFilters, "dispatch_filter_1", "seal_number_2", "shipping_type", 6, _
                                                 bDeleteHidden:=1)
                        Set e_rng = Nothing
                        ' count of rows?
                    End With
                    With out_ws
                        ' remove dups svc/trailer
                        Call remove_dups_mult_cols(out_ws, "service_agent", "trailer_number")

                        ' merge into out_ws and overwrite
                        ' EXPEDITES
                        Evaluate dbFilter_Exclude(out_ws, exWS, _
                                                  FindColumn(out_ws, "shipment_number", 1, 0), _
                                                  FindColumn(exWS, "shipment_number", 1, 0), _
                                                  bDeleteFiltered:=1)
                        ' dims
                        b = Bounds(out_ws)
                        Call copy_sheet_data(exWS, out_ws, "A2", "A" & b.lRow + 1, 0, 0)

                        ' CHARTER
                        Evaluate dbFilter_Exclude(out_ws, chWS, _
                                                  FindColumn(out_ws, "shipment_number", 1, 0), _
                                                  FindColumn(chWS, "shipment_number", 1, 0), _
                                                  bDeleteFiltered:=1)
                        ' dims
                        b = Bounds(out_ws)
                        Call copy_sheet_data(chWS, out_ws, "A2", "A" & b.lRow + 1, 0, 0)

                        ' Post filter/merge
                        ' remove via tunnel
                        Evaluate delete_rows_with_criteria(tWB, out_ws.name, "tunnel", FindColumn(out_ws, "trailer_number"))
                        Evaluate delete_rows_with_criteria(tWB, out_ws.name, "scrap", FindColumn(out_ws, "seal_number_2"))

                        ' Sort
                        ' Call sort_multi_fields(out_ws, "actshipmnt_end_time", xlAscending, "actualshipmentend", xlAscending)
                        Call SortCol(out_ws, "actshipmnt_end_time")
                        Call SortCol(out_ws, "actualshipmentend")

                        ' Totals
                        Dim s As boundsObj, d As boundsObj
                        ' dispatches
                        d = Bounds(out_ws)
                        d.sh.Cells(d.lRow + 1, d.headers("seal_number_2")).value = d.lRow - 1
                        ' expedites
                        d = Bounds(out_ws)
                        s = Bounds(exWS)
                        d.sh.Cells(d.lRow + 1, d.headers("seal_number_2")).value = s.lRow - 1
                        d.sh.Cells(d.lRow + 1, d.headers("seal_number_2")).Interior.ColorIndex = 6
                        ' charters
                        d = Bounds(out_ws)
                        s = Bounds(chWS)
                        d.sh.Cells(d.lRow + 1, d.headers("seal_number_2")).value = s.lRow - 1
                        d.sh.Cells(d.lRow + 1, d.headers("seal_number_2")).Interior.ColorIndex = 7

                    End With
                    ' New book
                    Set nWB = Workbooks.add
                    out_ws.Copy After:=nWB.Sheets(1)
                    Call Force_Del_Sheet(nWB, nWB.Sheets(1).name)

                    ' Save
                    Application.DisplayAlerts = False
                    nWB.SaveAs get_save_name(rName_folder(rName), rName, tCode, datevar)
                    Application.DisplayAlerts = True
                    Set out_ws = nWB.Sheets(out_ws.name)
                Case "mg_snapshot"
                    ' filter ovtransport_status = 1
                    ' Split(Replace(action, "filter:", ""), "|")(0)
                    fcol = FindColumn(out_ws, parse_arg(action, "filter:", 0), 1, 0)
                    xArg = parse_arg(action, "filter:", 1)
                    Call FilterColumn(out_ws.Parent, out_ws.name, fcol, xlFilterValues, xArg)
                Case Else                        ' default?
                    ' // example >> filter:ovtransport_status|=1 << example //
                    ' Split(Replace(action, "filter:", ""), "|")(0)
                    fcol = FindColumn(out_ws, parse_arg(action, "filter:", 0), 1, 0)
                    xArg = parse_arg(action, "filter:", 1)
                    If Contains(action, "delete:hidden") Then
                        Call FilterColumn(out_ws.Parent, out_ws.name, fcol, xlFilterValues, Split(xArg, " ")(0))
                        Evaluate deleteHidden(out_ws.Parent, out_ws.name)
                    Else
                        Call FilterColumn(out_ws.Parent, out_ws.name, fcol, xlFilterValues, xArg)
                    End If
                End Select                       ' rName

            Case "zmdesnr"
                ' filter from filtersheets
                Call RemoveDuplicates(out_ws, "serial_")
                Call Clear_AutoFilter(out_ws, 1)
                Evaluate advanced_filter(out_ws, get_rng(shFilters, rName & "_" & tCode & "_filter", 0, 0).CurrentRegion)
                Call SortCol(out_ws, "created_on")
                ' Delete hidden
                Call deleteHidden(out_ws.Parent, out_ws.name)
                ' Save as new book
                Set out_ws = save_sheet_new_book(out_ws, rName_folder(rName), rName, tCode, datevar)
                Set params_out.obj = out_ws
                params_out.type = "sheet"

            Case "lx03"
                ' Filter numeric
                Call FilterColumn(out_ws.Parent, out_ws.name, FindColumn(out_ws, "storage_bin", 1, 0), xlFilterValues, "<>*", "<>")
                ' Delete hidden
                Call deleteHidden(out_ws.Parent, out_ws.name)
                ' Sort by storage_type
                Call SortCol(out_ws, "storage_type")
                ' Save
                Set out_ws = save_sheet_new_book(out_ws, rName_folder(rName), rName, tCode, _
                                                 get_json_date(arg), 1, "MM-DD-YY", " ")
                Evaluate LetOrSetElement(params_out.obj, out_ws)

            Case Else
            End Select                           ' tCode
        Case Contains(action, "sort")
            Select Case tCode
            Case "zmdesnr"
                Dim sRng As Range
                Debug.Print arg("layout")
                ' Sort
                Call Clear_AutoFilter(out_ws, 1)
                Call sort_multi_fields(out_ws, Col_Letter(FindColumn(out_ws, "pallet", 1)), xlDescending, _
                                       Col_Letter(FindColumn(out_ws, "parent_serial_number", 1)), xlAscending)
                Call remove_dupes(out_ws, "lm_serial_number")
                ' Filter and delete
                Evaluate advanced_filter(out_ws, get_rng(shFilters, rName & "_" & tCode & "_filter_1", 0, 0).CurrentRegion)
                Call delete_visible(out_ws)
                out_ws.ShowAllData

                ' Remove partial dups
                Evaluate advanced_filter(out_ws, get_rng(shFilters, rName & "_" & tCode & "_filter_2", 0, 0).CurrentRegion)
                Call remove_dupes(out_ws, "parent_serial_number", dupRemoveVisible)
                out_ws.ShowAllData
                Call Clear_AutoFilter(out_ws, 1)
                ' New book
                Set out_ws = save_sheet_new_book(out_ws, rName_folder(rName), rName, tCode, datevar)
                Set params_out.obj = out_ws
                params_out.type = "sheet"
            Case Else
            End Select                           ' tCode
        Case Contains(action, "summarize"), Contains(action, "update")
            Select Case LCase(tCode)
            Case "lx04"
                Dim sName As String: sName = tCode & "_" & "summary"
                Dim hdrList As New BetterArray
                hdrList.Push "occupied", "empty", "usage", "load"
                ' Assign or Create sheet if doesn't exist, temp_ws should still exist
                ' modified create_or_get 220825 removed (rName)
                Set out_ws = create_or_get_sheet(tWB, UCase(rName & "_" & sName), hdrList)

                Dim target_range As Range, source_range As Range
                ' Get Dest
                Set target_range = lastRowInColLetter(out_ws, "A", 1, 0)
                Debug.Print target_range.address
                ' loop and paste values
                Dim xRng As Range
                target_range.value = temp_ws.name
                Set xRng = target_range.offset(0, 1)
                Do While out_ws.Cells(1, xRng.Column).value <> ""
                    xRng.value = get_rng(temp_ws, out_ws.Cells(1, xRng.Column).value, 1, 0, xlPart).value
                    Set xRng = xRng.offset(0, 1)
                Loop
                Debug.Print
            Case "vt11"
                ' out_ws
                Select Case True
                Case Contains(rName, "ob_")
                    Set out_ws = duplicatesheet(out_ws, "ob_vt11")
                    Call insert_column(out_ws.Parent, out_ws.name, "customer_name", 1)
                    ' Evaluate delete_column_by_header(out_ws.Parent, out_ws.name, "shipment_number")
                    ' Replace seal_number_2 with logistic for honda.
                    Call replace_column_values(out_ws, "logistic_number", "seal_number_2", "???-???*???")
                    Call replace_column_values(out_ws, "logistic_number", "seal_number_2", "FV??A")
                    ' replace 0:00 times with - in column
                    ' Dim searchVal As String, replaceVal As String
                    ' searchVal = "00:00:00"
                    ' Evaluate FilterColumn(tWB, out_ws.name, 16, xlFilterValues, "=" & #12:00:00 AM#)
                    ' replaceVal = "-"
                    ' Evaluate get_rng(out_ws, "acttranspstarttime", 1, 0, toLastRow:=True) _
                                    .Replace(searchVal, replaceVal, LookAt:=xlPart)
                    Call delete_column_by_header(out_ws.Parent, out_ws.name, "logistic_number")
                    ' insert comments on right
                    b = Bounds(out_ws)
                    Call insert_column(out_ws.Parent, out_ws.name, "current status", b.lCol + 1)
                    out_ws.Columns.AutoFit

                    ' Get customer name for VEND via shipto

                    ' lookup VEND
                    ' Call advanced_filter(out_ws, get_rng(shFilters, "ob_filter_1", 0, 1).CurrentRegion)
                    ' lookup routes with shipto
                    Evaluate regex_filter(out_ws, get_var(shFilters, "ob_filter_route_pattern"), get_var(shFilters, "ob_filter_route_header"))
                    ' Get cust name via shipto xd03
                    Call xd03_loop(sess, "xd03", rName, out_ws, "seal_number_2", "customer_name")
                    out_ws.Rows.Hidden = False
                    ' Call vt11_display_screen_info(out_ws, "shipment_number", sess, tCode, sap_var, arg, "logistic_number")
                    ' clear formats/filters
                    ' Call clear_formats(out_ws)
                    ' Debug.Print ""
                    ' wildcard match customer name
                    Evaluate wildcard_lookup(out_ws, "pattern_match", "seal_number_2", "customer_name")
                    ' sort
                    Call sort_multi_fields(out_ws, Col_Letter(FindColumn(out_ws, "descripof_shipment", 1, 0)), xlAscending, _
                                           Col_Letter(FindColumn(out_ws, "seal_number_2", 1, 0)), xlAscending)
                    ' remove dups
                    Call RemoveDuplicates_Mult_2(tWB, out_ws.name, "seal_number_2", "descripof_shipment")

                    ' Formatting
                    Evaluate formatting_ob(out_ws)

                    ' filter out 7 from status
                    Call FilterColumn(out_ws.Parent, out_ws.name, FindColumn(out_ws, "ovtransport_status", 1, 0), xlFilterValues, "<>7") '
                    ' multi filter
                    Call FilterColumn(out_ws.Parent, out_ws.name, FindColumn(out_ws, "descripof_shipment", 1, 0), _
                                      xlOr, replace_delim(Format(DateAdd("d", -1, TODAYS_DATE), "mm/dd"), "/") & _
                                      "*", replace_delim(Format(TODAYS_DATE, "mm/dd"), "/") & "*", , , 1)

                    ' replace select headers
                    Dim r As Range: Set r = get_rng(shLookup, "ob_replace_headers", 1, 0, toLastRow:=True)
                    Dim i
                    For Each i In r.Cells
                        With out_ws.Cells.Rows(1)
                            .Replace i.value, i.offset(0, 1).value
                        End With
                    Next i
                    Set r = Nothing

                    ' add comments to select headers
                    Set r = get_rng(shFilters, "ob_button_comments", 1, 1, toLastRow:=True)
                    Dim c
                    Set c = Nothing
                    For Each i In r.Cells
                        If Len(i) < 1 Then GoTo skip_comment
                        With out_ws.Cells.Rows(1)
                            Set c = .Find(i)
                            If Not c.Comment Is Nothing Then c.Comment.Delete
                            c.AddComment i.offset(0, 1).value
                            ' c.Comment.text i.Offset(0, 1).value
                        End With

skip_comment:
Next i
' reset comments to default
dbLog.log str_form & "Resetting comments to default" & str_form
Evaluate resetComments(out_ws.usedRange)
out_ws.Columns.AutoFit

                    ' resize last col
                    b = Bounds(out_ws)
                    out_ws.Columns(b.lCol).ColumnWidth = 27
                    Evaluate set_column_widths(out_ws, shLayout, "ob_col_widths")
                    ' hide shipment column
                    ' out_ws.Columns(FindColumn(out_ws, "shipment_number", 1, 0)).EntireColumn.Hidden = True

                    If Contains(action, "summarize") Then
                        ' New book
                        Set nWB = Workbooks.add
                        ' WS1
                        out_ws.Copy After:=nWB.Sheets(1)
                        Dim ws1 As Worksheet
                        Set ws1 = nWB.Sheets(out_ws.name)
                        ' rename
                        ws1.name = UCase(rName) & "_" & "display"

                        ' WS2
                        ' make hidden sheet and copy validation
                        Dim ws2 As Worksheet
                        Set ws2 = nWB.Sheets("Sheet1")
                        ws2.name = "validation"
                        shFilters.Cells.Find("ob_validation_1").offset(0, 1).CurrentRegion.Copy Destination:=ws2.Cells(1, 1)
                        ws2.Columns(1).EntireColumn.Delete
                        ws2.Visible = xlHidden
                        Evaluate add_validation(ws1, "current status", get_rng(ws2, "status", 0, 0, toLastRow:=True))

                        ' WS3
                        Dim ws3 As Worksheet
                        Set ws3 = duplicatesheet(ws1, UCase(rName & "_" & "data"))
                        Evaluate data_ob_sheet(ws3)
                        ' END WS3

                        ws1.Activate

                        ' Save
                        If arg.Exists("save_date") Then
                            nWB.SaveAs get_save_name(rName_folder(rName), rName, "", get_json_date(arg("save_date")), reverseDateName:=True)
                        Else
                            nWB.SaveAs get_save_name(rName_folder(rName), rName, "", datevar, reverseDateName:=True)
                        End If
                        Set out_ws = ws1
                    Else
                        Set temp_ws = out_ws
                        temp_ws.name = lValue & tCode
                        Set out_ws = Nothing
                    End If
                Case Else
                End Select ' true contains(rname, ob_)
            Case Else
            End Select  ' tcode (lx04, vt11)
        Case Contains(action, "sheet_out:")
            out_ws.name = parse_arg(action, "sheet_out:", 0, , 1)
            Set params_out.obj = out_ws
            params_out.type = "sheet"
        Case Else
        End Select  ' action contains

out:

        Application.StatusBar = False
    End With

    If Len(params_out.err) > 1 Then
        run_tcode_w_variant = params_out
        Exit Function
    End If
    If out_ws Is Nothing Then
        params_out.type = "sheet"
        params_out.name = temp_ws.name
        Set params_out.obj = temp_ws
    Else
        ' Use rName if provided
        If Len(rName) > 0 And Not Contains(out_ws.name, rName) Then
            out_ws.name = rename_sheet_if_exists(out_ws.Parent, UCase(rName) & "_" & out_ws.name, "_")
        End If
        If IsEmpty(params_out.obj) Then
            params_out.type = "book"
            Set params_out.obj = nWB
        ElseIf params_out.obj Is Nothing Then
            params_out.type = "book"
            Set params_out.obj = nWB
        End If
    End If
    run_tcode_w_variant = params_out
    Exit Function

errlv:

    MsgBox msg, vbExclamation
    GoTo out

End Function
