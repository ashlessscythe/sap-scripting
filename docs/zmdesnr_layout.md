If Not IsObject(application) Then
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
Set connection = application.Children(0)
End If
If Not IsObject(session) Then
Set session = connection.Children(0)
End If
If IsObject(WScript) Then
WScript.ConnectObject session, "on"
WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").resizeWorkingPane 137,26,false
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 5,"STATUS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "5"
session.findById("wnd[0]/mbar/menu[4]/menu[0]/menu[0]").select
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 3
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "3"
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 2
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "2"
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 0
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 9,"TKNUM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleColumn = "LGPLA"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "9"
