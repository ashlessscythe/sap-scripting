session.findById("wnd[0]").resizeWorkingPane 137,26,false
session.findById("wnd[0]/usr/txt%_IF_VSTEL_%_APP_%-TEXT").setFocus
session.findById("wnd[0]/usr/txt%_IF_VSTEL_%_APP_%-TEXT").caretPosition = 22
session.findById("wnd[0]").sendVKey 17
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[1]/usr/cntlALV*CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%\_IT_VBELN*%_APP_%-VALU_PUSH").press
session.findById("wnd[1]").sendVKey 24
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/lbl[1,16]").setFocus
session.findById("wnd[1]/usr/lbl[1,16]").caretPosition = 4
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "num_pkg"
session.findById("wnd[2]/usr/chkSCAN_STRING-START").setFocus
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 5
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[5]/menu[1]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]").close
