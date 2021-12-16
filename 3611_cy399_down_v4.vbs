Dim Arg
Set Arg = WScript.Arguments

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/ns_alr_87013611"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxt$1KOKRE").text = "561a"
session.findById("wnd[0]/usr/txt$1GJAHR").text = Arg(0) '2022
session.findById("wnd[0]/usr/ctxt$1PERIV").text = Arg(1) '10
session.findById("wnd[0]/usr/ctxt$1PERIB").text = Arg(1) '10
session.findById("wnd[0]/usr/ctxt_1KOSET-LOW").text = "cy-399cz"
session.findById("wnd[0]/usr/ctxt_1KOSET-LOW").setFocus
session.findById("wnd[0]/usr/ctxt_1KOSET-LOW").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/lbl[36,12]").setFocus
session.findById("wnd[0]/usr/lbl[36,12]").caretPosition = 11
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/lbl[1,1]").caretPosition = 26
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = -1
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BELNR"
session.findById("wnd[0]/tbar[1]/btn[40]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/lbl[1,7]").setFocus
session.findById("wnd[0]/usr/lbl[1,7]").caretPosition = 0
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\LegacyApp\Python36\py36wokspace"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "cy399_11231511.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
session.findById("wnd[1]/tbar[0]/btn[0]").press
