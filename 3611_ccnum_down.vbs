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
session.findById("wnd[0]/usr/ctxt_1KSTAR-LOW").text = Arg(2) 's90803
session.findById("wnd[0]/usr/ctxt_1KSTAR-LOW").setFocus
session.findById("wnd[0]/usr/ctxt_1KSTAR-LOW").caretPosition = 6
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/lbl[36,14]").setFocus
session.findById("wnd[0]/usr/lbl[36,14]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/lbl[1,1]").caretPosition = 26
session.findById("wnd[1]").sendVKey 2
'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleRow = 366
'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = 404
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "404"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").pressTotalRowCurrentCell
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
