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
session.findById("wnd[0]/tbar[0]/okcd").text = "/NS_ALR_87013611"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxt$1KOKRE").text = "561A"
session.findById("wnd[0]/usr/txt$1GJAHR").text = "2022"
session.findById("wnd[0]/usr/ctxt$1PERIV").text = "11"
session.findById("wnd[0]/usr/ctxt$1PERIB").text = "11"
session.findById("wnd[0]/usr/ctxt$1VERP").setFocus
session.findById("wnd[0]/usr/ctxt$1VERP").caretPosition = 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").unselectAll
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]").close
session.findById("wnd[0]/mbar/menu[0]/menu[3]").select
session.findById("wnd[1]").close
session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\LegacyApp\Python36\py36wokspace"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "S90820_3611_202111.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[0]").press
