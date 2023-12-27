import win32com.client

#Connect to your SAP
def get_sap_client():
    '''
        This function accesses SAP session running on your PC. Please be aware that you must first log in to your SAP account 
        before using this function.
    '''
    sap_gui = win32com.client.GetObject("SAPGUI") #Connect to already running instance of Application
    if not type(sap_gui) == win32com.client.CDispatch:
        return

    application = sap_gui.GetScriptingEngine #Trying to access SAP GUI
    if not type(application) == win32com.client.CDispatch:
        sap_gui = None
        return

    #SAP Instances Count
    for conn in range(application.Children.Count):
        #Get proper connection
        connection = application.Children(conn)

        #Get the session
        for sess in range(connection.Children.Count):
            #Loop through each connection and return session that are in 'Sesison Manager'
            session = connection.Children(sess)
            if session.Info.Transaction == "SESSION_MANAGER":
                return session
            else:
                return

ses = get_sap_client()

DESKTOP_PATH = ""
CREATED_FROM = "" #dd.mm.yyyy

ses.findById("wnd[0]/tbar[0]/okcd").text = "/nse16n"
ses.findById("wnd[0]").sendVKey(0)
ses.findById("wnd[0]/usr/ctxtGD-TAB").text = "vbak"
ses.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
ses.findById("wnd[0]").sendVKey(0)
ses.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnOPTION[1,2]").setFocus()
ses.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnOPTION[1,2]").press()
ses.findById("wnd[1]/usr/cntlGRID/shellcont/shell").setCurrentCell(6,"TEXT")
ses.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectedRows = "6"
ses.findById("wnd[1]/tbar[0]/btn[0]").press()
ses.findById("wnd[0]/usr/txtGD-MAX_LINES").text = "5000"
ses.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").text = CREATED_FROM
ses.findById("wnd[0]/tbar[1]/btn[8]").press()
#EXPORT WINDOW
ses.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
ses.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
#EXPORT PROPERTIES
ses.findById("wnd[1]/usr/ctxtDY_PATH").text = DESKTOP_PATH
ses.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export.xlsx"
#EXPORT SAVE
ses.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
ses.findById("wnd[1]/tbar[0]/btn[0]").press()