import win32com.client

#Connect to your SAP
def get_sap_client():
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