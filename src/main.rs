/*
Origin VBS Script:

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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nfpl9"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtFKKL1-GPART").text = "12345"
session.findById("wnd[0]/usr/cmbFKKL1-LSTYP").key = "OPEN"
session.findById("wnd[0]/usr/cmbFKKL1-LSTYP").setFocus
session.findById("wnd[0]/tbar[0]/btn[0]").press

And how this would be written as Rust:
*/

use sap_scripting::*;

fn main() -> crate::Result<()> {
    // Initialise the environment.
    let com_instance = SAPComInstance::new().expect("Couldn't get COM instance");
    let wrapper = com_instance.sap_wrapper().expect("Couldn't get SAP wrapper");
    let engine = wrapper.scripting_engine().expect("Couldn't get GuiApplication instance");

    let connection = match sap_scripting::GuiApplication_Impl::children(&engine)?.element_at(0)? {
        SAPComponent::GuiConnection(conn) => conn,
        _ => panic!("expected connection, but got something else!"),
    };
    eprintln!("Got connection");
    let session = match sap_scripting::GuiConnection_Impl::children(&connection)?.element_at(0)? {
        SAPComponent::GuiSession(session) => session,
        _ => panic!("expected session, but got something else!"),
    };

    // get user input for tcode to run
    println!("Enter tcode to run (default: lm01):");
    let mut tcode = String::new();
    match std::io::stdin().read_line(&mut tcode) {
        Ok(_) => {
            tcode = tcode.trim().to_owned();
            if tcode.is_empty() {
                tcode = "lm01".to_owned();
            }
            println!("You entered: {}", tcode);
        },
        Err(e) => eprintln!("Failed to read line: {}", e),
    }

    if let SAPComponent::GuiMainWindow(wnd) = session.find_by_id("wnd[0]".to_owned())? {
        // starting transaction
        eprintln!("Starting transaction");
        match session.start_transaction(tcode) {
            Ok(_) => eprintln!("Transaction started"),
            Err(e) => eprintln!("Error starting transaction: {:?}", e),
        }

        // status message (if any)
        match wnd.find_by_id("wnd[0]/sbar".to_owned()) {
            Ok(SAPComponent::GuiStatusbar(statusbar)) => {
                let text = statusbar.text()?;
                if text.len() == 0 {
                    eprintln!("No status message")
                } else {
                    eprintln!("Status message: {}", text)
                }
            }
            _ => eprintln!("No status bar"),
        }

    } else {
        panic!("no window!");
    }

    Ok(())
}