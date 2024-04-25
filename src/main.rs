use sap_scripting::*;
use std::io::stdin;
mod func;
use func::*;

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

    if let SAPComponent::GuiMainWindow(wnd) = session.find_by_id("wnd[0]".to_owned())? {
        // close all modal windows        
        close_all_modal_windows(&session)?;

        // get list
        let list =  match get_list_from_file("tcodes.txt") {
            Ok(list) => {
                eprintln!("Got list: {:?}", list);
                list
            },
            Err(e) => {
                eprintln!("Error getting list: {:?}", e);
                vec![]
            }
        };

        // ask user if ok to continue
        println!("found {} tcodes. Continue? (y/n): ", list.len());
        let mut continue_run = String::new();
        stdin().read_line(&mut continue_run).expect("Failed to read line");
        continue_run = continue_run.trim().to_owned();
        if ["y", "Y", ""].contains(&continue_run.as_str()) {
            for tcode in list {
                match start_user_tcode(&session, tcode, Some(false)) {
                    Ok(_) => {
                        handle_status_message(&session)?;
                        eprintln!("Tcode started")
                    },
                    Err(e) => eprintln!("Error starting tcode: {:?}", e),
                }
            }
        } else {
            eprintln!("Not runing tcodes. Exiting...");
            return Ok(());
        }

        let mut b_continue = handle_status_message(&session)?;

        // read objects on screen
        if b_continue {
            match session.find_by_id("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW".to_owned()) {
                Ok(SAPComponent::GuiTabStrip(ctab)) => {
                    let children = ctab.children().into_iter();
                    println!("Tabs: {}", children.count());
                }
                _ => eprintln!("No tab"),
            }
        }

        // enter
        if b_continue { b_continue = handle_status_message(&session)?; }

        if b_continue {
            match wnd.send_v_key(0) {
                Ok(_) => eprintln!("VKey sent"),
                Err(e) => eprintln!("Error sending VKey: {:?}", e),
            }
        }

        if b_continue { b_continue = handle_status_message(&session)?; }

        // inside tcode
        if b_continue {
            println!("Enter tab number (default: 1): ");
            match session.find_by_id("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW".to_owned()) {
                Ok(SAPComponent::GuiTabStrip(ctab)) => {
                    let children = ctab.children().into_iter();
                    println!("Tabs: {}", children.count());
                }
                _ => eprintln!("No tab"),
            }
        }

        let b_continue = match handle_status_message(&session) {
            Ok(b) => b,
            Err(e) => {
                eprintln!("Error handling status message: {:?}", e);
                false
            },
        }; 

        // ask if go back
        if b_continue {
            println!("Go back? (y/n): ");
            let mut go_back = String::new();
            stdin().read_line(&mut go_back).expect("Failed to read line");
            go_back = go_back.trim().to_owned();
            if ["y", "Y", ""].contains(&go_back.as_str()) {
                match session.start_transaction("SESSION_MANAGER".to_owned()) {
                    Ok(_) => {
                        eprintln!("main menu");
                        // check if window popup
                        close_modal_window(&session, None)?;
                    },
                    Err(e) => eprintln!("Error going to main menu: {:?}", e),
                }
            }
        }

    } else {
        panic!("no window!");
    }

    Ok(())
}