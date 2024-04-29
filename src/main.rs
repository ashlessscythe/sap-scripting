use std::vec;
use sap_scripting::*;

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

    let sess2 = match create_session(&connection, &session, 1) {
        Ok(sess) => sess,
        Err(e) => {
            eprintln!("Error creating session: {:?}", e);
            return Ok(());
        }
    };

    if let SAPComponent::GuiMainWindow(wnd) = session.find_by_id("wnd[0]".to_owned())? {
        // close all modal windows        
        close_all_modal_windows(&session)?;

        // ask user if list or single
        let run_single = prompt_bool("Run single tcode?").expect("error getting input");

        if run_single {
            // get tcode
            let tcode = "zmdetpc";
            match start_user_tcode(&session, tcode.to_owned(), Some(true)) {
                Ok(_) => {
                    handle_status_message(&session)?;
                    eprintln!("Tcode ({}) started", &tcode);

                    let info = get_session_info(&session).expect("Error getting session info");
                    println!("Current tcode: {:?}", info.transaction());
                    println!("Current screen: {:?}", info.screen_number());
                    println!("Current user: {:?}", info.user());

                    let info = get_session_info(&sess2).expect("Error getting sess2 info");
                    println!("Current tcode: {:?}", info.transaction());
                    println!("Current screen: {:?}", info.screen_number());
                    println!("Current user: {:?}", info.user());

                    // send variant key
                    send_vkey_main(&wnd, 17)?;

                    // get modal window
                    let w2 = get_modal_window(&session, &1).expect("Error getting modal window");

                    // clear prev text
                    set_text_modal(&w2, &1,"/usr/txtENAME-LOW", "")?;

                    // get variant name
                    let var_name = prompt_str("Variant name").expect("failed to get input");
                    
                    // enter variant name
                    set_text_modal(&w2, &1, "/usr/txtV-LOW", &var_name)?;

                    // send f8 (close modal window)
                    send_vkey_modal(&w2, 8)?;
                    handle_status_message(&session)?;
                    close_modal_window(&session, None)?;

                    // ask for delivery number
                    let del_num = prompt_str("Delivery number (default blank)").expect("failed to get input");

                    // enter delivery number
                    set_ctext_main(&session,"/usr/ctxtS_VBELN-LOW", &del_num)?;

                    // ask execute?
                    prompt_execute(&wnd, 8)?;
                    handle_status_message(&session)?;

                    // get value from table
                    let vals = match get_grid_values(&session, "VBELN") {
                        Ok(vals) => {
                            eprintln!("Got values: {:?}", vals);
                            vals
                        },
                        Err(e) => {
                            eprintln!("Error getting values: {:?}", e);
                            vec![]
                        }
                    };

                    println!("Vals count {}", vals.len());

                },
                Err(e) => eprintln!("Error starting tcode: {:?}", e),
            }
        } else {
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
            let run_mult = prompt_bool("Continue running tcodes?").expect("Error getting input");
                
            // if bool is true run, else don't
            if run_mult {
                for tcode in list {
                    match start_user_tcode(&session, tcode.to_owned(), Some(false)) {
                        Ok(_) => {
                            handle_status_message(&session)?;
                            eprintln!("Tcode started: {}", tcode)
                        },
                        Err(e) => eprintln!("Error starting tcode: {:?}", e),
                    }
                }
            }
        };

        // ask if go back
        let go_back = prompt_bool("Go back to main menu?").expect("Error getting input");
        if go_back {
            match start_user_tcode(&session, "SESSION_MANAGER".to_owned(), Some(false)) {
                Ok(_) => {
                    handle_status_message(&session)?;
                    close_all_modal_windows(&session)?;
                    eprintln!("Tcode started: session_manager")
                },
                Err(e) => eprintln!("Error starting tcode: {:?}", e),
            }
        }

    } else {
        eprintln!("No status bar");
    }
    // out
    Ok(())
}