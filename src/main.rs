use std::vec;
use sap_scripting::*;

mod func;
use func::*;


fn main() -> crate::Result<()> {
     // Initialise the environment.
    let com_instance = SAPComInstance::new().expect("Couldn't get COM instance");
    let wrapper = com_instance.sap_wrapper().expect("Couldn't get SAP wrapper");
    let engine = wrapper.scripting_engine().expect("Couldn't get GuiApplication instance");

    let connection = match sap_scripting::GuiApplication_Impl::children(&engine).unwrap().element_at(0) {
        Ok(SAPComponent::GuiConnection(conn)) => conn,
        _ => panic!("expected connection, but got something else!"),
    };

    // leave 0 alone
    let s0 = create_or_get_session(&connection, 0).expect("Error creating session");
    let w0 = match s0.find_by_id("wnd[0]".to_owned()) {
        Ok(SAPComponent::GuiMainWindow(w0)) => w0,
        _ => panic!("expected main window, but got something else!"),
    };

    // check if user logged in
    let info = get_session_info(&s0).expect("Error getting session info");
    if info.transaction().unwrap() == "S000".to_owned() {
        // log in
        let logged_in = log_in_sap(&s0).expect("Error logging in");
        // if not logged in, exit
        if !logged_in {
            eprintln!("Not logged in, exiting");
            return Ok(());
        }
    }; 

    let s1 = create_or_get_session(&connection, 1).expect("Error creating session");
    let w1 = match s1.find_by_id("wnd[0]".to_owned()) {
        Ok(SAPComponent::GuiMainWindow(w1)) => w1,
        _ => panic!("expected main window, but got something else!"),
    };
    // close all modal windows        
    close_all_modal_windows(&s1)?;
    
    let s2 = create_or_get_session(&connection, 2).expect("Error creating session");
    let w2 = match s2.find_by_id("wnd[0]".to_owned()) {
        Ok(SAPComponent::GuiMainWindow(w2)) => w2,
        _ => panic!("expected main window, but got something else!"),
    };

    // ask user if list or single
    let run_single = prompt_bool("Run single tcode?").expect("error getting input");

    if !run_single {
        // run list
        let list = get_list_from_file("tcodes.txt").expect("Error getting list from file");
        run_list(&s2, list, true);
    } else {
        // get tcode
        let tcode = "zmdetpc";
        let (_, t) = start_user_tcode(&s1, tcode.to_owned(), Some(true)).expect("Error starting tcode");
        
        let info = get_session_info(&s1).expect("Error getting session info");
        println!("Current tcode: {:?}", info.transaction());
        println!("Current screen: {:?}", info.screen_number());
        println!("Current user: {:?}", info.user());
        
        start_user_tcode(&s2, tcode.to_owned(), Some(false)).expect("Error starting tcode");
        let info = get_session_info(&s2).expect("Error getting sess info");
        println!("Current tcode: {:?}", info.transaction());
        println!("Current screen: {:?}", info.screen_number());
        println!("Current user: {:?}", info.user());

        // if tcode changed, do something
        if t != tcode {
            println!("tcode changed from {} to {}", tcode, t);
            println!("finding by name");
            let tab = match w1.find_by_id("wnd[0]/usr/tabsTABSTRIP_TABB1/tabpUCOMM2".to_owned()).expect("Error finding by id") {
               SAPComponent::GuiTab(tab) => {
                println!("found tab");
                    tab 
                },
                other => {
                    print_sap_component_type(&other);
                    panic!("expected tab, but got something else!");
                },
            };
            tab.select().expect("Error selecting tab");
        };

        // apply variant
        let variant_name = prompt_str("Variant name (default blank)").expect("failed to get input");
        apply_variant(&s1, &variant_name).expect("Error applying variant");

        // ask execute?
        prompt_execute(&w1, 8)?;
        handle_status_message(&s1)?;

        // get value from table
        let vals = match get_grid_values(&s1, "VBELN") {
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

    };

    // ask if go back
    let go_back = prompt_bool("Go back to main menu?").expect("Error getting input");
    if go_back {
        match start_user_tcode(&s1, "SESSION_MANAGER".to_owned(), Some(false)) {
            Ok(_) => {
                handle_status_message(&s1)?;
                close_all_modal_windows(&s1)?;
                eprintln!("Tcode started: session_manager")
            },
            Err(e) => eprintln!("Error starting tcode: {:?}", e),
        }
    }
    // out
    Ok(())
}

fn run_list(session: &GuiSession, list: Vec<String>, prompt: bool) {
    let run_mult: bool;
    if prompt { 
        // ask user if ok to continue
        run_mult = prompt_bool("Continue running tcodes?").expect("Error getting input");
    } else {
        run_mult = true;
    }
                
    // if bool is true run, else don't
    if run_mult {
        for tcode in list {
            match start_user_tcode(&session, tcode.to_owned(), Some(false)) {
                Ok(_) => {
                    handle_status_message(&session).expect("Error handling status message");
                    eprintln!("Tcode started: {}", tcode)
                },
                Err(e) => eprintln!("Error starting tcode: {:?}", e),
            }
        }
    }
}