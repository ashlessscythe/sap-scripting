use std::vec;
use sap_scripting::*;

mod test_func;

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
    let run_test = prompt_bool("Run test tcode?").expect("error getting input");

    if run_test {
        // test stuff?
        test_func::run_test_tcode(s2)?;
    } else {
        // run list
        let list = get_list_from_file("tcodes.txt").expect("Error getting list from file");
        run_list(&s2, list, true);
    };

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
