use std::{ops::Deref, os::windows::process, sync::Arc, vec};
use sap_scripting::*;
use test_func::{multi_tcode, run_test_tcode};
use tokio::runtime::Runtime;

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
    
    // ask user if list or single
    let run_test = prompt_bool("Run test tcode?").expect("error getting input");

    let conn = Arc::new(connection);

    if run_test {
        // test stuff?
        let result = async {
            // test_func::parallel_tcodes(&*conn, 1).expect("error running parallel");
            test_func::run_test_tcode(s1.into(), conn,1).await.expect("Error running test tcode");
        };
        // run block using executor
        let rt = Runtime::new().unwrap();
        rt.block_on(result);

    } else {
        let r = async {
            process_session(conn, 1).await.expect("Error processing session");
        };

        // run block using executor
        let mut rt = Runtime::new().unwrap();
        rt.block_on(r);
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

// async process session
async fn process_session(connection: Arc<GuiConnection>, idx: i32) -> Result<()> {
    let session_result = create_or_get_session_async(&connection, idx).await;
    match session_result {
        Ok(session) => {
            // Use session here
            multi_tcode(connection, 1).await.expect("Error running multi tcode");
            Ok(())
        },
        _ => {
            eprint!("Error creating session");
            Ok(())
        },
    }
}
