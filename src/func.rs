use sap_scripting::*;
use std::io::stdin;

pub fn get_list_from_file(file_path: &str) -> Result<Vec<String>> {
    let contents = std::fs::read_to_string(file_path)
        .expect("Something went wrong reading the file");
    let list: Vec<String> = contents.lines().map(|s| s.to_string()).collect();
    Ok(list)
}

pub fn handle_status_message(session: &GuiSession) -> Result<bool> {
    match session.find_by_id("wnd[0]/sbar".to_owned()) {
        Ok(SAPComponent::GuiStatusbar(status)) => {
            let text = status.text()?;
            if text.len() == 0 {
                eprintln!("No status message");
                Ok(true)
            } else {
                eprintln!("Status message: {}", text);
                Ok(false)
            }
        },
        _ => {
            eprintln!("No status bar");
            Ok(false)
        },
    }
}

// close all modal windows loop
pub fn close_all_modal_windows(session: &GuiSession) -> Result<String> {
    println!("Closing all modal windows");
    // start from 3 down
    let mut n_wnd = 3;
    while n_wnd >= 0 {
        match close_modal_window(&session, Some(n_wnd)) {
            Ok(_msg) => {
                // eprintln!("{}", msg);
                n_wnd -= 1;
            },
            Err(e) => {
                eprintln!("Error closing window: {:?}", e);
                break;
            },
        }
    }
    Ok("Closed all windows".to_string())
}

pub fn close_modal_window(session: &GuiSession, n_wnd: Option<i32>) -> Result<String> {
    let num  = match n_wnd {
        Some(n) => n,
        None => 1,
    };
    match session.find_by_id(format!("wnd[{}]", num).to_owned()) {
        Ok(SAPComponent::GuiModalWindow(wnd)) => {
            println!("Got window id: {}", wnd.id()?);
            wnd.close()?;
            Ok("Closed window".to_string())
        },
        _ => Ok("No window found".to_string()),
    }
}

pub fn start_user_tcode(session: &GuiSession, default_tcode: String, require_input: Option<bool>) -> Result<()> {
    // close any window if it exists
    match close_modal_window(&session, None) {
        Ok(msg) => eprintln!("{}", msg),
        Err(e) => eprintln!("Error closing window: {:?}", e),
    }

    // get user input for tcode to run
    let require_input = match require_input {
        Some(true) => true,
        _ => false,
    };

    let mut tcode = String::new();

    if require_input {
        println!("Enter tcode to run (default: {}):", default_tcode);
        stdin().read_line(&mut tcode).expect("Failed to read line");
        tcode = tcode.trim().to_owned();
        if tcode.is_empty() {
            tcode = default_tcode;
        }
        println!("You entered: {}", tcode);
    } else {
        tcode = default_tcode;
    }

    if let SAPComponent::GuiMainWindow(wnd) = session.find_by_id("wnd[0]".to_owned())? {
        println!("Got window id: {}", wnd.id()?);

        // starting transaction
        eprintln!("Starting transaction");
        match session.start_transaction(tcode) {
            Ok(_) => {
                eprintln!("Transaction started");
                Ok(())
            },
            Err(e) => {
                eprintln!("Error starting transaction: {:?}", e);
                Err(e)
            },
        }
    } else {
        panic!("no window!");
    }
}