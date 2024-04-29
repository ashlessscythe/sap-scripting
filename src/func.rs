use sap_scripting::*;
use std::io::stdin;
use std::io::{self, Write};

pub fn prompt_bool(prompt: &str) -> std::result::Result<bool, Box<dyn std::error::Error>> {
    let mut input = String::new();
    print!("{} (y/n): ", prompt);
    io::stdout().flush()?; // Make sure the prompt is immediately displayed
    io::stdin().read_line(&mut input)?;
    match input.trim().to_lowercase().as_str() {
        "y" | "yes" | "" => Ok(true),
        "n" | "no" => Ok(false),
        _ => Ok(false),
    }
}

// get user input for number
pub fn prompt_number(prompt: &str) -> std::result::Result<i32, Box<dyn std::error::Error>> {
    let mut input = String::new();
    print!("{}: ", prompt);
    io::stdout().flush()?; // Make sure the prompt is immediately displayed
    io::stdin().read_line(&mut input)?;
    let num = input.trim().parse::<i32>()?;
    Ok(num)
}

// get user str
pub fn prompt_str(prompt: &str) -> std::result::Result<String, Box<dyn std::error::Error>> {
    let mut input = String::new();
    print!("{}: ", prompt);
    io::stdout().flush()?; // Make sure the prompt is immediately displayed
    io::stdin().read_line(&mut input)?;
    Ok(input.trim().to_owned())
}

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
        _ => Ok("No modal window found".to_string()),
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

pub fn prompt_execute(wnd: &GuiMainWindow, vkey: i16) -> Result<()> {
    let execute = prompt_bool("Execute?").expect("failed to get input");

    if execute {
        // send vkey
        send_vkey_main(wnd, vkey)?;
    }

    Ok(())
}

// get session info
pub fn get_session_info(session: &GuiSession) -> Result<GuiSessionInfo> {
    match session.info() {
        Ok(info) => {
            // eprintln!("Got session info");
            Ok(info)
        },
        Err(e) => {
            eprintln!("Error getting session info: {:?}", e);
            Err(e)
        }
    }
}

// set modal window text
pub fn set_text_modal(wnd_obj: &GuiModalWindow, wnd_id: &i32, field_id: &str, text: &str) -> Result<()> {
    let search = format!("wnd[{}]{}", wnd_id, field_id); // field id should include the /usr/ prefix
    match wnd_obj.find_by_id(search.to_owned()) {
        Ok(SAPComponent::GuiTextField(txt)) => {
            match txt.set_text(text.to_owned()) {
                Ok(_) => {
                    eprintln!("Text set for field: {}", field_id);
                    Ok(())
                },
                Err(e) => {
                    eprintln!("Error setting text for field {}: {:?}", search, e);
                    Err(e)
                }
            }
        },
        _ => {
            eprintln!("No text field found with ID {}", search);
            Ok(())
        }
    }
}


// set gui text
pub fn set_text_main(session: &GuiSession, field_id: &str, text: &str) -> Result<()> {
    let find_id = format!("wnd[0]{}", field_id); // field id should include the /usr/ prefix
    match session.find_by_id(find_id.to_owned()) {
        Ok(SAPComponent::GuiTextField(txt)) => {
            match txt.set_text(text.to_owned()) {
                Ok(_) => {
                    eprintln!("Text set for field: {}", field_id);
                    Ok(())
                },
                Err(e) => {
                    eprintln!("Error setting text for field {}: {:?}", field_id, e);
                    Err(e)

                }
            }
        },
        _ => {
            eprintln!("No text field found with ID {}", find_id);
            Ok(())
        }
    }
}

pub fn set_ctext_main(session: &GuiSession, field_id: &str, text: &str) -> Result<()> {
    let find_id = format!("wnd[0]{}", field_id); // field id should include the /usr/ prefix
    match session.find_by_id(find_id.to_owned()) {
        Ok(SAPComponent::GuiCTextField(ctxt)) => {
            match ctxt.set_text(text.to_owned()) {
                Ok(_) => {
                    eprintln!("Text set for field: {}", field_id);
                    Ok(())
                },
                Err(e) => {
                    eprintln!("Error setting text for field {}: {:?}", field_id, e);
                    Err(e)

                }
            }
        },
        _ => {
            eprintln!("No text field found with ID {}", field_id);
            Ok(())
        }
    }
}

pub fn send_vkey_main(wnd: &GuiMainWindow, key: i16) -> Result<()> {
    match wnd.send_v_key(key) {
        Ok(_) => {
            eprintln!("Key {} sent successfully", key);
            Ok(())
        },
        Err(e) => {
            eprintln!("Error sending key {}: {:?}", key, e);
            Err(e)
        }
    }
}

pub fn send_vkey_modal(wnd: &GuiModalWindow, key: i16) -> Result<()> {
    match wnd.send_v_key(key) {
        Ok(_) => {
            eprintln!("Key {} sent successfully", key);
            Ok(())
        },
        Err(e) => {
            eprintln!("Error sending key {}: {:?}", key, e);
            Err(e)
        }
    }
}

// get wnd
pub fn get_modal_window(session: &GuiSession, wnd_id: &i32) -> std::result::Result<GuiModalWindow, String> {
    match session.find_by_id(format!("wnd[{}]", wnd_id).to_owned()) {
        Ok(SAPComponent::GuiModalWindow(wnd)) => Ok(wnd),
        _ => Err("expected modal window, but got something else!".to_owned()),
    }
}

pub fn get_grid(session: &GuiSession) -> std::result::Result<GuiGridView, String>{
    match session.find_by_id("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell".to_owned()) {
        Ok(SAPComponent::GuiGridView(grid)) => {
            eprint!("Got grid\n");
            Ok(grid)
        },
        _ => Err("expected grid, but got something else!".to_owned()),
    }
}

pub fn get_grid_values(session: &GuiSession, hdr: &str) -> std::result::Result<Vec<String>, String> {
    match get_grid(&session) {
        Ok(grid) => {
            let row_count = sap_scripting::GuiGridView_Impl::row_count(&grid)
                .map_err(|_| "Error getting row count".to_string())?;
            let mut vals = Vec::new();
            for i in 0..row_count {
                let val = sap_scripting::GuiGridView_Impl::get_cell_value(&grid, i, hdr.to_owned())
                    .map_err(|_| "Error getting cell value".to_string())?;
                vals.push(val);
            }
            Ok(vals)
        },
        Err(e) => Err(format!("Error getting grid: {:?}", e)),
    }
}

// create a new session
pub fn create_session(connection: &GuiConnection, session: &GuiSession, idx: i32) -> std::result::Result<GuiSession, String> {
    // Attempt to create a session using the connection
    let creation_result = sap_scripting::GuiSession_Impl::create_session(session);
    
    // Handle the result of the session creation attempt
    match creation_result {
        Ok(_) => {
            // If the session is successfully created, attempt to retrieve it
            // Assuming the new session is always added at the end
            let children = sap_scripting::GuiConnection_Impl::children(connection)
                .map_err(|e| format!("Failed to retrieve children after session creation: {:?}", e))?;
            
            // Get the new session
            let new_session = match children.element_at(idx) {
                Ok(SAPComponent::GuiSession(session)) => Ok(session),
                _ => return Err("Failed to retrieve new session".to_owned()),
            };
            return new_session;
        },
        Err(e) => Err(format!("Failed to create a new session: {:?}", e)),
    }
}
