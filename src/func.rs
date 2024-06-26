use sap_scripting::*;
use std::io::stdin;
use std::io::{self, Write};
use csv::Reader;
use chrono::prelude::*;
use rpassword;
use std::sync::Arc;

pub fn lvalue(s: Option<String>) -> String {
    // timestamp with s as suffix
    let s = match s {
        Some(s) => s,
        None => "".to_string(),
    };
    let now: DateTime<Local> = Local::now();
    let ts = now.format("%Y%m%d%H%M%S").to_string();
    format!("{}-{}", ts, s)
}

// go back
pub fn back_screen(wnd: &GuiMainWindow) -> Result<()> {
    for _ in 0..5 {
        send_vkey_main(wnd, 3).expect("Error sending vkey");
    }
    Ok(())
}

pub fn prompt_pass(prompt: &str) -> std::result::Result<String, Box<dyn std::error::Error>> {
    print!("{}: ", prompt);
    io::stdout().flush()?; // Make sure the prompt is immediately displayed
    let pass = rpassword::read_password()?;
    Ok(pass)
}

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

// get list from csv specific column
pub fn get_list_from_csv(file_path: &str, col: usize) -> Result<Vec<String>> {
    let mut rdr = csv::Reader::from_path(file_path).expect("Error reading csv");
    let mut list = Vec::new();
    for result in rdr.records() {
        let record = result.expect("Error reading record");
        list.push(record[col].to_string());
    }
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

pub fn print_sap_component_type(component: &SAPComponent) {
    match component {
        SAPComponent::GuiTabStrip(_) => println!("found tabstrip"),
        SAPComponent::GuiGridView(_) => println!("found gridview"),
        SAPComponent::GuiUserArea(_) => println!("found user area"),
        SAPComponent::GuiVComponent(_) => println!("found v component"),
        SAPComponent::GuiComponent(_) => println!("found component"),
        // Add more match arms for other variants of SAPComponent
        _ => println!("found unknown component"),
    }
}

// start threaded tcode
pub async fn start_tcode_async(session: &GuiSession, tcode: &String) -> Result<()> {
    // close any window if it exists
    match close_modal_window(&session, None) {
        Ok(msg) => eprintln!("{}", msg),
        Err(e) => eprintln!("Error closing window: {:?}", e),
    }

    // close modal windows if found
    close_all_modal_windows(session)?;

    // starting transaction
    eprintln!("Starting transaction");
    match session.start_transaction(tcode.to_owned()) {
        Ok(_) => {
            eprintln!("Transaction started");
            Ok(())
        },
        Err(e) => {
            eprintln!("Error starting transaction: {:?}", e);
            Err(e)
        },
    }
}

// tcode wit no user input
pub fn start_tcode(session: &GuiSession, tcode: String) -> Result<()> {
    // close any window if it exists
    match close_modal_window(&session, None) {
        Ok(msg) => eprintln!("{}", msg),
        Err(e) => eprintln!("Error closing window: {:?}", e),
    }

    // close modal windows if found
    close_all_modal_windows(session)?;

    // starting transaction
    eprintln!("Starting transaction");
    match session.start_transaction(tcode.clone()) {
        Ok(_) => {
            eprintln!("Transaction started");
            Ok(())
        },
        Err(e) => {
            eprintln!("Error starting transaction: {:?}", e);
            Err(e)
        },
    }
}
// tcode with user input
pub fn start_user_tcode(session: &GuiSession, default_tcode: String, require_input: Option<bool>) -> Result<((), String)> {
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

    // close modal windows if found
    close_all_modal_windows(session)?;

    // starting transaction
    eprintln!("Starting transaction");
    match session.start_transaction(tcode.clone()) {
        Ok(_) => {
            eprintln!("Transaction started");
            Ok(((), tcode))
        },
        Err(e) => {
            eprintln!("Error starting transaction: {:?}", e);
            Err(e)
        },
    }
}

// apply user variant
pub fn apply_variant(session: &GuiSession, var: &str) -> Result<()> {
    let wnd = get_main_window(session).expect("Error getting main window");

    // if blank var, return
    if var.is_empty() {
        println!("No variant specified, skipping...");
        return Ok(());
    } else {
        println!("Applying variant: {}", var);
    }

    // send variant key
    send_vkey_main(&wnd, 17)?;

    // get modal window
    let modal_window = get_modal_window(session, &1).expect("Error getting modal window");

    // clear prev text
    set_text_modal(&modal_window, &1,"/usr/txtENAME-LOW", "")?;

    // enter variant name
    set_text_modal(&modal_window, &1, "/usr/txtV-LOW", var)?;

    // send f8 (close modal window)
    send_vkey_modal(&modal_window, 8)?;
    handle_status_message(session)?;
    close_modal_window(session, None)?;

    Ok(())
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

// set password field
pub fn set_pass_main(session: &GuiSession, field_id: &str, text: &str) -> Result<()> {
    let find_id = format!("wnd[0]{}", field_id); // field id should include the /usr/ prefix
    match session.find_by_id(find_id.to_owned()) {
        Ok(SAPComponent::GuiPasswordField(pass)) => {
            match pass.set_text(text.to_owned()) {
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
            Ok(())
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

// press button
pub fn press_tbar_button_main(session: &GuiSession, tbar: i32, btn: i32) -> Result<()> {
    let find_id = format!("wnd[0]/tbar[{}]/btn[{}]", tbar, btn);
    match session.find_by_id(find_id.to_owned()) {
        Ok(SAPComponent::GuiButton(btn)) => {
            match btn.press() {
                Ok(_) => {
                    eprintln!("Button pressed: {}", find_id);
                    Ok(())
                },
                Err(e) => {
                    eprintln!("Error pressing button {}: {:?}", find_id, e);
                    Err(e)
                }
            }
        },
        _ => {
            eprintln!("No button found with ID {}", find_id);
            Ok(())
        }
    }
}

// get main wnd
pub fn get_main_window(session: &GuiSession) -> std::result::Result<GuiMainWindow, String> {
    match session.find_by_id("wnd[0]".to_owned()) {
        Ok(SAPComponent::GuiMainWindow(wnd)) => Ok(wnd),
        _ => Err("expected main window, but got something else!".to_owned()),
    }
}

// get modal wnd
pub fn get_modal_window(session: &GuiSession, wnd_id: &i32) -> std::result::Result<GuiModalWindow, String> {
    match session.find_by_id(format!("wnd[{}]", wnd_id).to_owned()) {
        Ok(SAPComponent::GuiModalWindow(wnd)) => Ok(wnd),
        _ => Err("expected modal window, but got something else!".to_owned()),
    }
}

// get tab
pub fn get_tab(session: &GuiSession, tab_id: &str) -> std::result::Result<GuiTab, String> {
    match session.find_by_id(format!("wnd[0]{}", tab_id.to_owned())) {
        Ok(SAPComponent::GuiTab(tab)) => {
            println!("Got tab");
            Ok(tab)
        },
        _ => Err("expected tab, but got something else!".to_owned()),
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

// get text from id
pub fn get_text_main(session: &GuiSession, field_id: &str) -> std::result::Result<String, String> {
    let find_id = format!("wnd[0]{}", field_id); // field id should include the /usr/ prefix
    println!("getting text from id: {}", find_id);
    match session.find_by_id(find_id.to_owned()) {
        Ok(SAPComponent::GuiTextField(txt)) => {
            println!("Got text field");
            let text = sap_scripting::GuiTextField_Impl::displayed_text(&txt)
                .map_err(|_| "Error getting text".to_string())?;
            Ok(text)
        },
        _ => Err("No text field found".to_string()),
    }
}

// async create or get session
pub async fn create_or_get_session_async(connection: &Arc<GuiConnection>, idx: i32) -> std::result::Result<GuiSession, String> {
    let children = sap_scripting::GuiConnection_Impl::children(&**connection)
        .map_err(|e| format!("Failed to retrieve children: {:?}", e))?; // get or create session @ idx
    match children.element_at(idx) {
        Ok(SAPComponent::GuiSession(sess)) => {
            println!("Returning existing session at index {}", idx);
            Ok(sess)  // Assuming GuiSession is cloneable. If not, you'll need to adjust this.
        },
        _ => {
            // If no session at that index, attempt to create a new session
            let sess = match sap_scripting::GuiConnection_Impl::children(&**connection).unwrap().element_at(0) {
                Ok(SAPComponent::GuiSession(sess)) => sess,
                _ => panic!("Error getting initial session"),
            };
            sap_scripting::GuiSession_Impl::create_session(&sess)
                .map_err(|e| format!("Failed to create a new session: {:?}", e))?; // wait for new session to be created
            std::thread::sleep(std::time::Duration::from_secs(3)); // Assume new session is created at the end of the list
            let new_children = sap_scripting::GuiConnection_Impl::children(&**connection)
                .map_err(|e| format!("Failed to retrieve children after session creation: {:?}", e))?;
            println!("New children count: {}", new_children.count().unwrap());
            match new_children.element_at(new_children.count().unwrap() - 1) {
                Ok(SAPComponent::GuiSession(new_sess)) => {
                    println!("Returning new session created at index {}", new_children.count().unwrap() - 1);
                    Ok(new_sess)  // Adjust if GuiSession is not cloneable
                },
                _ => Err("Failed to retrieve new session".to_owned()),
            }
        }
    }
}

// get or create session @ idx
pub fn create_or_get_session(connection: &GuiConnection, idx: i32) -> std::result::Result<GuiSession, String> {
    let children = sap_scripting::GuiConnection_Impl::children(connection)
        .map_err(|e| format!("Failed to retrieve children: {:?}", e))?;

    // Try to get the session at the specified index
    match children.element_at(idx) {
        Ok(SAPComponent::GuiSession(sess)) => {
            println!("Returning existing session at index {}", idx);
            Ok(sess)  // Assuming GuiSession is cloneable. If not, you'll need to adjust this.
        },
        _ => {
            // If no session at that index, attempt to create a new session
            let sess = match sap_scripting::GuiConnection_Impl::children(connection).unwrap().element_at(0) {
                Ok(SAPComponent::GuiSession(sess)) => sess,
                _ => panic!("Error getting initial session"),
            };
            sap_scripting::GuiSession_Impl::create_session(&sess)
                .map_err(|e| format!("Failed to create a new session: {:?}", e))?;

            // wait for new session to be created
            std::thread::sleep(std::time::Duration::from_secs(3));

            // Assume new session is created at the end of the list
            let new_children = sap_scripting::GuiConnection_Impl::children(connection)
                .map_err(|e| format!("Failed to retrieve children after session creation: {:?}", e))?;

            println!("New children count: {}", new_children.count().unwrap());

            match new_children.element_at(new_children.count().unwrap() - 1) {
                Ok(SAPComponent::GuiSession(new_sess)) => {
                    println!("Returning new session created at index {}", new_children.count().unwrap() - 1);
                    Ok(new_sess)  // Adjust if GuiSession is not cloneable
                },
                _ => Err("Failed to retrieve new session".to_owned()),
            }
        }
    }
}

// log in
pub fn log_in_sap(session: &GuiSession) -> Result<bool> {
    let wnd = match session.find_by_id("wnd[0]".to_owned()) {
        Ok(SAPComponent::GuiMainWindow(session)) => session,
        _ => panic!("expected main window, but got something else!"),
    };
    // login ask for user and pass
    println!("Not logged in, please enter user and pass");
    let user = prompt_str("User").expect("Error getting input");
    let pass = prompt_pass("Password").expect("Error getting input");
    set_text_main(session, "/usr/txtRSYST-MANDT", &"025".to_owned())?;
    set_text_main(session, "/usr/txtRSYST-BNAME", &user)?;
    set_pass_main(session, "/usr/pwdRSYST-BCODE", &pass)?;
    set_text_main(session, "/usr/txtRSYST-LANGU", &"EN".to_owned())?;
    send_vkey_main(&wnd, 0).expect("Error sending vkey");
    let logged_in = match handle_status_message(session) {
        Ok(true) => {
            eprintln!("Logged in");
            true
        },
        Ok(false) => {
            eprintln!("Error logging in");
            false
        },
        Err(e) => {
            eprintln!("Error logging in: {:?}", e);
            false
        },
    };
    Ok(logged_in)
}


// get grid values for headers to vec
pub fn get_grid_values_for_headers(session: &GuiSession, headers: &[String]) -> Result<Vec<Vec<String>>> {
    let mut full = Vec::new();
    headers.iter().for_each(|h| {
        let vals = match get_grid_values(session, h) {
            Ok(vals) => {
                println!("found {} values for header {}.", vals.len(), h);
                vals
            },
            Err(e) => {
                eprintln!("Error getting values: {:?}", e);
                vec![]
            }
        };
        full.push(vals);
        println!("Vals count as of header {} is {}", h, full.last().unwrap().len());
    });
    Ok(full)
}

// save grid to csv
pub fn save_grid_to_csv(data: Vec<Vec<String>>, headers: &[String], tcode: &str, filename: Option<String>) -> Result<String> {
    // write full to csv
    let mut csv = String::new();
    for i in 0..data[0].len() {
        for j in 0..data.len() {
            csv.push_str(&data[j][i]);
            csv.push(',');
        }
        csv.push('\n');
    }

    // header row with newline
    let header_row = format!("{},\n", headers.join(","));
    csv.insert_str(0, &header_row);
    let file_name = match filename {
        Some(f) => f,
        None => format!("{}.csv", lvalue(Some(tcode.to_owned()))),
    };
    std::fs::write(&file_name, csv).expect("Error writing to file");

    // close file
    println!("Wrote to file: {}", &file_name);

    Ok(file_name)
}