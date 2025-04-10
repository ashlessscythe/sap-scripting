use sap_scripting::*;
use std::time::Duration;
use std::thread;
use windows::core::Result;
use crate::utils::close_popups;
use crate::utils::sap_constants::STR_FORM;
use crate::utils::sap_ctrl_utils::hit_ctrl;

pub fn assert_tcode(session: &GuiSession, tcode: &str, wnd: Option<i32>) -> Result<bool> {
    let wnd_num = wnd.unwrap_or(0);
    
    // Start the transaction
    session.start_transaction(tcode.to_string())?;
    
    // Check for errors in status bar
    let err_msg = hit_ctrl(session, wnd_num, "/sbar", "Text", "Get", "")?;
    
    if err_msg.contains("exist") || err_msg.contains("autho") {
        println!("Error: {}", err_msg);
        return Ok(false);
    }
    
    if err_msg.is_empty() {
        return Ok(true);
    }
    
    // Log error message
    println!("{}{}{}", STR_FORM, err_msg, STR_FORM);
    
    Ok(false)
}

pub fn check_tcode(session: &GuiSession, tcode: &str, run: Option<bool>, _kill_popups: Option<bool>) -> Result<bool> {
    let run_val = run.unwrap_or(false);
    let b_kill_popups  = _kill_popups.unwrap_or(false);
    
    println!("Checking if tCode ({}) is active", tcode);
    
    match b_kill_popups {
        true => {
            close_popups(session)?;
        },
        _ => {}
    }
    
    // Get current transaction
    let current = session.info()?.transaction()?;
    
    // Check if on tCode
    if current.contains(tcode) {
        println!("tCode ({}) is active", tcode);
        return Ok(true);
    } else if run_val {
        // Run if requested
        println!("tCode mismatch, attempting to run tCode ({})", tcode);
        let _ = assert_tcode(session, tcode, None)?;
        thread::sleep(Duration::from_millis(500)); // Time_Event equivalent
        
        // Recursive call to check again
        return check_tcode(session, tcode, Some(false), Some(false));
    } else {
        println!("tCode mismatch. Current tCode is ({}), need ({})", current, tcode);
        return Ok(false);
    }
}
