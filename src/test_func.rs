use sap_scripting::*;
use crate::func::*;
use std::{collections::btree_map::Keys, fs};
use std::sync::{Arc, Mutex};
use std::thread;
use tokio::task;
use std::marker::{Sync, Send};

use rayon::prelude::*;
use std::result::Result;

struct SafeGuiSession {
    inner: Mutex<GuiSession>,
}

impl SafeGuiSession {
    fn new(session: GuiSession) -> Self {
        SafeGuiSession {
            inner: Mutex::new(session),
        }
    }
}

unsafe impl Sync for SafeGuiSession {}
unsafe impl Send for SafeGuiSession {}

// parallel
pub fn parallel_tcodes(conn: &GuiConnection, starting_idx: i32) -> Result<(), String> {
    let idx = starting_idx + 1;

    let s1 = SafeGuiSession::new(create_or_get_session(conn, idx).map_err(|e| e.to_string())?);
    let s2 = SafeGuiSession::new(create_or_get_session(conn, idx + 1).map_err(|e| e.to_string())?);
    let s3 = SafeGuiSession::new(create_or_get_session(conn, idx + 2).map_err(|e| e.to_string())?);

    let sessions = vec![s1, s2, s3];
    let tcodes = vec!["zmdesnr".to_string(), "zmdetpc".to_string(), "mmbe".to_string()];

    let arc_sessions = Arc::new(sessions);

    // Process each session in parallel
    let results: Vec<_> = arc_sessions.par_iter().zip(tcodes.par_iter()).map(|(session, tcode)| {
        match session.inner.lock() {
            Ok(guard) => match test_single_tcode(&*guard, tcode) {
                Ok(_) => Ok(format!("Transaction successful for {}", tcode)),
                Err(e) => Err(format!("Error executing transaction for {}: {}", tcode, e)),
            },
            Err(_) => Err("Mutex is poisoned".to_string()),
        }
    }).collect();

    // Check results and decide on the outcome
    results.into_iter().find(|result| result.is_err()).unwrap_or(Ok("error, in results?".to_string())).map(|_| ())

}

// test_single_tcode
pub fn test_single_tcode(session: &GuiSession, tcode: &String) -> Result<(), String> {
    let (_, t) = start_user_tcode(session, tcode.to_owned(), Some(false)).expect("Error starting tcode");
        
    let info = get_session_info(session).expect("Error getting session info");
    println!("Current tcode: {:?}", info.transaction());
    println!("Current screen: {:?}", info.screen_number());
    println!("Current user: {:?}", info.user());
        
    let w1 = get_main_window(session).expect("Error getting main window");

    // if tcode changed, do something
    if &t != tcode {
        println!("tcode changed from {} to {}", tcode, t);
        println!("finding by name");
    };

    // ask execute?
    // prompt_execute(&w1, 8)?;
    send_vkey_main(&w1, 8).expect("error handling message");
    handle_status_message(&session).expect("error handling message");

    send_vkey_main(&w1, 8).expect("error handling message");
    handle_status_message(&session).expect("error handling message");

    Ok(())
}

// start async
pub async fn multi_tcode(conn: Arc<GuiConnection>, starting_idx: i32) -> Result<(), String> {
    let s1 = Arc::new(create_or_get_session_async(&conn, 1).await.expect("Error creating session"));
    let s2 = Arc::new(create_or_get_session_async(&conn, 2).await.expect("Error creating session"));

    loop {

        let con1 = conn.clone();
        let con2 = conn.clone();
        run_test_tcode(s1.clone(), con1, starting_idx + 1).await.expect("Error running tcode");
        run_test_tcode(s2.clone(), con2, starting_idx + 2).await.expect("Error running tcode");

        // ask user if redo
        let redo = prompt_bool("Redo?").expect("Error getting input");
        if !redo {
            break;
        }
    }
    Ok(())
}

pub async fn run_test_tcode(session: Arc<GuiSession>, con: Arc<GuiConnection>, idx: i32) -> Result<(), String> {
    let tcode = "zmdetpc";
    let (_, t) = start_user_tcode(&session, tcode.to_owned(), Some(false)).expect("Error starting tcode");
        
    let info = get_session_info(&session).expect("Error getting session info");
    println!("Current tcode: {:?}", info.transaction());
    println!("Current screen: {:?}", info.screen_number());
    println!("Current user: {:?}", info.user());
        
    let w1 = get_main_window(&session).expect("Error getting main window");

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
    // let variant_name = prompt_str("Variant name (default blank)").expect("failed to get input");
    let variant_name = "tpc_yesterday".to_owned();
    apply_variant(&session, &variant_name).expect("Error applying variant");

    // ask execute?
    // prompt_execute(&w1, 8)?;
    send_vkey_main(&w1, 8).expect("error handling message");
    handle_status_message(&*session).expect("error handling message");

    // get values from table
    let headers = vec![
        "VBELN".to_owned(),
        "CHILD".to_owned(),
        "PALLET".to_owned(),
    ];
    let full = get_grid_values_for_headers(&session, &headers).expect("Error getting grid values");
    save_grid_to_csv(full.clone(), &headers, &tcode, None).expect("Error saving grid to csv");

    // create another session to lookup each value
    // read from csv column 1
    let list = &full[0].clone();
    println!("list has {} items", list.len());

    // create another session
    let s2 = create_or_get_session_async(&con.to_owned(), idx + 1).await.expect("Error creating session");
    let w2 = get_main_window(&s2).expect("Error getting main window");

    // let start = prompt_bool("Start tcodes?").expect("Error getting input");
    let start = true;
    if start {
        // iterate through first 10
        list.iter().take(3).for_each(|v| {
            match start_user_tcode(&s2, "VL03N".to_owned(), Some(false)) {
                Ok(_) => {
                    handle_status_message(&s2).expect("Error handling status message");
                    eprintln!("Tcode started: {}", v);
                    set_ctext_main(&s2, &"/usr/ctxtLIKP-VBELN".to_owned(), &v).expect("Error setting text");
                    send_vkey_main(&w2, 0).expect("Error sending vkey");
                    let tab_id = r#"/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01"#.to_owned();
                    let tab = get_tab(&s2, &tab_id).expect("Error getting tab");
                    tab.select().expect("Error selecting tab");
                    println!("Selected tab");
                    // count of packages
                    let packages_field = r#"/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/txtLIKP-ANZPK"#.to_owned();
                    let packages = get_text_main(&s2, &packages_field).expect("Error getting text");
                    println!("Packages for {} is {}", v, packages);
                    // header
                    press_tbar_button_main(&s2, 1, 8).expect("Error pressing button");   // hat button
                    let header_tab = r#"/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07"#.to_owned();
                    get_tab(&s2, &header_tab).expect("Error getting tab").select().expect("Error selecting tab");
                    println!("Selected header tab");
                    let text_box = r#"/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV50A:2110/txtLIKP-LIFEX"#.to_owned();
                    let dlv = get_text_main(&s2, &text_box.to_owned()).expect("Error getting text");
                    println!("delivery {} ext {}", v, dlv);
                },
                Err(e) => eprintln!("Error starting tcode: {:?}", e),
            }
        });
    } // end start

    // go back
    back_screen(&w1).expect("Error going back");
    back_screen(&w2).expect("Error going back");

    Ok(())

}
