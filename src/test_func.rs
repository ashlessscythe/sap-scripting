use sap_scripting::*;
use crate::func::*;
use std::fs;

pub fn run_test_tcode(session: GuiSession) -> Result<()> {
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
    send_vkey_main(&w1, 8)?;
    handle_status_message(&session)?;

    // get values from table
    let headers = vec![
        "VBELN".to_owned(),
        "CHILD".to_owned(),
        "PALLET".to_owned(),
    ];
    let mut full = Vec::new();
    headers.iter().for_each(|h| {
        let vals = match get_grid_values(&session, h) {
            Ok(vals) => {
                // write to file
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
    // write full to csv
    let mut csv = String::new();
    for i in 0..full[0].len() {
        for j in 0..full.len() {
            csv.push_str(&full[j][i]);
            csv.push(',');
        }
        csv.push('\n');
    }
    // header row with newline
    let header_row = format!("{},\n", headers.join(","));
    csv.insert_str(0, &header_row);
    let file_name = format!("{}.csv", lvalue(Some(tcode.to_owned())));
    fs::write(file_name, csv).expect("Error writing to file");

    Ok(())

}
