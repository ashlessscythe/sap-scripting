use sap_scripting::*;

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
        _ => panic!("expected modal window, but got something else!"),
    }
}

pub fn get_grid_value(session: &GuiSession, col_header: &str) -> SAPComponent {
    match session.find_by_id("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell".to_owned()) {
        Ok(SAPComponent::GuiGridView(grid)) => {
            println!("got grid with {:?} rows and {:?} columns", grid.row_count(), grid.column_count());
            let col = grid.get_column_position(col_header.to_owned()).expect("column not found");
            println!("cell 0,0: {:?}", grid.get_cell_value(0, col_header.to_owned()).unwrap());
            SAPComponent::GuiGridView(grid)
        },
        _ => panic!("expected grid, but got something else!"),
    }
}