use sap_scripting::*;
use std::env;
use std::io::{self, Write, stdout};
use std::time::Duration;
use std::thread;
use rand::Rng;
use dialoguer::Select;
use crossterm::{execute, terminal::{Clear, ClearType}};

mod utils;
mod vt11;
mod vt11_module;

use utils::*;
use utils::utils::{encrypt_data, decrypt_data, KEY_FILE_SUFFIX};
use vt11_module::run_vt11_module;

// Struct to hold login parameters
struct LoginParams {
    client_id: String,
    user: String,
    password: String,
    language: String,
}

// Convert LoginParams to ParamsStruct for use with utils functions
impl From<&LoginParams> for ParamsStruct {
    fn from(params: &LoginParams) -> Self {
        ParamsStruct {
            client_id: params.client_id.clone(),
            user: params.user.clone(),
            pass: params.password.clone(),
            language: params.language.clone(),
        }
    }
}

fn clear_screen() {
    execute!(stdout(), Clear(ClearType::All)).unwrap();
}

fn main() -> anyhow::Result<()> {
    // Initialize logging if needed
    // pretty_env_logger::init();

    // Initialize COM environment
    let com_instance = SAPComInstance::new().expect("Couldn't initialize COM environment");
    
    // Get SAP wrapper
    let wrapper = match com_instance.sap_wrapper() {
        Ok(w) => w,
        Err(e) => {
            eprintln!("Error getting SAP wrapper: {}", e);
            eprintln!("Make sure SAP GUI is installed and properly configured.");
            return Err(e.into());
        }
    };

    // Try to get the scripting engine
    let engine = match wrapper.scripting_engine() {
        Ok(e) => e,
        Err(e) => {
            eprintln!("Error getting SAP scripting engine: {}", e);
            eprintln!("Make sure SAP GUI is running and scripting is enabled.");
            return Err(e.into());
        }
    };

    // Get connection or create a new one
    let connection = match get_or_create_connection(&engine) {
        Ok(conn) => conn,
        Err(e) => {
            eprintln!("Error getting SAP connection: {}", e);
            return Err(e.into());
        }
    };
    
    // Get the first session
    let session: GuiSession = match GuiConnectionExt::children(&connection)?.element_at(0)?.downcast() {
        Some(s) => s,
        None => {
            eprintln!("Failed to get SAP session");
            return Err(windows::core::Error::from_win32().into());
        }
    };

    // Main application loop
    loop {
        clear_screen();

        // Check if already logged in
        let transaction = session.info()?.transaction()?;
        let is_logged_in = !transaction.contains("S000");
        
        // Create menu options based on login status
        let options = if is_logged_in {
            vec![
                "Log in to SAP",
                "VT11 - Shipment List Planning",
                "Log out of SAP",
                "Exit"
            ]
        } else {
            vec![
                "Log in to SAP",
                "VT11 - Shipment List Planning (Not available - Login required)",
                "Log out of SAP (Not available - Login required)",
                "Exit"
            ]
        };
        
        let choice = Select::new()
            .with_prompt("Choose an option")
            .items(&options)
            .default(0)
            .interact()
            .unwrap();
    
        match choice {
            0 => { 
                // Log in to SAP
                handle_login(&session)?;
            },
            1 => { 
                // Run VT11 module (only if logged in)
                if is_logged_in {
                    if let Err(e) = run_vt11_module(&session) {
                        eprintln!("Error running VT11 module: {}", e);
                        thread::sleep(Duration::from_secs(2));
                    }
                } else {
                    println!("You need to log in first.");
                    thread::sleep(Duration::from_secs(2));
                }
            },
            2 => { 
                // Log out of SAP (only if logged in)
                if is_logged_in {
                    if let Err(e) = handle_logout(&session) {
                        eprintln!("Error logging out: {}", e);
                        thread::sleep(Duration::from_secs(2));
                    }
                } else {
                    println!("You are not logged in.");
                    thread::sleep(Duration::from_secs(2));
                }
            },
            3 => {
                // Exit application
                clear_screen();
                println!("Exiting application...");
                return Ok(());
            },
            _ => {} // no-op
        }
    }
}

fn handle_login(session: &GuiSession) -> anyhow::Result<()> {
    // Check if already logged in
    let transaction = session.info()?.transaction()?;
    if !transaction.contains("S000") {
        println!("Already logged in. Current transaction: {}", transaction);
        thread::sleep(Duration::from_secs(2));
        // ask if refresh session (close popups)
        let options = vec!["Y", "N"];
        let choice = Select::new()
            .with_prompt("Refresh Session? (i.e. close windows, go back to main screen)")
            .items(&options)
            .default(0)
            .interact()
            .unwrap();
        match choice {
            0 => {
                close_popups(session)?;
                session.start_transaction("SESSION_MANAGER".into())?;
            },
            _ => {}     // no-op
        }
        return Ok(());
    }
    
    // Not logged in, perform login
    let login_params = get_login_parameters()?;
    login(session, &login_params)?;
    println!("Login successful!");
    thread::sleep(Duration::from_secs(2));
    
    Ok(())
}

fn handle_logout(session: &GuiSession) -> anyhow::Result<()> {
    if let Err(e) = session.end_transaction() {
        eprintln!("Error ending transaction: {}", e);
        return Err(e.into());
    }
    println!("Logged out of SAP");
    thread::sleep(Duration::from_secs(2));
    
    Ok(())
}

fn get_or_create_connection(engine: &GuiApplication) -> windows::core::Result<GuiConnection> {
    // Try to get existing connection
    if let Ok(children) = GuiApplicationExt::children(engine) {
        if children.count()? > 0 {
            if let Ok(component) = children.element_at(0) {
                if let Some(connection) = component.downcast::<GuiConnection>() {
                    return Ok(connection);
                }
            }
        }
    }
    
    // No existing connection, create a new one
    println!("No existing SAP connection found. Creating a new connection...");
    
    // In a real application, you might want to get the connection name from config
    // For now, we'll use a default or let the user specify
    let connection_name = env::var("SAP_CONNECTION_NAME")
        .unwrap_or_else(|_| "Production Instance".to_string());
    
    println!("Opening connection: {}", connection_name);
    
    // Open the connection
    let component = engine.open_connection(connection_name)?;
    
    // Convert to GuiConnection
    match component.downcast::<GuiConnection>() {
        Some(connection) => Ok(connection),
        None => {
            eprintln!("Failed to create SAP connection");
            Err(windows::core::Error::from_win32())
        }
    }
}

fn get_login_parameters() -> windows::core::Result<LoginParams> {
    // Default values
    let mut params = LoginParams {
        client_id: "025".to_string(),
        user: String::new(),
        password: String::new(),
        language: "EN".to_string(),
    };
    
    // Try to read from auth file
    let auth_path = match env::var("USERPROFILE") {
        Ok(profile) => format!("{}\\Documents\\SAP\\", profile),
        Err(_) => {
            eprintln!("Could not determine user profile directory");
            String::from(".\\")
        }
    };
    
    // Get instance ID from environment or use default
    let instance_id = env::var("SAP_INSTANCE_ID").unwrap_or_else(|_| "rs".to_string());
    let auth_file = format!("{}cryptauth_{}.txt", auth_path, instance_id);
    let key_file = format!("{}cryptauth_{}{}", auth_path, instance_id, KEY_FILE_SUFFIX);
    
    // Try to read credentials from file
    let mut ask_for_credentials = true;
    if let Ok(encrypted_data) = std::fs::read_to_string(&auth_file).map_err(|_| windows::core::Error::from_win32()) {
        if let Ok(key_data) = std::fs::read(&key_file).map_err(|_| windows::core::Error::from_win32()) {
            match decrypt_data(&encrypted_data, &key_data) {
                Ok(decrypted_data) => {
                    let lines: Vec<&str> = decrypted_data.split('\n').collect();
                    if lines.len() >= 2 {
                        params.user = lines[0].to_string();
                        params.password = lines[1].to_string();
                        ask_for_credentials = false;
                    }
                },
                Err(_) => {
                    eprintln!("Failed to decrypt credentials. They may be corrupted or using an old format.");
                }
            }
        }
    }
    
    // If credentials not found in file, ask user
    if ask_for_credentials {
        println!("Please enter your SAP credentials:");
        
        if params.user.is_empty() {
            print!("Username: ");
            io::stdout().flush().unwrap();
            io::stdin().read_line(&mut params.user).unwrap();
            params.user = params.user.trim().to_string();
        } else {
            println!("Username: {} (from saved credentials)", params.user);
        }
        
        if params.password.is_empty() {
            print!("Password: ");
            io::stdout().flush().unwrap();
            // In a real application, you would use a crate like rpassword to hide input
            io::stdin().read_line(&mut params.password).unwrap();
            params.password = params.password.trim().to_string();
        }
        
        // Ask if user wants to save credentials
        print!("Save credentials for future use? (y/n): ");
        io::stdout().flush().unwrap();
        let mut save_choice = String::new();
        io::stdin().read_line(&mut save_choice).unwrap();
        
        if save_choice.trim().to_lowercase() == "y" {
            save_credentials(&auth_path, &auth_file, &key_file, &params.user, &params.password)?;
        }
    }
    
    Ok(params)
}

fn login(session: &GuiSession, params: &LoginParams) -> windows::core::Result<()> {
    println!("Logging in to SAP...");
    
    // Find and fill client field
    if let Ok(client_field) = session.find_by_id("wnd[0]/usr/txtRSYST-MANDT".to_string()) {
        if let Some(text_field) = client_field.downcast::<GuiTextField>() {
            text_field.set_text(params.client_id.clone())?;
        }
    }
    
    // Find and fill username field
    if let Ok(user_field) = session.find_by_id("wnd[0]/usr/txtRSYST-BNAME".to_string()) {
        if let Some(text_field) = user_field.downcast::<GuiTextField>() {
            text_field.set_text(params.user.clone())?;
        }
    }
    
    // Find and fill password field
    if let Ok(pass_field) = session.find_by_id("wnd[0]/usr/pwdRSYST-BCODE".to_string()) {
        if let Some(password_field) = pass_field.downcast::<GuiPasswordField>() {
            password_field.set_text(params.password.clone())?;
        }
    }
    
    // Find and fill language field
    if let Ok(lang_field) = session.find_by_id("wnd[0]/usr/txtRSYST-LANGU".to_string()) {
        if let Some(text_field) = lang_field.downcast::<GuiTextField>() {
            text_field.set_text(params.language.clone())?;
        }
    }
    
    // Press Enter button
    if let Ok(enter_button) = session.find_by_id("wnd[0]/tbar[0]/btn[0]".to_string()) {
        if let Some(button) = enter_button.downcast::<GuiButton>() {
            button.press()?;
        }
    }
    
    // Wait a bit for login to process
    thread::sleep(Duration::from_millis(1000));
    
    // Check for multiple logon popup
    if let Ok(popup) = session.find_by_id("wnd[1]".to_string()) {
        if let Ok(popup_text) = popup.r_type() {
            if popup_text.contains("GuiModalWindow") {
                // Check if it's a multiple logon popup
                if let Ok(radio_button) = session.find_by_id("wnd[1]/usr/radMULTI_LOGON_OPT1".to_string()) {
                    if let Some(rb) = radio_button.downcast::<GuiRadioButton>() {
                        rb.select()?;
                        rb.set_focus()?;
                        
                        // Press Enter
                        if let Ok(window) = session.find_by_id("wnd[1]".to_string()) {
                            if let Some(modal_window) = window.downcast::<GuiModalWindow>() {
                                modal_window.send_v_key(0)?;
                            }
                        }
                    }
                }
            }
        }
    }
    
    // Check for error messages in status bar
    if let Ok(statusbar) = session.find_by_id("wnd[0]/sbar".to_string()) {
        if let Some(status) = statusbar.downcast::<GuiStatusbar>() {
            let message = status.text()?;

            match message {
                msg if msg.contains("incorrect") => {
                    eprintln!("Login failed: {}", msg);
                    return Err(windows::core::Error::from_win32())
                }
                msg if msg.contains("new password") => {
                    eprintln!("Password update required: {}", msg);
                    return Err(windows::core::Error::from_win32())
                }
                
                msg if msg.contains("exist") => {
                    eprintln!("Incorrect info or something not found");
                    return Err(windows::core::Error::from_win32())
                }

                _ => {} // no-op
            }
        }
    }
    
    // Close any remaining popups
    if let Ok(popup) = session.find_by_id("wnd[1]".to_string()) {
        if let Some(window) = popup.downcast::<GuiFrameWindow>() {
            window.close()?;
        }
    }
    
    Ok(())
}

fn save_credentials(auth_path: &str, auth_file: &str, key_file: &str, username: &str, password: &str) -> windows::core::Result<()> {
    // Create directory if it doesn't exist
    std::fs::create_dir_all(auth_path).map_err(|_| windows::core::Error::from_win32())?;
    
    // Generate or load encryption key
    let key = if let Ok(existing_key) = std::fs::read(key_file) {
        existing_key
    } else {
        // Generate a new random key
        let mut key = vec![0u8; 32]; // 256 bits for AES-256
        rand::thread_rng().fill(&mut key[..]);
        
        // Save the key
        std::fs::write(key_file, &key).map_err(|_| windows::core::Error::from_win32())?;
        key
    };
    
    // Encrypt and save credentials
    let content = format!("{}\n{}", username, password);
    match encrypt_data(&content, &key) {
        Ok(encrypted) => {
            std::fs::write(auth_file, encrypted).map_err(|_| windows::core::Error::from_win32())?;
            println!("Encrypted credentials saved to {}", auth_file);
            Ok(())
        },
        Err(_) => {
            eprintln!("Failed to encrypt credentials");
            Err(windows::core::Error::from_win32())
        }
    }
}
