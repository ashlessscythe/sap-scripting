use std::io::{self, Write};
use std::fmt::Debug;
use strum_macros::EnumIter;
use strum::IntoEnumIterator;

#[derive(EnumIter, Debug)]
enum ReadOptions {
    List,
    Single,
    Table,
    Chart,
}

pub fn prompt_enum() -> std::result::Result<ReadOptions, Box<dyn std::error::Error>> {
    let mut input = String::new();
    println!("Choose one of the following:");
    for (i, option) in ReadOptions::iter().enumerate() {
        println!("{}. {:?}", i + 1, option);
    }
    io::stdout().flush()?;
    io::stdin().read_line(&mut input)?;
    match input.trim().parse::<usize>() {
        Ok(i) if i > 0 && i <= ReadOptions::iter().count() => Ok(ReadOptions::iter().nth(i - 1).unwrap()),
        _ => Err("Invalid input".into()),
    }
}