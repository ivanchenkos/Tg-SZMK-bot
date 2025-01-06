use calamine::{
    open_workbook,
    Data, 
    DataType, 
    Range, 
    Reader, 
    Xlsx
};
use regex::{
    self, 
    Regex
};
use std::fmt::Display;

enum ParseResult {
    ParseList(Vec<String>),
    ParseSingle(String),
}

pub struct Cable {
    pub index: String,
    pub cable_type: CableType,
}


#[derive(Debug)]
pub enum CableType {
    Magistral,
    Transit,
    Local,
}

impl Display for CableType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match &self {
            CableType::Local => {
               write!(f, "Local")
            }
            CableType::Transit => {
                write!(f, "Transit")
            }
             CableType::Magistral => {
                write!(f, "Magistral")
            }
        }
    }
}

struct ParseRowResult {
    index: String,
    column_long: Option<Vec<String>>,
    column_short_first: Option<String>,
    column_short_second: Option<String>,
}

impl ParseRowResult {
    fn new() -> Self {
        ParseRowResult {
            index: String::new(),
            column_long: None,
            column_short_first: None,
            column_short_second: None,
        }
    }
}

pub fn get_sheet(file_name: &str, sheet_name: &str) -> Result<Range<Data>, String> {
    let mut excel_file: Xlsx<_> = match open_workbook(file_name) {
        Ok(excel_file) => excel_file,
        Err(_) => return Err("Can't open file".to_string())
    };

    match excel_file.worksheet_range(sheet_name) {
        Ok(sheet) => return Ok(sheet),
        Err(_) => return Err("Can't open worksheet".to_string())
    };
    
}

// parse chosen column of row
fn parse_column(row: &[Data], column: usize) -> Result<ParseResult, String> {
    match column {
        9 =>  {
            let column_value = match row[column].get_string() {
                Some(value) => value,
                None => return Err("Empty".to_string()),    
            };
            let regex_pattern = Regex::new(r"[ ();]").unwrap();
            let mut rooms: Vec<String> = regex_pattern.split(column_value).map(|m|m.to_string()).collect();

            rooms.retain(|s| s.len() == 5);

            Ok(ParseResult::ParseList(rooms))
        }
        11 | 14 => {
            match &row[column] {
                Data::Float(value) => {
                    return Ok(ParseResult::ParseSingle(value.to_string()))
                }
                Data::Int(value) => {
                    return Ok(ParseResult::ParseSingle(value.to_string()))
                }
                Data::String(value) => {
                    return Ok(ParseResult::ParseSingle(value.to_string()))
                }
                _ => return Err("Empty".to_string()),
                
            }
        }
        _ => Err("Wrong column".to_string())
    }
}

fn parse_row(row: &[Data]) -> ParseRowResult {
    let mut parse_row_result = ParseRowResult::new();


    parse_row_result.index = match row[0].get_string() {
        Some(value) => value.to_string(),
        None => String::from("None"),
    };
    
    match parse_column(row, 9) {
        Ok(ParseResult::ParseList(rooms)) => {
            parse_row_result.column_long = Some(rooms);
        }
        _ => {
            parse_row_result.column_long = None;
        }
    }

    match parse_column(row, 11) {
        Ok(ParseResult::ParseSingle(room)) => {
            parse_row_result.column_short_first = Some(room);
        }
        _ => {
            parse_row_result.column_short_first = None;
        }
    }

    match parse_column(row, 14) {
        Ok(ParseResult::ParseSingle(room)) => {
            parse_row_result.column_short_second = Some(room);
        }
        _ => {
            parse_row_result.column_short_second = None;
        }
    }

    parse_row_result
}

pub fn get_all_cables_in_room(room: String, sheet: &Range<Data>) -> Result<Vec<Cable>, &str> {
    let mut new_cables_list = Vec::new();
    let mut room_cutted: String = String::new();

    if room.chars().nth(0).unwrap() == '0' {
        room_cutted = room[1..].to_string();
    }
    else {
        room_cutted = room.clone();
    }
     
    println!("{:?}", room_cutted);

    for row in sheet.rows() {
        match row[2].get_string() {
            Some(value) => {
                if value == "ПМЛ" {
                    continue;
                }
            },
            None => (),
        }

        let build_result = parse_row(row);
        
        let is_long = match build_result.column_long {
            Some(column_value) => {
                column_value.contains(&room)
            },
            None => false
        };   
        let is_short_first = match &build_result.column_short_first {
            Some(column_value) => {
                // some rooms in xlsx file starts with 0 (ex. 07A10), so need to check 
                if column_value.chars().count() == 5 {
                    column_value == &room    
                } else {
                    column_value == &room_cutted    
                }
                
            },
            None => false
        };
        let is_short_second = match &build_result.column_short_second {
            Some(column_value) => {
                // some rooms in xlsx file starts with 0 (ex. 07A10), so need to check 
                if column_value.chars().count() == 5 {
                    column_value == &room    
                } else {
                    column_value == &room_cutted    
                }
            },
            None => false
        };

        if is_long && (!is_short_first && !is_short_second) {
            new_cables_list.push(Cable {
                index: build_result.index.clone(),
                cable_type: CableType::Transit,
            });
        }
        else if is_long && (is_short_first || is_short_second) {
            new_cables_list.push(Cable {
                index: build_result.index.clone(),
                cable_type: CableType::Magistral
            });
        }
        else if !is_long && (is_short_first || is_short_second) {
            new_cables_list.push(Cable {
                index: build_result.index.clone(),
                cable_type: CableType::Local
            })
        }
    }
    if new_cables_list.len() <= 0 {
        // super::write_log("Совпадений не найдено");
        println!("test");
        Err("Совпадений не найдено")
    } else {
        Ok(new_cables_list)    
    }
    
}

