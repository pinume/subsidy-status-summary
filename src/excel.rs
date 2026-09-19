use calamine::Data;
use rust_decimal::prelude::ToPrimitive;
use rust_xlsxwriter::Worksheet;

use crate::reader::{cell_to_decimal, cell_to_string};
use crate::styles::StylePool;

pub fn write_refund_cell(
    ws: &mut Worksheet,
    s: &StylePool,
    row: u32,
    col: u16,
    cell: &Data,
) -> Result<(), String> {
    match col {
        1 => {
            let val = cell_to_string(cell);
            if val.is_empty() {
                ws.write_string_with_format(row, col, "", &s.text_center)
                    .map_err(|e| e.to_string())?;
            } else {
                ws.write_string_with_format(row, col, val, &s.datetime)
                    .map_err(|e| e.to_string())?;
            }
        }
        8 | 9 | 10 | 18 => {
            if let Some(dec) = cell_to_decimal(cell) {
                ws.write_number_with_format(row, col, dec.to_f64().unwrap_or(0.0), &s.money)
                    .map_err(|e| e.to_string())?;
            } else {
                ws.write_string_with_format(row, col, "", &s.text_right)
                    .map_err(|e| e.to_string())?;
            }
        }
        11 => {
            if let Some(dec) = cell_to_decimal(cell) {
                ws.write_number_with_format(row, col, dec.to_f64().unwrap_or(0.0), &s.percent)
                    .map_err(|e| e.to_string())?;
            } else {
                ws.write_string_with_format(row, col, "", &s.text_right)
                    .map_err(|e| e.to_string())?;
            }
        }
        _ => {
            let val = cell_to_string(cell);
            ws.write_string_with_format(row, col, val, &s.text_left)
                .map_err(|e| e.to_string())?;
        }
    }
    Ok(())
}
