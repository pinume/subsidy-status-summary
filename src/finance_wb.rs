use std::collections::HashMap;
use std::path::Path;
use calamine::Data;
use rust_decimal::Decimal;
use rust_decimal::prelude::ToPrimitive;
use rust_xlsxwriter::{Workbook, Worksheet};

use crate::reader::{HeaderMap, cell_to_decimal, cell_to_string, read_sheet_rows};
use crate::styles::StylePool;
use crate::summary_wb::write_refund_cell;

/// Generate Workbook 2: 26年国补门店财务统筹表.xlsx (4 Sheets)
pub fn generate_store_finance_workbook(input_dir: &Path, output_path: &Path) -> Result<(), String> {
    let mut workbook = Workbook::new();
    let styles = StylePool::new();

    // 1. Load source data
    let store_occ_rows = read_sheet_rows(&input_dir.join("银联交易明细门店.xlsx"))?;
    let all_union_rows = read_sheet_rows(&input_dir.join("银联交易明细所有.xlsx"))?;
    let app_upload_rows = read_sheet_rows(&input_dir.join("已上传家电电脑.xlsx"))?;
    let dig_upload_rows = read_sheet_rows(&input_dir.join("已上传数码.xlsx"))?;
    let app_refund_rows = read_sheet_rows(&input_dir.join("回款明细家电电脑.xlsx"))?;
    let dig_refund_rows = read_sheet_rows(&input_dir.join("回款明细数码.xlsx"))?;
    let invoice_rows = read_sheet_rows(&input_dir.join("发票明细.xlsx"))?;

    // Extract store merchant codes
    let app_upload_h = HeaderMap::from_header_row(&app_upload_rows[0]);
    let dig_upload_h = HeaderMap::from_header_row(&dig_upload_rows[0]);
    let app_mch_idx = app_upload_h.find(&["商户号"]).ok_or("已上传家电电脑缺少商户号")?;
    let dig_mch_idx = dig_upload_h.find(&["商户号"]).ok_or("已上传数码缺少商户号")?;

    let app_store_code = cell_to_string(&app_upload_rows[1][app_mch_idx]);
    let dig_store_code = cell_to_string(&dig_upload_rows[1][dig_mch_idx]);

    // Build Sheet 1: 最终匹配表（全部的国补发生数据上匹配）
    build_final_match_sheet(
        &mut workbook,
        &styles,
        &store_occ_rows,
        &all_union_rows,
        &app_upload_rows,
        &dig_upload_rows,
        &invoice_rows,
    )?;

    // Build Sheet 2: 1.门店国补发生表（银联系统直接导出，不需要加工）
    build_store_occurrence_sheet(&mut workbook, &styles, &store_occ_rows)?;

    // Build Sheet 3: 2.门店累计回款表（事业部财务下发，门店筛选自己的）
    build_store_refund_sheet(
        &mut workbook,
        &styles,
        &app_refund_rows,
        &dig_refund_rows,
        &app_store_code,
        &dig_store_code,
    )?;

    // Build Sheet 4: 3.门店上传明细（从门店银联后台每月导出后汇总）
    build_store_upload_sheet(&mut workbook, &styles, &app_upload_rows, &dig_upload_rows)?;

    workbook
        .save(output_path)
        .map_err(|e| format!("保存工作簿二失败: {}", e))?;

    Ok(())
}

fn build_final_match_sheet(
    wb: &mut Workbook,
    s: &StylePool,
    store_occ: &[Vec<Data>],
    all_union: &[Vec<Data>],
    app_up: &[Vec<Data>],
    dig_up: &[Vec<Data>],
    invoices: &[Vec<Data>],
) -> Result<(), String> {
    let ws = wb.add_worksheet();
    ws.set_name("最终匹配表（全部的国补发生数据上匹配）").map_err(|e| e.to_string())?;
    ws.set_freeze_panes(1, 0).map_err(|e| e.to_string())?;

    // 27 columns widths
    let widths = [
        8.0, 20.0, 20.0, 14.0, 12.0, 22.0, 14.0, 14.0, 12.0, 16.0, 18.0, 12.0, 16.0, 18.0, 24.0,
        16.0, 28.0, 28.0, 14.0, 16.0, 14.0, 14.0, 16.0, 26.0, 16.0, 24.0, 14.0,
    ];
    for (col, w) in widths.iter().enumerate() {
        ws.set_column_width(col as u16, *w).map_err(|e| e.to_string())?;
    }

    // Headers
    let headers = [
        "序号", "清算时间", "交易时间", "终端号", "交易类型", "卡号", "交易金额", "清算金额", "手续费",
        "流水号", "检索号", "卡类型", "发卡行", "商户号", "商户名称", "分店简称", "商户订单号",
        "银商订单号", "交易方式", "分店", "优惠金额", "大类", "品牌", "产品型号", "状态", "发票号",
        "发票是否红冲",
    ];

    ws.set_row_height(0, 30.0).map_err(|e| e.to_string())?;
    for (col, h) in headers.iter().enumerate() {
        let fmt = if col == 0 || col == 21 || col == 24 || col == 26 {
            &s.col_header_center
        } else {
            &s.col_header
        };
        ws.write_string_with_format(0, col as u16, *h, fmt).map_err(|e| e.to_string())?;
    }

    // 1. Build Index for 银联交易明细所有 (84,907 rows)
    let all_h = HeaderMap::from_header_row(&all_union[0]);
    let all_ref_idx = all_h.find(&["交易参考号"]).ok_or("银联所有缺少交易参考号")?;
    let all_ord1_idx = all_h.find(&["银商订单号"]).ok_or("银联所有缺少银商订单号")?;
    let all_ord2_idx = all_h.find(&["银商原订单号"]);
    let all_cat_idx = all_h.find(&["商品类别"]).ok_or("银联所有缺少商品类别")?;
    let all_brand_idx = all_h.find(&["品牌"]).ok_or("银联所有缺少品牌")?;
    let all_model_idx = all_h.find(&["商品型号"]).ok_or("银联所有缺少商品型号")?;

    struct ProductInfo {
        cat: String,
        brand: String,
        model: String,
    }

    let mut ref_map: HashMap<String, ProductInfo> = HashMap::new();
    let mut order_map: HashMap<String, ProductInfo> = HashMap::new();

    for row in &all_union[1..] {
        let cat = cell_to_string(&row[all_cat_idx]);
        let brand = cell_to_string(&row[all_brand_idx]);
        let model = cell_to_string(&row[all_model_idx]);

        let ref_no = cell_to_string(&row[all_ref_idx]);
        if !ref_no.is_empty() && !ref_map.contains_key(&ref_no) {
            ref_map.insert(ref_no, ProductInfo {
                cat: cat.clone(),
                brand: brand.clone(),
                model: model.clone(),
            });
        }

        let ord1 = cell_to_string(&row[all_ord1_idx]);
        if !ord1.is_empty() && !order_map.contains_key(&ord1) {
            order_map.insert(ord1, ProductInfo {
                cat: cat.clone(),
                brand: brand.clone(),
                model: model.clone(),
            });
        }

        if let Some(idx2) = all_ord2_idx {
            let ord2 = cell_to_string(&row[idx2]);
            if !ord2.is_empty() && !order_map.contains_key(&ord2) {
                order_map.insert(ord2, ProductInfo { cat, brand, model });
            }
        }
    }

    // 2. Build Index for 门店上传明细 (appliance + digital)
    struct UploadInfo {
        status: String,
        invoice_no: String,
    }
    let mut upload_map: HashMap<String, UploadInfo> = HashMap::new();

    let index_upload = |rows: &[Vec<Data>], map: &mut HashMap<String, UploadInfo>| {
        let h = HeaderMap::from_header_row(&rows[0]);
        let ref_idx = h.find(&["检索参考号"]).unwrap();
        let st_idx = h.find(&["状态"]).unwrap();
        let inv_idx = h.find(&["发票号码"]).unwrap();

        for row in &rows[1..] {
            let r_no = cell_to_string(&row[ref_idx]);
            if !r_no.is_empty() {
                map.insert(r_no, UploadInfo {
                    status: cell_to_string(&row[st_idx]),
                    invoice_no: cell_to_string(&row[inv_idx]),
                });
            }
        }
    };
    index_upload(app_up, &mut upload_map);
    index_upload(dig_up, &mut upload_map);

    // 3. Build Index for 发票明细
    let inv_h = HeaderMap::from_header_row(&invoices[0]);
    let inv_no_idx = inv_h.find(&["数电发票号码"]).ok_or("发票明细缺少数电发票号码")?;
    let inv_type_idx = inv_h.find(&["开票类型"]).ok_or("发票明细缺少开票类型")?;
    let inv_st_idx = inv_h.find(&["开票状态"]).ok_or("发票明细缺少开票状态")?;

    struct InvoiceStatus {
        inv_type: String,
        inv_status: String,
    }
    let mut invoice_map: HashMap<String, InvoiceStatus> = HashMap::new();
    for row in &invoices[1..] {
        let no = cell_to_string(&row[inv_no_idx]);
        if !no.is_empty() {
            invoice_map.insert(no, InvoiceStatus {
                inv_type: cell_to_string(&row[inv_type_idx]),
                inv_status: cell_to_string(&row[inv_st_idx]),
            });
        }
    }

    // 4. Pre-resolve store occurrence column mappings for final match
    let store_h = HeaderMap::from_header_row(&store_occ[0]);
    let s_ref_idx = store_h.find(&["检索号"]).ok_or("门店发生表缺少检索号")?;
    let s_ord_idx = store_h.find(&["银商订单号"]).ok_or("门店发生表缺少银商订单号")?;
    let s_rem_idx = store_h.find(&["备注"]).ok_or("门店发生表缺少备注")?;

    let store_col_map: [(&str, u16); 20] = [
        ("清算时间", 1),
        ("交易时间", 2),
        ("终端号", 3),
        ("交易类型", 4),
        ("卡号", 5),
        ("交易金额", 6),
        ("清算金额", 7),
        ("手续费", 8),
        ("流水号", 9),
        ("检索号", 10),
        ("卡类型", 11),
        ("发卡行", 12),
        ("商户号", 13),
        ("商户名称", 14),
        ("分店简称", 15),
        ("商户订单号", 16),
        ("银商订单号", 17),
        ("交易方式", 18),
        ("分店", 19),
        ("优惠金额", 20),
    ];

    let resolved_store_cols: Vec<(usize, u16)> = store_col_map
        .iter()
        .filter_map(|(name, target_col)| {
            store_h.find(&[name]).map(|src_idx| (src_idx, *target_col))
        })
        .collect();

    for (r_idx, row) in store_occ.iter().enumerate().skip(1) {
        let cur_row = r_idx as u32; // Row 1 is header, data starts at row 1
        ws.set_row_height(cur_row, 22.0).map_err(|e| e.to_string())?;

        // Col 0: 序号 (1-based)
        ws.write_number_with_format(cur_row, 0, r_idx as f64, &s.int_count).map_err(|e| e.to_string())?;

        // Col 1-20: Pre-resolved from store occurrence
        for &(src_idx, target_col) in &resolved_store_cols {
            if src_idx < row.len() {
                let cell = &row[src_idx];
                match target_col {
                    1 | 2 => {
                        let val = cell_to_string(cell);
                        ws.write_string_with_format(cur_row, target_col, val, &s.datetime).map_err(|e| e.to_string())?;
                    }
                    6 | 7 | 8 | 20 => {
                        let val = cell_to_decimal(cell).unwrap_or(Decimal::ZERO).to_f64().unwrap_or(0.0);
                        ws.write_number_with_format(cur_row, target_col, val, &s.money_yuan).map_err(|e| e.to_string())?;
                    }
                    3 | 4 => {
                        let val = cell_to_string(cell);
                        ws.write_string_with_format(cur_row, target_col, val, &s.text_center).map_err(|e| e.to_string())?;
                    }
                    _ => {
                        let val = cell_to_string(cell);
                        ws.write_string_with_format(cur_row, target_col, val, &s.text_left).map_err(|e| e.to_string())?;
                    }
                }
            }
        }

        // Col 21, 22, 23: 大类, 品牌, 产品型号
        let ref_num = cell_to_string(&row[s_ref_idx]);
        let ord_num = cell_to_string(&row[s_ord_idx]);

        let prod = ref_map.get(&ref_num).or_else(|| order_map.get(&ord_num));
        let (cat, brand, model) = match prod {
            Some(p) => (p.cat.as_str(), p.brand.as_str(), p.model.as_str()),
            None => ("", "", ""),
        };
        ws.write_string_with_format(cur_row, 21, cat, &s.text_center).map_err(|e| e.to_string())?;
        ws.write_string_with_format(cur_row, 22, brand, &s.text_left).map_err(|e| e.to_string())?;
        ws.write_string_with_format(cur_row, 23, model, &s.text_left).map_err(|e| e.to_string())?;

        // Col 24: 状态
        let remark = cell_to_string(&row[s_rem_idx]);
        let up_info = upload_map.get(&ref_num);

        let status = if remark == "已退货" {
            "已退货"
        } else if let Some(info) = up_info {
            info.status.as_str()
        } else {
            "未提交"
        };
        ws.write_string_with_format(cur_row, 24, status, &s.text_center).map_err(|e| e.to_string())?;

        // Col 25: 发票号
        let invoice_no = match up_info {
            Some(info) if !info.invoice_no.is_empty() => info.invoice_no.as_str(),
            _ => "",
        };
        ws.write_string_with_format(cur_row, 25, invoice_no, &s.text_left).map_err(|e| e.to_string())?;

        // Col 26: 发票是否红冲
        let red_flush = if invoice_no.is_empty() {
            ""
        } else if let Some(inv) = invoice_map.get(invoice_no) {
            if inv.inv_type == "红票" || inv.inv_status == "已红冲" {
                "是"
            } else if inv.inv_type == "蓝票" && inv.inv_status == "开票完成" {
                "否"
            } else {
                ""
            }
        } else {
            ""
        };
        ws.write_string_with_format(cur_row, 26, red_flush, &s.text_center).map_err(|e| e.to_string())?;
    }

    Ok(())
}

fn build_store_occurrence_sheet(
    wb: &mut Workbook,
    s: &StylePool,
    store_occ: &[Vec<Data>],
) -> Result<(), String> {
    let ws = wb.add_worksheet();
    ws.set_name("1.门店国补发生表（银联系统直接导出，不需要加工）").map_err(|e| e.to_string())?;
    ws.set_freeze_panes(1, 0).map_err(|e| e.to_string())?;

    // 26 columns widths
    let widths = [
        20.0, 20.0, 14.0, 12.0, 22.0, 14.0, 14.0, 12.0, 12.0, 12.0, 16.0, 18.0, 12.0, 16.0, 18.0,
        24.0, 16.0, 28.0, 28.0, 14.0, 16.0, 14.0, 14.0, 14.0, 16.0, 18.0,
    ];
    for (col, w) in widths.iter().enumerate() {
        ws.set_column_width(col as u16, *w).map_err(|e| e.to_string())?;
    }

    // Row 1: Header (26 columns straight from source)
    ws.set_row_height(0, 30.0).map_err(|e| e.to_string())?;
    for col in 0..26 {
        let name = cell_to_string(&store_occ[0][col]);
        ws.write_string_with_format(0, col as u16, name, &s.col_header).map_err(|e| e.to_string())?;
    }

    // Data rows
    for (r_idx, row) in store_occ.iter().enumerate().skip(1) {
        let cur_row = r_idx as u32;
        ws.set_row_height(cur_row, 22.0).map_err(|e| e.to_string())?;

        for col in 0..26 {
            let cell = if col < row.len() { &row[col] } else { &Data::Empty };
            match col {
                // DateTime: 0:清算时间, 1:交易时间
                0 | 1 => {
                    let val = cell_to_string(cell);
                    ws.write_string_with_format(cur_row, col as u16, val, &s.datetime).map_err(|e| e.to_string())?;
                }
                // Center strings: 2:终端号, 3:交易类型
                2 | 3 => {
                    let val = cell_to_string(cell);
                    ws.write_string_with_format(cur_row, col as u16, val, &s.text_center).map_err(|e| e.to_string())?;
                }
                // Money: 5:交易金额, 6:清算金额, 7:手续费, 8:T0手续费, 9:D1手续费, 21:优惠金额, 22:分期手续费
                5 | 6 | 7 | 8 | 9 | 21 | 22 => {
                    if let Some(dec) = cell_to_decimal(cell) {
                        ws.write_number_with_format(cur_row, col as u16, dec.to_f64().unwrap_or(0.0), &s.money_yuan).map_err(|e| e.to_string())?;
                    } else {
                        ws.write_string_with_format(cur_row, col as u16, "", &s.text_right).map_err(|e| e.to_string())?;
                    }
                }
                _ => {
                    let val = cell_to_string(cell);
                    ws.write_string_with_format(cur_row, col as u16, val, &s.text_left).map_err(|e| e.to_string())?;
                }
            }
        }
    }

    Ok(())
}

fn build_store_refund_sheet(
    wb: &mut Workbook,
    s: &StylePool,
    app_refund: &[Vec<Data>],
    dig_refund: &[Vec<Data>],
    app_store_code: &str,
    dig_store_code: &str,
) -> Result<(), String> {
    let ws = wb.add_worksheet();
    ws.set_name("2.门店累计回款表（事业部财务下发，门店筛选自己的）").map_err(|e| e.to_string())?;
    ws.set_freeze_panes(1, 0).map_err(|e| e.to_string())?;

    let widths = [
        32.0, 20.0, 16.0, 26.0, 26.0, 22.0, 18.0, 14.0, 14.0, 14.0, 14.0, 12.0, 24.0, 14.0, 16.0,
        12.0, 16.0, 36.0, 14.0, 24.0, 18.0, 24.0, 48.0, 18.0,
    ];
    for (col, w) in widths.iter().enumerate() {
        ws.set_column_width(col as u16, *w).map_err(|e| e.to_string())?;
    }

    // Header (24 columns from source)
    ws.set_row_height(0, 30.0).map_err(|e| e.to_string())?;
    for col in 0..24 {
        let name = cell_to_string(&app_refund[0][col]);
        ws.write_string_with_format(0, col as u16, name, &s.col_header).map_err(|e| e.to_string())?;
    }

    let mut cur_row: u32 = 1;

    let append_rows = |ws: &mut Worksheet,
                       rows: &[Vec<Data>],
                       store_code: &str,
                       cur_row: &mut u32|
     -> Result<(), String> {
        let h = HeaderMap::from_header_row(&rows[0]);
        let mch_idx = h.find(&["核销商编"]).ok_or("回款明细缺少核销商编")?;

        for row in &rows[1..] {
            if cell_to_string(&row[mch_idx]) == store_code {
                ws.set_row_height(*cur_row, 22.0).map_err(|e| e.to_string())?;
                for col in 0..24 {
                    let cell = if col < row.len() { &row[col] } else { &Data::Empty };
                    write_refund_cell(ws, s, *cur_row, col as u16, cell)?;
                }
                *cur_row += 1;
            }
        }
        Ok(())
    };

    append_rows(ws, app_refund, app_store_code, &mut cur_row)?;
    append_rows(ws, dig_refund, dig_store_code, &mut cur_row)?;

    Ok(())
}

fn build_store_upload_sheet(
    wb: &mut Workbook,
    s: &StylePool,
    app_up: &[Vec<Data>],
    dig_up: &[Vec<Data>],
) -> Result<(), String> {
    let ws = wb.add_worksheet();
    ws.set_name("3.门店上传明细（从门店银联后台每月导出后汇总）").map_err(|e| e.to_string())?;
    ws.set_freeze_panes(1, 0).map_err(|e| e.to_string())?;

    // 61 column headers and widths as defined in Section 5.5.2
    let col_defs: [(&str, f64); 61] = [
        ("实时清分UUID", 28.0),
        ("商户号", 18.0),
        ("商户名称", 24.0),
        ("订单号", 28.0),
        ("交易日期", 14.0),
        ("交易金额", 14.0),
        ("检索参考号", 18.0),
        ("模版类型", 12.0),
        ("状态", 14.0),
        ("描述", 32.0),
        ("提交时间", 20.0),
        ("更新时间", 20.0),
        ("终端号", 14.0),
        ("分店id", 26.0),
        ("分店名", 18.0),
        ("所在地区", 20.0),
        ("详细地址", 30.0),
        ("地区编码", 12.0),
        ("tel", 16.0),
        ("发票号码", 24.0),
        ("发票金额", 14.0),
        ("购买方名称", 16.0),
        ("图片1", 26.0),
        ("S/N码", 24.0),
        ("是否属于 AI 产品", 16.0),
        ("IMEI1", 20.0),
        ("IMEI2", 20.0),
        ("图片1", 26.0),
        ("图片2", 26.0),
        ("图片3", 26.0),
        ("图片4", 26.0),
        ("img5", 26.0),
        ("img6", 26.0),
        ("img7", 26.0),
        ("img8", 26.0),
        ("img9", 26.0),
        ("img10", 26.0),
        ("img11", 26.0),
        ("img12", 26.0),
        ("img13", 26.0),
        ("img14", 26.0),
        ("img15", 26.0),
        ("签收时间", 20.0),
        ("remark", 20.0),
        ("EEG", 16.0),
        ("物流单号", 20.0),
        ("erpOrderNum", 20.0),
        ("ocrModify", 12.0),
        ("modifyStatus", 14.0),
        ("introduceInvoiceFlag", 18.0),
        ("是否交旧", 12.0),
        ("是否自提", 12.0),
        ("receiverName", 16.0),
        ("productCode", 18.0),
        ("subsideAmt", 14.0),
        ("productName", 28.0),
        ("交旧品类", 16.0),
        ("收货地址是否农村地区", 20.0),
        ("airConditionerKitInfo", 22.0),
        ("开票日期", 14.0),
        ("补贴金额", 14.0),
    ];

    for (col, (_, w)) in col_defs.iter().enumerate() {
        ws.set_column_width(col as u16, *w).map_err(|e| e.to_string())?;
    }

    // Row 1: Header
    ws.set_row_height(0, 30.0).map_err(|e| e.to_string())?;
    for (col, (name, _)) in col_defs.iter().enumerate() {
        let fmt = if [4, 7, 8, 10, 11, 12, 17, 24, 42, 47, 48, 49, 50, 51, 57, 59].contains(&col) {
            &s.col_header_center
        } else {
            &s.col_header
        };
        ws.write_string_with_format(0, col as u16, *name, fmt).map_err(|e| e.to_string())?;
    }

    // Pre-resolve column mappings for Appliance and Digital
    let app_h = HeaderMap::from_header_row(&app_up[0]);
    let dig_h = HeaderMap::from_header_row(&dig_up[0]);

    // Build 61-column index extractor parameterized by category
    let resolve_indices = |h: &HeaderMap, is_dig: bool| -> Vec<Option<usize>> {
        (0..61)
            .map(|col| match col {
                22 => h.find_occurrence("图片1", 1),
                25 | 26 => if is_dig { h.find(&[col_defs[col].0]) } else { None },
                27 => h.find_occurrence("图片1", 2),
                31 => h.find(&["img5", "图片5"]),
                32 => h.find(&["img6", "图片6"]),
                44 => if is_dig { None } else { h.find(&["EEG"]) },
                56 => h.find(&["交旧品类", "oldExchangeType"]),
                58 => if is_dig { None } else { h.find(&["airConditionerKitInfo"]) },
                _ => h.find(&[col_defs[col].0]),
            })
            .collect()
    };

    let app_map = resolve_indices(&app_h, false);
    let dig_map = resolve_indices(&dig_h, true);

    let mut cur_row: u32 = 1;

    let write_dataset = |ws: &mut Worksheet,
                         rows: &[Vec<Data>],
                         col_map: &[Option<usize>],
                         cur_row: &mut u32|
     -> Result<(), String> {
        for row in &rows[1..] {
            ws.set_row_height(*cur_row, 22.0).map_err(|e| e.to_string())?;
            for col in 0..61 {
                let cell = match col_map[col] {
                    Some(src_col) if src_col < row.len() => &row[src_col],
                    _ => &Data::Empty,
                };

                match col {
                    // Date: 4:交易日期, 59:开票日期
                    4 | 59 => {
                        let val = cell_to_string(cell);
                        ws.write_string_with_format(*cur_row, col as u16, val, &s.date).map_err(|e| e.to_string())?;
                    }
                    // DateTime: 10:提交时间, 11:更新时间, 42:签收时间
                    10 | 11 | 42 => {
                        let val = cell_to_string(cell);
                        ws.write_string_with_format(*cur_row, col as u16, val, &s.datetime).map_err(|e| e.to_string())?;
                    }
                    // Money: 5:交易金额, 20:发票金额, 54:subsideAmt, 60:补贴金额
                    5 | 20 | 54 | 60 => {
                        if let Some(dec) = cell_to_decimal(cell) {
                            ws.write_number_with_format(*cur_row, col as u16, dec.to_f64().unwrap_or(0.0), &s.money_yuan).map_err(|e| e.to_string())?;
                        } else {
                            ws.write_string_with_format(*cur_row, col as u16, "", &s.text_right).map_err(|e| e.to_string())?;
                        }
                    }
                    // Center text: 7:模版类型, 8:状态, 12:终端号, 17:地区编码, 24:是否属于 AI 产品, 47:ocrModify, 48:modifyStatus, 49:introduceInvoiceFlag, 50:是否交旧, 51:是否自提, 57:收货地址是否农村地区
                    7 | 8 | 12 | 17 | 24 | 47 | 48 | 49 | 50 | 51 | 57 => {
                        let val = cell_to_string(cell);
                        ws.write_string_with_format(*cur_row, col as u16, val, &s.text_center).map_err(|e| e.to_string())?;
                    }
                    _ => {
                        let val = cell_to_string(cell);
                        ws.write_string_with_format(*cur_row, col as u16, val, &s.text_left).map_err(|e| e.to_string())?;
                    }
                }
            }
            *cur_row += 1;
        }
        Ok(())
    };

    write_dataset(ws, app_up, &app_map, &mut cur_row)?;
    write_dataset(ws, dig_up, &dig_map, &mut cur_row)?;

    Ok(())
}
