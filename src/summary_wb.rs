use std::collections::HashMap;
use std::path::Path;
use calamine::Data;
use rust_decimal::Decimal;
use rust_decimal::prelude::ToPrimitive;
use rust_xlsxwriter::{Workbook, Worksheet};

use crate::reader::{HeaderMap, cell_to_decimal, cell_to_string, get_colored_row_indices, read_sheet_rows};
use crate::styles::StylePool;

/// Generate Workbook 1: 国补上传情况汇总.xlsx (5 Sheets)
pub fn generate_summary_workbook(input_dir: &Path, output_path: &Path) -> Result<(), String> {
    let mut workbook = Workbook::new();
    let styles = StylePool::new();

    // 1. Load source data
    let sales_rows = read_sheet_rows(&input_dir.join("销售用券情况统计.xlsx"))?;
    let app_upload_rows = read_sheet_rows(&input_dir.join("已上传家电电脑.xlsx"))?;
    let dig_upload_rows = read_sheet_rows(&input_dir.join("已上传数码.xlsx"))?;
    let invoice_rows = read_sheet_rows(&input_dir.join("发票明细.xlsx"))?;
    let app_refund_path = input_dir.join("回款明细家电电脑.xlsx");
    let dig_refund_path = input_dir.join("回款明细数码.xlsx");

    // Extract store merchant codes from uploaded files
    let app_upload_h = HeaderMap::from_header_row(&app_upload_rows[0]);
    let dig_upload_h = HeaderMap::from_header_row(&dig_upload_rows[0]);
    let app_mch_idx = app_upload_h.find(&["商户号"]).ok_or("已上传家电电脑缺少商户号")?;
    let dig_mch_idx = dig_upload_h.find(&["商户号"]).ok_or("已上传数码缺少商户号")?;
    
    let app_store_code = cell_to_string(&app_upload_rows[1][app_mch_idx]);
    let dig_store_code = cell_to_string(&dig_upload_rows[1][dig_mch_idx]);

    // Build Sheet 1: 汇总
    build_summary_sheet(&mut workbook, &styles, &sales_rows, &app_upload_rows, &dig_upload_rows)?;

    // Build Sheet 2: 品类品牌汇总
    build_category_brand_sheet(&mut workbook, &styles, &sales_rows)?;

    // Build Sheet 3: 审核失败明细
    build_failed_records_sheet(&mut workbook, &styles, &app_upload_rows, &dig_upload_rows, &invoice_rows)?;

    // Build Sheet 4: 异常回款明细
    build_refund_anomaly_sheet(
        &mut workbook,
        &styles,
        &app_refund_path,
        &dig_refund_path,
        &app_store_code,
        &dig_store_code,
    )?;

    // Build Sheet 5: 异常发票明细
    build_invoice_anomaly_sheet(&mut workbook, &styles, &input_dir.join("发票明细.xlsx"), &invoice_rows)?;

    workbook
        .save(output_path)
        .map_err(|e| format!("保存工作簿一失败: {}", e))?;

    Ok(())
}

fn build_summary_sheet(
    wb: &mut Workbook,
    s: &StylePool,
    sales: &[Vec<Data>],
    app_up: &[Vec<Data>],
    dig_up: &[Vec<Data>],
) -> Result<(), String> {
    let ws = wb.add_worksheet();
    ws.set_name("汇总").map_err(|e| e.to_string())?;

    ws.set_column_width(0, 38.0).map_err(|e| e.to_string())?;
    for col in 1..=6 {
        ws.set_column_width(col, 20.0).map_err(|e| e.to_string())?;
    }

    // 1. Metrics aggregation
    let sales_h = HeaderMap::from_header_row(&sales[0]);
    let cat_idx = sales_h.find(&["财务大类"]).ok_or("销售表缺少财务大类")?;
    let sub_idx = sales_h.find(&["补贴额"]).ok_or("销售表缺少补贴额")?;
    let qty_idx = sales_h.find(&["数量"]).ok_or("销售表缺少数量")?;

    let mut app_gen_amt = Decimal::ZERO;
    let mut app_gen_cnt = Decimal::ZERO;
    let mut dig_gen_amt = Decimal::ZERO;
    let mut dig_gen_cnt = Decimal::ZERO;

    for row in &sales[1..] {
        let cat = cell_to_string(&row[cat_idx]);
        let amt = cell_to_decimal(&row[sub_idx]).unwrap_or(Decimal::ZERO);
        let qty = cell_to_decimal(&row[qty_idx]).unwrap_or(Decimal::ZERO);

        if cat == "数码" {
            dig_gen_amt += amt;
            dig_gen_cnt += qty;
        } else {
            app_gen_amt += amt;
            app_gen_cnt += qty;
        }
    }

    // Aggregation helper for uploaded data
    let count_uploaded = |rows: &[Vec<Data>]| -> (HashMap<String, Decimal>, HashMap<String, i64>) {
        let h = HeaderMap::from_header_row(&rows[0]);
        let status_idx = h.find(&["状态"]).unwrap();
        let sub_idx = h.find(&["补贴金额"]).unwrap();

        let mut amt_map = HashMap::new();
        let mut cnt_map = HashMap::new();

        for row in &rows[1..] {
            let st = cell_to_string(&row[status_idx]);
            let amt = cell_to_decimal(&row[sub_idx]).unwrap_or(Decimal::ZERO);
            *amt_map.entry(st.clone()).or_insert(Decimal::ZERO) += amt;
            *cnt_map.entry(st).or_insert(0) += 1;
        }
        (amt_map, cnt_map)
    };

    let (app_amt, app_cnt) = count_uploaded(app_up);
    let (dig_amt, dig_cnt) = count_uploaded(dig_up);

    let get_val = |map: &HashMap<String, Decimal>, k: &str| *map.get(k).unwrap_or(&Decimal::ZERO);
    let get_cnt = |map: &HashMap<String, i64>, k: &str| *map.get(k).unwrap_or(&0);

    let app_paid_amt = get_val(&app_amt, "已回款");
    let app_paid_cnt = get_cnt(&app_cnt, "已回款");
    let dig_paid_amt = get_val(&dig_amt, "已回款");
    let dig_paid_cnt = get_cnt(&dig_cnt, "已回款");

    let app_unpaid_amt = app_gen_amt - app_paid_amt;
    let app_unpaid_cnt = app_gen_cnt.to_i64().unwrap_or(0) - app_paid_cnt;
    let dig_unpaid_amt = dig_gen_amt - dig_paid_amt;
    let dig_unpaid_cnt = dig_gen_cnt.to_i64().unwrap_or(0) - dig_paid_cnt;

    let app_pass_amt = get_val(&app_amt, "审核通过未回款");
    let app_pass_cnt = get_cnt(&app_cnt, "审核通过未回款");
    let dig_pass_amt = get_val(&dig_amt, "审核通过未回款");
    let dig_pass_cnt = get_cnt(&dig_cnt, "审核通过未回款");

    let app_wait_amt = get_val(&app_amt, "待审核");
    let app_wait_cnt = get_cnt(&app_cnt, "待审核");
    let dig_wait_amt = get_val(&dig_amt, "待审核");
    let dig_wait_cnt = get_cnt(&dig_cnt, "待审核");

    let app_fail_amt = get_val(&app_amt, "审核失败");
    let app_fail_cnt = get_cnt(&app_cnt, "审核失败");
    let dig_fail_amt = get_val(&dig_amt, "审核失败");
    let dig_fail_cnt = get_cnt(&dig_cnt, "审核失败");

    let app_unup_amt = app_unpaid_amt - app_pass_amt - app_wait_amt - app_fail_amt;
    let app_unup_cnt = app_unpaid_cnt - app_pass_cnt - app_wait_cnt - app_fail_cnt;
    let dig_unup_amt = dig_unpaid_amt - dig_pass_amt - dig_wait_amt - dig_fail_amt;
    let dig_unup_cnt = dig_unpaid_cnt - dig_pass_cnt - dig_wait_cnt - dig_fail_cnt;

    // Convert Yuan to Wan Yuan
    let to_wan = |d: Decimal| (d / Decimal::from(10000)).to_f64().unwrap_or(0.0);

    // Row 1: Title
    ws.set_row_height(0, 30.0).map_err(|e| e.to_string())?;
    ws.merge_range(0, 0, 0, 6, "国补上传情况汇总", &s.title).map_err(|e| e.to_string())?;

    // Row 2: Subtitle
    ws.set_row_height(1, 22.0).map_err(|e| e.to_string())?;
    ws.merge_range(1, 0, 1, 6, "数据截止口径：按所提供源文件全量数据统计；金额单位：万元", &s.subtitle).map_err(|e| e.to_string())?;

    // Row 3: Blank
    ws.set_row_height(2, 8.0).map_err(|e| e.to_string())?;

    // Row 4: Section 1 Header
    ws.set_row_height(3, 24.0).map_err(|e| e.to_string())?;
    ws.merge_range(3, 0, 3, 6, "一、金额与数量汇总", &s.section_header).map_err(|e| e.to_string())?;

    // Row 5 & 6: Col Headers
    ws.set_row_height(4, 30.0).map_err(|e| e.to_string())?;
    ws.set_row_height(5, 30.0).map_err(|e| e.to_string())?;
    ws.merge_range(4, 0, 5, 0, "项目", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(4, 1, 4, 2, "家电电脑", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(4, 3, 4, 4, "数码", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(4, 5, 4, 6, "合计", &s.col_header_center).map_err(|e| e.to_string())?;

    ws.write_string_with_format(5, 1, "金额（万元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(5, 2, "笔数", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(5, 3, "金额（万元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(5, 4, "笔数", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(5, 5, "金额（万元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(5, 6, "笔数", &s.col_header_center).map_err(|e| e.to_string())?;

    // Table 1 Rows (Rows 7-14, 0-based 6-13)
    let t1_data: [(&str, Decimal, i64, Decimal, i64); 7] = [
        ("国补发生额", app_gen_amt, app_gen_cnt.to_i64().unwrap_or(0), dig_gen_amt, dig_gen_cnt.to_i64().unwrap_or(0)),
        ("国补已回款额", app_paid_amt, app_paid_cnt, dig_paid_amt, dig_paid_cnt),
        ("未回款", app_unpaid_amt, app_unpaid_cnt, dig_unpaid_amt, dig_unpaid_cnt),
        ("审核通过未回款", app_pass_amt, app_pass_cnt, dig_pass_amt, dig_pass_cnt),
        ("待审核", app_wait_amt, app_wait_cnt, dig_wait_amt, dig_wait_cnt),
        ("审核失败", app_fail_amt, app_fail_cnt, dig_fail_amt, dig_fail_cnt),
        ("未上传", app_unup_amt, app_unup_cnt, dig_unup_amt, dig_unup_cnt),
    ];

    let mut cur_row = 6;
    for (i, (name, a_amt, a_cnt, d_amt, d_cnt)) in t1_data.iter().enumerate() {
        if i == 3 {
            // Row 10 (0-based 9): Tip row
            ws.set_row_height(cur_row, 22.0).map_err(|e| e.to_string())?;
            ws.merge_range(cur_row, 0, cur_row, 6, "未回款具体情况", &s.tip_row).map_err(|e| e.to_string())?;
            cur_row += 1;
        }

        ws.set_row_height(cur_row, 22.0).map_err(|e| e.to_string())?;
        ws.write_string_with_format(cur_row, 0, *name, &s.text_left).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 1, to_wan(*a_amt), &s.money_wan).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 2, *a_cnt as f64, &s.int_count).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 3, to_wan(*d_amt), &s.money_wan).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 4, *d_cnt as f64, &s.int_count).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 5, to_wan(*a_amt + *d_amt), &s.money_wan).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 6, (*a_cnt + *d_cnt) as f64, &s.int_count).map_err(|e| e.to_string())?;
        cur_row += 1;
    }

    // Row 15: Blank
    ws.set_row_height(cur_row, 8.0).map_err(|e| e.to_string())?;
    cur_row += 1;

    // Row 16: Section 2 Header
    ws.set_row_height(cur_row, 24.0).map_err(|e| e.to_string())?;
    ws.merge_range(cur_row, 0, cur_row, 6, "二、结构与比率分析", &s.section_header).map_err(|e| e.to_string())?;
    cur_row += 1;

    // Row 17 & 18: Table 2 Headers
    ws.set_row_height(cur_row, 30.0).map_err(|e| e.to_string())?;
    ws.set_row_height(cur_row + 1, 30.0).map_err(|e| e.to_string())?;
    ws.merge_range(cur_row, 0, cur_row + 1, 0, "指标", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(cur_row, 1, cur_row, 2, "家电电脑", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(cur_row, 3, cur_row, 4, "数码", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(cur_row, 5, cur_row, 6, "合计", &s.col_header_center).map_err(|e| e.to_string())?;

    ws.write_string_with_format(cur_row + 1, 1, "金额（万元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(cur_row + 1, 2, "占比/比率", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(cur_row + 1, 3, "金额（万元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(cur_row + 1, 4, "占比/比率", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(cur_row + 1, 5, "金额（万元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(cur_row + 1, 6, "占比/比率", &s.col_header_center).map_err(|e| e.to_string())?;
    cur_row += 2;

    // Table 2 Ratios
    let tot_gen_amt = app_gen_amt + dig_gen_amt;
    let tot_paid_amt = app_paid_amt + dig_paid_amt;
    let tot_unpaid_amt = app_unpaid_amt + dig_unpaid_amt;
    let tot_pass_amt = app_pass_amt + dig_pass_amt;
    let tot_wait_amt = app_wait_amt + dig_wait_amt;
    let tot_fail_amt = app_fail_amt + dig_fail_amt;
    let tot_unup_amt = app_unup_amt + dig_unup_amt;

    let div_pct = |num: Decimal, den: Decimal| -> f64 {
        if den.is_zero() {
            0.0
        } else {
            (num / den).to_f64().unwrap_or(0.0)
        }
    };

    let t2_data: [(&str, Decimal, f64, Decimal, f64, Decimal, f64); 8] = [
        ("国补发生额及占比", app_gen_amt, div_pct(app_gen_amt, tot_gen_amt), dig_gen_amt, div_pct(dig_gen_amt, tot_gen_amt), tot_gen_amt, 1.0),
        ("国补回款额及回款率", app_paid_amt, div_pct(app_paid_amt, app_gen_amt), dig_paid_amt, div_pct(dig_paid_amt, dig_gen_amt), tot_paid_amt, div_pct(tot_paid_amt, tot_gen_amt)),
        ("回款+审核通过未回款及综合回款率", app_paid_amt + app_pass_amt, div_pct(app_paid_amt + app_pass_amt, app_gen_amt), dig_paid_amt + dig_pass_amt, div_pct(dig_paid_amt + dig_pass_amt, dig_gen_amt), tot_paid_amt + tot_pass_amt, div_pct(tot_paid_amt + tot_pass_amt, tot_gen_amt)),
        ("未回款额", app_unpaid_amt, div_pct(app_unpaid_amt, app_gen_amt), dig_unpaid_amt, div_pct(dig_unpaid_amt, dig_gen_amt), tot_unpaid_amt, div_pct(tot_unpaid_amt, tot_gen_amt)),
        ("审核通过未回款", app_pass_amt, div_pct(app_pass_amt, app_unpaid_amt), dig_pass_amt, div_pct(dig_pass_amt, dig_unpaid_amt), tot_pass_amt, div_pct(tot_pass_amt, tot_unpaid_amt)),
        ("待审核", app_wait_amt, div_pct(app_wait_amt, app_unpaid_amt), dig_wait_amt, div_pct(dig_wait_amt, dig_unpaid_amt), tot_wait_amt, div_pct(tot_wait_amt, tot_unpaid_amt)),
        ("审核失败", app_fail_amt, div_pct(app_fail_amt, app_unpaid_amt), dig_fail_amt, div_pct(dig_fail_amt, dig_unpaid_amt), tot_fail_amt, div_pct(tot_fail_amt, tot_unpaid_amt)),
        ("未上传", app_unup_amt, div_pct(app_unup_amt, app_unpaid_amt), dig_unup_amt, div_pct(dig_unup_amt, dig_unpaid_amt), tot_unup_amt, div_pct(tot_unup_amt, tot_unpaid_amt)),
    ];

    for (i, (name, a_amt, a_pct, d_amt, d_pct, t_amt, t_pct)) in t2_data.iter().enumerate() {
        if i == 4 {
            // Row 23 (0-based 22): Tip row
            ws.set_row_height(cur_row, 22.0).map_err(|e| e.to_string())?;
            ws.merge_range(cur_row, 0, cur_row, 6, "未回款具体情况", &s.tip_row).map_err(|e| e.to_string())?;
            cur_row += 1;
        }

        ws.set_row_height(cur_row, 24.0).map_err(|e| e.to_string())?;
        ws.write_string_with_format(cur_row, 0, *name, &s.text_left).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 1, to_wan(*a_amt), &s.money_wan).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 2, *a_pct, &s.percent).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 3, to_wan(*d_amt), &s.money_wan).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 4, *d_pct, &s.percent).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 5, to_wan(*t_amt), &s.money_wan).map_err(|e| e.to_string())?;
        ws.write_number_with_format(cur_row, 6, *t_pct, &s.percent).map_err(|e| e.to_string())?;
        cur_row += 1;
    }

    // Row 28: Blank
    ws.set_row_height(cur_row, 8.0).map_err(|e| e.to_string())?;
    cur_row += 1;

    // Row 29: Notes Header
    ws.set_row_height(cur_row, 24.0).map_err(|e| e.to_string())?;
    ws.merge_range(cur_row, 0, cur_row, 6, "口径说明", &s.section_header).map_err(|e| e.to_string())?;
    cur_row += 1;

    let notes = [
        "1. 国补发生额及发生数量取自《销售用券情况统计.xlsx》；数量按“数量”字段净额统计，包含退货负数冲减。",
        "2. “财务大类=数码”计入数码，其余财务大类计入家电电脑。",
        "3. 已回款、审核通过未回款、待审核、审核失败取自两份“已上传”明细；金额汇总“补贴金额”，数量按明细记录数。",
        "4. 未回款=国补发生额-国补回款额；未上传=未回款-审核通过未回款-待审核-审核失败。",
        "5. 回款率=国补回款额/国补发生额；综合回款率=(国补回款额+审核通过未回款额)/国补发生额。",
    ];

    for note in notes {
        ws.set_row_height(cur_row, 28.0).map_err(|e| e.to_string())?;
        ws.merge_range(cur_row, 0, cur_row, 6, note, &s.text_left).map_err(|e| e.to_string())?;
        cur_row += 1;
    }

    Ok(())
}

fn build_category_brand_sheet(
    wb: &mut Workbook,
    s: &StylePool,
    sales: &[Vec<Data>],
) -> Result<(), String> {
    let ws = wb.add_worksheet();
    ws.set_name("品类品牌汇总").map_err(|e| e.to_string())?;

    ws.set_column_width(0, 16.0).map_err(|e| e.to_string())?;
    ws.set_column_width(1, 20.0).map_err(|e| e.to_string())?;
    for col in 2..=11 {
        ws.set_column_width(col, 16.0).map_err(|e| e.to_string())?;
    }

    // Title & Subtitle
    ws.set_row_height(0, 30.0).map_err(|e| e.to_string())?;
    ws.merge_range(0, 0, 0, 11, "品类品牌汇总", &s.title).map_err(|e| e.to_string())?;

    ws.set_row_height(1, 22.0).map_err(|e| e.to_string())?;
    ws.merge_range(1, 0, 1, 11, "数据源：销售用券情况统计.xlsx；已排除退货-原单、退货-退单；金额单位：元", &s.subtitle).map_err(|e| e.to_string())?;

    ws.set_row_height(2, 8.0).map_err(|e| e.to_string())?;

    // Header 2 layers
    ws.set_row_height(3, 30.0).map_err(|e| e.to_string())?;
    ws.set_row_height(4, 30.0).map_err(|e| e.to_string())?;
    ws.merge_range(3, 0, 4, 0, "财务大类", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(3, 1, 4, 1, "品牌", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(3, 2, 3, 3, "已回款", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(3, 4, 3, 5, "审核通过未回款", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(3, 6, 3, 7, "待审核", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(3, 8, 3, 9, "审核失败", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.merge_range(3, 10, 3, 11, "未上传", &s.col_header_center).map_err(|e| e.to_string())?;

    ws.write_string_with_format(4, 2, "补贴金额（元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(4, 3, "数量", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(4, 4, "补贴金额（元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(4, 5, "数量", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(4, 6, "补贴金额（元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(4, 7, "数量", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(4, 8, "补贴金额（元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(4, 9, "数量", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(4, 10, "补贴金额（元）", &s.col_header_center).map_err(|e| e.to_string())?;
    ws.write_string_with_format(4, 11, "数量", &s.col_header_center).map_err(|e| e.to_string())?;

    // Aggregation
    let h = HeaderMap::from_header_row(&sales[0]);
    let cat_idx = h.find(&["财务大类"]).unwrap();
    let brand_idx = h.find(&["品牌"]).unwrap();
    let sub_idx = h.find(&["补贴额"]).unwrap();
    let qty_idx = h.find(&["数量"]).unwrap();
    let remark_idx = h.find(&["备注"]).unwrap();

    // (Category, Brand) -> Status -> (Amount, Count)
    let mut matrix: HashMap<(String, String), HashMap<String, (Decimal, i64)>> = HashMap::new();

    for row in &sales[1..] {
        let remark = cell_to_string(&row[remark_idx]);
        if remark == "退货-原单" || remark == "退货-退单" {
            continue;
        }
        let cat = cell_to_string(&row[cat_idx]);
        let brand = cell_to_string(&row[brand_idx]);
        let amt = cell_to_decimal(&row[sub_idx]).unwrap_or(Decimal::ZERO);
        let qty = cell_to_decimal(&row[qty_idx]).and_then(|d| d.to_i64()).unwrap_or(0);

        let entry = matrix.entry((cat, brand)).or_default();
        let status_entry = entry.entry(remark).or_insert((Decimal::ZERO, 0));
        status_entry.0 += amt;
        status_entry.1 += qty;
    }

    // Standard base category and brand ordering; newly appeared categories/brands are appended dynamically
    let std_categories: [(&str, &[&str]); 7] = [
        ("厨卫", &["AO史密斯", "万家乐", "方太", "欧意", "海尔", "美的", "老板"]),
        ("洗衣机", &["博世", "小鸭", "海尔", "美的", "美菱", "西门子"]),
        ("冰箱", &["博世", "海尔", "美的", "美菱", "西门子"]),
        ("彩电", &["TCL", "创维", "海信", "海尔"]),
        ("空调", &["TCL", "奥克斯", "格力", "海信", "海尔", "科龙", "美的"]),
        ("小电", &["沁园"]),
        ("数码", &[
            "IQOO", "OPPO", "VIVO", "一加", "作业帮", "华为", "学而思", "小天才", "小米", "步步高",
            "真我（REALME）", "苹果", "荣耀",
        ]),
    ];

    // Collect all categories actually in matrix
    let mut all_cats: Vec<String> = Vec::new();
    for (cat, _) in std_categories {
        if matrix.keys().any(|(c, _)| c == cat) {
            all_cats.push(cat.to_string());
        }
    }
    for (c, _) in matrix.keys() {
        if !all_cats.contains(c) {
            all_cats.push(c.clone());
        }
    }

    let mut cat_brands: Vec<(String, Vec<String>)> = Vec::new();
    for cat in all_cats {
        let std_brands: Vec<String> = std_categories
            .iter()
            .find(|(c, _)| *c == cat)
            .map(|(_, bs)| bs.iter().map(|s| s.to_string()).collect())
            .unwrap_or_default();

        let mut brands: Vec<String> = Vec::new();
        // Add standard brands if present in matrix
        for b in &std_brands {
            if matrix.contains_key(&(cat.clone(), b.clone())) {
                brands.push(b.clone());
            }
        }
        // Add any newly encountered brands for this category
        let mut new_brands: Vec<String> = matrix
            .keys()
            .filter(|(c, b)| c == &cat && !std_brands.contains(b))
            .map(|(_, b)| b.clone())
            .collect();
        new_brands.sort();
        brands.extend(new_brands);

        if !brands.is_empty() {
            cat_brands.push((cat, brands));
        }
    }

    let statuses = ["已回款", "审核通过未回款", "待审核", "审核失败", "未上传"];

    let mut row_idx: u32 = 5; // 0-indexed row 5 = Excel Row 6
    for (cat, brands) in cat_brands {
        let start_row = row_idx;
        for brand in &brands {
            ws.set_row_height(row_idx, 22.0).map_err(|e| e.to_string())?;
            ws.write_string_with_format(row_idx, 0, &cat, &s.text_center).map_err(|e| e.to_string())?;
            ws.write_string_with_format(row_idx, 1, brand.as_str(), &s.text_left).map_err(|e| e.to_string())?;

            let empty_map = HashMap::new();
            let brand_map = matrix.get(&(cat.to_string(), brand.to_string())).unwrap_or(&empty_map);

            for (s_idx, st) in statuses.iter().enumerate() {
                let (amt, cnt) = brand_map.get(*st).copied().unwrap_or((Decimal::ZERO, 0));
                let col_amt = 2 + (s_idx as u16) * 2;
                let col_cnt = col_amt + 1;

                if amt.is_zero() && cnt == 0 {
                    ws.write_string_with_format(row_idx, col_amt, "-", &s.text_right).map_err(|e| e.to_string())?;
                    ws.write_string_with_format(row_idx, col_cnt, "-", &s.text_right).map_err(|e| e.to_string())?;
                } else {
                    ws.write_number_with_format(row_idx, col_amt, amt.to_f64().unwrap_or(0.0), &s.money_yuan).map_err(|e| e.to_string())?;
                    ws.write_number_with_format(row_idx, col_cnt, cnt as f64, &s.int_count).map_err(|e| e.to_string())?;
                }
            }
            row_idx += 1;
        }
        let end_row = row_idx - 1;
        if end_row > start_row {
            ws.merge_range(start_row, 0, end_row, 0, &cat, &s.text_center).map_err(|e| e.to_string())?;
        }
    }

    Ok(())
}

fn build_failed_records_sheet(
    wb: &mut Workbook,
    s: &StylePool,
    app_up: &[Vec<Data>],
    dig_up: &[Vec<Data>],
    invoices: &[Vec<Data>],
) -> Result<(), String> {
    let ws = wb.add_worksheet();
    ws.set_name("审核失败明细").map_err(|e| e.to_string())?;

    // Column widths: A=14, B=20, C=16, D=48, E=24, F=16, G=16, H=24, I=38, J=20
    let widths = [14.0, 20.0, 16.0, 48.0, 24.0, 16.0, 16.0, 24.0, 38.0, 20.0];
    for (col, w) in widths.iter().enumerate() {
        ws.set_column_width(col as u16, *w).map_err(|e| e.to_string())?;
    }

    // Invoice lookup: 发票号码 -> 主要商品名称
    let inv_h = HeaderMap::from_header_row(&invoices[0]);
    let inv_no_idx = inv_h.find(&["数电发票号码"]).ok_or("发票明细缺少数电发票号码")?;
    let inv_name_idx = inv_h.find(&["主要商品名称"]).ok_or("发票明细缺少主要商品名称")?;
    let mut invoice_name_map = HashMap::new();
    for row in &invoices[1..] {
        let no = cell_to_string(&row[inv_no_idx]);
        let name = cell_to_string(&row[inv_name_idx]);
        if !no.is_empty() {
            invoice_name_map.insert(no, name);
        }
    }

    // Title & Subtitle
    ws.set_row_height(0, 30.0).map_err(|e| e.to_string())?;
    ws.merge_range(0, 0, 0, 9, "审核失败明细", &s.title).map_err(|e| e.to_string())?;

    ws.set_row_height(1, 22.0).map_err(|e| e.to_string())?;
    ws.merge_range(1, 0, 1, 9, "数据源：已上传家电电脑.xlsx、已上传数码.xlsx、发票明细.xlsx；发票金额单位：元", &s.subtitle).map_err(|e| e.to_string())?;

    ws.set_row_height(2, 8.0).map_err(|e| e.to_string())?;

    // 1. Appliance Section
    let mut cur_row: u32 = 3;
    ws.set_row_height(cur_row, 24.0).map_err(|e| e.to_string())?;
    ws.merge_range(cur_row, 0, cur_row, 8, "家电电脑审核失败明细", &s.section_header).map_err(|e| e.to_string())?;
    cur_row += 1;

    let app_headers = ["交易日期", "检索参考号", "状态", "描述", "发票号码", "发票金额", "购买方名称", "S/N码", "商品名称"];
    ws.set_row_height(cur_row, 30.0).map_err(|e| e.to_string())?;
    for (col, h) in app_headers.iter().enumerate() {
        ws.write_string_with_format(cur_row, col as u16, *h, &s.col_header).map_err(|e| e.to_string())?;
    }
    cur_row += 1;

    let app_h = HeaderMap::from_header_row(&app_up[0]);
    let date_idx = app_h.find(&["交易日期"]).unwrap();
    let ref_idx = app_h.find(&["检索参考号"]).unwrap();
    let st_idx = app_h.find(&["状态"]).unwrap();
    let desc_idx = app_h.find(&["描述"]).unwrap();
    let inv_idx = app_h.find(&["发票号码"]).unwrap();
    let inv_amt_idx = app_h.find(&["发票金额"]).unwrap();
    let buyer_idx = app_h.find(&["购买方名称"]).unwrap();
    let sn_idx = app_h.find(&["S/N码", "sn码"]).unwrap();

    for row in &app_up[1..] {
        if cell_to_string(&row[st_idx]) == "审核失败" {
            ws.set_row_height(cur_row, 22.0).map_err(|e| e.to_string())?;
            let inv_no = cell_to_string(&row[inv_idx]);
            let prod_name = invoice_name_map.get(&inv_no).cloned().unwrap_or_default();

            ws.write_string_with_format(cur_row, 0, cell_to_string(&row[date_idx]), &s.date).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 1, cell_to_string(&row[ref_idx]), &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 2, "审核失败", &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 3, cell_to_string(&row[desc_idx]), &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 4, inv_no, &s.text_left).map_err(|e| e.to_string())?;
            let amt = cell_to_decimal(&row[inv_amt_idx]).unwrap_or(Decimal::ZERO).to_f64().unwrap_or(0.0);
            ws.write_number_with_format(cur_row, 5, amt, &s.money_yuan).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 6, cell_to_string(&row[buyer_idx]), &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 7, cell_to_string(&row[sn_idx]), &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 8, prod_name, &s.text_left).map_err(|e| e.to_string())?;
            cur_row += 1;
        }
    }

    // Blank row
    ws.set_row_height(cur_row, 8.0).map_err(|e| e.to_string())?;
    cur_row += 1;

    // 2. Digital Section
    ws.set_row_height(cur_row, 24.0).map_err(|e| e.to_string())?;
    ws.merge_range(cur_row, 0, cur_row, 9, "数码审核失败明细", &s.section_header).map_err(|e| e.to_string())?;
    cur_row += 1;

    let dig_headers = ["交易日期", "检索参考号", "状态", "描述", "发票号码", "发票金额", "购买方名称", "S/N码", "IMEI1", "IMEI2"];
    ws.set_row_height(cur_row, 30.0).map_err(|e| e.to_string())?;
    for (col, h) in dig_headers.iter().enumerate() {
        ws.write_string_with_format(cur_row, col as u16, *h, &s.col_header).map_err(|e| e.to_string())?;
    }
    cur_row += 1;

    let dig_h = HeaderMap::from_header_row(&dig_up[0]);
    let d_date_idx = dig_h.find(&["交易日期"]).unwrap();
    let d_ref_idx = dig_h.find(&["检索参考号"]).unwrap();
    let d_st_idx = dig_h.find(&["状态"]).unwrap();
    let d_desc_idx = dig_h.find(&["描述"]).unwrap();
    let d_inv_idx = dig_h.find(&["发票号码"]).unwrap();
    let d_inv_amt_idx = dig_h.find(&["发票金额"]).unwrap();
    let d_buyer_idx = dig_h.find(&["购买方名称"]).unwrap();
    let d_sn_idx = dig_h.find(&["S/N码", "sn码"]).unwrap();
    let d_imei1_idx = dig_h.find(&["IMEI1", "imei1"]).unwrap();
    let d_imei2_idx = dig_h.find(&["IMEI2", "imei2"]).unwrap();

    for row in &dig_up[1..] {
        if cell_to_string(&row[d_st_idx]) == "审核失败" {
            ws.set_row_height(cur_row, 22.0).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 0, cell_to_string(&row[d_date_idx]), &s.date).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 1, cell_to_string(&row[d_ref_idx]), &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 2, "审核失败", &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 3, cell_to_string(&row[d_desc_idx]), &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 4, cell_to_string(&row[d_inv_idx]), &s.text_left).map_err(|e| e.to_string())?;
            let amt = cell_to_decimal(&row[d_inv_amt_idx]).unwrap_or(Decimal::ZERO).to_f64().unwrap_or(0.0);
            ws.write_number_with_format(cur_row, 5, amt, &s.money_yuan).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 6, cell_to_string(&row[d_buyer_idx]), &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 7, cell_to_string(&row[d_sn_idx]), &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 8, cell_to_string(&row[d_imei1_idx]), &s.text_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(cur_row, 9, cell_to_string(&row[d_imei2_idx]), &s.text_left).map_err(|e| e.to_string())?;
            cur_row += 1;
        }
    }

    Ok(())
}

fn build_refund_anomaly_sheet(
    wb: &mut Workbook,
    s: &StylePool,
    app_refund_path: &Path,
    dig_refund_path: &Path,
    app_store_code: &str,
    dig_store_code: &str,
) -> Result<(), String> {
    let ws = wb.add_worksheet();
    ws.set_name("异常回款明细").map_err(|e| e.to_string())?;

    // 24 column widths
    let widths = [
        32.0, 20.0, 16.0, 26.0, 26.0, 22.0, 18.0, 14.0, 14.0, 14.0, 14.0, 12.0, 24.0, 14.0, 16.0,
        12.0, 16.0, 36.0, 14.0, 24.0, 18.0, 24.0, 48.0, 18.0,
    ];
    for (col, w) in widths.iter().enumerate() {
        ws.set_column_width(col as u16, *w).map_err(|e| e.to_string())?;
    }

    // Title & Subtitle
    ws.set_row_height(0, 30.0).map_err(|e| e.to_string())?;
    ws.merge_range(0, 0, 0, 23, "异常回款明细", &s.title).map_err(|e| e.to_string())?;

    ws.set_row_height(1, 22.0).map_err(|e| e.to_string())?;
    ws.merge_range(1, 0, 1, 23, "提取规则：筛选本店核销商编且单元格为粉色填充（颜色 #FFC7CE）；金额单位：元", &s.subtitle).map_err(|e| e.to_string())?;

    ws.set_row_height(2, 8.0).map_err(|e| e.to_string())?;

    let write_section = |ws: &mut Worksheet,
                         title: &str,
                         path: &Path,
                         store_code: &str,
                         start_row: &mut u32|
     -> Result<(), String> {
        let pink_rows = get_colored_row_indices(path, "FFFFC7CE")?;
        let rows = read_sheet_rows(path)?;
        let h = HeaderMap::from_header_row(&rows[0]);
        let mch_idx = h.find(&["核销商编"]).ok_or(format!("{:?} 缺少核销商编", path))?;

        ws.set_row_height(*start_row, 24.0).map_err(|e| e.to_string())?;
        ws.merge_range(*start_row, 0, *start_row, 23, title, &s.section_header).map_err(|e| e.to_string())?;
        *start_row += 1;

        // Write header row (24 columns from source)
        ws.set_row_height(*start_row, 34.0).map_err(|e| e.to_string())?;
        for col in 0..24 {
            let col_name = if col < rows[0].len() {
                cell_to_string(&rows[0][col])
            } else {
                String::new()
            };
            ws.write_string_with_format(*start_row, col as u16, col_name, &s.col_header).map_err(|e| e.to_string())?;
        }
        *start_row += 1;

        // Write data rows
        for (r_idx, row) in rows.iter().enumerate().skip(1) {
            let excel_row_num = (r_idx + 1) as u32;
            let mch = cell_to_string(&row[mch_idx]);

            // Dual filter: store_code AND pink fill
            if mch == store_code && pink_rows.contains(&excel_row_num) {
                ws.set_row_height(*start_row, 22.0).map_err(|e| e.to_string())?;
                for col in 0..24 {
                    let val = if col < row.len() { &row[col] } else { &Data::Empty };
                    write_refund_cell(ws, s, *start_row, col as u16, val)?;
                }
                *start_row += 1;
            }
        }
        Ok(())
    };

    let mut cur_row = 3;
    write_section(ws, "家电电脑异常回款明细", app_refund_path, app_store_code, &mut cur_row)?;

    // Blank row
    ws.set_row_height(cur_row, 8.0).map_err(|e| e.to_string())?;
    cur_row += 1;

    write_section(ws, "数码异常回款明细", dig_refund_path, dig_store_code, &mut cur_row)?;

    Ok(())
}

pub fn write_refund_cell(
    ws: &mut Worksheet,
    s: &StylePool,
    row: u32,
    col: u16,
    cell: &Data,
) -> Result<(), String> {
    match col {
        // Col 1: DateTime (交易完成时间)
        1 => {
            let val = cell_to_string(cell);
            if val.is_empty() {
                ws.write_string_with_format(row, col, "", &s.text_center).map_err(|e| e.to_string())?;
            } else {
                ws.write_string_with_format(row, col, val, &s.datetime).map_err(|e| e.to_string())?;
            }
        }
        // Money columns: 8 (销售金额), 9 (实收销售金额), 10 (补贴金额), 18 (发票金额)
        8 | 9 | 10 | 18 => {
            if let Some(dec) = cell_to_decimal(cell) {
                ws.write_number_with_format(row, col, dec.to_f64().unwrap_or(0.0), &s.money_yuan).map_err(|e| e.to_string())?;
            } else {
                ws.write_string_with_format(row, col, "", &s.text_right).map_err(|e| e.to_string())?;
            }
        }
        // Col 11: Percent (补贴比例)
        11 => {
            if let Some(dec) = cell_to_decimal(cell) {
                ws.write_number_with_format(row, col, dec.to_f64().unwrap_or(0.0), &s.percent).map_err(|e| e.to_string())?;
            } else {
                ws.write_string_with_format(row, col, "", &s.text_right).map_err(|e| e.to_string())?;
            }
        }
        // Plain text columns
        _ => {
            let val = cell_to_string(cell);
            ws.write_string_with_format(row, col, val, &s.text_left).map_err(|e| e.to_string())?;
        }
    }
    Ok(())
}

fn build_invoice_anomaly_sheet(
    wb: &mut Workbook,
    s: &StylePool,
    invoice_path: &Path,
    invoices: &[Vec<Data>],
) -> Result<(), String> {
    let ws = wb.add_worksheet();
    ws.set_name("异常发票明细").map_err(|e| e.to_string())?;

    // 8 columns: widths [20.0, 14.0, 24.0, 18.0, 40.0, 58.0, 16.0, 22.0]
    let widths = [20.0, 14.0, 24.0, 18.0, 40.0, 58.0, 16.0, 22.0];
    for (col, w) in widths.iter().enumerate() {
        ws.set_column_width(col as u16, *w).map_err(|e| e.to_string())?;
    }

    // Title & Subtitle
    ws.set_row_height(0, 30.0).map_err(|e| e.to_string())?;
    ws.merge_range(0, 0, 0, 7, "异常发票明细", &s.title).map_err(|e| e.to_string())?;

    ws.set_row_height(1, 22.0).map_err(|e| e.to_string())?;
    ws.merge_range(1, 0, 1, 7, "提取规则：发票明细黄色填充区域（颜色 #FFEB9C）；按匹配单据号升序排列", &s.subtitle).map_err(|e| e.to_string())?;

    ws.set_row_height(2, 8.0).map_err(|e| e.to_string())?;

    // Section Header
    ws.set_row_height(3, 24.0).map_err(|e| e.to_string())?;
    ws.merge_range(3, 0, 3, 7, "黄色异常发票明细", &s.section_header).map_err(|e| e.to_string())?;

    // Column Headers
    ws.set_row_height(4, 30.0).map_err(|e| e.to_string())?;
    for col in 0..8 {
        let h_name = cell_to_string(&invoices[0][col]);
        ws.write_string_with_format(4, col as u16, h_name, &s.col_header).map_err(|e| e.to_string())?;
    }

    let yellow_rows = get_colored_row_indices(invoice_path, "FFFFEB9C")?;
    let h = HeaderMap::from_header_row(&invoices[0]);
    let doc_no_idx = h.find(&["匹配单据号"]).ok_or("发票明细缺少匹配单据号")?;

    // Collect and sort yellow rows
    let mut collected: Vec<&Vec<Data>> = Vec::new();
    for (r_idx, row) in invoices.iter().enumerate().skip(1) {
        let excel_row_num = (r_idx + 1) as u32;
        if yellow_rows.contains(&excel_row_num) {
            collected.push(row);
        }
    }

    // Sort ascending by 匹配单据号 (empty at the end)
    collected.sort_by_key(|r| {
        let s = cell_to_string(&r[doc_no_idx]);
        (s.is_empty(), s)
    });

    let mut cur_row: u32 = 5;
    for row in collected {
        ws.set_row_height(cur_row, 22.0).map_err(|e| e.to_string())?;
        for col in 0..8 {
            let val = cell_to_string(&row[col]);
            if col == 0 {
                // 开票时间
                ws.write_string_with_format(cur_row, col as u16, val, &s.datetime).map_err(|e| e.to_string())?;
            } else {
                ws.write_string_with_format(cur_row, col as u16, val, &s.text_left).map_err(|e| e.to_string())?;
            }
        }
        cur_row += 1;
    }

    Ok(())
}
