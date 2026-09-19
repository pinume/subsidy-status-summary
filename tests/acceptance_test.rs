use std::path::Path;
use calamine::{Reader, open_workbook_auto};
use data_status_summary::{generate_store_finance_workbook, generate_summary_workbook};

#[test]
fn test_all_acceptance_benchmarks() {
    let source_dir = Path::new("/home/ubuntu/github/source_data");
    if !source_dir.exists() {
        eprintln!("Source directory not found, skipping test");
        return;
    }

    let out_dir = Path::new("/tmp/test_subsidy_summary");
    let _ = std::fs::create_dir_all(out_dir);

    let wb1_path = out_dir.join("国补上传情况汇总.xlsx");
    let wb2_path = out_dir.join("26年国补门店财务统筹表.xlsx");

    // 1. Run Workbook 1
    generate_summary_workbook(source_dir, &wb1_path).expect("Generate WB1 failed");
    assert!(wb1_path.exists(), "WB1 file should exist");

    let mut wb1 = open_workbook_auto(&wb1_path).expect("Open WB1 failed");
    let sheets1 = wb1.sheet_names();
    assert_eq!(
        sheets1,
        vec!["汇总", "品类品牌汇总", "审核失败明细", "异常回款明细", "异常发票明细"]
    );

    // Verify 品类品牌汇总 has 43 brand rows (Rows 6-48)
    let r_cat = wb1.worksheet_range("品类品牌汇总").unwrap();
    assert_eq!(r_cat.rows().count(), 48); // Row 1 to 48

    // Verify 审核失败明细 has 132 appliance + 74 digital = 206
    let r_fail = wb1.worksheet_range("审核失败明细").unwrap();
    // Count rows with status == "审核失败"
    let fail_cnt = r_fail
        .rows()
        .filter(|r| r.get(2).map(|c| c.to_string() == "审核失败").unwrap_or(false))
        .count();
    assert_eq!(fail_cnt, 206);

    // Verify 异常回款明细 has 44 + 2 = 46 rows (Row 6-49 and Row 53-54)
    let r_ref_anom = wb1.worksheet_range("异常回款明细").unwrap();
    let ref_anom_cnt = r_ref_anom
        .rows()
        .enumerate()
        .filter(|(idx, _)| {
            let row_num = idx + 1;
            (row_num >= 6 && row_num <= 49) || (row_num >= 53 && row_num <= 54)
        })
        .count();
    assert_eq!(ref_anom_cnt, 46);

    // Verify 异常发票明细 has 38 rows (Row 6-43)
    let r_inv_anom = wb1.worksheet_range("异常发票明细").unwrap();
    let inv_anom_cnt = r_inv_anom
        .rows()
        .enumerate()
        .filter(|(idx, _)| {
            let row_num = idx + 1;
            row_num >= 6 && row_num <= 43
        })
        .count();
    assert_eq!(inv_anom_cnt, 38);

    // 2. Run Workbook 2
    generate_store_finance_workbook(source_dir, &wb2_path).expect("Generate WB2 failed");
    assert!(wb2_path.exists(), "WB2 file should exist");

    let mut wb2 = open_workbook_auto(&wb2_path).expect("Open WB2 failed");
    let sheets2 = wb2.sheet_names();
    assert_eq!(
        sheets2,
        vec![
            "最终匹配表（全部的国补发生数据上匹配）",
            "1.门店国补发生表（银联系统直接导出，不需要加工）",
            "2.门店累计回款表（事业部财务下发，门店筛选自己的）",
            "3.门店上传明细（从门店银联后台每月导出后汇总）"
        ]
    );

    // Verify Sheet 2: 1.门店国补发生表 (13,971 data rows + 1 header = 13,972)
    let r_occ = wb2.worksheet_range("1.门店国补发生表（银联系统直接导出，不需要加工）").unwrap();
    assert_eq!(r_occ.rows().count(), 13972);
    assert_eq!(r_occ.rows().next().unwrap().len(), 26);

    // Verify Sheet 3: 2.门店累计回款表 (8,840 data rows + 1 header = 8,841)
    let r_ref = wb2.worksheet_range("2.门店累计回款表（事业部财务下发，门店筛选自己的）").unwrap();
    assert_eq!(r_ref.rows().count(), 8841);
    assert_eq!(r_ref.rows().next().unwrap().len(), 24);

    // Verify Sheet 4: 3.门店上传明细 (12,010 data rows + 1 header = 12,011)
    let r_up = wb2.worksheet_range("3.门店上传明细（从门店银联后台每月导出后汇总）").unwrap();
    assert_eq!(r_up.rows().count(), 12011);
    assert_eq!(r_up.rows().next().unwrap().len(), 61);

    // Verify Sheet 1: 最终匹配表
    let r_final = wb2.worksheet_range("最终匹配表（全部的国补发生数据上匹配）").unwrap();
    assert_eq!(r_final.rows().count(), 13972); // 13,971 data rows + 1 header
    assert_eq!(r_final.rows().next().unwrap().len(), 27);

    let mut ret_cnt = 0;
    let mut unsubmitted_cnt = 0;
    let mut invoice_cnt = 0;
    let mut red_yes_cnt = 0;
    let mut red_no_cnt = 0;
    let mut red_blank_cnt = 0;

    for row in r_final.rows().skip(1) {
        let status = row[24].to_string();
        if status == "已退货" {
            ret_cnt += 1;
        } else if status == "未提交" {
            unsubmitted_cnt += 1;
        }

        let inv_no = row[25].to_string();
        if !inv_no.is_empty() {
            invoice_cnt += 1;
        }

        let red = row[26].to_string();
        if red == "是" {
            red_yes_cnt += 1;
        } else if red == "否" {
            red_no_cnt += 1;
        } else {
            red_blank_cnt += 1;
        }
    }

    assert_eq!(ret_cnt, 344, "已退货单据应为 344 笔");
    assert_eq!(unsubmitted_cnt, 1617, "未提交单据应为 1617 笔");
    assert_eq!(invoice_cnt, 12023, "有发票号单据应为 12023 笔");
    assert_eq!(red_yes_cnt, 65, "红冲为'是'应为 65 笔");
    assert_eq!(red_no_cnt, 11957, "红冲为'否'应为 11957 笔");
    assert_eq!(red_blank_cnt, 1949, "红冲留空应为 1949 笔");

    println!("All acceptance benchmarks passed 100%!");
}
