use calamine::{Data, DataType, Reader, open_workbook_auto};
use data_status_summary::{generate_store_finance_workbook, generate_summary_workbook};
use std::collections::HashMap;
use std::path::Path;

#[test]
fn test_all_acceptance_benchmarks() {
    let default_source = Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .join("source_data");
    let source_dir = std::env::var_os("SUBSIDY_SOURCE_DIR")
        .map(Into::into)
        .unwrap_or(default_source);
    assert!(
        source_dir.is_dir(),
        "Source directory not found: {:?}; set SUBSIDY_SOURCE_DIR to run the acceptance test",
        source_dir
    );

    let out_dir = std::env::temp_dir().join(format!("test_subsidy_summary_{}", std::process::id()));
    let _ = std::fs::create_dir_all(&out_dir);

    let wb1_path = out_dir.join("国补上传情况汇总.xlsx");
    let wb2_path = out_dir.join("26年国补门店财务统筹表.xlsx");

    // 1. Run Workbook 1
    generate_summary_workbook(&source_dir, &wb1_path).expect("Generate WB1 failed");
    assert!(wb1_path.exists(), "WB1 file should exist");

    let mut wb1 = open_workbook_auto(&wb1_path).expect("Open WB1 failed");
    let sheets1 = wb1.sheet_names();
    assert_eq!(
        sheets1,
        vec![
            "汇总",
            "品类品牌汇总",
            "审核失败明细",
            "异常回款明细",
            "异常发票明细"
        ]
    );

    let r_summary = wb1.worksheet_range("汇总").unwrap();
    let number = |row, col| {
        r_summary
            .get_value((row, col))
            .and_then(DataType::as_f64)
            .unwrap()
    };
    assert!((number(6, 1) - 339.95).abs() < 0.01);
    assert!((number(6, 3) - 313.20).abs() < 0.01);
    assert!((number(6, 5) - 653.16).abs() < 0.01);
    assert!((number(7, 5) - 428.48).abs() < 0.01);
    assert!((number(8, 5) - 224.68).abs() < 0.01);
    assert!((number(20, 6) - 0.8345).abs() < 0.0001);

    // Verify 品类品牌汇总 has 43 brand rows (Rows 6-48)
    let r_cat = wb1.worksheet_range("品类品牌汇总").unwrap();
    assert_eq!(r_cat.rows().count(), 48); // Row 1 to 48

    // Verify 审核失败明细 has 132 appliance + 74 digital = 206
    let r_fail = wb1.worksheet_range("审核失败明细").unwrap();
    // Count rows with status == "审核失败"
    let fail_cnt = r_fail
        .rows()
        .filter(|r| r.get(2).map(|c| c == "审核失败").unwrap_or(false))
        .count();
    assert_eq!(fail_cnt, 206);
    let product_names: Vec<_> = r_fail
        .rows()
        .skip(5)
        .take(132)
        .map(|row| {
            let value = row[8].to_string();
            (value.is_empty(), value)
        })
        .collect();
    assert!(product_names.windows(2).all(|pair| pair[0] <= pair[1]));

    // Verify 异常回款明细 has 44 + 2 = 46 rows (Row 6-49 and Row 53-54)
    let r_ref_anom = wb1.worksheet_range("异常回款明细").unwrap();
    let ref_anom_cnt = r_ref_anom
        .rows()
        .enumerate()
        .filter(|(idx, _)| {
            let row_num = idx + 1;
            (6..=49).contains(&row_num) || (53..=54).contains(&row_num)
        })
        .count();
    assert_eq!(ref_anom_cnt, 46);
    for references in [
        r_ref_anom
            .rows()
            .skip(5)
            .take(44)
            .map(|row| {
                let value = row[2].to_string();
                (value.is_empty(), value)
            })
            .collect::<Vec<_>>(),
        r_ref_anom
            .rows()
            .skip(52)
            .take(2)
            .map(|row| {
                let value = row[2].to_string();
                (value.is_empty(), value)
            })
            .collect::<Vec<_>>(),
    ] {
        let counts = references
            .iter()
            .fold(HashMap::new(), |mut counts, (_, value)| {
                *counts.entry(value).or_insert(0) += 1;
                counts
            });
        let keys: Vec<_> = references
            .iter()
            .map(|(empty, value)| (*empty, counts[value] == 1, value))
            .collect();
        assert!(keys.windows(2).all(|pair| pair[0] <= pair[1]));
    }

    // Verify 异常发票明细 has 38 rows (Row 6-43)
    let r_inv_anom = wb1.worksheet_range("异常发票明细").unwrap();
    let inv_anom_cnt = r_inv_anom
        .rows()
        .enumerate()
        .filter(|(idx, _)| {
            let row_num = idx + 1;
            (6..=43).contains(&row_num)
        })
        .count();
    assert_eq!(inv_anom_cnt, 38);

    // 2. Run Workbook 2
    generate_store_finance_workbook(&source_dir, &wb2_path).expect("Generate WB2 failed");
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
    let r_occ = wb2
        .worksheet_range("1.门店国补发生表（银联系统直接导出，不需要加工）")
        .unwrap();
    assert_eq!(r_occ.rows().count(), 13972);
    assert_eq!(r_occ.rows().next().unwrap().len(), 26);
    assert!(matches!(r_occ.get_value((1, 0)), Some(Data::DateTime(_))));

    // Verify Sheet 3: 2.门店累计回款表 (8,840 data rows + 1 header = 8,841)
    let r_ref = wb2
        .worksheet_range("2.门店累计回款表（事业部财务下发，门店筛选自己的）")
        .unwrap();
    assert_eq!(r_ref.rows().count(), 8841);
    assert_eq!(r_ref.rows().next().unwrap().len(), 24);

    // Verify Sheet 4: 3.门店上传明细 (12,010 data rows + 1 header = 12,011)
    let r_up = wb2
        .worksheet_range("3.门店上传明细（从门店银联后台每月导出后汇总）")
        .unwrap();
    assert_eq!(r_up.rows().count(), 12011);
    assert_eq!(r_up.rows().next().unwrap().len(), 60);
    assert_eq!(r_up.rows().next().unwrap()[59], "开票日期");

    // Verify Sheet 1: 最终匹配表
    let r_final = wb2
        .worksheet_range("最终匹配表（全部的国补发生数据上匹配）")
        .unwrap();
    assert_eq!(r_final.rows().count(), 13972); // 13,971 data rows + 1 header
    assert_eq!(r_final.rows().next().unwrap().len(), 27);
    assert!(matches!(r_final.get_value((1, 1)), Some(Data::DateTime(_))));

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
