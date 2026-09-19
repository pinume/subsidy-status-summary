use std::fs::{self, OpenOptions};
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::time::Instant;

#[cfg(windows)]
use std::os::windows::ffi::OsStrExt;

use data_status_summary::{generate_store_finance_workbook, generate_summary_workbook};

const REQUIRED_FILES: [&str; 8] = [
    "销售用券情况统计.xlsx",
    "已上传家电电脑.xlsx",
    "已上传数码.xlsx",
    "回款明细家电电脑.xlsx",
    "回款明细数码.xlsx",
    "发票明细.xlsx",
    "银联交易明细所有.xlsx",
    "银联交易明细门店.xlsx",
];

fn main() {
    let input_dir = match prompt_input_directory() {
        Some(dir) => dir,
        None => {
            println!("程序已安全退出。");
            return;
        }
    };

    let output_dir = derive_output_directory(&input_dir);
    if let Err(e) = fs::create_dir_all(&output_dir) {
        eprintln!("创建输出目录 {:?} 失败: {}", output_dir, e);
        return;
    }

    println!();
    loop {
        println!("1. 国补上传情况汇总");
        println!("2. 26年国补门店财务统筹表");
        println!("0. Exit");
        print!("请输入选项编号 [0-2]: ");
        io::stdout().flush().unwrap();

        let mut choice = String::new();
        if io::stdin().read_line(&mut choice).is_err() {
            continue;
        }
        let choice = choice.trim();

        match choice {
            "1" => {
                run_job(
                    "国补上传情况汇总.xlsx",
                    &input_dir,
                    &output_dir,
                    generate_summary_workbook,
                );
                wait_for_return();
            }
            "2" => {
                run_job(
                    "26年国补门店财务统筹表.xlsx",
                    &input_dir,
                    &output_dir,
                    generate_store_finance_workbook,
                );
                wait_for_return();
            }
            "0" | "exit" | "quit" => {
                println!("程序已安全退出。");
                return;
            }
            _ => {
                println!("[错误] 无效选项，请输入 0 至 2 之间的数字编号！\n");
            }
        }
    }
}

fn prompt_input_directory() -> Option<PathBuf> {
    loop {
        print!("请输入源数据文件夹路径: ");
        io::stdout().flush().unwrap();

        let mut input = String::new();
        if io::stdin().read_line(&mut input).is_err() {
            continue;
        }

        let input = input.trim().trim_matches(['\'', '"']);
        if input == "0" || input.eq_ignore_ascii_case("exit") || input.eq_ignore_ascii_case("quit")
        {
            return None;
        }

        if input.is_empty() {
            continue;
        }

        let path = PathBuf::from(input);
        if !path.exists() {
            eprintln!("[错误] 路径不存在: {:?}", path);
            continue;
        }
        if !path.is_dir() {
            eprintln!("[错误] 所输路径不是文件夹: {:?}", path);
            continue;
        }

        // Validate required files
        let mut missing = Vec::new();
        for req in &REQUIRED_FILES {
            let file_path = path.join(req);
            if !file_path.exists() || !file_path.is_file() {
                missing.push(*req);
            }
        }

        if !missing.is_empty() {
            eprintln!("[错误] 该目录缺失以下关键源数据文件:");
            for f in missing {
                eprintln!("  - {}", f);
            }
            continue;
        }

        return Some(path);
    }
}

fn derive_output_directory(input_dir: &Path) -> PathBuf {
    let parent = input_dir.parent().unwrap_or(input_dir);
    parent.join("subsidy_summary")
}

fn wait_for_return() {
    print!("\n按回车键返回... ");
    io::stdout().flush().unwrap();
    let mut buf = String::new();
    let _ = io::stdin().read_line(&mut buf);
    println!();
}

fn run_job(
    name: &str,
    input_dir: &Path,
    output_dir: &Path,
    generator: impl FnOnce(&Path, &Path) -> Result<(), String>,
) {
    println!("\n正在处理: {} ...", name);
    let start = Instant::now();
    let target = output_dir.join(name);
    let tmp = output_dir.join(format!(".tmp_{}_{}", std::process::id(), name));
    let lock_path = output_dir.join(format!(".lock_{}", name));
    // ponytail: a crashed process can leave this lock behind; use OS locks if that becomes common.
    let _lock = match OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&lock_path)
    {
        Ok(lock) => lock,
        Err(e) => {
            eprintln!("已有相同报表任务正在运行，无法开始处理: {}", e);
            return;
        }
    };

    match generator(input_dir, &tmp) {
        Ok(_) => {
            if let Err(e) = replace_file(&tmp, &target) {
                eprintln!("保存失败: {}", e);
                let _ = fs::remove_file(&tmp);
            } else {
                println!("已完成 (耗时: {:.2?})", start.elapsed());
            }
        }
        Err(e) => {
            eprintln!("处理失败: {}", e);
            let _ = fs::remove_file(&tmp);
        }
    }
    let _ = fs::remove_file(lock_path);
}

#[cfg(not(windows))]
fn replace_file(source: &Path, target: &Path) -> io::Result<()> {
    fs::rename(source, target)
}

#[cfg(windows)]
fn replace_file(source: &Path, target: &Path) -> io::Result<()> {
    const MOVEFILE_REPLACE_EXISTING: u32 = 0x1;
    const MOVEFILE_WRITE_THROUGH: u32 = 0x8;

    #[link(name = "kernel32")]
    unsafe extern "system" {
        fn MoveFileExW(source: *const u16, target: *const u16, flags: u32) -> i32;
    }

    let source: Vec<u16> = source.as_os_str().encode_wide().chain(Some(0)).collect();
    let target: Vec<u16> = target.as_os_str().encode_wide().chain(Some(0)).collect();
    // SAFETY: both paths are NUL-terminated and remain alive for the duration of the call.
    let result = unsafe {
        MoveFileExW(
            source.as_ptr(),
            target.as_ptr(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH,
        )
    };
    if result == 0 {
        Err(io::Error::last_os_error())
    } else {
        Ok(())
    }
}
