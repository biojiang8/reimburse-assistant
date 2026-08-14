//! 报销助手 - Tauri 后端
//!
//! 职责：窗口/托盘/快捷键管理、文件对话框、调用 Python helper 处理发票
//! 与试剂耗材汇总（与 Electron 版共用同一套 Python 脚本）。

use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::Mutex;
use std::time::Duration;

use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use tauri::{
    menu::{Menu, MenuItem},
    tray::TrayIconBuilder,
    AppHandle, Emitter, Manager, WindowEvent,
};

const TOOL_TIMEOUT: Duration = Duration::from_secs(180);
const PYTHON_CANDIDATES: &[&str] = &[
    "python3",
    "python",
    "/Library/Frameworks/Python.framework/Versions/3.12/bin/python3",
];

/// 诊断日志：写入 ~/Library/Logs/reimburse-assistant.log
fn write_log(message: &str) {
    use std::io::Write;
    let path = dirs::home_dir()
        .map(|h| h.join("Library").join("Logs").join("reimburse-assistant.log"));
    if let Some(path) = path {
        if let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(&path)
        {
            let _ = writeln!(file, "[{}] {}", now_string(), message);
        }
    }
}

fn now_string() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0);
    let (h, m, s) = ((secs / 3600) % 24, (secs / 60) % 60, secs % 60);
    format!("{:02}:{:02}:{:02}", h, m, s)
}

// ---------------------------------------------------------------------------
// 状态与路径
// ---------------------------------------------------------------------------

struct AppState {
    /// 解析好的 Python 引擎（命令 + 工作目录）
    engine: Mutex<Option<Engine>>,
}

#[derive(Clone)]
struct Engine {
    command: String,
    cwd: PathBuf,
}

fn tool_root(app: &AppHandle) -> PathBuf {
    // 开发模式：tauri-app/src-tauri/helpers；打包后：resource dir/helpers
    let candidates = [
        app.path().resource_dir().ok().map(|d| d.join("helpers")),
        Some(Path::new(env!("CARGO_MANIFEST_DIR")).join("helpers")),
    ];
    candidates
        .into_iter()
        .flatten()
        .find(|p| p.is_dir())
        .unwrap_or_else(|| Path::new(env!("CARGO_MANIFEST_DIR")).join("helpers"))
}

/// 默认电子签路径：仅检查应用资源内的 signature.jpg（不再内置个人签名）。
/// 文件不存在时返回空路径，由前端要求用户自行选择电子签。
fn signature_path(app: &AppHandle) -> PathBuf {
    app.path()
        .resource_dir()
        .ok()
        .map(|d| d.join("helpers").join("signature.jpg"))
        .filter(|p| p.is_file())
        .unwrap_or_else(PathBuf::new)
}

fn default_output_dir() -> PathBuf {
    if let Some(docs) = dirs::document_dir() {
        return docs.join("报销助手");
    }
    PathBuf::from("报销助手")
}

// ---------------------------------------------------------------------------
// 子进程
// ---------------------------------------------------------------------------

struct ProcessResult {
    code: i32,
    stdout: String,
    stderr: String,
}

fn run_process(
    command: &str,
    args: &[String],
    cwd: &Path,
    envs: &[(&str, &str)],
    timeout: Duration,
) -> Result<ProcessResult, String> {
    let mut cmd = Command::new(command);
    cmd.args(args)
        .current_dir(cwd)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .env("PYTHONIOENCODING", "utf-8")
        .env("PYTHONUTF8", "1");
    for (key, value) in envs {
        cmd.env(key, value);
    }
    let mut child = cmd.spawn().map_err(|e| format!("无法启动进程：{e}"))?;
    let mut stdout = child.stdout.take().expect("stdout pipe");
    let mut stderr = child.stderr.take().expect("stderr pipe");
    let child = std::sync::Arc::new(Mutex::new(child));
    let (tx, rx) = std::sync::mpsc::channel::<ProcessResult>();

    let child_for_thread = child.clone();
    std::thread::spawn(move || {
        let mut out_str = String::new();
        let mut err_str = String::new();
        let _ = stdout.read_to_string(&mut out_str);
        let _ = stderr.read_to_string(&mut err_str);
        let code = child_for_thread
            .lock()
            .map(|mut c| c.wait().map(|s| s.code().unwrap_or(-1)).unwrap_or(-1))
            .unwrap_or(-1);
        let _ = tx.send(ProcessResult {
            code,
            stdout: out_str,
            stderr: err_str,
        });
    });

    match rx.recv_timeout(timeout) {
        Ok(result) => Ok(result),
        Err(_) => {
            let _ = child.lock().map(|mut c| c.kill());
            Err("处理超时，请检查 PDF 或订单表是否正常。".to_string())
        }
    }
}

fn parse_json_output(output: &str) -> Result<Value, String> {
    for line in output.lines().rev() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }
        if let Ok(value) = serde_json::from_str::<Value>(trimmed) {
            return Ok(value);
        }
    }
    Err("处理引擎返回了无法识别的结果。".to_string())
}

// ---------------------------------------------------------------------------
// Python 引擎解析
// ---------------------------------------------------------------------------

fn resolve_engine(app: &AppHandle) -> Result<Engine, String> {
    let state = app.state::<AppState>();
    if let Some(engine) = state.engine.lock().unwrap().clone() {
        return Ok(engine);
    }

    let root = tool_root(app);
    write_log(&format!(
        "resolve_engine: tool_root = {}, resource_dir = {:?}",
        root.display(),
        app.path().resource_dir().map(|d| d.display().to_string()).unwrap_or_default()
    ));

    // 1. PyInstaller 打包的 helper 二进制
    //    onedir: helpers/reimburse-helper/reimburse-helper；或平铺 helpers/reimburse-helper
    for name in ["reimburse-helper", "reimburse-helper.exe"] {
        for candidate in [
            root.join("reimburse-helper").join(name),
            root.join(name),
        ] {
            write_log(&format!(
                "resolve_engine: 检查 helper {} -> {}",
                candidate.display(),
                candidate.is_file()
            ));
            if candidate.is_file() {
                let engine = Engine {
                    command: candidate.to_string_lossy().to_string(),
                    cwd: root.clone(),
                };
                *state.engine.lock().unwrap() = Some(engine.clone());
                write_log(&format!("resolve_engine: 使用打包二进制 {}", engine.command));
                return Ok(engine);
            }
        }
    }

    // 2. 系统 Python + 脚本
    let scripts_ok = root.join("add_invoice_signature.py").is_file()
        && root.join("reagent_report.py").is_file();
    write_log(&format!("resolve_engine: 脚本可用 = {}", scripts_ok));
    if scripts_ok {
        for candidate in PYTHON_CANDIDATES {
            let probe = run_process(
                candidate,
                &[
                    "-c".into(),
                    "import fitz, PIL, openpyxl, docx; print('ready')".into(),
                ],
                &root,
                &[],
                Duration::from_secs(10),
            );
            match probe {
                Ok(result) if result.stdout.contains("ready") => {
                    let engine = Engine {
                        command: candidate.to_string(),
                        cwd: root.clone(),
                    };
                    *state.engine.lock().unwrap() = Some(engine.clone());
                    write_log(&format!("resolve_engine: 使用系统 Python {}", candidate));
                    return Ok(engine);
                }
                Ok(result) => write_log(&format!(
                    "resolve_engine: Python {} 探测失败 code={} stdout={:?} stderr={:?}",
                    candidate, result.code, result.stdout, result.stderr
                )),
                Err(error) => write_log(&format!(
                    "resolve_engine: Python {} 启动失败 {}",
                    candidate, error
                )),
            }
        }
        let message =
            "未找到可用的本地引擎。需要 Python、PyMuPDF、Pillow、openpyxl、python-docx；\
             或使用 helpers/reimburse-helper 打包二进制。"
                .to_string();
        write_log(&format!("resolve_engine: 失败 -> {}", message));
        return Err(message);
    }

    let message = format!(
        "未找到处理脚本：{}。请确认 helpers 目录包含 add_invoice_signature.py 与 reagent_report.py。",
        root.display()
    );
    write_log(&format!("resolve_engine: 失败 -> {}", message));
    Err(message)
}

fn run_helper(app: &AppHandle, args: &[String]) -> Result<ProcessResult, String> {
    let engine = resolve_engine(app)?;
    let root = tool_root(app);
    let is_binary = engine
        .command
        .contains("reimburse-helper");
    let mut full_args: Vec<String> = Vec::new();
    if is_binary {
        // helper 二进制：子命令 sign / reagent / summary / evidence / verify
        let first = args.first().map(String::as_str).unwrap_or("");
        let subcommand = if first == "--inspect-json" || first == "--json" {
            "sign"
        } else if first == "--summary-json" {
            "summary"
        } else if first == "--evidence-json" {
            "evidence"
        } else if first == "--verify-json" {
            "verify"
        } else {
            "reagent"
        };
        full_args.push(subcommand.to_string());
        full_args.extend(args.iter().cloned());
    } else {
        // Python 模式：python3 <对应脚本> ...
        let first = args.first().map(String::as_str).unwrap_or("");
        let script = if first == "--inspect-json" || first == "--json" {
            "add_invoice_signature.py"
        } else if first == "--summary-json" {
            "invoice_summary.py"
        } else if first == "--evidence-json" {
            "invoice_evidence.py"
        } else if first == "--verify-json" {
            "verify_package.py"
        } else {
            "reagent_report.py"
        };
        full_args.push(root.join(script).to_string_lossy().to_string());
        full_args.extend(args.iter().cloned());
    }
    write_log(&format!(
        "run_helper: {} {}",
        engine.command,
        full_args
            .iter()
            .map(|a| if a.contains(' ') { format!("{:?}", a) } else { a.clone() })
            .collect::<Vec<_>>()
            .join(" ")
    ));
    let result = run_process(&engine.command, &full_args, &engine.cwd, &[], TOOL_TIMEOUT)?;
    write_log(&format!(
        "run_helper: code={} stderr={:?}",
        result.code,
        result.stderr.chars().take(500).collect::<String>()
    ));
    Ok(result)
}

// ---------------------------------------------------------------------------
// 命令：默认值
// ---------------------------------------------------------------------------

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Defaults {
    output_directory: String,
    signature_path: String,
    engine_ready: bool,
    engine_label: String,
    engine_error: Option<String>,
}

#[tauri::command]
fn app_defaults(app: AppHandle) -> Defaults {
    let signature = signature_path(&app);
    let output = default_output_dir();
    std::fs::create_dir_all(&output).ok();
    match resolve_engine(&app) {
        Ok(engine) => Defaults {
            output_directory: output.to_string_lossy().to_string(),
            signature_path: signature.to_string_lossy().to_string(),
            engine_ready: true,
            engine_label: engine.command.clone(),
            engine_error: None,
        },
        Err(error) => Defaults {
            output_directory: output.to_string_lossy().to_string(),
            signature_path: signature.to_string_lossy().to_string(),
            engine_ready: false,
            engine_label: String::new(),
            engine_error: Some(error),
        },
    }
}

// ---------------------------------------------------------------------------
// 命令：文件对话框
// ---------------------------------------------------------------------------

#[tauri::command]
async fn select_files(app: AppHandle, kind: String) -> Result<Vec<String>, String> {
    use tauri_plugin_dialog::DialogExt;
    let builder = match kind.as_str() {
        "invoices" => app
            .dialog()
            .file()
            .set_title("选择电子发票")
            .add_filter("PDF 发票", &["pdf"]),
        "order" => app
            .dialog()
            .file()
            .set_title("选择供应商订单表")
            .add_filter("Excel 订单表", &["xlsx", "xls"]),
        "signature" => app
            .dialog()
            .file()
            .set_title("选择电子签图片")
            .add_filter("签名图片", &["jpg", "jpeg", "png"]),
        "photos" => app
            .dialog()
            .file()
            .set_title("选择实物照片")
            .add_filter("实物照片", &["jpg", "jpeg", "png"]),
        _ => return Err("未知的文件选择类型。".to_string()),
    };
    // 注意：必须用 async 命令 + blocking 对话框 API。同步命令运行在主线程，
    // 在主线程上调用 blocking_pick_* 会卡死整个应用（对话框永远无法弹出）。
    // async 命令运行在独立线程池，blocking_pick_* 在其上调用是官方推荐用法。
    let multi = kind == "invoices" || kind == "photos";
    if multi {
        let picked = builder
            .blocking_pick_files()
            .ok_or("未选择文件。".to_string())?;
        let paths: Vec<String> = picked
            .into_iter()
            .filter_map(|p| p.into_path().ok())
            .map(|p| p.to_string_lossy().to_string())
            .collect();
        write_log(&format!(
            "select_files({kind}): 选择了 {} 个文件 {:?}",
            paths.len(),
            paths
        ));
        Ok(paths)
    } else {
        let picked = builder
            .blocking_pick_file()
            .ok_or("未选择文件。".to_string())?;
        let paths = picked
            .into_path()
            .ok()
            .map(|p| vec![p.to_string_lossy().to_string()])
            .unwrap_or_default();
        write_log(&format!("select_files({kind}): 选择了 {:?}", paths));
        Ok(paths)
    }
}

#[tauri::command]
async fn select_directory(app: AppHandle, current: Option<String>) -> Result<Option<String>, String> {
    use tauri_plugin_dialog::DialogExt;
    let mut builder = app
        .dialog()
        .file()
        .set_title("选择文件夹");
    if let Some(path) = current {
        if PathBuf::from(&path).is_dir() {
            builder = builder.set_directory(&path);
        }
    }
    let picked = builder
        .blocking_pick_folder()
        .ok_or("未选择文件夹。".to_string())?;
    Ok(picked
        .into_path()
        .ok()
        .map(|p| p.to_string_lossy().to_string()))
}

// ---------------------------------------------------------------------------
// 命令：发票检查
// ---------------------------------------------------------------------------

#[tauri::command]
async fn inspect_invoices(app: AppHandle, paths: Vec<String>) -> Result<Value, String> {
    write_log(&format!(
        "inspect_invoices: 收到 {} 个路径 {:?}",
        paths.len(),
        paths
    ));
    let files = match validate_files(&paths, &[".pdf"]) {
        Ok(files) => files,
        Err(error) => {
            write_log(&format!("inspect_invoices: validate_files 失败 -> {error}"));
            return Err(error);
        }
    };
    let preview_dir = match preview_dir(&app) {
        Ok(dir) => dir,
        Err(error) => {
            write_log(&format!("inspect_invoices: preview_dir 失败 -> {error}"));
            return Err(error);
        }
    };
    write_log(&format!("inspect_invoices: 校验通过，preview_dir = {preview_dir}"));
    let mut args = vec![
        "--inspect-json".to_string(),
        "--company-side".to_string(),
        "seller".to_string(),
        "--amount-field".to_string(),
        "total".to_string(),
        "--preview-dir".to_string(),
        preview_dir,
    ];
    args.extend(files.iter().cloned());
    let result = match run_helper(&app, &args) {
        Ok(result) => result,
        Err(error) => {
            write_log(&format!("inspect_invoices: run_helper 失败 -> {error}"));
            return Err(error);
        }
    };
    if result.code != 0 {
        let error = if result.stderr.trim().is_empty() {
            result.stdout
        } else {
            result.stderr
        };
        write_log(&format!("inspect_invoices: 引擎返回非零 code={} -> {error}", result.code));
        return Err(error);
    }
    match parse_json_output(&result.stdout) {
        Ok(value) => Ok(value),
        Err(error) => {
            write_log(&format!("inspect_invoices: JSON 解析失败 -> {error}; stdout={:?}", result.stdout.chars().take(500).collect::<String>()));
            Err(error)
        }
    }
}

// ---------------------------------------------------------------------------
// 命令：发票处理
// ---------------------------------------------------------------------------

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ProcessPayload {
    files: Vec<String>,
    output_directory: Option<String>,
    signature_path: Option<String>,
    overwrite: Option<bool>,
    project_name: Option<String>,
    project_code: Option<String>,
}

#[tauri::command]
async fn process_invoices(
    app: AppHandle,
    window: tauri::Window,
    payload: ProcessPayload,
) -> Result<Value, String> {
    let files = validate_files(&payload.files, &[".pdf"])?;
    let output = validate_directory(
        payload
            .output_directory
            .unwrap_or_else(|| default_output_dir().to_string_lossy().to_string()),
    )?;
    let preview_dir = preview_dir(&app)?;
    let signature = payload
        .signature_path
        .map(|p| validate_files(&[p], &[".jpg", ".jpeg", ".png"]).map(|v| v[0].clone()))
        .transpose()?
        .unwrap_or_else(|| signature_path(&app).to_string_lossy().to_string());
    if !Path::new(&signature).is_file() {
        return Err("未找到电子签图片，请先在界面点击「电子签」选择签名图片。".to_string());
    }

    let mut results = Vec::new();
    let total = files.len();
    for (index, input) in files.iter().enumerate() {
        let _ = window.emit(
            "invoice:progress",
            json!({ "input": input, "index": index, "total": total, "status": "processing" }),
        );

        let mut args = vec![
            "--json".to_string(),
            "--company-side".to_string(),
            "seller".to_string(),
            "--amount-field".to_string(),
            "total".to_string(),
            "--output-dir".to_string(),
            output.clone(),
            "--signature".to_string(),
            signature.clone(),
            "--preview-dir".to_string(),
            preview_dir.clone(),
        ];
        if payload.overwrite.unwrap_or(false) {
            args.push("--overwrite".to_string());
        }
        args.push(input.clone());

        let mut outcome: Value = match run_helper(&app, &args) {
            Ok(result) if result.code == 0 => match parse_json_output(&result.stdout) {
                Ok(value) if value.get("ok").and_then(|v| v.as_bool()).unwrap_or(false) => {
                    value
                }
                Ok(value) => json!({ "ok": false, "input": input, "error": value.get("error").and_then(|v| v.as_str()).unwrap_or("处理失败") }),
                Err(error) => json!({ "ok": false, "input": input, "error": error }),
            },
            Ok(result) => {
                // 引擎出错时错误信息可能在 stdout（JSON 的 error 字段）而非 stderr
                let error = if result.stderr.trim().is_empty() {
                    parse_json_output(&result.stdout)
                        .ok()
                        .and_then(|v| v.get("error").and_then(|e| e.as_str()).map(String::from))
                        .unwrap_or_else(|| {
                            let message = result.stdout.trim();
                            if message.is_empty() {
                                "处理失败".to_string()
                            } else {
                                message.chars().take(300).collect()
                            }
                        })
                } else {
                    result.stderr.trim().chars().take(300).collect()
                };
                json!({ "ok": false, "input": input, "error": error })
            }
            Err(error) => json!({ "ok": false, "input": input, "error": error }),
        };

        // 每张发票签字成功后，同时生成配套的「发票明细 Excel」
        if outcome.get("ok").and_then(|v| v.as_bool()).unwrap_or(false) {
            match write_invoice_excel(
                &app,
                &output,
                &outcome,
                payload.project_name.as_deref().unwrap_or(""),
                payload.project_code.as_deref().unwrap_or(""),
            ) {
                Ok(excel_path) => {
                    if let Some(object) = outcome.as_object_mut() {
                        object.insert("excel".to_string(), Value::String(excel_path));
                    }
                }
                Err(error) => {
                    write_log(&format!(
                        "process_invoices: {} 明细 Excel 生成失败 -> {error}",
                        input
                    ));
                    if let Some(object) = outcome.as_object_mut() {
                        object.insert("excel_error".to_string(), Value::String(error));
                    }
                }
            }
        }

        let ok = outcome.get("ok").and_then(|v| v.as_bool()).unwrap_or(false);
        let _ = window.emit(
            "invoice:progress",
            json!({ "input": input, "index": index, "total": total, "status": if ok { "done" } else { "error" }, "result": outcome }),
        );
        results.push(outcome);
    }

    Ok(json!({ "output_directory": output, "results": results }))
}

/// 为一张已签字的发票生成配套「发票明细 Excel」，返回其路径。
fn write_invoice_excel(
    app: &AppHandle,
    output_dir: &str,
    result: &Value,
    project_name: &str,
    project_code: &str,
) -> Result<String, String> {
    let invoice_number = result
        .get("invoice_number")
        .and_then(|v| v.as_str())
        .map(|n| format!("dzfp{n}"))
        .unwrap_or_default();
    let filename = result
        .get("suggested_filename")
        .and_then(|v| v.as_str())
        .map(String::from)
        .or_else(|| {
            result
                .get("output")
                .and_then(|v| v.as_str())
                .and_then(|p| Path::new(p).file_name())
                .map(|n| n.to_string_lossy().to_string())
        })
        .unwrap_or_default();
    let record = json!({
        "project_name": project_name,
        "project_code": project_code,
        "invoice_number": invoice_number,
        "seller": result.get("seller").cloned().unwrap_or(Value::Null),
        "seller_tax_id": result.get("seller_tax_id").cloned().unwrap_or(Value::Null),
        "buyer": result.get("buyer").cloned().unwrap_or(Value::Null),
        "buyer_tax_id": result.get("buyer_tax_id").cloned().unwrap_or(Value::Null),
        "invoice_date": result.get("invoice_date").cloned().unwrap_or(Value::Null),
        "subtotal": result.get("subtotal").cloned().unwrap_or(Value::Null),
        "total": result.get("total").cloned().unwrap_or(Value::Null),
        "items": result.get("items").cloned().unwrap_or_else(|| json!([])),
        "filename": filename,
    });

    let cache_dir = app
        .path()
        .app_cache_dir()
        .map_err(|e| format!("无法获取缓存目录：{e}"))?;
    std::fs::create_dir_all(&cache_dir)
        .map_err(|e| format!("无法创建缓存目录 {}：{e}", cache_dir.display()))?;
    let json_path = cache_dir.join("invoice-excel-input.json");
    let json_text =
        serde_json::to_string(&record).map_err(|e| format!("明细数据序列化失败：{e}"))?;
    std::fs::write(&json_path, json_text)
        .map_err(|e| format!("无法写入明细数据 {}：{e}", json_path.display()))?;

    let xlsx_name = derive_package_name(&record, "发票明细.xlsx")
        .ok_or_else(|| "无法从发票信息推导明细 Excel 文件名。".to_string())?;
    let xlsx_path = Path::new(output_dir).join(&xlsx_name);
    let args = vec![
        "--summary-json".to_string(),
        "--json-file".to_string(),
        json_path.to_string_lossy().to_string(),
        "--output".to_string(),
        xlsx_path.to_string_lossy().to_string(),
    ];
    let result = run_helper(&app, &args)?;
    if result.code != 0 {
        return Err(if result.stderr.trim().is_empty() {
            result.stdout
        } else {
            result.stderr
        });
    }
    Ok(xlsx_path.to_string_lossy().to_string())
}

/// 从发票信息推导三件套配套文件名：
/// 「公司-金额-dzfp票号-{后缀}」，优先复用建议文件名里的前缀。
fn derive_package_name(record: &Value, suffix: &str) -> Option<String> {
    let base = record
        .get("filename")
        .and_then(|v| v.as_str())
        .unwrap_or("");
    let stem = base
        .strip_suffix(".pdf")
        .unwrap_or(base)
        .strip_suffix("-已签字")
        .unwrap_or_else(|| base.strip_suffix(".pdf").unwrap_or(base))
        .trim_end_matches('-');
    if !stem.is_empty() {
        return Some(format!("{stem}-{suffix}"));
    }
    let seller = record.get("seller").and_then(|v| v.as_str()).unwrap_or("");
    let total = record.get("total").and_then(|v| v.as_str()).unwrap_or("");
    let number = record
        .get("invoice_number")
        .and_then(|v| v.as_str())
        .unwrap_or("");
    if seller.is_empty() || total.is_empty() || number.is_empty() {
        return None;
    }
    let clean = |s: &str| {
        s.replace(['/', '\\', ':', '*', '?', '"', '<', '>', '|'], "_")
            .trim()
            .to_string()
    };
    Some(format!(
        "{}-{}-dzfp{}-{suffix}",
        clean(seller),
        clean(total),
        clean(number)
    ))
}

/// 派生实物证据 Word 文件名（与 PDF 前缀一致，后缀替换）。
fn derive_evidence_name(suggested_filename: &str) -> Option<String> {
    let stem = suggested_filename
        .strip_suffix("已签字.pdf")
        .or_else(|| suggested_filename.strip_suffix(".pdf"))
        .map(|s| s.trim_end_matches('-'));
    stem.map(|s| format!("{s}-实物证据.docx"))
}

// ---------------------------------------------------------------------------
// 命令：项目台账（记忆项目名称 + 经费代码）
// ---------------------------------------------------------------------------

#[derive(Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct ProjectEntry {
    name: String,
    code: String,
}

fn projects_file(app: &AppHandle) -> Result<PathBuf, String> {
    app.path()
        .app_config_dir()
        .map(|dir| dir.join("projects.json"))
        .map_err(|e| format!("无法获取配置目录：{e}"))
}

#[tauri::command]
fn load_projects(app: AppHandle) -> Result<Value, String> {
    let path = projects_file(&app)?;
    if !path.is_file() {
        return Ok(json!({ "projects": [] }));
    }
    let text = std::fs::read_to_string(&path)
        .map_err(|e| format!("无法读取项目台账 {}：{e}", path.display()))?;
    serde_json::from_str(&text).map_err(|e| format!("项目台账解析失败：{e}"))
}

#[tauri::command]
fn save_projects(app: AppHandle, projects: Vec<ProjectEntry>) -> Result<(), String> {
    let path = projects_file(&app)?;
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)
            .map_err(|e| format!("无法创建配置目录 {}：{e}", parent.display()))?;
    }
    let text = serde_json::to_string_pretty(&projects)
        .map_err(|e| format!("项目台账序列化失败：{e}"))?;
    std::fs::write(&path, text)
        .map_err(|e| format!("无法写入项目台账 {}：{e}", path.display()))
}

// ---------------------------------------------------------------------------
// 命令：实物证据 / 槽位校验
// ---------------------------------------------------------------------------

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct EvidencePayload {
    invoice: Value,
    photos: Vec<String>,
    output_directory: Option<String>,
    excel_path: Option<String>,
    project_name: Option<String>,
    project_code: Option<String>,
}

#[tauri::command]
async fn generate_evidence(app: AppHandle, payload: EvidencePayload) -> Result<Value, String> {
    let photos = validate_files(&payload.photos, &[".jpg", ".jpeg", ".png"])?;
    if photos.is_empty() {
        return Err("请先添加实物照片。".to_string());
    }
    let output = validate_directory(
        payload
            .output_directory
            .unwrap_or_else(|| default_output_dir().to_string_lossy().to_string()),
    )?;

    let invoice_number = payload
        .invoice
        .get("invoice_number")
        .and_then(|v| v.as_str())
        .unwrap_or("");
    let suggested = payload
        .invoice
        .get("suggested_filename")
        .and_then(|v| v.as_str())
        .unwrap_or("");
    let word_name = derive_evidence_name(suggested)
        .ok_or_else(|| "无法从发票信息推导实物证据文件名。".to_string())?;
    let word_path = Path::new(&output).join(&word_name);

    let record = json!({
        "invoice_number": format!("dzfp{invoice_number}"),
        "seller": payload.invoice.get("seller").cloned().unwrap_or(Value::Null),
        "seller_tax_id": payload.invoice.get("seller_tax_id").cloned().unwrap_or(Value::Null),
        "total": payload.invoice.get("total").cloned().unwrap_or(Value::Null),
        "invoice_date": payload.invoice.get("invoice_date").cloned().unwrap_or(Value::Null),
        "photos": photos,
    });

    let cache_dir = app
        .path()
        .app_cache_dir()
        .map_err(|e| format!("无法获取缓存目录：{e}"))?;
    std::fs::create_dir_all(&cache_dir)
        .map_err(|e| format!("无法创建缓存目录 {}：{e}", cache_dir.display()))?;
    let json_path = cache_dir.join("invoice-evidence-input.json");
    let json_text = serde_json::to_string(&record).map_err(|e| format!("证据数据序列化失败：{e}"))?;
    std::fs::write(&json_path, json_text)
        .map_err(|e| format!("无法写入证据数据 {}：{e}", json_path.display()))?;

    let args = vec![
        "--evidence-json".to_string(),
        "--json-file".to_string(),
        json_path.to_string_lossy().to_string(),
        "--output".to_string(),
        word_path.to_string_lossy().to_string(),
    ];
    let result = run_helper(&app, &args)?;
    if result.code != 0 {
        return Err(if result.stderr.trim().is_empty() {
            result.stdout
        } else {
            result.stderr
        });
    }
    write_log(&format!(
        "generate_evidence: {} 张照片 -> {}",
        record.get("photos").and_then(|v| v.as_array()).map(|a| a.len()).unwrap_or(0),
        word_path.display()
    ));

    // 同步更新配套的发票明细 Excel（写入照片文件列）
    let updated_excel = match payload.excel_path {
        Some(path) if !path.is_empty() => {
            let path = PathBuf::from(path);
            let excel_record = json!({
                "project_name": payload.project_name.clone().unwrap_or_default(),
                "project_code": payload.project_code.clone().unwrap_or_default(),
                "invoice_number": format!("dzfp{invoice_number}"),
                "seller": payload.invoice.get("seller").cloned().unwrap_or(Value::Null),
                "seller_tax_id": payload.invoice.get("seller_tax_id").cloned().unwrap_or(Value::Null),
                "buyer": payload.invoice.get("buyer").cloned().unwrap_or(Value::Null),
                "buyer_tax_id": payload.invoice.get("buyer_tax_id").cloned().unwrap_or(Value::Null),
                "invoice_date": payload.invoice.get("invoice_date").cloned().unwrap_or(Value::Null),
                "subtotal": payload.invoice.get("subtotal").cloned().unwrap_or(Value::Null),
                "total": payload.invoice.get("total").cloned().unwrap_or(Value::Null),
                "items": payload.invoice.get("items").cloned().unwrap_or_else(|| json!([])),
                "filename": suggested,
            });
            let json_path = cache_dir.join("invoice-excel-input.json");
            let json_text = serde_json::to_string(&excel_record)
                .map_err(|e| format!("明细数据序列化失败：{e}"))?;
            std::fs::write(&json_path, json_text)
                .map_err(|e| format!("无法写入明细数据 {}：{e}", json_path.display()))?;
            let args = vec![
                "--summary-json".to_string(),
                "--json-file".to_string(),
                json_path.to_string_lossy().to_string(),
                "--output".to_string(),
                path.to_string_lossy().to_string(),
            ];
            let result = run_helper(&app, &args)?;
            if result.code != 0 {
                return Err(if result.stderr.trim().is_empty() {
                    result.stdout
                } else {
                    result.stderr
                });
            }
            write_log(&format!(
                "generate_evidence: 明细 Excel 已更新 -> {}",
                path.display()
            ));
            Some(path.to_string_lossy().to_string())
        }
        _ => None,
    };

    Ok(json!({
        "ok": true,
        "output": word_path.to_string_lossy().to_string(),
        "excel": updated_excel,
        "photos": record.get("photos").and_then(|v| v.as_array()).map(|a| a.len()).unwrap_or(0),
    }))
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct VerifyItem {
    invoice_number: String,
    pdf: Option<String>,
    excel: Option<String>,
    word: Option<String>,
    photos: Option<Vec<String>>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct VerifyPayload {
    items: Vec<VerifyItem>,
}

#[tauri::command]
async fn verify_package(app: AppHandle, payload: VerifyPayload) -> Result<Value, String> {
    if payload.items.is_empty() {
        return Err("没有可校验的发票。".to_string());
    }
    let items: Vec<Value> = payload
        .items
        .iter()
        .map(|item| {
            json!({
                "invoice_number": item.invoice_number,
                "pdf": item.pdf.clone().unwrap_or_default(),
                "excel": item.excel.clone().unwrap_or_default(),
                "word": item.word.clone().unwrap_or_default(),
                "photos": item.photos.clone().unwrap_or_default(),
            })
        })
        .collect();

    let cache_dir = app
        .path()
        .app_cache_dir()
        .map_err(|e| format!("无法获取缓存目录：{e}"))?;
    std::fs::create_dir_all(&cache_dir)
        .map_err(|e| format!("无法创建缓存目录 {}：{e}", cache_dir.display()))?;
    let json_path = cache_dir.join("verify-input.json");
    let json_text = serde_json::to_string(&items).map_err(|e| format!("校验数据序列化失败：{e}"))?;
    std::fs::write(&json_path, json_text)
        .map_err(|e| format!("无法写入校验数据 {}：{e}", json_path.display()))?;

    let args = vec![
        "--verify-json".to_string(),
        "--json-file".to_string(),
        json_path.to_string_lossy().to_string(),
    ];
    let result = run_helper(&app, &args)?;
    if result.code != 0 {
        return Err(if result.stderr.trim().is_empty() {
            result.stdout
        } else {
            result.stderr
        });
    }
    parse_json_output(&result.stdout)
}

// ---------------------------------------------------------------------------
// 命令：打开路径
// ---------------------------------------------------------------------------

#[tauri::command]
fn open_path(value: String) -> Result<bool, String> {
    let path = PathBuf::from(&value);
    if !path.exists() {
        return Err("路径不存在。".to_string());
    }
    open_with_system(&path)
}

/// 读取图片并返回 base64 data URL（Tauri 的 WKWebView 不允许直接加载
/// file:// 图片，预览图以 data URL 方式传给前端）。
#[tauri::command]
fn read_image_base64(path: String) -> Result<String, String> {
    let path = PathBuf::from(&path);
    if !path.is_file() {
        return Err("图片不存在。".to_string());
    }
    let bytes = std::fs::read(&path).map_err(|e| format!("读取图片失败：{e}"))?;
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("png")
        .to_lowercase();
    let mime = match ext.as_str() {
        "jpg" | "jpeg" => "image/jpeg",
        "webp" => "image/webp",
        _ => "image/png",
    };
    use base64::Engine;
    Ok(format!(
        "data:{};base64,{}",
        mime,
        base64::engine::general_purpose::STANDARD.encode(&bytes)
    ))
}

#[tauri::command]
fn show_item(value: String) -> Result<bool, String> {
    let path = PathBuf::from(&value);
    if !path.exists() {
        return Err("文件不存在。".to_string());
    }
    #[cfg(target_os = "macos")]
    {
        let parent = path.parent().unwrap_or(&path);
        return open_with_system(parent);
    }
    #[cfg(not(target_os = "macos"))]
    open_with_system(&path)
}

fn open_with_system(path: &Path) -> Result<bool, String> {
    #[cfg(target_os = "macos")]
    let mut cmd = Command::new("open");
    #[cfg(target_os = "windows")]
    let mut cmd = Command::new("explorer");
    #[cfg(all(not(target_os = "macos"), not(target_os = "windows")))]
    let mut cmd = Command::new("xdg-open");
    cmd.arg(path);
    cmd.spawn()
        .map_err(|e| format!("无法打开路径：{e}"))?;
    Ok(true)
}

// ---------------------------------------------------------------------------
// 校验辅助
// ---------------------------------------------------------------------------

fn validate_files(values: &[String], allowed: &[&str]) -> Result<Vec<String>, String> {
    if values.is_empty() || values.len() > 200 {
        return Err("请选择 1 至 200 个文件。".to_string());
    }
    let mut resolved = Vec::new();
    for value in values {
        let path = PathBuf::from(value);
        let ext = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_lowercase())
            .unwrap_or_default();
        // 允许列表里的扩展名可能带点（如 ".pdf"），统一去掉前导点再比较。
        let matched = allowed
            .iter()
            .any(|a| a.trim_start_matches('.').eq_ignore_ascii_case(&ext));
        if !matched {
            return Err(format!("不支持的文件类型：{}", path.display()));
        }
        if !path.is_file() {
            return Err(format!("文件不存在：{}", path.display()));
        }
        resolved.push(path.to_string_lossy().to_string());
    }
    Ok(resolved)
}

fn validate_directory(value: String) -> Result<String, String> {
    let path = PathBuf::from(&value);
    std::fs::create_dir_all(&path).map_err(|e| format!("无法创建目录：{e}"))?;
    if !path.is_dir() {
        return Err("输出路径不是目录。".to_string());
    }
    Ok(path.to_string_lossy().to_string())
}

fn preview_dir(app: &AppHandle) -> Result<String, String> {
    let dir = app
        .path()
        .app_cache_dir()
        .map(|d| d.join("reimburse-previews"))
        .unwrap_or_else(|_| std::env::temp_dir().join("reimburse-previews"));
    std::fs::create_dir_all(&dir).map_err(|e| format!("无法创建预览目录：{e}"))?;
    Ok(dir.to_string_lossy().to_string())
}

// ---------------------------------------------------------------------------
// 托盘 / 快捷键 / 窗口行为
// ---------------------------------------------------------------------------

fn show_main_window(app: &AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.show();
        let _ = window.set_focus();
    }
}

fn setup_tray(app: &AppHandle) -> tauri::Result<()> {
    let show = MenuItem::with_id(app, "show", "打开报销助手", true, None::<&str>)?;
    let quit = MenuItem::with_id(app, "quit", "退出", true, None::<&str>)?;
    let menu = Menu::with_items(app, &[&show, &quit])?;

    TrayIconBuilder::with_id("main-tray")
        .tooltip("报销助手")
        .menu(&menu)
        .show_menu_on_left_click(false)
        .on_menu_event(move |app, event| match event.id.as_ref() {
            "show" => show_main_window(app),
            "quit" => app.exit(0),
            _ => {}
        })
        .on_tray_icon_event(|tray, event| {
            use tauri::tray::TrayIconEvent;
            if let TrayIconEvent::Click {
                button: tauri::tray::MouseButton::Left,
                ..
            } = event
            {
                show_main_window(&tray.app_handle());
            }
        })
        .build(app)?;

    use tauri_plugin_global_shortcut::{Code, GlobalShortcutExt, Modifiers, Shortcut, ShortcutState};
    let shortcut = Shortcut::new(Some(Modifiers::SUPER | Modifiers::SHIFT), Code::KeyR);
    let _ = app.global_shortcut().on_shortcut(shortcut, move |app, _shortcut, event| {
        if event.state == ShortcutState::Pressed {
            show_main_window(app);
        }
    });
    Ok(())
}

// ---------------------------------------------------------------------------
// 入口
// ---------------------------------------------------------------------------

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_global_shortcut::Builder::new().build())
        .manage(AppState {
            engine: Mutex::new(None),
        })
        .invoke_handler(tauri::generate_handler![
            app_defaults,
            select_files,
            select_directory,
            inspect_invoices,
            process_invoices,
            generate_evidence,
            verify_package,
            load_projects,
            save_projects,
            open_path,
            show_item,
            read_image_base64,
        ])
        .setup(|app| {
            setup_tray(app.handle())?;
            let window = app.get_webview_window("main").expect("main window");
            let window_handle = window.clone();
            window.on_window_event(move |event| {
                if let WindowEvent::CloseRequested { api, .. } = event {
                    api.prevent_close();
                    let _ = window_handle.hide();
                }
            });
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
