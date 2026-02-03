use std::collections::HashMap;
use std::env;
use std::fs;
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::process::Command;
use calamine::{open_workbook, Reader, Xlsx};
use log::{error, info, warn};

fn find_project_root() -> Option<PathBuf> {
    let mut current = env::current_dir().ok()?;
    
    loop {
        let cargo_toml = current.join("Cargo.toml");
        if cargo_toml.exists() {
            return Some(current);
        }
        
        match current.parent() {
            Some(parent) => current = parent.to_path_buf(),
            None => return None,
        }
    }
}

fn check_qpdf_installed() -> bool {
    match Command::new("qpdf").arg("--version").output() {
        Ok(output) => output.status.success(),
        Err(_) => false,
    }
}

fn main() {
    // Initialize logger with default level if not set
    if env::var("RUST_LOG").is_err() {
        env::set_var("RUST_LOG", "info");
    }
    env_logger::init();
    
    println!("Starting PDF page insertion tool...");
    
    // Check if qpdf is installed
    if !check_qpdf_installed() {
        println!("\n=== ERROR ===");
        println!("qpdf is not installed or not in PATH.");
        println!("Please install qpdf:");
        println!("  - Windows: Download from https://github.com/qpdf/qpdf/releases");
        println!("  - Or use: choco install qpdf");
        println!("  - Or use: winget install qpdf");
        println!("  - Make sure qpdf is in your system PATH");
        error!("qpdf not found");
        return;
    }
    println!("✓ qpdf found");
    
    // Get current working directory (where compare.xlsx and bia.pdf should be)
    let source_dir = match find_project_root() {
        Some(dir) => dir,
        None => {
            match env::current_dir() {
                Ok(dir) => dir,
                Err(e) => {
                    error!("Failed to get current directory: {}", e);
                    return;
                }
            }
        }
    };
    
    // Validate required files exist in source directory
    let excel_path = source_dir.join("compare.xlsx");
    let bia_path = source_dir.join("bia.pdf");
    
    if !excel_path.exists() {
        error!("compare.xlsx not found in directory: {}", source_dir.display());
        println!("ERROR: compare.xlsx not found in: {}", source_dir.display());
        return;
    }
    
    if !bia_path.exists() {
        error!("bia.pdf not found in directory: {}", source_dir.display());
        println!("ERROR: bia.pdf not found in: {}", source_dir.display());
        return;
    }
    
    // Get page count from bia.pdf using pdfcpu
    println!("Loading bia.pdf from: {}", bia_path.display());
    let bia_page_count = match get_pdf_page_count(&bia_path) {
        Ok(count) => count,
        Err(e) => {
            error!("Failed to get page count from bia.pdf: {}", e);
            println!("ERROR: Failed to get page count from bia.pdf: {}", e);
            return;
        }
    };
    println!("bia.pdf has {} pages", bia_page_count);
    
    // Prompt for directory path (where PDF files to process are located)
    print!("Enter directory path: ");
    io::stdout().flush().unwrap();
    
    let mut input = String::new();
    io::stdin().read_line(&mut input).expect("Failed to read input");
    let dir_path = input.trim();
    
    if dir_path.is_empty() {
        error!("Directory path cannot be empty");
        return;
    }
    
    let base_dir = Path::new(dir_path);
    
    // Validate directory exists
    if !base_dir.exists() || !base_dir.is_dir() {
        error!("Directory does not exist: {}", dir_path);
        return;
    }
    
    info!("Reading compare.xlsx...");
    let mappings = match read_excel_mappings(&excel_path) {
        Ok(m) => m,
        Err(e) => {
            error!("Failed to read compare.xlsx: {}", e);
            println!("ERROR: Failed to read compare.xlsx: {}", e);
            return;
        }
    };
    
    info!("Found {} mappings in Excel file", mappings.len());
    
    // Scan child directories for PDF files
    let pdf_files = match scan_child_directories(base_dir) {
        Ok(files) => files,
        Err(e) => {
            error!("Failed to scan directories: {}", e);
            println!("ERROR: Failed to scan directories: {}", e);
            return;
        }
    };
    
    if pdf_files.is_empty() {
        warn!("No PDF files found in child directories");
        println!("ERROR: No PDF files found in child directories!");
        return;
    }
    
    info!("Found {} PDF files in subdirectories", pdf_files.len());
    println!("\nProcessing {} files...\n", pdf_files.len());
    
    // Process PDFs
    let mut processed = 0;
    let mut skipped = 0;
    let mut errors = 0;
    
    for pdf_path in pdf_files {
        match process_pdf_with_qpdf(&pdf_path, &bia_path, &mappings, bia_page_count) {
            Ok(true) => {
                processed += 1;
                let filename = pdf_path.file_name().and_then(|n| n.to_str()).unwrap_or("unknown");
                println!("✓ {}", filename);
                info!("Processed: {}", pdf_path.display());
            }
            Ok(false) => {
                skipped += 1;
                let filename = pdf_path.file_name().and_then(|n| n.to_str()).unwrap_or("unknown");
                println!("⊘ {} (skipped - no match in Excel)", filename);
                warn!("Skipped: {} (no match in Excel)", pdf_path.display());
            }
            Err(e) => {
                errors += 1;
                let filename = pdf_path.file_name().and_then(|n| n.to_str()).unwrap_or("unknown");
                println!("✗ {} - Error: {}", filename, e);
                error!("Error processing {}: {}", pdf_path.display(), e);
            }
        }
    }
    
    // Summary
    println!("\n=== Summary ===");
    println!("Processed: {}", processed);
    println!("Skipped: {}", skipped);
    println!("Errors: {}", errors);
    info!("Summary: {} processed, {} skipped, {} errors", processed, skipped, errors);
    
    // Keep terminal open for user to see results
    println!("\nPress Enter to close...");
    io::stdout().flush().unwrap();
    let mut _input = String::new();
    let _ = io::stdin().read_line(&mut _input);
}

fn get_pdf_page_count(pdf_path: &Path) -> Result<usize, Box<dyn std::error::Error>> {
    // Use qpdf to get page count
    let output = Command::new("qpdf")
        .args(["--show-npages", pdf_path.to_str().unwrap()])
        .output()?;
    
    if !output.status.success() {
        return Err(format!("qpdf failed: {}", String::from_utf8_lossy(&output.stderr)).into());
    }
    
    let stdout = String::from_utf8_lossy(&output.stdout);
    let count = stdout.trim().parse::<usize>()?;
    
    Ok(count)
}

fn read_excel_mappings(excel_path: &Path) -> Result<HashMap<String, u32>, Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook(excel_path)?;
    let mut mappings = HashMap::new();
    
    if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
        for row in range.rows() {
            if row.len() < 2 {
                continue;
            }
            
            // Column A: filename
            let filename_cell = &row[0];
            let filename = match filename_cell {
                calamine::Data::String(s) => s.trim().to_string(),
                calamine::Data::Float(f) => f.to_string(),
                calamine::Data::Int(i) => i.to_string(),
                _ => continue,
            };
            
            if filename.is_empty() {
                continue;
            }
            
            // Column B: page number
            let page_cell = &row[1];
            let page_num = match page_cell {
                calamine::Data::Int(i) => *i as u32,
                calamine::Data::Float(f) => *f as u32,
                _ => continue,
            };
            
            if page_num == 0 {
                continue;
            }
            
            // Store 0-based page index
            let page_index = page_num - 1;
            
            // Normalize filename: remove path, keep only filename
            let filename_only = Path::new(&filename)
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or(&filename)
                .to_string();
            
            mappings.insert(filename_only, page_index);
        }
    }
    
    Ok(mappings)
}

fn normalize_filename(filename: &str) -> String {
    // Remove path, keep only filename
    let filename_only = Path::new(filename)
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or(filename)
        .to_string();
    
    // Remove .pdf extension for comparison
    filename_only
        .strip_suffix(".pdf")
        .or_else(|| filename_only.strip_suffix(".PDF"))
        .unwrap_or(&filename_only)
        .to_string()
}

fn extract_base_name(filename: &str) -> String {
    let normalized = normalize_filename(filename);
    
    // Remove anything after and including parentheses: "hoa (1)" -> "hoa", "hoa(2)" -> "hoa"
    // This allows hoa.pdf, hoa (1).pdf, hoa (2).pdf, etc. to all match "hoa"
    let base = if let Some(pos) = normalized.find(" (") {
        normalized[..pos].trim().to_string()
    } else if let Some(pos) = normalized.find('(') {
        normalized[..pos].trim().to_string()
    } else {
        normalized
    };
    
    base
}

fn match_pdf_name(pdf_filename: &str, mappings: &HashMap<String, u32>) -> Option<u32> {
    let pdf_base = normalize_filename(pdf_filename);
    
    // Try exact match first: "hoa" matches "hoa"
    if let Some(&page) = mappings.get(&pdf_base) {
        return Some(page);
    }
    
    // Try with .pdf extension: "hoa" matches "hoa.pdf"
    let pdf_with_ext = format!("{}.pdf", pdf_base);
    if let Some(&page) = mappings.get(&pdf_with_ext) {
        return Some(page);
    }
    
    // Only match files with "(1)" - the first duplicate, ignore (2), (3), etc.
    // "hoa (1).pdf" -> extract base "hoa" and check if has "(1)"
    
    // Check if this is a "(1)" file (the first duplicate)
    let has_number_one = pdf_filename.contains("(1)") || pdf_filename.contains("(1).");
    
    if has_number_one {
        let pdf_base_name = extract_base_name(pdf_filename);
        
        // Check all mappings for exact base name match
        // "hoa (1).pdf" extracts "hoa", matches Excel "hoa"
        if let Some(&page) = mappings.get(&pdf_base_name) {
            return Some(page);
        }
        
        // Check if any Excel entry matches when we extract its base name
        for (excel_filename, &page) in mappings.iter() {
            let excel_base_name = extract_base_name(excel_filename);
            
            // Match base names: both extract to same base name
            if pdf_base_name == excel_base_name {
                return Some(page);
            }
        }
    }
    
    None
}

fn scan_child_directories(base_dir: &Path) -> Result<Vec<PathBuf>, Box<dyn std::error::Error>> {
    let mut pdf_files = Vec::new();
    
    // Scan only direct child directories (one level deep)
    for entry in fs::read_dir(base_dir)? {
        let entry = entry?;
        let path = entry.path();
        
        if path.is_dir() {
            // Scan PDF files in this child directory
            for file_entry in fs::read_dir(&path)? {
                let file_entry = file_entry?;
                let file_path = file_entry.path();
                
                if file_path.is_file() {
                    if let Some(ext) = file_path.extension() {
                        if ext.eq_ignore_ascii_case("pdf") {
                            pdf_files.push(file_path);
                        }
                    }
                }
            }
        }
    }
    
    Ok(pdf_files)
}

fn process_pdf_with_qpdf(
    pdf_path: &Path,
    bia_path: &Path,
    mappings: &HashMap<String, u32>,
    bia_page_count: usize,
) -> Result<bool, Box<dyn std::error::Error>> {
    let filename = pdf_path
        .file_name()
        .and_then(|n| n.to_str())
        .ok_or("Invalid filename")?;
    
    // Match PDF with Excel entries
    let page_index = match match_pdf_name(filename, mappings) {
        Some(idx) => idx,
        None => return Ok(false), // No match, skip
    };
    
    // Convert to 1-based page number
    let page_number = page_index + 1;
    
    // Validate page number
    if page_number as usize > bia_page_count {
        return Err(format!(
            "Page number {} exceeds bia.pdf page count ({})",
            page_number,
            bia_page_count
        )
        .into());
    }
    
    println!("  Inserting page {} from bia.pdf", page_number);
    
    // Create temp file for output
    let temp_dir = env::temp_dir();
    let temp_output_pdf = temp_dir.join(format!("merged_output_{}.pdf", std::process::id()));
    
    // Use qpdf to combine: page from bia.pdf first, then all pages from target PDF
    // qpdf --empty --pages bia.pdf N target.pdf -- output.pdf
    // Use --warning-exit-0 to return success even with warnings (common in non-standard PDFs)
    let output = Command::new("qpdf")
        .args([
            "--warning-exit-0",
            "--empty",
            "--pages",
            bia_path.to_str().unwrap(),
            &page_number.to_string(),
            pdf_path.to_str().unwrap(),
            "--",
            temp_output_pdf.to_str().unwrap(),
        ])
        .output()?;
    
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(format!("Failed to merge PDFs with qpdf: {}", stderr).into());
    }
    
    // Verify output exists
    if !temp_output_pdf.exists() {
        return Err("Failed to create merged PDF".into());
    }
    
    // Replace original file with merged output
    fs::copy(&temp_output_pdf, pdf_path)?;
    
    // Clean up temp file
    let _ = fs::remove_file(&temp_output_pdf);
    
    Ok(true)
}
