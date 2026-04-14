New-Item -ItemType Directory -Force -Path test_data
New-Item -ItemType Directory -Force -Path scripts

$test_files = @(
    "ACC_Ph3_SDD_8-04-26.docx",
    "ACC_Ph3_WCCA-furtherEdits_31032026.xlsx",
    "Doc_Review_UltraSmall_HDD (1) (1).xlsx",
    "Doc_Review_UltraSmall_HDD (1).xlsx",
    "Doc_Review_UltraSmall_HDD.xlsx",
    "Review_Report_ACC_Ph3_SDD_8-04-26_Pro_20260413_1354.xlsx",
    "TICO-ULTRASMALL-PH3-CONCEPT_HSIS_27_02_26 (1).xlsx",
    "Ultrasmall-CAE-CFD-Endo-san-comments.xlsx",
    "Ultrasmall_Ph4_HSIS_Review.xlsx",
    "Ultrasmall_Ph4_HardwareDesignDocument_08_04_2026.docx",
    "Ultrasmall_Ph4_WCCA_SCTM_Review.xlsx",
    "extracted_comments.txt",
    "new_doc_summary.txt",
    "rd.txt",
    "report_analysis.json",
    "wcca_output.txt"
)

foreach ($file in $test_files) {
    if (Test-Path ".\$file") {
        Move-Item -Path ".\$file" -Destination "test_data\" -Force
    }
}

$script_files = @(
    "analyze_report.py",
    "analyze_reviews.py",
    "check_wcca.py",
    "debug_excel.py",
    "extract_comments.py",
    "generate_test_docs.py"
)

foreach ($file in $script_files) {
    if (Test-Path ".\$file") {
        Move-Item -Path ".\$file" -Destination "scripts\" -Force
    }
}

git add .
git commit -m "chore: reorganize workspace into scripts and test_data directories"
git push
