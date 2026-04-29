# Run Step3 with file-based output (avoids PowerShell subprocess hanging with asyncio)
$env:DEEPSEEK_API_KEY = "sk-5d7585b8e85b4d28a25e8f60e696ac78"
$env:DEEPSEEK_BASE_URL = "https://api.deepseek.com"
$env:DEEPSEEK_FLASH_MODEL = "deepseek-v4-flash"
$env:DEEPSEEK_PRO_MODEL = "deepseek-v4-pro"

$script = "d:\桌面\clientab-main\skills\strategic-emerging-monthly-report\scripts\deepseek_clean_summarize_excel.py"
$input = "d:\桌面\clientab-main\temp-data\summaries_deepseek_final_cleaned.xlsx"
$output = "d:\桌面\clientab-main\temp-data\summaries_deepseek_final_summarized.xlsx"
$log = "d:\桌面\clientab-main\temp-data\step3_output.log"

$proc = Start-Process -FilePath python -ArgumentList "-u","""$script""","-i","""$input""","-o","""$output""","--concurrency-flash","20","--concurrency-pro","5","--max-input-chars","12000","--max-cleaned-chars","10000","--limit","3" -NoNewWindow -PassThru -RedirectStandardOutput "$env:TEMP\step3_out.txt" -RedirectStandardError "$env:TEMP\step3_err.txt"
Write-Host "PID: $($proc.Id)"
$proc.WaitForExit()
Write-Host "Exit code: $($proc.ExitCode)"
Write-Host "--- stdout ---"
Get-Content "$env:TEMP\step3_out.txt" -ErrorAction SilentlyContinue
Write-Host "--- stderr ---"
Get-Content "$env:TEMP\step3_err.txt" -ErrorAction SilentlyContinue
