"""Run step3 with live output streaming"""
import subprocess, sys, os, time

script = r"d:\桌面\clientab-main\skills\strategic-emerging-monthly-report\scripts\deepseek_clean_summarize_excel.py"
input_f = r"d:\桌面\clientab-main\temp-data\summaries_deepseek_final_cleaned.xlsx"
output_f = r"d:\桌面\clientab-main\temp-data\summaries_deepseek_final_summarized.xlsx"

env = os.environ.copy()
env['PYTHONUNBUFFERED'] = '1'

proc = subprocess.Popen(
    [sys.executable, "-u", script,
     "-i", input_f,
     "-o", output_f,
     "--concurrency-flash", "20",
     "--concurrency-pro", "5",
     "--max-input-chars", "12000",
     "--max-cleaned-chars", "10000"],
    stdout=subprocess.PIPE,
    stderr=subprocess.STDOUT,
    text=True,
    env=env,
    encoding="utf-8",
    errors="replace",
)

print(f"PID: {proc.pid}", flush=True)
for line in proc.stdout:
    print(line, end="", flush=True)

exit_code = proc.wait()
print(f"\nExit code: {exit_code}", flush=True)
sys.exit(exit_code)
