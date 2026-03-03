import subprocess, sys
result = subprocess.run([sys.executable, 'docx_to_ppt.py'], capture_output=True, text=True, encoding='utf-8')
with open('run_log.txt', 'w', encoding='utf-8') as f:
    f.write(result.stdout)
    if result.stderr:
        f.write('\n--- STDERR ---\n')
        f.write(result.stderr)
print("Done. Check run_log.txt")
