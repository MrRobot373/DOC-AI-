Remove-Item -Path "WIP_Test Rpt_US" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "__pycache__" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "uploads" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "reports" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "reorg.ps1" -Force -ErrorAction SilentlyContinue

New-Item -ItemType Directory -Force -Path "backend"

Move-Item -Path "app.py" -Destination "backend\" -Force
Move-Item -Path "doc_parser.py" -Destination "backend\" -Force
Move-Item -Path "doc_fixer.py" -Destination "backend\" -Force
Move-Item -Path "report_generator.py" -Destination "backend\" -Force
Move-Item -Path "review_engine.py" -Destination "backend\" -Force
Move-Item -Path "requirements.txt" -Destination "backend\" -Force
Move-Item -Path "templates" -Destination "backend\" -Force
Move-Item -Path "tests" -Destination "backend\" -Force

git add -A
