#!/bin/bash
# Initializes git repository with a clean, multi-step commit history for Zeckathon evaluation
git init
git branch -M main

git add README.md requirements.txt .gitignore
git commit -m "chore: initialize repository structure and project documentation"

git add data/master_90_parameters.csv data/peer_financials_fy26_fy29.csv data/cams_3statement_inputs.csv
git commit -m "feat(data): add 90-parameter dataset, peer financials, and CAMS 3-statement inputs"

git add src/dcf_engine.py src/peer_valuation.py
git commit -m "feat(valuation): implement CAMS 3-statement DCF engine and peer relative valuation model"

git add src/scenario_stresstest.py
git commit -m "feat(scenarios): model mandatory -30% F&O and -20% equity drawdown shock across candidates"

git add src/generate_charts.py src/build_excel_model.py run_pipeline.py
git commit -m "feat(pipeline): add dynamic openpyxl workbook generator, charts, and CLI pipeline"

git add outputs/ docs/ init_git_history.sh
git commit -m "docs(deliverables): generate final Excel financial model, charts, and IC memo"

echo "Git repository initialized with 6 structured commits! Run 'git log --oneline' to verify."
