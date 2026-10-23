POWERSHELL ?= pwsh
PYTHON ?= python
TRUTH_DIR ?= truth/pdf-word
TRUTH_FILTER ?= *
TRUTH_IMAGE_DIR ?= output/pdf/word-truth-images
DOTNET ?= dotnet
CLI_PROJECT ?= DocxportNet.Cli/DocxportNet.Cli.csproj
CLI_DLL ?= DocxportNet.Cli/bin/Release/net8.0/docxport.dll
FLOW_DIR ?= output/pdf/current-flow
SCORE_DIR ?= output/pdf/word-image-scores
BASELINE ?= truth/flow-word-visual-baseline.csv
SCORE_TOLERANCE ?= 0.001

.PHONY: help truth truth-missing check-truth truth-image render-flow check-flow score-current score-diff score baseline-update

help:
	@echo make truth          Export all DOCX fixtures through Word
	@echo make truth-missing  Export only DOCX fixtures without a Word PDF
	@echo make truth-image    Refresh local scoring images from Word PDFs
	@echo make render-flow    Export the selected DOCX fixtures with Flow
	@echo make score-current Compare existing Flow PDFs with cached Word images
	@echo make score-diff    Compare the current score with an optional baseline
	@echo make score         Run all checks, render, and score
	@echo Set TRUTH_FILTER=pattern to select fixture names and TRUTH_DIR=path to change the output directory.

truth:
	$(POWERSHELL) -NoProfile -File tools/doc-docx/Run-WordPdfTruth.ps1 -TruthDirectory "$(TRUTH_DIR)" -NameFilter "$(TRUTH_FILTER)"

truth-missing:
	$(POWERSHELL) -NoProfile -File tools/doc-docx/Run-WordPdfTruth.ps1 -TruthDirectory "$(TRUTH_DIR)" -NameFilter "$(TRUTH_FILTER)" -MissingOnly

check-truth:
	$(PYTHON) tools/doc-docx/Check-PdfInputs.py truth --fixture-dir fixtures/docx --pdf-dir "$(TRUTH_DIR)" --name-filter "$(TRUTH_FILTER)"

truth-image:
	$(PYTHON) tools/doc-docx/Render-WordTruthImages.py --pdf-dir "$(TRUTH_DIR)" --image-dir "$(TRUTH_IMAGE_DIR)" --name-filter "$(TRUTH_FILTER)"

render-flow:
	$(DOTNET) build "$(CLI_PROJECT)" -c Release
	$(PYTHON) tools/doc-docx/Run-FlowCliCorpus.py --fixture-dir fixtures/docx --output-dir "$(FLOW_DIR)" --name-filter "$(TRUTH_FILTER)" --dotnet "$(DOTNET)" --cli-dll "$(CLI_DLL)"

check-flow:
	$(PYTHON) tools/doc-docx/Check-PdfInputs.py flow --fixture-dir fixtures/docx --pdf-dir "$(FLOW_DIR)" --name-filter "$(TRUTH_FILTER)"

score-current:
	$(PYTHON) tools/doc-docx/Score-WordTruthImages.py --flow-dir "$(FLOW_DIR)" --truth-pdf-dir "$(TRUTH_DIR)" --truth-image-dir "$(TRUTH_IMAGE_DIR)" --output-dir "$(SCORE_DIR)" --name-filter "$(TRUTH_FILTER)"

score-diff:
	$(PYTHON) tools/doc-docx/Compare-WordVisualBaseline.py --current "$(SCORE_DIR)/flow-word-visual-results.csv" --baseline "$(BASELINE)" --output "$(SCORE_DIR)/flow-word-baseline-diff.csv" --tolerance "$(SCORE_TOLERANCE)"

score:
	$(MAKE) check-truth TRUTH_FILTER="$(TRUTH_FILTER)"
	$(MAKE) truth-image TRUTH_FILTER="$(TRUTH_FILTER)"
	$(MAKE) render-flow TRUTH_FILTER="$(TRUTH_FILTER)"
	$(MAKE) check-flow TRUTH_FILTER="$(TRUTH_FILTER)"
	$(MAKE) score-current TRUTH_FILTER="$(TRUTH_FILTER)"

baseline-update:
	$(PYTHON) tools/doc-docx/Copy-ScoreBaseline.py --current "$(SCORE_DIR)/flow-word-visual-results.csv" --baseline "$(BASELINE)"
