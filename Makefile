POWERSHELL ?= pwsh
TRUTH_DIR ?= truth/pdf-word
TRUTH_FILTER ?= *

.PHONY: help truth truth-missing

help:
	@echo make truth          Export all DOCX fixtures through Word
	@echo make truth-missing  Export only DOCX fixtures without a Word PDF
	@echo Set TRUTH_FILTER=pattern to select fixture names and TRUTH_DIR=path to change the output directory.

truth:
	$(POWERSHELL) -NoProfile -File tools/doc-docx/Run-WordPdfTruth.ps1 -TruthDirectory "$(TRUTH_DIR)" -NameFilter "$(TRUTH_FILTER)"

truth-missing:
	$(POWERSHELL) -NoProfile -File tools/doc-docx/Run-WordPdfTruth.ps1 -TruthDirectory "$(TRUTH_DIR)" -NameFilter "$(TRUTH_FILTER)" -MissingOnly
