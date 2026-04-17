# Report Generator
# Usage:
#   make          - fetch fresh data from GitHub + generate all reports (docx + pdf + html)
#   make html     - fetch fresh data + generate HTML reports only
#   make pdf      - fetch fresh data + generate DOCX + PDF reports only
#   make setup    - create .env from .env.example (run once after cloning)
#   make auth     - save your gh token to .env (run once after cloning)
#   make fetch    - fetch GitHub project data only (no report generation)
#   make clean    - delete generated reports and image cache
#   make pages    - copy the main HTML report to index.html

-include .env
export

QUERY       = query.graphql
RAW_JSON    = project_raw.json
PRETTY_JSON = project_pretty.json
SCRIPT      = generate_report.js
ENV_FILE    = .env

.PHONY: all html pdf setup auth fetch clean pages

all: fetch
	@echo "Generating all reports (DOCX + PDF + HTML)..."
	node $(SCRIPT)

html: fetch
	@echo "Generating HTML reports only..."
	node $(SCRIPT) --html

pdf: fetch
	@echo "Generating DOCX + PDF reports only..."
	node $(SCRIPT) --pdf

setup:
	@if [ -f $(ENV_FILE) ]; then \
		echo "$(ENV_FILE) already exists, skipping."; \
	else \
		cp .env.example $(ENV_FILE); \
		echo "Created $(ENV_FILE) from .env.example — fill in your values then run: make auth"; \
	fi

auth:
	@echo "Saving GitHub token to $(ENV_FILE)..."
	@sed -i 's|^GITHUB_TOKEN=.*|GITHUB_TOKEN=$(shell gh auth token)|' $(ENV_FILE)
	@echo "Token updated in $(ENV_FILE) (gitignored)"

fetch:
	@echo "Fetching GitHub project data..."
	gh api graphql -F query=@$(QUERY) -F org=$(PROJECT_ORG) -F projectNumber=$(PROJECT_NUMBER) > $(RAW_JSON)
	@echo "Raw JSON written to $(RAW_JSON)"
	python3 -m json.tool $(RAW_JSON) > $(PRETTY_JSON)
	@echo "Pretty-printed JSON written to $(PRETTY_JSON)"

pages:
	@echo "Copying main HTML report to index.html..."
	cp html/*_Test_Report.html index.html

clean:
	@echo "Cleaning up..."
	-rm -f $(RAW_JSON) $(PRETTY_JSON)
	-rm -rf docx/ pdf/ .img_cache/ html/
	@echo "Cleaned up"

# for GlobTek use only
upload:
	printf "cd $(WEBSITE_PATH)\nput $(PWD)/$(REPORT_FILE_PATH) index.html\nbye\n" | sshpass -p "$(SFTP_PASS)" sftp -o StrictHostKeyChecking=no $(SFTP_USER)@$(SFTP_HOST)