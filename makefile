# GTM965500P Report Generator
# Usage:
#   make          - fetch fresh data from GitHub + generate all reports (docx + pdf + html)
#   make html     - fetch fresh data + generate HTML reports only
#   make pdf      - fetch fresh data + generate DOCX + PDF reports only
#   make auth 	  - save your gh token to .env (run once after cloning)
#   make fetch    - fetch GitHub project data only (no report generation)
#   make clean    - delete generated reports and image cache

QUERY       = query.graphql
RAW_JSON    = project_issues.json
PRETTY_JSON = GTM965500P_pretty.json
SCRIPT      = generate_report.js
ENV_FILE    = .env

.PHONY: all html pdf auth fetch clean

all: fetch
	@echo "Generating all reports (DOCX + PDF + HTML)..."
	node $(SCRIPT)

html: fetch
	@echo "Generating HTML reports only..."
	node $(SCRIPT) --html

pdf: fetch
	@echo "Generating DOCX + PDF reports only..."
	node $(SCRIPT) --pdf

auth: 
	@echo "Saving GitHub token to $(ENV_FILE)..."
	@echo "GITHUB_TOKEN=$$(gh auth token)" > $(ENV_FILE)
	@echo "Token saved to $(ENV_FILE) (gitignored)"


# Always re-fetch — don't rely on file timestamps
fetch:
	@test -f $(ENV_FILE) || { echo "!! No .env file found. Run 'make auth'"; exit 1; }
	@echo "Fetching GitHub project data..."
	gh api graphql -F query=@$(QUERY) > $(RAW_JSON)
	@echo "✓ Raw JSON written to $(RAW_JSON)"
	python3 -m json.tool $(RAW_JSON) > $(PRETTY_JSON)
	@echo "✓ Pretty-printed JSON written to $(PRETTY_JSON)"

clean:
	rm -f $(RAW_JSON) $(PRETTY_JSON)
	rm -rf docx/ pdf/ .img_cache/ html/
	@echo "✓ Cleaned up"
