.PHONY: build preview etl data capacity clean

build:
	@echo "{\"date\": \"$$(git log -1 --format=%cI)\"}" > data/last_updated.json
	yarn build

preview:
	yarn preview

etl: data

# Refresh the cited refinery-capacity reference table from its sources.
# Committed to the repo, so `data` does not require it each run.
capacity:
	uv run --with requests scripts/build_capacity.py

data:
	@mkdir -p data
	uv run --with duckdb --with openpyxl scripts/build_data.py

clean:
	rm -rf docs/.observable data/data.duckdb data/raw
