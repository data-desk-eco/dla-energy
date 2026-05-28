.PHONY: build preview etl data clean

build:
	@echo "{\"date\": \"$$(git log -1 --format=%cI)\"}" > data/last_updated.json
	yarn build

preview:
	yarn preview

etl: data

data:
	@mkdir -p data
	uv run --with duckdb --with openpyxl scripts/build_data.py

clean:
	rm -rf docs/.observable data/data.duckdb data/raw
