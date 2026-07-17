# archgen — common tasks. Usage: `make <target>`.
# Override vars: `make render SPEC=examples/sample_spec.yaml OUT=out.pptx`

PY ?= .venv/bin/python
ARCHGEN ?= .venv/bin/archgen
SPEC ?= examples/sample_spec.yaml
OUT ?= out.pptx

.DEFAULT_GOAL := help
.PHONY: help setup icons test lint render preview gallery doctor schema clean

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | \
		awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-10s\033[0m %s\n", $$1, $$2}'

setup: ## Create venv, install package + dev deps, populate icons
	python -m venv .venv
	$(PY) -m pip install -e ".[dev]"
	$(PY) -m archgen.registry.fetch

icons: ## Rebuild the icon manifest from local assets/icons
	$(PY) -m archgen.registry.fetch --no-download

test: ## Run the test suite
	$(PY) -m pytest -q

lint: ## Validate the spec ($(SPEC))
	$(ARCHGEN) lint $(SPEC)

# RENDER_FLAGS defaults to the agent workflow (CLAUDE.md §4): best-fit direction
# and a layout report that fails on error findings. RENDER_FLAGS= to disable.
RENDER_FLAGS ?= --auto-direction --strict

render: ## Render $(SPEC) -> $(OUT) (with layout report)
	$(ARCHGEN) render $(SPEC) -o $(OUT) $(RENDER_FLAGS)

preview: ## Render + open a PNG preview of $(OUT) (macOS QuickLook)
	$(ARCHGEN) render $(SPEC) -o $(OUT) $(RENDER_FLAGS)
	@rm -rf .preview && mkdir -p .preview
	@qlmanage -t -s 1600 -o .preview $(OUT) >/dev/null 2>&1 && open .preview/*.png

gallery: ## Render every example (+ bespoke shine board) to docs/gallery/*.png
	./scripts/build_gallery.sh

doctor: ## Report layout/converter capabilities
	$(ARCHGEN) doctor

schema: ## Print the spec JSON Schema
	$(ARCHGEN) schema

clean: ## Remove build/render/cache artifacts
	rm -rf .archgen_cache .preview out *.pptx *.egg-info
	find . -name __pycache__ -type d -prune -exec rm -rf {} +
