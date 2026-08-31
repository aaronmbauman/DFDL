# DFDL specification build.
#
# Metanorma runs inside a pinned Docker image, so nothing needs to be
# installed locally and every build is reproducible.
#
#   make xml     semantic XML only (fast)
#   make html    HTML
#   make pdf     ISO-formatted PDF (slow)
#   make all     pdf + html
#   make lint    ruff + black over tools/ (host tools, not the container)
#   make check   lint, then the validators in tools/ against the semantic XML
#   make clean   remove build output
#
# EDITION selects which edition to build; output lands in build/$(EDITION)/.
# Both editions render identically-named artifacts, so they need separate
# directories.
#
#   make EDITION=ogf pdf
#
# ISO's house font (Cambria) is proprietary; --continue-without-fonts lets
# the build fall back to Noto Sans instead of prompting for a licence.
#
# The build also writes a relaton/ bibliography cache at the repo root; it is
# gitignored and kept between builds so references resolve offline.

IMAGE := metanorma/metanorma:alpine-1.17.0
STEM  := dfdl

EDITION  ?= iso
SPEC_iso := spec/dfdl.adoc
SPEC_ogf := spec/dfdl-ogf.adoc
SPEC     := $(SPEC_$(EDITION))
BUILD    := build/$(EDITION)

# Metanorma names its output after the input file; the artifacts are
# renamed to $(STEM) so both editions land under the same names.
SRCSTEM  := $(basename $(notdir $(SPEC)))

DOCKER   := docker run --rm -v "$(CURDIR):/metanorma" $(IMAGE)
MN_FLAGS := --no-install-fonts --continue-without-fonts

.PHONY: all xml html pdf lint check clean

# One pdf pass also emits the HTML and the XML, so `all` is just `pdf`.
all: pdf

xml:  FORMATS := xml
html: FORMATS := xml,html
pdf:  FORMATS := xml,html,pdf

# Metanorma writes its output next to the input file, so move the artifacts
# into $(BUILD)/ afterwards.
xml html pdf:
	@mkdir -p $(BUILD)
	$(DOCKER) metanorma compile -t iso -x $(FORMATS) $(MN_FLAGS) $(SPEC)
	@for f in $(dir $(SPEC))$(SRCSTEM).*; do \
		case "$$f" in *.adoc) continue ;; esac; \
		mv -f "$$f" "$(BUILD)/$(STEM)$${f#$(dir $(SPEC))$(SRCSTEM)}"; \
	done
	@echo "Output in $(BUILD)/"

# ruff and black run on the host, not in the Metanorma container.
lint:
	@if ls tools/*.py >/dev/null 2>&1; then \
		ruff check tools/ && black --check tools/; \
	else \
		echo "No Python in tools/ yet; nothing to lint."; \
	fi

# Runs every executable in tools/, plus any tools/check*.py, over the
# build. Tolerant of tools/ being empty or absent.
check: lint html
	@checkers=$$({ find tools -maxdepth 1 -type f -perm -u+x; \
		ls tools/check*.py; } 2>/dev/null | sort -u); \
	if [ -z "$$checkers" ]; then \
		echo "No validators in tools/ yet; nothing to check."; \
	else \
		for c in $$checkers; do \
			echo "==> $$c"; \
			if [ -x "$$c" ]; then "$$c" $(BUILD)/$(STEM).xml; \
			else python3 "$$c" $(BUILD)/$(STEM).xml; fi || exit 1; \
		done; \
	fi

clean:
	rm -rf build .ruff_cache
