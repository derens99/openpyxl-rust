# Development commands for openpyxl-rust
# Install just: https://github.com/casey/just

set shell := ["bash", "-euc"]

# List available commands
default:
    @just --list

# Build the Rust extension in release mode
build:
    uv run maturin develop --release

# Run the test suite
test *args:
    uv run pytest tests/ -q {{args}}

# Run Python lint and format checks
lint-py:
    uvx ruff check python/ tests/
    uvx ruff format --check python/ tests/

# Run Rust lint and format checks
lint-rs:
    cargo fmt --check
    cargo clippy -- -D warnings

# Run all lints (Python + Rust)
lint: lint-py lint-rs

# Auto-fix Python lint issues and format
fix:
    uvx ruff check --fix python/ tests/
    uvx ruff format python/ tests/

# Run cargo-deny license and advisory checks
deny:
    cargo deny check

# Run benchmarks against openpyxl
bench:
    uv run python benchmarks/bench_vs_openpyxl.py

# Install all dev dependencies
setup:
    uv sync --group dev
    uv run maturin develop --release

# Run the full CI pipeline locally (lint, deny, build, test)
ci: lint deny build test

# Clean build artifacts
clean:
    cargo clean
    rm -rf dist/ build/ *.egg-info/
