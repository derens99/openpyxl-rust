# Contributing to openpyxl_rust

Thanks for your interest in contributing!

## Development Setup

### Prerequisites
- Python 3.9+
- Rust toolchain (install via [rustup](https://rustup.rs/))
- [maturin](https://github.com/PyO3/maturin) (`pip install maturin`)

### Getting Started

```bash
# Clone the repo
git clone https://github.com/derens99/openpyxl-rust.git
cd openpyxl-rust

# Create a virtual environment
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows

# Install dev dependencies
pip install maturin pytest openpyxl ruff pre-commit

# Build the Rust extension
maturin develop --release

# Run tests
pytest tests/ -q

# Set up pre-commit hooks
pre-commit install
```

### Project Structure

```
src/              # Rust code (PyO3 bindings + rust_xlsxwriter)
python/           # Python package (openpyxl-compatible API)
tests/            # pytest test suite
benchmarks/       # Performance benchmarks
```

### Running Checks

```bash
# Python linting + formatting
ruff check python/ tests/ benchmarks/
ruff format python/ tests/ benchmarks/

# Rust linting + formatting
cargo clippy
cargo fmt

# Tests
pytest tests/ -q
```

## Submitting Changes

1. Fork the repo and create a branch from `main`
2. Make your changes
3. Ensure `ruff check`, `ruff format --check`, `cargo clippy`, `cargo fmt --check`, and `pytest` all pass
4. Submit a pull request

## Reporting Bugs

Use the [bug report template](https://github.com/derens99/openpyxl-rust/issues/new?template=bug_report.md) to file issues.

## Code Style

- **Python**: Formatted with ruff (config in `pyproject.toml`)
- **Rust**: Formatted with rustfmt, linted with clippy
