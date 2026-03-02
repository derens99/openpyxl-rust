# Contributing to openpyxl_rust

Thanks for your interest in contributing!

## Development Setup

### Prerequisites
- Python 3.10+
- Rust toolchain (install via [rustup](https://rustup.rs/))
- [uv](https://docs.astral.sh/uv/) (recommended) or pip
- [just](https://github.com/casey/just) (optional, for dev commands)

### Getting Started

```bash
# Clone the repo
git clone https://github.com/derens99/openpyxl-rust.git
cd openpyxl-rust

# Install dev dependencies and build (using just)
just setup

# Or manually:
uv sync --group dev
uv run maturin develop --release

# Run tests
just test
# or: uv run pytest tests/ -q

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
# All lints (Python + Rust)
just lint

# Auto-fix Python lint issues
just fix

# License + vulnerability audit
just deny

# Tests
just test

# Benchmarks
just bench
```

Or without just:

```bash
# Python
uvx ruff check python/ tests/
uvx ruff format --check python/ tests/

# Rust
cargo fmt --check
cargo clippy -- -D warnings
cargo deny check

# Tests
uv run pytest tests/ -q
```

## Submitting Changes

1. Fork the repo and create a branch from `dev`
2. Make your changes
3. Run `just lint` and `just test` to verify
4. Submit a pull request

## Reporting Bugs

Use the [bug report template](https://github.com/derens99/openpyxl-rust/issues/new?template=bug_report.md) to file issues.

## Code Style

- **Python**: Formatted and linted with [ruff](https://docs.astral.sh/ruff/) (config in `pyproject.toml`)
- **Rust**: Formatted with rustfmt, linted with clippy (config in `Cargo.toml` `[lints]`)
