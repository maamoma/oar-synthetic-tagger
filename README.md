# Synthetic Tagger

Automated tagging tool for New Relic Synthetic monitors using NerdGraph API with CMDB enrichment.

## Overview

This tool reads entity data from CSV files, optionally enriches it via CMDB API lookups, and applies tags to New Relic Synthetic entities using the NerdGraph mutation API.

## Versions

| Script | Purpose |
|--------|---------|
| `syntheticTagger.py` | Original v1 tagger |
| `syntheticTagger_v2.py` | Full 13-tag tagger with CMDB lookup, pipe-split fallback |
| `syntheticTagger_v3.py` | Direct tag push from CSV (no CMDB lookup) |
| `syntheticTagger_v4_GB.py` | GB-specific 7-tag tagger with CMDB alias support |

## Setup

1. Create a virtual environment:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```

2. Install dependencies:
   ```bash
   pip install openpyxl requests pyodbc
   ```

3. Copy `config/config.json.template` to `config/config.json` and fill in credentials.

4. Place your input CSV in the `synthetic_tag/` directory.

## Configuration

See `config/config.json.template` for the expected format. Required values:
- **New Relic API Key** (`NRAK-*`)
- **CMDB API subscription key** (for APIM endpoint)
- **SQL Server credentials** (for reporting, optional)

## Usage

```bash
cd synthetic_tagger/synthetic_tag
python syntheticTagger_v3.py
```

Each script will prompt or read from its configured CSV, compare existing tags, and push updates to New Relic.
