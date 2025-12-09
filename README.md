# IACI — Internationalization Capability Index

This repository contains the research pipeline for computing the Internationalization Capability Index (IACI) for Chinese private universities. The modular implementation lives under `src/iaci_index`, while legacy step scripts are retained for reference under `legacy_scripts/`.

## Project structure
```
├── data/
│   ├── raw/
│   ├── interim/
│   └── processed/
├── src/iaci_index/        # Canonical pipeline modules
├── scripts/               # Thin entrypoints that call the package
├── legacy_scripts/        # Archived flat step scripts (no new development)
├── drivers/               # Selenium drivers
├── notebooks/             # Exploratory analyses (not the main pipeline)
├── environment.yml
└── README.md
```

The package exposes clear functions such as `run_step1`, `run_step2`, `run_step3`, `run_step4`, `build_lri_features`, `build_arii_features`, `build_tli_features`, `run_pca_for_intl_index`, `compute_iaci_4d`, and `prettify_scores`. Formulas, outputs, and filenames match the legacy scripts.

## Environment setup
Create the conda environment and activate it:

```bash
conda env create -f environment.yml
conda activate iaci
```

## Running the pipeline
Official full run:

```bash
python scripts/run_full_pipeline.py
```

Individual steps can be invoked as needed:

```bash
python scripts/run_step1_crawling.py
python scripts/run_step4_enrichment.py
python scripts/run_step5_features.py
python scripts/run_step5_pca.py
python scripts/run_step5_iaci_4d.py
```

## Notes
- Notebooks in the root directory are exploratory and are not part of the canonical pipeline.
- The `legacy_scripts/` directory keeps the old flat step files for archival/reference only; new work should use `scripts/` and `src/iaci_index/`.
- All outputs retain their original file names (e.g., `step5_A11_IACI_final_4D.xlsx`).
- API-dependent steps expect the same environment variables as before (e.g., `MOONSHOT_API_KEY`).
