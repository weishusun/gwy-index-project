# IACI — Internationalization Capability Index

This repository organizes the research pipeline for building the Internationalization Capability Index (IACI) into a modular Python package.

## Project structure
```
├── data/
│   ├── raw/
│   ├── interim/
│   └── processed/
├── src/
│   └── iaci_index/
│       ├── config.py
│       ├── data_io.py
│       ├── crawling/
│       │   ├── step1_school_list.py
│       │   ├── step2_official_site_search.py
│       │   ├── step2_extra_info_urls.py
│       │   ├── step3_metrics_crawler.py
│       │   └── offline_cache.py
│       ├── enrichment/
│       │   ├── kimi_api.py
│       │   ├── step4_llm_metrics_completion.py
│       │   └── text_cleaning.py
│       ├── features/
│       │   ├── language.py
│       │   ├── asean.py
│       │   └── text_intl.py
│       ├── modeling/
│       │   ├── pca_model.py
│       │   └── iaci_composite.py
│       └── utils/
│           ├── common.py
│           └── logging_utils.py
├── scripts/
│   ├── run_step1_crawling.py
│   ├── run_step4_enrichment.py
│   ├── run_step5_features.py
│   ├── run_step5_pca.py
│   ├── run_step5_iaci_4d.py
│   └── run_full_pipeline.py
└── environment.yml
```

Each module mirrors the original step scripts while exposing clear entry points (e.g., `run_step1`, `run_step2`, `build_lri_features`, `run_pca_for_intl_index`, `compute_iaci_4d`, `prettify_scores`). The behavior, file names, and formulas remain unchanged.

## Running the pipeline
Install dependencies with Conda:

```bash
conda env create -f environment.yml
conda activate iaci
```

Run individual steps:

```bash
python scripts/run_step1_crawling.py
python scripts/run_step4_enrichment.py
python scripts/run_step5_features.py
python scripts/run_step5_pca.py
python scripts/run_step5_iaci_4d.py
```

Or execute the full workflow:

```bash
python scripts/run_full_pipeline.py
```

## Notes
- All outputs retain their original file names (e.g., `step5_A11_IACI_final_4D.xlsx`).
- API-dependent steps expect the same environment variables as before (e.g., `MOONSHOT_API_KEY`).
- The modular layout makes it easier to run notebooks or scripts with consistent imports while preserving existing calculations.
