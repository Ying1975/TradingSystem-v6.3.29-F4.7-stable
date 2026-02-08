# TradingSystem v6.3.29-F4.7 — Stable Core Scripts 

This repository contains the **stable core logic** for TradingSystem v6.3.29.

It is intended to be used as a **drop-in replacement** for the corresponding files
in an existing v6.3.29 TradingSystem environment.

---

## Included files (ONLY these)

This repo intentionally contains **only three files**:

- **strategy_score.py**
  - Single source of truth for:
    - Short squeeze pressure light
    - Turnover rate light
    - Trade value light
  - Market-grouped trade value ranking (TWSE / TWO separated)
  - Composite score calculation

- **daily_auto_run_final.py**
  - Main daily pipeline
  - FULL / Decision outputs
  - Delegates all scoring & lights to `strategy_score.py`

- **export_top20.py**
  - Top20 exporter
  - Reuses the same unified logic
  - No duplicated scoring or light rules

> ⚠️ No other files are expected or supported in this repo.

---

## Guarantees (v6.3.29 only)

- FULL / Decision / Top20 lights are **fully consistent**
- No dtype pollution:
  - No `LOW`
  - No emoji
  - No string written into numeric columns
- Trade value lights and rankings are computed **per market**
  - TWSE and TPEx (TWO) are never mixed

---

## What this repo is NOT

- ❌ Not a full TradingSystem repository
- ❌ Not a data downloader
- ❌ Not responsible for Yahoo / TWSE / TPEx connectivity
- ❌ Not guaranteed to work on versions other than v6.3.29

This repo assumes:
- Existing v6.3.29 folder structure
- Existing data/cache mechanism

---

## Version lock (IMPORTANT)

This code is **frozen for TradingSystem v6.3.29**.

- Do **NOT** mix with v6.4+ without review
- Do **NOT** partially copy logic into other versions
- If logic changes are needed:
  - Fork this repo
  - Bump version explicitly

---

## How to use

1. Replace the following files in your local `TradingSystem` directory:
   - `strategy_score.py`
   - `daily_auto_run_final.py`
   - `export_top20.py`

2. Run:
```bash
python oneclick_daily_run.py --mode pre
