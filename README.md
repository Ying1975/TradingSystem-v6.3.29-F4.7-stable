# TradingSystem v6.3.29-F4.7 (Stable)

This repository contains the **stable core scripts** of TradingSystem v6.3.29-F4.7.

## Included files
- strategy_score.py  
  - Unified lights logic (turnover / trade value / short squeeze)
  - Market-grouped trade value ranking
- daily_auto_run_final.py  
  - Main pipeline
- export_top20.py  
  - Top20 exporter (no duplicated logic)

## Guarantees
- FULL / Decision / Top20 lights are **fully consistent**
- No dtype pollution (no LOW / emoji written into numeric columns)
- Trade value lights are computed **per market (TWSE / TWO)**

## How to use
Replace these files in your local TradingSystem directory and run:
```bash
python oneclick_daily_run.py --mode pre
