invoice_group_excel
====================

A Python library for automatic Excel cost aggregation, filtering, grouping and multi‑sheet report generation.

Features:
- Auto-detects numeric columns and sums them
- Works with arbitrary column names
- Generates multiple sheets based on filters
- Supports grouping and sorting
- Configurable formatting

Installation:
```
pip install invoice_group_excel
```

Usage:
```
from invoice_group_excel import InvoiceAggregator, CFG
ag = InvoiceAggregator(CFG.input_data_file, ['Sheet1'])
ag.save_and_style(CFG.output_data_file)
```
