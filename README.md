# CRS School Data Scraper

This project scrapes data from the Chinese Ministry of Education's CRS platform:
https://www.crs.jsj.edu.cn

## Features

- Extracts school cooperation programs by country
- Cleans and structures data (region, category, university, program)
- Exports formatted Excel reports with:
  - Merged cells
  - Sorted by pinyin
  - Clickable links
- Extracts detail page data:
  - Level (Bachelor/Master)
  - Duration
  - Major / Course

## Output

- `crs_school_list.xlsx`
- `crs_detail_data.xlsx`

## Installation

```bash
pip install -r requirements.txt