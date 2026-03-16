# KPI Availability ETL

This repository contains the ETL process and application for processing KPI Availability and STS Raw data. 

## Features
- Generates KPI Reports based on master assets and raw STS data.
- Capable of outputting synced formatted Excel files for final reporting.
- Utilizes `app.py` as the main entry point to run processing logic or web application UI.

## Structure
- `app.py`: Main application script.
- `etl_kpi.py`: Script for KPI data extraction and transformation.
- `etl_raw_sts.py`: Script for Raw STS data Extraction and transformation.
- `custom_order.py`: Custom ordering configurations and mapping details.
- `data/`: Contains raw input forms and data sources.
- `output/`: Processed Excel reports are dumped here.
