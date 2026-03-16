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

## How to Run Locally

Follow these steps to run the Streamlit application from scratch on your local machine:

1. **Clone the repository**
   ```bash
   git clone https://github.com/azahrulsmavo-design/available_car.git
   cd available_car
   ```

2. **Create a Virtual Environment (Recommended)**
   It's best practice to use a virtual environment so dependencies don't conflicts.
   ```bash
   # On Windows
   python -m venv venv
   venv\Scripts\activate

   # On macOS/Linux
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install Dependencies**
   Install the required packages from the `requirements.txt` file.
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the Application**
   Start the Streamlit development server.
   ```bash
   python -m streamlit run app.py
   ```
   
   The application should automatically open in your default web browser at `http://localhost:8501`.
