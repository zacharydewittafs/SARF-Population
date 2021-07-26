## Description

This script reads in completed P&P Excel files and generates a populated SARF PDF file containing all user information from the file 
corresponding to each of the Excel files

## Requirements
- Windows OS
- Python (or the Anaconda distrubution of Python)
- PyPDF2 and pandas python libraries

## Initial Setup 
- Install Python via Windows Installer (https://www.python.org/downloads/release/python-376/) or the Anaconda Python 3.X distribution on your computer (https://www.anaconda.com/distribution/)
    - If installing via the Windows Installer make sure to select 'Add Python 3.X to PATH' before starting the installation 
    - If installing Anaconda distrubution, follow ONLY step #1 in ETL Pipeline Onboarding doc https://docs.google.com/document/d/1xrOD9rUI2Y_bRssm9nvC2YQbxR1jmXqp/edit
- Install Git on your computer: https://git-scm.com/downloads
- Clone this Repository onto your computer
    - git clone https://github.com/DonnaJacksonAFS/sarf_generation.git
- In the directory you've cloned this repository into, run the following command. This installs all of the other Python package dependencies on your computer.
    - pip install -r requirements.txt

## How to Use
1. Create a P&P_Files and SARF_Template folder in the main project
2. Place all completed P&P Excel files (.xls and .xlsx only) in the P&P_Files directory
3. Place the SARF template PDF file in the SARF_Template directory (only one file should exist in this directory)
4. Run (double-click) the **generate_sarfs.bat** file  - Note: if you have the Anaconda distribution of python, then run **generate_sarfs_anaconda.bat**
