# Excel Validator

Excel Validator is a Python-based tool for validating Excel files using configurable validation rules defined in YAML files.

The project provides a simple **Streamlit web interface** for uploading Excel files, applying validation rules, reviewing errors, and downloading the validated Excel file.

## Features

- Upload Excel files through the browser
- Upload YAML validation configuration
- Select the worksheet to validate
- Validate Excel data using configurable rules
- Display validation progress
- Show total validation errors
- Display broken cells and validation messages
- Generate an annotated Excel file
- Download the validated Excel file
- Option to write validation messages into Excel cells

## Supported Validators

- Base Validator
- Choice Validator
- Conditional Validator
- Country Validator
- DateTime Validator
- Email Validator
- Excel Date Validator
- Length Validator
- NotBlank Validator
- Regex Validator
- Type Validator
- URL Validator
- Numeric Range Validator
- Date Range Validator

Example YAML rules:

```yaml
validators:
	columns:
		A:
			- URL:
					trim: true
		B:
			- NumericRange:
					min: 0
					max: 100
		C:
			- DateRange:
					min: "2020-01-01"
					max: "2030-12-31"
```

## Requirements

- Python 3.9 or higher
- Required Python libraries listed in `requirements.txt`
- Pandas and NumPy are used for error analysis and export

## Installation

Clone the repository and install the dependencies:

```bash
pip install -r requirements.txt

```
<img width="1902" height="871" alt="image" src="https://github.com/user-attachments/assets/122730b1-2dae-428e-9265-4ce5734ff8a8" />


