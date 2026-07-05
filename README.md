# Decode_project_data-analytics
# Excel Data Cleaning Project

## Project Overview

This project demonstrates the process of cleaning and preparing a raw sales dataset using Microsoft Excel. The objective was to improve data quality by identifying and correcting common issues such as missing values, duplicate records, inconsistent formatting, and validation errors.

## Dataset

The dataset contains customer order information, including:

* Order ID
* Order Date
* Customer ID
* Product
* Quantity
* Unit Price
* Shipping Address
* Payment Method
* Order Status
* Tracking Number
* Items in Cart
* Coupon Code
* Referral Source
* Total Price

## Objectives

* Handle missing values
* Check and remove duplicate records
* Standardize date formats
* Clean text fields using Excel functions
* Validate numeric values
* Verify calculated fields
* Prepare a clean dataset for analysis

## Data Cleaning Steps

### Missing Values

* Identified blank values in the CouponCode column.
* Replaced missing values with "No Coupon".

### Duplicate Records

* Checked duplicate OrderID values using Excel's Remove Duplicates tool.
* Confirmed there were no duplicate records.

### Text Cleaning

* Removed unnecessary spaces using the TRIM() function.
* Standardized text formatting.

### Date Formatting

* Converted dates into a consistent format.

### Validation

* Verified that TotalPrice = Quantity × UnitPrice.
* Checked for invalid numeric values.
* Confirmed no missing values remained.

## Excel Features Used

* Remove Duplicates
* TRIM()
* IF()
* COUNTIF()
* Find & Replace
* Conditional Formatting
* Sort & Filter
* Excel Tables
* Freeze Panes
* AutoFit Columns

## Final Result

The dataset was successfully cleaned and formatted for further analysis. It contains consistent values, standardized formatting, validated calculations, and improved overall data quality.

## Skills Demonstrated

* Data Cleaning
* Data Validation
* Data Preparation
* Microsoft Excel
* Spreadsheet Formatting
* Analytical Thinking

## Project Files

* raw_dataset.xlsx
* cleaned_dataset.xlsx

## Author

Your Name
Aspiring Data Analyst
