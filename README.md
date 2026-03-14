# CFDI XML to Excel Parser

Python script developed to extract data from CFDI XML files and
generate an Excel report for invoice tracking.

## Problem

The client only had CFDI XML files and needed to keep track of
their invoices in Excel.

Manually extracting the information was time-consuming.

## Solution

This script parses CFDI XML files and extracts key data such as:

- UUID
- Issuer RFC
- Receiver RFC
- Invoice date
- Subtotal
- VAT
- Total

The data is automatically exported into an Excel spreadsheet.

## Technologies

- Python
- XML parsing
- Pandas / OpenPyXL

## Extra Feature

A visual completion notification (fireworks animation) is displayed
when the process finishes successfully.

## Use case

Useful for accountants or businesses that need to consolidate
CFDI data into Excel reports.
