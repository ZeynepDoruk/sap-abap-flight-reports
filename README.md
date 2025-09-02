Z_SFLIGHT_REPORT
Overview

Z_SFLIGHT_REPORT is an ABAP report designed to display flight reservation data in a structured and interactive ALV (ABAP List Viewer) grid. The report allows users to filter, sort, and export flight reservation details to Excel, providing a comprehensive view of bookings, customer details, flight information, and occupancy rates.



Purpose

The report aims to provide SAP users with a tool to:
View flight reservations with customizable filters.
Analyze occupancy rates of flights.
Sort and organize data dynamically.
Export filtered results directly to Excel for reporting purposes.

Features

Selection Screen Filters
Users can filter data by:
Carrier Name (p_carna) – airline name
Customer Name (p_name)
Departure City (p_citfr)
Arrival City (p_cityto)
Flight Date Range (p_fldafr to p_fldato)
Booking Class (p_class) – economy, business, first, etc.
Data Handling
Fetches data from custom structure ZRESERVATION_DATA and related standard SAP tables:
SCARR – Airline details
SPFLI – Flight route details
SBOOK – Booking details
SFLIGHT – Flight schedule and capacity
SCUSTOM – Customer information
Calculates occupancy rate for each flight.
Removes duplicate records for cleaner display.
Default sorting by Airline Name, Reservation Number, and Occupancy Rate.

ALV Display

Dynamic column headers with descriptive names.
Column width optimization and zebra striping.
Editable rows for inline updates.
Top-of-page commentary including report title, current date, and user information.

Excel Export

Export ALV results to Excel (.xlsx) via ABAP2XLSX integration.
Both interactive (user-triggered) and batch job exports supported.
First row is frozen for better readability.

Job Mode

When run in batch mode (sy-batch = 'X'), the report automatically generates Excel output without user interaction.
Usage
Execute transaction Z_SFLIGHT_REPORT.
Enter filter criteria on the selection screen.
Click Execute to view the results in ALV grid.
Optional: Click Excel button in toolbar to export the data.
Batch mode execution automatically exports the report to Excel at a predefined path.

Technical Details

ALV Grid Functions:
REUSE_ALV_GRID_DISPLAY_LVC for dynamic display.
Custom field catalog (gs_field_catalog) for mapping structure fields to user-friendly column names.
Layout settings (gs_layout) for optimized display and editing.
Sorting: Multi-level sorting implemented using lvc_t_sort structure.

Excel Export:
Uses zcl_excel and zcl_excel_worksheet classes.
Conversion from internal table to Excel handled by zcl_excel_converter.
Saved using cl_gui_frontend_services=>gui_download.
Error Handling: Try-catch blocks implemented during Excel export for exception safety.

Notes

Occupancy rate (OCC_RATE) of 1 indicates a fully booked flight.
Default filters are applied if user does not enter values (e.g., % wildcard for names and cities, current date for start of flight date, 9999-12-31 for end of flight date).
Maximum 200 rows are fetched per execution to avoid performance issues.

Dependencies

Custom data structure ZRESERVATION_DATA.
Custom ABAP classes: ZCL_EXCEL, ZCL_EXCEL_WORKSHEET, ZCL_EXCEL_WRITER_2007.
Standard SAP tables: SCARR, SPFLI, SBOOK, SFLIGHT, SCUSTOM.

License

This project is intended for educational and internal SAP development purposes.

Author
Name: Zeynep
Date: 10-12-2024
