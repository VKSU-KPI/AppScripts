# General Election Automation Scripts

This repository contains the core Google Apps Script files used to manage the interactive dashboard for election tracking. These scripts work together to ensure data integrity, prevent user errors, and automate the creation of standardized election documents based on predefined templates. The codebase is organized into modular files, utilizing a central configuration architecture to make maintenance and scaling straightforward.

### `config.gs`

The `config.gs` file serves as the single source of truth for all global variables used across the project. It stores crucial cell references, sheet names, and arrays defining which election types strictly require faculty inputs. By isolating these constants from the main logic, we ensure that any future structural changes to the spreadsheet only require a single, quick update in this file. This prevents hardcoded values from scattering across the executable scripts and makes the entire system highly resilient.

### `election_faculty_menu.gs`

The `election_faculty_menu.gs` script is responsible for dynamically updating the user interface of the dashboard. It runs automatically in the background using the built-in `onEdit` trigger every time a user modifies a cell. The script monitors the primary control cell for selecting the election type, ensuring it only executes when modifications occur on the designated dashboard sheet. When a user selects an election type that requires a faculty division, the script dynamically generates a data validation dropdown menu populated with the available faculties. Conversely, if the chosen election does not require a faculty, it explicitly clears the data validation and content to prevent accidental assignment, keeping the reporting structure perfectly logical.

### `election_sheet.gs`

The `election_sheet.gs` script handles the automated generation of brand new spreadsheet tabs for fresh election instances. Triggered manually via a designated button on the dashboard, it reads the final configuration of the election type and faculty. After performing strict validation checks to prevent incomplete data entry, it locates the corresponding hidden template by appending a standardized suffix to the election name. It then creates a complete duplicate of the template, applies the strictly formatted naming convention, verifies that a sheet with this name does not already exist to prevent overwriting, and finally resets the dashboard inputs to a clean state.

### `append_election_cycle.gs`

The `append_election_cycle.gs` script introduces the ability to manage recurring election cycles within a single, continuous document. Instead of creating a brand new tab for every event, this manually triggered function appends a fresh, empty copy of the template directly below the existing historical data on the targeted faculty sheet. Its primary technical achievement is the sophisticated handling of complex conditional formatting. It actively intercepts the formatting rules immediately after pasting, scans for strict absolute row references, and mathematically shifts the row numbers down based on the exact offset calculated during the paste. This safely severs the formatting connection between old and new tables while preserving relative references natively, seamlessly maintaining an independent, multi-cycle history on a single page.