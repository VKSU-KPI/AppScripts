# General Election Automation Scripts

This folder contains the core Google Apps Script files used to manage the interactive dashboard for election tracking. These scripts work together to ensure data integrity, prevent user errors, and automate the creation of standardized election documents based on predefined templates. The codebase is organized into modular files, utilizing a central configuration architecture and shared utility functions to make maintenance and scaling straightforward.

### `config.gs`

The `config.gs` file serves as the single source of truth for all global variables used across the project. It stores crucial cell references, sheet names, history indexing coordinates, and arrays defining which election types strictly require faculty or dormitory inputs. By isolating these constants from the main logic, we ensure that any future structural changes to the spreadsheet only require a single, quick update in this file. This prevents hardcoded values from scattering across the executable scripts and makes the entire system highly resilient.

### `utils.gs`

The `utils.gs` file is a centralized helper module designed to keep the codebase clean and avoid code duplication. It contains shared functions used by multiple other scripts, such as reading the current input parameters from the dashboard and securely wiping those inputs clear after a successful operation. By routing these repetitive tasks through a single utility file, we ensure that any changes to how data is read or cleared only need to be updated in one place.

### `election_menus.gs`

The `election_menus.gs` script is responsible for dynamically updating the user interface of the dashboard. It runs automatically in the background using the built-in `onEdit` trigger every time a user modifies a cell. The script monitors the primary control cell for selecting the election type, ensuring it only executes when modifications occur on the designated dashboard sheet. When a user selects an election type that requires a faculty or dormitory division, the script dynamically generates a data validation dropdown menu populated with the available options. To prevent cross-contamination, selecting a faculty-based election automatically clears any lingering dormitory data, and vice versa.

### `election_sheet.gs`

The `election_sheet.gs` script handles the automated generation of brand new spreadsheet tabs for fresh election instances. Triggered manually via a designated button on the dashboard, it reads the final configuration of the election type, faculty, and dormitory parameters. After performing strict validation checks to prevent incomplete data entry, it locates the corresponding hidden template. It creates a complete duplicate of the template, applies the strictly formatted naming convention (using a hashtag prefix specifically for dormitories), and verifies that a sheet with this exact name does not already exist. Upon successful creation, it triggers the history indexer to log the new sheet and finally resets the dashboard inputs to a clean state.

### `election_cycle.gs`

The `election_cycle.gs` script introduces the ability to manage recurring election cycles within a single, continuous document. Instead of creating a brand new tab for every event, this manually triggered function appends a fresh, empty copy of the template directly below the existing historical data on the targeted tracking sheet. Its primary technical achievement is the sophisticated handling of complex conditional formatting. It actively intercepts the formatting rules immediately after pasting, scans for strict absolute row references, and mathematically shifts the row numbers down based on the exact offset calculated during the paste. This safely severs the formatting connection between old and new tables while seamlessly maintaining an independent, multi-cycle history on a single page.

### `history.gs`

The `history.gs` script acts as an automated indexer, keeping the main overview sheet perfectly organized. Whenever a completely new election tab is generated, this script constructs a direct, clickable hyperlink using the unique internal ID of the newly created sheet. For unique university-wide elections, it safely writes the link into a designated fixed cell. For sub-divisions like faculties and dormitories, it scans a predefined column range, finds the first available empty row, inserts the localized formula, and then forces the spreadsheet to physically update the visual interface before automatically sorting the entire column alphabetically. This guarantees a clean, dynamically updating table of contents for the entire election process.