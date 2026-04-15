# HLI Project GHG Calcs Builder

This package lets you build a Windows `.exe` from GitHub Actions, so you do not need Python installed locally.

## What it does
- Preserves the original Excel template formatting by editing the workbook directly
- Populates columns A:N
- Leaves O:S intact unless you enable the formula compatibility repair
- Leaves the `Emission Factors` tab intact
- Converts `Legal` weight to `44,000`
- Blanks imported `0`, `NAME?`, and `#NAME?`
- Sorts by `Delivery Date` oldest first
- Estimates missing distances online
- Uses `rail = 90% of road distance`

## Data security
The app processes workbook contents locally on the computer where the `.exe` is run. It does not store shipment data in any external database. The only data that may leave the device is `Origin`, `Destination`, and `Mode` when a missing distance needs to be estimated using online map/geocoding services.

## Build a Windows EXE using GitHub
1. Create a new GitHub repository.
2. Upload all files from this folder, preserving the `.github/workflows` folder.
3. Open the repository on GitHub.
4. Go to **Actions**.
5. Run **Build Windows EXE**.
6. When the workflow finishes, open the run and download the artifact named **HLI_Project_GHG_Calcs_Builder_Windows**.
7. Extract it and run `HLI_Project_GHG_Calcs_Builder.exe`.

## Optional local build on a Windows machine
If you later have Python available on Windows, double-click `build_windows.bat`.


## Data Security

This application is designed to protect sensitive client data by performing all spreadsheet processing locally on the user’s device. Input files are not uploaded, stored, or transmitted to any external servers or databases.

To support distance estimation where required, the application makes limited external requests to third-party routing and geocoding services. These requests include only non-sensitive geographic information (i.e., origin and destination locations) necessary to calculate transportation distances. No client-specific data, emissions data, financial information, or full spreadsheet contents are transmitted externally.

The application does not retain, log, or store any data beyond the local session. All outputs are generated and saved directly on the user’s device.
