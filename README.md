# HLI Project GHG Calcs Builder

This package builds a Windows `.exe` from GitHub Actions so you do not need Python installed locally.

## What this version does
- Preserves the original Excel template formatting by editing the workbook directly
- Populates columns A:N and leaves the `Emission Factors` tab intact
- Converts `Legal` weight to `44,000`
- Blanks imported `0`, `NAME?`, and `#NAME?`
- Sorts by `Delivery Date` oldest first
- Keeps distance blank only for `Storage`
- Uses stronger mode-specific distance logic:
  - **Road:** OSRM online road-network routing when available, with calibrated fallback floors
  - **Rail:** corridor / hub routing for covered North American lanes, then road-network proxy × rail factor, then calibrated fallback
  - **Sea:** maritime-network routing through `searoute` when available, else calibrated sea fallback
  - **Inland Water:** Atlantic ICW corridor logic when applicable, else calibrated inland-water fallback
- Reuses cached lane estimates for better speed
- Converts all externally sourced route distances to miles before writing the output workbook

## Data security
The app processes workbook contents locally on the computer where the `.exe` is run. It does not store shipment data in any external database.

The only data that may leave the device is limited location data needed for routing/geocoding when a missing distance needs to be estimated. Distances from those services are normalized to miles before the workbook is written. Specifically:
- `Origin`
- `Destination`
- mode-driven routing requests to public routing/geocoding services

No full workbook contents, client totals, financial data, or emissions results are sent externally.

## Build a Windows EXE using GitHub
1. Create a new GitHub repository.
2. Upload all files from this folder, preserving the `.github/workflows` folder.
3. Open the repository on GitHub.
4. Go to **Actions**.
5. Run **Build Windows EXE**.
6. When the workflow finishes, open the run and download the artifact named **HLI_Project_GHG_Calcs_Builder_Windows**.
7. Extract it and run `HLI_Project_GHG_Calcs_Builder.exe`.

## Suggested QA lanes after rebuild
- Oakland, CA → Cancun, Mexico by Road
- Rotterdam, Netherlands → New York City, NY by Sea
- Myrtle Beach, SC → Miami, FL by Inland Water
- Boston, MA → Ithaca, NY by Rail

## Notes
- The rail logic is materially improved, but it is still not a full FRA/NARN graph implementation.
- Sea routing is stronger than simple crow-flies estimation, but still intended for GHG distance estimation rather than vessel navigation.
- Output workbook distances are always written in miles.
- If a routing service is unavailable, the app still returns a calibrated estimate so non-Storage rows do not remain blank.
