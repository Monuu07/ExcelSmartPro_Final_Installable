# ExcelSmart Pro — repaired build source

This repo contains the repaired source layout for building `ExcelSmart Pro.xlam` with a Windows/Excel-backed GitHub Actions workflow.

## Root cause fixed
The earlier sandbox-generated XLAM had a corrupted VBA module during packaging. The source modules in this repo are the corrected source-of-truth. Build the final XLAM on Windows using the workflow included here.

## Repo layout
- `.github/workflows/build_xlam.yml`
- `src/ExcelSmartPro.xlam/Modules/*.bas`
- `src/ExcelSmartPro.xlam/XMLsource/...`

## How to build on GitHub
1. Create a new GitHub repository.
2. Upload the CONTENTS of this zip to the repo root (do not upload the zip itself).
3. Open the `Actions` tab.
4. Run `Build ExcelSmart Pro XLAM`.
5. Wait for the workflow to finish.
6. Download the `ExcelSmartPro-XLAM` artifact.
7. Unzip the artifact and use the generated `.xlam` file in Excel.

## How to install in Excel
1. Right-click the `.xlam` file -> Properties -> Unblock.
2. Move it to a local folder such as `Documents\\ExcelSmartPro\\`.
3. In Excel: File -> Options -> Add-ins.
4. Manage: Excel Add-ins -> Go -> Browse.
5. Pick the `.xlam` file and enable it.
6. Restart Excel.
