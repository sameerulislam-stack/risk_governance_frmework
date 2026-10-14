# Build the Windows dashboard package

Run `build_exe.ps1` from PowerShell. The script uses the project's `.venv`
environment, installs application and build dependencies, builds a Windows
executable folder, and creates `release/JLI_Risk_Governance_Windows.zip`.
Application data is excluded from the source repository. Initialize
`data/risk.db` and place any authorized workbooks to bundle in `data` before
building.

The package starts the local web server, opens the dashboard in a browser, and
includes the current Equity database, W1/W2/W3 workbooks, and Trade Monitoring
workbook. On first start,
the bundled data is copied to `%LOCALAPPDATA%\JLI Risk Governance\data`; later
uploads and database updates are kept there. The recipient should extract the
whole folder from the ZIP and run `JLI Risk Governance.exe`.

The Trade Risk Analysis workbook is included. Later Trade Monitoring and
portfolio workbook uploads are saved in the user's data folder. The report
pages use Tailwind CSS and ApexCharts from public CDNs, so an internet
connection is required for their styling and charts.

The package contains the current Equity database and portfolio workbooks.
Share it only with people authorized to access those files. This local
executable is a per-user dashboard, not a hosted multi-user service.
