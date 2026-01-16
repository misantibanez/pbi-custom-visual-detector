# Power BI Custom Visual Identifier

This tool helps identify custom visuals in Power BI reports published in Microsoft Fabric by exporting and analyzing PBIX files.

## Overview

The tool provides two authentication methods to scan Power BI reports and detect custom visuals:

- **`get_reports_pbi_interactive.py`**: For user authentication (Admin users)
- **`get_reports_pbi_sp.py`**: For service principal authentication (automated scenarios)

Both scripts:
1. Connect to Power BI using the specified authentication method
2. Scan all accessible workspaces and reports
3. Export reports as PBIX files (when possible)
4. Extract and analyze visual information from the PBIX Layout JSON
5. Generate a CSV report with detailed results

## Prerequisites

### For Interactive User Authentication (`get_reports_pbi_interactive.py`)

1. **Admin User Account**
   - Power BI Administrator role or workspace access
   - No app registration needed (uses Azure CLI public client)

2. **Enable Power BI Service**
   - Go to [Power BI Admin Portal](https://app.powerbi.com/admin-portal/tenantSettings)
   - Ensure you have access to the workspaces you want to scan

### For Service Principal Authentication (`get_reports_pbi_sp.py`)

1. **Azure AD Service Principal (App Registration)**
   - Go to [Azure Portal](https://portal.azure.com)
   - Navigate to Azure Active Directory > App Registrations
   - Click "New registration"
   - Name: "PowerBI Custom Visual Scanner"
   - Supported account types: "Accounts in this organizational directory only"
   - Click "Register"

2. **Create Client Secret**
   - In your app registration, go to "Certificates & secrets"
   - Click "New client secret"
   - Add a description and select expiration period
   - Click "Add"
   - **⚠️ IMPORTANT**: Copy the secret value immediately (you won't be able to see it again)

3. **Configure API Permissions**
   - In your app registration, go to "API permissions"
   - Click "Add a permission"
   - Select "Power BI Service"
   - Select "Application permissions" (not Delegated)
   - Add these permissions:
     - `Tenant.Read.All` - Required for tenant-wide access
     - `Tenant.ReadWrite.All` - May be needed for export operations
   - Click "Add permissions"
   - Click "Grant admin consent for [Your Organization]" (admin consent required)

4. **Enable Power BI Service Admin**
   - Go to [Power BI Admin Portal](https://app.powerbi.com/admin-portal/tenantSettings)
   - Under "Developer settings" > "Service principals can use Power BI APIs"
   - Enable and add your service principal to the list (or a security group containing it)

5. **Note about Workspace Access**
   - With tenant-level permissions, the service principal can access all workspaces
   - You DON'T need to add it as Member/Admin to individual workspaces
   - However, ensure the service principal has proper API permissions enabled

6. **Get Your IDs**
   - From App Registration:
     - Copy the "Application (client) ID"
     - Copy the "Directory (tenant) ID"
     - Copy the "Client Secret" (from step 2)

## Setup

1. **Install dependencies:**
   ```powershell
   pip install -r requirements.txt
   ```

2. **Configure the script:**
   
   **For Interactive Authentication:**
   - Open `get_reports_pbi_interactive.py`
   - Update `TENANT_ID` with your tenant ID
   - No other configuration needed (uses device code flow)
   
   **For Service Principal:**
   - Open `get_reports_pbi_sp.py`
   - Replace `CLIENT_ID` with your application (client) ID
   - Replace `TENANT_ID` with your directory (tenant) ID
   - Replace `CLIENT_SECRET` with your client secret value

## Usage

### Interactive User Authentication

```powershell
python get_reports_pbi_interactive.py
```

- You'll be prompted to visit https://microsoft.com/devicelogin
- Enter the provided code
- Sign in with your admin account
- The script will continue automatically after authentication

### Service Principal Authentication

```powershell
python get_reports_pbi_sp.py
```

- Authenticates automatically with client credentials (no user interaction needed)
- Ideal for scheduled/automated scans

## Output

Both scripts generate:

### Console Output
For each report, you'll see:
- Report name and ID
- Export method used (Direct Export, Page Listing Only, Failed)
- Number of pages
- Total visuals found
- Custom visuals detected (if any)

### CSV Report
A timestamped CSV file (`pbi_custom_visuals_report_YYYYMMDD_HHMMSS.csv`) with columns:
- **workspace**: Workspace name
- **report**: Report name
- **report_id**: Unique report ID
- **method**: Method used (Direct Export, Page Listing Only, Failed)
- **num_pages**: Number of pages in the report
- **is_directlake**: Whether the report uses DirectLake (Yes/No/Unknown)
- **total_visuals**: Total number of visuals detected
- **custom_visuals**: Number of custom visuals detected

### Summary Statistics
- Total reports analyzed
- Reports with custom visuals
- DirectLake reports (cannot be exported)
- Successful PBIX exports

## How It Works

1. **Authentication**: 
   - Interactive: Device Code Flow with user credentials
   - Service Principal: Client credentials flow (app + secret)

2. **PBIX Export**: Attempts to export reports as PBIX files using Power BI REST API

3. **PBIX Parsing**: 
   - PBIX files are ZIP archives
   - Extracts `Report/Layout` file containing JSON metadata
   - Decodes using UTF-16-LE encoding (PBIX standard)
   - Parses visual definitions from `sections[].visualContainers[]`

4. **Custom Visual Detection**: 
   - Compares visual types against built-in visuals list
   - Identifies custom visuals by patterns:
     - Contains dots (e.g., `publisher.visualname`)
     - Very long names (>25 characters)
     - GUID-like identifiers
     - Prefixes like `PBI_CV_`

## Limitations

### DirectLake Restriction (Critical)
- **Microsoft Fabric DirectLake datasets CANNOT be exported as PBIX**
- Error: `ExportData_DisabledForModelWithDirectLakeMode`
- This is a platform restriction, not a permissions issue
- Affects most modern Fabric reports
- The script will detect and mark these as DirectLake in the CSV

### Workarounds
- Only non-DirectLake reports (Import/DirectQuery) can be fully analyzed
- DirectLake reports show page count but no visual details
- For complete analysis, reports must use traditional datasets

### Other Limitations
- Requires appropriate permissions to export PBIX files
- Only works with reports you have access to
- Admin monitoring reports have special restrictions
- Some custom visuals might be misidentified if they use naming conventions similar to built-in visuals

## Troubleshooting

### Interactive Authentication Issues

**Device code not working:**
- Ensure you can access https://microsoft.com/devicelogin
- Check if your account has 2FA enabled (device flow supports it)
- Try clearing browser cookies/cache

**Authentication succeeds but can't access reports:**
- Verify your account has Power BI Admin role or workspace access
- Check workspace permissions

### Service Principal Issues

**Authentication fails:**
- Ensure your service principal has the correct API permissions (Application permissions, not Delegated)
- Verify admin consent has been granted for the tenant
- Check that service principals are enabled in Power BI Admin Portal
- Verify the client secret hasn't expired

**Can't export reports (403 Forbidden):**
- Enable "Service principals can use Power BI APIs" in Power BI Admin Portal
- Add your service principal to the allowed list or security group
- Verify you have `Tenant.ReadWrite.All` permission with admin consent

### Export Issues

**"ExportData_DisabledForModelWithDirectLakeMode" error:**
- This is expected for DirectLake reports
- The script will detect and report these
- No workaround available - Microsoft platform limitation

**"PowerBINotAuthorizedException" for admin reports:**
- Admin monitoring reports require special tenant admin permissions
- These reports cannot be exported even with service principal

**No visuals detected in exported PBIX:**
- Check that the PBIX file was saved successfully
- Verify file size is reasonable (>0 bytes)
- Inspect the PBIX manually using `inspect_pbix.py` tool

### General Issues

**Script runs slow:**
- Large tenants with many workspaces take time
- Consider filtering to specific workspaces
- Network latency affects export speed

**CSV file not generated:**
- Check for file permission issues
- Verify the script completed without errors
- Look for error messages in console output

## Architecture Notes

- Uses Power BI REST API v1.0 (`/myorg` endpoints)
- PBIX files are saved temporarily as `report_{id}.pbix` for inspection
- UTF-16-LE decoding is required for PBIX Layout files
- Built-in visual detection uses comprehensive type list
