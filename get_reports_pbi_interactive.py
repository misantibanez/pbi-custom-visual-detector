"""
Script to identify custom visuals in Power BI reports - User Authentication Version

Uses Device Code Flow for interactive user authentication (no password needed).
"""

import requests
import json
import time
import zipfile
import io
import csv
import os
from datetime import datetime
from msal import PublicClientApplication
from typing import List, Dict, Optional

# Configuration
CLIENT_ID = "client-id"  # Azure CLI Public Client ID (Microsoft-owned)
TENANT_ID = "tenant-id"  # Your Tenant ID
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]
PBI_API_BASE = "https://api.powerbi.com/v1.0/myorg"


def get_access_token_interactive() -> Optional[str]:
    """
    Authenticate using Device Code Flow (user interactive).
    User will see a code to enter at microsoft.com/devicelogin
    """
    app = PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY
    )
    
    # Try to get token from cache first
    accounts = app.get_accounts()
    if accounts:
        print("Found cached authentication, attempting silent login...")
        result = app.acquire_token_silent(SCOPE, account=accounts[0])
        if result and "access_token" in result:
            print("OK Authentication successful (cached)")
            return result["access_token"]
    
    # If no cache, use device flow
    flow = app.initiate_device_flow(scopes=SCOPE)
    
    if "user_code" not in flow:
        print(f"ERROR Failed to create device flow: {flow.get('error_description')}")
        return None
    
    print("\n" + "="*60)
    print("AUTHENTICATION REQUIRED")
    print("="*60)
    print(flow["message"])
    print("="*60 + "\n")
    
    # Wait for user to authenticate
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        print("OK Authentication successful!")
        return result["access_token"]
    else:
        print(f"ERROR Authentication failed")
        print(f"Error: {result.get('error')}")
        print(f"Error description: {result.get('error_description')}")
        return None


def get_workspaces(access_token: str, use_admin_api: bool = True, exclude_personal: bool = True,
                   capacity_ids: List[str] = None) -> List[Dict]:
    """Get all workspaces. Use admin API to get ALL workspaces in tenant."""
    headers = {"Authorization": f"Bearer {access_token}"}
    
    if use_admin_api:
        # Admin API returns ALL workspaces in the tenant
        url = f"{PBI_API_BASE}/admin/groups?$top=5000"
    else:
        # Regular API only returns workspaces user has access to
        url = f"{PBI_API_BASE}/groups"
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    workspaces = response.json().get("value", [])
    
    # Filter out personal workspaces if requested
    if exclude_personal:
        workspaces = [ws for ws in workspaces if ws.get("type") != "PersonalGroup"]
    
    # Filter by capacity IDs if provided
    if capacity_ids:
        capacity_ids_lower = [c.lower() for c in capacity_ids]
        workspaces = [ws for ws in workspaces if ws.get("capacityId", "").lower() in capacity_ids_lower]
    
    return workspaces


def get_reports_in_workspace(access_token: str, workspace_id: str) -> List[Dict]:
    """Get all reports in a specific workspace."""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{PBI_API_BASE}/groups/{workspace_id}/reports"
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    return response.json().get("value", [])


def get_report_pages(access_token: str, workspace_id: str, report_id: str) -> List[Dict]:
    """Get pages in a report."""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{PBI_API_BASE}/groups/{workspace_id}/reports/{report_id}/pages"
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json().get("value", [])
    except:
        return []


def export_report_as_pbix(access_token: str, workspace_id: str, report_id: str) -> Optional[bytes]:
    """
    Export report as PBIX file.
    Returns bytes of the PBIX file if successful.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{PBI_API_BASE}/groups/{workspace_id}/reports/{report_id}/Export"
    
    try:
        response = requests.get(url, headers=headers, timeout=60)
        if response.status_code == 200:
            return response.content
        else:
            error_msg = response.text
            if "ExportData_DisabledForModelWithDirectLakeMode" in error_msg:
                return None  # DirectLake restriction
            return None
    except Exception as e:
        return None


def extract_visuals_from_pbix(pbix_content: bytes) -> List[Dict]:
    """
    Extract visual information from PBIX file.
    PBIX is a ZIP archive containing JSON files with report metadata.
    """
    visuals = []
    
    try:
        with zipfile.ZipFile(io.BytesIO(pbix_content)) as zip_file:
            # Look for Layout files which contain visual definitions
            for file_name in zip_file.namelist():
                if "Layout" in file_name and not file_name.endswith("/"):
                    print(f"    Found layout file: {file_name}")
                    
                    try:
                        # PBIX files typically use UTF-16 LE encoding
                        layout_content = zip_file.read(file_name).decode('utf-16-le')
                        layout_data = json.loads(layout_content)
                        
                        # Parse sections and visual containers
                        if "sections" in layout_data:
                            for section in layout_data["sections"]:
                                section_name = section.get("displayName", "Unnamed Section")
                                
                                if "visualContainers" in section:
                                    for container in section["visualContainers"]:
                                        if "config" in container:
                                            config_str = container["config"]
                                            config = json.loads(config_str)
                                            
                                            # Extract visual type
                                            visual_type = config.get("singleVisual", {}).get("visualType", "Unknown")
                                            
                                            visual_info = {
                                                "name": config.get("name", "Unnamed"),
                                                "type": visual_type,
                                                "is_custom": is_custom_visual(visual_type),
                                                "page": section_name
                                            }
                                            
                                            visuals.append(visual_info)
                    except UnicodeDecodeError:
                        # Try UTF-8 if UTF-16 fails
                        try:
                            layout_content = zip_file.read(file_name).decode('utf-8')
                            layout_data = json.loads(layout_content)
                            
                            if "sections" in layout_data:
                                for section in layout_data["sections"]:
                                    section_name = section.get("displayName", "Unnamed Section")
                                    
                                    if "visualContainers" in section:
                                        for container in section["visualContainers"]:
                                            if "config" in container:
                                                config_str = container["config"]
                                                config = json.loads(config_str)
                                                
                                                visual_type = config.get("singleVisual", {}).get("visualType", "Unknown")
                                                
                                                visual_info = {
                                                    "name": config.get("name", "Unnamed"),
                                                    "type": visual_type,
                                                    "is_custom": is_custom_visual(visual_type),
                                                    "page": section_name
                                                }
                                                
                                                visuals.append(visual_info)
                        except Exception as e2:
                            print(f"    Error decoding layout: {e2}")
    except Exception as e:
        print(f"  Error extracting visuals from PBIX: {e}")
    
    return visuals


def is_custom_visual(visual_type: str) -> bool:
    """
    Determine if a visual type is a custom visual.
    Built-in visuals have simple names like 'clusteredBarChart', 'lineChart', etc.
    Custom visuals typically have longer names with dots or special patterns.
    """
    # List of known built-in visual types
    builtin_visuals = {
        'clusteredBarChart', 'clusteredColumnChart', 'hundredPercentStackedBarChart',
        'hundredPercentStackedColumnChart', 'lineChart', 'areaChart', 'stackedAreaChart',
        'lineStackedColumnComboChart', 'lineClusteredColumnComboChart', 'ribbonChart',
        'waterfallChart', 'funnelChart', 'scatterChart', 'pieChart', 'donutChart',
        'gauge', 'card', 'multiRowCard', 'kpi', 'slicer', 'table', 'matrix',
        'filledMap', 'map', 'shape', 'image', 'textbox', 'treemap', 'basicShape',
        'actionButton', 'columnChart', 'barChart', 'pivotTable'
    }
    
    # If it's in the built-in list, it's not custom
    if visual_type.lower() in builtin_visuals:
        return False
    
    # Custom visuals often have:
    # - Dots in the name (e.g., 'PBI_CV_xxxxxxxx' or 'organization.visualName')
    # - Very long names (>25 chars)
    # - Special prefixes like 'PBI_CV_'
    if '.' in visual_type or len(visual_type) > 25 or visual_type.startswith('PBI_CV_'):
        return True
    
    return False


def analyze_workspace_reports(access_token: str, workspace_id: str, workspace_name: str, capacity_id: str = "") -> List[Dict]:
    """Analyze all reports in a workspace. Returns list of analysis results."""
    print(f"\n{'='*64}")
    print(f"{'='*16}                                                Analyzing workspace: {workspace_name}")
    print(f"{'='*64}")
    print(f"{'='*16}                                                ", end="")
    
    # Get reports
    reports = get_reports_in_workspace(access_token, workspace_id)
    print(f"Found {len(reports)} reports\n")
    
    results = []
    
    for report in reports:
        report_name = report.get("name", "Unnamed Report")
        report_id = report.get("id")
        web_url = report.get("webUrl", "")
        
        print(f"\n{'-'*64}")
        print(f"{'-'*16}                                                Report: {report_name}")
        print(f"Report ID: {report_id}")
        
        # Initialize result record
        result = {
            "workspace": workspace_name,
            "workspace_id": workspace_id,
            "capacity_id": capacity_id,
            "report": report_name,
            "report_id": report_id,
            "method": "Failed",
            "num_pages": 0,
            "is_directlake": "Unknown",
            "total_visuals": 0,
            "custom_visuals": 0
        }
        
        # Try to export and analyze PBIX
        print(f"  Attempting PBIX export...")
        pbix_content = export_report_as_pbix(access_token, workspace_id, report_id)
        
        if pbix_content:
            print(f"  Extracting visuals from PBIX...")
            
            # Save PBIX for debugging
            filename = f"report_{report_id[:8]}.pbix"
            with open(filename, 'wb') as f:
                f.write(pbix_content)
            print(f"  Saved PBIX: {filename}")
            
            # Extract visuals
            visuals = extract_visuals_from_pbix(pbix_content)
            
            # Delete PBIX after analysis
            try:
                os.remove(filename)
                print(f"  Deleted PBIX: {filename}")
            except Exception as e:
                print(f"  Warning: Could not delete PBIX: {e}")
            
            if visuals:
                print(f"  Total visuals found: {len(visuals)}")
                
                # Group by page
                pages = {}
                for visual in visuals:
                    page = visual["page"]
                    if page not in pages:
                        pages[page] = []
                    pages[page].append(visual)
                
                print(f"\n  Report structure:")
                for page_name, page_visuals in pages.items():
                    print(f"    Page '{page_name}': {len(page_visuals)} visuals")
                
                # Check for custom visuals
                custom_visuals = [v for v in visuals if v["is_custom"]]
                
                # Update result
                result["method"] = "Direct Export"
                result["total_visuals"] = len(visuals)
                result["custom_visuals"] = len(custom_visuals)
                result["is_directlake"] = "No"
                result["num_pages"] = len(pages)
                
                if custom_visuals:
                    print(f"\n  CUSTOM VISUALS DETECTED ({len(custom_visuals)}):")
                    for cv in custom_visuals:
                        print(f"    - Type: {cv['type']}")
                        print(f"      Page: {cv['page']}")
                        print(f"      Name: {cv['name']}")
                else:
                    print(f"\n  No custom visuals detected")
            else:
                print(f"  WARNING: Could not extract visual information from PBIX")
                result["method"] = "Direct Export (No Visuals)"
                result["is_directlake"] = "No"
        else:
            print(f"  Export failed (likely DirectLake restriction)")
            result["is_directlake"] = "Yes"
            
            # Try to at least get page info
            pages = get_report_pages(access_token, workspace_id, report_id)
            if pages:
                print(f"  Report has {len(pages)} page(s):")
                for page in pages:
                    print(f"    - {page.get('displayName', page.get('name', 'Unnamed'))}")
                
                result["method"] = "Page Listing Only"
                result["num_pages"] = len(pages)
            else:
                result["method"] = "Failed"
            
            print(f"\n  NOTE: Cannot extract visual details via API")
        
        print(f"  LINK: {web_url}")
        results.append(result)
    
    return results


def main():
    print("Power BI Custom Visual Identifier - Interactive User Auth")
    print("="*60)
    print()
    
    # Authenticate
    print("Authenticating with user credentials...")
    access_token = get_access_token_interactive()
    
    if not access_token:
        print("ERROR: Failed to authenticate")
        return
    
    # Ask for capacity filter
    print("\nFilter by Capacity ID?")
    print("  - Enter capacity IDs separated by comma")
    print("  - Or press Enter to show all workspaces")
    capacity_input = input("Capacity IDs: ").strip()
    
    capacity_ids = None
    if capacity_input:
        capacity_ids = [c.strip() for c in capacity_input.split(",") if c.strip()]
        print(f"Filtering by {len(capacity_ids)} capacity ID(s)")
    
    # Get workspaces
    print("\nFetching workspaces...")
    workspaces = get_workspaces(access_token, capacity_ids=capacity_ids)
    print(f"Found {len(workspaces)} workspaces\n")
    
    # Create CSV file and write header immediately
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"pbi_custom_visuals_report_{timestamp}.csv"
    fieldnames = ['workspace', 'workspace_id', 'capacity_id', 'report', 'report_id', 'method', 'num_pages', 
                  'is_directlake', 'total_visuals', 'custom_visuals']
    
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
    
    print(f"CSV file created: {csv_filename}")
    print("Results will be saved progressively...\n")
    
    # Collect all results
    all_results = []
    
    # Analyze each workspace
    for workspace in workspaces:
        workspace_name = workspace.get("name", "Unnamed Workspace")
        workspace_id = workspace.get("id")
        capacity_id = workspace.get("capacityId", "")
        
        results = analyze_workspace_reports(access_token, workspace_id, workspace_name, capacity_id)
        all_results.extend(results)
        
        # Append results to CSV after each workspace
        if results:
            with open(csv_filename, 'a', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                for result in results:
                    writer.writerow(result)
            print(f"  [Saved {len(results)} report(s) to CSV]")
    
    # Summary
    print(f"\n{'='*60}")
    total_reports = len(all_results)
    reports_with_custom = sum(1 for r in all_results if r['custom_visuals'] > 0)
    directlake_reports = sum(1 for r in all_results if r['is_directlake'] == 'Yes')
    successful_exports = sum(1 for r in all_results if 'Export' in r['method'])
    
    print(f"\nCSV report generated: {csv_filename}")
    print(f"{'='*60}")
    print(f"SUMMARY:")
    print(f"  Total reports analyzed: {total_reports}")
    print(f"  Reports with custom visuals: {reports_with_custom}")
    print(f"  DirectLake reports: {directlake_reports}")
    print(f"  Successful PBIX exports: {successful_exports}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
