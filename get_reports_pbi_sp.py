"""
Script to identify custom visuals in Power BI reports published in Microsoft Fabric.

This script uses the Power BI Scanner API to:
1. Authenticate with Azure AD using service principal
2. Scan workspaces for detailed metadata
3. Extract visual information from reports
4. Identify custom visuals
"""

import requests
import json
import time
import csv
from datetime import datetime
from msal import ConfidentialClientApplication
from typing import List, Dict, Optional

# Configuration
CLIENT_ID = "client-id"  # Service Principal (App) ID
TENANT_ID = "tenant-id"  # Tenant ID
CLIENT_SECRET = "secret-id"  # Service Principal Secret
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]
PBI_API_BASE = "https://api.powerbi.com/v1.0/myorg"


def get_access_token(client_id: str, tenant_id: str, client_secret: str) -> Optional[str]:
    """
    Authenticate and get access token using service principal (client credentials flow).
    """
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )
    
    # Acquire token using client credentials
    result = app.acquire_token_for_client(scopes=SCOPE)
    
    if "access_token" in result:
        print("OK Authentication successful with service principal")
        return result["access_token"]
    else:
        print(f"ERROR Authentication failed")
        print(f"Error: {result.get('error')}")
        print(f"Error description: {result.get('error_description')}")
        return None


def get_workspaces(access_token: str) -> List[Dict]:
    """
    Get all workspaces accessible to the user.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{PBI_API_BASE}/groups"
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    return response.json().get("value", [])


def get_reports_in_workspace(access_token: str, workspace_id: str) -> List[Dict]:
    """
    Get all reports in a specific workspace.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{PBI_API_BASE}/groups/{workspace_id}/reports"
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    return response.json().get("value", [])


def clone_report(access_token: str, workspace_id: str, report_id: str, report_name: str) -> Optional[str]:
    """
    Clone a report to try exporting the clone (may have fewer restrictions).
    Returns the clone ID if successful.
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    url = f"{PBI_API_BASE}/groups/{workspace_id}/reports/{report_id}/Clone"
    clone_name = f"temp_analysis_{report_id[:8]}"
    
    body = {
        "name": clone_name,
        "targetWorkspaceId": workspace_id
    }
    
    try:
        response = requests.post(url, headers=headers, json=body)
        if response.status_code in [200, 201]:
            clone_id = response.json().get("id")
            print(f"  Cloned as: {clone_name} (ID: {clone_id})")
            return clone_id
        else:
            return None
    except:
        return None


def delete_report(access_token: str, workspace_id: str, report_id: str):
    """Delete a report (cleanup clones)."""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{PBI_API_BASE}/groups/{workspace_id}/reports/{report_id}"
    
    try:
        requests.delete(url, headers=headers)
    except:
        pass


def export_report_as_pbix(access_token: str, workspace_id: str, report_id: str, is_clone: bool = False) -> Optional[bytes]:
    """
    Try to export/download report as PBIX file.
    This requires specific permissions but would give us full access to visuals.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{PBI_API_BASE}/groups/{workspace_id}/reports/{report_id}/Export"
    
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.content
        else:
            error_msg = response.json().get("error", {}).get("code", "Unknown")
            if not is_clone:  # Only print error for original, not clone attempts
                print(f"  Direct export failed: {error_msg}")
            return None
    except Exception as e:
        return None


def extract_visuals_from_pbix(pbix_content: bytes) -> List[Dict]:
    """
    Extract visual information from PBIX file.
    PBIX is a ZIP archive containing JSON files with report metadata.
    """
    import zipfile
    import io
    
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


def get_report_pages(access_token: str, workspace_id: str, report_id: str) -> list:
    """
    Get all pages in a report using the regular API.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{PBI_API_BASE}/groups/{workspace_id}/reports/{report_id}/pages"
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json().get("value", [])
    except Exception as e:
        print(f"Error getting pages: {e}")
        return []


def scan_workspace(access_token: str, workspace_id: str) -> Optional[str]:
    """
    Initiate a workspace scan using the Scanner API.
    Returns the scan ID.
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    # Use admin API endpoint
    url = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo"
    
    # Request body with ALL options enabled to get maximum metadata
    # Including visual information requires "Enhance admin APIs responses with detailed metadata" in Admin Portal
    body = {
        "workspaces": [workspace_id],
        "datasetExpressions": True,  # Enable to get DAX expressions
        "datasetSchema": True,       # Enable to get dataset schema
        "datasourceDetails": True,   # Enable to get datasource details
        "getArtifactUsers": True,    # Enable to get user info
        "lineage": True              # Enable to get lineage info
    }
    
    print(f"📤 Request body: {json.dumps(body, indent=2)}")
    
    response = requests.post(url, headers=headers, json=body)
    
    if response.status_code == 202:
        # Scan accepted, get scan ID from Location header
        location = response.headers.get("Location", "")
        scan_id = location.split("/")[-1] if location else None
        print(f"✓ Scan accepted - Scan ID: {scan_id}")
        return scan_id
    else:
        print(f"❌ Scan request failed: {response.status_code}")
        print(f"Response: {response.text}")
        print(f"Response headers: {dict(response.headers)}")
        return None


def get_scan_status(access_token: str, scan_id: str) -> Optional[str]:
    """
    Check the status of a workspace scan.
    Returns 'Succeeded', 'Running', or None if failed.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanStatus/{scan_id}"
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json().get("status")
    except Exception as e:
        print(f"Error checking scan status: {e}")
        return None


def get_scan_result(access_token: str, scan_id: str) -> Optional[Dict]:
    """
    Get the result of a completed workspace scan.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanResult/{scan_id}"
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"Error getting scan result: {e}")
        return None


def extract_visuals_from_scan(scan_data: Dict, debug: bool = False) -> Dict[str, List[Dict]]:
    """
    Extract visual information from Scanner API result.
    Returns a dictionary mapping report IDs to their visuals.
    """
    report_visuals = {}
    
    # Save scan data for debugging if needed
    if debug:
        with open("scan_debug.json", "w", encoding="utf-8") as f:
            json.dump(scan_data, f, indent=2)
        print("📝 Scan data saved to scan_debug.json")
    
    try:
        workspaces = scan_data.get("workspaces", [])
        
        if debug:
            print(f"\n🔍 DEBUG: Found {len(workspaces)} workspaces in scan data")
        
        for workspace in workspaces:
            reports = workspace.get("reports", [])
            
            if debug:
                print(f"🔍 DEBUG: Workspace has {len(reports)} reports")
            
            for report in reports:
                report_id = report.get("id")
                report_name = report.get("name", "Unknown")
                pages = report.get("pages", [])
                
                if debug:
                    print(f"\n🔍 DEBUG: Report '{report_name}' ({report_id})")
                    print(f"   Pages: {len(pages)}")
                
                all_visuals = []
                
                for page in pages:
                    page_name = page.get("name", "Unnamed Page")
                    visuals = page.get("visuals", [])
                    
                    if debug and visuals:
                        print(f"   Page '{page_name}': {len(visuals)} visuals")
                    
                    for visual in visuals:
                        visual_type = visual.get("visualType", "Unknown")
                        
                        if debug:
                            print(f"      - Type: {visual_type}")
                        
                        visual_info = {
                            "name": visual.get("name", "Unnamed"),
                            "type": visual_type,
                            "is_custom": is_custom_visual(visual_type),
                            "page": page_name
                        }
                        
                        all_visuals.append(visual_info)
                
                report_visuals[report_id] = {
                    "name": report_name,
                    "visuals": all_visuals
                }
    
    except Exception as e:
        print(f"Error extracting visuals from scan: {e}")
        import traceback
        traceback.print_exc()
    
    return report_visuals


def is_custom_visual(visual_type: str) -> bool:
    """
    Determine if a visual type is a custom visual.
    Built-in visuals have specific type names, custom visuals typically have longer identifiers.
    """
    if not visual_type or visual_type == "Unknown":
        return False
    
    # Common built-in visual types (comprehensive list)
    builtin_visuals = {
        "barChart", "clusteredBarChart", "clusteredColumnChart", "columnChart",
        "lineChart", "areaChart", "lineClusteredColumnComboChart", "lineStackedColumnComboChart",
        "pieChart", "donutChart", "funnel", "gauge", "card", "multiRowCard",
        "table", "matrix", "slicer", "map", "filledMap", "shape", "image",
        "textbox", "scatterChart", "pivotTable", "treemap", "waterfallChart",
        "hundredPercentStackedBarChart", "hundredPercentStackedColumnChart",
        "ribbonChart", "kpi", "decompositionTreeVisual",
        # Additional built-in visuals
        "stackedBarChart", "stackedColumnChart", "lineStackedAreaChart",
        "hundredPercentStackedAreaChart", "stackedAreaChart",
        "ribbon", "actionButton", "basicShape"
    }
    
    # Check if it's a known built-in visual
    if visual_type in builtin_visuals:
        return False
    
    # Custom visuals have specific patterns:
    # 1. Contains alphanumeric GUID-like strings (longer identifiers)
    # 2. Contains dots (package notation like "publisher.visualname")
    # 3. Has very long names (>25 characters)
    
    if "." in visual_type:  # e.g., "PBI_CV_xxxxx" or "publisher.visual"
        return True
    
    if len(visual_type) > 25:  # Very long names are typically custom
        return True
    
    # If starts with common custom visual prefixes
    custom_prefixes = ["PBI_CV", "custom", "Custom"]
    if any(visual_type.startswith(prefix) for prefix in custom_prefixes):
        return True
    
    # Default: if not in built-in list and matches patterns, consider it custom
    # This is more conservative - unknown short names won't be flagged
    return False


def analyze_workspace_reports(access_token: str, workspace_id: str, workspace_name: str) -> List[Dict]:
    """
    Analyze all reports in a workspace for custom visuals.
    Attempts multiple methods: Direct export, Clone+Export, and page listing.
    Returns list of dictionaries with analysis results.
    """
    print(f"\n{'='*80}")
    print(f"Analyzing workspace: {workspace_name}")
    print(f"{'='*80}")
    
    reports = get_reports_in_workspace(access_token, workspace_id)
    print(f"Found {len(reports)} reports\n")
    
    results = []
    
    for report in reports:
        report_name = report.get("name", "Unknown")
        report_id = report.get("id")
        
        # Skip if it's already a temp analysis clone
        if "temp_analysis_" in report_name or "temp_clone_for_analysis" in report_name:
            continue
        
        print(f"\n{'-'*80}")
        print(f"Report: {report_name}")
        print(f"Report ID: {report_id}")
        
        # Initialize result record
        result = {
            "workspace": workspace_name,
            "report": report_name,
            "report_id": report_id,
            "method": "Failed",
            "num_pages": 0,
            "is_directlake": "Unknown",
            "total_visuals": 0,
            "custom_visuals": 0
        }
        
        pbix_content = None
        clone_id = None
        
        # METHOD 1: Try direct PBIX export
        print("  [Method 1] Direct PBIX export...")
        pbix_content = export_report_as_pbix(access_token, workspace_id, report_id, is_clone=False)
        is_directlake = False
        clone_id = None
        
        # METHOD 2: If direct export fails, try clone + export
        if not pbix_content:
            print("  [Method 2] Clone + Export approach...")
            is_directlake = True  # Likely DirectLake if export failed
            result["is_directlake"] = "Yes"
            clone_id = clone_report(access_token, workspace_id, report_id, report_name)
            
            if clone_id:
                # Wait a moment for clone to be ready
                import time
                time.sleep(2)
                
                # Try to export the clone
                print(f"  Attempting to export clone...")
                pbix_content = export_report_as_pbix(access_token, workspace_id, clone_id, is_clone=True)
                
                if pbix_content:
                    print(f"  SUCCESS Clone exported ({len(pbix_content)} bytes)")
                else:
                    print(f"  Clone export also failed (DirectLake restriction)")
        
        # If we got PBIX content, extract visuals
        if pbix_content:
            print(f"  Extracting visuals from PBIX...")
            
            # Save PBIX for inspection
            pbix_filename = f"report_{report_id[:8]}.pbix"
            with open(pbix_filename, "wb") as f:
                f.write(pbix_content)
            print(f"  Saved PBIX: {pbix_filename}")
            
            visuals = extract_visuals_from_pbix(pbix_content)
            
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
                
                # Filter custom visuals
                custom_visuals = [v for v in visuals if v["is_custom"]]
                
                # Update result
                result["method"] = "Direct Export"
                result["total_visuals"] = len(visuals)
                result["custom_visuals"] = len(custom_visuals)
                result["is_directlake"] = "No"
                result["num_pages"] = len(pages)
                
                if custom_visuals:
                    print(f"\n  ** CUSTOM VISUALS FOUND ({len(custom_visuals)}) **")
                    for visual in custom_visuals:
                        print(f"    - {visual['name']}")
                        print(f"      Type: {visual['type']}")
                        print(f"      Page: {visual['page']}")
                else:
                    print(f"\n  No custom visuals detected")
            else:
                print("  WARNING: Could not extract visual information from PBIX")
                result["method"] = "Direct Export (No Visuals)"
                result["is_directlake"] = "No"
        else:
            # METHOD 3: Fallback to page listing only
            print("  [Method 3] Basic page listing (no visual details)...")
            pages = get_report_pages(access_token, workspace_id, report_id)
            
            if pages:
                print(f"  Report has {len(pages)} page(s):")
                for page in pages:
                    print(f"    - {page.get('displayName')}")
                print(f"\n  NOTE: Cannot extract visual details via API")
                print(f"  LINK: {report.get('webUrl', 'N/A')}")
                
                result["method"] = "Page Listing Only"
                result["num_pages"] = len(pages)
            else:
                print("  ERROR: Could not retrieve page information")
                result["method"] = "Failed"
        
        # Cleanup: delete clone if created
        if clone_id:
            print(f"  Cleaning up clone...")
            delete_report(access_token, workspace_id, clone_id)
        
        results.append(result)
    
    return results


def main():
    """
    Main function to scan Power BI reports for custom visuals.
    """
    print("Power BI Custom Visual Identifier")
    print("==================================\n")
    
    # Get access token
    print("Authenticating with service principal...")
    access_token = get_access_token(CLIENT_ID, TENANT_ID, CLIENT_SECRET)
    
    if not access_token:
        print("Failed to authenticate")
        return
    
    # Get workspaces
    print("Fetching workspaces...")
    workspaces = get_workspaces(access_token)
    print(f"Found {len(workspaces)} workspaces\n")
    
    # Collect all results
    all_results = []
    
    # Option 1: Analyze all workspaces
    for workspace in workspaces:
        workspace_name = workspace.get("name", "Unknown")
        workspace_id = workspace.get("id")
        
        try:
            results = analyze_workspace_reports(access_token, workspace_id, workspace_name)
            all_results.extend(results)
        except Exception as e:
            print(f"Error analyzing workspace {workspace_name}: {e}")
    
    # Generate CSV report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"pbi_custom_visuals_report_{timestamp}.csv"
    
    print(f"\n{'='*80}")
    print("Generating CSV report...")
    
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['workspace', 'report', 'report_id', 'method', 'num_pages', 
                      'is_directlake', 'total_visuals', 'custom_visuals']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for result in all_results:
            writer.writerow(result)
    
    # Summary
    total_reports = len(all_results)
    reports_with_custom = sum(1 for r in all_results if r['custom_visuals'] > 0)
    directlake_reports = sum(1 for r in all_results if r['is_directlake'] == 'Yes')
    successful_exports = sum(1 for r in all_results if 'Export' in r['method'])
    
    print(f"\nCSV report generated: {csv_filename}")
    print(f"{'='*80}")
    print(f"SUMMARY:")
    print(f"  Total reports analyzed: {total_reports}")
    print(f"  Reports with custom visuals: {reports_with_custom}")
    print(f"  DirectLake reports: {directlake_reports}")
    print(f"  Successful PBIX exports: {successful_exports}")
    print(f"{'='*80}\n")
    
    # Option 2: Analyze specific workspace (uncomment and modify as needed)
    # specific_workspace_id = "YOUR_WORKSPACE_ID"
    # analyze_workspace_reports(access_token, specific_workspace_id, "My Workspace")


if __name__ == "__main__":
    main()
