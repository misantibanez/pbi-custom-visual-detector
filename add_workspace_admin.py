"""
Script to add a user as Admin to a Power BI workspace.

Uses Device Code Flow for interactive user authentication.
Requires Fabric Administrator permissions to add users to any workspace.
"""

import requests
from msal import PublicClientApplication
from typing import Optional, List, Dict

# Configuration
CLIENT_ID = "client-id"  # Azure CLI Public Client ID
TENANT_ID = "tenant-id"  # Your Tenant ID
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]
PBI_API_BASE = "https://api.powerbi.com/v1.0/myorg"


def get_access_token_interactive() -> Optional[str]:
    """Authenticate using Device Code Flow."""
    app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    
    accounts = app.get_accounts()
    if accounts:
        print("Found cached authentication, attempting silent login...")
        result = app.acquire_token_silent(SCOPE, account=accounts[0])
        if result and "access_token" in result:
            print("✓ Authentication successful (cached)")
            return result["access_token"]
    
    flow = app.initiate_device_flow(scopes=SCOPE)
    
    if "user_code" not in flow:
        print(f"✗ Failed to create device flow: {flow.get('error_description')}")
        return None
    
    print("\n" + "="*60)
    print("AUTHENTICATION REQUIRED")
    print("="*60)
    print(flow["message"])
    print("="*60 + "\n")
    
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        print("✓ Authentication successful!")
        return result["access_token"]
    else:
        print(f"✗ Authentication failed: {result.get('error_description')}")
        return None


def get_workspaces(access_token: str, use_admin_api: bool = True, exclude_personal: bool = True) -> List[Dict]:
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
    
    return workspaces


def get_workspace_users(access_token: str, workspace_id: str, use_admin_api: bool = True) -> List[Dict]:
    """Get all users in a workspace."""
    headers = {"Authorization": f"Bearer {access_token}"}
    
    if use_admin_api:
        url = f"{PBI_API_BASE}/admin/groups/{workspace_id}/users"
    else:
        url = f"{PBI_API_BASE}/groups/{workspace_id}/users"
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json().get("value", [])
    except:
        return []


def user_exists_in_workspace(access_token: str, workspace_id: str, user_email: str) -> bool:
    """Check if a user already has access to a workspace."""
    users = get_workspace_users(access_token, workspace_id, use_admin_api=True)
    user_email_lower = user_email.lower()
    
    for user in users:
        email = user.get("emailAddress", "").lower()
        upn = user.get("userPrincipalName", "").lower()
        if user_email_lower == email or user_email_lower == upn:
            return True
    return False


def add_user_to_workspace(access_token: str, workspace_id: str, user_email: str, 
                          access_right: str = "Admin", use_admin_api: bool = True) -> bool:
    """
    Add a user to a workspace with specified access right.
    
    Args:
        access_token: Bearer token for authentication
        workspace_id: GUID of the workspace
        user_email: Email address of the user to add
        access_right: One of 'Admin', 'Contributor', 'Member', 'Viewer'
        use_admin_api: If True, uses Admin API (requires Fabric Admin permissions)
    
    Returns:
        True if successful, False otherwise
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    # Use Admin API for tenant-wide access
    if use_admin_api:
        url = f"{PBI_API_BASE}/admin/groups/{workspace_id}/users"
    else:
        url = f"{PBI_API_BASE}/groups/{workspace_id}/users"
    
    payload = {
        "emailAddress": user_email,
        "groupUserAccessRight": access_right
    }
    
    try:
        # First check if user already exists in workspace
        if user_exists_in_workspace(access_token, workspace_id, user_email):
            print(f"ℹ User already has access to this workspace")
            return True
        
        response = requests.post(url, headers=headers, json=payload)
        
        if response.status_code == 200:
            print(f"✓ Successfully added '{user_email}' as {access_right} to workspace")
            return True
        elif response.status_code == 400:
            response_text = response.text
            
            # User already exists in workspace
            if "AlreadyExists" in response_text or "already exists" in response_text.lower():
                print(f"ℹ User already has access to this workspace")
                return True
            elif "NotSupported" in response_text:
                print(f"✗ Operation not supported for this workspace type")
                return False
            else:
                print(f"✗ Failed to add user: {response_text}")
                return False
        elif response.status_code == 401:
            print(f"✗ Not authorized for this workspace")
            return False
        else:
            print(f"✗ Failed to add user. Status: {response.status_code}")
            print(f"  Response: {response.text}")
            return False
            
    except Exception as e:
        print(f"✗ Error adding user: {e}")
        return False


def update_user_in_workspace(access_token: str, workspace_id: str, user_email: str, 
                             access_right: str = "Admin", use_admin_api: bool = True) -> bool:
    """
    Update an existing user's permissions in a workspace.
    
    Args:
        access_token: Bearer token for authentication
        workspace_id: GUID of the workspace
        user_email: Email address of the user to update
        access_right: One of 'Admin', 'Contributor', 'Member', 'Viewer'
        use_admin_api: If True, uses Admin API (requires Fabric Admin permissions)
    
    Returns:
        True if successful, False otherwise
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    # Use Admin API for tenant-wide access
    if use_admin_api:
        url = f"{PBI_API_BASE}/admin/groups/{workspace_id}/users"
    else:
        url = f"{PBI_API_BASE}/groups/{workspace_id}/users"
    
    payload = {
        "emailAddress": user_email,
        "groupUserAccessRight": access_right
    }
    
    try:
        response = requests.put(url, headers=headers, json=payload)
        
        if response.status_code == 200:
            print(f"✓ Successfully updated '{user_email}' to {access_right}")
            return True
        else:
            print(f"✗ Failed to update user. Status: {response.status_code}")
            print(f"  Response: {response.text}")
            return False
            
    except Exception as e:
        print(f"✗ Error updating user: {e}")
        return False


def find_workspace_by_name(workspaces: List[Dict], name: str) -> Optional[Dict]:
    """Find a workspace by name (case-insensitive partial match)."""
    name_lower = name.lower()
    for ws in workspaces:
        if name_lower in ws.get("name", "").lower():
            return ws
    return None


def main():
    print("Power BI Workspace Admin Manager")
    print("="*60)
    print()
    
    # Authenticate
    print("Authenticating...")
    access_token = get_access_token_interactive()
    
    if not access_token:
        print("✗ Failed to authenticate")
        return
    
    # Get workspaces
    print("\nFetching workspaces...")
    workspaces = get_workspaces(access_token)
    print(f"Found {len(workspaces)} workspaces\n")
    
    # List workspaces
    print("Available workspaces:")
    print("-"*60)
    for i, ws in enumerate(workspaces, 1):
        print(f"  {i}. {ws['name']}")
        print(f"     ID: {ws['id']}")
    print("-"*60)
    
    # Get user input
    print("\nOptions:")
    print("  - Enter workspace number (1, 2, etc.)")
    print("  - Enter workspace name (partial match)")
    print("  - Enter workspace ID (GUID)")
    print("  - Enter 'all' to add user to ALL workspaces")
    print()
    
    workspace_input = input("Select workspace: ").strip()
    
    if not workspace_input:
        print("No workspace selected. Exiting.")
        return
    
    # Determine target workspaces
    target_workspaces = []
    
    if workspace_input.lower() == 'all':
        target_workspaces = workspaces
        print(f"\nWill add user to ALL {len(workspaces)} workspaces")
    elif workspace_input.isdigit():
        idx = int(workspace_input) - 1
        if 0 <= idx < len(workspaces):
            target_workspaces = [workspaces[idx]]
        else:
            print("Invalid workspace number")
            return
    else:
        # Try to find by name or ID
        ws = find_workspace_by_name(workspaces, workspace_input)
        if ws:
            target_workspaces = [ws]
        else:
            # Try by ID
            for ws in workspaces:
                if ws['id'] == workspace_input:
                    target_workspaces = [ws]
                    break
        
        if not target_workspaces:
            print(f"No workspace found matching '{workspace_input}'")
            return
    
    # Get user email to add
    user_email = input("\nEnter user email to add as Admin: ").strip()
    
    if not user_email or '@' not in user_email:
        print("Invalid email address")
        return
    
    # Get access level
    print("\nAccess levels:")
    print("  1. Admin (full control)")
    print("  2. Member (edit + share)")
    print("  3. Contributor (edit only)")
    print("  4. Viewer (read only)")
    
    access_choice = input("Select access level [1]: ").strip() or "1"
    
    access_map = {
        "1": "Admin",
        "2": "Member", 
        "3": "Contributor",
        "4": "Viewer"
    }
    
    access_right = access_map.get(access_choice, "Admin")
    
    # Confirm
    print(f"\n{'='*60}")
    print("CONFIRMATION")
    print(f"{'='*60}")
    print(f"User: {user_email}")
    print(f"Access Level: {access_right}")
    print(f"Target Workspaces: {len(target_workspaces)}")
    for ws in target_workspaces:
        print(f"  - {ws['name']}")
    print(f"{'='*60}")
    
    confirm = input("\nProceed? (y/n): ").strip().lower()
    
    if confirm != 'y':
        print("Operation cancelled.")
        return
    
    # Add user to workspace(s)
    print(f"\nAdding user to workspace(s)...")
    print("-"*60)
    
    success_count = 0
    fail_count = 0
    
    for ws in target_workspaces:
        print(f"\nWorkspace: {ws['name']}")
        if add_user_to_workspace(access_token, ws['id'], user_email, access_right):
            success_count += 1
        else:
            fail_count += 1
    
    # Summary
    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    print(f"  Successful: {success_count}")
    print(f"  Failed: {fail_count}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
