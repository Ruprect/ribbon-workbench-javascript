"""
Dataverse Solution Checker
==========================
Checks all Power Platform environments for:
1. A specific solution (Develop1SmartButtons)
2. Non-Microsoft web resources in the Default solution

Requirements:
    pip install msal requests

Author: Abakion A/S
"""

import csv
import msal
import os
import requests
import sys
from datetime import datetime
from typing import Optional

# Configuration
SOLUTION_TO_FIND = "Develop1SmartButtons"

# Microsoft publisher prefixes to exclude when checking web resources
MICROSOFT_PREFIXES = (
    "msdyn_",
    "mscrm_",
    "Microsoft",
    "cc_",        # Common Microsoft component prefix
    "msdynce_",
    "msdynmkt_",
    "msfp_",
    "mspcat_",
    "adminsi_",
    "powerpagesite_",
    "powerpagecomponent_",
)

# Azure AD App Registration (use the well-known Power Platform CLI client ID)
CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d"  # Power Platform CLI
AUTHORITY = "https://login.microsoftonline.com/organizations"

# Token cache file path (in user's home directory)
TOKEN_CACHE_FILE = os.path.join(os.path.expanduser("~"), ".dataverse_checker_cache.json")

# API Scopes
BAP_SCOPE = ["https://api.bap.microsoft.com/.default"]
DATAVERSE_SCOPE_TEMPLATE = "https://{}.crm4.dynamics.com/.default"


def load_token_cache() -> msal.SerializableTokenCache:
    """
    Load token cache from file if it exists.
    """
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        try:
            with open(TOKEN_CACHE_FILE, "r") as f:
                cache.deserialize(f.read())
            print(f"✅ Loaded cached credentials from {TOKEN_CACHE_FILE}")
        except Exception as e:
            print(f"⚠️  Could not load token cache: {e}")
    return cache


def save_token_cache(cache: msal.SerializableTokenCache) -> None:
    """
    Save token cache to file if it has changed.
    """
    if cache.has_state_changed:
        try:
            with open(TOKEN_CACHE_FILE, "w") as f:
                f.write(cache.serialize())
        except Exception as e:
            print(f"⚠️  Could not save token cache: {e}")


def get_token_with_device_code(scopes: list[str], cache: msal.SerializableTokenCache, silent_only: bool = False) -> Optional[str]:
    """
    Authenticate using device code flow.
    Returns access token or None if failed.

    If silent_only=True, only try cached tokens without prompting for login.
    """
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )

    # Try to get token from cache first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            print(f"   ✅ Using cached token for {accounts[0].get('username', 'unknown user')}")
            return result["access_token"]

    # If silent only, don't prompt for login
    if silent_only:
        return None

    # Initiate device code flow
    flow = app.initiate_device_flow(scopes=scopes)
    
    if "user_code" not in flow:
        print(f"Failed to create device flow: {flow.get('error_description', 'Unknown error')}")
        return None
    
    print("\n" + "=" * 60)
    print("DEVICE CODE AUTHENTICATION")
    print("=" * 60)
    print(f"\n{flow['message']}\n")
    print("=" * 60 + "\n")
    
    # Wait for user to authenticate
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        print("✅ Authentication successful!\n")
        return result["access_token"]
    else:
        print(f"❌ Authentication failed: {result.get('error_description', 'Unknown error')}")
        return None


def get_environments(token: str) -> list[dict]:
    """
    Get all Power Platform environments the user has access to.
    Uses the BAP API.
    """
    url = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments"
    params = {
        "api-version": "2020-10-01"
    }
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }
    
    response = requests.get(url, headers=headers, params=params)
    
    if response.status_code == 200:
        return response.json().get("value", [])
    elif response.status_code == 403:
        # Try non-admin endpoint
        url = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments"
        response = requests.get(url, headers=headers, params=params)
        if response.status_code == 200:
            return response.json().get("value", [])
    
    print(f"❌ Failed to get environments: {response.status_code}")
    print(response.text)
    return []


def get_dataverse_token(env_url: str, cache: msal.SerializableTokenCache) -> Optional[str]:
    """
    Get a token for a specific Dataverse environment.
    Returns token and whether it was from cache.
    """
    # Extract the base URL without protocol
    base_url = env_url.replace("https://", "").replace("http://", "").rstrip("/")
    scope = [f"https://{base_url}/.default"]

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )

    # Try to get token from cache
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scope, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    # Need interactive auth for this scope
    flow = app.initiate_device_flow(scopes=scope)
    if "user_code" not in flow:
        return None

    print(f"\n⚠️  Additional authentication needed for: {base_url}")
    print(f"   {flow['message']}\n")

    result = app.acquire_token_by_device_flow(flow)

    # Save cache after new authentication
    save_token_cache(cache)

    return result.get("access_token")


def check_solution_in_environment(env_url: str, token: str, solution_name: str) -> Optional[dict]:
    """
    Check if a solution exists in a Dataverse environment.
    Returns solution info if found, None otherwise.
    """
    # Ensure URL ends with /
    if not env_url.endswith("/"):
        env_url += "/"
    
    # Query the solutions table
    url = f"{env_url}api/data/v9.2/solutions"
    params = {
        "$select": "uniquename,friendlyname,version,installedon,ismanaged",
        "$filter": f"uniquename eq '{solution_name}'"
    }
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0"
    }
    
    try:
        response = requests.get(url, headers=headers, params=params, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            solutions = data.get("value", [])
            if solutions:
                return solutions[0]
        elif response.status_code == 401:
            return {"error": "Unauthorized - token may have expired"}
        elif response.status_code == 403:
            return {"error": "Access denied - insufficient permissions"}
        else:
            return {"error": f"HTTP {response.status_code}"}
    except requests.exceptions.Timeout:
        return {"error": "Request timeout"}
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}

    return None


def get_non_microsoft_js_webresources(env_url: str, token: str) -> dict:
    """
    Get all non-Microsoft, unmanaged JavaScript web resources from the Default solution.
    Returns dict with 'webresources' list or 'error' key.
    """
    if not env_url.endswith("/"):
        env_url += "/"

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0"
    }

    # First, get the Default solution ID
    solutions_url = f"{env_url}api/data/v9.2/solutions"
    params = {
        "$select": "solutionid,uniquename",
        "$filter": "uniquename eq 'Default'"
    }

    try:
        response = requests.get(solutions_url, headers=headers, params=params, timeout=30)

        if response.status_code != 200:
            return {"error": f"Failed to get Default solution: HTTP {response.status_code}"}

        solutions = response.json().get("value", [])
        if not solutions:
            return {"error": "Default solution not found"}

        default_solution_id = solutions[0]["solutionid"]

        # Query web resources in the Default solution via solutioncomponent
        # Component type 61 = Web Resource
        components_url = f"{env_url}api/data/v9.2/solutioncomponents"
        params = {
            "$select": "objectid",
            "$filter": f"_solutionid_value eq {default_solution_id} and componenttype eq 61"
        }

        response = requests.get(components_url, headers=headers, params=params, timeout=30)

        if response.status_code != 200:
            return {"error": f"Failed to get solution components: HTTP {response.status_code}"}

        components = response.json().get("value", [])

        if not components:
            return {"webresources": []}

        # Get the web resource IDs
        webresource_ids = [c["objectid"] for c in components]

        # Query web resources by IDs (batch in groups to avoid URL length limits)
        all_webresources = []
        batch_size = 50

        for i in range(0, len(webresource_ids), batch_size):
            batch_ids = webresource_ids[i:i + batch_size]
            filter_conditions = " or ".join([f"webresourceid eq {wrid}" for wrid in batch_ids])

            wr_url = f"{env_url}api/data/v9.2/webresourceset"
            params = {
                "$select": "name,displayname,webresourcetype,ismanaged,createdon,modifiedon",
                "$filter": f"({filter_conditions}) and webresourcetype eq 3"  # 3 = JavaScript
            }

            response = requests.get(wr_url, headers=headers, params=params, timeout=30)

            if response.status_code == 200:
                all_webresources.extend(response.json().get("value", []))

        # Filter out Microsoft and managed web resources
        non_microsoft_unmanaged = []
        for wr in all_webresources:
            name = wr.get("name", "")
            is_managed = wr.get("ismanaged", False)
            if not is_managed and not any(name.startswith(prefix) for prefix in MICROSOFT_PREFIXES):
                non_microsoft_unmanaged.append(wr)

        return {"webresources": non_microsoft_unmanaged}

    except requests.exceptions.Timeout:
        return {"error": "Request timeout"}
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}


def write_webresources_csv(env_name: str, webresources: list, output_dir: str = ".") -> str:
    """
    Write web resources to a CSV file for an environment.
    Returns the file path.
    """
    # Sanitize environment name for filename
    safe_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in env_name)
    safe_name = safe_name.strip().replace(' ', '_')

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"webresources_{safe_name}_{timestamp}.csv"
    filepath = os.path.join(output_dir, filename)

    with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['name', 'displayname', 'ismanaged', 'createdon', 'modifiedon']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        writer.writeheader()
        for wr in webresources:
            writer.writerow({
                'name': wr.get('name', ''),
                'displayname': wr.get('displayname', ''),
                'ismanaged': wr.get('ismanaged', ''),
                'createdon': wr.get('createdon', ''),
                'modifiedon': wr.get('modifiedon', '')
            })

    return filepath


def main():
    print("\n" + "=" * 60)
    print("DATAVERSE SOLUTION CHECKER")
    print(f"Looking for: {SOLUTION_TO_FIND}")
    print("Also checking: Unmanaged, non-Microsoft JavaScript (.js) in Default solution")
    print("=" * 60)

    # Track CSV files created
    csv_files_created = []
    
    # Load token cache from disk (if exists)
    cache = load_token_cache()

    # Step 1: Authenticate to BAP API
    print("\n📝 Step 1: Authenticating to Power Platform...")
    bap_token = get_token_with_device_code(BAP_SCOPE, cache)

    # Save cache after authentication
    save_token_cache(cache)
    
    if not bap_token:
        print("❌ Could not authenticate. Exiting.")
        sys.exit(1)
    
    # Step 2: Get all environments
    print("📝 Step 2: Fetching environments...")
    environments = get_environments(bap_token)
    
    if not environments:
        print("❌ No environments found or access denied.")
        sys.exit(1)
    
    print(f"\n✅ Found {len(environments)} environment(s)\n")
    
    # Step 3: Check each environment for the solution
    print("=" * 60)
    print("CHECKING ENVIRONMENTS")
    print("=" * 60)
    
    results = []
    dataverse_envs = []
    
    # First, identify environments with Dataverse
    for env in environments:
        props = env.get("properties", {})
        display_name = props.get("displayName", "Unknown")
        env_type = props.get("environmentSku", "Unknown")
        linked_env = props.get("linkedEnvironmentMetadata", {})
        instance_url = linked_env.get("instanceUrl", "")
        
        if instance_url:
            dataverse_envs.append({
                "name": display_name,
                "type": env_type,
                "url": instance_url,
                "id": env.get("name", "")
            })
        else:
            print(f"⏭️  {display_name} ({env_type}) - No Dataverse database")
    
    if not dataverse_envs:
        print("\n❌ No environments with Dataverse found.")
        sys.exit(1)
    
    print(f"\n📊 {len(dataverse_envs)} environment(s) with Dataverse to check\n")
    
    # Check each Dataverse environment
    for i, env in enumerate(dataverse_envs, 1):
        print(f"\n[{i}/{len(dataverse_envs)}] {env['name']} ({env['type']})")
        print(f"    URL: {env['url']}")
        
        # Get token for this environment
        dv_token = get_dataverse_token(env["url"], cache)
        
        if not dv_token:
            print(f"    ❌ Could not get token for this environment")
            results.append({**env, "status": "Auth Failed", "solution": None})
            continue
        
        # Check for solution
        solution = check_solution_in_environment(env["url"], dv_token, SOLUTION_TO_FIND)

        if solution:
            if "error" in solution:
                print(f"    ⚠️  Solution check error: {solution['error']}")
                results.append({**env, "status": "Error", "solution": None, "error": solution["error"]})
            else:
                version = solution.get("version", "Unknown")
                is_managed = "Managed" if solution.get("ismanaged") else "Unmanaged"
                friendly = solution.get("friendlyname", "")
                print(f"    ✅ FOUND: {friendly} v{version} ({is_managed})")
                results.append({**env, "status": "Found", "solution": solution})
        else:
            print(f"    ❌ Solution not found")
            results.append({**env, "status": "Not Found", "solution": None})

        # Check for non-Microsoft JavaScript web resources in Default solution
        wr_result = get_non_microsoft_js_webresources(env["url"], dv_token)

        if "error" in wr_result:
            print(f"    ⚠️  JS web resources check error: {wr_result['error']}")
            results[-1]["webresources"] = None
            results[-1]["webresources_error"] = wr_result["error"]
        else:
            webresources = wr_result["webresources"]
            results[-1]["webresources"] = webresources
            if webresources:
                print(f"    📦 Found {len(webresources)} non-Microsoft JS file(s) in Default solution:")
                for wr in webresources[:5]:  # Show first 5
                    wr_name = wr.get("name", "Unknown")
                    print(f"       • {wr_name}")
                if len(webresources) > 5:
                    print(f"       ... and {len(webresources) - 5} more")

                # Write to CSV
                csv_path = write_webresources_csv(env["name"], webresources)
                csv_files_created.append({"env": env["name"], "path": csv_path, "count": len(webresources)})
                print(f"    💾 Saved to: {csv_path}")
            else:
                print(f"    ✅ No non-Microsoft JS files in Default solution")
    
    # Summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)

    # Solution summary
    print(f"\n--- {SOLUTION_TO_FIND} Solution ---")
    found_envs = [r for r in results if r["status"] == "Found"]
    not_found_envs = [r for r in results if r["status"] == "Not Found"]
    error_envs = [r for r in results if r["status"] in ["Error", "Auth Failed"]]

    print(f"\n✅ Found in {len(found_envs)} environment(s):")
    for env in found_envs:
        sol = env["solution"]
        print(f"   • {env['name']}: v{sol.get('version', '?')} ({'Managed' if sol.get('ismanaged') else 'Unmanaged'})")

    if not_found_envs:
        print(f"\n❌ Not found in {len(not_found_envs)} environment(s):")
        for env in not_found_envs:
            print(f"   • {env['name']}")

    if error_envs:
        print(f"\n⚠️  Errors in {len(error_envs)} environment(s):")
        for env in error_envs:
            print(f"   • {env['name']}: {env.get('error', env['status'])}")

    # Web resources summary
    print(f"\n--- Unmanaged, Non-Microsoft JavaScript Files in Default Solution ---")
    envs_with_webresources = [r for r in results if r.get("webresources")]
    envs_clean = [r for r in results if r.get("webresources") is not None and len(r.get("webresources", [])) == 0]

    if envs_with_webresources:
        print(f"\n📦 Found non-Microsoft JS files in {len(envs_with_webresources)} environment(s):")
        for env in envs_with_webresources:
            wr_list = env["webresources"]
            print(f"\n   {env['name']} ({len(wr_list)} JS file(s)):")
            for wr in wr_list:
                wr_name = wr.get("name", "Unknown")
                wr_display = wr.get("displayname", "")
                display_info = f" ({wr_display})" if wr_display else ""
                print(f"      • {wr_name}{display_info}")
    else:
        print(f"\n✅ No non-Microsoft JS files found in Default solution")

    if envs_clean:
        print(f"\n✅ Clean environments (no non-Microsoft JS files): {len(envs_clean)}")
        for env in envs_clean:
            print(f"   • {env['name']}")

    # CSV files summary
    if csv_files_created:
        print(f"\n--- CSV Files Created ---")
        for csv_info in csv_files_created:
            print(f"   • {csv_info['env']}: {csv_info['path']} ({csv_info['count']} records)")

    print("\n" + "=" * 60)
    print("Done!")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()