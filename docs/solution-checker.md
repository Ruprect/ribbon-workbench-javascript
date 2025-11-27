# Solution Checker Documentation

A Python utility to scan all Power Platform environments for solutions and unmanaged web resources.

## Overview

The Solution Checker (`dataverse_solution_checker.py`) scans all Power Platform environments you have access to and reports:

1. **Solution Presence**: Whether a specific solution (default: `Develop1SmartButtons`) exists in each environment
2. **Unmanaged JavaScript**: Non-Microsoft, unmanaged JavaScript web resources in the Default solution

This helps identify:
- Where legacy solutions are deployed
- Environments with custom JavaScript that should be moved to proper solutions
- Potential cleanup candidates before migrations

## Requirements

- Python 3.8+
- Required packages:
  ```bash
  pip install msal requests
  ```

## Quick Start

```bash
# Install dependencies
pip install msal requests

# Run the checker
python dataverse_solution_checker.py
```

On first run, you'll be prompted to authenticate using device code flow:

```
============================================================
DEVICE CODE AUTHENTICATION
============================================================

To sign in, use a web browser to open https://microsoft.com/devicelogin
and enter the code XXXXXXXX to authenticate.

============================================================
```

## Authentication

### Device Code Flow

The script uses the Power Platform CLI's public client ID for authentication. This means:

- No app registration required
- Works with any Microsoft 365 account
- Respects your existing Power Platform permissions
- Supports MFA and conditional access policies

### Token Caching

Tokens are cached in your home directory:
- **Windows**: `C:\Users\<username>\.dataverse_checker_cache.json`
- **macOS/Linux**: `~/.dataverse_checker_cache.json`

On subsequent runs, cached tokens are used automatically. If tokens expire, they're refreshed silently using the refresh token.

To force re-authentication, delete the cache file.

## Configuration

Edit the constants at the top of `dataverse_solution_checker.py`:

```python
# Solution to search for
SOLUTION_TO_FIND = "Develop1SmartButtons"

# Prefixes to exclude from web resource check
MICROSOFT_PREFIXES = (
    "msdyn_",
    "mscrm_",
    "Microsoft",
    # ... add more as needed
)
```

### Microsoft Prefixes

Web resources starting with these prefixes are excluded from the report:

| Prefix | Description |
|--------|-------------|
| `msdyn_` | Dynamics 365 first-party apps |
| `mscrm_` | Legacy CRM components |
| `Microsoft` | Microsoft components |
| `cc_` | Common controls |
| `msdynce_` | Customer Engagement |
| `msdynmkt_` | Marketing |
| `msfp_` | Forms Pro |
| `mspcat_` | Power CAT tools |
| `adminsi_` | Admin components |
| `powerpagesite_` | Power Pages |
| `powerpagecomponent_` | Power Pages components |

## Output

### Console Output

The script displays:

1. **Environment List**: All environments with/without Dataverse
2. **Per-Environment Results**:
   - Solution found/not found with version
   - Unmanaged JavaScript web resources found
   - CSV file created (if web resources found)
3. **Summary**:
   - Environments with the solution
   - Environments without the solution
   - Environments with unmanaged JavaScript
   - List of CSV files created

### Example Output

```
============================================================
DATAVERSE SOLUTION CHECKER
Looking for: Develop1SmartButtons
Also checking: Unmanaged, non-Microsoft JavaScript (.js) in Default solution
============================================================

✅ Loaded cached credentials from C:\Users\user\.dataverse_checker_cache.json

📝 Step 1: Authenticating to Power Platform...
   ✅ Using cached token for user@contoso.com

📝 Step 2: Fetching environments...

✅ Found 5 environment(s)

============================================================
CHECKING ENVIRONMENTS
============================================================

⏭️  Power Pages (Developer) - No Dataverse database

📊 4 environment(s) with Dataverse to check

[1/4] Production (Production)
    URL: https://contoso.crm4.dynamics.com/
    ✅ FOUND: Smart Buttons v1.0.0.5 (Managed)
    📦 Found 3 non-Microsoft JS file(s) in Default solution:
       • new_customscript.js
       • new_legacycode.js
       • contoso_helper.js
    💾 Saved to: webresources_Production_20231127_143052.csv

[2/4] Development (Sandbox)
    URL: https://contoso-dev.crm4.dynamics.com/
    ❌ Solution not found
    ✅ No non-Microsoft JS files in Default solution
```

### CSV Output

For each environment with unmanaged JavaScript, a CSV file is created:

**Filename format**: `webresources_{EnvironmentName}_{timestamp}.csv`

**Columns**:
| Column | Description |
|--------|-------------|
| `name` | Web resource logical name |
| `displayname` | Display name |
| `ismanaged` | Always `False` (filtered) |
| `createdon` | Creation date |
| `modifiedon` | Last modified date |

## Filtering Logic

The script identifies web resources that are:

1. **In the Default solution**: Components added directly without a custom solution
2. **JavaScript type**: `webresourcetype = 3`
3. **Unmanaged**: `ismanaged = false`
4. **Non-Microsoft**: Name doesn't start with any Microsoft prefix

## API Endpoints Used

### Business Application Platform (BAP) API
- `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments`
- Retrieves list of Power Platform environments

### Dataverse Web API
- `{env}/api/data/v9.2/solutions` - Query solutions
- `{env}/api/data/v9.2/solutioncomponents` - Query solution components
- `{env}/api/data/v9.2/webresourceset` - Query web resources

## Troubleshooting

### "No environments found"

1. Verify you have Power Platform access
2. Try the non-admin endpoint (automatic fallback)
3. Check your Microsoft 365 account has appropriate licenses

### "Could not get token for this environment"

1. The environment may require additional authentication
2. You may not have access to that specific environment
3. The environment URL may be incorrect

### "Access denied - insufficient permissions"

You need at least System Customizer or equivalent role to query solutions and web resources.

### Token Cache Issues

Delete the cache file to force re-authentication:
```bash
# Windows
del %USERPROFILE%\.dataverse_checker_cache.json

# macOS/Linux
rm ~/.dataverse_checker_cache.json
```

## Extending the Script

### Check for Different Solution

```python
SOLUTION_TO_FIND = "YourSolutionUniqueName"
```

### Check for Other Web Resource Types

Modify the `webresourcetype` filter in `get_non_microsoft_js_webresources`:

| Type | Value |
|------|-------|
| HTML | 1 |
| CSS | 2 |
| JavaScript | 3 |
| XML | 4 |
| PNG | 5 |
| JPG | 6 |
| GIF | 7 |
| XAP | 8 |
| XSL | 9 |
| ICO | 10 |
| SVG | 11 |
| RESX | 12 |

### Add Custom Prefix Exclusions

```python
MICROSOFT_PREFIXES = (
    "msdyn_",
    "mscrm_",
    # ... existing prefixes
    "yourcompany_",  # Add your own managed prefixes to exclude
)
```

## Security Considerations

- The token cache file contains sensitive authentication tokens
- Add `.dataverse_checker_cache.json` to `.gitignore`
- The script uses Microsoft's official Power Platform CLI client ID
- All API calls use HTTPS
- Tokens are stored with user-only file permissions where supported
