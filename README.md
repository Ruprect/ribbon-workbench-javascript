# Dynamics 365 / Power Platform Tools

A collection of tools for Microsoft Dynamics 365 and Power Platform development, created by Abakion A/S.

## Tools

### [Flow Caller (flowCaller.js)](docs/flowcaller.md)

A JavaScript library for triggering Power Automate flows directly from Dynamics 365 ribbon/command bar buttons. Designed as a modern replacement for Ribbon Workbench/Smart Buttons functionality.

**Key Features:**
- Call Power Automate flows from ribbon buttons
- Support for single records (forms) and multiple records (grids)
- Confirmation dialogs before execution
- Multi-language support (English, Danish, German)
- Environment variable-based flow URL configuration
- Progress indicators and error handling

### [Solution Checker (dataverse_solution_checker.py)](docs/solution-checker.md)

A Python utility to scan all Power Platform environments you have access to, checking for:
1. Presence of a specific solution (e.g., Develop1SmartButtons)
2. Unmanaged, non-Microsoft JavaScript web resources in the Default solution

**Key Features:**
- Device code authentication (no app registration required)
- Persistent token caching between runs
- Scans all environments automatically
- Exports findings to CSV files per environment
- Filters out Microsoft and managed web resources

## Quick Start

### Flow Caller

1. Upload `flowCaller.js` as a web resource in your Dynamics 365 solution
2. Create an environment variable containing your Power Automate flow HTTP trigger URL
3. Add a command bar button calling `Abakion.FlowCaller.callFlowWithConfirmation`

See [full documentation](docs/flowcaller.md) for detailed setup instructions.

### Solution Checker

```bash
# Install dependencies
pip install msal requests

# Run the checker
python dataverse_solution_checker.py
```

See [full documentation](docs/solution-checker.md) for configuration options.

## License

See [LICENSE](LICENSE) file.
