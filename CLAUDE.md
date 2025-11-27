# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A collection of tools for Microsoft Dynamics 365 / Power Platform:

1. **flowCaller.js** - JavaScript library for calling Power Automate flows from Dynamics 365 ribbon/command bar buttons. Replaces functionality previously provided by Ribbon Workbench/Smart Buttons.

2. **dataverse_solution_checker.py** - Python utility to scan all Power Platform environments for a specific solution using device code authentication.

## Architecture

### flowCaller.js

Client-side JavaScript library using the `Abakion.FlowCaller` namespace. Key concepts:

- **Environment Variables**: Flow URLs are stored in Dataverse environment variables, retrieved via FetchXML against `environmentvariabledefinition` and `environmentvariablevalue` tables
- **Localization**: Built-in support for English, Danish, and German. Auto-detects from Xrm user settings LCID or falls back to browser language
- **Context Handling**: Supports three contexts:
  - Command Designer with `selectedItemReferences` (multi-select grids)
  - Ribbon Workbench grid context via `primaryControl.getGrid()`
  - Form context via `primaryControl.data.entity`

Main entry points:
- `callFlowWithConfirmation()` - Primary method for grid/form with confirmation dialog
- `callFlowWithCustomFields()` - For forms with additional field data in payload
- `callFlow()` - Simple wrapper for backward compatibility

### dataverse_solution_checker.py

Standalone Python script using MSAL for authentication. Uses the Power Platform CLI client ID for device code flow. Queries the BAP API for environments, then checks each Dataverse instance for a specific solution.

Dependencies: `msal`, `requests`

## Dynamics 365 / Power Platform Context

This code runs in the Dynamics 365 browser environment where the `Xrm` global object is available. Key Xrm APIs used:
- `Xrm.WebApi.retrieveMultipleRecords()` - Dataverse queries
- `Xrm.Navigation.openConfirmDialog()` / `openAlertDialog()` - User dialogs
- `Xrm.Utility.showProgressIndicator()` / `closeProgressIndicator()` - Progress UI
- `Xrm.Utility.getGlobalContext().userSettings` - User locale settings
