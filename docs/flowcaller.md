# Flow Caller Documentation

A JavaScript library for calling Power Automate flows from Dynamics 365 ribbon/command bar buttons.

## Overview

Flow Caller (`flowCaller.js`) provides a simple way to trigger Power Automate HTTP-triggered flows from Dynamics 365 model-driven apps. It handles:

- Confirmation dialogs before execution
- Single record (form) and multiple record (grid) scenarios
- Progress indicators during flow execution
- Localized messages (English, Danish, German)
- Environment variable-based configuration
- Error handling and user feedback

## Installation

1. **Upload the Web Resource**
   - Navigate to your solution in Power Apps
   - Add a new Web Resource of type JavaScript (JS)
   - Upload `flowCaller.js`
   - Set the name (e.g., `new_flowcaller.js`)
   - Publish the web resource

2. **Create Environment Variable**
   - In your solution, add an Environment Variable
   - Set the Schema Name (e.g., `new_MyFlowUrl`)
   - Set the Data Type to "Text"
   - Set the Default Value to your Power Automate flow's HTTP trigger URL
   - Alternatively, set the Current Value for environment-specific URLs

3. **Create the Power Automate Flow**
   - Create a flow with "When an HTTP request is received" trigger
   - Configure the trigger to accept POST requests
   - Expected JSON payload:
     ```json
     {
       "id": "record-guid-without-braces",
       "entityName": "account"
     }
     ```

## Usage

### Method 1: Command Designer (Recommended)

In the modern Command Designer:

1. Create a new command
2. Add action: "Run JavaScript"
3. Library: Select your uploaded web resource
4. Function: `Abakion.FlowCaller.callFlowWithConfirmation`
5. Add parameters in order:
   - **PrimaryControl** (from dropdown)
   - **SelectedControlSelectedItemReferences** (for grids) or leave empty for forms
   - **String**: Environment variable schema name (e.g., `"new_MyFlowUrl"`)
   - **String**: Confirmation message (e.g., `"Process {0} record(s)?"`)
   - **String**: Success message (e.g., `"Records processed successfully"`)
   - **Boolean**: Refresh timeline after completion (`true` or `false`)

### Method 2: Ribbon Workbench

Configure the command with these parameters:

| Parameter | Type | Value |
|-----------|------|-------|
| PrimaryControl | CrmParameter | PrimaryControl |
| SelectedItemReferences | CrmParameter | SelectedControlSelectedItemReferences |
| EnvVarName | String | Your environment variable name |
| ConfirmationText | String | Confirmation message with {0} placeholder |
| SuccessMessage | String | Success message |
| RefreshTimeline | Boolean | true/false |

## API Reference

### `Abakion.FlowCaller.callFlowWithConfirmation`

Main method for triggering flows with confirmation dialog.

```javascript
Abakion.FlowCaller.callFlowWithConfirmation(
    primaryControl,           // Form or grid context
    selectedItemReferences,   // Array of selected records (null for forms)
    envVarName,              // Environment variable schema name
    confirmationText,        // Confirmation message ({0} = record count)
    successMessage,          // Success message ({0} = processed count)
    refreshTimeline          // Boolean - refresh timeline after completion
);
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `primaryControl` | Object | The form or grid execution context |
| `selectedItemReferences` | Array/null | Selected records from Command Designer, or null for forms |
| `envVarName` | String | Schema name of the environment variable containing the flow URL |
| `confirmationText` | String | Message shown in confirmation dialog. Use `{0}` for record count |
| `successMessage` | String | Message shown after successful execution. Use `{0}` for count |
| `refreshTimeline` | Boolean | Whether to refresh the timeline control after completion (forms only) |

### `Abakion.FlowCaller.callFlowWithCustomFields`

Sends additional form field values to the flow.

```javascript
Abakion.FlowCaller.callFlowWithCustomFields(
    primaryControl,    // Form context only
    envVarName,       // Environment variable schema name
    confirmationText, // Confirmation message
    fieldNames        // Array of field logical names to include
);
```

**Payload sent to flow:**
```json
{
    "id": "record-guid",
    "entityName": "account",
    "fieldname1": "value1",
    "fieldname2": "value2"
}
```

### `Abakion.FlowCaller.callFlow`

Simple method with default parameters (for backward compatibility).

```javascript
Abakion.FlowCaller.callFlow(primaryControl);
```

Uses environment variable `aba_DefaultFlowUrl` and default confirmation text.

### `Abakion.FlowCaller.setLanguage`

Manually set the UI language.

```javascript
Abakion.FlowCaller.setLanguage("da"); // Danish
Abakion.FlowCaller.setLanguage("de"); // German
Abakion.FlowCaller.setLanguage("en"); // English (default)
```

## Localization

The library automatically detects the user's language from:
1. Dynamics 365 user settings (LCID)
2. Browser language (fallback)

### Supported Languages

| Language | Code | LCID |
|----------|------|------|
| English | en | 1033 |
| Danish | da | 1030 |
| German | de | 1031 |

### Adding a New Language

Edit the `LOCALIZED_STRINGS` object in `flowCaller.js`:

```javascript
"fr": {
    noRecordsSelected: "Veuillez sélectionner au moins un enregistrement.",
    // ... add all other strings
}
```

Then add the LCID mapping in `initializeLanguage`:

```javascript
var lcidMap = {
    1033: "en",
    1030: "da",
    1031: "de",
    1036: "fr", // French
};
```

## Environment Variables

Flow URLs are retrieved from Dataverse environment variables. This allows:

- Different URLs per environment (dev, test, prod)
- Easy URL updates without code changes
- Secure storage of flow URLs

The library queries both:
1. `environmentvariablevalue` (environment-specific value)
2. `environmentvariabledefinition.defaultvalue` (fallback)

Priority is given to the environment-specific value.

## Examples

### Grid Command - Process Multiple Records

```javascript
// Command Designer configuration
Function: Abakion.FlowCaller.callFlowWithConfirmation
Parameters:
  1. PrimaryControl
  2. SelectedControlSelectedItemReferences
  3. "new_ProcessRecordsFlowUrl"
  4. "Process {0} selected record(s)?"
  5. "{0} record(s) processed successfully"
  6. false
```

### Form Command - Send Email with Timeline Refresh

```javascript
// Command Designer configuration
Function: Abakion.FlowCaller.callFlowWithConfirmation
Parameters:
  1. PrimaryControl
  2. (empty/null)
  3. "new_SendEmailFlowUrl"
  4. "Send email for this record?"
  5. "Email sent successfully"
  6. true
```

### Form Command - Include Custom Fields

```javascript
// Call from a custom script
Abakion.FlowCaller.callFlowWithCustomFields(
    formContext,
    "new_CreateDocumentFlowUrl",
    "Create document with current values?",
    ["name", "emailaddress1", "telephone1"]
);
```

## Error Handling

The library provides user-friendly error messages for:

- No records selected
- Environment variable not found
- Flow URL not configured
- HTTP errors from flow execution
- Network errors

All errors are displayed using `Xrm.Navigation.openAlertDialog`.

## Troubleshooting

### Flow URL Not Found

1. Verify the environment variable schema name matches exactly
2. Check that the environment variable has a value set
3. Ensure the user has read access to environment variables

### Flow Returns Error

1. Check the flow's run history in Power Automate
2. Verify the flow accepts the expected JSON payload
3. Ensure the flow's HTTP trigger allows anonymous access or has proper authentication

### Records Not Processing

1. Check browser console for JavaScript errors
2. Verify the web resource is published
3. Ensure the command is properly configured with all parameters

## Security Considerations

- Flow URLs should use HTTPS
- Consider using Azure AD authentication on your flows
- Environment variables are readable by users with appropriate privileges
- Flow execution is subject to Power Automate licensing and throttling
