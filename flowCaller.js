/**
 * Abakion Flow Caller
 * Used to call Power Automate flows from Dynamics 365 ribbon/command bar
 * Supports: 
 * - Single and multiple records
 * - Confirmation dialog
 * - Environment variables for flow URLs and strings
 * - Progress indicators
 */

var Abakion = Abakion || {};

Abakion.FlowCaller = {
    // Cache for loaded strings to avoid repeated API calls
    _stringsCache: null,
    
    // Fallback strings if environment variable fails to load
    _defaultStrings: {
        noRecordsSelected: "Please select at least one record.",
        envVarNotFound: "Environment variable not found: {0}",
        envVarRetrievalFailed: "Environment variable retrieval failed: {0}",
        processing: "Processing {0} record(s)",
		completed: "Completed",
        defaultSuccessMessage: "{0} record(s) processed successfully",
        recordsFailed: "{0} record(s) failed",
        confirmTitle: "Confirm Action",
        confirmButton: "Yes",
        cancelButton: "Cancel",
		flowUrlNotFound: "Flow URL not found: {0}",
        flowExecutionFailed: "Flow execution failed for record {0}: HTTP {1}",
        networkError: "Network error for record {0}",
        timelineRefreshWarning: "Timeline refresh failed",
		notSupported: "Not supported",
		onlySingleRecord: "This function only support single records from form.",
		configurationError: "Configuration Error",
		success: "Success",
		error: "error",
		errorWhenCallingFlow: "Error when calling flow: {0}",
		errorWhenRetrievingVariable: "Error when retrieving environment variable: {0}",
		flowCompletedSuccessfully: "Flow completed successfully!",
		flowStartedSuccessfully: "Flow started successfully!",
		sendingDataToFlow: "Sending data to flow..."
    },
    
    /**
     * Load strings from environment variable
     * Caches result to avoid repeated API calls
     */
    _loadStrings: function() {
        var self = this;
        
        // Return cached strings if already loaded
        if (self._stringsCache) {
            return Promise.resolve(self._stringsCache);
        }
        
        return self._getEnvironmentVariable("rup_JavaScriptFlowCallerStrings")
            .then(function(jsonString) {
                try {
                    self._stringsCache = JSON.parse(jsonString);
                    return self._stringsCache;
                } catch (error) {
                    console.warn("Failed to parse strings from environment variable, using defaults:", error);
                    self._stringsCache = self._defaultStrings;
                    return self._defaultStrings;
                }
            })
            .catch(function(error) {
                console.warn("Failed to load strings from environment variable, using defaults:", error);
                self._stringsCache = self._defaultStrings;
                return self._defaultStrings;
            });
    },
    
    /**
     * Get localized string with optional formatting
     */
    _getString: function(key, values) {
        var self = this;
        var strings = self._stringsCache || self._defaultStrings;
        var template = strings[key] || key;
        
        if (values) {
            return self._formatString(template, values);
        }
        return template;
    },
    
    /**
     * Format string with placeholders {0}, {1}, etc.
     */
    _formatString: function(template, values) {
        if (!Array.isArray(values)) {
            values = [values];
        }
        return template.replace(/\{(\d+)\}/g, function(match, index) {
            return typeof values[index] !== 'undefined' ? values[index] : match;
        });
    },
    /**
     * Calls flow with confirmation dialog - supports multiple records
     * @param {object} primaryControl - Form or grid context
     * @param {array} selectedItemReferences - Array of selected records (from Command Designer)
     * @param {string} envVarName - Name of environment variable with flow URL
     * @param {string} confirmationText - Text for confirmation (use {0} for number of records)
     * @param {string} successMessage - Success message (use {0} for number of records). E.g. "Email(s) created"
     * @param {boolean} refreshTimeline - Should timeline be refreshed after flow? (only for single record)
     */
    callFlowWithConfirmation: function(primaryControl, selectedItemReferences, envVarName, confirmationText, successMessage, refreshTimeline) {
        var self = this;
		debugger;
        
        // Load strings first
        self._loadStrings().then(function() {
            var selectedRecords = [];
            var entityName = "";
            
            // Check if we have selectedItemReferences (Command Designer)
            if (selectedItemReferences && selectedItemReferences.length > 0) {
                // Command Designer - use selectedItemReferences
                selectedRecords = selectedItemReferences;
                entityName = selectedItemReferences[0].TypeName || selectedItemReferences[0].etn || selectedItemReferences[0].entityType;
                
            } else if (primaryControl.getGrid) {
                // Ribbon Workbench - Grid context (view)
                var grid = primaryControl.getGrid();
                var rows = grid.getSelectedRows();
                
                if (rows.getLength() === 0) {
                    Xrm.Navigation.openAlertDialog({
                        text: self._getString("noRecordsSelected")
                    });
                    return;
                }
                
                var allRows = rows.getAll();
                for (var i = 0; i < allRows.length; i++) {
                    selectedRecords.push({
                        Id: allRows[i].getData().entity.getId(),
                        TypeName: allRows[i].getData().entity.getEntityName()
                    });
                }
                entityName = allRows[0].getData().entity.getEntityName();
                
            } else if (primaryControl.data && primaryControl.data.entity) {
                // Form context (single record)
                var recordId = primaryControl.data.entity.getId().replace(/[{}]/g, "");
                entityName = primaryControl.data.entity.getEntityName();
                
                selectedRecords = [{
                    Id: recordId,
                    TypeName: entityName
                }];
            } else {
                Xrm.Navigation.openAlertDialog({
                    text: self._getString("noRecordsSelected")
                });
                return;
            }
            
            if (selectedRecords.length === 0) {
                Xrm.Navigation.openAlertDialog({
                    text: self._getString("noRecordsSelected")
                });
                return;
            }
            
            // Build confirmation text with count
            var confirmText = confirmationText.replace("{0}", selectedRecords.length);
            
            // Show confirmation dialog
            var confirmStrings = {
                text: confirmText,
                title: self._getString("confirmTitle"),
                confirmButtonLabel: self._getString("confirmButton"),
                cancelButtonLabel: self._getString("cancelButton")
            };
            
            Xrm.Navigation.openConfirmDialog(confirmStrings).then(function(result) {
                if (result.confirmed) {
                    // Get flow URL from environment variable
                    self._getEnvironmentVariable(envVarName).then(function(flowUrl) {
                        if (!flowUrl) {
                            Xrm.Navigation.openAlertDialog({
                                text: self._getString("envVarNotFound", envVarName)
                            });
                            return;
                        }
                        
                        // Process records
                        self._processRecords(selectedRecords, entityName, flowUrl, primaryControl, successMessage, refreshTimeline);
                        
                    }).catch(function(error) {
                        Xrm.Navigation.openAlertDialog({
                            text: self._getString("envVarRetrievalFailed", error.message)
                        });
                    });
                }
            });
        });
    },
    
    /**
     * Simple version without confirmation (backward compatibility)
     * @param {object} primaryControl - Form context
     */
    callFlow: function(primaryControl) {
        this.callFlowWithConfirmation(primaryControl, null, "aba_DefaultFlowUrl", "Do you want to run this action?", null, false);
    },
    
    /**
     * Version with custom data from form fields
     * @param {object} primaryControl - Form context
     * @param {string} envVarName - Environment variable name
     * @param {string} confirmationText - Confirmation text
     * @param {array} fieldNames - Array of field names to include in payload
     */
    callFlowWithCustomFields: function(primaryControl, envVarName, confirmationText, fieldNames) {
        var self = this;
        var formContext = primaryControl;
        
        // Validation - only form context
        if (primaryControl.getGrid) {
            Xrm.Navigation.openAlertDialog({
                text: self._getString("onlySingleRecord"),
                title: self._getString("notSupported")
            });
            return;
        }
        
        var recordId = formContext.data.entity.getId().replace(/[{}]/g, "");
        var entityName = formContext.data.entity.getEntityName();
        
        // Show confirmation
        var confirmStrings = {
            text: confirmationText.replace("{0}", "1"),
			title: self._getString("confirmTitle"),
			confirmButtonLabel: self._getString("confirmButton"),
			cancelButtonLabel: self._getString("cancelButton")
        };
        
        Xrm.Navigation.openConfirmDialog(confirmStrings).then(function(result) {
            if (result.confirmed) {
                self._getEnvironmentVariable(envVarName).then(function(flowUrl) {
                    if (!flowUrl) {
                        Xrm.Navigation.openAlertDialog({
                            text: self._getString("flowUrlNotFound", envVarName),
                            title: self._getString("configurationError")
                        });
                        return;
                    }
                    
                    // Build custom payload
                    var payload = {
                        id: recordId,
                        entityName: entityName
                    };
                    
                    // Add custom fields
                    if (fieldNames && fieldNames.length > 0) {
                        for (var i = 0; i < fieldNames.length; i++) {
                            var fieldName = fieldNames[i];
                            var attribute = formContext.getAttribute(fieldName);
                            if (attribute) {
                                payload[fieldName] = attribute.getValue();
                            }
                        }
                    }
                    
                    Xrm.Utility.showProgressIndicator(self._getString("sendingDataToFlow"));
                    
                    fetch(flowUrl, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify(payload)
                    })
                    .then(function(response) {
                        Xrm.Utility.closeProgressIndicator();
                        
                        if (response.ok) {
                            return response.json().then(function(data) {
                                Xrm.Navigation.openAlertDialog({
                                    text: self._getString("flowCompletedSuccessfully"),
                                    title: self._getString("success")
                                }).then(function() {
                                    formContext.data.refresh(false);
                                });
                            }).catch(function() {
                                Xrm.Navigation.openAlertDialog({
                                    text: self._getString("flowStartedSuccessfully"),
                                    title: self._getString("success")
                                });
                            });
                        } else {
                            throw new Error('HTTP ' + response.status);
                        }
                    })
                    .catch(function(error) {
                        Xrm.Utility.closeProgressIndicator();
                        Xrm.Navigation.openAlertDialog({
                            text: self._getString("errorWhenCallingFlow", error.message),
                            title: self._getString("error")
                        });
                        console.error("Abakion Flow call error:", error);
                    });
                    
                }).catch(function(error) {
                    Xrm.Navigation.openAlertDialog({
                        text: self._getString("errorWhenRetrievingVariable", error.message),
                        title: self._getString("error")
                    });
                });
            }
        });
    },
    
    /**
     * PRIVATE: Get environment variable value
     */
    _getEnvironmentVariable: function(variableName) {
        return new Promise(function(resolve, reject) {
            var fetchXml = [
                "<fetch top='1'>",
                "  <entity name='environmentvariabledefinition'>",
                "    <attribute name='defaultvalue' />",
                "    <filter>",
                "      <condition attribute='schemaname' operator='eq' value='" + variableName + "' />",
                "    </filter>",
                "    <link-entity name='environmentvariablevalue' from='environmentvariabledefinitionid' to='environmentvariabledefinitionid' link-type='outer'>",
                "      <attribute name='value' />",
                "    </link-entity>",
                "  </entity>",
                "</fetch>"
            ].join("");
            
            Xrm.WebApi.retrieveMultipleRecords("environmentvariabledefinition", "?fetchXml=" + encodeURIComponent(fetchXml))
                .then(function(result) {
                    if (result.entities.length > 0) {
                        // Prioritize value over defaultvalue
                        var value = result.entities[0]["environmentvariablevalue1.value"] || 
                                   result.entities[0].defaultvalue;
                        resolve(value);
                    } else {
                        reject(new Error("Environment variable not found: " + variableName));
                    }
                })
                .catch(function(error) {
                    reject(error);
                });
        });
    },
    
    /**
     * PRIVATE: Process records and call flow
     */
    _processRecords: function(selectedRecords, entityName, flowUrl, primaryControl, successMessage, refreshTimeline) {
        var self = this;
        var totalRecords = selectedRecords.length;
        var processedCount = 0;
        var failedCount = 0;
        var isFormContext = !primaryControl.getGrid;
        
        Xrm.Utility.showProgressIndicator(self._getString("processing", totalRecords));
        
        // Process each record
        var promises = selectedRecords.map(function(record) {
            // Handle different record formats
            var recordId = record.Id || record.id || record.getId();
            
            // Remove curly braces if present
            if (recordId) {
                recordId = recordId.toString().replace(/[{}]/g, "");
            }
            
            var payload = {
                id: recordId,
                entityName: entityName
            };
            
            return fetch(flowUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            })
            .then(function(response) {
                if (response.ok) {
                    processedCount++;
                } else {
                    failedCount++;
                    console.error(self._getString("flowExecutionFailed", [recordId, response.status]));
                }
            })
            .catch(function(error) {
                failedCount++;
                console.error(self._getString("networkError", recordId), error);
            });
        });
        
        // Wait for all requests
        Promise.all(promises).then(function() {
            Xrm.Utility.closeProgressIndicator();
            
            // Build success message
            var message;
            if (successMessage) {
                // Use custom success message
                message = successMessage.replace("{0}", processedCount);
            } else {
                // Fallback to standard message
                message = self._getString("defaultSuccessMessage", processedCount);
            }
            
            if (failedCount > 0) {
                message += "\n" + self._getString("recordsFailed", failedCount);
            }
            
            Xrm.Navigation.openAlertDialog({
                text: message,
                title: self._getString("completed")
            }).then(function() {
                // Refresh timeline if it's form context and refreshTimeline is true
                if (isFormContext && refreshTimeline === true) {
                    try {
                        // Modern way to refresh timeline (replaces Xrm.Page)
                        var timelineControl = primaryControl.getControl("Timeline");
                        if (timelineControl && timelineControl.refresh) {
                            timelineControl.refresh();
                        }
                    } catch (error) {
                        console.warn(self._getString("timelineRefreshWarning"), error);
                    }
                }
                
                // Refresh grid if it's grid context
                if (primaryControl.getGrid) {
                    primaryControl.getGrid().refresh();
                } else {
                    // Refresh form
                    primaryControl.data.refresh(false);
                }
            });
        });
    }
};

/**
 * USAGE EXAMPLES:
 * 
 * COMMAND DESIGNER - GRID (Multiple records):
 * ============================================
 * Function: Abakion.FlowCaller.callFlowWithConfirmation
 * Parameters (in this order):
 * 1. Primary Control (from dropdown)
 * 2. Selected Items (from dropdown - select "SelectedControlSelectedItemReferences")
 * 3. "environmentVariable_with_flow_link" (String - environment variable name)
 * 4. "Do you want to update these {0} records?" (String - confirmation text)
 * 5. "Email(s) created" (String - success message, use {0} for count)
 * 6. false (Boolean - refresh timeline? false for grid)
 * 
 * COMMAND DESIGNER - FORM (Single record with timeline refresh):
 * ==============================================================
 * Function: Abakion.FlowCaller.callFlowWithConfirmation
 * Parameters:
 * 1. Primary Control (from dropdown)
 * 2. null (or leave empty - not relevant for form)
 * 3. "environmentVariable_with_flow_link" (String)
 * 4. "Do you want to send this email?" (String)
 * 5. "Email created" (String - do NOT use {0} for single record)
 * 6. true (Boolean - refresh timeline)
 * 
 * SUCCESS MESSAGE EXAMPLES:
 * ===============================
 * Grid (multiple): "Email(s) created" → "3 Email(s) created"
 * Grid (multiple): "{0} opportunity(ies) updated" → "5 opportunity(ies) updated"
 * Form (single): "Email created" → "Email created" (no {0})
 * Form (single): "Record updated successfully" → "Record updated successfully"
 * 
 * WITH CUSTOM FIELDS:
 * ==================
 * Abakion.FlowCaller.callFlowWithCustomFields(
 *     primaryControl,
 *     "Send order for approval?",
 *     "aba_OrderApprovalFlowUrl",
 *     ["aba_amount", "aba_customer", "aba_priority"]
 * );
 */
