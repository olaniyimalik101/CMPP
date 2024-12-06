function updateRelatedFields(executionContext) {
    var formContext = executionContext.getFormContext();
    
    // Check if the form type is 'Create' (form type 1)
    if (formContext.ui.getFormType() !== 1) {
        return; 
    }

    var relatedCase = formContext.getAttribute("cmpp_case").getValue();

    if (relatedCase != null && relatedCase.length > 0) {
        var caseRecordId = relatedCase[0].id; 
        var caseEntityType = relatedCase[0].entityType; 

        Xrm.WebApi.retrieveRecord(caseEntityType, caseRecordId, "?$select=cmpp_anumber,cmpp_cellphone,cmpp_dob,cmpp_email,cmpp_firstname,cmpp_lastname,cmpp_preferredlanguage").then(
            function success(result) {
                console.log(result);
                const aNumber = result.cmpp_anumber;
                const phoneNo = result.cmpp_cellphone;
                const dob = result.cmpp_dob;
                const email = result.cmpp_email;
                const firstname = result.cmpp_firstname;
                const lastname = result.cmpp_lastname;
                const language = result.cmpp_preferredlanguage;

                var formattedDob = null;
                if (dob) {
                    formattedDob = new Date(dob);
                    if (isNaN(formattedDob.getTime())) {
                        formattedDob = null;
                    }
                }

                formContext.getAttribute("cmpp_anumber").setValue(aNumber);
                formContext.getAttribute("cmpp_dateofbirth").setValue(formattedDob);
                formContext.getAttribute("cmpp_email").setValue(email);
                formContext.getAttribute("cmpp_firstname").setValue(firstname);
                formContext.getAttribute("cmpp_firstlastname").setValue(lastname);
                formContext.getAttribute("cmpp_phoneno").setValue(phoneNo);
                formContext.getAttribute("cmpp_preferredlanguage").setValue(language); 

            },
            function error(error) {
                Xrm.Navigation.openErrorDialog({ message: "Error: " + error.message, details: error.message });
            }
        );
    } else {
        console.log("No lookup field value set.");
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function runConsentEmail(primaryControl) {
    let formContext = primaryControl;

    // Get the GUID of the entity record
    var guid = formContext.data.entity.getId();

    if (!guid) {
        Xrm.Navigation.openAlertDialog({
            text: "The record must first be saved",
            title: "Record Not Saved",
            confirmButtonLabel: "OK"
        });
        return; // Exit the function if the record is not saved
    }

    debugger;
     // Name of the action (workflow)
    var actionName = "ExecuteWorkflow"; 

    var executeWorkflowRequest = {
        entity: {
            id: "34e84019-64b2-ef11-b8e9-001dd83058e3", 
            entityType: "workflow"  
        },
        EntityId: {guid}, 
        getMetadata: function () {
            return {
                boundParameter: "entity",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.workflow",  
                        "structuralProperty": 5
                    },
                    "EntityId": {
                        "typeName": "Edm.Guid",  
                        "structuralProperty": 1
                    }
                },
                operationType: 0,  
                operationName: actionName  
            };
        }
    };
    debugger;
    // Display a progress indicator while processing
    Xrm.Utility.showProgressIndicator("Processing Email...");

    // Execute the workflow using Xrm.WebApi
    Xrm.WebApi.online.execute(executeWorkflowRequest).then(
        function success(result) {
            if (result.ok) {  
             debugger;
                console.log("Workflow executed successfully.:");
            }
            Xrm.Utility.closeProgressIndicator(); 
        },
        function (error) {
            // Handle error if the workflow execution fails
            Xrm.Navigation.openErrorDialog({ message: "Error: " + error.message, details: error.message });
            Xrm.Utility.closeProgressIndicator(); 
        }
    );
}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function BeginFileImport(primaryControl) {
    let formContext = primaryControl;
    let duplicateOption = formContext.getAttribute("cmpp_duplicatehandling").getValue();
    let pluginCalled = false;

    // Check if duplicateOption is null or undefined
    if (duplicateOption == null) {
        var mesgTitle = "Duplicate Handling Is Empty";
        var details = "Duplicate Handling is compulsory. Please select an option.";
        Xrm.Navigation.openErrorDialog({ message: mesgTitle, details: details });
        return;
    }
    debugger;
    if (!pluginCalled) {
        var objActionCallRequest = null;
        let actionName = "cmpp_FileValidationandImport";

        var parameters = {};
        var entity = {};
        entity.id = formContext.data.entity.getId().replace("{", "").replace("}", "");
        entity.entityType = "cmpp_importfile";
        parameters.entity = entity;
        parameters.DuplicateHandlingOption = duplicateOption;

        try {
            objActionCallRequest = {
                entity: parameters.entity,
                DuplicateHandlingOption: parameters.DuplicateHandlingOption,
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.cmpp_importfile",
                                "structuralProperty": 5
                            },
                            "DuplicateHandlingOption": {
                                "typeName": "Edm.Int32",
                                "structuralProperty": 1
                            }
                        },
                        operationType: 0,
                        operationName: actionName
                    };
                }
            };

            Xrm.Utility.showProgressIndicator("Processing File Validation and Import...");
            debugger;
            Xrm.WebApi.online.execute(objActionCallRequest).then(
                function success(result) {
                    if (result.status === 200) {  // Checking for success status
                        result.json().then(function (response) {
                            debugger;
                            if (response) {
                                console.log("Respons: " + response.Response);
                                debugger;
                                showValidationandImportUpdate(response.Response);
                            }
                        });
                    } else {
                        Xrm.Navigation.openErrorDialog({ message: "Issue in File Import", details: "Import execution failed" });
                    }
                    Xrm.Utility.closeProgressIndicator();
                },
                function (ex) {
                    Xrm.Navigation.openErrorDialog({ message: "Error: " + ex.message, details: ex.message });
                    Xrm.Utility.closeProgressIndicator();
                }
            );
        } catch (ex) {
            Xrm.Navigation.openErrorDialog({ message: ex.message, details: ex.message });
            Xrm.Utility.closeProgressIndicator();
        }
    }
}

function showValidationandImportUpdate(validationResponse) {
    debugger;
    // Close the progress indicator before displaying the result
    Xrm.Utility.closeProgressIndicator();

    // If validation is not successful
    if (!validationResponse.IsSuccess) {
        // Parse the JSON response if it's a string
        var parsedResponse;
        if (typeof validationResponse === "string") {
            try {
                parsedResponse = JSON.parse(validationResponse);  // Parsing the JSON response string
            } catch (e) {
                console.error("Error parsing JSON response:", e);
                parsedResponse = validationResponse; // In case parsing fails, use the original response
            }
        } else {
            parsedResponse = validationResponse; // If already an object, use it as is
        }

        // Ensure failureReason exists and is an array, otherwise handle it as a string or empty array
        var errorMessages = "";
        if (Array.isArray(parsedResponse.failureReason)) {
            errorMessages = parsedResponse.failureReason.join("\n");
        } else if (typeof parsedResponse.failureReason === "string") {
            errorMessages = parsedResponse.failureReason;
        } else {
            errorMessages = "Unknown error occurred."; // In case failureReason is undefined or not valid
        }
      debugger;
        var message = "Validation Failed:\n" + errorMessages;
        var alertStrings = {
            confirmButtonLabel: "OK",
            text: message,
            title: "Validation Errors"
        };

        var alertOptions = { height: 300, width: 450 };
        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
    }
    // If validation is successful
    else {
        var successMessage = "File validation is successful. Import will run in the background. Please check the Import Record tab for import details.";
        var successAlertStrings = {
            confirmButtonLabel: "OK",
            text: successMessage,
            title: "Validation Successful"
        };
        debugger;
        var successAlertOptions = { height: 150, width: 450 };
        Xrm.Navigation.openAlertDialog(successAlertStrings, successAlertOptions);
    }
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function updateFileName(executionContext) {
    var formContext = executionContext.getFormContext();
    
    // Check if the form type is "create"
    var formType = formContext.ui.getFormType();
    if (formType !== 1) { // 1 indicates the Create form type
        return; // Exit if not in create mode
    }

    // Get the current value of the file name attribute
    var fileName = formContext.getAttribute("cmpp_name").getValue();
    
    if (fileName === null || fileName.trim() === "") {
        const now = new Date();

        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0'); // Months are zero-indexed
        const day = String(now.getDate()).padStart(2, '0');

        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');

        const currentDateTime = `${year}${month}${day}${hours}${minutes}${seconds}`;

        formContext.getAttribute("cmpp_name").setValue(currentDateTime);
        formContext.data.save();
    }
}


function saveImport(executionContext) {
    var formContext = executionContext.getFormContext();

    var importStarted = formContext.getAttribute("cmpp_startimport").getValue();

    if (importStarted) {
        formContext.data.save();
    }
}



function saveCreatecasetoggle(executionContext) {
    var formContext = executionContext.getFormContext();

    var proceedCaseCreation = formContext.getAttribute("cmpp_proceedwithcasecreation").getValue();

    if (proceedCaseCreation) {
        formContext.data.save();
    }
}


/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

