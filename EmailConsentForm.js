function BeginFileImport(primaryControl) {
    let formContext = primaryControl;
    var orgUrl = Xrm.Utility.getGlobalContext().getClientUrl(); 
    var apiEndpoint = "/api/data/v9.0/ExportPdfDocument";
    var firstName = formContext.getAttribute("cmpp_firstname").getValue();
    var lastName = formContext.getAttribute("cmpp_lastname").getValue();
    var emailaddress = formContext.getAttribute("cmpp_email").getValue();
    var caseId = formContext.data.entity.getId().slice(1, -1);
    var base64Pdf = "";
    let pluginCalled = false;

    // Check if duplicateOption is null or undefined
    if (emailaddress  == null) {
        var mesgTitle = "Email address Is Empty";
        var details = "Provide an email address for the participant";
        Xrm.Navigation.openErrorDialog({ message: mesgTitle, details: details });
        return;
    }

    var requestData = {
        "EntityTypeCode": 11085,
        "SelectedTemplate": {
            "@odata.type": "Microsoft.Dynamics.CRM.documenttemplate",
            "documenttemplateid": "d916580b-28be-ef11-b8e9-001dd8305876"
        },
        "SelectedRecords": `["${caseId}"]`
    };

    // Construct the full API URL
    var fullUrl = orgUrl + apiEndpoint;
    
    // Show the progress indicator before starting the API call
    Xrm.Utility.showProgressIndicator("Export Consent Form...");

    // Use XMLHttpRequest to make the POST request
    var xhr = new XMLHttpRequest();
    xhr.open("POST", fullUrl, true);
    xhr.setRequestHeader("Accept", "application/json");
    xhr.setRequestHeader("Content-Type", "application/json");

    // Success callback
    xhr.onload = async function () {
        // Close the progress indicator when the API call completes (either success or failure)
        Xrm.Utility.closeProgressIndicator();

        if (xhr.status === 200) {
            console.log("PDF Export successful");
            var response = JSON.parse(xhr.responseText);

            // Check if PdfFile exists and is not null or empty
            if (response.PdfFile && response.PdfFile.trim() !== "") {
                base64Pdf = response.PdfFile;
                debugger;
                
            } else {
                console.error("No PDF data found in the response.");
            }
        } else {
            console.error("Error exporting PDF", xhr.responseText);
        }
    };

    // Error callback
    xhr.onerror = function () {
        // Close the progress indicator in case of an error
        Xrm.Utility.closeProgressIndicator();
        console.error("Request failed", xhr.statusText);
    };

    // Send the request with the data
    xhr.send(JSON.stringify(requestData));
    
    debugger;
    if (!pluginCalled) {
        var objActionCallRequest = null;
        let actionName = "cmpp_SendConsentEmail";

        var parameters = {};
        var entity = {};
        entity.id = caseId;
        entity.entityType = "cmpp_case";
        parameters.entity = entity;
        parameters.Base64Pdf = base64Pdf;
        entity.FirstName = firstName;
        parameters.LastName = lastname;
        parameters.EmailAddress = emailaddress;

        try {
            objActionCallRequest = {
                entity: parameters.entity,               
                entity: parameters.entity, 
                Base64Pdf: parameters.Base64Pdf,               
                FirstName: parameters.FirstName,
                LastName: parameters.LastName,
                EmailAddress: parameters.EmailAddress,
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.cmpp_case",
                                "structuralProperty": 5
                            },
                            "Base64Pdf": {
                                "typeName": "Edm.String",
                                "structuralProperty": 1
                            },
                            "FirstName": {
                                "typeName": "Edm.String",
                                "structuralProperty": 1
                            },
                            "LastName": {
                                "typeName": "Edm.String",
                                "structuralProperty": 1
                            },
                            "EmailAddress": {
                                "typeName": "Edm.String",
                                "structuralProperty": 1
                            }
                        },
                        operationType: 0,
                        operationName: actionName
                    };
                }
            };

            Xrm.Utility.showProgressIndicator("Processing...");
            debugger;
            Xrm.WebApi.online.execute(objActionCallRequest).then(
                function success(result) {
                    if (result.status === 200) {  // Checking for success status
                        result.json().then(function (response) {
                            debugger;
                            if (response) {
                                console.log("Response: " + response.IsSuccess);
                                debugger;
                            }
                        });
                    } else {
                        Xrm.Navigation.openErrorDialog({ message: "Issue in Emailing Consent Form", details: "Sending Consent Form failed" });
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


