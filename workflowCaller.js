function runConsentEmail(primaryControl) {
    let formContext = primaryControl;

    // Get the GUID of the entity record
    var guid = formContext.data.entity.getId();  

     // Name of the action (workflow)
    var actionName = "ExecuteWorkflow"; 

    var executeWorkflowRequest = {
        entity: {
            id: "5d08630a-b21f-428a-a222-db8b4d4f6e60", 
            entityType: "workflow"  
        },
        EntityId: guid, 
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
    
    // Display a progress indicator while processing
    Xrm.Utility.showProgressIndicator("Processing Email...");

    // Execute the workflow using Xrm.WebApi
    Xrm.WebApi.online.execute(executeWorkflowRequest).then(
        function success(result) {
            if (result.ok) {  
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
