function fetchMotorClaimDetail(executionContext) {
  let formContext = executionContext.getFormContext();
  let nationalId = formContext.getAttribute("vrp_nationalid").getValue();
  let claimnumber = formContext.getAttribute("vrp_name").getValue();
  let lineOfBiz = formContext.getAttribute("ey_lineofbusiness").getValue();
  let claimDetailFetched = false;



  if (!claimDetailFetched) {
    if (nationalId == null || claimnumber == null) {
      return;
    }



    if (lineOfBiz != null && lineOfBiz == 364840001) {
      var objActionCallRequest = null;
      let actionName = "vrp_GetMotorClaimsDetails";



      var parameters = {};
      var entity = {};
      entity.id = formContext.data.entity.getId().replace("{", "").replace("}", "");
      entity.entityType = "vrp_claim";
      parameters.entity = entity;
      parameters.NationalID = nationalId;
      parameters.ClaimNumber = claimnumber;



      try {
        objActionCallRequest = {
          entity: parameters.entity,
          NationalID: parameters.NationalID,
          ClaimNumber: parameters.ClaimNumber,



          getMetadata: function () {
            return {
              boundParameter: "entity",
              parameterTypes: {
                "entity": {
                  "typeName": "mscrm.vrp_claim",
                  "structuralProperty": 5
                },
                "NationalID": {
                  "typeName": "Edm.String",
                  "structuralProperty": 1
                },
                "ClaimNumber": {
                  "typeName": "Edm.String",
                  "structuralProperty": 1
                }
              },
              operationType: 0,
              operationName: actionName
            };
          }
        };



        Xrm.Utility.showProgressIndicator("Fetching Motor Claim Detail...");
        Xrm.WebApi.online.execute(objActionCallRequest).then(
          function success(result) {
            console.log(result);
            if (result.ok) {
              result.json().then(function (response) {
                if (response != null) {
                  Xrm.Utility.closeProgressIndicator();
                  if (response.IsSuccess) {
                    claimDetailFetched = true;
                    console.log('API successfully retrieved the claim detail');
                    formContext.data.save(); // Save the form context
                    Xrm.Utility.closeProgressIndicator();
                  }
                }
              });
            } else {
              console.log("Issue from Plugin");
            }
          },
          function (ex) {
            Xrm.Navigation.openErrorDialog({ message: "Error :: " + ex.message, details: ex.message });
            Xrm.Utility.closeProgressIndicator();
          }
        );
      }
      catch (ex) {
        Xrm.Navigation.openErrorDialog({ message: ex.message, details: ex.message });
        Xrm.Utility.closeProgressIndicator();
      }
    }
  }
}
