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
