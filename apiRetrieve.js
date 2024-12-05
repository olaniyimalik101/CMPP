function retrieveLookupFieldValue(executionContext) {
    var formContext = executionContext.getFormContext();
    
    var relatedCase = formContext.getAttribute("cmpp_case").getValue();

    if (relatedCase != null && relatedCase.length > 0) {
        var caseRecordId = relatedCase[0].id; 
        var caseEntityType = relatedCase[0].entityType; 

        Xrm.WebApi.retrieveRecord(caseEntityType, caseRecordId, "?$select=cmpp_anumber,cmpp_cellphone,cmpp_dob,cmpp_email,cmpp_firstname, cmpp_lastname,").then(
            function success(result) {
                console.log("Related Record Data: ", result);
                const aNumber = result.cmpp_anumber;
                const phoneNo = result.cmpp_cellphone;
                const dob = result.cmpp_dob;
                const email = result.cmpp_email;
                const firstname = result.cmpp_firstname;
                const lastname = result.cmpp_lastname;

                formContext.getAttribute("cmpp_anumber").setValue(aNumber);
                formContext.getAttribute("cmpp_dateofbirth").setValue(dob);
                formContext.getAttribute("cmpp_email").setValue(email);
                formContext.getAttribute("cmpp_firstname").setValue(firstname);
                formContext.getAttribute("cmpp_firstlastname").setValue(lastname);
                formContext.getAttribute("cmpp_phoneno").setValue(phoneNo);

                formContext.data.save();

            },
            function error(error) {
                Xrm.Navigation.openErrorDialog({ message: "Error: " + error.message, details: error.message });
            }
        );
    } else {
        console.log("No lookup field value set.");
    }
}
