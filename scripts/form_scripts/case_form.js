/*************************************************************
 * Enrollment Case Form Script.
 * Last update 11/12/2024
 ************************************************************/
var cmpp;
(function (cmpp) {
    (function (cmpp_case) {
        let formContext;
        /** configure this event handler to be called in form OnLoad */
        function OnLoad(executionContext) {
            formContext = executionContext.getFormContext();
            //install other event handlers here
            let dobAttr = formContext.getAttribute("cmpp_dob");
            dobAttr && dobAttr.addOnChange(validateDayOfBirth);
        }
        cmpp_case.OnLoad = OnLoad;
        /** configure this event handler to be called in form OnSave */
        function OnSave(executionContext) {
            formContext = executionContext.getFormContext();
        }
        cmpp_case.OnSave = OnSave;
        /** enfore cmpp_dob not in future */
        function validateDayOfBirth(eCtx) {
            let formCtx = eCtx.getFormContext();
            let dobAttr = formCtx.getAttribute("cmpp_dob");
            let dobAttrValue = dobAttr ? dobAttr.getValue() : null;
            if (dobAttrValue == null)
                return;
            let currentDate = new Date();
            let failed = false;
            let errorMessage = "";
            if (dobAttrValue >= currentDate) {
                failed = true;
                errorMessage = "date value cannot be in the future";
            }
            let dobCtr = formCtx.getControl("cmpp_dob");
            if (failed) {
                dobCtr.setNotification(errorMessage, "cmpp_dob_unique_notification_id");
            }
            else {
                dobCtr.clearNotification("cmpp_dob_unique_notification_id");
            }
            return failed;
        }
    })(cmpp.cmpp_case || (cmpp.cmpp_case = {}));
})(cmpp || (cmpp = {}));

