var cmpp;
(function (cmpp) {
    (function (cmpp_importfile) {
        /** configure this event handler to be called in form OnLoad */
        function OnLoad(executionContext) {
            executionContext.getFormContext();
            //install other event handlers here
        }
        cmpp_importfile.OnLoad = OnLoad;
        /** configure this event handler to be called in form OnSave */
        function OnSave(executionContext) {
            executionContext.getFormContext();
            update(executionContext);
        }
        cmpp_importfile.OnSave = OnSave;
        function update(executionContext) {
            var formContext = executionContext.getFormContext();
            var filename = formContext.getAttribute("cmpp_name").getValue();
            alert(filename);
        }
    })(cmpp.cmpp_importfile || (cmpp.cmpp_importfile = {}));
})(cmpp || (cmpp = {}));

