using Microsoft.Xrm.Sdk;
using System;
using Vrp.Crm.C360.Plugins.Common.Handler;



namespace Vrp.Crm.C360.Plugins.Plugins
{
    public class GetMotorClaimDetailPluginActionExecuter : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext pluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(pluginExecutionContext.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));



            string nationalid;
            string claimnumber;



            try
            {
                tracingService.Trace("GetClaimDetailPluginActionExecuter Code Execution Started");



                var targetEntity = (EntityReference)pluginExecutionContext.InputParameters["Target"];
                tracingService.Trace("TargetEntity: " + targetEntity.LogicalName + ", TargetEntityID: " + targetEntity.Id);



                if (pluginExecutionContext.InputParameters.Contains("NationalID") && pluginExecutionContext.InputParameters.Contains("ClaimNumber"))
                {
                    nationalid = pluginExecutionContext.InputParameters["NationalID"].ToString();
                    claimnumber = pluginExecutionContext.InputParameters["ClaimNumber"].ToString();



                    MotorClaimDetailHandler MotorclaimDetailHander = new MotorClaimDetailHandler(service, tracingService);



                    tracingService.Trace("Calling ExecuteGetMotorClaimDetailAction");



                    MotorclaimDetailHander.ExecuteGetMotorClaimDetailAction(nationalid, claimnumber, targetEntity);



                    tracingService.Trace("ExecuteGetMotorClaimDetailAction Ends");



                    pluginExecutionContext.OutputParameters["IsSuccess"] = true;//execution is successful
                }
                else
                {
                    tracingService.Trace("NationalId or ClaimNumber is missing");
                }
                tracingService.Trace("GetMotorClaimDetailPluginActionExecuter Code Execution Ended");
            }



            catch (Exception ex)
            {



                tracingService.Trace(ex.Message);
                ExceptionLog.ExceptionLogs("Middleware", string.Empty, string.Empty, null, ex.Message.ToString(), ex.Message.ToString(), "Vrp.Crm.C360.GetMotorClaimDetailPluginActionExecuter", service);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
 
