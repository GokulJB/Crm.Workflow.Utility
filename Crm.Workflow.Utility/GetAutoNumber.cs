// <copyright file="GetAutoNumber.cs" >
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Gokul JB</author>
// <date>6/19/2018 </date>
// <summary>Implements the GetAutoNumber Workflow Activity.</summary>
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Crm.Workflow.Utility
{

    public sealed class GetAutoNumber : CodeActivity
    {
        #region Public Variable

        [ReferenceTargetAttribute("uu_autonumberconfig")]
        [Input("Auto Number Reference")]
        [RequiredArgument]
        [ArgumentDescription("It will pass Auto number config refference")]
        public InArgument<EntityReference> autoNumberEntity { get; set; }

        [Output("Formatted Auto Number")]
        [ArgumentDescription("It return Formatted auto Number")]
        public OutArgument<string> returnAutoNumber { get; set; }

        [Output("Numeric Auto Number")]
        [ArgumentDescription("It return un Formatted auto Number")]
        public OutArgument<int> ReturnUNFormattedAutoNumber { get; set; }
        

        #endregion

        #region Private Variable

        private EntityReference autoNumberRef = null;
        private int autoNumber;
        private string formatedAutoNumber = string.Empty;
        ITracingService tracingService { get; set; }

        # endregion



        protected override void Execute(CodeActivityContext executioncontext)
        {


            int lastNumber;
            int defaultDigit;
            int incrementValue;
            string formatVal = string.Empty;
            Entity entity = null;

            IWorkflowContext context = executioncontext.GetExtension<IWorkflowContext>();
            if (context == null) throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");

             tracingService = executioncontext.GetExtension<ITracingService>();
            if (tracingService == null) throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");

            IOrganizationServiceFactory serviceFactory = executioncontext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (null == service) throw new InvalidPluginExecutionException("Failed to retrieve service.");

            try
            {

                autoNumberRef = autoNumberEntity.Get(executioncontext);
                tracingService.Trace("Auto Number Logical Name:{0}, ID:{1}, Name:{2} ", autoNumberRef.LogicalName, autoNumberRef.Id, autoNumberRef.Name);

                entity = service.Retrieve(autoNumberRef.LogicalName, autoNumberRef.Id, new ColumnSet(true));


                if (entity != null)
                {
                    //entity = null;
                    tracingService.Trace("Start Auto Number Generation");

                    lastNumber = entity.GetAttributeValue<int>("uu_lastnumber");
                    defaultDigit = entity.GetAttributeValue<int>("uu_defaultdegit");
                    incrementValue = entity.GetAttributeValue<int>("uu_increment");
                    formatVal = entity.GetAttributeValue<string>("uu_format");

                    tracingService.Trace("Auto Number details - Last Number:{0}, DefaultDigit:{1}, Increment Value:{2}, Formatted Value:{3} ", lastNumber, defaultDigit, incrementValue, formatVal);

                    if (defaultDigit == 0 || incrementValue == 0 ) return;

                    returnAutoNumber.Set(executioncontext, FetchAutoNumber(lastNumber, defaultDigit, incrementValue, formatVal));

                    tracingService.Trace("Generated Number:{0}, Formatted Number:{1}", autoNumber, formatedAutoNumber);
                    ReturnUNFormattedAutoNumber.Set(executioncontext, autoNumber);
                    entity.Attributes["uu_lastnumber"] = autoNumber;
                    service.Update(entity);
                }

                tracingService.Trace("End Auto Number Generation");
            }
            catch (Exception ex)
            {
                tracingService.Trace(ex.ToString());
                throw new InvalidWorkflowException(ex.ToString());
            }

        }

        /// <summary>
        /// GetAutoNumber to return next formatted Value
        /// </summary>
        /// <param name="lastNumber"></param>
        /// <param name="defaultDigit"></param>
        /// <param name="incrementValue"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public string FetchAutoNumber(int lastNumber, int defaultDigit, int incrementValue, string format)
        {
            int autoNumLength;
            

            autoNumber = lastNumber + incrementValue;
            autoNumLength = autoNumber.ToString("D").Length;
            formatedAutoNumber = autoNumLength < defaultDigit ? autoNumber.ToString("D" + defaultDigit.ToString()) : Convert.ToString(autoNumber);
            tracingService.Trace("Formate Value: {0}, Formatted Number: {1}, Auto Number: {2}", format, formatedAutoNumber, autoNumber);
            formatedAutoNumber = string.IsNullOrEmpty(format)? formatedAutoNumber:format.Contains("{0}") ? string.Format(format, formatedAutoNumber) : format + formatedAutoNumber;

            return formatedAutoNumber;
        }
    }
}
