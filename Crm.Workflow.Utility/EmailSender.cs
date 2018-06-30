// <copyright file="EmailSender.cs" company="Test EVN">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Gokul JB</author>
// <date>6/26/2017 5:39:04 PM</date>
// <summary>Implements the EmailSender Workflow Activity.</summary>
namespace Crm.Workflow.Utility
{
    using System;
    using System.Activities;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Workflow;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Query;
    using System.ComponentModel.DataAnnotations;
    using System.Linq;
    public sealed class EmailSender : CodeActivity
    {
        /// <summary>
        /// Input Parameters
        /// </summary>


        [Input("Email Template Name")]
        [RequiredArgument]
        [Default("None")]
        [ArgumentDescription("The text contain email template name")]
        public InArgument<string> EmailTemplateName { get; set; }

        // Add New Feature for pass Draft email : Vasanth 13/07/2017

        [ReferenceTarget("email")]
        [Input("Draft Email Reference")]
        public InArgument<EntityReference> DraftEmailReference { get; set; }

        [Input("From Address")]
        [ReferenceTarget("queue")]
        [RequiredArgument]
        [ArgumentDescription("Sender from address(Type of Queue)")]
        public InArgument<EntityReference> QueueFromAddress { get; set; }

        [Input("To Address")]
        [RequiredArgument]
        [ArgumentDescription("The text contains  To address of email")]
        public InArgument<string> ToAddress { get; set; }


        [Input("Category")]
        [ArgumentDescription("The text useful for to pass category value if required")]
        public InArgument<string> Category { get; set; }

        [Output("ReturnValue")]
        [ArgumentDescription("User friendly message, if validation got fail")]
        public OutArgument<string> returnValue { get; set; }

        #region Priavate variables
        /// <summary>
        /// Private variables
        /// </summary>
        private ITracingService tracingService { get; set; }
        private IOrganizationService service { get; set; }
        private CodeActivityContext executionContext { get; set; }
        private IWorkflowContext context { get; set; }
        private Entity queue_Entity = null;
        private Entity party_queue_Entity = null;
        private EntityCollection fromQueueActivityParty = null;
        EntityCollection activityPartis = null;
        private string templateName = string.Empty;
        private string queryString = string.Empty;
        private string category = string.Empty;
        private EntityReference draftEmailRef = null;


        #endregion

        #region Constant value

        private const string emailTemplateEntity = "template";
        private const string activityPartyEntity = "activityparty";
        private const string activityPartyPartyidField = "partyid";

        #endregion


        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext codeActivityContext)
        {
            executionContext = codeActivityContext;
            // Create the tracing service
            tracingService = executionContext.GetExtension<ITracingService>();

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered EmailSender.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("EmailSender.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            service = serviceFactory.CreateOrganizationService(context.UserId);

            if (null == service)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve service.");
            }

            try
            {
                SendEmail();
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting EmailSender.Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        #region Private Methods to send Email
        /// <summary>
        /// SendEmail method useful for get email and prepare email template to send to customer
        /// </summary>
        private void SendEmail()
        {
            tracingService.Trace("Entered EmailSender.SendEmail()");
            try
            {
                templateName = EmailTemplateName.Get(executionContext);
                queryString = ToAddress.Get(executionContext);
                category = Category.Get(executionContext);

                tracingService.Trace("EmailSender.SendEmail(), Template Name:{0} InPutEmailString: {1}, Category: {2}", queryString, templateName, category);

                if (!string.IsNullOrEmpty(templateName) && !string.IsNullOrEmpty(queryString))
                {
                    EntityCollection emails = null;
                    if (queryString.TrimStart().StartsWith("<"))
                    {
                        emails = GetEmailsByFetchXML();
                    }
                    else
                    {
                        emails = GetEmailByString();
                    }

                    if (null != emails && emails.Entities.Count > 0)
                    {
                        PrepareEmail(emails);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// GetEmailsByFetchXML method useful for to get list of emails based on fetchXML query
        /// </summary>
        /// <returns></returns>
        private EntityCollection GetEmailsByFetchXML()
        {
            try
            {
                tracingService.Trace("Entered EmailSender.GetEmailsByFetchXML()");

                string contextPrimaryId = Convert.ToString(context.PrimaryEntityId);
                EntityCollection result = service.RetrieveMultiple(new FetchExpression(queryString));

                if (null != result && result.Entities.Count > 0)
                {
                    activityPartis = new EntityCollection();
                    activityPartis.EntityName = activityPartyEntity;

                    for (int i = 0; i < result.Entities.Count; i++)
                    {
                        Entity entity = result.Entities[i];

                        foreach (var item in entity.Attributes.Keys)
                        {
                            tracingService.Trace("EmailSender.GetEmailsByFetchXML(), Item Name: {0}, Item Value: {1}", item, entity.Attributes[item]);

                            string itemValue = GetValuefromAttribute(entity.Attributes[item]);
                            if (!string.IsNullOrEmpty(itemValue) && contextPrimaryId != itemValue)
                            {
                                if (IsValidEmail(itemValue))
                                {
                                    Entity email = new Entity();
                                    email.LogicalName = activityPartyEntity;
                                    email.Attributes["addressused"] = itemValue;
                                    activityPartis.Entities.Add(email);
                                }
                            }
                            else
                            {
                                //resultOutPut.Set(executionContext, contextPrimaryId +"---"+ Convert.ToString(entity.Attributes[item])); 
                            }
                        }
                    }

                    tracingService.Trace("EmailSender.GetEmailsByFetchXML(), Email Count: {0}", activityPartis.Entities.Count);
                    return activityPartis.Entities.Count > 0 ? activityPartis : null;
                }
                else
                {
                    returnValue.Set(executionContext, "No record found for this contact, please check the condition");

                    return activityPartis;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private EntityCollection GetEmailByString()
        {
            tracingService.Trace("Entered EmailSender.GetEmailByString()");
            try
            {
                string[] separators = { ",", ";", " " };
                string[] arrayEmails = queryString.Split(separators, StringSplitOptions.RemoveEmptyEntries);

                if (arrayEmails.Length > 0)
                {
                    tracingService.Trace("EmailSender.GetEmailsByFetchXML(), EmailCount: {0}", Convert.ToString(arrayEmails.Length));
                    activityPartis = new EntityCollection();
                    activityPartis.EntityName = activityPartyEntity;
                    // To restrict duplicate emails
                    arrayEmails = arrayEmails.Distinct().ToArray();
                    foreach (var email in arrayEmails)
                    {
                        tracingService.Trace("EmailSender.GetEmailsByFetchXML(), Input Email: {0}", email);
                        Entity entity = new Entity(activityPartyEntity);
                        entity.Attributes["addressused"] = email;
                        activityPartis.Entities.Add(entity);
                    }

                }
                return activityPartis;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// PreapreEmailTemplate method useful for prepare email and email template
        /// </summary>
        /// <param name="toActivityParty"></param>
        private void PrepareEmail(EntityCollection toActivityParty)
        {
            tracingService.Trace("EmailSender.PreapreEmailTemplate()");
            try
            {
                #region Prepare Email format

                // Creating Email 'from' recipient activity party entity object
                //Entity emailFromReciepent = new Entity(activityPartyEntity);

                // Assigning receiver email address to activity party addressused attribute
                //wod_EmailToReciepent["participationtypemask"] = new OptionSetValue(0);

                EntityReference in_queueEntityReference = QueueFromAddress.Get(executionContext);
                if (in_queueEntityReference != null)
                {
                    queue_Entity = service.Retrieve(in_queueEntityReference.LogicalName, in_queueEntityReference.Id, new ColumnSet(true));
                    if (queue_Entity != null)
                    {
                        party_queue_Entity = new Entity(activityPartyEntity);
                        party_queue_Entity[activityPartyPartyidField] = in_queueEntityReference;

                        fromQueueActivityParty = new EntityCollection();
                        fromQueueActivityParty.Entities.Add(party_queue_Entity);
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("Check Workflow Parameters: Error Generating 'From' Field from Queue input parameter.  Ensure Queue is Configured Correctly.");
                    }
                }
                else
                {
                    throw new InvalidPluginExecutionException("Check Workflow Parameters: Error Generating 'From' Field from Queue input parameter.  Ensure Queue is Configured Correctly.");
                }

                // Setting from user account
                //emailFromReciepent.Attributes[activityPartyPartyidField] = new EntityReference("systemuser", context.UserId);

                // Creating Email entity object


                #endregion

                #region Construct Email

                draftEmailRef = DraftEmailReference.Get(executionContext);
                Entity emailEntity;

                if (draftEmailRef == null)
                {
                    emailEntity = new Entity("email");

                    // Setting email entity 'to' attribute value
                    emailEntity.Attributes["to"] = toActivityParty;
                    emailEntity.Attributes["from"] = fromQueueActivityParty;
                    if (!string.IsNullOrEmpty(category))
                        emailEntity.Attributes["category"] = category;

                    SendEmailByTemplate(templateName, emailEntity);
                }
                else
                {
                    emailEntity = service.Retrieve(draftEmailRef.LogicalName, draftEmailRef.Id, new ColumnSet(true));
                    if (emailEntity == null)
                    {
                        throw new InvalidPluginExecutionException("Email Entity is empty");
                    }
                    else
                    {
                        // Setting email entity 'to' attribute value
                        emailEntity.Attributes["to"] = toActivityParty;
                        emailEntity.Attributes["from"] = fromQueueActivityParty;
                        if (!string.IsNullOrEmpty(category))
                            emailEntity.Attributes["category"] = category;

                        service.Update(emailEntity);

                        SendEmailByDraft(emailEntity);
                    }

                #endregion                

                }

                #region Send Email With Template



                #endregion

            }
            catch (Exception ex)
            {
                throw ex;

            }
        }
        /// <summary>
        /// SendEmailWithTemplate method useful for to appened email temaplate with Email and send to customer
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="templateName"></param>
        /// <param name="entiy"></param>
        /// <param name="context"></param>
        private void SendEmailByTemplate(string templateName, Entity entiy)
        {
            tracingService.Trace("EmailSender.SendEmailByTemplate()");
            try
            {
                Guid templateId = Guid.Empty;

                // Get Template Id by Name

                Entity template = GetTemplateByName(templateName);

                if (template != null && template.Id != null)
                {
                    tracingService.Trace("EmailSender.SendEmailByTemplate(), Template ID: {0}, Entity Name:{1}", Convert.ToString(template.Id), context.PrimaryEntityName);
                    var emailUsingTemplateReq = new SendEmailFromTemplateRequest
                    {
                        Target = entiy,
                        TemplateId = template.Id,
                        RegardingId = context.PrimaryEntityId,
                        RegardingType = context.PrimaryEntityName
                    };

                    var emailUsingTemplateResp = (SendEmailFromTemplateResponse)service.Execute(emailUsingTemplateReq);
                    // resultOutPut.Set(executionContext, Convert.ToString(template.Id));
                }

                else
                {
                    tracingService.Trace("EmailSender.SendEmailByTemplate()===> No Template found");
                    // “****No email template exists with the given name ****”);

                }
            }
            catch (Exception ex)
            {
                throw ex;

            }
        }

        /// <summary>
        /// GetTemplateByName method useful for to get template Id from template entity by template name
        /// </summary>
        /// <param name="title"></param>
        /// <param name="crmService"></param>
        /// <returns></returns>
        private Entity GetTemplateByName(string title)
        {
            tracingService.Trace("EmailSender.GetTemplateByName()");

            try
            {
                var query = new QueryExpression();

                query.EntityName = emailTemplateEntity;

                var filter = new FilterExpression();

                var condition1 = new ConditionExpression("title", ConditionOperator.Equal, new object[] { title });

                filter.AddCondition(condition1);

                query.Criteria = filter;

                EntityCollection allTemplates = service.RetrieveMultiple(query);

                Entity emailTemplate = null;

                if (allTemplates.Entities.Count > 0)
                {
                    tracingService.Trace("EmailSender.GetTemplateByName(), Template Count:{0}", Convert.ToString(allTemplates.Entities.Count));

                    emailTemplate = allTemplates.Entities[0];

                }
                else
                {
                    returnValue.Set(executionContext, "Template " + '"' + title + '"' + " is not available in our record, please check the record ");
                }

                return emailTemplate;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// GetValuefromAttribute
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        private static string GetValuefromAttribute(object p)
        {
            try
            {
                //if (p.ToString() == "Microsoft.Xrm.Sdk.EntityReference")
                //{
                //    return ((EntityReference)p).Name;
                //}
                //if (p.ToString() == "Microsoft.Xrm.Sdk.OptionSetValue")
                //{
                //    return ((OptionSetValue)p).Value.ToString();
                //}
                //if (p.ToString() == "Microsoft.Xrm.Sdk.Money")
                //{
                //    return ((Money)p).Value.ToString();
                //}
                if (p.ToString() == "Microsoft.Xrm.Sdk.AliasedValue")
                {
                    if ((p as AliasedValue).Value is OptionSetValue)
                        return ((p as AliasedValue).Value as OptionSetValue).Value.ToString();
                    else if ((p as AliasedValue).Value is EntityReference)
                        return ((EntityReference)(((AliasedValue)(p)).Value)).Name;
                    else if ((p as AliasedValue).Value is Money) // Handling DataType Money for "fixedfreighttotalsku" field & FA removing issue fix
                        return ((Money)(p as AliasedValue).Value).Value.ToString();
                    else
                        return ((AliasedValue)p).Value.ToString();
                }
                else
                {
                    return p.ToString();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// IsValidEmail method useful for validate email address
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public bool IsValidEmail(string source)
        {

            return new EmailAddressAttribute().IsValid(source);
        }
        #endregion

        #region Methods for Send Draft Email

        private void SendEmailByDraft(Entity entity)
        {
            tracingService.Trace("EmailSender.SendEmailByDraft()");
            try
            {
                GetTrackingTokenEmailRequest trackingTokenEmailRequest = new GetTrackingTokenEmailRequest();
                GetTrackingTokenEmailResponse trackingTokenEmailResponse = null;

                var draftEmailReq = new SendEmailRequest
                {
                    EmailId = entity.Id,
                    IssueSend = true,

                };

                trackingTokenEmailResponse = (GetTrackingTokenEmailResponse)service.Execute(trackingTokenEmailRequest);

                // setting email tracking token
                draftEmailReq.TrackingToken = trackingTokenEmailResponse.TrackingToken;

                // send request
                service.Execute(draftEmailReq);

            }
            catch (Exception ex)
            {
                throw ex;

            }
        }

        #endregion

    }

}