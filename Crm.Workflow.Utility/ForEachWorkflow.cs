// <copyright file="InvokeWorkflow.cs" >
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Gokul JB</author>
// <date>6/19/2018 </date>
// <summary>Implements the InvokeWorkflow Workflow Activity.</summary>
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Crm.Workflow.Utility
{
    public sealed class ForEachWorkflow : CodeActivity
    {
        #region Public Variable

        [RequiredArgument]
        [Input("Child Entity Name")]
        public InArgument<string> childEntityName { get; set; }

         [RequiredArgument]
        [Input("Relationship Name")]
        public InArgument<string> relationshipName { get; set; }

        [RequiredArgument]
        [ReferenceTarget("workflow")]
        [Input("Child Workflow Name")]
        public InArgument<EntityReference> childWorkflow { get; set; }

        #endregion

        #region Private Variable
        private ITracingService tracingService = null;
        #endregion

        protected override void Execute(CodeActivityContext executionContext)
        {
            
            string relatedEntityName = string.Empty;
            string relationshipSchmemaName = string.Empty;
            EntityReference workflowReference = null;
            EntityCollection entityCollection = null;
            IWorkflowContext context = null;
            IOrganizationServiceFactory serviceFactory= null;
            IOrganizationService service = null;
            

            try
            {
                
                context = executionContext.GetExtension<IWorkflowContext>();
                if (context == null) throw new InvalidPluginExecutionException("Failed to retrieve workflow context");
                tracingService = executionContext.GetExtension<ITracingService>();
                if (tracingService == null) throw new InvalidPluginExecutionException("Failed to initialize tracing service");
                serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                if (serviceFactory == null) throw new InvalidPluginExecutionException("Failed to retrieve Service factory");
                service = serviceFactory.CreateOrganizationService(context.UserId);
                if (service == null) throw new InvalidPluginExecutionException("Failed to initialize Organization Service");

                tracingService.Trace("After initializing service.....");
                workflowReference = childWorkflow.Get(executionContext);
                relatedEntityName = Convert.ToString(childEntityName.Get(executionContext));
                relationshipSchmemaName = Convert.ToString(relationshipName.Get(executionContext));
                

                tracingService.Trace("Related Entity Name: {0}, Relationsip Name: {1}, Workflow Name: {2}", relatedEntityName, relationshipName, workflowReference.Name);
                if (workflowReference != null && relatedEntityName != null && relationshipSchmemaName != null)
                {
                    tracingService.Trace("Call GetAssociatedEntityItems method");
                    //Call GetAssociatedEntityItems method to related entity collection
                    entityCollection = GetAssociatedEntityItems(context.PrimaryEntityName, context.PrimaryEntityId, relatedEntityName, relationshipSchmemaName, service);

                    if (entityCollection != null)
                    {
                        tracingService.Trace("Call InvokeWorkflows method");
                        //Call InvokeWorkflows method to invoke workflows dynamically
                        InvokeWorkflows(workflowReference.Id, entityCollection, service);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            


        }

        /// <summary>
        /// GetAssociatedEntityItems method is to get related entity details
        /// </summary>
        /// <param name="primaryEntityName"></param>
        /// <param name="_primaryEntityId"></param>
        /// <param name="relatedEntityName"></param>
        /// <param name="relationshipName"></param>
        /// <param name="serviceProxy"></param>
        /// <returns></returns>
        private EntityCollection GetAssociatedEntityItems(string primaryEntityName, Guid _primaryEntityId, string relatedEntityName, string relationshipName, IOrganizationService serviceProxy)
        {

            try
            {

                tracingService.Trace("Entering into InvokeWorkflows method....");
                EntityCollection result = null;
                QueryExpression relatedEntityQuery = new QueryExpression();
                relatedEntityQuery.EntityName = relatedEntityName;
                relatedEntityQuery.ColumnSet = new ColumnSet(false);

                Relationship relationship = new Relationship();
                relationship.SchemaName = relationshipName;
                //relationship.PrimaryEntityRole = EntityRole.Referencing;
                RelationshipQueryCollection relatedEntity = new RelationshipQueryCollection();
                relatedEntity.Add(relationship, relatedEntityQuery);

                RetrieveRequest request = new RetrieveRequest();
                request.RelatedEntitiesQuery = relatedEntity;
                request.ColumnSet = new ColumnSet(true);
                request.Target = new EntityReference
                {
                    Id = _primaryEntityId,
                    LogicalName = primaryEntityName
                };

                RetrieveResponse response = (RetrieveResponse)serviceProxy.Execute(request);
                RelatedEntityCollection relatedEntityCollection = response.Entity.RelatedEntities;

                tracingService.Trace("After get the RelatedEntityCollection");
                if (relatedEntityCollection != null)
                {
                    tracingService.Trace("RelatedEntityCollection Count: {0}, RelatedEntityCollection.Values.Count:{1}", relatedEntityCollection.Count, relatedEntityCollection.Values.Count);
                    if (relatedEntityCollection.Count > 0 && relatedEntityCollection.Values.Count > 0)
                    {
                        result = (EntityCollection)relatedEntityCollection.Values.ElementAt(0);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// InvokeWorkflows method is to call workflow dynamically based on passed parameters
        /// </summary>
        /// <param name="_workflowGuid"></param>
        /// <param name="entityResult"></param>
        /// <param name="serviceProxy"></param>
        private void InvokeWorkflows(Guid _workflowGuid, EntityCollection entityResult, IOrganizationService serviceProxy)
        {

            tracingService.Trace("Related Entity Collection Count: {0}", entityResult.Entities.Count);
           
            entityResult.Entities.ToList().ForEach(entity =>
            {
                ExecuteWorkflowRequest request = new ExecuteWorkflowRequest
                {
                    EntityId = entity.Id,
                    WorkflowId = _workflowGuid
                };

                serviceProxy.Execute(request); //run the workflow
            });
        }
    }



}
