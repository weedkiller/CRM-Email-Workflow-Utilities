﻿using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.Email
{
    public class CcQueueMembers : WorkFlowActivityBase
    {
        public CcQueueMembers() : base(typeof(CcQueueMembers)) { }

        [RequiredArgument]
        [Input("Email To Send")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailToSend { get; set; }

        [RequiredArgument]
        [Input("Recipient Queue")]
        [ReferenceTarget("queue")]
        public InArgument<EntityReference> RecipientQueue { get; set; }

        [RequiredArgument]
        [Default("false")]
        [Input("Include Owner?")]
        public InArgument<bool> IncludeOwner { get; set; }

        [RequiredArgument]
        [Default("false")]
        [Input("Send Email?")]
        public InArgument<bool> SendEmail { get; set; }

        [Output("Users Added")]
        public OutArgument<int> UsersAdded { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            EntityReference emailToSend = EmailToSend.Get(context);
            EntityReference recipientQueue = RecipientQueue.Get(context);
            bool sendEmail = SendEmail.Get(context);
            bool includeOwner = IncludeOwner.Get(context);

            List<Entity> ccList = new List<Entity>();

            Entity email = RetrieveEmail(localContext.OrganizationService, emailToSend.Id);

            if (email == null)
            {
                UsersAdded.Set(context, 0);
                return;
            }

            //Add any pre-defined recipients specified to the array               
            foreach (Entity activityParty in email.GetAttributeValue<EntityCollection>("cc").Entities)
            {
                ccList.Add(activityParty);
            }

            bool is2011 = Is2011(localContext.OrganizationService);
            EntityCollection users = is2011
                ? GetQueueOwner(localContext.OrganizationService, recipientQueue.Id)
                : GetQueueMembers(localContext.OrganizationService, recipientQueue.Id);

            if (!is2011)
            {
                if (includeOwner)
                    users.Entities.AddRange(GetQueueOwner(localContext.OrganizationService, recipientQueue.Id).Entities);
            }

            ccList = ProcessUsers(users, ccList);

            //Update the email
            email["cc"] = ccList.ToArray();
            localContext.OrganizationService.Update(email);

            //Send
            if (sendEmail)
            {
                SendEmailRequest request = new SendEmailRequest
                {
                    EmailId = emailToSend.Id,
                    TrackingToken = string.Empty,
                    IssueSend = true
                };

                localContext.OrganizationService.Execute(request);
            }

            UsersAdded.Set(context, ccList.Count);
        }

        private static Entity RetrieveEmail(IOrganizationService service, Guid emailId)
        {
            return service.Retrieve("email", emailId, new ColumnSet("cc"));
        }

        private static List<Entity> ProcessUsers(EntityCollection queueMembers, List<Entity> ccList)
        {
            foreach (Entity e in queueMembers.Entities)
            {
                Entity activityParty =
                    new Entity("activityparty")
                    {
                        ["partyid"] = new EntityReference("systemuser", e.Id)
                    };

                if (ccList.Any(t => t.GetAttributeValue<EntityReference>("partyid").Id == e.Id)) continue;

                ccList.Add(activityParty);
            }

            return ccList;
        }

        private static bool Is2011(IOrganizationService service)
        {
            //Check if 2011
            RetrieveVersionRequest request = new RetrieveVersionRequest();
            OrganizationResponse response = service.Execute(request);

            return response.Results["Version"].ToString().StartsWith("5");
        }

        private static EntityCollection GetQueueOwner(IOrganizationService service, Guid queueId)
        {
            //Retrieve the queue owner
            Entity queue = service.Retrieve("queue", queueId, new ColumnSet("ownerid"));

            if (queue == null) return new EntityCollection();

            Entity owner = new Entity("systemuser") { Id = queue.GetAttributeValue<EntityReference>("ownerid").Id };

            EntityCollection ownerCollection = new EntityCollection();
            ownerCollection.Entities.Add(owner);

            return ownerCollection;
        }

        private static EntityCollection GetQueueMembers(IOrganizationService service, Guid queueId)
        {
            //Query for the business unit members
            QueryExpression query = new QueryExpression
            {
                EntityName = "systemuser",
                ColumnSet = new ColumnSet(false),
                LinkEntities =
                {
                    new LinkEntity
                    {
                        LinkFromEntityName = "systemuser",
                        LinkFromAttributeName = "systemuserid",
                        LinkToEntityName = "queuemembership",
                        LinkToAttributeName = "systemuserid",
                        Columns = new ColumnSet("systemuserid"),
                        LinkCriteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression
                                {
                                    AttributeName = "queueid",
                                    Operator = ConditionOperator.Equal,
                                    Values = { queueId }
                                }
                            }
                        }
                    }
                },
                Criteria = new FilterExpression
                {
                    Conditions = {
                        new ConditionExpression
                        {
                            AttributeName = "isdisabled",
                            Operator = ConditionOperator.Equal,
                            Values = { false }
                        }
                    }
                }
            };

            return service.RetrieveMultiple(query);
        }
    }
}