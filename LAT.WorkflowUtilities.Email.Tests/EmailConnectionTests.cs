using FakeXrmEasy;
using FakeXrmEasy.FakeMessageExecutors;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class EmailConnectionTests
    {
        #region Test Initialization and Cleanup
        // Use ClassInitialize to run code before running the first test in the class
        [ClassInitialize()]
        public static void ClassInitialize(TestContext testContext) { }

        // Use ClassCleanup to run code after all tests in a class have run
        [ClassCleanup()]
        public static void ClassCleanup() { }

        // Use TestInitialize to run code before running each test 
        [TestInitialize()]
        public void TestMethodInitialize() { }

        // Use TestCleanup to run code after each test has run
        [TestCleanup()]
        public void TestMethodCleanup() { }
        #endregion

        [TestMethod]
        public void EmailConnection_1_Account_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext { PrimaryEntityId = Guid.NewGuid() };

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity account = new Entity("account")
            {
                Id = Guid.NewGuid(),
                ["emailaddress1"] = "test@test.com",
                ["statecode"] = new OptionSetValue(0)
            };

            Entity connectionRole = new Entity("connectionrole")
            {
                Id = Guid.NewGuid()
            };

            Entity connection = new Entity("connection")
            {
                Id = Guid.NewGuid(),
                ["record1id"] = new EntityReference("systemuser", workflowContext.PrimaryEntityId),
                ["record1objecttypecode"] = new OptionSetValue(8),
                ["record2id"] = account.ToEntityReference(),
                ["record2objecttypecode"] = new OptionSetValue(1),
                ["record2roleid"] = connectionRole.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "ConnectionRole", connectionRole.ToEntityReference() },
                { "IncludeOppositeConnection", false },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, account, connectionRole, connection });
            var fakeRetrieveMetadataChangesRequest = new FakeRetrieveMetadataChangesRequestExecutor("account");
            xrmFakedContext.AddFakeMessageExecutor<RetrieveMetadataChangesRequest>(fakeRetrieveMetadataChangesRequest);

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailConnection>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailConnection_1_Account_0_Existing_Reverse_Connection()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext { PrimaryEntityId = Guid.NewGuid() };

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity account = new Entity("account")
            {
                Id = Guid.NewGuid(),
                ["emailaddress1"] = "test@test.com",
                ["statecode"] = new OptionSetValue(0)
            };

            Entity connectionRole = new Entity("connectionrole")
            {
                Id = Guid.NewGuid()
            };

            Entity connection = new Entity("connection")
            {
                Id = Guid.NewGuid(),
                ["record2id"] = new EntityReference("systemuser", workflowContext.PrimaryEntityId),
                ["record2objecttypecode"] = new OptionSetValue(8),
                ["record2roleid"] = connectionRole.Id,
                ["record1id"] = account.ToEntityReference(),
                ["record1objecttypecode"] = new OptionSetValue(1)
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "ConnectionRole", connectionRole.ToEntityReference() },
                { "IncludeOppositeConnection", true },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, account, connectionRole, connection });
            var fakeRetrieveMetadataChangesRequest = new FakeRetrieveMetadataChangesRequestExecutor("account");
            xrmFakedContext.AddFakeMessageExecutor<RetrieveMetadataChangesRequest>(fakeRetrieveMetadataChangesRequest);

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailConnection>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailConnection_1_Account_1_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext { PrimaryEntityId = Guid.NewGuid() };

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id
            };

            Guid id2 = Guid.NewGuid();
            Entity activityParty = new Entity("activityparty")
            {
                Id = id2,
                ["activitypartyid"] = id2,
                ["activityid"] = email.ToEntityReference(),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection to = new EntityCollection();
            to.Entities.Add(activityParty);
            email["to"] = to;

            Entity account = new Entity("account")
            {
                Id = Guid.NewGuid(),
                ["emailaddress1"] = "test@test.com",
                ["statecode"] = new OptionSetValue(0)
            };

            Entity connectionRole = new Entity("connectionrole")
            {
                Id = Guid.NewGuid()
            };

            Entity connection = new Entity("connection")
            {
                Id = Guid.NewGuid(),
                ["record1id"] = new EntityReference("systemuser", workflowContext.PrimaryEntityId),
                ["record1objecttypecode"] = new OptionSetValue(8),
                ["record2id"] = account.ToEntityReference(),
                ["record2objecttypecode"] = new OptionSetValue(1),
                ["record2roleid"] = connectionRole.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "ConnectionRole", connectionRole.ToEntityReference() },
                { "IncludeOppositeConnection", false },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, account, connectionRole, connection, activityParty });
            var fakeRetrieveMetadataChangesRequest = new FakeRetrieveMetadataChangesRequestExecutor("account");
            xrmFakedContext.AddFakeMessageExecutor<RetrieveMetadataChangesRequest>(fakeRetrieveMetadataChangesRequest);

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailConnection>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailConnection_1_Account_Without_Email_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext { PrimaryEntityId = Guid.NewGuid() };

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity account = new Entity("account")
            {
                Id = Guid.NewGuid(),
                ["statecode"] = new OptionSetValue(0)
            };

            Entity connectionRole = new Entity("connectionrole")
            {
                Id = Guid.NewGuid()
            };

            Entity connection = new Entity("connection")
            {
                Id = Guid.NewGuid(),
                ["record1id"] = new EntityReference("systemuser", workflowContext.PrimaryEntityId),
                ["record1objecttypecode"] = new OptionSetValue(8),
                ["record2id"] = account.ToEntityReference(),
                ["record2objecttypecode"] = new OptionSetValue(1),
                ["record2roleid"] = connectionRole.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "ConnectionRole", connectionRole.ToEntityReference() },
                { "IncludeOppositeConnection", false },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, account, connectionRole, connection });
            var fakeRetrieveMetadataChangesRequest = new FakeRetrieveMetadataChangesRequestExecutor("account");
            xrmFakedContext.AddFakeMessageExecutor<RetrieveMetadataChangesRequest>(fakeRetrieveMetadataChangesRequest);

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailConnection>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailConnection_1_Account_Inactive_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext { PrimaryEntityId = Guid.NewGuid() };

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity account = new Entity("account")
            {
                Id = Guid.NewGuid(),
                ["emailaddress1"] = "test@test.com",
                ["statecode"] = new OptionSetValue(1)
            };

            Entity connectionRole = new Entity("connectionrole")
            {
                Id = Guid.NewGuid()
            };

            Entity connection = new Entity("connection")
            {
                Id = Guid.NewGuid(),
                ["record1id"] = new EntityReference("systemuser", workflowContext.PrimaryEntityId),
                ["record1objecttypecode"] = new OptionSetValue(8),
                ["record2id"] = account.ToEntityReference(),
                ["record2objecttypecode"] = new OptionSetValue(1),
                ["record2roleid"] = connectionRole.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "ConnectionRole", connectionRole.ToEntityReference() },
                { "IncludeOppositeConnection", false },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, account, connectionRole, connection });
            var fakeRetrieveMetadataChangesRequest = new FakeRetrieveMetadataChangesRequestExecutor("account");
            xrmFakedContext.AddFakeMessageExecutor<RetrieveMetadataChangesRequest>(fakeRetrieveMetadataChangesRequest);

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailConnection>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailConnection_2_Accounts_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext { PrimaryEntityId = Guid.NewGuid() };

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity account1 = new Entity("account")
            {
                Id = Guid.NewGuid(),
                ["emailaddress1"] = "test1@test.com",
                ["statecode"] = new OptionSetValue(0)
            };

            Entity account2 = new Entity("account")
            {
                Id = Guid.NewGuid(),
                ["emailaddress1"] = "test2@test.com",
                ["statecode"] = new OptionSetValue(0)
            };

            Entity connectionRole = new Entity("connectionrole")
            {
                Id = Guid.NewGuid()
            };

            Entity connection1 = new Entity("connection")
            {
                Id = Guid.NewGuid(),
                ["record1id"] = new EntityReference("systemuser", workflowContext.PrimaryEntityId),
                ["record1objecttypecode"] = new OptionSetValue(8),
                ["record2id"] = account1.ToEntityReference(),
                ["record2objecttypecode"] = new OptionSetValue(1),
                ["record2roleid"] = connectionRole.Id
            };

            Entity connection2 = new Entity("connection")
            {
                Id = Guid.NewGuid(),
                ["record1id"] = new EntityReference("systemuser", workflowContext.PrimaryEntityId),
                ["record1objecttypecode"] = new OptionSetValue(8),
                ["record2id"] = account2.ToEntityReference(),
                ["record2objecttypecode"] = new OptionSetValue(1),
                ["record2roleid"] = connectionRole.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "ConnectionRole", connectionRole.ToEntityReference() },
                { "IncludeOppositeConnection", false },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, account1, account2, connectionRole, connection1, connection2 });
            var fakeRetrieveMetadataChangesRequest = new FakeRetrieveMetadataChangesRequestExecutor("account");
            xrmFakedContext.AddFakeMessageExecutor<RetrieveMetadataChangesRequest>(fakeRetrieveMetadataChangesRequest);

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailConnection>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailConnection_1_SystemUser_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext { PrimaryEntityId = Guid.NewGuid() };

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity systemUser = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test@test.com",
                ["isdisabled"] = false
            };

            Entity connectionRole = new Entity("connectionrole")
            {
                Id = Guid.NewGuid()
            };

            Entity connection = new Entity("connection")
            {
                Id = Guid.NewGuid(),
                ["record1id"] = new EntityReference("systemuser", workflowContext.PrimaryEntityId),
                ["record1objecttypecode"] = new OptionSetValue(8),
                ["record2id"] = systemUser.ToEntityReference(),
                ["record2objecttypecode"] = new OptionSetValue(8),
                ["record2roleid"] = connectionRole.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "ConnectionRole", connectionRole.ToEntityReference() },
                { "IncludeOppositeConnection", false },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, systemUser, connectionRole, connection });
            var fakeRetrieveMetadataChangesRequest = new FakeRetrieveMetadataChangesRequestExecutor("systemuser");
            xrmFakedContext.AddFakeMessageExecutor<RetrieveMetadataChangesRequest>(fakeRetrieveMetadataChangesRequest);

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailConnection>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailConnection_1_SystemUser_Without_Email_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext { PrimaryEntityId = Guid.NewGuid() };

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity systemUser = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false
            };

            Entity connectionRole = new Entity("connectionrole")
            {
                Id = Guid.NewGuid()
            };

            Entity connection = new Entity("connection")
            {
                Id = Guid.NewGuid(),
                ["record1id"] = new EntityReference("systemuser", workflowContext.PrimaryEntityId),
                ["record1objecttypecode"] = new OptionSetValue(8),
                ["record2id"] = systemUser.ToEntityReference(),
                ["record2objecttypecode"] = new OptionSetValue(8),
                ["record2roleid"] = connectionRole.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "ConnectionRole", connectionRole.ToEntityReference() },
                { "IncludeOppositeConnection", false },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, systemUser, connectionRole, connection });
            var fakeRetrieveMetadataChangesRequest = new FakeRetrieveMetadataChangesRequestExecutor("systemuser");
            xrmFakedContext.AddFakeMessageExecutor<RetrieveMetadataChangesRequest>(fakeRetrieveMetadataChangesRequest);

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailConnection>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailConnection_1_SystemUser_Disabled_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext { PrimaryEntityId = Guid.NewGuid() };

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity systemUser = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test@test.com",
                ["isdisabled"] = true
            };

            Entity connectionRole = new Entity("connectionrole")
            {
                Id = Guid.NewGuid()
            };

            Entity connection = new Entity("connection")
            {
                Id = Guid.NewGuid(),
                ["record1id"] = new EntityReference("systemuser", workflowContext.PrimaryEntityId),
                ["record1objecttypecode"] = new OptionSetValue(8),
                ["record2id"] = systemUser.ToEntityReference(),
                ["record2objecttypecode"] = new OptionSetValue(8),
                ["record2roleid"] = connectionRole.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "ConnectionRole", connectionRole.ToEntityReference() },
                { "IncludeOppositeConnection", false },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, systemUser, connectionRole, connection });
            var fakeRetrieveMetadataChangesRequest = new FakeRetrieveMetadataChangesRequestExecutor("systemuser");
            xrmFakedContext.AddFakeMessageExecutor<RetrieveMetadataChangesRequest>(fakeRetrieveMetadataChangesRequest);

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailConnection>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        private class FakeRetrieveMetadataChangesRequestExecutor : IFakeMessageExecutor
        {
            readonly string _logicalName;

            public FakeRetrieveMetadataChangesRequestExecutor(string logicalName)
            {
                _logicalName = logicalName;
            }

            public bool CanExecute(OrganizationRequest request)
            {
                return request is RetrieveMetadataChangesRequest;
            }

            public Type GetResponsibleRequestType()
            {
                return typeof(RetrieveMetadataChangesRequest);
            }

            public OrganizationResponse Execute(OrganizationRequest request, XrmFakedContext ctx)
            {
                RetrieveMetadataChangesResponse response = new RetrieveMetadataChangesResponse();

                EntityMetadataCollection metadataCollection = new EntityMetadataCollection();
                EntityMetadata entityMetadata = new EntityMetadata { LogicalName = _logicalName };
                metadataCollection.Add(entityMetadata);
                ParameterCollection results = new ParameterCollection { { "EntityMetadata", metadataCollection } };
                response.Results = results;

                return response;
            }
        }
    }
}