using FakeXrmEasy;
using FakeXrmEasy.FakeMessageExecutors;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class EmailQueueMembersTests
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
        public void EmailQueueMembers_No_Members_Without_Owner_1_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

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

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = new EntityReference("systemuser", Guid.NewGuid())
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientQueue", queue.ToEntityReference() },
                { "IncludeOwner", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityParty, queue });
            var fakeRetrieveVersionRequestExecutor = new FakeRetrieveVersionRequestExecutor(false);
            xrmFakedContext.AddFakeMessageExecutor<RetrieveVersionRequest>(fakeRetrieveVersionRequestExecutor);

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailQueueMembers>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailQueueMembers_No_Members_Without_Owner_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = new EntityReference("systemuser", Guid.NewGuid())
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientQueue", queue.ToEntityReference() },
                { "IncludeOwner", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, queue });
            var fakeRetrieveVersionRequestExecutor = new FakeRetrieveVersionRequestExecutor(false);
            xrmFakedContext.AddFakeMessageExecutor<RetrieveVersionRequest>(fakeRetrieveVersionRequestExecutor);

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailQueueMembers>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailQueueMembers_1_Member_Without_Owner_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = new EntityReference("systemuser", Guid.NewGuid())
            };

            Entity systemUser = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["isdisabled"] = false
            };

            Guid id2 = Guid.NewGuid();
            Entity queueMembership = new Entity("queuemembership")
            {
                Id = id2,
                ["queuemembership"] = id2,
                ["systemuserid"] = systemUser.Id,
                ["queueid"] = queue.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientQueue", queue.ToEntityReference() },
                { "IncludeOwner", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, queue, systemUser, queueMembership });
            var fakeRetrieveVersionRequestExecutor = new FakeRetrieveVersionRequestExecutor(false);
            xrmFakedContext.AddFakeMessageExecutor<RetrieveVersionRequest>(fakeRetrieveVersionRequestExecutor);

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailQueueMembers>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailQueueMembers_1_Member_Without_Owner_2_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Guid idAp1 = Guid.NewGuid();
            Entity activityParty1 = new Entity("activityparty")
            {
                Id = idAp1,
                ["activitypartyid"] = idAp1,
                ["activityid"] = new EntityReference("email", id),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            Guid idAp2 = Guid.NewGuid();
            Entity activityParty2 = new Entity("activityparty")
            {
                Id = idAp2,
                ["activitypartyid"] = idAp2,
                ["activityid"] = new EntityReference("email", id),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection to = new EntityCollection();
            to.Entities.Add(activityParty1);
            to.Entities.Add(activityParty2);
            email["to"] = to;

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = new EntityReference("systemuser", Guid.NewGuid())
            };

            Entity systemUser = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["isdisabled"] = false
            };

            Guid id2 = Guid.NewGuid();
            Entity queueMembership = new Entity("queuemembership")
            {
                Id = id2,
                ["queuemembership"] = id2,
                ["systemuserid"] = systemUser.Id,
                ["queueid"] = queue.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientQueue", queue.ToEntityReference() },
                { "IncludeOwner", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, queue, systemUser, queueMembership, activityParty1, activityParty2 });
            var fakeRetrieveVersionRequestExecutor = new FakeRetrieveVersionRequestExecutor(false);
            xrmFakedContext.AddFakeMessageExecutor<RetrieveVersionRequest>(fakeRetrieveVersionRequestExecutor);

            const int expected = 3;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailQueueMembers>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailQueueMembers_2_Members_Without_Owner_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = new EntityReference("systemuser", Guid.NewGuid())
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["isdisabled"] = false
            };

            Entity systemUser2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["isdisabled"] = false
            };

            Guid id2 = Guid.NewGuid();
            Entity queueMembership1 = new Entity("queuemembership")
            {
                Id = id2,
                ["queuemembership"] = id2,
                ["systemuserid"] = systemUser1.Id,
                ["queueid"] = queue.Id
            };

            Guid id3 = Guid.NewGuid();
            Entity queueMembership2 = new Entity("queuemembership")
            {
                Id = id3,
                ["queuemembership"] = id3,
                ["systemuserid"] = systemUser2.Id,
                ["queueid"] = queue.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientQueue", queue.ToEntityReference() },
                { "IncludeOwner", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, queue, systemUser1, systemUser2, queueMembership1, queueMembership2 });
            var fakeRetrieveVersionRequestExecutor = new FakeRetrieveVersionRequestExecutor(false);
            xrmFakedContext.AddFakeMessageExecutor<RetrieveVersionRequest>(fakeRetrieveVersionRequestExecutor);

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailQueueMembers>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailQueueMembers_2_Members_1_Disabled_Without_Owner_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = new EntityReference("systemuser", Guid.NewGuid())
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["isdisabled"] = false
            };

            Entity systemUser2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["isdisabled"] = true
            };

            Guid id2 = Guid.NewGuid();
            Entity queueMembership1 = new Entity("queuemembership")
            {
                Id = id2,
                ["queuemembership"] = id2,
                ["systemuserid"] = systemUser1.Id,
                ["queueid"] = queue.Id
            };

            Guid id3 = Guid.NewGuid();
            Entity queueMembership2 = new Entity("queuemembership")
            {
                Id = id3,
                ["queuemembership"] = id3,
                ["systemuserid"] = systemUser2.Id,
                ["queueid"] = queue.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientQueue", queue.ToEntityReference() },
                { "IncludeOwner", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, queue, systemUser1, systemUser2, queueMembership1, queueMembership2 });
            var fakeRetrieveVersionRequestExecutor = new FakeRetrieveVersionRequestExecutor(false);
            xrmFakedContext.AddFakeMessageExecutor<RetrieveVersionRequest>(fakeRetrieveVersionRequestExecutor);

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailQueueMembers>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailQueueMembers_1_Member_With_Same_Owner_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

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
                ["isdisabled"] = false
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = systemUser.ToEntityReference()
            };

            Guid id2 = Guid.NewGuid();
            Entity queueMembership = new Entity("queuemembership")
            {
                Id = id2,
                ["queuemembership"] = id2,
                ["systemuserid"] = systemUser.Id,
                ["queueid"] = queue.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientQueue", queue.ToEntityReference() },
                { "IncludeOwner", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, queue, systemUser, queueMembership });
            var fakeRetrieveVersionRequestExecutor = new FakeRetrieveVersionRequestExecutor(false);
            xrmFakedContext.AddFakeMessageExecutor<RetrieveVersionRequest>(fakeRetrieveVersionRequestExecutor);

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailQueueMembers>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailQueueMembers_1_Member_Without_Owner_0_Existing_2011()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = new EntityReference("systemuser", Guid.NewGuid())
            };

            Entity systemUser = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["isdisabled"] = false
            };

            Guid id2 = Guid.NewGuid();
            Entity queueMembership = new Entity("queuemembership")
            {
                Id = id2,
                ["queuemembership"] = id2,
                ["systemuserid"] = systemUser.Id,
                ["queueid"] = queue.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientQueue", queue.ToEntityReference() },
                { "IncludeOwner", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, queue, systemUser, queueMembership });
            var fakeRetrieveVersionRequestExecutor = new FakeRetrieveVersionRequestExecutor(true);
            xrmFakedContext.AddFakeMessageExecutor<RetrieveVersionRequest>(fakeRetrieveVersionRequestExecutor);

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailQueueMembers>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        private class FakeRetrieveVersionRequestExecutor : IFakeMessageExecutor
        {
            private readonly string _version;
            public FakeRetrieveVersionRequestExecutor(bool is2011)
            {
                _version = is2011 ? "5.0.9690.4376" : "8.1.0.512";
            }

            public bool CanExecute(OrganizationRequest request)
            {
                return request is RetrieveVersionRequest;
            }

            public Type GetResponsibleRequestType()
            {
                return typeof(RetrieveVersionRequest);
            }

            public OrganizationResponse Execute(OrganizationRequest request, XrmFakedContext ctx)
            {
                OrganizationResponse response = new OrganizationResponse
                {
                    ResponseName = "RetrieveVersionRequest",
                    Results = new ParameterCollection { { "Version", _version } }
                };

                return response;
            }
        }
    }
}