using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class EmailTeamTests
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
        public void EmailTeam_1_Team_1_User_1_Existing()
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
                ["activityid"] = new EntityReference("email", id),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection to = new EntityCollection();
            to.Entities.Add(activityParty);
            email["to"] = to;

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid()
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["isdisabled"] = false
            };

            Entity teammembership = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientTeam", team.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityParty, team, systemUser1, teammembership });
            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailTeam>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailTeam_1_Team_1_User_Without_Email_1_Existing()
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
                ["activityid"] = new EntityReference("email", id),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection to = new EntityCollection();
            to.Entities.Add(activityParty);
            email["to"] = to;

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid()
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false
            };

            Entity teammembership = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientTeam", team.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityParty, team, systemUser1, teammembership });
            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailTeam>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailTeam_1_Team_2_Users_1_Existing()
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
                ["activityid"] = new EntityReference("email", id),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection to = new EntityCollection();
            to.Entities.Add(activityParty);
            email["to"] = to;

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid()
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["isdisabled"] = false
            };

            Entity teammembership1 = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity systemUser2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["isdisabled"] = false
            };

            Entity teammembership2 = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser2.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientTeam", team.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityParty, team, systemUser1, teammembership1, systemUser2, teammembership2 });
            const int expected = 3;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailTeam>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailTeam_1_Team_2_Users_1_Disabled_1_Existing()
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
                ["activityid"] = new EntityReference("email", id),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection to = new EntityCollection();
            to.Entities.Add(activityParty);
            email["to"] = to;

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid()
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["isdisabled"] = false
            };

            Entity teammembership1 = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity systemUser2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["isdisabled"] = true
            };

            Entity teammembership2 = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser2.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientTeam", team.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityParty, team, systemUser1, teammembership1, systemUser2, teammembership2 });
            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailTeam>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }
    }
}