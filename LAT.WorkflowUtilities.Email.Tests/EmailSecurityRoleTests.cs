using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class EmailSecurityRoleTests
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
        public void EmailSecurityRole_1_Role_0_Users_0_Existing()
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

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientRole", Guid.NewGuid().ToString() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email });
            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailSecurityRole>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailSecurityRole_1_Role_1_User_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid roleId = Guid.NewGuid();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["isdisabled"] = false
            };

            Entity systemUserRoles = new Entity("systemuserroles")
            {
                Id = Guid.NewGuid(),
                ["roleid"] = roleId,
                ["systemuserid"] = systemUser1.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientRole", roleId.ToString() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, systemUser1, systemUserRoles });
            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailSecurityRole>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailSecurityRole_1_Role_2_Users_0_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid roleId = Guid.NewGuid();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["isdisabled"] = false
            };

            Entity systemUser2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["isdisabled"] = false
            };

            Entity systemUserRoles1 = new Entity("systemuserroles")
            {
                Id = Guid.NewGuid(),
                ["roleid"] = roleId,
                ["systemuserid"] = systemUser1.Id
            };

            Entity systemUserRoles2 = new Entity("systemuserroles")
            {
                Id = Guid.NewGuid(),
                ["roleid"] = roleId,
                ["systemuserid"] = systemUser2.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientRole", roleId.ToString() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, systemUser1, systemUserRoles1, systemUser2, systemUserRoles2 });
            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailSecurityRole>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailSecurityRole_1_Role_2_Users_1_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid roleId = Guid.NewGuid();

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

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["isdisabled"] = false
            };

            Entity systemUser2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["isdisabled"] = false
            };

            Entity systemUserRoles1 = new Entity("systemuserroles")
            {
                Id = Guid.NewGuid(),
                ["roleid"] = roleId,
                ["systemuserid"] = systemUser1.Id
            };

            Entity systemUserRoles2 = new Entity("systemuserroles")
            {
                Id = Guid.NewGuid(),
                ["roleid"] = roleId,
                ["systemuserid"] = systemUser2.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientRole", roleId.ToString() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, systemUser1, systemUserRoles1, systemUser2, systemUserRoles2, activityParty });
            const int expected = 3;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailSecurityRole>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailSecurityRole_1_Role_2_Users_1_Disabled_1_Existing()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid roleId = Guid.NewGuid();

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

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["isdisabled"] = false
            };

            Entity systemUser2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["isdisabled"] = true
            };

            Entity systemUserRoles1 = new Entity("systemuserroles")
            {
                Id = Guid.NewGuid(),
                ["roleid"] = roleId,
                ["systemuserid"] = systemUser1.Id
            };

            Entity systemUserRoles2 = new Entity("systemuserroles")
            {
                Id = Guid.NewGuid(),
                ["roleid"] = roleId,
                ["systemuserid"] = systemUser2.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientRole", roleId.ToString() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, systemUser1, systemUserRoles1, systemUser2, systemUserRoles2, activityParty });
            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailSecurityRole>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }
    }
}