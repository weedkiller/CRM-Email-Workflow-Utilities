using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class EmailBusinessUnitTests
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
        public void EmailBusinessUnit_No_Users_Business_Unit_With_No_Existing_Users()
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

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", new EntityReference("businessunit", Guid.NewGuid()) },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnit>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnit_No_Users_Business_Unit_With_1_Existing_User()
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

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity user1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, user1 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnit>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnit_2_Users_Business_Unit_With_No_Existing_Users()
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

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity user1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity user2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, user1, user2 });

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnit>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnit_2_Users_Business_Unit_1_Invalid_AccessMode_With_No_Existing_Users()
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

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity user1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity user2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 5,
                ["isdisabled"] = false
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, user1, user2 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnit>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnit_2_Users_Business_Unit_1_Disabled_With_No_Existing_Users()
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

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity user1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity user2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = true
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, user1, user2 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnit>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnit_2_Users_Business_Unit_With_1_Existing_User()
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

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity user1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity user2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityParty, businessUnit, user1, user2 });

            const int expected = 3;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnit>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnit_2_Users_Business_Unit_1_User_Child_Business_Unit_1_User_With_1_Existing_User()
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

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity user1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity user2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity childBusinessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid(),
                ["parentbusinessunitid"] = businessUnit.ToEntityReference()
            };

            Entity user3 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test3@test.com",
                ["businessunitid"] = childBusinessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "IncludeChildren", true},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityParty, businessUnit, user1, user2, childBusinessUnit, user3 });

            const int expected = 4;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnit>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnit_2_Users_Business_Unit_1_User_Child_Business_Unit_2_Users_With_1_Existing_User()
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

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity user1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity user2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity childBusinessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid(),
                ["parentbusinessunitid"] = businessUnit.ToEntityReference()
            };

            Entity user3 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test3@test.com",
                ["businessunitid"] = childBusinessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity user4 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test4@test.com",
                ["businessunitid"] = childBusinessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "IncludeChildren", true},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityParty, businessUnit, user1, user2, childBusinessUnit, user3, user4 });

            const int expected = 5;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnit>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnit_2_Users_Business_Unit_1_User_Child_Business_Unit_2_Users__Empty_Child_Business_Unit_With_1_Existing_User()
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

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity user1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test1@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity user2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test2@test.com",
                ["businessunitid"] = businessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity childBusinessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid(),
                ["parentbusinessunitid"] = businessUnit.ToEntityReference()
            };

            Entity user3 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test3@test.com",
                ["businessunitid"] = childBusinessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity user4 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = "test4@test.com",
                ["businessunitid"] = childBusinessUnit.Id,
                ["accessmode"] = 1,
                ["isdisabled"] = false
            };

            Entity childBusinessUnit2 = new Entity("businessunit")
            {
                Id = Guid.NewGuid(),
                ["parentbusinessunitid"] = businessUnit.ToEntityReference()
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "IncludeChildren", true},
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityParty, businessUnit, user1, user2, childBusinessUnit, user3, user4, childBusinessUnit2 });

            const int expected = 5;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnit>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }
    }
}