using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class CheckAttachmentsTests
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
        public void CheckAttachments_Zero_Attachments()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToCheck", new EntityReference("email", email.Id) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email });

            const int expectedAttachmentCount = 0;
            const bool expectedHasAttachments = false;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<CheckAttachments>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expectedAttachmentCount, result["AttachmentCount"]);
            Assert.AreEqual(expectedHasAttachments, result["HasAttachments"]);
        }

        [TestMethod]
        public void CheckAttachments_1_Attachment()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id
            };
            Entity activityMimeAttachment = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["activityid"] = email.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToCheck", new EntityReference("email", email.Id) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment });

            const int expectedAttachmentCount = 1;
            const bool expectedHasAttachments = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<CheckAttachments>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expectedAttachmentCount, result["AttachmentCount"]);
            Assert.AreEqual(expectedHasAttachments, result["HasAttachments"]);
        }
    }
}