using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class DeleteAttachmentByNameTests
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
        public void DeleteAttachmentByName_No_Match()
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
                //["activityid"] = id,
                ["objectid"] = email.Id,
                ["filename"] = "text.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "FileName", "test.txt"},
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { activityMimeAttachment });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachmentByName>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachmentByName_1_Match()
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
                ["objectid"] = email.Id,
                ["filename"] = "test.txt"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "FileName", "test.txt"},
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachmentByName>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachmentByName_2_Matches()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id
            };

            Entity activityMimeAttachment1 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filename"] = "test.txt"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filename"] = "test.txt"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "FileName", "test.txt"},
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachmentByName>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachmentByName_1_Of2_Match()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id
            };

            Entity activityMimeAttachment1 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filename"] = "test.txt"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filename"] = "test.docs"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "FileName", "test.txt"},
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachmentByName>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }
    }
}