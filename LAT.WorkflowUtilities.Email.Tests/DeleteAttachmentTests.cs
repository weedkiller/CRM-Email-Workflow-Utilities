using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class DeleteAttachmentTests
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
        public void DeleteAttachment_Delete_0_Greater_Than_10000()
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
                ["filesize"] = 5000,
                ["filename"] = "text.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 10000},
                { "DeleteSizeMin", 0 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_1_Greater_Than_10000()
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
                ["filesize"] = 100000,
                ["filename"] = "text.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 10000},
                { "DeleteSizeMin", 0 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_2_Greater_Than_10000()
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
                ["filesize"] = 500000,
                ["filename"] = "text1.docx"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 500000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 10000},
                { "DeleteSizeMin", 0 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_1_Of_2_Greater_Than_10000()
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
                ["filesize"] = 5000,
                ["filename"] = "text1.docx"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 500000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 10000},
                { "DeleteSizeMin", 0 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_0_Less_Than_10000()
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
                ["filesize"] = 50000,
                ["filename"] = "text.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 10000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1 });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_1_Less_Than_10000()
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
                ["filesize"] = 5000,
                ["filename"] = "text.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 10000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException), "Minimum Size Cannot Be Greater Than Maximum Size")]
        public void DeleteAttachment_Invalid_Range()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 100000},
                { "DeleteSizeMin", 200000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_2_Less_Than_10000()
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
                ["filesize"] = 5000,
                ["filename"] = "text1.docx"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 3000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 10000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_1_Of_2_Less_Than_5000()
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
                ["filesize"] = 5000,
                ["filename"] = "text1.docx"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 500000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 5000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_0_Between_3000_And_75000()
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
                ["filesize"] = 5000,
                ["filename"] = "text1.docx"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 50000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 75000},
                { "DeleteSizeMin", 3000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_1_Between_3000_And_75000()
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
                ["filesize"] = 5000,
                ["filename"] = "text1.docx"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 500000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 75000},
                { "DeleteSizeMin", 3000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_0_PDF()
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
                ["filesize"] = 100000,
                ["filename"] = "text.docx"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 5000,
                ["filename"] = "text.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_1_PDF()
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
                ["filesize"] = 100000,
                ["filename"] = "text.pdf"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_2_PDF()
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
                ["filesize"] = 100000,
                ["filename"] = "text1.pdf"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 100000,
                ["filename"] = "text2.pdf"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_1_Of_2_PDF()
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
                ["filesize"] = 100000,
                ["filename"] = "text1.pdf"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 100000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_1_PDF_Less_Than_75000()
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
                ["filesize"] = 5000,
                ["filename"] = "text1.pdf"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 5000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 5000 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_0_PDF_Between_3000_And_75000()
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
                ["filesize"] = 5000,
                ["filename"] = "text1.pdf"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 50000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 75000},
                { "DeleteSizeMin", 3000 },
                { "Extensions", "pdf" },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_2_PDF_DOCX_Multiple_Extensions()
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
                ["filesize"] = 100000,
                ["filename"] = "text1.pdf"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 100000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf,docx" },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }

        [TestMethod]
        public void DeleteAttachment_Delete_0()
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
                ["filesize"] = 100000,
                ["filename"] = "text1.pdf"
            };

            Entity activityMimeAttachment2 = new Entity("activitymimeattachment")
            {
                Id = Guid.NewGuid(),
                ["objectid"] = email.Id,
                ["filesize"] = 100000,
                ["filename"] = "text2.docx"
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailWithAttachments", email.ToEntityReference() },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 0 },
                { "Extensions", null },
                { "AppendNotice", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, activityMimeAttachment1, activityMimeAttachment2 });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DeleteAttachment>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["NumberOfAttachmentsDeleted"]);
        }
    }
}