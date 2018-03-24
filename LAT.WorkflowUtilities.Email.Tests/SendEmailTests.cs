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
    public class SendEmailTests
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
        public void SendEmail_Send()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            var fakeSendEmailRequestExecutor = new FakeSendEmailRequestExecutor();
            xrmFakedContext.AddFakeMessageExecutor<SendEmailRequest>(fakeSendEmailRequestExecutor);

            const bool expected = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<SendEmail>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["EmailSent"]);
        }

        private class FakeSendEmailRequestExecutor : IFakeMessageExecutor
        {
            public bool CanExecute(OrganizationRequest request)
            {
                return request is SendEmailRequest;
            }

            public Type GetResponsibleRequestType()
            {
                return typeof(SendEmailRequest);
            }

            public OrganizationResponse Execute(OrganizationRequest request, XrmFakedContext ctx)
            {
                SendEmailResponse sendResponse = new SendEmailResponse();

                return sendResponse;
            }
        }
    }
}