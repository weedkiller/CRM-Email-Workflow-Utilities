using FakeXrmEasy;
using FakeXrmEasy.FakeMessageExecutors;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class CreateTemplateTests
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
        public void CreateTemplate_Create()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "Template", new EntityReference("template", Guid.NewGuid()) },
                { "RecordUrl", "https://test.crm.dynamics.com:443/main.aspx?etc=1&id=76ff58d2-ae20-e811-a97e-000d3a367d35&histKey=208282481&newWindow=true&pagetype=entityrecord"}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();            var fakeRetrieveMetadataChangesRequest = new FakeRetrieveMetadataChangesRequestExecutor();
            xrmFakedContext.AddFakeMessageExecutor<RetrieveMetadataChangesRequest>(fakeRetrieveMetadataChangesRequest);
            var fakeInstantiateTemplateRequestExecutor = new FakeInstantiateTemplateRequestExecutor();
            xrmFakedContext.AddFakeMessageExecutor<InstantiateTemplateRequest>(fakeInstantiateTemplateRequestExecutor);

            const string expectedTemplateSubject = "Test Subject";
            const string expectedTemplateDescription = "Test Description";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<CreateTemplate>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expectedTemplateSubject, result["TemplateSubject"]);
            Assert.AreEqual(expectedTemplateDescription, result["TemplateDescription"]);
        }

        private class FakeRetrieveMetadataChangesRequestExecutor : IFakeMessageExecutor
        {
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
                EntityMetadata entityMetadata = new EntityMetadata { LogicalName = "account" };
                metadataCollection.Add(entityMetadata);
                ParameterCollection results = new ParameterCollection { { "EntityMetadata", metadataCollection } };
                response.Results = results;

                return response;
            }
        }

        private class FakeInstantiateTemplateRequestExecutor : IFakeMessageExecutor
        {
            public bool CanExecute(OrganizationRequest request)
            {
                return request is InstantiateTemplateRequest;
            }

            public Type GetResponsibleRequestType()
            {
                return typeof(InstantiateTemplateRequest);
            }

            public OrganizationResponse Execute(OrganizationRequest request, XrmFakedContext ctx)
            {
                InstantiateTemplateResponse response =
                    new InstantiateTemplateResponse { Results = new ParameterCollection() };

                Entity email = new Entity("email")
                {
                    ["subject"] = "Test Subject",
                    ["description"] = "Test Description"
                };

                EntityCollection results = new EntityCollection();
                results.Entities.Add(email);
                response.Results.Add("EntityCollection", results);

                return response;
            }
        }
    }
}