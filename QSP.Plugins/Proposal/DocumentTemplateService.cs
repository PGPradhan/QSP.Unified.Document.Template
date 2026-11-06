using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace QSP.Plugins.Proposal
{
    internal class DocumentTemplateService
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _tracing;

        internal DocumentTemplateService(
            IOrganizationService service,
            ITracingService tracing)
        {
            _service = service;
            _tracing = tracing;
        }

        internal byte[] GetTemplate(string templateName)
        {
            _tracing.Trace(
                $"Retrieving Document Template: {templateName}");

            var query = new QueryExpression("documenttemplate")
            {
                ColumnSet = new ColumnSet(
                    "documenttemplateid",
                    "name",
                    "content",
                    "associatedentitytypecode"
                ),
                TopCount = 2
            };

            query.Criteria.AddCondition(
                "name",
                ConditionOperator.Equal,
                templateName);

            var templates =
                _service.RetrieveMultiple(query).Entities;

            if (templates.Count == 0)
            {
                throw new InvalidPluginExecutionException(
                    $"Document Template '{templateName}' was not found.");
            }

            if (templates.Count > 1)
            {
                throw new InvalidPluginExecutionException(
                    $"Multiple Document Templates named '{templateName}' were found.");
            }

            var template = templates.First();

            var associatedEntity =
                template.GetAttributeValue<string>(
                    "associatedentitytypecode");

            _tracing.Trace(
                $"Template associated entity: {associatedEntity}");

            var content =
                template.GetAttributeValue<string>("content");

            if (string.IsNullOrWhiteSpace(content))
            {
                throw new InvalidPluginExecutionException(
                    "The QSP Document Template does not contain document content.");
            }

            byte[] templateBytes;

            try
            {
                templateBytes =
                    Convert.FromBase64String(content);
            }
            catch (FormatException)
            {
                throw new InvalidPluginExecutionException(
                    "The QSP Document Template content is not valid Base64.");
            }

            if (templateBytes.Length == 0)
            {
                throw new InvalidPluginExecutionException(
                    "The QSP Document Template is empty.");
            }

            _tracing.Trace(
                $"Document Template retrieved. Size: {templateBytes.Length} bytes.");

            return templateBytes;
        }
    }
}