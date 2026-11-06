using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using QSP.Plugins.Constants;
using System;

namespace QSP.Plugins.Proposal
{
    public class GenerateQspProposalPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            var tracing = (ITracingService)
                serviceProvider.GetService(typeof(ITracingService));

            var factory = (IOrganizationServiceFactory)
                serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            var service =
                factory.CreateOrganizationService(context.UserId);

            try
            {
                tracing.Trace(
                    "GenerateQspProposalPlugin started.");

                if (!context.InputParameters.Contains("Target"))
                {
                    throw new InvalidPluginExecutionException(
                        "Target parameter is missing.");
                }

                var target =
                    context.InputParameters["Target"]
                    as EntityReference;

                if (target == null)
                {
                    throw new InvalidPluginExecutionException(
                        "Target must be an EntityReference.");
                }

                if (target.LogicalName !=
                    QspProposalConstants.QuoteTable)
                {
                    throw new InvalidPluginExecutionException(
                        "Invalid QSP Quote target.");
                }

                tracing.Trace(
                    $"Quote Id: {target.Id}");

                var quote = service.Retrieve(
                    QspProposalConstants.QuoteTable,
                    target.Id,
                    new ColumnSet(
                        QspProposalConstants.BusinessUnitLookup
                    )
                );

                tracing.Trace(
                    "Quote retrieved successfully.");

                var businessUnit =
                    quote.GetAttributeValue<EntityReference>(
                        QspProposalConstants.BusinessUnitLookup);

                if (businessUnit == null)
                {
                    throw new InvalidPluginExecutionException(
                        "Business Unit is required to generate the proposal.");
                }
                if (businessUnit.LogicalName != QspProposalConstants.BusinessUnitTable)
                {
                    throw new InvalidPluginExecutionException(
                        "The Quote Business Unit lookup contains an unexpected record type.");
                }

                tracing.Trace(
                    $"Business Unit Id: {businessUnit.Id}");

                var logoService =
                    new BusinessUnitLogoService(
                        service,
                        tracing);

                var logoBytes =
                    logoService.GetLogo(
                        businessUnit.Id);

                tracing.Trace(
                    $"Business Unit logo ready. Size: {logoBytes.Length} bytes.");

                var templateService =
                    new DocumentTemplateService(
                        service,
                        tracing);

                var templateBytes =
                    templateService.GetTemplate(
                        QspProposalConstants.TemplateName);

                var wordService = new WordDocumentService(tracing);

                var documentBytes = wordService.Generate(templateBytes,logoBytes);

                tracing.Trace(
                    $"Final document generated. Size: {documentBytes.Length} bytes.");

                tracing.Trace(
                    $"Word template ready. Size: {templateBytes.Length} bytes.");

                tracing.Trace(
                    "GenerateQspProposalPlugin completed successfully.");
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                tracing.Trace(
                    $"GenerateQspProposalPlugin Error: {ex}");

                throw new InvalidPluginExecutionException(
                    $"Unable to generate QSP proposal: {ex.Message}",
                    ex);
            }
        }
    }
}