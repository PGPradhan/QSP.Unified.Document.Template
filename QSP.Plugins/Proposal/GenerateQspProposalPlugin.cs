using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using QSP.Plugins.Constants;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace QSP.Plugins.Proposal
{
    public class GenerateQspProposalPlugin : IPlugin
    {
        public void Execute(
            IServiceProvider serviceProvider)
        {
            var context =
                (IPluginExecutionContext)
                serviceProvider.GetService(
                    typeof(IPluginExecutionContext));

            var tracing =
                (ITracingService)
                serviceProvider.GetService(
                    typeof(ITracingService));

            var factory =
                (IOrganizationServiceFactory)
                serviceProvider.GetService(
                    typeof(IOrganizationServiceFactory));

            var service =
                factory.CreateOrganizationService(
                    context.UserId);

            try
            {
                tracing.Trace(
                    "GenerateQspProposalPlugin started.");

                if (!context.InputParameters.Contains(
                    "Target"))
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

                var quote =
                    service.Retrieve(
                        QspProposalConstants.QuoteTable,
                        target.Id,
                        new ColumnSet(
                            QspProposalConstants.BusinessUnitLookup,
                            QspProposalConstants.CustomerLookup,
                            QspProposalConstants.ActualQuote,
                            QspProposalConstants.OpportunityName,
                            QspProposalConstants.QuoteNumber));

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

                if (businessUnit.LogicalName !=
                    QspProposalConstants.BusinessUnitTable)
                {
                    throw new InvalidPluginExecutionException(
                        "The Quote Business Unit lookup contains an unexpected record type.");
                }

                var customerReference =
                    quote.GetAttributeValue<EntityReference>(
                        QspProposalConstants.CustomerLookup);

                if (customerReference == null)
                {
                    throw new InvalidPluginExecutionException(
                        "QSP Customer is required to generate the proposal.");
                }

                if (customerReference.LogicalName !=
                    QspProposalConstants.CustomerTable)
                {
                    throw new InvalidPluginExecutionException(
                        "The QSP Customer lookup contains an unexpected record type.");
                }

                tracing.Trace(
                    $"Business Unit Id: {businessUnit.Id}");

                tracing.Trace(
                    $"QSP Customer Id: {customerReference.Id}");

                var customer =
                    service.Retrieve(
                        QspProposalConstants.CustomerTable,
                        customerReference.Id,
                        new ColumnSet(
                            QspProposalConstants.CustomerName,
                            QspProposalConstants.CustomerContactName,
                            QspProposalConstants.CustomerAddressLine1,
                            QspProposalConstants.CustomerAddressCounty,
                            QspProposalConstants.CustomerAddressCountry));

                tracing.Trace(
                    "QSP Customer retrieved successfully.");

                var logoService =
                    new BusinessUnitLogoService(
                        service,
                        tracing);

                var logoBytes =
                    logoService.GetLogo(
                        businessUnit.Id);

                tracing.Trace(
                    $"Business Unit logo ready. Size: {logoBytes.Length} bytes.");

                var proposal =
                    BuildProposalData(
                        quote,
                        customer,
                        logoBytes,
                        tracing);

                var templateService =
                    new DocumentTemplateService(
                        service,
                        tracing);

                var templateBytes =
                    templateService.GetTemplate(
                        QspProposalConstants.TemplateName);

                tracing.Trace(
                    $"Word template ready. Size: {templateBytes.Length} bytes.");

                var wordService =
                    new WordDocumentService(
                        tracing);

                var documentBytes =
                    wordService.Generate(
                        templateBytes,
                        proposal);

                tracing.Trace(
                    $"Final document generated. Size: {documentBytes.Length} bytes.");

                var fileName =
                    $"QSP-Proposal-{target.Id}.docx";

                var documentBase64 =
                    Convert.ToBase64String(
                        documentBytes);

                context.OutputParameters["pp_FileName"] =
                    fileName;

                context.OutputParameters["pp_Document"] =
                    documentBase64;

                tracing.Trace(
                    $"Proposal output prepared. File: {fileName}");

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

        private static ProposalResult BuildProposalData(
            Entity quote,
            Entity customer,
            byte[] logoBytes,
            ITracingService tracing)
        {
            var actualQuote =
                quote.GetAttributeValue<Money>(
                    QspProposalConstants.ActualQuote);

            var proposal =
                new ProposalResult
                {
                    ProposalDate =
                        DateTime.UtcNow.ToString(
                            "d-MMM-yyyy",
                            CultureInfo.InvariantCulture),

                    QuoteNumber = quote.GetAttributeValue<string>(
                            QspProposalConstants.QuoteNumber) ?? string.Empty,

                    ClientName =
                        customer.GetAttributeValue<string>(
                            QspProposalConstants.CustomerName)
                        ?? string.Empty,

                    ContactName =
                        customer.GetAttributeValue<string>(
                            QspProposalConstants.CustomerContactName)
                        ?? string.Empty,

                    ProjectName =
                        quote.GetAttributeValue<string>(
                            QspProposalConstants.OpportunityName)
                        ?? string.Empty,

                    SiteAddress =
                        BuildAddress(
                            customer),

                    SellPrice =
                        actualQuote != null
                            ? actualQuote.Value.ToString(
                                "N2",
                                CultureInfo.InvariantCulture)
                            : string.Empty,

                    LogoBytes =
                        logoBytes
                };

            tracing.Trace(
                $"Proposal Date: {proposal.ProposalDate}");

            tracing.Trace(
                $"Client: {proposal.ClientName}");

            tracing.Trace(
                $"Contact: {proposal.ContactName}");

            tracing.Trace(
                $"Project: {proposal.ProjectName}");

            tracing.Trace(
                $"Site Address: {proposal.SiteAddress}");

            tracing.Trace(
                $"Sell Price: {proposal.SellPrice}");

            return proposal;
        }

        private static string BuildAddress(
            Entity customer)
        {
            var addressParts =
                new List<string>
                {
                    customer.GetAttributeValue<string>(
                        QspProposalConstants.CustomerAddressLine1),

                    customer.GetAttributeValue<string>(
                        QspProposalConstants.CustomerAddressCounty),

                    customer.GetAttributeValue<string>(
                        QspProposalConstants.CustomerAddressCountry)
                };

            return string.Join(
                ", ",
                addressParts
                    .Where(
                        value =>
                            !string.IsNullOrWhiteSpace(
                                value))
                    .Select(
                        value =>
                            value.Trim()));
        }
    }
}