using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using QSP.Plugins.Constants;
using System;

namespace QSP.Plugins.Proposal
{
    internal class BusinessUnitLogoService
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _tracing;

        internal BusinessUnitLogoService(
            IOrganizationService service,
            ITracingService tracing)
        {
            _service = service;
            _tracing = tracing;
        }

        internal byte[] GetLogo(Guid businessUnitId)
        {
            _tracing.Trace(
                $"Retrieving proposal logo for Business Unit: {businessUnitId}");

            var businessUnit = _service.Retrieve(
                QspProposalConstants.BusinessUnitTable,
                businessUnitId,
                new ColumnSet(
                    QspProposalConstants.BusinessUnitName,
                    QspProposalConstants.ProposalLogo
                )
            );

            var businessUnitName =
                businessUnit.GetAttributeValue<string>(
                    QspProposalConstants.BusinessUnitName);

            _tracing.Trace(
                $"Business Unit: {businessUnitName}");

            var logo =
                businessUnit.GetAttributeValue<byte[]>(
                    QspProposalConstants.ProposalLogo);

            if (logo == null || logo.Length == 0)
            {
                throw new InvalidPluginExecutionException(
                    $"Proposal Logo is not configured for Business Unit '{businessUnitName}'.");
            }

            _tracing.Trace(
                $"Proposal logo retrieved. Size: {logo.Length} bytes.");

            return logo;
        }
    }
}