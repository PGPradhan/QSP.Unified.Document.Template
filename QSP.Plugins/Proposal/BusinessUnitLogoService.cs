using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using QSP.Plugins.Constants;
using System;
using System.IO;

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
                    QspProposalConstants.BusinessUnitName));

            var businessUnitName =
                businessUnit.GetAttributeValue<string>(
                    QspProposalConstants.BusinessUnitName);

            _tracing.Trace(
                $"Business Unit: {businessUnitName}");

            try
            {
                var logo =
                    DownloadFullSizeLogo(
                        businessUnitId);

                if (logo == null ||
                    logo.Length == 0)
                {
                    throw new InvalidPluginExecutionException(
                        $"Proposal Logo is not configured for Business Unit '{businessUnitName}'.");
                }

                _tracing.Trace(
                    $"Full-size proposal logo retrieved. Size: {logo.Length} bytes.");

                return logo;
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _tracing.Trace(
                    $"Full-size logo download failed: {ex}");

                throw new InvalidPluginExecutionException(
                    $"Unable to retrieve Proposal Logo for Business Unit '{businessUnitName}': {ex.Message}",
                    ex);
            }
        }

        private byte[] DownloadFullSizeLogo(
            Guid businessUnitId)
        {
            var initializeRequest =
                new InitializeFileBlocksDownloadRequest
                {
                    Target =
                        new EntityReference(
                            QspProposalConstants.BusinessUnitTable,
                            businessUnitId),

                    FileAttributeName =
                        QspProposalConstants.ProposalLogo
                };

            var initializeResponse =
                (InitializeFileBlocksDownloadResponse)
                _service.Execute(
                    initializeRequest);

            if (string.IsNullOrWhiteSpace(
                initializeResponse.FileContinuationToken))
            {
                throw new InvalidPluginExecutionException(
                    "Unable to initialize Proposal Logo download.");
            }

            _tracing.Trace(
                $"Proposal logo download initialized. File size: {initializeResponse.FileSizeInBytes} bytes.");

            const long blockSize =
                4 * 1024 * 1024;

            long offset = 0;

            using (var output =
                new MemoryStream())
            {
                while (offset <
                    initializeResponse.FileSizeInBytes)
                {
                    var remaining =
                        initializeResponse.FileSizeInBytes -
                        offset;

                    var requestedSize =
                        Math.Min(
                            blockSize,
                            remaining);

                    var downloadRequest =
                        new DownloadBlockRequest
                        {
                            FileContinuationToken =
                                initializeResponse.FileContinuationToken,

                            Offset =
                                offset,

                            BlockLength =
                                requestedSize
                        };

                    var downloadResponse =
                        (DownloadBlockResponse)
                        _service.Execute(
                            downloadRequest);

                    if (downloadResponse.Data == null ||
                        downloadResponse.Data.Length == 0)
                    {
                        break;
                    }

                    output.Write(
                        downloadResponse.Data,
                        0,
                        downloadResponse.Data.Length);

                    offset +=
                        downloadResponse.Data.Length;
                }

                var logoBytes =
                    output.ToArray();

                _tracing.Trace(
                    $"Proposal logo download completed. Downloaded: {logoBytes.Length} bytes.");

                return logoBytes;
            }
        }
    }
}