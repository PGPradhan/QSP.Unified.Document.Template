var QSP = QSP || {};

QSP.ProposalRibbon = {
    generateProposal: async function (primaryControl) {
        try {
            var formContext = primaryControl;

            if (!formContext) {
                throw new Error("Form context is not available.");
            }

            var quoteId = formContext.data.entity.getId();

            if (!quoteId) {
                await Xrm.Navigation.openAlertDialog({
                    title: "Generate Proposal",
                    text: "Please save the Quote before generating the proposal."
                });
                return;
            }

            quoteId = quoteId.replace(/[{}]/g, "");

            Xrm.Utility.showProgressIndicator("Generating proposal...");

            var result = await QSP.ProposalRibbon.callGenerateProposalApi(quoteId);

            if (!result || !result.pp_Document) {
                throw new Error("The proposal document was not returned.");
            }

            var fileName = result.pp_FileName || "QSP-Proposal.docx";

            QSP.ProposalRibbon.downloadDocument(
                result.pp_Document,
                fileName
            );
        }
        catch (error) {
            console.error("Generate Proposal Error:", error);

            await Xrm.Navigation.openErrorDialog({
                message: QSP.ProposalRibbon.getErrorMessage(error)
            });
        }
        finally {
            Xrm.Utility.closeProgressIndicator();
        }
    },

    callGenerateProposalApi: async function (quoteId) {
        var request = {
            entity: {
                entityType: "pp_qspquotepricing01",
                id: quoteId
            },

            getMetadata: function () {
                return {
                    boundParameter: "entity",

                    parameterTypes: {
                        entity: {
                            typeName: "mscrm.pp_qspquotepricing01",
                            structuralProperty: 5
                        }
                    },

                    operationType: 0,
                    operationName: "pp_GenerateQSPProposal"
                };
            }
        };

        var response = await Xrm.WebApi.online.execute(request);

        if (!response.ok) {
            var errorText = await response.text();

            throw new Error(
                errorText || "Proposal generation failed."
            );
        }

        return await response.json();
    },

    downloadDocument: function (base64Document, fileName) {
        var byteCharacters = atob(base64Document);
        var byteNumbers = new Uint8Array(byteCharacters.length);

        for (var i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }

        var blob = new Blob(
            [byteNumbers],
            {
                type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
        );

        var url = window.URL.createObjectURL(blob);
        var link = document.createElement("a");

        link.href = url;
        link.download = fileName || "QSP-Proposal.docx";
        link.style.display = "none";

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        setTimeout(function () {
            window.URL.revokeObjectURL(url);
        }, 1000);
    },

    getErrorMessage: function (error) {
        if (!error) {
            return "An unexpected error occurred while generating the proposal.";
        }

        if (error.message) {
            try {
                var parsed = JSON.parse(error.message);

                if (parsed.error && parsed.error.message) {
                    return parsed.error.message;
                }
            }
            catch (e) {
            }

            return error.message;
        }

        return "An unexpected error occurred while generating the proposal.";
    }
};