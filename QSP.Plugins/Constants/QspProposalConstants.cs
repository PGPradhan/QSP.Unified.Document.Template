namespace QSP.Plugins.Constants
{
    internal static class QspProposalConstants
    {
        // Quote
        internal const string QuoteTable = "pp_qspquotepricing01";
        internal const string QuoteId = "pp_qspquotepricing01id";

        // Quote fields
        internal const string Client = "pp_name";
        internal const string Contact = "pp_qspcontactname";
        internal const string ProjectName = "pp_opportunityname";
        internal const string BusinessUnitLookup = "pp_businessunit";
        internal const string AddressLine1 = "pp_addressline1";
        internal const string AddressCounty = "pp_addresscounty";
        internal const string AddressCountry = "pp_addresscountry";
        internal const string SellPrice = "pp_actualquote";

        // Business Unit
        internal const string BusinessUnitTable = "pp_businessunit";
        internal const string BusinessUnitId = "pp_businessunitid";
        internal const string BusinessUnitName = "pp_name";
        internal const string ProposalLogo = "pp_proposallogo";

        // Word controls
        internal const string LogoTag = "QSP_BU_LOGO";
        internal const string ProposalDateTag = "QSP_PROPOSAL_DATE";

        // Word template
        internal const string TemplateName =
            "QSP Quote Pricing 01 2026-11-5 12-40-8";
    }
}