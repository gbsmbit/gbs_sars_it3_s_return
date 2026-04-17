Module SARSIT3Consts

    'Public Enum Months
    '    None
    '    January
    '    February
    '    March
    '    April
    '    May
    '    June
    '    July
    '    August
    '    September
    '    October
    '    November
    '    December
    'End Enum

    'Public Enum TranFlows
    '    Debit
    '    Credit
    'End Enum

    Public Enum ClientFields
        cifId
        clnttype
        invId
        cifRegDate
        cifNO
        Deleted
        agID
        cifInslDate
        GetCalendar
        cifBadCredit
        CreatedBy
        ModifiedBy
        ModifiedDate
        CreatedDate
        GBSClientInd
        Comment
        ScreeningExemptInd
        ReviewInd
        LastReviewDate
        ClientRelationshipID
        RiskRatingID
        UNMismatchInd
        ExcludeDOConsolidateInd
        DOConsolidateInd
        UNComplianceStatusID
        cifFICAID
        DueForReviewDate
        SourceOfIncome
        OFACComplianceStatusID
        HMTComplianceStatusID
        ClientCreditRiskID
        AccountCommPrefID
        GeneralCommPrefID
        LPAFirmCode
        DateExtractedCIF
        TFSComplianceStatusID
        cifIdIndividual
        ttId
        lngId
        ctyIdIndividual
        occId
        rId
        mId
        cifiIDNO
        cifiFirstName
        cifiSurname
        cifiInitials
        cifiAltName
        cifiStaffMember
        cifiTaxNO
        cifiDOB
        cifiForeignIDNO
        cifiPassportNO
        cifiTelHome
        cifiTelWork
        cifiTelAlternate
        cifiCell
        cifiFaxHome
        cifiFaxWork
        cifiEmail
        cifiWebsite
        cifIIncome
        cifIdPhoto
        gndId
        DeletedIndividual
        auditIdIndividual
        PEPInd
        PEPComplianceStatusID
        cifiAltEmail
        cifiNotes
        siIdIndividual
        swIdIndividual
        sicIdIndividual
        DateExtractedCIFIndividual
        cifIdCompany
        secId
        companytype
        cifcRegNO
        cifcVATNO
        cifcName
        cifcTradingAs
        cifcTaxNO
        ctyIdCompany
        cifcTel1
        cifcTel2
        cifcTel3
        cifcCell
        cifcFax1
        cifcFax2
        cifcEmail
        cifcWebsite
        auditIdCompany
        DeletedCompany
        cifcAddressee
        cifcContactPerson
        cifcAltEmail
        cifcNotes
        dId
        nbId
        oId
        sicIdCompany
        swIdCompany
        siIdCompany
        cifcObjectives
        DateExtractedCIFCompany
    End Enum

    Public Enum AccountFields
        taxCIF
        taxPaid
        taxAccrued
        taxAccountNo
        atId
        accRegDate
        asId
        accId
    End Enum

    Public Enum AddressFields
        Status
        Line1
        Line2
        Line3
        Line4
        Suburb
        City
        PostCode
        Province
    End Enum

End Module
