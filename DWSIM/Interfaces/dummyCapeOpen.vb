
Namespace CAPEOPEN110
    Public Enum eCapePhaseStatus
        CAPE_UNKNOWNPHASESTATUS = 0
        CAPE_ATEQUILIBRIUM = 1
        CAPE_ESTIMATES = 2
    End Enum
    Public Enum eCapeCalculationCode
        CAPE_NO_CALCULATION = 0
        CAPE_LOG_FUGACITY_COEFFICIENTS = 1
        CAPE_T_DERIVATIVE = 2
        CAPE_P_DERIVATIVE = 4
        CAPE_MOLE_NUMBERS_DERIVATIVES = 8
    End Enum
End Namespace

