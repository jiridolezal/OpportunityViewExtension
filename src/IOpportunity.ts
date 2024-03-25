export interface IOpportunity {
    sfaLeadId: string,
    sfaCustomer: string,
    sfaLeadName: string,
    sfaRfpDay: string,
    sfaSalerStringId: string,
    sfaBidManagerStringId: string,
    sfaGarantStringId: string,
    sfaLegalStringId: string,
    sfaGoNoGo: string,
    sfaGenChannel: string,
    sfaOpportunityId: string,
    sfaTeamId: string,
    sfaReasonOfLost: string,
    sfaOpportunityPhase: string,
    sfaTeamDone: boolean | null
}