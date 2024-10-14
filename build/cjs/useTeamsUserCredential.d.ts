import { TeamsUserCredentialAuthConfig, TeamsUserCredential } from "@microsoft/teamsfx";
export type TeamsContextWithCredential = {
    /**
     * Instance of TeamsUserCredential.
     */
    teamsUserCredential?: TeamsUserCredential;
    /**
     * Status of data loading.
     */
    loading: boolean;
    /**
     * Error information.
     */
    error: unknown;
    /**
     * Indicates that current environment is in Teams
     */
    inTeams?: boolean;
    /**
     * Teams theme string.
     */
    themeString: string;
    /**
     * Teams context object.
     */
    context?: any;
};
/**
 * Initialize TeamsFx SDK with customized configuration.
 *
 * @param authConfig - custom configuration to override default ones.
 * @returns TeamsContextWithCredential object
 *
 * @public
 */
export declare function useTeamsUserCredential(authConfig: TeamsUserCredentialAuthConfig): TeamsContextWithCredential;
//# sourceMappingURL=useTeamsUserCredential.d.ts.map