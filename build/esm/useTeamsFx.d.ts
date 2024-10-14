import { TeamsFx } from "@microsoft/teamsfx";
export type TeamsFxContext = {
    /**
     * Instance of TeamsFx.
     */
    teamsfx?: TeamsFx;
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
 * @param teamsfxConfig - custom configuration to override default ones.
 * @returns TeamsFxContext object
 *
 * @public
 */
export declare function useTeamsFx(teamsfxConfig?: Record<string, string>): TeamsFxContext;
//# sourceMappingURL=useTeamsFx.d.ts.map