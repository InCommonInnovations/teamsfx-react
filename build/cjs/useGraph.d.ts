import { Data } from "./useData";
import { TeamsFx, TeamsUserCredential } from "@microsoft/teamsfx";
import { Client } from "@microsoft/microsoft-graph-client";
type GraphOption = {
    scope?: string[];
    teamsfx?: TeamsFx;
};
type GraphOptionWithCredential = {
    scope?: string[];
    credential?: TeamsUserCredential;
};
/**
 * Helper function to call Microsoft Graph API with authentication.
 * @deprecated Please use {@link useGraphWithCredential} instead.
 *
 * @param fetchGraphDataAsync - async function of how to call Graph API and fetch data.
 * @param options - teamsfx instance and OAuth resource scope.
 * @returns data, loading status, error and reload function
 *
 * @public
 */
export declare function useGraph<T>(fetchGraphDataAsync: (graph: Client, teamsfx: TeamsFx, scope: string[]) => Promise<T>, options?: GraphOption): Data<T>;
/**
 * Helper function to call Microsoft Graph API with authentication.
 *
 * @param fetchGraphDataAsync - async function of how to call Graph API and fetch data.
 * @param options - Authentication configuration and OAuth resource scope.
 * @returns data, loading status, error and reload function
 *
 * @public
 */
export declare function useGraphWithCredential<T>(fetchGraphDataAsync: (graph: Client, credential: TeamsUserCredential, scope: string[]) => Promise<T>, options?: GraphOptionWithCredential): Data<T>;
export {};
//# sourceMappingURL=useGraph.d.ts.map