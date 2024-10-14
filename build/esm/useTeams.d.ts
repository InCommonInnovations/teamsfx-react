import { app } from "@microsoft/teams-js";
/**
 * Microsoft Teams React hook
 * @param options optional options
 * @returns A tuple with properties and methods
 * properties:
 *  - inTeams: boolean = true if inside Microsoft Teams
 *  - fullscreen: boolean = true if in full screen mode
 *  - themeString: string - representation of the theme (default, dark or contrast)
 *  - context - the Microsoft Teams JS SDK context
 * methods:
 *  - setTheme - manually set the theme
 */
export declare function useTeams(options?: {
    initialTheme?: string;
    setThemeHandler?: (theme?: string) => void;
}): [
    {
        inTeams?: boolean;
        fullScreen?: boolean;
        themeString: string;
        context?: app.Context;
        loading?: boolean;
    }
];
//# sourceMappingURL=useTeams.d.ts.map