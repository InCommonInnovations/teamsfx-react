// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { __assign, __awaiter, __generator } from "tslib";
import { LogLevel, setLogLevel, setLogFunction, TeamsFx, IdentityType } from "@microsoft/teamsfx";
import { useTeams } from "./useTeams";
import { useData } from "./useData";
/**
 * Initialize TeamsFx SDK with customized configuration.
 *
 * @param teamsfxConfig - custom configuration to override default ones.
 * @returns TeamsFxContext object
 *
 * @public
 */
export function useTeamsFx(teamsfxConfig) {
    var _this = this;
    var result = useTeams({})[0];
    var _a = useData(function () { return __awaiter(_this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            if (process.env.NODE_ENV === "development") {
                setLogLevel(LogLevel.Verbose);
                setLogFunction(function (level, message) {
                    console.log(message);
                });
            }
            return [2 /*return*/, new TeamsFx(IdentityType.User, teamsfxConfig)];
        });
    }); }), data = _a.data, error = _a.error, loading = _a.loading;
    return __assign({ teamsfx: data, error: error, loading: loading }, result);
}
