"use strict";
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
Object.defineProperty(exports, "__esModule", { value: true });
exports.useTeamsFx = useTeamsFx;
var tslib_1 = require("tslib");
var teamsfx_1 = require("@microsoft/teamsfx");
var useTeams_1 = require("./useTeams");
var useData_1 = require("./useData");
/**
 * Initialize TeamsFx SDK with customized configuration.
 *
 * @param teamsfxConfig - custom configuration to override default ones.
 * @returns TeamsFxContext object
 *
 * @public
 */
function useTeamsFx(teamsfxConfig) {
    var _this = this;
    var result = (0, useTeams_1.useTeams)({})[0];
    var _a = (0, useData_1.useData)(function () { return tslib_1.__awaiter(_this, void 0, void 0, function () {
        return tslib_1.__generator(this, function (_a) {
            if (process.env.NODE_ENV === "development") {
                (0, teamsfx_1.setLogLevel)(teamsfx_1.LogLevel.Verbose);
                (0, teamsfx_1.setLogFunction)(function (level, message) {
                    console.log(message);
                });
            }
            return [2 /*return*/, new teamsfx_1.TeamsFx(teamsfx_1.IdentityType.User, teamsfxConfig)];
        });
    }); }), data = _a.data, error = _a.error, loading = _a.loading;
    return tslib_1.__assign({ teamsfx: data, error: error, loading: loading }, result);
}
