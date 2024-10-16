"use strict";
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
Object.defineProperty(exports, "__esModule", { value: true });
exports.useTeamsUserCredential = useTeamsUserCredential;
var tslib_1 = require("tslib");
var teamsfx_1 = require("@microsoft/teamsfx");
var useTeams_1 = require("./useTeams");
var useData_1 = require("./useData");
/**
 * Initialize TeamsFx SDK with customized configuration.
 *
 * @param authConfig - custom configuration to override default ones.
 * @returns TeamsContextWithCredential object
 *
 * @public
 */
function useTeamsUserCredential(authConfig) {
    var _a;
    var result = (0, useTeams_1.useTeams)({})[0];
    var _b = (0, useData_1.useData)(function () {
        if (process.env.NODE_ENV === "development") {
            (0, teamsfx_1.setLogLevel)(teamsfx_1.LogLevel.Verbose);
            (0, teamsfx_1.setLogFunction)(function (level, message) {
                console.log(message);
            });
        }
        return Promise.resolve(new teamsfx_1.TeamsUserCredential(authConfig));
    }), data = _b.data, error = _b.error, loading = _b.loading;
    return tslib_1.__assign(tslib_1.__assign({}, result), { teamsUserCredential: data, error: error, loading: loading || ((_a = result.loading) !== null && _a !== void 0 ? _a : true) });
}
