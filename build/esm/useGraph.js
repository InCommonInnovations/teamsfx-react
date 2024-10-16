// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { __assign, __awaiter, __generator } from "tslib";
import { useData } from "./useData";
import { TeamsFx, createMicrosoftGraphClient, ErrorWithCode, TeamsUserCredential, } from "@microsoft/teamsfx";
import { Client, GraphError } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { useState } from "react";
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
export function useGraph(fetchGraphDataAsync, options) {
    var _this = this;
    var _a = __assign({ scope: ["User.Read"], teamsfx: new TeamsFx() }, options), scope = _a.scope, teamsfx = _a.teamsfx;
    var _b = useState(false), needConsent = _b[0], setNeedConsent = _b[1];
    var _c = useData(function () { return __awaiter(_this, void 0, void 0, function () {
        var err_1, helpLink, graph, graphData, err_2;
        var _a, _b;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    if (!needConsent) return [3 /*break*/, 4];
                    _c.label = 1;
                case 1:
                    _c.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, teamsfx.login(scope)];
                case 2:
                    _c.sent();
                    setNeedConsent(false);
                    return [3 /*break*/, 4];
                case 3:
                    err_1 = _c.sent();
                    if (err_1 instanceof ErrorWithCode && ((_a = err_1.message) === null || _a === void 0 ? void 0 : _a.includes("CancelledByUser"))) {
                        helpLink = "https://aka.ms/teamsfx-auth-code-flow";
                        err_1.message +=
                            '\nIf you see "AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application" ' +
                                "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
                                "cli (version < 0.11.0). Please refer to the help link for how to fix the issue: ".concat(helpLink);
                    }
                    throw err_1;
                case 4:
                    _c.trys.push([4, 6, , 7]);
                    graph = createMicrosoftGraphClient(teamsfx, scope);
                    return [4 /*yield*/, fetchGraphDataAsync(graph, teamsfx, scope)];
                case 5:
                    graphData = _c.sent();
                    return [2 /*return*/, graphData];
                case 6:
                    err_2 = _c.sent();
                    if (err_2 instanceof GraphError && ((_b = err_2.code) === null || _b === void 0 ? void 0 : _b.includes("UiRequiredError"))) {
                        // Silently fail for user didn't consent error
                        setNeedConsent(true);
                    }
                    else {
                        throw err_2;
                    }
                    return [3 /*break*/, 7];
                case 7: return [2 /*return*/];
            }
        });
    }); }), data = _c.data, error = _c.error, loading = _c.loading, reload = _c.reload;
    return { data: data, error: error, loading: loading, reload: reload };
}
/**
 * Helper function to call Microsoft Graph API with authentication.
 *
 * @param fetchGraphDataAsync - async function of how to call Graph API and fetch data.
 * @param options - Authentication configuration and OAuth resource scope.
 * @returns data, loading status, error and reload function
 *
 * @public
 */
export function useGraphWithCredential(fetchGraphDataAsync, options) {
    var _this = this;
    var credential;
    if (!(options === null || options === void 0 ? void 0 : options.credential)) {
        var authConfig = {
            clientId: process.env.REACT_APP_CLIENT_ID,
            initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
        };
        credential = new TeamsUserCredential(authConfig);
    }
    else {
        credential = options === null || options === void 0 ? void 0 : options.credential;
    }
    var scope;
    if (!(options === null || options === void 0 ? void 0 : options.scope)) {
        scope = ["User.Read"];
    }
    else {
        scope = options.scope;
    }
    var _a = useState(false), needConsent = _a[0], setNeedConsent = _a[1];
    var _b = useData(function () { return __awaiter(_this, void 0, void 0, function () {
        var err_3, helpLink, authProvider, graph, graphData, err_4;
        var _a, _b;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    if (!needConsent) return [3 /*break*/, 4];
                    _c.label = 1;
                case 1:
                    _c.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, credential.login(scope)];
                case 2:
                    _c.sent();
                    setNeedConsent(false);
                    return [3 /*break*/, 4];
                case 3:
                    err_3 = _c.sent();
                    if (err_3 instanceof ErrorWithCode && ((_a = err_3.message) === null || _a === void 0 ? void 0 : _a.includes("CancelledByUser"))) {
                        helpLink = "https://aka.ms/teamsfx-auth-code-flow";
                        err_3.message +=
                            '\nIf you see "AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application" ' +
                                "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
                                "cli (version < 0.11.0). Please refer to the help link for how to fix the issue: ".concat(helpLink);
                    }
                    throw err_3;
                case 4:
                    _c.trys.push([4, 6, , 7]);
                    authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: scope });
                    graph = Client.initWithMiddleware({
                        authProvider: authProvider,
                    });
                    return [4 /*yield*/, fetchGraphDataAsync(graph, credential, scope)];
                case 5:
                    graphData = _c.sent();
                    return [2 /*return*/, graphData];
                case 6:
                    err_4 = _c.sent();
                    if (err_4 instanceof GraphError && ((_b = err_4.code) === null || _b === void 0 ? void 0 : _b.includes("UiRequiredError"))) {
                        // Silently fail for user didn't consent error
                        setNeedConsent(true);
                    }
                    else {
                        throw err_4;
                    }
                    return [3 /*break*/, 7];
                case 7: return [2 /*return*/];
            }
        });
    }); }), data = _b.data, error = _b.error, loading = _b.loading, reload = _b.reload;
    return { data: data, error: error, loading: loading, reload: reload };
}
