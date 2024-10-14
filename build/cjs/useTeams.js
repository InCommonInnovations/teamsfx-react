"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.useTeams = useTeams;
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
var react_1 = require("react");
var react_dom_1 = require("react-dom");
var teams_js_1 = require("@microsoft/teams-js");
var getTheme = function () {
    var urlParams = new URLSearchParams(window.location.search);
    var theme = urlParams.get("theme");
    return theme == null ? undefined : theme;
};
/**
 * Microsoft Teams React hook
 * @param options optional options
 * @returns A tuple with properties and methods
 * properties:
 *  - inTeams: boolean = true if inside Microsoft Teams
 *  - fullscreen: boolean = true if in full screen mode
 *  - theme: Fluent UI Theme
 *  - themeString: string - representation of the theme (default, dark or contrast)
 *  - context - the Microsoft Teams JS SDK context
 * methods:
 *  - setTheme - manually set the theme
 */
function useTeams(options) {
    var _a = (0, react_1.useState)(undefined), loading = _a[0], setLoading = _a[1];
    var _b = (0, react_1.useState)(undefined), inTeams = _b[0], setInTeams = _b[1];
    var _c = (0, react_1.useState)(undefined), fullScreen = _c[0], setFullScreen = _c[1];
    var _d = (0, react_1.useState)("default"), themeString = _d[0], setThemeString = _d[1];
    var initialTheme = (0, react_1.useState)(options && options.initialTheme ? options.initialTheme : getTheme())[0];
    var _e = (0, react_1.useState)(undefined), context = _e[0], setContext = _e[1];
    (0, react_1.useEffect)(function () {
        teams_js_1.app
            .initialize()
            .then(function () {
            teams_js_1.app
                .getContext()
                .then(function (context) {
                (0, react_dom_1.unstable_batchedUpdates)(function () {
                    setInTeams(true);
                    setContext(context);
                    setFullScreen(context.page.isFullScreen);
                });
                teams_js_1.pages.registerFullScreenHandler(function (isFullScreen) {
                    setFullScreen(isFullScreen);
                });
                setLoading(false);
            })
                .catch(function () {
                setLoading(false);
                setInTeams(false);
            });
        })
            .catch(function () {
            setLoading(false);
            setInTeams(false);
        });
    }, []);
    return [
        { inTeams: inTeams, fullScreen: fullScreen, context: context, themeString: themeString, loading: loading },
    ];
}
