// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { useEffect, useState } from "react";
import { unstable_batchedUpdates as batchedUpdates } from "react-dom";
import { app, pages } from "@microsoft/teams-js";
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
export function useTeams(options) {
    var _a = useState(undefined), loading = _a[0], setLoading = _a[1];
    var _b = useState(undefined), inTeams = _b[0], setInTeams = _b[1];
    var _c = useState(undefined), fullScreen = _c[0], setFullScreen = _c[1];
    var _d = useState("default"), themeString = _d[0], setThemeString = _d[1];
    var initialTheme = useState(options && options.initialTheme ? options.initialTheme : getTheme())[0];
    var _e = useState(undefined), context = _e[0], setContext = _e[1];
    useEffect(function () {
        app
            .initialize()
            .then(function () {
            app
                .getContext()
                .then(function (context) {
                batchedUpdates(function () {
                    setInTeams(true);
                    setContext(context);
                    setFullScreen(context.page.isFullScreen);
                });
                pages.registerFullScreenHandler(function (isFullScreen) {
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
