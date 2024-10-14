// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import 'react-native-get-random-values';
import { useEffect, useState } from "react";
import { unstable_batchedUpdates as batchedUpdates } from "react-dom";
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
export function useTeams(options?: {
  initialTheme?: string;
  setThemeHandler?: (theme?: string) => void;
}): [
  {
    inTeams?: boolean;
    fullScreen?: boolean;
    context?: app.Context;
    loading?: boolean;
  }
] {
  const [loading, setLoading] = useState<boolean | undefined>(undefined);
  const [inTeams, setInTeams] = useState<boolean | undefined>(undefined);
  const [fullScreen, setFullScreen] = useState<boolean | undefined>(undefined);
  const [context, setContext] = useState<app.Context | undefined>(undefined);

  useEffect(() => {

    app
      .initialize()
      .then(() => {
        app
          .getContext()
          .then((context) => {
            batchedUpdates(() => {
              setInTeams(true);
              setContext(context);
              setFullScreen(context.page.isFullScreen);
            });
            setLoading(false);
          })
          .catch(() => {
            setLoading(false);
            setInTeams(false);
          });
      })
      .catch(() => {
        setLoading(false);
        setInTeams(false);
      });
  }, []);

  return [
    { inTeams, fullScreen, context, loading },
  ];
}
