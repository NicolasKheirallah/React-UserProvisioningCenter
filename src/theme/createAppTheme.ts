import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { Theme } from '@fluentui/react-components';
import {
  teamsDarkTheme,
  teamsHighContrastTheme,
  teamsLightTheme,
  webLightTheme
} from '@fluentui/react-components';
import { createV9ThemeFromSp, type ISpThemeEffects, type ISpThemePalette } from './spThemeToV9';

export type TeamsAppTheme = 'default' | 'dark' | 'contrast';

/**
 * Resolves the Fluent v9 theme for the current host. Teams dark/high-contrast
 * take precedence (SharePoint theme tokens are not meaningful there); on
 * SharePoint pages the section's theme variant is bridged so the web part
 * follows site branding and inverted (dark) variants.
 */
export function createAppTheme(
  spTheme: IReadonlyTheme | undefined,
  teamsTheme: TeamsAppTheme | undefined
): Theme {
  if (teamsTheme === 'dark') {
    return teamsDarkTheme;
  }
  if (teamsTheme === 'contrast') {
    return teamsHighContrastTheme;
  }
  if (spTheme?.palette) {
    const base: Theme = spTheme.isInverted ? teamsDarkTheme : webLightTheme;
    return createV9ThemeFromSp(
      spTheme.palette as ISpThemePalette,
      spTheme.effects as ISpThemeEffects,
      !!spTheme.isInverted,
      base
    );
  }
  return teamsTheme === 'default' ? teamsLightTheme : webLightTheme;
}
