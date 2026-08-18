/**
 * cloudConfig.ts
 *
 * Resolves which Microsoft 365 / Entra ID national cloud this server talks
 * to, based on environment variables. This module has no imports from MSAL
 * or the Graph SDK. `resolveCloud` is a pure function of its `env` argument;
 * `getCloud()` is the only place that reads `process.env` implicitly, and it
 * memoises the result. If the configuration is invalid, `getCloud()` prints
 * a one-line error to stderr and calls `process.exit(1)` — the documented
 * clean-startup-exit behaviour — rather than letting the exception propagate.
 *
 * Environment variables:
 * - M365_CLOUD (optional): selects a preset national cloud. The value is
 *   normalised (trimmed, lower-cased, `_` replaced with `-`) before being
 *   matched against the accepted names/aliases below. Unset or empty
 *   defaults to 'global'.
 *     - 'global'   (aliases: commercial, public, gcc)
 *     - 'gcc-high' (aliases: gcchigh, usgov, usgovhigh)
 *     - 'dod'      (alias: usgovdod)
 *   Any other value throws.
 * - GRAPH_BASE_URL (optional): overrides the Microsoft Graph origin for the
 *   selected cloud, e.g. https://graph.microsoft.us. Must be an https
 *   origin with no path (other than `/`), query, or hash. A missing scheme
 *   is tolerated (`https://` is prepended) and a trailing slash is
 *   tolerated (stripped). Applied on top of the M365_CLOUD preset.
 * - LOGIN_AUTHORITY_HOST (optional): overrides the Entra ID login origin
 *   for the selected cloud, e.g. https://login.microsoftonline.us. Same
 *   validation rules as GRAPH_BASE_URL.
 *
 * If either override resolves to an origin different from the selected
 * preset's corresponding value, the resolved cloud's `name` becomes
 * 'custom'.
 */

export type CloudName = 'global' | 'gcc-high' | 'dod' | 'custom';

export interface CloudConfig {
  name: CloudName;
  authorityHost: string; // https origin, no trailing slash, e.g. https://login.microsoftonline.us
  authorityHostname: string; // host[:port] (MSAL knownAuthorities compares host:port), e.g. login.microsoftonline.us
  graphBaseUrl: string; // https origin, no trailing slash, e.g. https://graph.microsoft.us
  graphHost: string; // bare hostname, no port (Graph SDK customHosts compares hostname only), e.g. graph.microsoft.us
  // True only when LOGIN_AUTHORITY_HOST produced an origin different from the
  // selected preset's authorityHost. Lets authConfig.ts set MSAL's
  // knownAuthorities only when a custom login host actually needs trusting.
  customAuthority: boolean;
}

export const GRAPH_SCOPES = [
  'User.Read',
  'Calendars.Read',
  'Calendars.ReadWrite',
  'Mail.Send',
  'Mail.ReadWrite',
  'Mail.Read',
  'People.Read',
] as const;

type PresetName = 'global' | 'gcc-high' | 'dod';

// Presets hold only the origins; host/hostname are always derived from them
// via normalizeOrigin so there is a single source of truth for parsing.
const PRESETS: Record<PresetName, { authorityHost: string; graphBaseUrl: string }> = {
  'global': {
    authorityHost: 'https://login.microsoftonline.com',
    graphBaseUrl: 'https://graph.microsoft.com',
  },
  'gcc-high': {
    authorityHost: 'https://login.microsoftonline.us',
    graphBaseUrl: 'https://graph.microsoft.us',
  },
  'dod': {
    authorityHost: 'https://login.microsoftonline.us',
    graphBaseUrl: 'https://dod-graph.microsoft.us',
  },
};

// Maps every accepted normalised M365_CLOUD string (including each
// preset's own canonical name) to the preset it selects.
const ALIASES: Record<string, PresetName> = {
  'global': 'global',
  'commercial': 'global',
  'public': 'global',
  'gcc': 'global',
  'gcc-high': 'gcc-high',
  'gcchigh': 'gcc-high',
  'usgov': 'gcc-high',
  'usgovhigh': 'gcc-high',
  'dod': 'dod',
  'usgovdod': 'dod',
};

// Canonical preset names, plus the aliases whose target differs from their
// own key (i.e. excludes 'global', 'gcc-high', 'dod' themselves, which are
// already covered by the canonical-names list).
const CANONICAL_NAMES = Object.keys(PRESETS) as PresetName[];
const ALIAS_NAMES = Object.keys(ALIASES).filter((key) => key !== ALIASES[key]);

function resolvePresetName(rawValue: string | undefined): PresetName {
  const trimmed = (rawValue ?? '').trim();
  if (trimmed === '') {
    return 'global';
  }

  const normalized = trimmed.toLowerCase().replace(/_/g, '-');
  const presetName = ALIASES[normalized];
  if (!presetName) {
    throw new Error(
      `Invalid M365_CLOUD "${rawValue}". Expected one of: ${CANONICAL_NAMES.join(', ')} (aliases: ${ALIAS_NAMES.join(', ')}).`,
    );
  }

  return presetName;
}

function normalizeOrigin(value: string, varName: string, example: string): { origin: string; host: string; hostname: string } {
  const trimmedValue = value.trim();

  const invalid = (): never => {
    throw new Error(`Invalid ${varName} "${value}": must be an https origin such as ${example}`);
  };

  const withScheme = trimmedValue.includes('://') ? trimmedValue : `https://${trimmedValue}`;

  let url: URL;
  try {
    url = new URL(withScheme);
  } catch {
    return invalid();
  }

  if (
    url.protocol !== 'https:' ||
    url.pathname !== '/' ||
    url.search !== '' ||
    url.hash !== '' ||
    url.username !== '' ||
    url.password !== ''
  ) {
    return invalid();
  }

  return { origin: url.origin, host: url.host, hostname: url.hostname };
}

export function resolveCloud(env: NodeJS.ProcessEnv = process.env): CloudConfig {
  const presetName = resolvePresetName(env.M365_CLOUD);
  const preset = PRESETS[presetName];

  // Run the preset origins through the same normalizeOrigin path as
  // overrides, so authorityHostname/graphHost are always derived rather
  // than duplicated.
  let { origin: authorityHost, host: authorityHostname } = normalizeOrigin(
    preset.authorityHost,
    'LOGIN_AUTHORITY_HOST',
    'https://login.microsoftonline.us',
  );
  let { origin: graphBaseUrl, hostname: graphHost } = normalizeOrigin(
    preset.graphBaseUrl,
    'GRAPH_BASE_URL',
    'https://graph.microsoft.us',
  );
  let isCustom = false;
  let customAuthority = false;

  const graphBaseUrlOverride = env.GRAPH_BASE_URL?.trim();
  if (graphBaseUrlOverride) {
    const { origin, hostname } = normalizeOrigin(graphBaseUrlOverride, 'GRAPH_BASE_URL', 'https://graph.microsoft.us');
    if (origin !== graphBaseUrl) {
      isCustom = true;
    }
    graphBaseUrl = origin;
    graphHost = hostname;
  }

  const loginAuthorityHostOverride = env.LOGIN_AUTHORITY_HOST?.trim();
  if (loginAuthorityHostOverride) {
    const { origin, host } = normalizeOrigin(loginAuthorityHostOverride, 'LOGIN_AUTHORITY_HOST', 'https://login.microsoftonline.us');
    if (origin !== authorityHost) {
      isCustom = true;
      customAuthority = true;
    }
    authorityHost = origin;
    authorityHostname = host;
  }

  return {
    name: isCustom ? 'custom' : presetName,
    authorityHost,
    authorityHostname,
    graphBaseUrl,
    graphHost,
    customAuthority,
  };
}

// graphBaseUrl doubles as the OAuth resource for these scopes (GRAPH_BASE_URL
// is used both as the request origin AND as the token audience), so callers
// must pass the real Graph resource origin for the cloud, never a proxy.
export function qualifyScopes(scopes: readonly string[], graphBaseUrl: string): string[] {
  const base = graphBaseUrl.endsWith('/') ? graphBaseUrl.slice(0, -1) : graphBaseUrl;
  return scopes.map((scope) => (scope.includes('://') ? scope : `${base}/${scope}`));
}

let resolved: CloudConfig | undefined;
/**
 * Resolve the cloud from process.env once and cache it. On invalid configuration
 * print the one-line error to stderr and exit(1) — imports are hoisted, so this is
 * the earliest point at which startup validation can fail cleanly.
 */
export function getCloud(): CloudConfig {
  if (!resolved) {
    try {
      resolved = resolveCloud();
    } catch (error) {
      console.error(`outlook-mcp: ${error instanceof Error ? error.message : String(error)}`);
      process.exit(1);
    }
  }
  return resolved;
}
