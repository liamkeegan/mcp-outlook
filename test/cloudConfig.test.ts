import { describe, it, expect } from 'vitest';
import { resolveCloud, qualifyScopes, GRAPH_SCOPES, getCloud } from '../src/cloudConfig.js';

describe('resolveCloud', () => {
  it('defaults to global when the env is empty', () => {
    const result = resolveCloud({});
    expect(result.name).toBe('global');
    expect(result.authorityHost).toBe('https://login.microsoftonline.com');
    expect(result.authorityHostname).toBe('login.microsoftonline.com');
    expect(result.graphBaseUrl).toBe('https://graph.microsoft.com');
    expect(result.graphHost).toBe('graph.microsoft.com');
  });

  it('defaults to global when M365_CLOUD is an empty string', () => {
    const result = resolveCloud({ M365_CLOUD: '' });
    expect(result.name).toBe('global');
  });

  it('resolves the global preset by canonical name', () => {
    const result = resolveCloud({ M365_CLOUD: 'global' });
    expect(result).toEqual({
      name: 'global',
      authorityHost: 'https://login.microsoftonline.com',
      authorityHostname: 'login.microsoftonline.com',
      graphBaseUrl: 'https://graph.microsoft.com',
      graphHost: 'graph.microsoft.com',
      customAuthority: false,
    });
  });

  it('resolves the gcc-high preset by canonical name', () => {
    const result = resolveCloud({ M365_CLOUD: 'gcc-high' });
    expect(result).toEqual({
      name: 'gcc-high',
      authorityHost: 'https://login.microsoftonline.us',
      authorityHostname: 'login.microsoftonline.us',
      graphBaseUrl: 'https://graph.microsoft.us',
      graphHost: 'graph.microsoft.us',
      customAuthority: false,
    });
  });

  it('resolves the dod preset by canonical name', () => {
    const result = resolveCloud({ M365_CLOUD: 'dod' });
    expect(result).toEqual({
      name: 'dod',
      authorityHost: 'https://login.microsoftonline.us',
      authorityHostname: 'login.microsoftonline.us',
      graphBaseUrl: 'https://dod-graph.microsoft.us',
      graphHost: 'dod-graph.microsoft.us',
      customAuthority: false,
    });
  });

  it.each(['commercial', 'public', 'gcc'])('resolves alias "%s" to global', (alias) => {
    expect(resolveCloud({ M365_CLOUD: alias }).name).toBe('global');
  });

  it.each(['gcchigh', 'usgov', 'usgovhigh'])('resolves alias "%s" to gcc-high', (alias) => {
    expect(resolveCloud({ M365_CLOUD: alias }).name).toBe('gcc-high');
  });

  it('resolves alias "usgovdod" to dod', () => {
    expect(resolveCloud({ M365_CLOUD: 'usgovdod' }).name).toBe('dod');
  });

  it('normalises whitespace, case, and underscores', () => {
    expect(resolveCloud({ M365_CLOUD: ' GCC-High ' }).name).toBe('gcc-high');
    expect(resolveCloud({ M365_CLOUD: 'gcc_high' }).name).toBe('gcc-high');
    expect(resolveCloud({ M365_CLOUD: 'DoD' }).name).toBe('dod');
  });

  it('throws the exact contract error message for an invalid M365_CLOUD value', () => {
    expect(() => resolveCloud({ M365_CLOUD: 'nope' })).toThrow(
      'Invalid M365_CLOUD "nope". Expected one of: global, gcc-high, dod (aliases: commercial, public, gcc, gcchigh, usgov, usgovhigh, usgovdod).',
    );
  });

  it('applies GRAPH_BASE_URL on top of a preset and marks the cloud custom', () => {
    const result = resolveCloud({ M365_CLOUD: 'gcc-high', GRAPH_BASE_URL: 'https://graph.contoso.us' });
    expect(result.name).toBe('custom');
    expect(result.graphBaseUrl).toBe('https://graph.contoso.us');
    expect(result.graphHost).toBe('graph.contoso.us');
    // authority is untouched
    expect(result.authorityHost).toBe('https://login.microsoftonline.us');
    // a GRAPH_BASE_URL-only override doesn't touch the login authority
    expect(result.customAuthority).toBe(false);
  });

  it('treats a whitespace-only GRAPH_BASE_URL as unset', () => {
    const result = resolveCloud({ M365_CLOUD: 'gcc-high', GRAPH_BASE_URL: '  ' });
    expect(result.name).toBe('gcc-high');
    expect(result.graphBaseUrl).toBe('https://graph.microsoft.us');
    expect(result.graphHost).toBe('graph.microsoft.us');
  });

  it('rejects a GRAPH_BASE_URL containing userinfo', () => {
    expect(() => resolveCloud({ GRAPH_BASE_URL: 'https://user:pw@graph.microsoft.us' })).toThrow(
      'Invalid GRAPH_BASE_URL "https://user:pw@graph.microsoft.us": must be an https origin such as https://graph.microsoft.us',
    );
  });

  it('strips a trailing slash from GRAPH_BASE_URL', () => {
    const result = resolveCloud({ GRAPH_BASE_URL: 'https://graph.microsoft.com/' });
    expect(result.graphBaseUrl).toBe('https://graph.microsoft.com');
  });

  it('tolerates a schemeless GRAPH_BASE_URL', () => {
    const result = resolveCloud({ GRAPH_BASE_URL: 'graph.microsoft.us' });
    expect(result.graphBaseUrl).toBe('https://graph.microsoft.us');
    expect(result.graphHost).toBe('graph.microsoft.us');
  });

  it('rejects a non-https GRAPH_BASE_URL with the exact contract message', () => {
    expect(() => resolveCloud({ GRAPH_BASE_URL: 'http://graph.microsoft.us' })).toThrow(
      'Invalid GRAPH_BASE_URL "http://graph.microsoft.us": must be an https origin such as https://graph.microsoft.us',
    );
  });

  it('rejects a GRAPH_BASE_URL containing a path', () => {
    expect(() => resolveCloud({ GRAPH_BASE_URL: 'https://graph.microsoft.us/v1.0' })).toThrow(
      'Invalid GRAPH_BASE_URL "https://graph.microsoft.us/v1.0": must be an https origin such as https://graph.microsoft.us',
    );
  });

  it('applies a LOGIN_AUTHORITY_HOST override and marks the cloud custom', () => {
    const result = resolveCloud({ LOGIN_AUTHORITY_HOST: 'https://login.contoso.com' });
    expect(result.name).toBe('custom');
    expect(result.authorityHost).toBe('https://login.contoso.com');
    expect(result.authorityHostname).toBe('login.contoso.com');
    // graph is untouched
    expect(result.graphBaseUrl).toBe('https://graph.microsoft.com');
    expect(result.customAuthority).toBe(true);
  });

  it('rejects a non-https LOGIN_AUTHORITY_HOST with the exact contract message', () => {
    expect(() => resolveCloud({ LOGIN_AUTHORITY_HOST: 'http://login.microsoftonline.us' })).toThrow(
      'Invalid LOGIN_AUTHORITY_HOST "http://login.microsoftonline.us": must be an https origin such as https://login.microsoftonline.us',
    );
  });

  it('keeps a non-default port in authorityHostname (host:port), separate from graphHost (hostname only)', () => {
    const result = resolveCloud({ LOGIN_AUTHORITY_HOST: 'https://login.example.com:8443' });
    expect(result.authorityHostname).toBe('login.example.com:8443');
    expect(result.authorityHost).toBe('https://login.example.com:8443');
  });

  it('rejects a LOGIN_AUTHORITY_HOST containing userinfo', () => {
    expect(() => resolveCloud({ LOGIN_AUTHORITY_HOST: 'https://user:pw@login.microsoftonline.us' })).toThrow(
      'Invalid LOGIN_AUTHORITY_HOST "https://user:pw@login.microsoftonline.us": must be an https origin such as https://login.microsoftonline.us',
    );
  });

  it('keeps the preset name when overrides equal the preset values', () => {
    const result = resolveCloud({
      M365_CLOUD: 'gcc-high',
      GRAPH_BASE_URL: 'https://graph.microsoft.us',
      LOGIN_AUTHORITY_HOST: 'https://login.microsoftonline.us',
    });
    expect(result.name).toBe('gcc-high');
  });
});

describe('getCloud', () => {
  it('is exported as a function without reading process.env at import time', () => {
    // NOTE: do not call getCloud() here — on invalid configuration it calls
    // process.exit(1), which would kill the test runner. Importing the
    // module must not read process.env; only calling getCloud() does.
    expect(typeof getCloud).toBe('function');
  });
});

describe('qualifyScopes', () => {
  it('prefixes all scopes with the graph base URL', () => {
    const result = qualifyScopes(GRAPH_SCOPES, 'https://graph.microsoft.us');
    expect(result).toEqual([
      'https://graph.microsoft.us/User.Read',
      'https://graph.microsoft.us/Calendars.Read',
      'https://graph.microsoft.us/Calendars.ReadWrite',
      'https://graph.microsoft.us/Mail.Send',
      'https://graph.microsoft.us/Mail.ReadWrite',
      'https://graph.microsoft.us/Mail.Read',
      'https://graph.microsoft.us/People.Read',
    ]);
  });

  it('leaves an already-qualified scope untouched', () => {
    const result = qualifyScopes(['https://graph.microsoft.us/User.Read'], 'https://graph.microsoft.us');
    expect(result).toEqual(['https://graph.microsoft.us/User.Read']);
  });

  it('does not produce a double slash when the base URL has a trailing slash', () => {
    const result = qualifyScopes(['User.Read'], 'https://graph.microsoft.us/');
    expect(result).toEqual(['https://graph.microsoft.us/User.Read']);
  });
});
