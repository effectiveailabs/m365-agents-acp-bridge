# Security

## Reporting

After the public GitHub repository exists, please report vulnerabilities through GitHub Security Advisories for `effectiveailabs/m365-agents-acp-bridge`.

Until then, report privately to the repository maintainers.

## Secret Handling

The bridge is designed to consume short-lived delegated Microsoft access tokens in `external_token` mode. It must not store Microsoft refresh tokens.

Do not include the following in issues, logs, tests, fixtures, or examples:

- Microsoft access tokens,
- Microsoft refresh tokens,
- Authorization headers,
- client secrets,
- connection strings,
- direct-connect URLs,
- tenant-private payloads.

Use `env:` or `file:` secret refs in configuration examples.
