# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| Latest (`main`) | Yes |
| Tagged releases | Yes (latest tag only) |
| Older tags | No — please update to the latest release |

---

## Reporting a Vulnerability

If you discover a security vulnerability in any script in this repository, **please do not open a public GitHub issue**.

Instead, report it privately:

1. Go to the [Security tab](https://github.com/nakamas-it/powershell-scripts/security) of this repository
2. Click **"Report a vulnerability"** to open a private advisory
3. Include as much detail as possible:
   - Which script is affected
   - Steps to reproduce the issue
   - Potential impact
   - Suggested fix (optional)

You will receive an acknowledgement within **48 hours**. We aim to release a patched version within **7 days** of confirmation, depending on severity.

---

## Security Design Principles

All scripts in this repository follow these principles:

- **No credential storage** — scripts never store, log, or transmit passwords or secrets
- **No network uploads** — scripts only read local system data; nothing is sent externally
- **Read-only by design** — scripts collect metadata (file names, sizes, dates) and do not modify system state unless explicitly stated
- **No global execution policy changes** — scripts never call `Set-ExecutionPolicy` at machine or user scope
- **HTTPS only** — all downloads use HTTPS with TLS 1.2+

---

## Safe Usage

### Execution Policy
Run scripts with a per-process bypass rather than changing global policy:
```powershell
powershell -ExecutionPolicy Bypass -File .\ScriptName.ps1
```

### Verify Integrity Before Running
Always check the SHA256 hash of a downloaded script against the hash published in the release notes:
```powershell
(Get-FileHash .\ScriptName.ps1 -Algorithm SHA256).Hash
```

### Pulling from GitHub
Always use HTTPS raw URLs. For production use, pin to a specific release tag rather than `main`:
```powershell
# Pinned to a release tag (recommended for production)
irm https://raw.githubusercontent.com/nakamas-it/powershell-scripts/v1.0.0/Run.ps1 | iex

# Latest (acceptable for personal/lab use)
irm https://raw.githubusercontent.com/nakamas-it/powershell-scripts/main/Run.ps1 | iex
```

### Enterprise / RMM Deployment
For deployment via Intune, N-central, or other RMM tools, sign scripts with a trusted code-signing certificate before distributing:
```powershell
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert
Set-AuthenticodeSignature -FilePath .\ScriptName.ps1 -Certificate $cert
```

---

## Scope

This policy applies to all `.ps1` files, configuration files, and workflows in this repository.
Third-party tools or dependencies referenced in documentation are outside this scope.
