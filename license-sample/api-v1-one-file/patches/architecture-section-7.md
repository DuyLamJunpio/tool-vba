## 7. Excel/Word client (VBA) design

Owner-decided on 29/10/2026 (DEC-107): the Windows Excel client uses device-code activation and closes its own workbook when access cannot be granted.

```text
Standard module: modLicenseApiV1 (one-file sample in tool-vba)
  Auto_Open -> local signature/claims/clock checks -> validate-device when needed
  device-code -> approval on Store -> redeem (no account JWT in Excel)
  License_PremiumAllowed / License_FeatureAllowed -> guarded session revalidation
  WinHTTP HTTPS -> TLS 1.2/1.3, system proxy or optional configured override
  DPAPI cache in current-user Registry -> ciphertext only
ThisWorkbook (merge handlers into the owner's existing module)
  BeforeClose -> hide premium sheets
  BeforeSave / AfterSave -> persist hidden sheets; restore entitled session view
Sheets
  Owner-created visible guide: explain enabling macros and activation
  Configured premium worksheets: xlSheetVeryHidden in the distributed file
```

Rules:

- Without entitlement, hide configured premium sheets, show fixed Vietnamese guidance, then close **only ThisWorkbook with SaveChanges:=False** after OK/X. Never quit Excel or close another workbook. During startup, a new device code may be offered once when activation is required or unusable cache was successfully cleared; No/X cancels and closes this workbook.
- `Cfg_ProtectedSheetNames` defaults to an empty list. The workbook author creates a separate visible guide, configures premium sheet names, merges close/save event handlers and ships the saved file with those sheets `xlSheetVeryHidden`. Valid entitlement makes them `xlSheetVisible`; loss of entitlement or close hides them again. Do not use server-issued sheet passwords or sheet protection as the licensing mechanism. Invalid sheet configuration fails closed. Hidden sheets deter ordinary sharing; they do not encrypt content or prevent deliberate extraction. With macros disabled, only the saved visibility state applies.
- Excel uses only `device-code`, `redeem` and `validate-device`; approval and account login stay on Store. It never receives or stores an account JWT. Do not store activation tokens, fingerprints or device codes in sheets, logs, URLs or dialogs. Only `user_code` is shown. Cache ciphertext in HKCU holds the activation token, fingerprint, successful-validation time and greatest observed UTC time; DPAPI uses the same fixed public product-specific entropy in both directions. Undecryptable older cache is treated as absent, without a decryption error, so upgrading may require activation once.
- Before granting access, verify Ed25519 and all required claims. Select a public key by header `kid`, require payload `kid` to match, and reject an unknown key or invalid signature. Support two explicit key pairs; the sample preserves its first `k2026-10` public key and leaves the second pair empty. Never embed a private key or shared secret. All verification errors fail closed.
- The maximum observed UTC time only increases and does not move the successful-validation anchor. A clock earlier than that maximum or the verified token's `iat` fails with `CLOCK_ROLLBACK`; ask the customer to synchronize Windows time and reopen. Keep the cache. This is a local check, not protection against deleting or restoring the whole cache.
- Preserve DEC-104/DEC-105: a token past only `exp` is still a renewal credential for `validate-device`, not an offline entitlement and not a reason to require a new device code. Fresh tokens must not be expired. Signed `rvi` and `ogh` govern validation and offline grace; `ogh = 0` requires successful online validation when opening. A server refusal never permits offline fallback. When a guarded feature is called after revalidation is due during a session, try online once per `rvi` cycle with a reentry guard and existing timeouts; a successful renewal is transparent. There is no background timer.
- Clear cached credentials for local fingerprint mismatch or failed cached-token verification, and for HTTP 4xx except 429 with `TOKEN_INVALID`, `DEVICE_REVOKED` or `DEVICE_MISMATCH`. Offer a new code only after successful clearing. Do not clear for network errors, rate limits, request timestamp errors, HTTP 5xx, `LICENSE_EXPIRED` or `LICENSE_SUSPENDED`; renewal/unblocking can reuse the credential. A clock-watermark write retains the token and validation anchor.
- Read fingerprint identifiers without changing the OS. Use `%SystemDrive%` for the volume serial: C: produces the same hash inputs as before, while an existing installation on another system drive may need reactivation. Use Windows user proxy/PAC settings with WinHTTP defaults as fallback, or an optional explicit proxy; never disable certificate checks or follow redirects. Product/version/API/activation-page configuration comes from the approved release, not guessed values.

---

