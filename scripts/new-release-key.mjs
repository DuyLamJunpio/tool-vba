// Generates the ECDSA P-256 key pair used to sign company VBA packages (docs/tool-release-api-contract.md §1.2).
// Run once:  node scripts/new-release-key.mjs
// - The PRIVATE key goes only into the API host's secret store as TOOL_RELEASE_SIGNING_PRIVATE_KEY.
//   Never commit it, never paste it into chat or tickets.
// - The PUBLIC key goes into src/ToolVba.App/release.props (ToolVbaReleasePublicKey). It is not secret.
import { generateKeyPairSync } from 'node:crypto';

const { privateKey, publicKey } = generateKeyPairSync('ec', { namedCurve: 'P-256' });
const privatePem = privateKey.export({ type: 'pkcs8', format: 'pem' }).toString().trim();
const publicSpki = publicKey.export({ type: 'spki', format: 'der' }).toString('base64');

process.stdout.write(
  [
    '=== TOOL_RELEASE_SIGNING_PRIVATE_KEY (API secret store only, one line) ===',
    privatePem.replace(/\n/g, '\\n'),
    '',
    '=== ToolVbaReleasePublicKey (src/ToolVba.App/release.props) ===',
    publicSpki,
    '',
  ].join('\n'),
);
