// Structural checks only. This does not replace Debug > Compile VBAProject.
import { readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const targets = process.argv.slice(2);
if (!targets.length) targets.push('modLicenseApiV1.bas', 'TestResults/modLicenseRegression.bas');

function uncomment(line) {
  let quoted = false;
  for (let i = 0; i < line.length; i++) {
    if (line[i] === '"') {
      if (quoted && line[i + 1] === '"') { i++; continue; }
      quoted = !quoted;
    } else if (line[i] === "'" && !quoted) return line.slice(0, i);
  }
  return line;
}

function argumentsOf(text) {
  if (!text.trim()) return [];
  let quoted = false;
  let depth = 0;
  let start = 0;
  const result = [];
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') {
      if (quoted && text[i + 1] === '"') { i++; continue; }
      quoted = !quoted;
    } else if (!quoted) {
      if (c === '(') depth++;
      if (c === ')') depth--;
      if (c === ',' && depth === 0) { result.push(text.slice(start, i).trim()); start = i + 1; }
    }
  }
  assert(!quoted && depth === 0, 'Unbalanced call expression');
  result.push(text.slice(start).trim());
  return result;
}

assert.equal(argumentsOf('Array("a,b", ""), Cfg_ProductKey(), "x"').length, 3);
assert.equal(uncomment(`value = "it's ""quoted""" ' comment`).trim(), `value = "it's ""quoted"""`);

const checkedCalls = new Set([
  'CacheSave', 'CacheLoad', 'VerifyTokenClaims', 'AcceptVerifiedState', 'StageVerifiedCredential',
  'ClockCanUse', 'ObserveClock', 'RevalidationCycle', 'CanAttemptRefresh', 'AttemptRevalidation',
  'FindPublicKey', 'KeyPairsValid', 'TokenKeyIdsMatch', 'NativeCacheSaveTest', 'NativeCacheLoadTest',
  'ShouldClearCacheForServerError', 'ShouldRememberServerDenial', 'IsServerRefusal', 'ApiErrorCode',
  'RememberServerDenial', 'SaveClockWatermark', 'AdoptSharedCacheState', 'RevalidationDue',
]);
// Procedures the 11/11/2026 DEC-107 addendum requires; a missing one must fail here, before Excel.
const requiredProcedures = ['ShouldRememberServerDenial', 'IsServerRefusal', 'ApiErrorCode',
  'RememberServerDenial', 'SaveClockWatermark', 'AdoptSharedCacheState', 'RevalidationDue'];
const reserved = /\b(?:Dim|Const|ByVal|ByRef)\s+(?:Input|Open|Close|Print|Set|Get|Name|Date|Time|Error|Type|Select|Option|Function|Sub|Loop|Next|For|Else|If|Then|Public|Private)\b/i;

for (const target of targets) {
  const bytes = readFileSync(resolve(root, target));
  assert([...bytes].every(value => value < 128), `${target}: ASCII required`);
  const physical = bytes.toString('ascii').split(/\r?\n/);
  assert(physical.every(line => line.length < 1024), `${target}: VBA physical line limit`);
  const logical = [];
  let pending = '';
  let first = 0;
  for (let i = 0; i < physical.length; i++) {
    const line = uncomment(physical[i]).trim();
    if (!pending) first = i + 1;
    if (/\s_\s*$/.test(line)) { pending += line.replace(/\s_\s*$/, ' '); continue; }
    logical.push({ text: pending + line, line: first });
    pending = '';
  }
  assert(!pending, `${target}: dangling continuation`);
  const procedures = new Map();
  let current = null;
  const blocks = [];
  for (const entry of logical) {
    const text = entry.text;
    assert(!reserved.test(text), `${target}:${entry.line}: reserved declaration identifier`);
    const declaration = /^(?:Public|Private)\s+(Sub|Function)\s+(\w+)\(([\s\S]*)\)(?:\s+As\s+\w+(?:\(\))?)?$/i.exec(text);
    if (declaration) {
      assert(!current, `${target}:${entry.line}: nested procedure`);
      const [, kind, name, params] = declaration;
      assert(!procedures.has(name.toLowerCase()), `${target}: duplicate procedure ${name}`);
      const parameters = argumentsOf(params);
      current = { kind: kind.toLowerCase(), name, entries: [], max: parameters.length,
        min: parameters.filter(value => !/^Optional\b/i.test(value)).length };
      procedures.set(name.toLowerCase(), current);
      continue;
    }
    if (!current) continue;
    if (/^End (Sub|Function)$/i.test(text)) {
      assert(text.toLowerCase() === `end ${current.kind}`, `${target}:${entry.line}: wrong procedure end`);
      assert(!blocks.length, `${target}:${entry.line}: unclosed block in ${current.name}`);
      current = null;
      continue;
    }
    current.entries.push(entry);
    let opened = null;
    if (/^If\s.+\sThen$/i.test(text)) opened = 'if';
    else if (/^For\s/i.test(text)) opened = 'for';
    else if (/^Do(?:\s|$)/i.test(text)) opened = 'do';
    else if (/^Select Case\s/i.test(text)) opened = 'select';
    else if (/^With\s/i.test(text)) opened = 'with';
    if (opened) blocks.push(opened);
    const closing = /^End If$/i.test(text) ? 'if' : /^Next(?:\s|$)/i.test(text) ? 'for' :
      /^Loop(?:\s|$)/i.test(text) ? 'do' : /^End Select$/i.test(text) ? 'select' : /^End With$/i.test(text) ? 'with' : null;
    if (closing) assert(blocks.pop() === closing, `${target}:${entry.line}: unmatched ${closing} block`);
  }
  assert(!current, `${target}: unfinished procedure`);
  for (const name of requiredProcedures) {
    assert(procedures.has(name.toLowerCase()), `${target}: missing procedure ${name}`);
  }
  for (const procedure of procedures.values()) {
    for (const entry of procedure.entries) {
      for (const name of checkedCalls) {
        const pattern = new RegExp(`\\b${name}\\b`, 'ig');
        for (const match of entry.text.matchAll(pattern)) {
          const suffix = entry.text.slice(match.index + name.length).trimStart();
          if (suffix.startsWith('=')) continue;
          let args;
          if (suffix.startsWith('(')) {
            let quoted = false;
            let depth = 0;
            let end = -1;
            for (let i = 0; i < suffix.length; i++) {
              if (suffix[i] === '"') {
                if (quoted && suffix[i + 1] === '"') { i++; continue; }
                quoted = !quoted;
              } else if (!quoted) {
                if (suffix[i] === '(') depth++;
                if (suffix[i] === ')' && --depth === 0) { end = i; break; }
              }
            }
            assert(end >= 0, `${target}:${entry.line}: unbalanced ${name} call`);
            args = argumentsOf(suffix.slice(1, end));
          } else {
            const prefix = entry.text.slice(0, match.index).trim();
            if (prefix && !/Then$/i.test(prefix)) continue;
            args = argumentsOf(suffix);
          }
          const callee = procedures.get(name.toLowerCase());
          assert(callee, `${target}:${entry.line}: missing ${name}`);
          assert(args.length >= callee.min && args.length <= callee.max,
            `${target}:${entry.line}: ${name} has ${args.length} arguments; expected ${callee.min}-${callee.max}`);
        }
      }
    }
  }
  console.log(`${target}: structural/arity checks passed (${procedures.size} procedures). Not an Excel compile.`);
}
