// Maintenance only (needs Node.js + the CLI). Normal use of the toolkit does not run this.
//
// Snapshots the Power BI visual metadata that the Microsoft powerbi-report-authoring-cli validates against
// (visual types, roles, formatting objects and properties, visual container objects) into
// resources/pbir_metadata.json, so bhayankar_validate.py can run the same checks with no Node.js and no internet.
//
//   npm install -g @microsoft/powerbi-report-authoring-cli@<version>
//   node tools/snapshot_pbir_metadata.mjs            (writes resources/pbir_metadata.json)
//
// Source packages are MIT licensed (Copyright (c) Microsoft Corporation); see NOTICE.md.
import { execSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const npmRoot = execSync('npm root -g', { encoding: 'utf8' }).trim();
const cliDir = join(npmRoot, '@microsoft', 'powerbi-report-authoring-cli');
const cliPkg = JSON.parse(readFileSync(join(cliDir, 'package.json'), 'utf8'));
const schemaPkg = JSON.parse(readFileSync(join(cliDir, 'node_modules', '@microsoft', 'powerbi-core-visual-schema', 'package.json'), 'utf8'));
const cli = await import(pathToFileURL(join(cliDir, 'dist', 'index.js')).href);
const provider = await cli.getMetadataProvider();

function props(properties) {
  const out = {};
  for (const [name, def] of Object.entries(properties || {}).sort(([a], [b]) => a.localeCompare(b))) {
    out[name] = def.kind === 'enum' && Array.isArray(def.enumValues) ? { kind: 'enum', enumValues: def.enumValues } : def.kind;
  }
  return out;
}

const visuals = {};
const capabilities = {};
for (const vt of provider.getVisualTypes()) {
  const e = provider.getVisualEntry(vt);
  const roles = {};
  for (const [name, r] of Object.entries(e.roles || {})) roles[name] = r.kind;
  visuals[vt] = {
    capabilityKey: e.capabilityKey,
    roles,
    requiredRoles: e.requiredRoles || [],
    maxPerRole: e.maxPerRole || {},
    deprecated: Boolean(e.deprecated),
    ...(e.deprecated ? { deprecatedReason: e.deprecatedReason || null, preferredAlternative: e.preferredAlternative || null } : {}),
  };
  const key = e.capabilityKey || vt;
  if (!(key in capabilities)) {
    const cap = provider.getCapabilityEntry(key);
    capabilities[key] = cap ? Object.fromEntries(Object.entries(cap.objects || {}).sort(([a], [b]) => a.localeCompare(b))
      .map(([name, o]) => [name, props(o.properties)])) : null;
  }
}
const vcos = {};
for (const name of provider.getVcoNames()) vcos[name] = props(provider.getVcoEntry(name).properties);

const out = {
  source: `${cliPkg.name} ${cliPkg.version} with ${schemaPkg.name} ${schemaPkg.version} (MIT, Copyright (c) Microsoft Corporation)`,
  cli_version: cliPkg.version,
  constants: { visualNameMax: 50, pageNameMax: 50, slicerHeaderFloor: 28, slicerSelectorFloor: 32, defaultPadding: 8, layoutTolerance: 10 },
  visuals,
  capabilities,
  vcos,
};
const target = join(here, '..', 'resources', 'pbir_metadata.json');
writeFileSync(target, JSON.stringify(out, null, 1) + '\n');
console.log(`${Object.keys(visuals).length} visual types, ${Object.keys(capabilities).length} capability sets, ` +
  `${Object.keys(vcos).length} container objects -> ${target}`);
