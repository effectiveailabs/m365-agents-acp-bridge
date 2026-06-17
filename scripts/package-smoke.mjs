import { execFileSync } from 'node:child_process';
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const packDir = mkdtempSync(join(tmpdir(), 'm365-agents-acp-bridge-pack-'));

try {
  execFileSync(pnpmBin(), ['pack', '--pack-destination', packDir], {
    cwd: root,
    stdio: 'pipe',
  });

  const tarball = readdirSync(packDir).find((file) => file.endsWith('.tgz'));
  if (!tarball) {
    throw new Error('pnpm pack did not produce a tarball');
  }

  const tarballPath = join(packDir, tarball);
  const tarEntries = execFileSync('tar', ['-tzf', tarballPath], {
    encoding: 'utf8',
  })
    .split('\n')
    .filter(Boolean);

  assertIncludes(tarEntries, 'package/dist/index.js');
  assertIncludes(tarEntries, 'package/dist/cli.js');
  assertIncludes(tarEntries, 'package/dist/index.d.ts');
  assertIncludes(tarEntries, 'package/README.md');
  assertIncludes(tarEntries, 'package/LICENSE');
  assertNotIncluded(tarEntries, /^package\/src\//);
  assertNotIncluded(tarEntries, /^package\/tests\//);
  assertNotIncluded(tarEntries, /^package\/\.env/);

  execFileSync('tar', ['-xzf', tarballPath, '-C', packDir], {
    stdio: 'pipe',
  });

  const packageDir = join(packDir, 'package');
  const consumerDir = join(packDir, 'consumer');
  mkdirSync(consumerDir);
  writeFileSync(
    join(consumerDir, 'package.json'),
    JSON.stringify({ private: true, type: 'module' }, null, 2),
  );

  execFileSync(pnpmBin(), ['add', tarballPath], {
    cwd: consumerDir,
    stdio: 'pipe',
  });

  await import(
    pathToFileURL(
      join(consumerDir, 'node_modules/@effectiveai/m365-agents-acp-bridge/dist/index.js'),
    ).href
  );

  const help = execFileSync(pnpmBin(), ['exec', 'm365-acp', '--help'], {
    cwd: consumerDir,
    encoding: 'utf8',
  });
  if (!help.includes('M365 Agents ACP Bridge')) {
    throw new Error('CLI help did not print the expected bridge heading');
  }

  const forbiddenPatterns = [
    /short-lived-token/i,
    /super-secret/i,
    /secret\.example/i,
    /eyJsecret/i,
  ];
  for (const file of walk(packageDir)) {
    const contents = readFileSync(file, 'utf8');
    for (const pattern of forbiddenPatterns) {
      if (pattern.test(contents)) {
        throw new Error(`Packaged file contains test secret pattern ${pattern}: ${file}`);
      }
    }
  }

  console.log(`Package smoke passed: ${tarball}`);
} finally {
  rmSync(packDir, { recursive: true, force: true });
}

function pnpmBin() {
  return process.platform === 'win32' ? 'pnpm.cmd' : 'pnpm';
}

function assertIncludes(entries, expected) {
  if (!entries.includes(expected)) {
    throw new Error(`Package is missing ${expected}`);
  }
}

function assertNotIncluded(entries, pattern) {
  const match = entries.find((entry) => pattern.test(entry));
  if (match) {
    throw new Error(`Package includes unexpected file ${match}`);
  }
}

function* walk(directory) {
  for (const entry of readdirSync(directory)) {
    const path = join(directory, entry);
    if (!existsSync(path)) {
      continue;
    }
    const stat = statSync(path);
    if (stat.isDirectory()) {
      yield* walk(path);
    } else {
      yield path;
    }
  }
}
