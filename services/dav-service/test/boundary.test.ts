import { readFileSync, readdirSync, statSync } from 'node:fs';
import { builtinModules } from 'node:module';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

/**
 * dav-service may depend on the gateway contract and its own npm dependencies,
 * never on the demo site (or any other app). This is the architectural rule
 * from the brief, enforced on every test run.
 */
const serviceRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const pkg = JSON.parse(readFileSync(join(serviceRoot, 'package.json'), 'utf8')) as {
  dependencies?: Record<string, string>;
  devDependencies?: Record<string, string>;
};
const allowedPackages = new Set(Object.keys(pkg.dependencies ?? {}));

function sourceFiles(dir: string): string[] {
  return readdirSync(dir).flatMap((entry) => {
    const path = join(dir, entry);
    if (statSync(path).isDirectory()) return sourceFiles(path);
    return /\.[cm]?[jt]s$/.test(entry) ? [path] : [];
  });
}

const IMPORT = /(?:import|export)\s[^'"]*?from\s*['"]([^'"]+)['"]|import\s*\(\s*['"]([^'"]+)['"]\s*\)|import\s+['"]([^'"]+)['"]|require\(\s*['"]([^'"]+)['"]\s*\)/g;

function importsOf(file: string): string[] {
  const text = readFileSync(file, 'utf8');
  return [...text.matchAll(IMPORT)].map((m) => m[1] ?? m[2] ?? m[3] ?? m[4]!);
}

function packageName(specifier: string): string {
  const parts = specifier.split('/');
  return specifier.startsWith('@') ? parts.slice(0, 2).join('/') : parts[0]!;
}

describe('dav-service dependency boundary', () => {
  const files = sourceFiles(join(serviceRoot, 'src'));

  it('finds the sources', () => {
    expect(files.length).toBeGreaterThan(5);
  });

  it.each(files.map((f) => [relative(serviceRoot, f), f]))('%s imports only allowed modules', (_name, file) => {
    for (const specifier of importsOf(file)) {
      if (specifier.startsWith('.')) {
        const target = resolve(dirname(file), specifier);
        expect(relative(join(serviceRoot, 'src'), target).startsWith('..'), `${specifier} leaves src/`).toBe(false);
        continue;
      }
      if (specifier.startsWith('node:') || builtinModules.includes(specifier)) continue;
      expect(specifier).not.toMatch(/demo-site|apps\//);
      expect(allowedPackages.has(packageName(specifier)), `${specifier} is not a declared dependency`).toBe(true);
    }
  });

  it('declares no dependency on the demo site', () => {
    const deps = { ...pkg.dependencies, ...pkg.devDependencies };
    for (const [name, spec] of Object.entries(deps)) {
      expect(name).not.toMatch(/demo-site/);
      expect(spec).not.toMatch(/demo-site|apps\//);
    }
  });
});
