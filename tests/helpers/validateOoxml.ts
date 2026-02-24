import * as fs from 'fs/promises';
import * as path from 'path';
import * as os from 'os';
import { randomUUID } from 'crypto';
import { spawn } from 'child_process';

function resolveValidatorBinary(): string {
  const platform = process.platform;
  const arch = process.arch;
  let rid: string;
  if (platform === 'darwin' && arch === 'arm64') rid = 'osx-arm64';
  else if (platform === 'darwin' && arch === 'x64') rid = 'osx-x64';
  else if (platform === 'linux' && arch === 'x64') rid = 'linux-x64';
  else if (platform === 'linux' && arch === 'arm64') rid = 'linux-arm64';
  else if (platform === 'win32' && arch === 'x64') rid = 'win-x64';
  else throw new Error(`No OOXML validator binary for platform: ${platform} ${arch}`);

  const binName = rid.startsWith('win-') ? 'ooxml-validator.exe' : 'ooxml-validator';
  return path.join(
    __dirname,
    '..',
    '..',
    'node_modules',
    '@xarsh',
    'ooxml-validator',
    'bin',
    rid,
    binName
  );
}

function runValidator(
  filePath: string
): Promise<{ ok: boolean; errors?: Array<{ description: string }> }> {
  const cmd = resolveValidatorBinary();
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, [filePath], { stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (d: Buffer) => {
      stdout += d.toString();
    });
    child.stderr.on('data', (d: Buffer) => {
      stderr += d.toString();
    });
    child.on('error', (err) =>
      reject(new Error(`Failed to spawn OOXML validator: ${err.message}`))
    );
    child.on('close', (code) => {
      if (code !== 0) {
        return reject(
          new Error(`OOXML Validator exited with code ${code}. stderr: ${stderr || stdout}`)
        );
      }
      try {
        const trimmed = stdout.trim();
        if (!trimmed) return resolve({ ok: false, errors: [] });
        resolve(JSON.parse(trimmed));
      } catch (e: any) {
        reject(new Error(`Failed to parse validator output: ${e.message}\n${stdout}`));
      }
    });
  });
}

export async function validateOoxml(buffer: Buffer): Promise<void> {
  const tempPath = path.join(os.tmpdir(), `ooxml-validate-${randomUUID()}.docx`);
  try {
    await fs.writeFile(tempPath, buffer);
    const result = await runValidator(tempPath);
    if (!result.ok && result.errors && result.errors.length > 0) {
      throw new Error(
        `OOXML validation failed:\n${result.errors.map((e) => `  - ${e.description}`).join('\n')}`
      );
    }
  } finally {
    await fs.unlink(tempPath).catch(() => {});
  }
}
