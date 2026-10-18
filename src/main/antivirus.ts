import { spawn, execFile, type ChildProcess } from 'child_process';
import { existsSync, readdirSync } from 'fs';
import { writeFile, readFile, unlink, mkdtemp, rmdir } from 'fs/promises';
import { tmpdir } from 'os';
import { join } from 'path';
import { createInterface } from 'readline';
import { log } from './logger';

export interface AvResult {
  status: 'clean' | 'infected' | 'unavailable' | 'error';
  /** Nome da ameaça quando o motor informa (ClamAV/MpCmdRun; AMSI não informa). */
  threat?: string;
  /** Motor que respondeu: 'amsi' | 'defender' | 'clamd' | 'clamscan'. */
  engine?: string;
}

// ---------------- Windows: AMSI (qualquer AV registrado como provedor) ----------------
// AmsiScanBuffer é a interface oficial que antivírus locais implementam —
// Windows Defender, Avast, Kaspersky etc. Um processo PowerShell persistente
// inicializa o AMSI uma vez e responde "CLEAN"/"INFECTED" por arquivo,
// sem o custo de abrir um processo novo por mensagem.

const AMSI_PS1 = `
$src = @'
using System;
using System.Runtime.InteropServices;
public static class AmsiNative {
    [DllImport("amsi.dll", CallingConvention=CallingConvention.StdCall)]
    public static extern int AmsiInitialize([MarshalAs(UnmanagedType.LPWStr)] string appName, out IntPtr amsiContext);
    [DllImport("amsi.dll", CallingConvention=CallingConvention.StdCall)]
    public static extern void AmsiUninitialize(IntPtr amsiContext);
    [DllImport("amsi.dll", CallingConvention=CallingConvention.StdCall)]
    public static extern int AmsiScanBuffer(IntPtr amsiContext, byte[] buffer, uint length, [MarshalAs(UnmanagedType.LPWStr)] string contentName, IntPtr session, out uint result);
}
'@
try { Add-Type -TypeDefinition $src } catch { Write-Output 'INITFAIL'; exit 1 }
$ctx = [IntPtr]::Zero
try { $hr = [AmsiNative]::AmsiInitialize('UnitMail', [ref]$ctx) } catch { Write-Output 'INITFAIL'; exit 1 }
if ($hr -ne 0 -or $ctx -eq [IntPtr]::Zero) { Write-Output 'INITFAIL'; exit 1 }
Write-Output 'READY'
while ($true) {
  $line = [Console]::In.ReadLine()
  if ($null -eq $line -or $line -eq 'QUIT') { break }
  try {
    $bytes = [IO.File]::ReadAllBytes($line)
    $r = [uint32]0
    [void][AmsiNative]::AmsiScanBuffer($ctx, $bytes, [uint32]$bytes.Length, 'unitmail-message.eml', [IntPtr]::Zero, [ref]$r)
    # AMSI_RESULT >= 0x8000 (32768) = ameaça detectada
    if ($r -ge 32768) { Write-Output 'INFECTED' } else { Write-Output 'CLEAN' }
  } catch { Write-Output 'SCANERR' }
}
if ($ctx -ne [IntPtr]::Zero) { [AmsiNative]::AmsiUninitialize($ctx) | Out-Null }
`;

class AmsiScanner {
  private proc: ChildProcess | null = null;
  private pending: Array<(line: string) => void> = [];
  private scriptPath: string | null = null;
  /** null = nunca tentou; true/false = AMSI disponível ou não nesta sessão. */
  private ok: boolean | null = null;

  private async ensure(): Promise<boolean> {
    if (this.ok !== null) return this.ok;
    try {
      const dir = await mkdtemp(join(tmpdir(), 'unitmail-av-'));
      this.scriptPath = join(dir, 'amsi.ps1');
      await writeFile(this.scriptPath, AMSI_PS1, 'utf-8');
      const proc = spawn(
        'powershell',
        ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', this.scriptPath],
        { stdio: ['pipe', 'pipe', 'ignore'], windowsHide: true },
      );
      this.proc = proc;
      const rl = createInterface({ input: proc.stdout! });
      const ready = new Promise<boolean>((resolve) => {
        let settled = false;
        rl.on('line', (line) => {
          const l = line.trim();
          if (!settled) {
            settled = true;
            resolve(l === 'READY');
            if (l !== 'READY') return;
          } else {
            const p = this.pending.shift();
            if (p) p(l);
          }
        });
        rl.on('close', () => {
          if (!settled) resolve(false);
          this.ok = false;
          while (this.pending.length) this.pending.shift()?.('SCANERR');
        });
        setTimeout(() => {
          if (!settled) resolve(false);
        }, 15000);
      });
      proc.on('error', () => {
        this.ok = false;
      });
      this.ok = await ready;
      if (!this.ok) {
        try {
          proc.kill();
        } catch {
          // já morto
        }
        this.proc = null;
      }
      return this.ok;
    } catch {
      this.ok = false;
      return false;
    }
  }

  /** 'CLEAN' | 'INFECTED' | 'INITFAIL' | 'SCANERR' */
  async scanFile(path: string): Promise<string> {
    if (!(await this.ensure())) return 'INITFAIL';
    return new Promise((resolve) => {
      this.pending.push(resolve);
      try {
        this.proc!.stdin!.write(path + '\n');
      } catch {
        this.pending.pop();
        this.ok = false;
        resolve('SCANERR');
      }
    });
  }
}

const amsi = new AmsiScanner();

// ---------------- Windows: antivírus de terceiros (ESET, Avast…) ----------------

let wmiProducts: string[] | null = null;

/**
 * Antivírus registrados na Central de Segurança do Windows
 * (root\SecurityCenter2 → AntiVirusProduct) — ESET, Avast, Kaspersky,
 * Defender e qualquer outro produto instalado aparecem aqui.
 */
function wmiAvProducts(): Promise<string[]> {
  if (wmiProducts) return Promise.resolve(wmiProducts);
  return new Promise((resolve) => {
    execFile(
      'powershell',
      [
        '-NoProfile',
        '-NonInteractive',
        '-Command',
        'Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct ' +
          '-ErrorAction SilentlyContinue | Select-Object -ExpandProperty displayName',
      ],
      { timeout: 15000, windowsHide: true },
      (err, stdout) => {
        wmiProducts = err
          ? []
          : [...new Set(String(stdout).split(/\r?\n/).map((s) => s.trim()).filter(Boolean))];
        resolve(wmiProducts);
      },
    );
  });
}

/**
 * Scanner de linha de comando do ESET (ecls.exe) — presente em NOD32
 * Antivirus / ESET Endpoint. Sai 0 = limpo, 1 = ameaça encontrada.
 */
function findEcls(): string | null {
  const roots = [
    process.env['ProgramFiles'] || 'C:\\Program Files',
    process.env['ProgramFiles(x86)'],
  ];
  for (const r of roots) {
    if (!r) continue;
    const eset = join(r, 'ESET');
    if (!existsSync(eset)) continue;
    try {
      for (const d of readdirSync(eset)) {
        const p = join(eset, d, 'ecls.exe');
        if (existsSync(p)) return p;
      }
    } catch {
      // pasta parcialmente acessível — segue
    }
  }
  return null;
}

function eclsScan(path: string): Promise<AvResult> {
  const exe = findEcls();
  if (!exe) return Promise.resolve({ status: 'unavailable' });
  return new Promise((resolve) => {
    execFile(
      exe,
      ['-no-quarantine', '-no-logall', path],
      { timeout: 120000, windowsHide: true },
      (err, stdout) => {
        const out = String(stdout || '');
        const code = (err as NodeJS.ErrnoException & { code?: number })?.code;
        if (!err || code === 0) return resolve({ status: 'clean', engine: 'eset' });
        // exit 1 = ameaça encontrada; a saída traz threat="nome"
        if (code === 1 || /threat="/i.test(out)) {
          const m = out.match(/threat="([^"]+)"/i);
          return resolve({ status: 'infected', threat: m?.[1], engine: 'eset' });
        }
        resolve({ status: 'error', engine: 'eset' });
      },
    );
  });
}

/**
 * Fallback genérico para qualquer AV residente (ESET, Avast etc.): a
 * proteção em tempo real inspeciona o arquivo no momento em que é gravado.
 * Se ele desaparece (foi para quarentena) ou a leitura é negada → ameaça.
 * Arquivo presente e legível → tratado como limpo (o AV não reagiu).
 */
async function onAccessScan(path: string): Promise<AvResult> {
  await new Promise((r) => setTimeout(r, 800)); // dá tempo do residente agir
  if (!existsSync(path)) return { status: 'infected', engine: 'onaccess' };
  try {
    await readFile(path);
    return { status: 'clean', engine: 'onaccess' };
  } catch {
    return { status: 'infected', engine: 'onaccess' };
  }
}

function findMpCmdRun(): string | null {
  try {
    const base = 'C:\\ProgramData\\Microsoft\\Windows Defender\\Platform';
    if (!existsSync(base)) return null;
    const dirs = readdirSync(base)
      .filter((d) => existsSync(join(base, d, 'MpCmdRun.exe')))
      .sort()
      .reverse(); // versão mais nova primeiro
    return dirs.length ? join(base, dirs[0], 'MpCmdRun.exe') : null;
  } catch {
    return null;
  }
}

function mpCmdScan(path: string): Promise<AvResult> {
  const exe = findMpCmdRun();
  if (!exe) return Promise.resolve({ status: 'unavailable' });
  return new Promise((resolve) => {
    execFile(
      exe,
      ['-Scan', '-ScanType', '3', '-File', path, '-DisableRemediation'],
      { timeout: 60000, windowsHide: true },
      (err, stdout) => {
        const out = String(stdout || '');
        // Código 0 = sem ameaças; código 2 / texto "threat" = ameaça encontrada
        if (!err) return resolve({ status: 'clean', engine: 'defender' });
        const m = out.match(/Threat\s*:\s*(.+)/i) || out.match(/found\s+(.+)/i);
        if (/threat|virus|malware/i.test(out)) {
          return resolve({ status: 'infected', threat: m?.[1]?.trim(), engine: 'defender' });
        }
        resolve({ status: 'error', engine: 'defender' });
      },
    );
  });
}

// ---------------- Linux: ClamAV (clamdscan ou clamscan) ----------------

function which(cmd: string): Promise<boolean> {
  return new Promise((resolve) => {
    execFile('which', [cmd], { timeout: 5000 }, (err) => resolve(!err));
  });
}

let clamCmd: 'clamdscan' | 'clamscan' | null | undefined;

async function clamScan(path: string): Promise<AvResult> {
  if (clamCmd === undefined) {
    clamCmd = (await which('clamdscan')) ? 'clamdscan' : (await which('clamscan')) ? 'clamscan' : null;
  }
  if (!clamCmd) return { status: 'unavailable' };
  const args =
    clamCmd === 'clamdscan' ? ['--fdpass', '--no-summary', path] : ['--no-summary', path];
  return new Promise((resolve) => {
    execFile(clamCmd!, args, { timeout: 120000 }, (err, stdout) => {
      const out = String(stdout || '');
      if (!err) return resolve({ status: 'clean', engine: clamCmd! });
      // exit 1 = vírus encontrado ("arquivo: NomeDoVirus FOUND"); 2 = erro real
      if (err.code === 1 || /FOUND/i.test(out)) {
        const m = out.match(/:\s*(.+?)\s+FOUND/i);
        return resolve({ status: 'infected', threat: m?.[1], engine: clamCmd! });
      }
      resolve({ status: 'error', engine: clamCmd! });
    });
  });
}

// ---------------- API ----------------

let detected: { available: boolean; engine: string } | null = null;

/**
 * Detecta o antivírus local uma vez por sessão:
 *  - Windows: AMSI (interface que ESET, Avast, Defender e outros AVs
 *    registrados implementam) → ecls.exe (ESET) → MpCmdRun (Defender) →
 *    proteção em tempo real de qualquer AV listado na Central de Segurança.
 *  - Linux: ClamAV (clamdscan/clamscan).
 * Outras plataformas: indisponível (o caller segue sem verificação).
 */
export async function avStatus(): Promise<{ available: boolean; engine: string }> {
  if (detected) return detected;
  if (process.platform === 'win32') {
    const products = await wmiAvProducts();
    const thirdParty = products.find((p) => !/defender/i.test(p));
    const avName = thirdParty || products[0];
    // Sonda real: escreve um arquivo e manda o AMSI verificar — prova que a
    // interface está operante (o AV instalado responde por ela).
    const probe = join(tmpdir(), `unitmail-probe-${Date.now()}.eml`);
    await writeFile(probe, 'probe').catch(() => {});
    const amsiOk = await amsi
      .scanFile(probe)
      .then((r) => r !== 'INITFAIL')
      .catch(() => false);
    unlink(probe).catch(() => {});
    if (amsiOk) {
      detected = {
        available: true,
        engine: avName ? `${avName} (verificação do sistema)` : 'antivírus do sistema',
      };
    } else if (findEcls()) {
      detected = { available: true, engine: 'ESET (verificação por demanda)' };
    } else if (findMpCmdRun()) {
      detected = { available: true, engine: 'Microsoft Defender' };
    } else if (products.length) {
      // AV registrado sem interface de demanda — a proteção em tempo real
      // inspeciona arquivos na gravação/leitura.
      detected = { available: true, engine: `${avName} (proteção em tempo real)` };
    } else {
      detected = { available: false, engine: 'nenhum antivírus detectado' };
    }
  } else if (process.platform === 'linux') {
    if (await which('clamdscan')) detected = { available: true, engine: 'ClamAV (daemon)' };
    else if (await which('clamscan')) detected = { available: true, engine: 'ClamAV' };
    else detected = { available: false, engine: 'antivírus não encontrado (instale o ClamAV)' };
  } else {
    detected = { available: false, engine: 'sem suporte nesta plataforma' };
  }
  return detected;
}

/**
 * Motor de verificação decidido UMA vez por sessão: a primeira mensagem
 * percorre a cadeia (AMSI → ESET → Defender → proteção em tempo real) e as
 * seguintes vão direto no motor escolhido. Sem isso, cada email tentava a
 * cadeia inteira — e um MpCmdRun lento ou timeout de 60s travava o sync por
 * minutos na primeira mensagem.
 */
let scanEngine: 'amsi' | 'eset' | 'defender' | 'onaccess' | 'none' | null = null;
let scanErrCount = 0;

/** Disjuntor: motor falhando repetidamente desliga a verificação na sessão. */
function trackErr(res: AvResult): AvResult {
  if (res.status === 'error' && ++scanErrCount >= 5) scanEngine = 'none';
  return res;
}

export async function scanEmailBuffer(buf: Buffer, suggestedName = 'message.eml'): Promise<AvResult> {
  if (scanEngine === 'none') return { status: 'unavailable' };
  const dir = await mkdtemp(join(tmpdir(), 'unitmail-av-'));
  const file = join(dir, suggestedName.replace(/[<>:"/\\|?*]/g, '_'));
  try {
    await writeFile(file, buf);
    if (process.platform === 'win32') {
      // Motor decidido → vai direto nele, sem percorrer a cadeia
      if (scanEngine === 'amsi') {
        const r = await amsi.scanFile(file);
        if (r === 'INFECTED') return { status: 'infected', engine: 'amsi' };
        if (r === 'CLEAN') return { status: 'clean', engine: 'amsi' };
        scanEngine = null; // AMSI morreu no meio — decide de novo
      } else if (scanEngine === 'eset') {
        return trackErr(await eclsScan(file));
      } else if (scanEngine === 'defender') {
        return trackErr(await mpCmdScan(file));
      } else if (scanEngine === 'onaccess') {
        return trackErr(await onAccessScan(file));
      }
      // Primeira mensagem: percorre a cadeia e trava a decisão
      const r = await amsi.scanFile(file);
      if (r === 'CLEAN' || r === 'INFECTED') {
        scanEngine = 'amsi';
        return { status: r === 'INFECTED' ? 'infected' : 'clean', engine: 'amsi' };
      }
      // AMSI falhou — tenta o scanner do ESET, depois o Defender por demanda
      const es = await eclsScan(file);
      if (es.status !== 'unavailable') {
        scanEngine = 'eset';
        return trackErr(es);
      }
      const mp = await mpCmdScan(file);
      if (mp.status !== 'unavailable') {
        scanEngine = 'defender';
        return trackErr(mp);
      }
      // Qualquer AV residente registrado na Central de Segurança cobre por
      // proteção em tempo real: se o arquivo sumir/ficar ilegível → ameaça.
      if ((await wmiAvProducts()).length) {
        scanEngine = 'onaccess';
        return trackErr(await onAccessScan(file));
      }
      scanEngine = 'none';
      return { status: 'unavailable' };
    }
    if (process.platform === 'linux') return await clamScan(file);
    return { status: 'unavailable' };
  } catch (e) {
    log.error('antivirus: scan falhou', e);
    return { status: 'error' };
  } finally {
    unlink(file).then(() => rmdir(dir)).catch(() => {});
  }
}
