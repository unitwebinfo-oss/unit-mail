import net from 'net';
import tls from 'tls';

/**
 * Cliente POP3 mínimo (RFC 1939). O pacote `node-pop3` do npm tem build CJS
 * quebrado, então implementamos o protocolo diretamente — ele é simples:
 * comandos de texto, respostas "+OK"/"-ERR", e respostas multilinha
 * terminadas por uma linha contendo só ".".
 *
 * Toda a leitura passa por UM buffer de bytes: uma espera pendente é ou de
 * uma linha (status/cabeçalho) ou de bytes até um terminador (corpo RETR).
 * Como o protocolo é pedido/resposta, só existe uma espera por vez.
 */
type Pending =
  | { kind: 'line'; resolve: (l: string) => void; reject: (e: Error) => void }
  | { kind: 'bytes'; term: Buffer; resolve: (b: Buffer) => void; reject: (e: Error) => void };

export class Pop3Client {
  private socket: net.Socket | tls.TLSSocket | null = null;
  private buffer = Buffer.alloc(0);
  private pending: Pending | null = null;
  private waitTimer: ReturnType<typeof setTimeout> | null = null;

  /** Tempo máximo esperando dados do servidor (renovado a cada chunk recebido). */
  private static READ_TIMEOUT_MS = 60000;

  private clearTimer(): void {
    if (this.waitTimer) {
      clearTimeout(this.waitTimer);
      this.waitTimer = null;
    }
  }

  private armTimer(): void {
    this.clearTimer();
    this.waitTimer = setTimeout(() => {
      this.waitTimer = null;
      this.failWaiter(new Error('Timeout esperando resposta do servidor POP3'));
    }, Pop3Client.READ_TIMEOUT_MS);
  }

  /** Falha a espera pendente (socket fechado/erro/timeout). */
  private failWaiter(e: Error): void {
    const p = this.pending;
    this.pending = null;
    this.clearTimer();
    p?.reject(e);
  }

  async connect(
    host: string,
    port: number,
    secure: boolean,
    user: string,
    pass: string,
    allowSelfSigned = false,
  ): Promise<void> {
    this.socket = await new Promise((resolve, reject) => {
      const sock: net.Socket | tls.TLSSocket = secure
        ? tls.connect({ host, port, servername: host, rejectUnauthorized: !allowSelfSigned })
        : net.connect({ host, port });
      const onErr = (e: Error) => reject(e);
      sock.once('error', onErr);
      sock.once(secure ? 'secureConnect' : 'connect', () => {
        sock.off('error', onErr);
        resolve(sock);
      });
      sock.setTimeout(30000, () => {
        sock.destroy(new Error('Timeout na conexão POP3'));
      });
    });
    this.socket!.on('data', (d) => this.onData(d));
    this.socket!.on('error', (e) => this.failWaiter(e));
    this.socket!.on('close', () => this.failWaiter(new Error('Conexão POP3 encerrada pelo servidor')));
    this.socket!.on('end', () => this.failWaiter(new Error('Conexão POP3 encerrada')));
    // Banner de boas-vindas
    const banner = await this.readLine();
    if (!banner.startsWith('+OK')) throw new Error('Servidor POP3 não respondeu +OK');

    await this.expectOk(`USER ${user}`);
    await this.expectOk(`PASS ${pass}`);
  }

  private onData(d: Buffer): void {
    this.buffer = Buffer.concat([this.buffer, d]);
    if (this.pending) this.armTimer();
    this.pump();
  }

  /** Entrega ao aguardador pendente o que já está no buffer, se completo. */
  private pump(): void {
    const p = this.pending;
    if (!p) return;
    if (p.kind === 'line') {
      const idx = this.buffer.indexOf('\r\n');
      if (idx === -1) return;
      const line = this.buffer.subarray(0, idx).toString('utf-8');
      this.buffer = this.buffer.subarray(idx + 2);
      this.pending = null;
      this.clearTimer();
      p.resolve(line);
    } else {
      const idx = this.buffer.indexOf(p.term);
      if (idx === -1) return;
      const out = this.buffer.subarray(0, idx);
      this.buffer = this.buffer.subarray(idx + p.term.length);
      this.pending = null;
      this.clearTimer();
      p.resolve(out);
    }
  }

  private readLine(): Promise<string> {
    return new Promise((resolve, reject) => {
      this.pending = { kind: 'line', resolve, reject };
      this.armTimer();
      this.pump();
    });
  }

  /** Lê bytes crus até o terminador (corpo de RETR), sem decodificar. */
  private readBytesUntil(term: Buffer): Promise<Buffer> {
    return new Promise((resolve, reject) => {
      this.pending = { kind: 'bytes', term, resolve, reject };
      this.armTimer();
      this.pump();
    });
  }

  private write(line: string): void {
    this.socket!.write(line + '\r\n');
  }

  private async expectOk(line: string): Promise<string> {
    this.write(line);
    const res = await this.readLine();
    if (!res.startsWith('+OK')) {
      const cmd = line.split(' ')[0];
      const hint =
        cmd === 'PASS' || cmd === 'USER'
          ? ' — verifique usuário e senha da conta em Configurações → Email'
          : '';
      throw new Error(`POP3 falhou: ${cmd} → ${res}${hint}`);
    }
    return res;
  }

  /** Resposta de linha única (STAT, etc.) */
  private async singleLine(cmd: string): Promise<string> {
    return this.expectOk(cmd);
  }

  /** Resposta multilinha terminada por "." (LIST, UIDL, TOP) */
  private async multiLine(cmd: string): Promise<string[]> {
    await this.expectOk(cmd);
    const lines: string[] = [];
    for (;;) {
      const line = await this.readLine();
      if (line === '.') break;
      // dot-stuffing: linha iniciada com "." tem um "." extra
      lines.push(line.startsWith('.') ? line.slice(1) : line);
    }
    return lines;
  }

  /** Total de mensagens na caixa */
  async stat(): Promise<number> {
    const res = await this.singleLine('STAT');
    return parseInt(res.split(/\s+/)[1] || '0', 10);
  }

  /** Mapa número da mensagem → UIDL */
  async uidl(): Promise<Array<{ num: number; uidl: string }>> {
    const lines = await this.multiLine('UIDL');
    return lines
      .map((l) => {
        const [num, uidl] = l.trim().split(/\s+/);
        return { num: parseInt(num, 10), uidl };
      })
      .filter((x) => !isNaN(x.num) && x.uidl);
  }

  /** Cabeçalhos da mensagem (TOP n 0) */
  async top(num: number): Promise<string> {
    return (await this.multiLine(`TOP ${num} 0`)).join('\n');
  }

  /** Mensagem RFC822 completa — bytes crus (não decodifica: preserva 8bit/anexos) */
  async retr(num: number): Promise<Buffer> {
    await this.expectOk(`RETR ${num}`);
    const raw = await this.readBytesUntil(Buffer.from('\r\n.\r\n'));
    // dot-unstuffing: qualquer linha do corpo iniciada com ".." perde o 1º ponto.
    // Feito em latin1 (mapeia byte↔char 1:1) para não corromper bytes não-UTF8.
    return Buffer.from(raw.toString('latin1').replace(/^\.\./gm, '.'), 'latin1');
  }

  /** Marca mensagem para exclusão (aplicada no QUIT). */
  async dele(num: number): Promise<void> {
    await this.expectOk(`DELE ${num}`);
  }

  /** Aborta o comando em andamento (cancelamento do sync): rejeita a espera
   *  pendente na hora e derruba o socket sem QUIT. Diferente do timeout de
   *  leitura, o usuário não precisa esperar o servidor responder. */
  abort(): void {
    this.failWaiter(new Error('sincronização cancelada'));
    try {
      this.socket?.destroy();
    } catch {
      // já morto
    }
    this.socket = null;
  }

  async quit(): Promise<void> {
    try {
      await this.expectOk('QUIT');
    } catch {
      // ignora
    }
    this.clearTimer();
    try {
      this.socket?.destroy();
    } catch {
      // ignora
    }
    this.socket = null;
  }
}
