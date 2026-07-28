export type LogLevel = "info" | "warn" | "error";

export interface Logger {
  info(event: string, details?: Readonly<Record<string, unknown>>): void;
  warn(event: string, details?: Readonly<Record<string, unknown>>): void;
  error(event: string, details?: Readonly<Record<string, unknown>>): void;
}

export class StderrJsonLogger implements Logger {
  info(event: string, details: Readonly<Record<string, unknown>> = {}): void {
    this.write("info", event, details);
  }

  warn(event: string, details: Readonly<Record<string, unknown>> = {}): void {
    this.write("warn", event, details);
  }

  error(event: string, details: Readonly<Record<string, unknown>> = {}): void {
    this.write("error", event, details);
  }

  private write(
    level: LogLevel,
    event: string,
    details: Readonly<Record<string, unknown>>,
  ): void {
    process.stderr.write(
      `${JSON.stringify({ timestamp: new Date().toISOString(), level, event, ...details })}\n`,
    );
  }
}
