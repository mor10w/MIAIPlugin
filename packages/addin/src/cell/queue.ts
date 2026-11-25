/**
 * Kø for regnearksfunktionen =MIAI.ASK() (#80). Excel kalder funktionen for alle celler på én gang;
 * køen sender kun få til modellen ad gangen, deler ét kald mellem celler med samme forespørgsel og
 * kan stoppes fra task pane.
 */

/** Brugeren trykkede "Stop MIAI-funktioner". */
export class CellStopped extends Error {
  constructor() {
    super("stopped");
    this.name = "CellStopped";
  }
}

/** Excel har afbrudt kaldet (input ændret, cellen slettet). Resultatet bruges ikke. */
export class CellCanceled extends Error {
  constructor() {
    super("canceled");
    this.name = "CellCanceled";
  }
}

export interface QueueState {
  running: number;
  queued: number;
}

interface Job {
  key: string;
  work: (signal: AbortSignal) => Promise<string>;
  abort: AbortController;
  /** Hvor mange celler venter på dette job. */
  waiters: number;
  promise: Promise<string>;
  resolve: (v: string) => void;
  reject: (e: unknown) => void;
  started: boolean;
  /** Hvorfor jobbet blev afbrudt, så fejlen bliver den rigtige. */
  ended?: Error;
}

export interface Ticket {
  promise: Promise<string>;
  /** Denne celle venter ikke længere. Jobbet afbrydes, når ingen celler venter på det. */
  cancel(): void;
}

export class CellQueue {
  private waiting: Job[] = [];
  private running = new Set<Job>();
  private byKey = new Map<string, Job>();
  private listeners = new Set<(s: QueueState) => void>();

  /** `limit` læses, hver gang der er plads, så en ændret indstilling virker med det samme. */
  constructor(private readonly limit: () => number) {}

  get state(): QueueState {
    return { running: this.running.size, queued: this.waiting.length };
  }

  subscribe(fn: (s: QueueState) => void): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  /** Sætter et kald i kø. Kald med samme nøgle, mens det første er i gang, deler resultatet. */
  enqueue(key: string, work: (signal: AbortSignal) => Promise<string>): Ticket {
    let job = this.byKey.get(key);
    if (!job) {
      let resolve!: (v: string) => void;
      let reject!: (e: unknown) => void;
      const promise = new Promise<string>((res, rej) => {
        resolve = res;
        reject = rej;
      });
      // En afbrudt celle har ingen, der venter på løftet; det må ikke blive en ubehandlet fejl.
      promise.catch(() => {});
      job = { key, work, abort: new AbortController(), waiters: 0, promise, resolve, reject, started: false };
      this.byKey.set(key, job);
      this.waiting.push(job);
    }
    const j = job;
    j.waiters++;
    let canceled = false;
    const ticket: Ticket = {
      promise: j.promise,
      cancel: () => {
        if (canceled) return;
        canceled = true;
        if (--j.waiters === 0) this.end(j, new CellCanceled());
      },
    };
    this.pump();
    return ticket;
  }

  /** Tømmer køen og afbryder de kald, der er i gang. Returnerer antallet af stoppede kald. */
  stopAll(): number {
    const jobs = [...this.waiting, ...this.running];
    for (const j of jobs) this.end(j, new CellStopped());
    return jobs.length;
  }

  /** Afslutter et job før tid: ud af køen, eller afbryd kaldet, hvis det er i gang. */
  private end(job: Job, why: Error): void {
    if (job.ended || this.byKey.get(job.key) !== job) return;
    job.ended = why;
    this.byKey.delete(job.key);
    if (job.started) job.abort.abort();
    else this.waiting = this.waiting.filter((j) => j !== job);
    job.reject(why);
    this.emit();
    this.pump();
  }

  private pump(): void {
    const max = Math.max(1, Math.floor(this.limit()) || 1);
    while (this.running.size < max && this.waiting.length) {
      const job = this.waiting.shift()!;
      job.started = true;
      this.running.add(job);
      void job
        .work(job.abort.signal)
        .then(
          (v) => {
            if (!job.ended) job.resolve(v);
          },
          (e) => {
            if (!job.ended) job.reject(e);
          },
        )
        .finally(() => {
          this.running.delete(job);
          if (this.byKey.get(job.key) === job) this.byKey.delete(job.key);
          this.emit();
          this.pump();
        });
    }
    this.emit();
  }

  private emit(): void {
    const s = this.state;
    for (const fn of this.listeners) fn(s);
  }
}
