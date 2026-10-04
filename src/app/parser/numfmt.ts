/** Number-format classification shared by the .xlsx and .xls readers. */

export const enum Kind { None, Date, Time, DateTime, Percent }

const BUILTIN: Record<number, Kind> = {
  9: Kind.Percent, 10: Kind.Percent,
  14: Kind.Date, 15: Kind.Date, 16: Kind.Date, 17: Kind.Date,
  18: Kind.Time, 19: Kind.Time, 20: Kind.Time, 21: Kind.Time,
  22: Kind.DateTime, 45: Kind.Time, 46: Kind.Time, 47: Kind.Time,
};
for (const id of [27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 57, 58]) BUILTIN[id] = Kind.Date;
for (const id of [32, 33, 55, 56]) BUILTIN[id] = Kind.Time;

export function builtinKind(id: number): Kind {
  return BUILTIN[id] ?? Kind.None;
}

export function kindOfFormat(code: string): Kind {
  const c = code.split(';')[0]
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, '');
  if (c.includes('%')) return Kind.Percent;
  if (/^\s*general\s*$/i.test(c)) return Kind.None;
  const hasTime = /[hs]/i.test(c);
  const hasDate = /[dy]/i.test(c) || (!hasTime && /m/i.test(c));
  if (hasDate && hasTime) return Kind.DateTime;
  if (hasDate) return Kind.Date;
  if (hasTime) return Kind.Time;
  return Kind.None;
}

const pad = (n: number) => (n < 10 ? '0' : '') + n;

export function formatSerial(serial: number, kind: Kind, date1904: boolean): string {
  const days = date1904 ? serial + 1462 : serial;
  const d = new Date(Math.round((days - 25569) * 86400) * 1000);
  if (Number.isNaN(d.getTime())) return formatGeneral(serial);
  const date = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  const time = `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
  if (kind === Kind.Date) return date;
  if (kind === Kind.Time) return time;
  return `${date} ${time}`;
}

export function formatGeneral(n: number): string {
  return Number.isInteger(n) ? String(n) : String(parseFloat(n.toPrecision(15)));
}

export function formatNumber(n: number, kind: Kind, date1904: boolean): string {
  if (kind === Kind.Percent) return formatGeneral(n * 100) + '%';
  if (kind !== Kind.None) return formatSerial(n, kind, date1904);
  return formatGeneral(n);
}
