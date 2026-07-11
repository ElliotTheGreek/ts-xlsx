/**
 * Excel serial-date conversion — port of openpyxl/utils/datetime.py
 * (from_excel / to_excel). Excel serials are timezone-agnostic day counts, so
 * all math is done in UTC and Dates are constructed/read via their UTC
 * components — no local-timezone drift on round-trip.
 */

/** 1899-12-30 (the 1900 date system's effective epoch, accounting for the bug). */
export const WINDOWS_EPOCH_MS = Date.UTC(1899, 11, 30);
/** 1904-01-01 (the 1904 / Mac date system). */
export const MAC_EPOCH_MS = Date.UTC(1904, 0, 1);

const MS_PER_DAY = 86_400_000;

/** Excel serial → Date. `is1904` selects the 1904 date system. */
export function fromExcel(value: number, is1904 = false): Date {
  const epoch = is1904 ? MAC_EPOCH_MS : WINDOWS_EPOCH_MS;
  let day = Math.floor(value);
  const fraction = value - day;
  const diffMs = Math.round(fraction * MS_PER_DAY);
  // The infamous 1900 leap-year bug: serials 1..59 are shifted by one day.
  if (value > 0 && value < 60 && !is1904) day += 1;
  return new Date(epoch + day * MS_PER_DAY + diffMs);
}

/** Date → Excel serial. `is1904` selects the 1904 date system. */
export function toExcel(date: Date, is1904 = false): number {
  const epoch = is1904 ? MAC_EPOCH_MS : WINDOWS_EPOCH_MS;
  let day = Math.floor((date.getTime() - epoch) / MS_PER_DAY);
  const fraction =
    (date.getUTCHours() * 3600 +
      date.getUTCMinutes() * 60 +
      date.getUTCSeconds() +
      date.getUTCMilliseconds() / 1000) /
    86400;
  if (day > 0 && day <= 60 && !is1904) day -= 1;
  return day + fraction;
}
