export function parseTimeToMinutes(input) {
  let value = String(input ?? "").trim();
  if (!value) throw new Error("Duration is empty");

  if (value.includes(":")) {
    const parts = value.split(":");
    if (parts.length !== 2) throw new Error("Invalid time: " + value);
    const hh = parseInt(parts[0].trim(), 10);
    const mm = parseInt(parts[1].trim(), 10);
    if (!Number.isFinite(hh) || !Number.isFinite(mm)) throw new Error("Invalid time: " + value);
    if (mm < 0 || mm > 59) throw new Error("Minutes must be between 0 and 59: " + value);
    return hh * 60 + mm;
  }

  value = value.replace(",", ".");
  const hours = parseFloat(value);
  if (!Number.isFinite(hours)) throw new Error("Invalid time: " + value);
  const minutes = Math.round(hours * 60);
  if (minutes <= 0) throw new Error("Duration must be greater than 0");
  return minutes;
}

export function hhmm(totalMinutes) {
  const hours = Math.floor(totalMinutes / 60);
  const mins = totalMinutes % 60;
  return `${hours}:${String(mins).padStart(2, "0")}`;
}
