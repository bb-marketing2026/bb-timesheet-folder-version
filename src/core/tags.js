// Client tag: the client's short name at the start of each Project Description in the timesheet
// (for example "RAWCS" in "RAWCS Content Hub"). People pick a generic activity tag such as
// "Content Planning" and the app writes "RAWCS Content Planning".

const lc = (s) => String(s || '').trim().toLowerCase();

/**
 * Best guess from the timesheet: the client's short name, only if most existing tags start with it.
 * Candidates come from the Department / Client cell: "RAWCS - Rotary Australia ..." gives "RAWCS",
 * plus the initials and first word of the name. Timesheets whose tags do not start with the client
 * (for example Weld East: "Website Management") get no client tag.
 */
export function detectClientTag(info) {
  const client = String(info?.client || '').trim();
  const tags = info?.tags || [];
  const short = client.split(' - ')[0].trim();
  const words = client.replace(/[^A-Za-z0-9 ]+/g, ' ').split(/\s+/).filter(Boolean);
  const initials = words.filter((w) => /^[A-Z]/.test(w)).map((w) => w[0]).join('');
  const candidates = [...new Set([short, initials, words[0]].filter((c) => c && c.length <= 20))];
  if (!tags.length) return short.length <= 12 && !short.includes(' ') ? short : '';
  const starting = (p) => tags.filter((t) => lc(t).startsWith(`${lc(p)} `)).length;
  return candidates.find((c) => starting(c) >= Math.ceil(tags.length / 2)) || '';
}

/** "RAWCS Content Hub" -> "Content Hub" when the client tag is "RAWCS". */
export function stripClient(tag, client) {
  const t = String(tag || '').trim();
  if (!client) return t;
  return lc(t).startsWith(`${lc(client)} `) ? t.slice(client.trim().length).trim() : t;
}

/** "Content Planning" -> "RAWCS Content Planning". Leaves tags that already start with the client alone. */
export function withClient(tag, client) {
  const t = String(tag || '').trim();
  const c = String(client || '').trim();
  if (!c || !t) return t;
  if (lc(t) === lc(c) || lc(t).startsWith(`${lc(c)} `)) return t;
  return `${c} ${t}`;
}

/** Generic activity tags for the picker: existing tags without the client part, most used first. */
export function genericTags(tags, client) {
  const seen = new Set();
  const out = [];
  for (const t of tags || []) {
    const g = stripClient(t, client);
    if (g && !seen.has(lc(g))) { seen.add(lc(g)); out.push(g); }
  }
  return out;
}
