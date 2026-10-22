// Client tag: the client's short name at the start of each Project Description in the timesheet
// (for example "RAWCS" in "RAWCS Content Hub"). People pick a generic activity tag such as
// "Content Planning" and the app writes "RAWCS Content Planning".

const lc = (s) => String(s || '').trim().toLowerCase();

/** Best guess from the timesheet: the client's short name, if most existing tags start with it. */
export function detectClientTag(info) {
  const fromClient = String(info?.client || '').split(' - ')[0].trim();
  const tags = info?.tags || [];
  if (!tags.length) return fromClient.length <= 20 ? fromClient : '';
  const starting = (p) => tags.filter((t) => lc(t).startsWith(`${lc(p)} `)).length;
  if (fromClient && starting(fromClient) >= Math.ceil(tags.length / 2)) return fromClient;
  const counts = new Map();
  for (const t of tags) { const w = t.trim().split(/\s+/)[0]; if (w && t.trim().includes(' ')) counts.set(w, (counts.get(w) || 0) + 1); }
  const [best, n] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0] || [];
  return best && n >= Math.ceil(tags.length / 2) ? best : '';
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
