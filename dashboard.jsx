// dashboard.jsx
// Personal dashboard: weather, Outlook mail/calendar, sports ticker, news,
// notes (with handwriting), to-dos, AI assistant.
// All state syncs via OneDrive (Microsoft Graph) when signed in; otherwise
// falls back to localStorage.

const { useState, useEffect, useRef, useCallback, useMemo } = React;

// ---------- storage helpers ----------
const LS = {
  get(key, fallback) {
    try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; }
    catch { return fallback; }
  },
  set(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch {}
  }
};

// ---------- time helpers ----------
function fmtTime(d) {
  return new Date(d).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
}
function fmtRelative(d) {
  const t = new Date(d).getTime();
  const diff = Date.now() - t;
  const m = Math.floor(diff / 60000);
  if (m < 1) return "now";
  if (m < 60) return m + "m";
  const h = Math.floor(m / 60);
  if (h < 24) return h + "h";
  const days = Math.floor(h / 24);
  if (days < 7) return days + "d";
  return new Date(d).toLocaleDateString([], { month: "short", day: "numeric" });
}
function fmtDayShort(d) {
  return new Date(d).toLocaleDateString([], { weekday: "short" });
}
function todayISO() { return new Date().toISOString().slice(0, 10); }

// =====================================================================
// MICROSOFT GRAPH (MSAL) — auth, mail, calendar, OneDrive sync
// =====================================================================
const GRAPH_SCOPES = [
  "User.Read",
  "Mail.Read",
  "Calendars.Read",
  "Files.ReadWrite.AppFolder",
  "Tasks.ReadWrite"
];

let msalInstance = null;
function getMsal(clientId) {
  if (!clientId) return null;
  if (msalInstance && msalInstance._cid === clientId) return msalInstance;
  if (!window.msal) return null;
  msalInstance = new window.msal.PublicClientApplication({
    auth: {
      clientId,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: window.location.origin + "/"
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
  });
  msalInstance._cid = clientId;
  msalInstance.initialize().catch(() => {});
  return msalInstance;
}

async function msSignIn(clientId) {
  const m = getMsal(clientId);
  if (!m) throw new Error("MSAL not ready (missing client ID)");
  await m.initialize();
  const result = await m.loginPopup({ scopes: GRAPH_SCOPES, prompt: "select_account" });
  m.setActiveAccount(result.account);
  return result.account;
}

async function msGetToken(clientId) {
  const m = getMsal(clientId);
  if (!m) return null;
  await m.initialize();
  let acct = m.getActiveAccount();
  if (!acct) {
    const accts = m.getAllAccounts();
    if (!accts.length) return null;
    acct = accts[0];
    m.setActiveAccount(acct);
  }
  try {
    const r = await m.acquireTokenSilent({ account: acct, scopes: GRAPH_SCOPES });
    return r.accessToken;
  } catch (e) {
    const r = await m.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    return r.accessToken;
  }
}

async function graphFetch(clientId, path, opts) {
  const token = await msGetToken(clientId);
  if (!token) throw new Error("Not signed in to Microsoft");
  const url = path.startsWith("http") ? path : "https://graph.microsoft.com/v1.0" + path;
  const headers = Object.assign({ Authorization: "Bearer " + token }, (opts && opts.headers) || {});
  const res = await fetch(url, Object.assign({}, opts, { headers }));
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error("Graph " + res.status + ": " + txt.slice(0, 200));
  }
  const ct = res.headers.get("content-type") || "";
  if (ct.includes("application/json")) return res.json();
  return res.text();
}

// =====================================================================
// SETTINGS
// =====================================================================
const SETTINGS_KEY = "dashSettings.v1";
const DEFAULT_SETTINGS = {
  msClientId: "",
  newsApiKey: "",
  newsSources: "bbc-news,reuters,associated-press",
  anthropicKey: "",
  anthropicModel: "claude-sonnet-4-6",
  weatherUnit: "imperial",
  sportsLeagues: ["nfl", "nba", "mlb", "nhl"],
  favTeams: "",
  location: null
};

function useSettings() {
  const [settings, setSettings] = useState(() => Object.assign({}, DEFAULT_SETTINGS, LS.get(SETTINGS_KEY, {})));
  useEffect(() => { LS.set(SETTINGS_KEY, settings); }, [settings]);
  return [settings, setSettings];
}

// =====================================================================
// THEME
// =====================================================================
function useTheme() {
  const [theme, setThemeState] = useState(() => (LS.get("dashTheme", { theme: "dark" }).theme));
  const setTheme = (t) => {
    setThemeState(t);
    LS.set("dashTheme", { theme: t });
    document.documentElement.setAttribute("data-theme", t);
    const meta = document.getElementById("themeColorMeta");
    if (meta) meta.setAttribute("content", t === "light" ? "#f6f5f1" : "#0c1530");
  };
  useEffect(() => { document.documentElement.setAttribute("data-theme", theme); }, [theme]);
  return [theme, setTheme];
}

// =====================================================================
// CLOCK
// =====================================================================
function useClock() {
  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(t);
  }, []);
  return now;
}

// =====================================================================
// WEATHER (Open-Meteo, no key)
// =====================================================================
const WMO_DESC = {
  0: ["Clear", "☀️"], 1: ["Mainly clear", "🌤"], 2: ["Partly cloudy", "⛅"], 3: ["Overcast", "☁️"],
  45: ["Fog", "🌫"], 48: ["Rime fog", "🌫"],
  51: ["Light drizzle", "🌦"], 53: ["Drizzle", "🌦"], 55: ["Heavy drizzle", "🌧"],
  61: ["Light rain", "🌦"], 63: ["Rain", "🌧"], 65: ["Heavy rain", "🌧"],
  71: ["Light snow", "🌨"], 73: ["Snow", "🌨"], 75: ["Heavy snow", "❄️"],
  80: ["Showers", "🌦"], 81: ["Showers", "🌧"], 82: ["Violent showers", "⛈"],
  95: ["Thunderstorm", "⛈"], 96: ["Storm w/ hail", "⛈"], 99: ["Severe storm", "⛈"]
};

function WeatherCard({ settings, setSettings }) {
  const [data, setData] = useState(null);
  const [err, setErr] = useState(null);
  const [loading, setLoading] = useState(false);

  const refresh = useCallback(async () => {
    setLoading(true); setErr(null);
    try {
      let loc = settings.location;
      if (!loc) {
        loc = await new Promise((resolve, reject) => {
          if (!navigator.geolocation) return reject(new Error("Geolocation unavailable"));
          navigator.geolocation.getCurrentPosition(
            (p) => resolve({ lat: p.coords.latitude, lon: p.coords.longitude, name: "Current location" }),
            (e) => reject(new Error(e.message)),
            { timeout: 8000 }
          );
        });
        setSettings(s => ({ ...s, location: loc }));
      }
      const unit = settings.weatherUnit === "metric" ? "celsius" : "fahrenheit";
      const wind = settings.weatherUnit === "metric" ? "kmh" : "mph";
      const url = "https://api.open-meteo.com/v1/forecast"
        + "?latitude=" + loc.lat + "&longitude=" + loc.lon
        + "&current=temperature_2m,apparent_temperature,weather_code,wind_speed_10m,relative_humidity_2m"
        + "&daily=temperature_2m_max,temperature_2m_min,weather_code"
        + "&temperature_unit=" + unit + "&wind_speed_unit=" + wind
        + "&timezone=auto&forecast_days=5";
      const res = await fetch(url);
      if (!res.ok) throw new Error("Weather fetch failed");
      const j = await res.json();
      setData({ loc, ...j });
    } catch (e) {
      setErr(e.message);
    } finally {
      setLoading(false);
    }
  }, [settings.location, settings.weatherUnit, setSettings]);

  useEffect(() => { refresh(); const t = setInterval(refresh, 15 * 60 * 1000); return () => clearInterval(t); }, [refresh]);

  const c = data && data.current;
  const d = data && data.daily;
  const desc = c ? (WMO_DESC[c.weather_code] || ["—", "•"]) : null;
  const unitSym = settings.weatherUnit === "metric" ? "°C" : "°F";

  return (
    <div className="card span-3 row-2">
      <h2>Weather <div className="actions"><button className="icon-btn" onClick={refresh} title="Refresh">↻</button></div></h2>
      {loading && !data && <div className="muted"><span className="spinner"/> Loading…</div>}
      {err && <div className="error">{err}</div>}
      {data && (
        <div className="weather">
          <div>
            <div className="muted">{data.loc.name}</div>
            <div className="now">
              <div className="temp">{Math.round(c.temperature_2m)}°</div>
              <div>
                <div style={{fontSize: 28}}>{desc[1]}</div>
                <div className="desc">{desc[0]}</div>
              </div>
            </div>
            <div className="meta">
              <span>Feels {Math.round(c.apparent_temperature)}{unitSym}</span>
              <span>Wind {Math.round(c.wind_speed_10m)} {settings.weatherUnit === "metric" ? "km/h" : "mph"}</span>
              <span>Hum {c.relative_humidity_2m}%</span>
            </div>
          </div>
          {d && (
            <div className="forecast">
              {d.time.slice(0, 5).map((t, i) => (
                <div className="day" key={t}>
                  <div className="lbl">{i === 0 ? "Today" : fmtDayShort(t)}</div>
                  <div style={{fontSize: 18}}>{(WMO_DESC[d.weather_code[i]] || ["", "•"])[1]}</div>
                  <div className="hi">{Math.round(d.temperature_2m_max[i])}°</div>
                  <div className="lo">{Math.round(d.temperature_2m_min[i])}°</div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// =====================================================================
// SPORTS TICKER (ESPN public site.api — no key)
// =====================================================================
const ESPN_LEAGUES = [
  { id: "nfl",      path: "football/nfl",         label: "NFL" },
  { id: "nba",      path: "basketball/nba",       label: "NBA" },
  { id: "wnba",     path: "basketball/wnba",      label: "WNBA" },
  { id: "mlb",      path: "baseball/mlb",         label: "MLB" },
  { id: "nhl",      path: "hockey/nhl",           label: "NHL" },
  { id: "mls",      path: "soccer/usa.1",         label: "MLS" },
  { id: "epl",      path: "soccer/eng.1",         label: "EPL" },
  { id: "ncaaf",    path: "football/college-football", label: "NCAAF" },
  { id: "ncaam",    path: "basketball/mens-college-basketball", label: "NCAAM" }
];

function SportsTicker({ settings }) {
  const [events, setEvents] = useState([]);
  const [news, setNews] = useState([]);
  const fav = useMemo(() =>
    settings.favTeams.split(",").map(s => s.trim().toLowerCase()).filter(Boolean),
    [settings.favTeams]);

  const refresh = useCallback(async () => {
    const leagues = ESPN_LEAGUES.filter(l => settings.sportsLeagues.includes(l.id));
    const out = [];
    const newsOut = [];
    await Promise.all(leagues.map(async (lg) => {
      try {
        const url = "https://site.api.espn.com/apis/site/v2/sports/" + lg.path + "/scoreboard";
        const res = await fetch(url);
        if (res.ok) {
          const j = await res.json();
          (j.events || []).forEach(ev => {
            const comp = ev.competitions && ev.competitions[0];
            if (!comp) return;
            const home = comp.competitors.find(c => c.homeAway === "home");
            const away = comp.competitors.find(c => c.homeAway === "away");
            if (!home || !away) return;
            const status = ev.status && ev.status.type;
            out.push({
              id: lg.id + "-" + ev.id,
              league: lg.label,
              home: { name: home.team.abbreviation || home.team.shortDisplayName, score: home.score },
              away: { name: away.team.abbreviation || away.team.shortDisplayName, score: away.score },
              state: status ? status.state : "pre",
              short: status ? status.shortDetail : "",
              link: ev.links && ev.links[0] && ev.links[0].href
            });
          });
        }
      } catch {}
      try {
        const newsUrl = "https://site.api.espn.com/apis/site/v2/sports/" + lg.path + "/news?limit=4";
        const nRes = await fetch(newsUrl);
        if (nRes.ok) {
          const nj = await nRes.json();
          (nj.articles || []).slice(0, 3).forEach((art, idx) => {
            const link = (art.links && art.links.web && art.links.web.href)
                       || (art.links && art.links.mobile && art.links.mobile.href);
            const headline = (art.headline || art.title || "").trim();
            if (!headline) return;
            newsOut.push({
              id: lg.id + "-news-" + (art.id || idx),
              league: lg.label,
              headline: headline.length > 120 ? headline.slice(0, 117) + "…" : headline,
              link
            });
          });
        }
      } catch {}
    }));
    const order = { in: 0, pre: 1, post: 2 };
    out.sort((a, b) => (order[a.state] - order[b.state]));
    if (fav.length) {
      out.sort((a, b) => {
        const af = fav.some(f => a.home.name.toLowerCase().includes(f) || a.away.name.toLowerCase().includes(f));
        const bf = fav.some(f => b.home.name.toLowerCase().includes(f) || b.away.name.toLowerCase().includes(f));
        return (bf ? 1 : 0) - (af ? 1 : 0);
      });
    }
    setEvents(out);
    setNews(newsOut);
  }, [settings.sportsLeagues, fav]);

  useEffect(() => { refresh(); const t = setInterval(refresh, 60000); return () => clearInterval(t); }, [refresh]);

  const tickerItems = useMemo(() => {
    const games = events.map(e => ({ ...e, _kind: "game" }));
    const headlines = news.map(n => ({ ...n, _kind: "news" }));
    const merged = [];
    let gi = 0, ni = 0;
    while (gi < games.length || ni < headlines.length) {
      if (gi < games.length) merged.push(games[gi++]);
      if (gi < games.length) merged.push(games[gi++]);
      if (ni < headlines.length) merged.push(headlines[ni++]);
    }
    return merged;
  }, [events, news]);

  if (!tickerItems.length) {
    return (
      <div className="ticker"><span className="label">Sports</span>
        <div className="muted" style={{padding: "0 8px"}}>No games or headlines in selected leagues right now.</div>
      </div>
    );
  }
  const items = tickerItems.concat(tickerItems);
  return (
    <div className="ticker">
      <span className="label">Live · Sports</span>
      <div className="ticker-track">
        {items.map((e, i) => e._kind === "news" ? (
          <a className="ticker-item news" key={e.id + "-" + i} href={e.link} target="_blank" rel="noreferrer">
            <span className="news-tag">{e.league} News</span>
            <span className="headline">{e.headline}</span>
          </a>
        ) : (
          <a className="ticker-item" key={e.id + "-" + i} href={e.link} target="_blank" rel="noreferrer">
            <span className="team">{e.league}:</span>
            <span className="team">{e.away.name}</span>
            <span className="score">{e.away.score}</span>
            <span className="team">@</span>
            <span className="team">{e.home.name}</span>
            <span className="score">{e.home.score}</span>
            <span className={"status " + (e.state === "in" ? "live" : "")}>{e.state === "in" ? "● LIVE " : ""}{e.short}</span>
          </a>
        ))}
      </div>
    </div>
  );
}

// =====================================================================
// MAIL (Outlook)
// =====================================================================
function MailCard({ settings, signedIn, onSignIn }) {
  const [mail, setMail] = useState([]);
  const [err, setErr] = useState(null);
  const [loading, setLoading] = useState(false);

  const load = useCallback(async () => {
    if (!signedIn || !settings.msClientId) return;
    setLoading(true); setErr(null);
    try {
      const j = await graphFetch(settings.msClientId,
        "/me/mailFolders/Inbox/messages?$top=12&$select=id,subject,from,bodyPreview,receivedDateTime,isRead,webLink&$orderby=receivedDateTime desc"
      );
      setMail(j.value || []);
    } catch (e) { setErr(e.message); }
    finally { setLoading(false); }
  }, [signedIn, settings.msClientId]);

  useEffect(() => { load(); const t = setInterval(load, 5 * 60 * 1000); return () => clearInterval(t); }, [load]);

  return (
    <div className="card span-4 row-2">
      <h2>Inbox <span className="badge">Outlook</span>
        <div className="actions">
          {signedIn ? <button className="icon-btn" onClick={load} title="Refresh">↻</button>
                    : <button className="icon-btn primary" onClick={onSignIn}>Sign in</button>}
        </div>
      </h2>
      {!signedIn && <div className="empty">Sign in with Microsoft to view your Outlook inbox.</div>}
      {loading && !mail.length && <div className="muted"><span className="spinner"/> Loading mail…</div>}
      {err && <div className="error">{err}</div>}
      <div className="list">
        {mail.map(m => (
          <a className={"list-item " + (m.isRead ? "read" : "")} key={m.id} href={m.webLink} target="_blank" rel="noreferrer">
            <span className="dot"/>
            <div className="body">
              <div className="meta1">
                <span className="from">{m.from && m.from.emailAddress && (m.from.emailAddress.name || m.from.emailAddress.address)}</span>
                <span className="when">{fmtRelative(m.receivedDateTime)}</span>
              </div>
              <div className="subj">{m.subject || "(no subject)"}</div>
              <div className="preview">{m.bodyPreview}</div>
            </div>
          </a>
        ))}
      </div>
    </div>
  );
}

// =====================================================================
// CALENDAR (Outlook)
// =====================================================================
function CalendarCard({ settings, signedIn, onSignIn }) {
  const [events, setEvents] = useState([]);
  const [err, setErr] = useState(null);
  const [loading, setLoading] = useState(false);

  const load = useCallback(async () => {
    if (!signedIn || !settings.msClientId) return;
    setLoading(true); setErr(null);
    try {
      const start = new Date(); start.setHours(0,0,0,0);
      const end = new Date(start); end.setDate(end.getDate() + 7);
      const path = "/me/calendarView?startDateTime=" + start.toISOString()
        + "&endDateTime=" + end.toISOString()
        + "&$select=id,subject,start,end,location,isAllDay,webLink&$orderby=start/dateTime&$top=20";
      const j = await graphFetch(settings.msClientId, path);
      setEvents(j.value || []);
    } catch (e) { setErr(e.message); }
    finally { setLoading(false); }
  }, [signedIn, settings.msClientId]);

  useEffect(() => { load(); const t = setInterval(load, 10 * 60 * 1000); return () => clearInterval(t); }, [load]);

  const grouped = useMemo(() => {
    const g = {};
    events.forEach(e => {
      const d = new Date(e.start.dateTime + (e.start.timeZone === "UTC" ? "Z" : ""));
      const k = d.toDateString();
      (g[k] = g[k] || []).push({ ...e, _d: d });
    });
    return Object.keys(g).map(k => ({ day: k, items: g[k] }));
  }, [events]);

  return (
    <div className="card span-5 row-2">
      <h2>Upcoming <span className="badge">Calendar</span>
        <div className="actions">
          {signedIn ? <button className="icon-btn" onClick={load}>↻</button>
                    : <button className="icon-btn primary" onClick={onSignIn}>Sign in</button>}
        </div>
      </h2>
      {!signedIn && <div className="empty">Sign in to view your next 7 days.</div>}
      {loading && !events.length && <div className="muted"><span className="spinner"/> Loading events…</div>}
      {err && <div className="error">{err}</div>}
      {!loading && !err && signedIn && !events.length && <div className="empty">No upcoming events in the next 7 days.</div>}
      <div className="list">
        {grouped.map(g => (
          <div key={g.day} style={{marginTop: 6}}>
            <div className="muted" style={{fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em", padding: "4px 10px"}}>
              {new Date(g.day).toLocaleDateString([], { weekday: "long", month: "short", day: "numeric" })}
            </div>
            {g.items.map(ev => (
              <a className="list-item" key={ev.id} href={ev.webLink} target="_blank" rel="noreferrer">
                <span className="dot"/>
                <div className="body">
                  <div className="meta1">
                    <span className="from">{ev.subject}</span>
                    <span className="when">{ev.isAllDay ? "All day" : fmtTime(ev._d)}</span>
                  </div>
                  {ev.location && ev.location.displayName &&
                    <div className="preview">{ev.location.displayName}</div>}
                </div>
              </a>
            ))}
          </div>
        ))}
      </div>
    </div>
  );
}

// =====================================================================
// NEWS — NewsAPI if key set; otherwise Hacker News (free, no key)
// =====================================================================
function NewsCard({ settings }) {
  const [items, setItems] = useState([]);
  const [err, setErr] = useState(null);
  const [loading, setLoading] = useState(false);

  const load = useCallback(async () => {
    setLoading(true); setErr(null);
    try {
      if (settings.newsApiKey) {
        const url = "https://newsapi.org/v2/top-headlines?sources="
          + encodeURIComponent(settings.newsSources) + "&pageSize=15&apiKey=" + settings.newsApiKey;
        const r = await fetch(url);
        if (!r.ok) throw new Error("NewsAPI " + r.status);
        const j = await r.json();
        setItems((j.articles || []).map(a => ({
          title: a.title, source: a.source && a.source.name,
          url: a.url, when: a.publishedAt, image: a.urlToImage
        })));
      } else {
        const idsRes = await fetch("https://hacker-news.firebaseio.com/v0/topstories.json");
        const ids = (await idsRes.json()).slice(0, 15);
        const stories = await Promise.all(ids.map(id =>
          fetch("https://hacker-news.firebaseio.com/v0/item/" + id + ".json").then(r => r.json())
        ));
        setItems(stories.filter(Boolean).map(s => ({
          title: s.title, source: "Hacker News",
          url: s.url || ("https://news.ycombinator.com/item?id=" + s.id),
          when: s.time * 1000, image: null
        })));
      }
    } catch (e) { setErr(e.message); }
    finally { setLoading(false); }
  }, [settings.newsApiKey, settings.newsSources]);

  useEffect(() => { load(); const t = setInterval(load, 15 * 60 * 1000); return () => clearInterval(t); }, [load]);

  return (
    <div className="card span-4 row-3">
      <h2>Headlines <span className="badge">{settings.newsApiKey ? "NewsAPI" : "HN"}</span>
        <div className="actions"><button className="icon-btn" onClick={load}>↻</button></div>
      </h2>
      {loading && !items.length && <div className="muted"><span className="spinner"/> Loading…</div>}
      {err && <div className="error">{err}</div>}
      <div className="news-list">
        {items.map((n, i) => (
          <a className="news-item" key={i} href={n.url} target="_blank" rel="noreferrer">
            {n.image && (
              <img className="news-img" src={n.image} alt="" loading="lazy"
                   onError={(e) => { e.currentTarget.style.display = "none"; }}/>
            )}
            <div className="news-body">
              <div className="src">{n.source} · <span className="when">{fmtRelative(n.when)}</span></div>
              <div className="title">{n.title}</div>
            </div>
          </a>
        ))}
      </div>
    </div>
  );
}

// =====================================================================
// SYNCED STATE (notes + todos) — OneDrive when signed in, else local
// =====================================================================
const SYNC_KEY = "dashSyncedData.v1";
const SYNC_PATH = "/me/drive/special/approot:/dashboard.json";

const EMPTY_DATA = { notes: [], todos: [], _v: 1, _ts: 0 };

function useSyncedData(settings, signedIn) {
  const [data, setData] = useState(() => Object.assign({}, EMPTY_DATA, LS.get(SYNC_KEY, {})));
  const [syncing, setSyncing] = useState(false);
  const [lastSync, setLastSync] = useState(null);
  const saveTimer = useRef(null);
  const pendingSave = useRef(false);

  useEffect(() => {
    if (!signedIn || !settings.msClientId) return;
    let alive = true;
    (async () => {
      try {
        setSyncing(true);
        const j = await graphFetch(settings.msClientId, SYNC_PATH + ":/content").catch(e => {
          if (String(e).includes("404")) return null;
          throw e;
        });
        if (!alive) return;
        if (j && j._ts && j._ts > (data._ts || 0)) {
          setData(j);
          LS.set(SYNC_KEY, j);
        } else if (data._ts) {
          await pushRemote(data);
        }
        setLastSync(new Date());
      } catch (e) {
        console.warn("Sync load failed:", e.message);
      } finally {
        if (alive) setSyncing(false);
      }
    })();
    return () => { alive = false; };
  }, [signedIn, settings.msClientId]);

  async function pushRemote(payload) {
    if (!signedIn || !settings.msClientId) return;
    try {
      setSyncing(true);
      await graphFetch(settings.msClientId, SYNC_PATH + ":/content", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      setLastSync(new Date());
    } catch (e) {
      console.warn("Sync push failed:", e.message);
    } finally {
      setSyncing(false);
    }
  }

  const update = useCallback((updater) => {
    setData(prev => {
      const next = (typeof updater === "function" ? updater(prev) : updater);
      next._ts = Date.now();
      LS.set(SYNC_KEY, next);
      pendingSave.current = next;
      if (saveTimer.current) clearTimeout(saveTimer.current);
      saveTimer.current = setTimeout(() => { if (pendingSave.current) pushRemote(pendingSave.current); }, 1500);
      return next;
    });
  }, [signedIn, settings.msClientId]);

  return { data, update, syncing, lastSync, forceSync: () => pushRemote(data) };
}

// =====================================================================
// HANDWRITING CANVAS
// =====================================================================
function HandwritingCanvas({ onSave, onClose, theme }) {
  const wrapRef = useRef(null);
  const canvasRef = useRef(null);
  const [color, setColor] = useState(theme === "light" ? "#1a1a1a" : "#e9e6d6");
  const [size, setSize] = useState(2.5);
  const drawing = useRef(false);
  const last = useRef({ x: 0, y: 0, p: 0.5 });

  const COLORS = ["#1a1a1a", "#e9e6d6", "#f7d44b", "#c12d2d", "#2f7a3e", "#4a8eff"];

  useEffect(() => {
    const canvas = canvasRef.current;
    const wrap = wrapRef.current;
    function resize() {
      const dpr = window.devicePixelRatio || 1;
      const rect = wrap.getBoundingClientRect();
      canvas.width = rect.width * dpr;
      canvas.height = rect.height * dpr;
      const ctx = canvas.getContext("2d");
      ctx.scale(dpr, dpr);
      ctx.lineCap = "round"; ctx.lineJoin = "round";
    }
    resize();
    window.addEventListener("resize", resize);
    return () => window.removeEventListener("resize", resize);
  }, []);

  function getPos(e) {
    const r = canvasRef.current.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top, p: e.pressure || 0.5 };
  }

  function onPointerDown(e) {
    e.preventDefault();
    canvasRef.current.setPointerCapture(e.pointerId);
    drawing.current = true;
    last.current = getPos(e);
  }
  function onPointerMove(e) {
    if (!drawing.current) return;
    e.preventDefault();
    const p = getPos(e);
    const ctx = canvasRef.current.getContext("2d");
    ctx.strokeStyle = color;
    const width = size * (0.5 + (p.p || 0.5) * 1.5);
    ctx.lineWidth = width;
    ctx.beginPath();
    ctx.moveTo(last.current.x, last.current.y);
    ctx.lineTo(p.x, p.y);
    ctx.stroke();
    last.current = p;
  }
  function onPointerUp(e) {
    drawing.current = false;
    try { canvasRef.current.releasePointerCapture(e.pointerId); } catch {}
  }

  function clear() {
    const c = canvasRef.current;
    c.getContext("2d").clearRect(0, 0, c.width, c.height);
  }

  function save() {
    onSave(canvasRef.current.toDataURL("image/png"));
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <header>
          <h3>Handwriting</h3>
          <div style={{marginLeft: "auto"}} className="muted">Apple Pencil & touch supported</div>
        </header>
        <div className="canvas-tools">
          {COLORS.map(c => (
            <span key={c} className={"swatch " + (c === color ? "active" : "")}
                  style={{background: c}} onClick={() => setColor(c)}/>
          ))}
          <label className="muted" style={{display:"flex", alignItems:"center", gap: 6}}>
            Size <input type="range" min="1" max="10" step="0.5" value={size} onChange={e => setSize(+e.target.value)}/>
          </label>
          <button className="icon-btn" onClick={clear}>Clear</button>
        </div>
        <div className="canvas-wrap" ref={wrapRef}>
          <canvas ref={canvasRef}
            onPointerDown={onPointerDown}
            onPointerMove={onPointerMove}
            onPointerUp={onPointerUp}
            onPointerCancel={onPointerUp}
          />
        </div>
        <footer>
          <button className="icon-btn" onClick={onClose}>Cancel</button>
          <button className="icon-btn primary" onClick={save}>Insert into note</button>
        </footer>
      </div>
    </div>
  );
}

// =====================================================================
// NOTES PANEL — card grid that opens an editor modal
// =====================================================================
function NotesPanel({ data, update, theme }) {
  const [editingId, setEditingId] = useState(null);

  function newNote() {
    const id = "n_" + Math.random().toString(36).slice(2, 9);
    const note = { id, title: "Untitled", html: "", updatedAt: Date.now() };
    update(d => ({ ...d, notes: [note, ...d.notes] }));
    setEditingId(id);
  }
  function delNote(id) {
    update(d => ({ ...d, notes: d.notes.filter(n => n.id !== id) }));
    if (editingId === id) setEditingId(null);
  }
  function patch(id, fields) {
    update(d => ({
      ...d,
      notes: d.notes.map(n => n.id === id ? { ...n, ...fields, updatedAt: Date.now() } : n)
    }));
  }

  const editing = data.notes.find(n => n.id === editingId) || null;
  const sorted = data.notes.slice().sort((a, b) => b.updatedAt - a.updatedAt);

  return (
    <div className="card span-7 row-3">
      <h2>Notes <span className="badge">{data.notes.length}</span>
        <div className="actions">
          <button className="icon-btn primary" onClick={newNote}>＋ New note</button>
        </div>
      </h2>
      <div className="notes-grid">
        <button className="note-card empty-card" onClick={newNote}>＋ New note</button>
        {sorted.map(n => {
          const preview = (n.html || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
          return (
            <button key={n.id} className="note-card" onClick={() => setEditingId(n.id)}>
              <div className="note-card-title">{n.title || "Untitled"}</div>
              <div className="note-card-preview">{preview || "Empty note"}</div>
              <div className="note-card-meta">{fmtRelative(n.updatedAt)}</div>
            </button>
          );
        })}
      </div>
      {editing && (
        <NoteEditorModal
          note={editing}
          theme={theme}
          onPatch={(fields) => patch(editing.id, fields)}
          onDelete={() => { if (confirm("Delete this note?")) { delNote(editing.id); } }}
          onClose={() => setEditingId(null)}
        />
      )}
    </div>
  );
}

function NoteEditorModal({ note, theme, onPatch, onDelete, onClose }) {
  const editorRef = useRef(null);
  const [showCanvas, setShowCanvas] = useState(false);

  useEffect(() => {
    if (editorRef.current) {
      editorRef.current.innerHTML = note.html || "";
      setTimeout(() => { try { editorRef.current && editorRef.current.focus(); } catch {} }, 30);
    }
  }, [note.id]);

  useEffect(() => {
    function onKey(e) { if (e.key === "Escape") onClose(); }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);

  function onEditorInput() { onPatch({ html: editorRef.current.innerHTML }); }
  function exec(cmd, value) {
    document.execCommand(cmd, false, value);
    editorRef.current.focus();
    onEditorInput();
  }
  function insertImage(dataUrl) {
    editorRef.current.focus();
    document.execCommand("insertImage", false, dataUrl);
    onEditorInput();
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal modal-note" onClick={e => e.stopPropagation()}>
        <header>
          <input className="note-title-input" value={note.title}
                 onChange={e => onPatch({ title: e.target.value })}
                 placeholder="Untitled"/>
          <button className="icon-btn" onClick={onDelete} title="Delete">Delete</button>
          <button className="icon-btn primary" onClick={onClose}>Done</button>
        </header>
        <div className="notes-toolbar">
          <button className="icon-btn" onClick={() => exec("bold")}><b>B</b></button>
          <button className="icon-btn" onClick={() => exec("italic")}><i>I</i></button>
          <button className="icon-btn" onClick={() => exec("underline")}><u>U</u></button>
          <button className="icon-btn" onClick={() => exec("insertUnorderedList")}>• List</button>
          <button className="icon-btn" onClick={() => exec("insertOrderedList")}>1. List</button>
          <button className="icon-btn" onClick={() => exec("formatBlock", "h2")}>H</button>
          <button className="icon-btn" onClick={() => exec("formatBlock", "blockquote")}>❝</button>
          <button className="icon-btn" onClick={() => setShowCanvas(true)}>✎ Draw</button>
        </div>
        <div ref={editorRef} className="notes-editor"
             contentEditable suppressContentEditableWarning
             onInput={onEditorInput}/>
        {showCanvas && (
          <HandwritingCanvas
            theme={theme}
            onSave={(d) => { insertImage(d); setShowCanvas(false); }}
            onClose={() => setShowCanvas(false)}
          />
        )}
      </div>
    </div>
  );
}

// =====================================================================
// TODOS
// =====================================================================
function TodoPanel({ data, update }) {
  const [text, setText] = useState("");
  const [filter, setFilter] = useState("active");

  function add() {
    const t = text.trim();
    if (!t) return;
    const todo = { id: "t_" + Math.random().toString(36).slice(2, 9), text: t, done: false, createdAt: Date.now(), due: null };
    update(d => ({ ...d, todos: [todo, ...d.todos] }));
    setText("");
  }
  function toggle(id) {
    update(d => ({ ...d, todos: d.todos.map(t => t.id === id ? { ...t, done: !t.done, doneAt: !t.done ? Date.now() : null } : t) }));
  }
  function del(id) {
    update(d => ({ ...d, todos: d.todos.filter(t => t.id !== id) }));
  }
  function clearDone() {
    update(d => ({ ...d, todos: d.todos.filter(t => !t.done) }));
  }

  const visible = data.todos.filter(t =>
    filter === "all" ? true : filter === "active" ? !t.done : t.done
  );

  return (
    <div className="card span-5 row-3">
      <h2>To-do <span className="badge">{data.todos.filter(t => !t.done).length}</span>
        <div className="actions">
          <select value={filter} onChange={e => setFilter(e.target.value)}
                  style={{background: "var(--bg-card)", color: "var(--fg)", border: "1px solid var(--border)", borderRadius: 8, padding: "4px 8px"}}>
            <option value="active">Active</option>
            <option value="all">All</option>
            <option value="done">Done</option>
          </select>
          <button className="icon-btn" onClick={clearDone}>Clear done</button>
        </div>
      </h2>
      <div className="todo-input">
        <input placeholder="Add a task… (press Enter)"
          value={text} onChange={e => setText(e.target.value)}
          onKeyDown={e => { if (e.key === "Enter") add(); }}/>
        <button className="icon-btn primary" onClick={add}>Add</button>
      </div>
      <div className="todo-list" style={{marginTop: 10}}>
        {visible.length === 0 && <div className="empty">{filter === "active" ? "All caught up." : "Nothing here."}</div>}
        {visible.map(t => (
          <div key={t.id} className={"todo " + (t.done ? "done" : "")}>
            <input type="checkbox" checked={t.done} onChange={() => toggle(t.id)}/>
            <span className="text">{t.text}</span>
            {t.due && <span className="due">{new Date(t.due).toLocaleDateString()}</span>}
            <button className="del" onClick={() => del(t.id)}>✕</button>
          </div>
        ))}
      </div>
    </div>
  );
}

// =====================================================================
// AI ASSISTANT (Anthropic API)
// =====================================================================
function AIPanel({ settings, data, update }) {
  const [messages, setMessages] = useState([
    { role: "system", content: "Assistant ready. Ask me to add tasks, plan your day, summarize notes, or organize what's on your plate." }
  ]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const logRef = useRef(null);

  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight;
  }, [messages, busy]);

  function pushTodos(items) {
    if (!Array.isArray(items) || !items.length) return 0;
    const newOnes = items.map(text => ({
      id: "t_" + Math.random().toString(36).slice(2, 9),
      text: String(text).slice(0, 240), done: false, createdAt: Date.now()
    }));
    update(d => ({ ...d, todos: [...newOnes, ...d.todos] }));
    return newOnes.length;
  }

  function pushNote(title, body) {
    const id = "n_" + Math.random().toString(36).slice(2, 9);
    const html = String(body || "").split(/\n\n+/).map(p => "<p>" + p.replace(/\n/g, "<br>") + "</p>").join("");
    update(d => ({ ...d, notes: [{ id, title: title || "AI note", html, updatedAt: Date.now() }, ...d.notes] }));
    return id;
  }

  async function send() {
    const txt = input.trim();
    if (!txt || busy) return;
    if (!settings.anthropicKey) {
      setMessages(m => [...m, { role: "user", content: txt }, { role: "system", content: "Add an Anthropic API key in Settings to enable the assistant." }]);
      setInput("");
      return;
    }

    const userMsg = { role: "user", content: txt };
    const next = [...messages.filter(m => m.role !== "system"), userMsg];
    setMessages(m => [...m, userMsg]);
    setInput("");
    setBusy(true);

    const context = {
      now: new Date().toISOString(),
      todos_active: data.todos.filter(t => !t.done).slice(0, 30).map(t => t.text),
      todos_done_today: data.todos.filter(t => t.done && t.doneAt && (Date.now() - t.doneAt) < 86400000).map(t => t.text),
      notes_titles: data.notes.slice(0, 20).map(n => n.title),
      weather_unit: settings.weatherUnit
    };

    const system = `You are a concise personal assistant embedded in a dashboard.
Help the user stay organized. You can call tools to add tasks or notes.
Be brief and direct. Prefer bullet lists for plans.

Current context (JSON):
${JSON.stringify(context, null, 2)}`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": settings.anthropicKey,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true"
        },
        body: JSON.stringify({
          model: settings.anthropicModel || "claude-sonnet-4-6",
          max_tokens: 1024,
          system,
          tools: [
            {
              name: "add_todos",
              description: "Add one or more to-do items to the user's list.",
              input_schema: { type: "object", properties: { items: { type: "array", items: { type: "string" } } }, required: ["items"] }
            },
            {
              name: "add_note",
              description: "Create a new note in the user's notes panel.",
              input_schema: { type: "object", properties: { title: { type: "string" }, body: { type: "string" } }, required: ["title", "body"] }
            }
          ],
          messages: next.map(m => ({ role: m.role, content: m.content }))
        })
      });

      if (!res.ok) {
        const t = await res.text();
        throw new Error("API " + res.status + ": " + t.slice(0, 200));
      }
      const j = await res.json();
      let textOut = "";
      const actions = [];
      (j.content || []).forEach(block => {
        if (block.type === "text") textOut += block.text;
        else if (block.type === "tool_use") {
          if (block.name === "add_todos") {
            const n = pushTodos(block.input.items);
            actions.push(`Added ${n} task${n === 1 ? "" : "s"}.`);
          } else if (block.name === "add_note") {
            pushNote(block.input.title, block.input.body);
            actions.push(`Created note "${block.input.title}".`);
          }
        }
      });
      const reply = (textOut.trim() + (actions.length ? "\n\n— " + actions.join(" ") : "")).trim() || "(done)";
      setMessages(m => [...m, { role: "assistant", content: reply }]);
    } catch (e) {
      setMessages(m => [...m, { role: "system", content: "Error: " + e.message }]);
    } finally {
      setBusy(false);
    }
  }

  const suggestions = [
    "Plan my day",
    "Summarize my open tasks",
    "What needs my attention?",
    "Add: pick up groceries, call dentist"
  ];

  return (
    <div className="card span-12">
      <h2>Assistant <span className="badge">Claude</span></h2>
      <div className="ai">
        <div className="ai-log" ref={logRef}>
          {messages.map((m, i) => (
            <div key={i} className={"ai-msg " + m.role}>{m.content}</div>
          ))}
          {busy && <div className="ai-msg assistant"><span className="spinner"/> thinking…</div>}
        </div>
        <div className="ai-suggestions">
          {suggestions.map(s => <button key={s} onClick={() => setInput(s)}>{s}</button>)}
        </div>
        <div className="ai-input">
          <textarea rows="2" placeholder="Ask anything — try 'plan my day' or 'add buy milk'…"
            value={input} onChange={e => setInput(e.target.value)}
            onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }}/>
          <button className="icon-btn primary" onClick={send} disabled={busy}>Send</button>
        </div>
      </div>
    </div>
  );
}

// =====================================================================
// SETTINGS MODAL
// =====================================================================
function SettingsModal({ settings, setSettings, theme, setTheme, onClose, signedIn, onSignIn, onSignOut, account, onForceSync, lastSync, syncing }) {
  const [s, setS] = useState(settings);

  function save() {
    setSettings(s);
    onClose();
  }

  function toggleLeague(id) {
    setS(prev => ({
      ...prev,
      sportsLeagues: prev.sportsLeagues.includes(id)
        ? prev.sportsLeagues.filter(x => x !== id)
        : [...prev.sportsLeagues, id]
    }));
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <header>
          <h3>Settings</h3>
          <div style={{marginLeft: "auto"}}>
            <div className="theme-toggle">
              <button className={theme === "dark" ? "active" : ""} onClick={() => setTheme("dark")}>Dark</button>
              <button className={theme === "light" ? "active" : ""} onClick={() => setTheme("light")}>Light</button>
            </div>
          </div>
        </header>
        <div className="body">
          <div className="settings">

            <div className="section">
              <label>Microsoft Account</label>
              {signedIn ? (
                <div style={{display: "flex", gap: 8, alignItems: "center"}}>
                  <div className="muted">Signed in{account ? " as " + (account.username || account.name) : ""}.</div>
                  <button className="icon-btn" onClick={onForceSync} disabled={syncing}>
                    {syncing ? <><span className="spinner"/> Syncing…</> : "Sync now"}
                  </button>
                  <button className="icon-btn" onClick={onSignOut}>Sign out</button>
                </div>
              ) : (
                <div style={{display: "flex", gap: 8, alignItems: "center"}}>
                  <button className="icon-btn primary" onClick={onSignIn} disabled={!s.msClientId}>Sign in</button>
                  {!s.msClientId && <span className="hint">Add a Client ID below first.</span>}
                </div>
              )}
              {lastSync && <div className="hint">Last sync {fmtRelative(lastSync)}</div>}
            </div>

            <div className="section">
              <label>Microsoft Azure App Client ID</label>
              <input type="text" placeholder="00000000-0000-0000-0000-000000000000"
                value={s.msClientId} onChange={e => setS({ ...s, msClientId: e.target.value })}/>
              <div className="hint">
                Create a free Azure App Registration at portal.azure.com → App registrations → New.
                Set platform to "Single-page application" and add redirect URI: <code>{window.location.origin + "/"}</code>.
                Required permissions: User.Read, Mail.Read, Calendars.Read, Files.ReadWrite.AppFolder, Tasks.ReadWrite.
                Used for Outlook mail, calendar, and cross-device sync of notes/to-dos via OneDrive.
              </div>
            </div>

            <div className="section">
              <label>Anthropic API Key (AI Assistant)</label>
              <input type="password" placeholder="sk-ant-…"
                value={s.anthropicKey} onChange={e => setS({ ...s, anthropicKey: e.target.value })}/>
              <div className="row" style={{marginTop: 6}}>
                <div>
                  <label>Model</label>
                  <select value={s.anthropicModel} onChange={e => setS({ ...s, anthropicModel: e.target.value })}>
                    <option value="claude-opus-4-7">Claude Opus 4.7 (most capable)</option>
                    <option value="claude-sonnet-4-6">Claude Sonnet 4.6 (balanced)</option>
                    <option value="claude-haiku-4-5-20251001">Claude Haiku 4.5 (fastest)</option>
                  </select>
                </div>
              </div>
              <div className="hint">
                Get a key at console.anthropic.com → API keys. Stored locally in your browser only;
                requests go directly to Anthropic from this page.
              </div>
            </div>

            <div className="section">
              <label>NewsAPI Key (optional)</label>
              <input type="password" placeholder="leave blank to use Hacker News"
                value={s.newsApiKey} onChange={e => setS({ ...s, newsApiKey: e.target.value })}/>
              <label style={{marginTop: 6}}>News sources (NewsAPI source IDs, comma-separated)</label>
              <input type="text" value={s.newsSources} onChange={e => setS({ ...s, newsSources: e.target.value })}/>
              <div className="hint">
                Get a free key at newsapi.org. Note: NewsAPI's free Developer plan only allows requests from localhost — for a deployed dashboard, leave this blank and the headlines will fall back to Hacker News, or upgrade NewsAPI to a paid plan.
              </div>
            </div>

            <div className="section">
              <label>Weather Units</label>
              <select value={s.weatherUnit} onChange={e => setS({ ...s, weatherUnit: e.target.value })}>
                <option value="imperial">Imperial (°F, mph)</option>
                <option value="metric">Metric (°C, km/h)</option>
              </select>
              {s.location && (
                <div style={{display: "flex", gap: 8, alignItems: "center", marginTop: 6}}>
                  <span className="muted">Location: {s.location.name} ({s.location.lat.toFixed(2)}, {s.location.lon.toFixed(2)})</span>
                  <button className="icon-btn" onClick={() => setS({ ...s, location: null })}>Re-detect</button>
                </div>
              )}
            </div>

            <div className="section">
              <label>Sports Leagues (ESPN)</label>
              <div style={{display: "flex", flexWrap: "wrap", gap: 6}}>
                {ESPN_LEAGUES.map(lg => (
                  <button key={lg.id}
                    className={"icon-btn " + (s.sportsLeagues.includes(lg.id) ? "primary" : "")}
                    onClick={() => toggleLeague(lg.id)}>
                    {lg.label}
                  </button>
                ))}
              </div>
              <label style={{marginTop: 8}}>Favorite teams (comma-separated, abbreviations or names — pinned to front of ticker)</label>
              <input type="text" placeholder="e.g. KC, LAL, NYY"
                value={s.favTeams} onChange={e => setS({ ...s, favTeams: e.target.value })}/>
            </div>

          </div>
        </div>
        <footer>
          <button className="icon-btn" onClick={onClose}>Cancel</button>
          <button className="icon-btn primary" onClick={save}>Save</button>
        </footer>
      </div>
    </div>
  );
}

// =====================================================================
// LINK VIEWER (in-app browser modal)
// =====================================================================
function LinkViewerModal({ url, onClose }) {
  const [loaded, setLoaded] = useState(false);
  const [blocked, setBlocked] = useState(false);
  const iframeRef = useRef(null);
  const hostname = useMemo(() => {
    try { return new URL(url).hostname.replace(/^www\./, ""); }
    catch { return url; }
  }, [url]);

  function handleLoad() {
    // X-Frame-Options / CSP blocked iframes still fire `load`. Detect by
    // sniffing contentDocument: cross-origin success throws; same-origin
    // empty/about:blank means the browser refused to render the page.
    setTimeout(() => {
      const f = iframeRef.current;
      if (!f) return;
      try {
        const doc = f.contentDocument;
        if (doc !== null) {
          const empty = !doc.body
            || doc.body.children.length === 0
            || (doc.URL === "about:blank" && doc.body.innerHTML.trim() === "");
          if (empty) { setBlocked(true); return; }
        }
      } catch {
        // Cross-origin access denied — actual page loaded successfully.
      }
      setLoaded(true);
    }, 60);
  }

  useEffect(() => {
    const t = setTimeout(() => { if (!loaded && !blocked) setBlocked(true); }, 3500);
    return () => clearTimeout(t);
  }, [loaded, blocked]);

  useEffect(() => {
    function onKey(e) { if (e.key === "Escape") onClose(); }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal modal-browser" onClick={e => e.stopPropagation()}>
        <header>
          <span className="browser-host">{hostname}</span>
          <a className="icon-btn" href={url} target="_blank" rel="noreferrer"
             data-external="1" onClick={(e) => e.stopPropagation()}>Open ↗</a>
          <button className="icon-btn primary" onClick={onClose}>Close</button>
        </header>
        <div className="iframe-wrap">
          {!blocked && (
            <iframe ref={iframeRef} src={url}
              referrerPolicy="no-referrer"
              sandbox="allow-scripts allow-same-origin allow-popups allow-forms allow-popups-to-escape-sandbox"
              onLoad={handleLoad}/>
          )}
          {!loaded && !blocked && (
            <div className="iframe-status">
              <span className="spinner"/>
              <div className="muted">Loading {hostname}…</div>
            </div>
          )}
          {blocked && (
            <div className="iframe-status">
              <div className="blocked-title">Can't preview {hostname}</div>
              <div className="muted" style={{maxWidth: 380}}>
                This site blocks itself from being shown inside other apps. Tap below to open it in your browser.
              </div>
              <a className="icon-btn primary" href={url} target="_blank" rel="noreferrer"
                 data-external="1" onClick={onClose}>Open in browser ↗</a>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// =====================================================================
// APP
// =====================================================================
function App() {
  const [settings, setSettings] = useSettings();
  const [theme, setTheme] = useTheme();
  const [showSettings, setShowSettings] = useState(false);
  const [signedIn, setSignedIn] = useState(false);
  const [account, setAccount] = useState(null);
  const [authError, setAuthError] = useState(null);
  const [linkViewer, setLinkViewer] = useState(null);
  const now = useClock();

  useEffect(() => {
    function onClick(e) {
      if (e.defaultPrevented) return;
      if (e.button !== 0) return;
      if (e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
      let el = e.target;
      while (el && el !== document.body) {
        if (el.tagName === "A") break;
        el = el.parentElement;
      }
      if (!el || el.tagName !== "A") return;
      if (el.getAttribute("data-external") === "1") return;
      const href = el.getAttribute("href") || el.href;
      if (!href || !/^https?:\/\//i.test(href)) return;
      try {
        const u = new URL(href);
        if (u.origin === window.location.origin) return;
        if (u.hostname.includes("login.microsoftonline.com")) return;
        if (u.hostname.includes("login.live.com")) return;
        if (u.hostname.includes("login.windows.net")) return;
      } catch { return; }
      e.preventDefault();
      setLinkViewer(href);
    }
    document.addEventListener("click", onClick, true);
    return () => document.removeEventListener("click", onClick, true);
  }, []);

  useEffect(() => {
    if (!settings.msClientId || !window.msal) return;
    (async () => {
      try {
        const m = getMsal(settings.msClientId);
        if (!m) return;
        await m.initialize();
        const accts = m.getAllAccounts();
        if (accts.length) {
          m.setActiveAccount(accts[0]);
          setAccount(accts[0]);
          setSignedIn(true);
        }
      } catch {}
    })();
  }, [settings.msClientId]);

  async function signIn() {
    setAuthError(null);
    if (!settings.msClientId) {
      setShowSettings(true);
      return;
    }
    try {
      const acct = await msSignIn(settings.msClientId);
      setAccount(acct);
      setSignedIn(true);
    } catch (e) {
      setAuthError(e.message);
    }
  }
  async function signOut() {
    try {
      const m = getMsal(settings.msClientId);
      await m.logoutPopup({ account: account });
    } catch {}
    setAccount(null);
    setSignedIn(false);
  }

  const synced = useSyncedData(settings, signedIn);

  const greeting = useMemo(() => {
    const h = now.getHours();
    if (h < 5) return "Up late";
    if (h < 12) return "Good morning";
    if (h < 18) return "Good afternoon";
    return "Good evening";
  }, [now]);

  return (
    <div className="app">
      <div className="topbar">
        <h1>Dashboard</h1>
        <span className="greeting">{greeting}{account && account.name ? ", " + account.name.split(" ")[0] : ""} · {now.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric" })}</span>
        <span className="clock">{now.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}</span>
        <div className="topbar-actions">
          {synced.syncing && <span className="muted" style={{display:"inline-flex", alignItems:"center", gap:6}}><span className="spinner"/> sync</span>}
          <div className="theme-toggle">
            <button className={theme === "dark" ? "active" : ""} onClick={() => setTheme("dark")}>Dark</button>
            <button className={theme === "light" ? "active" : ""} onClick={() => setTheme("light")}>Light</button>
          </div>
          <button className="icon-btn" onClick={() => setShowSettings(true)}>⚙ Settings</button>
        </div>
      </div>

      <SportsTicker settings={settings}/>

      {authError && <div className="error" style={{padding: "8px 18px"}}>{authError}</div>}

      <div className="grid">
        <WeatherCard settings={settings} setSettings={setSettings}/>
        <MailCard settings={settings} signedIn={signedIn} onSignIn={signIn}/>
        <CalendarCard settings={settings} signedIn={signedIn} onSignIn={signIn}/>
        <NotesPanel data={synced.data} update={synced.update} theme={theme}/>
        <NewsCard settings={settings}/>
        <TodoPanel data={synced.data} update={synced.update}/>
        <AIPanel settings={settings} data={synced.data} update={synced.update}/>
      </div>

      {showSettings && (
        <SettingsModal
          settings={settings} setSettings={setSettings}
          theme={theme} setTheme={setTheme}
          onClose={() => setShowSettings(false)}
          signedIn={signedIn} account={account}
          onSignIn={signIn} onSignOut={signOut}
          onForceSync={synced.forceSync} lastSync={synced.lastSync} syncing={synced.syncing}
        />
      )}

      {linkViewer && (
        <LinkViewerModal url={linkViewer} onClose={() => setLinkViewer(null)}/>
      )}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<App/>);
