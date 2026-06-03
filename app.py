import streamlit as st
import pandas as pd
import requests
import json
import re
import io
from datetime import date, datetime
import plotly.graph_objects as go

# ── Konfigurace ───────────────────────────────────────────────────────────────
CLIENT_ID  = "9c5552a6-4492-4e81-b96b-d1ea74abb8ed"
AUTHORITY  = "https://login.microsoftonline.com/consumers/oauth2/v2.0"
SCOPES     = "Calendars.ReadWrite offline_access"
TOKEN_FILE = "nafta_token.json"
KAPACITA_NADRZE = 7000  # litrů

# ── Autentizace ───────────────────────────────────────────────────────────────

def _gist_config() -> tuple:
    """Vrátí (github_token, gist_id) ze secrets, nebo (None, None)."""
    try:
        g = st.secrets.get("gist", {})
        gt = g.get("github_token")
        gi = g.get("gist_id")
        if gt and gi:
            return gt, gi
    except Exception:
        pass
    return None, None

def _load_gist_tokens() -> dict:
    github_token, gist_id = _gist_config()
    if not github_token:
        return {}
    try:
        r = requests.get(
            f"https://api.github.com/gists/{gist_id}",
            headers={"Authorization": f"token {github_token}",
                     "Accept": "application/vnd.github.v3+json"},
            timeout=5,
        )
        if r.status_code == 200:
            content = r.json()["files"]["nafta_token.json"]["content"]
            return json.loads(content)
    except Exception:
        pass
    return {}

def _save_gist_tokens(data: dict):
    github_token, gist_id = _gist_config()
    if not github_token:
        return
    try:
        requests.patch(
            f"https://api.github.com/gists/{gist_id}",
            headers={"Authorization": f"token {github_token}",
                     "Accept": "application/vnd.github.v3+json"},
            json={"files": {"nafta_token.json": {"content": json.dumps(data)}}},
            timeout=5,
        )
    except Exception:
        pass

def _load_tokens() -> dict:
    # 1) Čerstvě obnovené tokeny v session state (cloud i lokál)
    if '_saved_tokens' in st.session_state:
        return st.session_state['_saved_tokens']
    # 2) GitHub Gist (vždy aktuální refresh token)
    gist_tokens = _load_gist_tokens()
    if gist_tokens.get("refresh_token"):
        return gist_tokens
    # 3) Lokální soubor
    try:
        with open(TOKEN_FILE) as f:
            return json.load(f)
    except Exception:
        pass
    # 4) Streamlit Secrets (počáteční nastavení / záloha)
    try:
        tok = st.secrets.get("nafta_token", {})
        if tok.get("refresh_token"):
            return dict(tok)
    except Exception:
        pass
    return {}

def _save_tokens(data: dict):
    # Vždy uložit do session state (funguje všude)
    st.session_state['_saved_tokens'] = data
    # Uložit do GitHub Gist (trvalá persistence na cloudu)
    _save_gist_tokens(data)
    # Pokusit se uložit do souboru (funguje lokálně)
    try:
        with open(TOKEN_FILE, "w") as f:
            json.dump(data, f)
    except Exception:
        pass

def _refresh(refresh_token: str) -> str | None:
    r = requests.post(f"{AUTHORITY}/token", data={
        "client_id": CLIENT_ID, "grant_type": "refresh_token",
        "refresh_token": refresh_token, "scope": SCOPES,
    })
    d = r.json()
    if "access_token" in d:
        _save_tokens({"access_token": d["access_token"],
                      "refresh_token": d.get("refresh_token", refresh_token)})
        return d["access_token"]
    return None

def get_valid_token() -> str | None:
    tokens = _load_tokens()
    if tokens.get("refresh_token"):
        return _refresh(tokens["refresh_token"])
    return None

def start_device_flow() -> dict:
    r = requests.post(f"{AUTHORITY}/devicecode", data={"client_id": CLIENT_ID, "scope": SCOPES})
    return r.json()

def poll_device_token(device_code: str) -> str | None:
    r = requests.post(f"{AUTHORITY}/token", data={
        "client_id": CLIENT_ID,
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "device_code": device_code,
    })
    d = r.json()
    if "access_token" in d:
        _save_tokens({"access_token": d["access_token"], "refresh_token": d.get("refresh_token")})
        return d["access_token"]
    return None

# ── Microsoft Graph ───────────────────────────────────────────────────────────

def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Prefer": 'outlook.body-content-type="text"',
    }

def get_calendar_id(token: str) -> str | None:
    r = requests.get("https://graph.microsoft.com/v1.0/me/calendars", headers=_headers(token))
    for cal in r.json().get("value", []):
        if cal["name"] == "Palivo":
            return cal["id"]
    return None

def fetch_events(token: str, cal_id: str | None, date_from: date, date_to: date) -> list:
    base = (f"https://graph.microsoft.com/v1.0/me/calendars/{cal_id}/events"
            if cal_id else "https://graph.microsoft.com/v1.0/me/events")
    params = {
        "$filter": (f"start/dateTime ge '{date_from}T00:00:00' "
                    f"and start/dateTime le '{date_to}T23:59:59'"),
        "$select": "id,subject,body,start",
        "$orderby": "start/dateTime",
        "$top": 999,
    }
    events, url = [], base
    while url:
        r = requests.get(url, headers=_headers(token), params=params)
        data = r.json()
        events.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None
    return events

def mark_as_paid(token: str, event_id: str, body_raw: str) -> bool:
    new_body = body_raw.rstrip() + "\nZaplaceno: Ano"
    r = requests.patch(
        f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
        headers={**_headers(token), "Content-Type": "application/json"},
        json={"body": {"contentType": "text", "content": new_body}},
    )
    return r.status_code in (200, 204)

def get_tank_info(token: str) -> tuple:
    """Vrátí (stav_nadrze, ceny_df, df_tankování_vse) ze všech historických záznamů."""
    cal_id = get_calendar_id(token)
    events = fetch_events(token, cal_id, date(2020, 1, 1), date.today())
    df = parse_events(events)
    empty_ceny = pd.DataFrame(columns=["datum", "cena_za_litr"])
    empty_tank = pd.DataFrame()
    if df.empty:
        return 0.0, empty_ceny, empty_tank
    df_d = df[df["typ"] == "Doplnění"]
    df_t = df[df["typ"] == "Tankování"]
    doplneno = df_d["litry"].sum()
    odebrano = df_t["litry"].sum()
    level = max(0.0, doplneno - odebrano)
    ceny = df_d[["datum", "cena_za_litr"]].dropna(subset=["cena_za_litr"]).sort_values("datum").reset_index(drop=True)
    return level, ceny, df_t.reset_index(drop=True)

def _cena_pro_datum(ceny_df: pd.DataFrame, datum) -> float | None:
    """Cena za litr z posledního Doplnění na nebo před daným datem."""
    rel = ceny_df[ceny_df["datum"] <= datum]
    if rel.empty:
        return None
    return float(rel.iloc[-1]["cena_za_litr"])

# ── Parsování záznamů ─────────────────────────────────────────────────────────

def _clean(text: str) -> str:
    text = re.sub(r"<br\s*/?>|</p>|</div>|</li>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in text.split("\n")]
    return "\n".join(ln for ln in lines if ln)

def _find(pattern: str, text: str, cast=str):
    m = re.search(pattern, text)
    if m:
        try:
            return cast(m.group(1).strip())
        except Exception:
            return None
    return None

def parse_events(events: list) -> pd.DataFrame:
    rows = []
    for ev in events:
        event_id   = ev.get("id", "")
        subj       = ev.get("subject", "")
        body_raw   = ev.get("body", {}).get("content", "")
        body       = _clean(body_raw)
        dt         = datetime.fromisoformat(ev["start"]["dateTime"].rstrip("Z")).date()
        zaplaceno  = bool(re.search(r"Zaplaceno:\s*Ano", body))

        if subj.startswith("Doplnění nafty:"):
            rows.append({
                "event_id": event_id, "body_raw": body_raw,
                "datum": dt, "typ": "Doplnění",
                "uzivatel": None, "spz": None, "kategorie": None, "platba": None,
                "litry":        _find(r"Litry:\s*([\d.,]+)\s*L",         body, float),
                "cena_za_litr": _find(r"Cena za litr:\s*([\d.,]+)\s*Kč", body, float),
                "celkem_kc":    _find(r"Celkem:\s*([\d.,]+)\s*Kč",       body, float),
                "tachometr": None, "zaplaceno": None,
            })

        elif subj.startswith(("Tankování nafty:", "Odběr nafty:")):
            platba = _find(r"Platba:\s*([^\n]+)", body)
            rows.append({
                "event_id": event_id, "body_raw": body_raw,
                "datum": dt, "typ": "Tankování",
                "uzivatel":  _find(r"Uživatel:\s*([^\n]+)",  body),
                "spz":       _find(r"SPZ:\s*([^\n]+)",       body),
                "kategorie": _find(r"Kategorie:\s*([^\n]+)", body),
                "platba":    platba,
                "litry":     _find(r"Litry:\s*([\d.,]+)\s*L", body, float),
                "cena_za_litr": None, "celkem_kc": None,
                "tachometr": _find(r"Stav tachometru:\s*(\d+)\s*km", body, int),
                "zaplaceno": zaplaceno,
            })

    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df["litry"] = df["litry"].apply(
        lambda x: float(str(x).replace(",", ".")) if x is not None else None)
    return df

# ── Analýza spotřeby ─────────────────────────────────────────────────────────

def calculate_consumption(df_tank: pd.DataFrame) -> pd.DataFrame:
    df_t = df_tank[df_tank["tachometr"].notna()].copy()
    df_t = df_t.sort_values(["spz", "datum"]).reset_index(drop=True)
    results = []
    for spz, grp in df_t.groupby("spz"):
        grp = grp.reset_index(drop=True)
        for i in range(len(grp) - 1):
            tach_od = grp.loc[i,   "tachometr"]
            tach_do = grp.loc[i+1, "tachometr"]
            if tach_do <= tach_od:
                continue
            vzdalenost = tach_do - tach_od
            d_od = grp.loc[i,   "datum"]
            d_do = grp.loc[i+1, "datum"]
            if d_od == d_do:
                litry = float(grp.loc[i + 1, "litry"] or 0)
            else:
                litry = df_tank[
                    (df_tank["spz"] == spz) &
                    (df_tank["datum"] > d_od) &
                    (df_tank["datum"] <= d_do)
                ]["litry"].sum()
            if litry > 0:
                results.append({
                    "SPZ": spz,
                    "Od": d_od, "Do": d_do,
                    "Vzdálenost (km)": int(vzdalenost),
                    "Litry": round(float(litry), 1),
                    "Spotřeba (L/100km)": round(float(litry) / vzdalenost * 100, 1),
                })
    return pd.DataFrame(results) if results else pd.DataFrame()

# ── Excel export ──────────────────────────────────────────────────────────────

def to_excel(df: pd.DataFrame, df_tank_vse: pd.DataFrame | None = None) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_t = df[df["typ"] == "Tankování"].copy()
        df_d = df[df["typ"] == "Doplnění"].copy()

        df_t[["datum","uzivatel","spz","kategorie","platba","litry","tachometr","zaplaceno"]].rename(columns={
            "datum":"Datum","uzivatel":"Uživatel","spz":"SPZ","kategorie":"Kategorie",
            "platba":"Platba","litry":"Litry (L)","tachometr":"Tachometr (km)","zaplaceno":"Zaplaceno",
        }).to_excel(writer, sheet_name="Tankování", index=False)

        df_d[["datum","litry","cena_za_litr","celkem_kc"]].rename(columns={
            "datum":"Datum","litry":"Litry (L)","cena_za_litr":"Cena/L (Kč)","celkem_kc":"Celkem (Kč)",
        }).to_excel(writer, sheet_name="Doplnění", index=False)

        if not df_t.empty:
            df_t.groupby("uzivatel")["litry"].agg(Tankování="count", Litry="sum").reset_index().rename(
                columns={"uzivatel":"Uživatel","Litry":"Litry (L)"}
            ).to_excel(writer, sheet_name="Podle uživatelů", index=False)

            df_t.groupby("spz")["litry"].agg(Tankování="count", Litry="sum").reset_index().rename(
                columns={"spz":"SPZ","Litry":"Litry (L)"}
            ).to_excel(writer, sheet_name="Podle SPZ", index=False)

            df_t.groupby("kategorie")["litry"].agg(Tankování="count", Litry="sum").reset_index().rename(
                columns={"kategorie":"Kategorie","Litry":"Litry (L)"}
            ).to_excel(writer, sheet_name="Podle kategorií", index=False)

            # Dluhy
            dluhy = df_t[(df_t["platba"] == "Na dluh") & (~df_t["zaplaceno"])][
                ["datum","uzivatel","spz","litry"]
            ].rename(columns={"datum":"Datum","uzivatel":"Uživatel","spz":"SPZ","litry":"Litry (L)"})
            dluhy.to_excel(writer, sheet_name="Nesplacené dluhy", index=False)

            spotreba = calculate_consumption(df_tank_vse if df_tank_vse is not None else df_t)
            if not spotreba.empty:
                spotreba.to_excel(writer, sheet_name="Spotřeba", index=False)

    return buf.getvalue()

# ── Vizualizace nádrže ────────────────────────────────────────────────────────

def tank_gauge(level: float) -> go.Figure:
    pct = level / KAPACITA_NADRZE * 100
    color = "red" if pct < 20 else ("orange" if pct < 40 else "steelblue")
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=level,
        number={"suffix": " L", "valueformat": ".0f", "font": {"size": 36}},
        title={"text": f"Stav nádrže  ({pct:.0f} %)", "font": {"size": 16}},
        gauge={
            "axis": {"range": [0, KAPACITA_NADRZE], "ticksuffix": " L"},
            "bar":  {"color": color},
            "steps": [
                {"range": [0,               KAPACITA_NADRZE * 0.20], "color": "#ffe0e0"},
                {"range": [KAPACITA_NADRZE * 0.20, KAPACITA_NADRZE * 0.40], "color": "#fff3cd"},
                {"range": [KAPACITA_NADRZE * 0.40, KAPACITA_NADRZE],        "color": "#e0f0ff"},
            ],
            "threshold": {"line": {"color": "black", "width": 3},
                          "thickness": 0.75, "value": level},
        },
    ))
    fig.update_layout(height=300, margin=dict(t=60, b=40, l=30, r=30))
    return fig

# ── Streamlit UI ──────────────────────────────────────────────────────────────

st.set_page_config(page_title="Nafta – přehled", layout="wide", page_icon="⛽")

# --- PIN ochrana ---
if not st.session_state.get("pin_ok"):
    st.markdown("""
        <style>
        .pin-box { max-width: 320px; margin: 15vh auto 0 auto; text-align: center; }
        .pin-box input { text-align: center; letter-spacing: 4px; font-size: 1.4rem; }
        </style>
        <div class="pin-box"><h2>⛽ Nafta</h2></div>
    """, unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1, 1])
    with col:
        pin = st.text_input("PIN", type="password", label_visibility="collapsed",
                            placeholder="")
        if st.button("Vstoupit", type="primary", use_container_width=True):
            spravny_pin = st.secrets.get("app", {}).get("pin", "")
            if pin == spravny_pin:
                st.session_state.pin_ok = True
                st.rerun()
            else:
                st.error("Nesprávný PIN")
    st.stop()

# --- Přihlášení ---
if "token" not in st.session_state:
    token = get_valid_token()
    if token:
        st.session_state.token = token
    else:
        st.title("⛽ Nafta – přihlášení")
        if "flow" not in st.session_state:
            with st.spinner("Připravuji přihlášení..."):
                st.session_state.flow = start_device_flow()
        flow = st.session_state.flow
        st.info(f"Přejdi na **{flow.get('verification_uri','microsoft.com/devicelogin')}** a zadej kód:")
        st.code(flow.get("user_code", ""), language=None)
        st.caption("Po přihlášení klikni na tlačítko níže.")
        if st.button("✅ Přihlásil jsem se", type="primary"):
            tok = poll_device_token(flow["device_code"])
            if tok:
                st.session_state.token = tok
                del st.session_state.flow
                st.rerun()
            else:
                st.error("Token ještě není k dispozici – zkus za chvíli.")
        st.stop()

token = st.session_state.token

# --- Stav nádrže + historické ceny (načítáme jednou za sezení) ---
if "tank_level" not in st.session_state:
    with st.spinner("Načítám stav nádrže..."):
        st.session_state.tank_level, st.session_state.ceny_df, st.session_state.df_tank_vse = get_tank_info(token)

# --- Hlavní stránka ---
st.title("⛽ Nafta – přehled")

# Stav nádrže nahoře
col_gauge, col_metrics = st.columns([1, 2])
with col_gauge:
    st.plotly_chart(tank_gauge(st.session_state.tank_level),
                    use_container_width=True, config={"displayModeBar": False})

# --- Datum + Načíst ---
col1, col2 = st.columns(2)
with col1:
    _d = date.today()
    _m, _y = _d.month - 2, _d.year
    if _m <= 0: _m += 12; _y -= 1
    date_from = st.date_input("Od", value=date(_y, _m, 1))
with col2:
    date_to = st.date_input("Do", value=date.today())
aktualizovat = st.button("🔄 Aktualizovat", type="primary")

def _nacti(token, date_from, date_to):
    cal_id = get_calendar_id(token)
    events = fetch_events(token, cal_id, date_from, date_to)
    st.session_state.df = parse_events(events)
    st.session_state.nacteno_pro = (date_from, date_to)
    st.session_state.tank_level, st.session_state.ceny_df, st.session_state.df_tank_vse = get_tank_info(token)

potreba_nacist = (
    "df" not in st.session_state
    or aktualizovat
    or st.session_state.get("nacteno_pro") != (date_from, date_to)
)
if potreba_nacist:
    with st.spinner("Načítám záznamy z kalendáře Palivo..."):
        _nacti(token, date_from, date_to)

df: pd.DataFrame = st.session_state.df
df_t = df[df["typ"] == "Tankování"] if not df.empty else pd.DataFrame()
df_d = df[df["typ"] == "Doplnění"]  if not df.empty else pd.DataFrame()

# Metriky vedle gauge
with col_metrics:
    st.write("")
    st.write("")
    m1, m2 = st.columns(2)
    m1.metric("Tankování (období)", f"{len(df_t)}")
    m2.metric("Odebráno (období)",  f"{df_t['litry'].sum():.1f} L" if not df_t.empty else "0 L")
    m3, m4 = st.columns(2)
    m3.metric("Doplnění (období)",  f"{len(df_d)}")
    m4.metric("Zaplaceno celkem",   f"{df_d['celkem_kc'].sum():.0f} Kč" if not df_d.empty else "0 Kč")

if df.empty:
    st.warning("V daném období nejsou žádné záznamy.")
    st.stop()

# --- Excel export ---
st.download_button("⬇ Stáhnout Excel", data=to_excel(df, st.session_state.get("df_tank_vse")),
    file_name=f"nafta_{date_from}_{date_to}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()

# ── DLUHY ────────────────────────────────────────────────────────────────────
dluhy = df_t[(df_t["platba"] == "Na dluh") & (~df_t["zaplaceno"].fillna(False))].copy()
ceny_df = st.session_state.get("ceny_df", pd.DataFrame(columns=["datum", "cena_za_litr"]))

if not dluhy.empty:
    # Přidat cenu a částku z posledního Doplnění před datem tankování
    dluhy["cena_za_litr"] = dluhy["datum"].apply(lambda d: _cena_pro_datum(ceny_df, d))
    dluhy["castka_kc"] = dluhy.apply(
        lambda r: round(r["litry"] * r["cena_za_litr"], 0)
        if (r["litry"] and r["cena_za_litr"]) else None, axis=1
    )

    st.subheader("💳 Nesplacené dluhy")

    # Souhrn dluhů per uživatel
    dluh_sum = dluhy.groupby("uzivatel").agg(
        Tankování=("litry", "count"),
        **{"Litry (L)": ("litry", "sum")},
        **{"Částka (Kč)": ("castka_kc", "sum")},
    ).reset_index().rename(columns={"uzivatel": "Uživatel"})
    dluh_sum["Litry (L)"] = dluh_sum["Litry (L)"].round(1)
    dluh_sum["Částka (Kč)"] = dluh_sum["Částka (Kč)"].apply(
        lambda x: f"{x:.0f} Kč" if pd.notna(x) else "–"
    )
    st.dataframe(dluh_sum, use_container_width=False, hide_index=True)

    st.write("**Označit jako zaplaceno:**")
    for _, row in dluhy.iterrows():
        c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 1, 1, 1, 1])
        c1.write(str(row["datum"]))
        c2.write(str(row["uzivatel"] or ""))
        c3.write(str(row["spz"] or ""))
        c4.write(f"{row['litry']:.1f} L" if row["litry"] else "")
        castka = row.get("castka_kc")
        c5.write(f"{castka:.0f} Kč" if castka else "–")
        if c6.button("✅ Zaplatil", key=f"paid_{row['event_id']}"):
            with st.spinner("Ukládám..."):
                ok = mark_as_paid(token, row["event_id"], row["body_raw"])
            if ok:
                del st.session_state["df"]
                del st.session_state["nacteno_pro"]
                st.session_state.tank_level, st.session_state.ceny_df, st.session_state.df_tank_vse = get_tank_info(token)
                st.rerun()
            else:
                st.error("Nepodařilo se uložit.")
    st.divider()

# ── TABY ─────────────────────────────────────────────────────────────────────
tab_t, tab_u, tab_spz, tab_kat, tab_d = st.tabs(
    ["🚗 Tankování", "👤 Uživatelé", "🔧 Vozidla & Spotřeba", "🏷 Kategorie", "🛢 Doplnění"])

with tab_t:
    cols = ["datum","uzivatel","spz","kategorie","platba","litry","tachometr","zaplaceno"]
    rename = {"datum":"Datum","uzivatel":"Uživatel","spz":"SPZ","kategorie":"Kategorie",
              "platba":"Platba","litry":"Litry (L)","tachometr":"Tachometr (km)","zaplaceno":"Zaplaceno"}
    st.dataframe(df_t[cols].rename(columns=rename), use_container_width=True, hide_index=True)

with tab_u:
    if not df_t.empty:
        grp = df_t.groupby("uzivatel")["litry"].agg(Tankování="count", Litry="sum").reset_index()
        grp.columns = ["Uživatel","Tankování","Litry (L)"]
        grp["Litry (L)"] = grp["Litry (L)"].round(1)
        grp = grp.sort_values("Litry (L)", ascending=False)
        c1, c2 = st.columns([2, 1])
        with c1:
            st.bar_chart(grp.set_index("Uživatel")["Litry (L)"])
        with c2:
            st.dataframe(grp, use_container_width=True, hide_index=True)

with tab_spz:
    if not df_t.empty:
        grp = df_t.groupby("spz")["litry"].agg(Tankování="count", Litry="sum").reset_index()
        grp.columns = ["SPZ","Tankování","Litry (L)"]
        grp["Litry (L)"] = grp["Litry (L)"].round(1)
        grp = grp.sort_values("Litry (L)", ascending=False)
        c1, c2 = st.columns([2, 1])
        with c1:
            st.bar_chart(grp.set_index("SPZ")["Litry (L)"])
        with c2:
            st.dataframe(grp, use_container_width=True, hide_index=True)

        df_tank_vse = st.session_state.get("df_tank_vse", df_t)
        spotreba = calculate_consumption(df_tank_vse)
        if not spotreba.empty:
            st.subheader("⛽ Průměrná spotřeba")
            # Průměr na SPZ
            avg = spotreba.groupby("SPZ")["Spotřeba (L/100km)"].mean().reset_index()
            avg["Spotřeba (L/100km)"] = avg["Spotřeba (L/100km)"].round(1)
            c1, c2 = st.columns([1, 1])
            with c1:
                st.bar_chart(avg.set_index("SPZ")["Spotřeba (L/100km)"])
            with c2:
                st.dataframe(spotreba, use_container_width=True, hide_index=True)
        else:
            st.info("Pro výpočet spotřeby jsou potřeba alespoň dvě tankování se stavem tachometru u stejného vozidla.")

with tab_kat:
    if not df_t.empty:
        grp = df_t.groupby("kategorie")["litry"].agg(Tankování="count", Litry="sum").reset_index()
        grp.columns = ["Kategorie","Tankování","Litry (L)"]
        grp["Litry (L)"] = grp["Litry (L)"].round(1)
        grp = grp.sort_values("Litry (L)", ascending=False)
        c1, c2 = st.columns([2, 1])
        with c1:
            st.bar_chart(grp.set_index("Kategorie")["Litry (L)"])
        with c2:
            st.dataframe(grp, use_container_width=True, hide_index=True)

with tab_d:
    if not df_d.empty:
        st.dataframe(
            df_d[["datum","litry","cena_za_litr","celkem_kc"]].rename(columns={
                "datum":"Datum","litry":"Litry (L)",
                "cena_za_litr":"Cena/L (Kč)","celkem_kc":"Celkem (Kč)"}),
            use_container_width=True, hide_index=True)
        c1, c2, c3 = st.columns(3)
        c1.metric("Celkem doplněno", f"{df_d['litry'].sum():.1f} L")
        c2.metric("Průměrná cena/L", f"{df_d['cena_za_litr'].mean():.2f} Kč")
        c3.metric("Celkem zaplaceno", f"{df_d['celkem_kc'].sum():.0f} Kč")
