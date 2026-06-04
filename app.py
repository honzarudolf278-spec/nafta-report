import streamlit as st
import pandas as pd
import requests
import json
import re
import io
import time
from datetime import date, datetime
import plotly.graph_objects as go

# ── Konfigurace ───────────────────────────────────────────────────────────────
CLIENT_ID  = "9c5552a6-4492-4e81-b96b-d1ea74abb8ed"
AUTHORITY  = "https://login.microsoftonline.com/consumers/oauth2/v2.0"
SCOPES     = "Calendars.ReadWrite Contacts.ReadWrite offline_access"
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

def _patch_event(token: str, event_id: str, subject: str, body: str) -> bool:
    payload = {"subject": subject, "body": {"contentType": "text", "content": body}}
    r = requests.patch(f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
                       headers={**_headers(token), "Content-Type": "application/json"}, json=payload)
    if r.status_code == 401:
        t2 = get_valid_token()
        if t2:
            r = requests.patch(f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
                               headers={**_headers(t2), "Content-Type": "application/json"}, json=payload)
    return r.status_code in (200, 204)

def upravit_tankovani(token: str, event_id: str, uzivatel: str, spz: str, kategorie: str,
                      platba: str, litry: float, tachometr, body_raw: str) -> bool:
    zaplaceno_radek = "\nZaplaceno: Ano" if re.search(r"Zaplaceno:\s*Ano", body_raw) else ""
    popis = (f"Uživatel: {uzivatel}\nSPZ: {spz}\nKategorie: {kategorie}\n"
             f"Platba: {platba}\nLitry: {litry} L")
    if tachometr: popis += f"\nStav tachometru: {int(tachometr)} km"
    popis += zaplaceno_radek
    return _patch_event(token, event_id, f"Tankování nafty: {spz} - {litry:.1f} L", popis)

def upravit_doplneni(token: str, event_id: str, litry: float, cena: float) -> bool:
    celkem = litry * cena
    popis = f"Litry: {litry} L\nCena za litr: {cena} Kč\nCelkem: {celkem:.2f} Kč"
    return _patch_event(token, event_id, f"Doplnění nafty: {litry:.1f} L", popis)

def mark_as_paid(token: str, event_id: str, body_raw: str) -> bool:
    # Nepřidávej duplicitně
    if re.search(r"Zaplaceno:\s*Ano", body_raw):
        return True
    new_body = body_raw.rstrip() + "\nZaplaceno: Ano"
    r = requests.patch(
        f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
        headers={**_headers(token), "Content-Type": "application/json"},
        json={"body": {"contentType": "text", "content": new_body}},
    )
    if r.status_code == 401:
        token2 = get_valid_token()
        if token2:
            r = requests.patch(
                f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
                headers={**_headers(token2), "Content-Type": "application/json"},
                json={"body": {"contentType": "text", "content": new_body}},
            )
    return r.status_code in (200, 204)

# ── Správa kontaktů ───────────────────────────────────────────────────────────

def _get_or_create_contact_folder(token: str, name: str) -> str | None:
    """Vrátí ID složky kontaktů dle názvu, nebo ji vytvoří."""
    r = requests.get("https://graph.microsoft.com/v1.0/me/contactFolders",
                     headers=_headers(token))
    if r.status_code == 200:
        for f in r.json().get("value", []):
            if f["displayName"].lower() == name.lower():
                return f["id"]
    # Vytvoř složku
    r2 = requests.post("https://graph.microsoft.com/v1.0/me/contactFolders",
                       headers={**_headers(token), "Content-Type": "application/json"},
                       json={"displayName": name})
    if r2.status_code == 201:
        return r2.json()["id"]
    return None

def nacist_zamestnance_web(token: str) -> list[dict]:
    """Načte zaměstnance ze složky Zaměstnanci (pro webové rozhraní)."""
    r = requests.get("https://graph.microsoft.com/v1.0/me/contactFolders",
                     headers=_headers(token))
    folder_id = None
    if r.status_code == 200:
        for f in r.json().get("value", []):
            if f["displayName"].lower() == "zaměstnanci":
                folder_id = f["id"]
                break
    if not folder_id:
        return []
    url = f"https://graph.microsoft.com/v1.0/me/contactFolders/{folder_id}/contacts?$select=id,displayName,givenName,surname,title,personalNotes"
    r2 = requests.get(url, headers=_headers(token))
    if r2.status_code != 200:
        return []
    result = []
    for c in r2.json().get("value", []):
        notes = c.get("personalNotes", "")
        pin_m   = re.search(r"Heslo:\s*(\S+)",                          notes)
        limit_m = re.search(r"proplácen[eé]:\s*(\d+)", notes, re.IGNORECASE)
        spz_m   = re.search(r"SPZ:\s*([A-Z0-9 ]+)",    notes, re.IGNORECASE)
        result.append({
            "id":       c["id"],
            "jmeno":    c.get("displayName", ""),
            "givenName":c.get("givenName", ""),
            "surname":  c.get("surname", ""),
            "title":    c.get("title", ""),
            "pin":      pin_m.group(1)            if pin_m   else "",
            "limit":    limit_m.group(1)          if limit_m else "0",
            "spz":      spz_m.group(1).strip().upper() if spz_m else "",
        })
    return result

def pridat_zamestnance(token: str, titul: str, jmeno: str, prijmeni: str,
                       pin: str, limit: int, spz: str) -> bool:
    folder_id = _get_or_create_contact_folder(token, "Zaměstnanci")
    if not folder_id:
        return False
    notes = f"Heslo: {pin}"
    if limit > 0:
        notes += f"\nproplácení: {limit}"
    if spz:
        notes += f"\nSPZ: {spz.upper()}"
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/me/contactFolders/{folder_id}/contacts",
        headers={**_headers(token), "Content-Type": "application/json"},
        json={"title": titul, "givenName": jmeno, "surname": prijmeni,
              "personalNotes": notes},
    )
    return r.status_code == 201

def upravit_zamestnance(token: str, contact_id: str, titul: str, jmeno: str,
                        prijmeni: str, pin: str, limit: int, spz: str) -> bool:
    notes = f"Heslo: {pin}"
    if limit > 0:
        notes += f"\nproplácení: {limit}"
    if spz:
        notes += f"\nSPZ: {spz.upper()}"
    r = requests.patch(
        f"https://graph.microsoft.com/v1.0/me/contacts/{contact_id}",
        headers={**_headers(token), "Content-Type": "application/json"},
        json={"title": titul, "givenName": jmeno, "surname": prijmeni,
              "personalNotes": notes},
    )
    return r.status_code in (200, 204)

def smazat_zamestnance(token: str, contact_id: str) -> bool:
    r = requests.delete(f"https://graph.microsoft.com/v1.0/me/contacts/{contact_id}",
                        headers=_headers(token))
    return r.status_code == 204

def upravit_korekci(token: str, event_id: str, litry: float, poznamka: str = "") -> bool:
    nazev = f"Korekce nádrže: {litry:.0f} L"
    popis = f"Litry: {litry} L"
    if poznamka:
        popis += f"\nPoznámka: {poznamka}"
    r = requests.patch(
        f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
        headers={**_headers(token), "Content-Type": "application/json"},
        json={"subject": nazev, "body": {"contentType": "text", "content": popis}},
    )
    if r.status_code == 401:
        token2 = get_valid_token()
        if token2:
            r = requests.patch(
                f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
                headers={**_headers(token2), "Content-Type": "application/json"},
                json={"subject": nazev, "body": {"contentType": "text", "content": popis}},
            )
    return r.status_code in (200, 204)

def smazat_udalost(token: str, event_id: str) -> bool:
    r = requests.delete(
        f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
        headers=_headers(token))
    if r.status_code == 401:
        token2 = get_valid_token()
        if token2:
            r = requests.delete(
                f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
                headers=_headers(token2))
    return r.status_code == 204

def _get_vozidla_folder_id(token: str) -> str | None:
    """Najde složku Vozidla (top-level nebo child)."""
    r = requests.get("https://graph.microsoft.com/v1.0/me/contactFolders",
                     headers=_headers(token))
    if r.status_code != 200:
        return None
    for f in r.json().get("value", []):
        if f["displayName"].lower() == "vozidla":
            return f["id"]
        # Hledej v podsložkách
        rc = requests.get(
            f"https://graph.microsoft.com/v1.0/me/contactFolders/{f['id']}/childFolders",
            headers=_headers(token))
        if rc.status_code == 200:
            for c in rc.json().get("value", []):
                if c["displayName"].lower() == "vozidla":
                    return c["id"]
    return None

def _get_or_create_vozidla_folder(token: str) -> str | None:
    """Vrátí ID složky Vozidla, nebo ji vytvoří jako podsložku Zaměstnanci."""
    fid = _get_vozidla_folder_id(token)
    if fid:
        return fid
    # Vytvoř pod složkou Zaměstnanci
    zam_id = _get_or_create_contact_folder(token, "Zaměstnanci")
    if not zam_id:
        return None
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/me/contactFolders/{zam_id}/childFolders",
        headers={**_headers(token), "Content-Type": "application/json"},
        json={"displayName": "Vozidla"})
    if r.status_code == 201:
        return r.json()["id"]
    return None

def nacist_vozidla_web(token: str) -> list[dict]:
    fid = _get_vozidla_folder_id(token)
    if not fid:
        return []
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/me/contactFolders/{fid}/contacts?$select=id,givenName,surname",
        headers=_headers(token))
    if r.status_code != 200:
        return []
    result = []
    for c in r.json().get("value", []):
        spz = (c.get("surname") or "").strip().upper()
        typ = (c.get("givenName") or "").strip()
        if spz:
            result.append({"id": c["id"], "typ": typ, "spz": spz})
    return sorted(result, key=lambda x: x["spz"])

def pridat_vozidlo(token: str, typ: str, spz: str) -> bool:
    fid = _get_or_create_vozidla_folder(token)
    if not fid:
        return False
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/me/contactFolders/{fid}/contacts",
        headers={**_headers(token), "Content-Type": "application/json"},
        json={"givenName": typ, "surname": spz.upper()})
    return r.status_code == 201

def upravit_vozidlo(token: str, contact_id: str, typ: str, spz: str) -> bool:
    r = requests.patch(
        f"https://graph.microsoft.com/v1.0/me/contacts/{contact_id}",
        headers={**_headers(token), "Content-Type": "application/json"},
        json={"givenName": typ, "surname": spz.upper()})
    return r.status_code in (200, 204)

def smazat_vozidlo(token: str, contact_id: str) -> bool:
    r = requests.delete(f"https://graph.microsoft.com/v1.0/me/contacts/{contact_id}",
                        headers=_headers(token))
    return r.status_code == 204

# ── Korekce nádrže ────────────────────────────────────────────────────────────

def ulozit_korekci_nadrze(token: str, litry: float, poznamka: str = "") -> bool:
    cal_id = get_calendar_id(token)
    url = (f"https://graph.microsoft.com/v1.0/me/calendars/{cal_id}/events"
           if cal_id else "https://graph.microsoft.com/v1.0/me/events")
    now = datetime.now()
    popis = f"Litry: {litry} L"
    if poznamka:
        popis += f"\nPoznámka: {poznamka}"
    r = requests.post(
        url,
        headers={**_headers(token), "Content-Type": "application/json"},
        json={
            "subject": f"Korekce nádrže: {litry:.0f} L",
            "body": {"contentType": "text", "content": popis},
            "start": {"dateTime": now.isoformat(), "timeZone": "Europe/Prague"},
            "end":   {"dateTime": now.replace(minute=now.minute+1).isoformat(), "timeZone": "Europe/Prague"},
        },
    )
    return r.status_code == 201

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
    df_k = df[df["typ"] == "Korekce"]
    # Pokud existuje korekce, použij poslední jako výchozí bod
    if not df_k.empty:
        last_k = df_k.sort_values("datum").iloc[-1]
        baseline = float(last_k["litry"] or 0.0)
        cutoff   = last_k["datum"]
        df_d = df_d[df_d["datum"] > cutoff]
        df_t = df_t[df_t["datum"] > cutoff]
    else:
        baseline = 0.0
    doplneno = df_d["litry"].sum()
    odebrano = df_t["litry"].sum()
    level = max(0.0, baseline + doplneno - odebrano)
    ceny = df[df["typ"] == "Doplnění"][["datum", "cena_za_litr"]].dropna(subset=["cena_za_litr"]).sort_values("datum").reset_index(drop=True)
    return level, ceny, df[df["typ"] == "Tankování"].reset_index(drop=True)

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

        elif subj.startswith("Korekce nádrže:"):
            litry_subj = _find(r"Korekce nádrže:\s*([\d.,]+)", subj, float)
            rows.append({
                "event_id": event_id, "body_raw": body_raw,
                "datum": dt, "typ": "Korekce",
                "uzivatel": None, "spz": None, "kategorie": None, "platba": None,
                "litry": litry_subj or _find(r"Litry:\s*([\d.,]+)\s*L", body, float),
                "cena_za_litr": None, "celkem_kc": None,
                "tachometr": None, "zaplaceno": None,
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

st.set_page_config(page_title="Nafta", layout="wide", page_icon="⛽")

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
        st.text_input("Uživatel", placeholder="uživatel", autocomplete="username",
                      label_visibility="collapsed", key="_login_user")
        pin = st.text_input("PIN", type="password", label_visibility="collapsed",
                            placeholder="PIN", autocomplete="current-password")
        if st.button("Vstoupit", type="primary", use_container_width=True):
            spravny_pin  = st.secrets.get("app", {}).get("pin", "")
            spravny_user = st.secrets.get("app", {}).get("username", "martver_user")
            zadany_user  = st.session_state.get("_login_user", "")
            if zadany_user == spravny_user and pin == spravny_pin:
                st.session_state.pin_ok = True
                st.rerun()
            else:
                st.error("Nesprávné přihlašovací údaje")
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
st.title("⛽ Nafta")

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
    st.session_state.zaplacene_ids = set()
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
else:
    st.download_button("⬇ Stáhnout Excel", data=to_excel(df, st.session_state.get("df_tank_vse")),
        file_name=f"nafta_{date_from}_{date_to}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()

# ── DLUHY ────────────────────────────────────────────────────────────────────
if "zaplacene_ids" not in st.session_state:
    st.session_state.zaplacene_ids = set()

_paid = st.session_state.zaplacene_ids
dluhy = df_t[
    (df_t["platba"] == "Na dluh") &
    (~df_t["zaplaceno"].fillna(False)) &
    (~df_t["event_id"].isin(_paid))
].copy() if not df_t.empty else pd.DataFrame()
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
                st.session_state.zaplacene_ids.add(row["event_id"])
                st.rerun()
            else:
                st.error("Nepodařilo se uložit.")
    st.divider()

# ── TABY ─────────────────────────────────────────────────────────────────────
tab_t, tab_u, tab_spz, tab_kat, tab_d, tab_kor, tab_admin = st.tabs(
    ["🚗 Tankování", "👤 Uživatelé", "🔧 Vozidla & Spotřeba", "🏷 Kategorie", "🛢 Doplnění", "📋 Korekce", "⚙️ Správa"])

with tab_t:
    if df_t.empty:
        st.info("V daném období nejsou žádná tankování.")
    else:
        cols = ["datum","uzivatel","spz","kategorie","platba","litry","tachometr","zaplaceno"]
        rename = {"datum":"Datum","uzivatel":"Uživatel","spz":"SPZ","kategorie":"Kategorie",
                  "platba":"Platba","litry":"Litry (L)","tachometr":"Tachometr (km)","zaplaceno":"Zaplaceno"}
        _bt1, _bt2 = st.columns([1, 9])
        if _bt1.button("☑ Vše", key="sa_t_on"):
            st.session_state.sa_tank = True; st.rerun()
        if _bt2.button("☐ Nic", key="sa_t_off"):
            st.session_state.sa_tank = False; st.rerun()
        df_t_disp = df_t[cols].rename(columns=rename).copy()
        df_t_disp.insert(0, "☑", st.session_state.get("sa_tank", False))
        edited_t = st.data_editor(df_t_disp, use_container_width=True, hide_index=True,
                                  column_config={"☑": st.column_config.CheckboxColumn("☑", width="small")})
        ids_t = df_t[edited_t["☑"].values]["event_id"].tolist()
        if ids_t:
            if st.button(f"🗑 Smazat vybrané ({len(ids_t)})", type="primary", key="del_tank"):
                with st.spinner("Mažu záznamy..."):
                    ok_count = sum(smazat_udalost(token, eid) for eid in ids_t)
                st.success(f"Smazáno {ok_count} z {len(ids_t)} záznamů")
                st.session_state.sa_tank = False
                del st.session_state["df"], st.session_state["nacteno_pro"]
                st.rerun()

        if len(ids_t) == 1:
            row_t = df_t[df_t["event_id"] == ids_t[0]].iloc[0]
            with st.expander("✏️ Upravit vybraný záznam"):
                ec1, ec2, ec3 = st.columns(3)
                et_uziv  = ec1.text_input("Uživatel", value=str(row_t["uzivatel"] or ""), key="et_uz")
                et_spz   = ec2.text_input("SPZ", value=str(row_t["spz"] or ""), key="et_spz")
                et_litry = ec3.number_input("Litry (L)", value=float(row_t["litry"] or 0), min_value=0.0, step=0.5, key="et_l")
                ec4, ec5, ec6 = st.columns(3)
                kat_options = ['Služební', 'Osobní proplacené', 'Osobní neproplácené', 'Ostatní']
                plat_options = ['Firma', 'Hotově', 'Na dluh', 'QR platba']
                et_kat   = ec4.selectbox("Kategorie", kat_options,
                    index=kat_options.index(row_t["kategorie"]) if row_t["kategorie"] in kat_options else 0, key="et_kat")
                et_plat  = ec5.selectbox("Platba", plat_options,
                    index=plat_options.index(row_t["platba"]) if row_t["platba"] in plat_options else 0, key="et_pl")
                et_tach  = ec6.text_input("Tachometr (km)", value=str(int(row_t["tachometr"])) if pd.notna(row_t["tachometr"]) else "", key="et_tach")
                if st.button("💾 Uložit změny", key="et_save"):
                    with st.spinner("Ukládám..."):
                        ok = upravit_tankovani(token, ids_t[0], et_uziv, et_spz, et_kat, et_plat,
                                               et_litry, et_tach or None, row_t["body_raw"])
                    if ok:
                        st.success("Uloženo")
                        st.session_state.sa_tank = False
                        del st.session_state["df"], st.session_state["nacteno_pro"]
                        st.rerun()
                    else:
                        st.error("Chyba při ukládání")

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
    else:
        st.info("V daném období nejsou žádná tankování.")

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
            avg = spotreba.groupby("SPZ")["Spotřeba (L/100km)"].mean().reset_index()
            avg["Spotřeba (L/100km)"] = avg["Spotřeba (L/100km)"].round(1)
            c1, c2 = st.columns([1, 1])
            with c1:
                st.bar_chart(avg.set_index("SPZ")["Spotřeba (L/100km)"])
            with c2:
                st.dataframe(spotreba, use_container_width=True, hide_index=True)
        else:
            st.info("Pro výpočet spotřeby jsou potřeba alespoň dvě tankování se stavem tachometru u stejného vozidla.")
    else:
        st.info("V daném období nejsou žádná tankování.")

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
    else:
        st.info("V daném období nejsou žádné kategorie.")

with tab_d:
    if not df_d.empty:
        _bd1, _bd2 = st.columns([1, 9])
        if _bd1.button("☑ Vše", key="sa_d_on"):
            st.session_state.sa_dopl = True; st.rerun()
        if _bd2.button("☐ Nic", key="sa_d_off"):
            st.session_state.sa_dopl = False; st.rerun()
        df_d_disp = df_d[["datum","litry","cena_za_litr","celkem_kc"]].rename(columns={
            "datum":"Datum","litry":"Litry (L)",
            "cena_za_litr":"Cena/L (Kč)","celkem_kc":"Celkem (Kč)"}).copy()
        df_d_disp.insert(0, "☑", st.session_state.get("sa_dopl", False))
        edited_d = st.data_editor(df_d_disp, use_container_width=True, hide_index=True,
                                  column_config={"☑": st.column_config.CheckboxColumn("☑", width="small")})
        ids_d = df_d[edited_d["☑"].values]["event_id"].tolist()
        if ids_d:
            if st.button(f"🗑 Smazat vybrané ({len(ids_d)})", type="primary", key="del_dopl"):
                with st.spinner("Mažu záznamy..."):
                    ok_count = sum(smazat_udalost(token, eid) for eid in ids_d)
                st.success(f"Smazáno {ok_count} z {len(ids_d)} záznamů")
                st.session_state.sa_dopl = False
                del st.session_state["df"], st.session_state["nacteno_pro"]
                st.rerun()

        if len(ids_d) == 1:
            row_d = df_d[df_d["event_id"] == ids_d[0]].iloc[0]
            with st.expander("✏️ Upravit vybraný záznam"):
                dc1, dc2 = st.columns(2)
                ed_litry = dc1.number_input("Litry (L)", value=float(row_d["litry"] or 0), min_value=0.0, step=1.0, key="ed_l")
                ed_cena  = dc2.number_input("Cena za litr (Kč)", value=float(row_d["cena_za_litr"] or 0), min_value=0.0, step=0.1, format="%.2f", key="ed_c")
                if st.button("💾 Uložit změny", key="ed_save"):
                    with st.spinner("Ukládám..."):
                        ok = upravit_doplneni(token, ids_d[0], ed_litry, ed_cena)
                    if ok:
                        st.success("Uloženo")
                        st.session_state.sa_dopl = False
                        del st.session_state["df"], st.session_state["nacteno_pro"]
                        st.session_state.tank_level, st.session_state.ceny_df, st.session_state.df_tank_vse = get_tank_info(token)
                        st.rerun()
                    else:
                        st.error("Chyba při ukládání")

        c1, c2, c3 = st.columns(3)
        c1.metric("Celkem doplněno", f"{df_d['litry'].sum():.1f} L")
        c2.metric("Průměrná cena/L", f"{df_d['cena_za_litr'].mean():.2f} Kč")
        c3.metric("Celkem zaplaceno", f"{df_d['celkem_kc'].sum():.0f} Kč")
    else:
        st.info("V daném období nejsou žádná doplnění.")

with tab_kor:
    # Načteme korekce vždy z celé historie
    if "df_kor_cache" not in st.session_state:
        with st.spinner("Načítám korekce..."):
            _cal = get_calendar_id(token)
            _all_ev = fetch_events(token, _cal, date(2020, 1, 1), date.today())
            _df_all = parse_events(_all_ev)
            st.session_state.df_kor_cache = (
                _df_all[_df_all["typ"] == "Korekce"].sort_values("datum", ascending=False).reset_index(drop=True)
                if not _df_all.empty else pd.DataFrame()
            )
    df_kor = st.session_state.df_kor_cache

    if df_kor.empty:
        st.info("Zatím žádné korekce.")
    else:
        # Tlačítka Označit vše / Nic
        _bk1, _bk2 = st.columns([1, 9])
        if _bk1.button("☑ Vše", key="sa_k_on"):
            st.session_state.sa_kor = True; st.rerun()
        if _bk2.button("☐ Nic", key="sa_k_off"):
            st.session_state.sa_kor = False; st.rerun()

        # Tabulka se zaškrtávacím sloupcem
        df_kor_disp = df_kor[["datum","litry"]].rename(columns={"datum":"Datum","litry":"Litry (L)"}).copy()
        df_kor_disp["Poznámka"] = df_kor["body_raw"].apply(
            lambda b: _find(r"Poznámka:\s*([^\n]+)", _clean(b)) or "")
        df_kor_disp.insert(0, "☑", st.session_state.get("sa_kor", False))
        edited_k = st.data_editor(df_kor_disp, use_container_width=True, hide_index=True,
                                  column_config={"☑": st.column_config.CheckboxColumn("☑", width="small")})
        ids_k = df_kor[edited_k["☑"].values]["event_id"].tolist()
        if ids_k:
            if st.button(f"🗑 Smazat vybrané ({len(ids_k)})", type="primary", key="del_kor"):
                with st.spinner("Mažu..."):
                    ok_count = sum(smazat_udalost(token, eid) for eid in ids_k)
                st.success(f"Smazáno {ok_count} z {len(ids_k)}")
                st.session_state.sa_kor = False
                del st.session_state["df_kor_cache"], st.session_state["nacteno_pro"]
                st.session_state.tank_level, st.session_state.ceny_df, st.session_state.df_tank_vse = get_tank_info(token)
                st.rerun()

        st.divider()
        st.markdown("**Upravit korekci:**")
        for _, row in df_kor.iterrows():
            poz = _find(r"Poznámka:\s*([^\n]+)", _clean(row["body_raw"])) or ""
            with st.expander(f"{row['datum']} — {row['litry']:.0f} L" + (f" — {poz}" if poz else "")):
                kc1, kc2, kc3 = st.columns([2, 3, 1])
                new_l = kc1.number_input("Litry", 0.0, float(KAPACITA_NADRZE),
                                         float(row["litry"] or 0), 50.0, "%.0f",
                                         key=f"kor_l_{row['event_id']}")
                new_p = kc2.text_input("Poznámka", value=poz, key=f"kor_p_{row['event_id']}")
                kc3.write(""); kc3.write("")
                if kc3.button("💾", key=f"kor_s_{row['event_id']}"):
                    if upravit_korekci(token, row["event_id"], new_l, new_p):
                        del st.session_state["df_kor_cache"]
                        st.session_state.tank_level, st.session_state.ceny_df, st.session_state.df_tank_vse = get_tank_info(token)
                        st.rerun()
                    else:
                        st.error("Chyba")

with tab_admin:
    st.subheader("⚙️ Správa")

    # ── Korekce nádrže ────────────────────────────────────────────────────────
    st.markdown("### 🛢 Korekce stavu nádrže")
    st.caption(f"Kapacita nádrže: {KAPACITA_NADRZE} L &nbsp;|&nbsp; Aktuální stav: {st.session_state.tank_level:.0f} L")

    col_k1, col_k2 = st.columns([2, 1])
    with col_k1:
        korekce_litry = st.number_input(
            "Nastavit skutečný stav nádrže na (litrů):",
            min_value=0.0, max_value=float(KAPACITA_NADRZE),
            value=0.0, step=50.0, format="%.0f"
        )
        korekce_poznamka = st.text_input("Poznámka (volitelné)", placeholder="např. Fyzická kontrola nádrže")
    with col_k2:
        st.write("")
        st.write("")
        if st.button("💾 Uložit korekci", type="primary"):
            with st.spinner("Ukládám..."):
                ok = ulozit_korekci_nadrze(token, korekce_litry, korekce_poznamka)
            if ok:
                st.success(f"Stav nádrže nastaven na {korekce_litry:.0f} L")
                del st.session_state["nacteno_pro"]
                st.session_state.tank_level, st.session_state.ceny_df, st.session_state.df_tank_vse = get_tank_info(token)
                st.rerun()
            else:
                st.error("Nepodařilo se uložit.")
        if st.button("🚫 Prázdná nádrž (0 L)"):
            with st.spinner("Ukládám..."):
                ok = ulozit_korekci_nadrze(token, 0.0, "Prázdná nádrž")
            if ok:
                st.success("Stav nádrže nastaven na 0 L")
                del st.session_state["nacteno_pro"]
                st.session_state.tank_level, st.session_state.ceny_df, st.session_state.df_tank_vse = get_tank_info(token)
                st.rerun()
            else:
                st.error("Nepodařilo se uložit.")

    st.divider()

    # ── Správa uživatelů ──────────────────────────────────────────────────────
    st.markdown("### 👥 Správa uživatelů")

    if "zam_list" not in st.session_state:
        with st.spinner("Načítám uživatele..."):
            st.session_state.zam_list = nacist_zamestnance_web(token)

    if st.button("🔄 Obnovit seznam"):
        with st.spinner("Načítám..."):
            st.session_state.zam_list = nacist_zamestnance_web(token)

    zam_list = st.session_state.zam_list
    if zam_list:
        for z in zam_list:
            with st.expander(f"**{z['jmeno']}**  —  PIN: {'*' * len(z['pin'])}  |  Limit: {z['limit']} Kč"):
                er1, er2, er3 = st.columns([1, 2, 2])
                new_titul    = er1.text_input("Titul",    value=z.get("title", ""),    key=f"ti_{z['id']}", placeholder="Ing.")
                new_jmeno    = er2.text_input("Jméno",    value=z.get("givenName", ""), key=f"jm_{z['id']}")
                new_prijmeni = er3.text_input("Příjmení", value=z.get("surname", ""),   key=f"pr_{z['id']}")
                ec1, ec2, ec3, ec4 = st.columns([2, 2, 2, 1])
                new_pin   = ec1.text_input("PIN",          value=z["pin"],               key=f"pin_{z['id']}")
                new_limit = ec2.number_input("Limit (Kč)", value=int(z["limit"] or 0),   min_value=0, step=100, key=f"lim_{z['id']}")
                new_spz   = ec3.text_input("Pref. SPZ",    value=z["spz"],               key=f"spz_{z['id']}")
                ec4.write("")
                ec4.write("")
                if ec4.button("💾", key=f"save_{z['id']}", help="Uložit změny"):
                    ok = upravit_zamestnance(token, z["id"], new_titul, new_jmeno, new_prijmeni, new_pin, new_limit, new_spz)
                    if ok:
                        st.success("Uloženo")
                        st.session_state.zam_list = nacist_zamestnance_web(token)
                        st.rerun()
                    else:
                        st.error("Chyba při ukládání")
                if ec4.button("🗑", key=f"del_{z['id']}", help="Smazat uživatele"):
                    ok = smazat_zamestnance(token, z["id"])
                    if ok:
                        st.success("Smazáno")
                        st.session_state.zam_list = nacist_zamestnance_web(token)
                        st.rerun()
                    else:
                        st.error("Chyba při mazání")
    else:
        st.info("Složka Zaměstnanci je prázdná nebo nebyla nalezena.")

    st.markdown("#### ➕ Přidat uživatele")
    with st.form("novy_zamestnanec"):
        fc0, fc1, fc2 = st.columns([1, 2, 2])
        n_titul   = fc0.text_input("Titul", placeholder="Ing.")
        n_jmeno   = fc1.text_input("Jméno *")
        n_prijmeni= fc2.text_input("Příjmení *")
        fc3, fc4, fc5 = st.columns(3)
        n_pin     = fc3.text_input("PIN *")
        n_limit   = fc4.number_input("Limit proplácení (Kč)", min_value=0, step=100)
        n_spz     = fc5.text_input("Preferovaná SPZ")
        submitted = st.form_submit_button("Přidat", type="primary")
        if submitted:
            if not n_jmeno or not n_prijmeni or not n_pin:
                st.error("Jméno, příjmení a PIN jsou povinné.")
            else:
                with st.spinner("Ukládám..."):
                    ok = pridat_zamestnance(token, n_titul, n_jmeno, n_prijmeni, n_pin, n_limit, n_spz)
                if ok:
                    st.success(f"Uživatel {n_jmeno} {n_prijmeni} přidán.")
                    st.session_state.zam_list = nacist_zamestnance_web(token)
                    st.rerun()
                else:
                    st.error("Nepodařilo se přidat uživatele. Zkontrolujte oprávnění (Contacts.ReadWrite).")

    st.divider()

    # ── Správa vozidel ────────────────────────────────────────────────────────
    st.markdown("### 🚗 Správa vozidel")

    if "voz_list" not in st.session_state:
        with st.spinner("Načítám vozidla..."):
            st.session_state.voz_list = nacist_vozidla_web(token)

    if st.button("🔄 Obnovit seznam vozidel"):
        with st.spinner("Načítám..."):
            st.session_state.voz_list = nacist_vozidla_web(token)

    voz_list = st.session_state.voz_list
    if voz_list:
        for v in voz_list:
            with st.expander(f"**{v['spz']}** — {v['typ']}"):
                vc1, vc2, vc3 = st.columns([2, 2, 1])
                v_typ = vc1.text_input("Typ vozidla", value=v["typ"], key=f"vtyp_{v['id']}")
                v_spz = vc2.text_input("SPZ", value=v["spz"], key=f"vspz_{v['id']}")
                vc3.write("")
                vc3.write("")
                if vc3.button("💾", key=f"vsave_{v['id']}", help="Uložit"):
                    ok = upravit_vozidlo(token, v["id"], v_typ, v_spz)
                    if ok:
                        st.success("Uloženo")
                        st.session_state.voz_list = nacist_vozidla_web(token)
                        st.rerun()
                    else:
                        st.error("Chyba při ukládání")
                if vc3.button("🗑", key=f"vdel_{v['id']}", help="Smazat"):
                    ok = smazat_vozidlo(token, v["id"])
                    if ok:
                        st.success("Smazáno")
                        st.session_state.voz_list = nacist_vozidla_web(token)
                        st.rerun()
                    else:
                        st.error("Chyba při mazání")
    else:
        st.info("Složka Vozidla je prázdná nebo nebyla nalezena.")

    st.markdown("#### ➕ Přidat vozidlo")
    with st.form("nove_vozidlo"):
        vf1, vf2 = st.columns(2)
        nv_typ = vf1.text_input("Typ vozidla (např. Octavia)")
        nv_spz = vf2.text_input("SPZ (např. 9AF 1888)")
        if st.form_submit_button("Přidat vozidlo", type="primary"):
            if not nv_spz:
                st.error("SPZ je povinná.")
            else:
                with st.spinner("Ukládám..."):
                    ok = pridat_vozidlo(token, nv_typ, nv_spz)
                if ok:
                    st.success(f"Vozidlo {nv_spz} přidáno.")
                    st.session_state.voz_list = nacist_vozidla_web(token)
                    st.rerun()
                else:
                    st.error("Nepodařilo se přidat vozidlo.")
