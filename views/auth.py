"""
auth.py
-------
A minimal login gate for the web-hosted version of this tool. This only
matters once the app is reachable over the internet - anyone who reaches
the URL must sign in with one of the accounts below before seeing any
reconciliation data. Not used when running the app locally/offline.

Passwords are stored here as salted SHA-256 hashes, never plaintext. To
add or change a user, call hash_password("newpassword") once (e.g. in a
throwaway python -c one-liner) to get its hash, then update USERS below.

The sign-in screen's look (see _LOGIN_CSS/_render_login_card below) reuses
the same brass/ink "RecoMatrix" visual identity as the public landing page
(recomatrix.com) - same fonts, palette and card treatment - so the jump
from the public site into this screen doesn't feel like two different
products. Styling is done via Streamlit's own stable data-testid
selectors and the documented st.container(key=...) -> .st-key-<key> CSS
hook, not by guessing at Streamlit's internal auto-generated class names
(those change between Streamlit versions; the testid/key attributes are
Streamlit's public, stable contract for exactly this kind of styling).
"""

import hashlib
import hmac
import streamlit as st

# A fixed salt is a reasonable tradeoff for a small internal tool with a
# handful of named users - it's not protecting against a large-scale
# credential-stuffing/rainbow-table attack the way a per-user random salt
# would, just against a password being stored in plain, readable text.
SALT = "reco-tool-2026"


def hash_password(password: str) -> str:
    return hashlib.sha256((SALT + password).encode()).hexdigest()


# username -> sha256(SALT + password)
USERS = {
    "Shreepriya": hash_password("Shree@123"),
    "Yashashwini": hash_password("Yash@123"),
    "Chethan": hash_password("Chethan@123"),
}


_BRAND_MARK_SVG = (
    '<svg viewBox="0 0 32 32" fill="none">'
    '<rect x="0.75" y="0.75" width="30.5" height="30.5" rx="7" stroke="currentColor" stroke-width="1.5"/>'
    '<path d="M9 16.5L13.6 21L23 10.5" stroke="currentColor" stroke-width="1.7" '
    'stroke-linecap="round" stroke-linejoin="round"/></svg>'
)

_LOGIN_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');

:root {
  --rm-bg: #F6F4EE;
  --rm-surface: #FFFFFF;
  --rm-ink: #14201F;
  --rm-ink-soft: #4B5654;
  --rm-ink-faint: #7C8785;
  --rm-line: rgba(20,32,31,0.14);
  --rm-line-strong: rgba(20,32,31,0.24);
  --rm-accent: #A97E14;
  --rm-accent-strong: #8E6A10;
  --rm-teal: #1F7A63;
  --rm-critical: #A5433B;
}
@media (prefers-color-scheme: dark) {
  :root {
    --rm-bg: #0E1B1A;
    --rm-surface: #142524;
    --rm-ink: #F2EEE3;
    --rm-ink-soft: #B9C3C0;
    --rm-ink-faint: #7E8B88;
    --rm-line: rgba(242,238,227,0.14);
    --rm-line-strong: rgba(242,238,227,0.26);
    --rm-accent: #DCB454;
    --rm-accent-strong: #EAC876;
    --rm-teal: #54D8B8;
    --rm-critical: #E58076;
  }
}

/* Full-page cream canvas, card vertically + horizontally centered */
[data-testid="stAppViewContainer"], [data-testid="stApp"] { background: var(--rm-bg) !important; }
[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stToolbar"], [data-testid="stMainMenu"], [data-testid="stAppDeployButton"] { visibility: hidden !important; }
footer { visibility: hidden !important; }

[data-testid="stMain"] {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 100vh;
}
[data-testid="stMainBlockContainer"] {
  max-width: 448px;
  width: 100%;
  padding-top: 6vh;
  padding-bottom: 6vh;
}

.rm-login-brand {
  display: flex; align-items: center; justify-content: center; gap: 10px;
  margin-bottom: 30px;
}
.rm-login-brand svg { width: 30px; height: 30px; color: var(--rm-ink); flex: none; }
.rm-login-brand span {
  font-family: 'Instrument Serif', Georgia, serif; font-style: italic;
  font-size: 1.55rem; color: var(--rm-ink);
}

/* The sign-in card itself - a Streamlit container targeted via its stable
   st-key-* class (see _render_login_card below: st.container(key="rm_login_card")) */
div.st-key-rm_login_card {
  background: var(--rm-surface);
  border: 1px solid var(--rm-line);
  border-radius: 8px;
  padding: 40px 36px 30px;
  box-shadow: 0 1px 2px rgba(20,32,31,0.06), 0 20px 44px -20px rgba(20,32,31,0.24);
  position: relative;
  overflow: hidden;
}
div.st-key-rm_login_card::before {
  content: ""; position: absolute; inset: 0 0 auto 0; height: 4px;
  background: linear-gradient(90deg, var(--rm-accent), var(--rm-teal));
}

.rm-login-eyebrow {
  display: block; text-align: center;
  font-family: 'IBM Plex Mono', monospace; font-size: 0.7rem; font-weight: 600;
  letter-spacing: 0.12em; text-transform: uppercase; color: var(--rm-accent-strong);
  margin-bottom: 12px;
}
.rm-login-title {
  font-family: 'Instrument Serif', Georgia, serif; font-weight: 400; font-size: 1.9rem;
  color: var(--rm-ink); text-align: center; margin: 0 0 6px;
}
.rm-login-sub {
  font-family: 'IBM Plex Sans', sans-serif; font-size: 0.88rem; color: var(--rm-ink-soft);
  text-align: center; margin: 0 0 26px;
}
.rm-login-fine {
  margin-top: 22px; font-size: 0.78rem; color: var(--rm-ink-faint); text-align: center;
  font-family: 'IBM Plex Sans', sans-serif; line-height: 1.5;
}

/* Form fields - targeted via Streamlit's stable data-testid hooks rather
   than its internal (version-dependent) generated class names. */
div[data-testid="stTextInput"] label p {
  font-family: 'IBM Plex Sans', sans-serif !important; font-size: 0.72rem !important;
  font-weight: 600 !important; text-transform: uppercase; letter-spacing: 0.06em;
  color: var(--rm-ink-faint) !important;
}
div[data-testid="stTextInputRootElement"] {
  background: var(--rm-bg) !important;
  border: 1px solid var(--rm-line-strong) !important;
  border-radius: 4px !important;
  box-shadow: none !important;
}
div[data-testid="stTextInputRootElement"]:focus-within {
  border-color: var(--rm-accent) !important;
  box-shadow: 0 0 0 1px var(--rm-accent) !important;
}
div[data-testid="stTextInputRootElement"] input {
  font-family: 'IBM Plex Sans', sans-serif !important;
  color: var(--rm-ink) !important;
  padding: 11px 12px !important;
}

div[data-testid="stFormSubmitButton"] button {
  width: 100%;
  background: var(--rm-accent) !important;
  color: #14201F !important;
  border: none !important;
  border-radius: 4px !important;
  font-family: 'IBM Plex Sans', sans-serif !important;
  font-weight: 600 !important;
  padding: 12px !important;
  margin-top: 6px;
  transition: background 0.15s ease;
}
div[data-testid="stFormSubmitButton"] button:hover { background: var(--rm-accent-strong) !important; }
div[data-testid="stFormSubmitButton"] button p { color: inherit !important; }

@media (max-width: 520px) {
  [data-testid="stMainBlockContainer"] { padding-top: 4vh; max-width: 94vw; }
  div.st-key-rm_login_card { padding: 32px 22px 24px; }
}
</style>
"""


def _render_login_card():
    """Renders the sign-in card's chrome (brand, heading, form) - the
    actual username/password/submit widgets are real Streamlit widgets
    (needed for session_state + st.form to work), just styled via the CSS
    above. Returns (username, password, submitted)."""
    st.markdown(_LOGIN_CSS, unsafe_allow_html=True)
    st.markdown(
        f'<div class="rm-login-brand">{_BRAND_MARK_SVG}<span>RecoMatrix</span></div>',
        unsafe_allow_html=True,
    )

    with st.container(key="rm_login_card"):
        st.markdown(
            """
            <span class="rm-login-eyebrow">Authorised Portal Access</span>
            <div class="rm-login-title">Sign in</div>
            <div class="rm-login-sub">Enter your credentials to continue.</div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign in")
        st.markdown(
            '<div class="rm-login-fine">Access is by authorised account only.'
            '<br>Forgot your credentials? Contact your administrator.</div>',
            unsafe_allow_html=True,
        )

    return username, password, submitted


def check_login() -> bool:
    """Renders a login form and returns True once the visitor has signed
    in with a valid username/password (remembered for the rest of this
    browser session). Call this right after st.set_page_config(), before
    building anything else - render nothing else until it returns True."""
    if st.session_state.get("authenticated"):
        return True

    username, password, submitted = _render_login_card()

    if submitted:
        expected = USERS.get(username)
        if expected and hmac.compare_digest(expected, hash_password(password)):
            st.session_state["authenticated"] = True
            st.session_state["username"] = username
            st.rerun()
        else:
            st.error("Incorrect username or password.")

    return False


def logout_button():
    """Small sidebar footer showing who's signed in, with a log-out button.
    Call this once, inside the same `with st.sidebar:` block app.py already
    uses for the Reco Tool header."""
    if st.session_state.get("authenticated"):
        st.caption(f"Signed in as **{st.session_state.get('username')}**")
        if st.button("Log out"):
            st.session_state["authenticated"] = False
            st.rerun()
