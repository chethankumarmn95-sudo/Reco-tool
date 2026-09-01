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


def check_login() -> bool:
    """Renders a login form and returns True once the visitor has signed
    in with a valid username/password (remembered for the rest of this
    browser session). Call this right after st.set_page_config(), before
    building anything else - render nothing else until it returns True."""
    if st.session_state.get("authenticated"):
        return True

    st.markdown("## Reco Tool — Sign in")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in")

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
