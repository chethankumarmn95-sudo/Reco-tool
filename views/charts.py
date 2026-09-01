"""
charts.py
---------
Wraps plotly chart creation so a missing/outdated plotly install degrades
to Streamlit's built-in charts instead of crashing the whole page. This is
exactly what broke last time: a fresh environment that hadn't re-run
`pip install -r requirements.txt` after plotly was added had no plotly at
all, and the ImportError took down the entire Dashboard tab.
"""

import streamlit as st

try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False


def plotly_missing_notice():
    st.warning(
        "Charts are running in basic mode because the `plotly` package isn't installed. "
        "Run `pip install -r requirements.txt` in the reco_tool folder and restart the app "
        "for the full interactive charts."
    )


def bar_grouped(df, x, y_cols, title, height=360):
    if PLOTLY_AVAILABLE:
        fig = px.bar(df, x=x, y=y_cols, barmode="group", title=title,
                     labels={"value": "Amount (₹)", x: x.title(), "variable": ""})
        fig.update_layout(height=height, margin=dict(t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.caption(title)
        st.bar_chart(df.set_index(x)[y_cols])


def donut(df, names, values, title, height=360):
    if PLOTLY_AVAILABLE:
        fig = px.pie(df, names=names, values=values, title=title, hole=0.45)
        fig.update_layout(height=height, margin=dict(t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.caption(title)
        st.bar_chart(df.set_index(names)[values])


def horizontal_bar(df, x, y, title, height=340):
    if PLOTLY_AVAILABLE:
        fig = px.bar(df.sort_values(x, ascending=True), x=x, y=y, orientation="h",
                     title=title, labels={x: x.title(), y: ""})
        fig.update_layout(height=height, margin=dict(t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.caption(title)
        st.bar_chart(df.set_index(y)[x])


def line_trend(df, x, y, title, height=320):
    if PLOTLY_AVAILABLE:
        fig = px.line(df, x=x, y=y, title=title, markers=True)
        fig.update_layout(height=height, margin=dict(t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.caption(title)
        st.line_chart(df.set_index(x)[y])


def health_gauge(score_pct, title="Reconciliation Health Score"):
    if PLOTLY_AVAILABLE:
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=score_pct,
            number={"suffix": "%"},
            title={"text": title},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": "#6C5CE7"},
                "steps": [
                    {"range": [0, 60], "color": "#FDE2E1"},
                    {"range": [60, 85], "color": "#FFF3D6"},
                    {"range": [85, 100], "color": "#DFF5E1"},
                ],
            },
        ))
        fig.update_layout(height=260, margin=dict(t=40, b=10, l=20, r=20))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.metric(title, f"{score_pct:.1f}%")
