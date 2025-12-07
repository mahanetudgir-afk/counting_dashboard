"""
Streamlit Data Analyzer & Live Poll-Day Dashboard
File: streamlit_app.py

Features implemented in this single-file Streamlit app:
- Upload Excel/CSV and paste Google Sheet URL (public) or use OAuth-based private Google Sheets access (instructions + helper).
- Auto-detection of column types (number/date/category) and automatic time column detection.
- Booth-wise comparison templates + pre-built political-polling dashboards.
- Automatic gender inference from name columns (local heuristic using `gender-guesser` and optional Genderize.io API fallback).
- Real-time auto-refresh (polling) for live poll-day operations with adjustable interval.
- Exports: polished Excel (multiple sheets) and multipage PDF export (charts + summary).

Dependencies (pip):
streamlit pandas numpy plotly openpyxl xlrd xlsxwriter requests gender-guesser gspread oauth2client reportlab pillow

Notes on Google OAuth / private sheets:
- Two options: (A) Service Account (recommended for server-to-server access) — you create a GCP service account, share the sheet with the service account email, and provide the JSON key file to the deployed app (securely). (B) OAuth 2.0 user flow (requires a small backend to securely hold client secrets and exchange tokens). Below the code I include helper functions and comments for both.

To run locally:
1) pip install -r requirements.txt (or the list above)
2) streamlit run streamlit_app.py

"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import plotly.express as px
import requests
from datetime import datetime
from gender_guesser.detector import Detector
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from PIL import Image
import base64

# Optional Google Sheets
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GS_AVAILABLE = True
except Exception:
    GS_AVAILABLE = False

# ------------------------- Helpers -------------------------

def infer_column_types(df, sample_n=100):
    cols = []
    for c in df.columns:
        sample = df[c].dropna().astype(str).head(sample_n)
        numeric = sample.map(lambda x: is_number(x)).sum()
        dates = sample.map(lambda x: is_date(x)).sum()
        unique = sample.nunique()
        total = len(sample) if len(sample)>0 else 1
        if dates/total > 0.6:
            t = 'date'
        elif numeric/total > 0.7:
            t = 'number'
        elif unique < 50:
            t = 'category'
        else:
            t = 'text'
        cols.append({'column':c, 'type':t, 'unique':unique})
    return pd.DataFrame(cols)


def is_number(x):
    try:
        float(str(x).replace(',',''))
        return True
    except:
        return False


def is_date(x):
    try:
        pd.to_datetime(x)
        return True
    except:
        return False


def read_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return None
    name = uploaded_file.name.lower()
    if name.endswith('.csv'):
        return pd.read_csv(uploaded_file)
    else:
        return pd.read_excel(uploaded_file, engine='openpyxl')


def fetch_google_sheet_csv(sheet_url):
    # Accepts a normal googlesheet url and converts to export CSV
    # Sheet must be shared 'anyone with link' or be accessible to credentials used
    if '/d/' in sheet_url:
        try:
            id_part = sheet_url.split('/d/')[1].split('/')[0]
            csv_url = f'https://docs.google.com/spreadsheets/d/{id_part}/export?format=csv'
            r = requests.get(csv_url)
            r.raise_for_status()
            return pd.read_csv(io.StringIO(r.text))
        except Exception as e:
            st.error(f'Error fetching public Google Sheet: {e}')
            return None
    else:
        st.error('Unrecognized Google Sheet URL')
        return None


def infer_gender_series(names_series, use_genderize=False, genderize_api_key=None):
    det = Detector(case_sensitive=False)
    genders = []
    for name in names_series.fillna(''):
        first = str(name).strip().split()[0] if str(name).strip() else ''
        if not first:
            genders.append(None); continue
        # local heuristic
        g = det.get_gender(first)
        # map detector outputs to male/female/unknown
        if g in ['male','mostly_male']:
            genders.append('male'); continue
        if g in ['female','mostly_female']:
            genders.append('female'); continue
        # fallback: use Genderize.io if requested
        if use_genderize and genderize_api_key:
            try:
                resp = requests.get('https://api.genderize.io', params={'name': first, 'apikey': genderize_api_key}).json()
                if 'gender' in resp and resp['gender']:
                    genders.append(resp['gender']); continue
            except:
                pass
        genders.append('unknown')
    return pd.Series(genders)


def make_pivot_booth(df, booth_col, value_col, agg='sum'):
    if agg=='sum':
        pivot = df.groupby(booth_col)[value_col].sum().reset_index().sort_values(value_col, ascending=False)
    elif agg=='mean':
        pivot = df.groupby(booth_col)[value_col].mean().reset_index().sort_values(value_col, ascending=False)
    else:
        pivot = df.groupby(booth_col)[value_col].count().reset_index().sort_values(value_col, ascending=False)
    return pivot


def export_excel_bytes(df, charts=[]):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='data', index=False)
        workbook = writer.book
        # optionally add charts as images in a sheet
        for i, img in enumerate(charts):
            # img is PIL image
            img_byte = io.BytesIO()
            img.save(img_byte, format='PNG')
            img_byte.seek(0)
            worksheet = writer.sheets.get('Charts')
            if worksheet is None:
                worksheet = workbook.add_worksheet('Charts')
                writer.sheets['Charts'] = worksheet
            worksheet.insert_image(1, i*10, f'chart{i}.png', {'image_data': img_byte})
    out.seek(0)
    return out.read()


def export_pdf_bytes(summary_text, chart_images=[]):
    out = io.BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph('Data Analyzer Summary', styles['Title']))
    story.append(Spacer(1,12))
    for line in summary_text.split('\n'):
        story.append(Paragraph(line, styles['Normal']))
        story.append(Spacer(1,6))
    for img in chart_images:
        # Convert PIL Image to temporary file-like
        b = io.BytesIO()
        img.save(b, format='PNG')
        b.seek(0)
        story.append(RLImage(b, width=450, height=250))
        story.append(Spacer(1,12))
    doc.build(story)
    out.seek(0)
    return out.read()

# ------------------------- Streamlit UI -------------------------

st.set_page_config(page_title='Data Analyzer — Polling Dashboard', layout='wide')
st.title('Data Analyzer — Polling & Booth Dashboard (Streamlit)')

# Sidebar: data source
st.sidebar.header('Data source')
upload = st.sidebar.file_uploader('Upload Excel/CSV', type=['csv','xlsx','xls'])
use_gsheet = st.sidebar.text_input('Or paste a public Google Sheet URL (CSV export)')

# Optional: service account credentials for gspread (private sheets)
st.sidebar.markdown('---')
st.sidebar.subheader('Private Google Sheets (Service account)')
sa_file = st.sidebar.file_uploader('Upload service account JSON (optional)', type=['json'])

# Auto-refresh settings
st.sidebar.markdown('---')
st.sidebar.subheader('Live refresh (polling)')
auto_refresh = st.sidebar.checkbox('Enable auto-refresh (poll every N seconds)', value=False)
refresh_interval = st.sidebar.number_input('Interval (seconds)', value=30, min_value=5)

# Genderize API (optional)
st.sidebar.markdown('---')
st.sidebar.subheader('Gender inference')
use_genderize = st.sidebar.checkbox('Use Genderize.io API fallback?', value=False)
genderize_key = None
if use_genderize:
    genderize_key = st.sidebar.text_input('Genderize.io API key (optional)')

# Load data
df = None
if upload is not None:
    df = read_uploaded_file(upload)
    st.sidebar.success('Loaded uploaded file')
elif use_gsheet:
    df = fetch_google_sheet_csv(use_gsheet)
    if df is not None:
        st.sidebar.success('Loaded public Google Sheet')
elif sa_file is not None and GS_AVAILABLE:
    # Use service account to access a private sheet (user must still provide sheet ID below)
    st.sidebar.info('Service Account uploaded — use the Sheet ID below to load private sheet')
    try:
        creds_json = sa_file.getvalue()
        creds_dict = creds_json
        # We'll pass this to a helper once user provides sheet id
    except Exception as e:
        st.sidebar.error('Could not read service account file')

# If service account and sheet id provided
if GS_AVAILABLE and sa_file is not None:
    private_sheet_id = st.sidebar.text_input('Private Sheet ID (for service account)', '')
    if private_sheet_id and st.sidebar.button('Load private sheet'):
        try:
            scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_dict(sa_file.getvalue(), scope)
            client = gspread.authorize(creds)
            sh = client.open_by_key(private_sheet_id)
            ws = sh.get_worksheet(0)
            df = pd.DataFrame(ws.get_all_records())
            st.sidebar.success('Loaded private sheet')
        except Exception as e:
            st.sidebar.error(f'Failed to load private sheet: {e}')

if df is None:
    st.info('Upload a file or paste a Google Sheet URL to begin. You can also upload a service account JSON and private sheet id to access private Google Sheets.')
    st.stop()

# Post-load processing
st.write(f'Data loaded — {df.shape[0]} rows, {df.shape[1]} columns')
col_types = infer_column_types(df)
st.dataframe(col_types)

# Allow user to pick key columns
st.sidebar.markdown('---')
st.sidebar.subheader('Auto-detect columns')
possible_booth_cols = [c for c in df.columns if 'booth' in c.lower() or 'polling' in c.lower() or 'station' in c.lower()]
booth_col = st.sidebar.selectbox('Booth / Station column (for booth-wise comparison)', options=['']+list(df.columns), index=0)
name_col = st.sidebar.selectbox('Name column (for gender inference)', options=['']+list(df.columns), index=0)
value_col = st.sidebar.selectbox('Numeric column (votes / counts)', options=['']+list(df.select_dtypes(include=[np.number]).columns), index=0)

# Gender inference
if name_col:
    st.write('Running gender inference on name column...')
    genders = infer_gender_series(df[name_col], use_genderize=use_genderize, genderize_api_key=genderize_key)
    df['_inferred_gender'] = genders
    st.write(df[[name_col, '_inferred_gender']].head(10))

# Booth-wise comparison templates
st.sidebar.markdown('---')
st.sidebar.subheader('Templates')
template = st.sidebar.selectbox('Choose a template', options=['Default summary','Booth comparison — top booths','Time-series by booth (if timestamp)','Gender split by booth'])

st.header('Dashboard')

col1, col2 = st.columns((2,1))

with col1:
    if template=='Default summary':
        st.subheader('Quick summaries')
        st.metric('Rows', df.shape[0])
        if value_col:
            st.metric('Total (sum)', float(df[value_col].sum()))
            st.metric('Mean', float(df[value_col].mean()))
        st.subheader('Top columns')
        st.dataframe(col_types.sort_values('type').head(20))

    elif template=='Booth comparison — top booths':
        st.subheader('Booth-wise comparison')
        if not booth_col or not value_col:
            st.warning('Pick both a booth column and a numeric value column in the sidebar')
        else:
            pivot = make_pivot_booth(df, booth_col, value_col, agg='sum')
            st.dataframe(pivot.head(50))
            fig = px.bar(pivot.head(30), x=booth_col, y=value_col, title='Votes by booth (top 30)')
            st.plotly_chart(fig, use_container_width=True)

    elif template=='Time-series by booth (if timestamp)':
        st.subheader('Time-series by booth')
        time_cols = [c for c,t in zip(col_types['column'], col_types['type']) if t=='date']
        time_col = st.selectbox('Choose time column', options=['']+time_cols)
        if time_col and booth_col and value_col:
            tmp = df[[time_col, booth_col, value_col]].copy()
            tmp[time_col] = pd.to_datetime(tmp[time_col])
            agg = tmp.groupby([pd.Grouper(key=time_col, freq='15T'), booth_col])[value_col].sum().reset_index()
            fig = px.line(agg, x=time_col, y=value_col, color=booth_col, title='Time-series by booth')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info('Need a time column, booth column, and value column to build this template')

    elif template=='Gender split by booth':
        st.subheader('Gender split')
        if '_inferred_gender' not in df.columns or not booth_col:
            st.warning('Run gender inference (choose name column) and select booth column in sidebar')
        else:
            tmp = df.groupby([booth_col, '_inferred_gender']).size().reset_index(name='count')
            fig = px.bar(tmp, x=booth_col, y='count', color='_inferred_gender', title='Gender split by booth')
            st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader('Controls & Export')
    if st.button('Download Excel (with charts)'):
        # Create simple charts as images using plotly and PIL
        charts = []
        try:
            if template=='Booth comparison — top booths' and booth_col and value_col:
                pivot = make_pivot_booth(df, booth_col, value_col)
                fig = px.bar(pivot.head(30), x=booth_col, y=value_col, title='Votes by booth')
                img_bytes = fig.to_image(format='png')
                charts.append(Image.open(io.BytesIO(img_bytes)))
        except Exception as e:
            st.error('Could not render chart image for Excel: '+str(e))
        excel_bytes = export_excel_bytes(df, charts=charts)
        st.download_button('Download .xlsx', data=excel_bytes, file_name='data_with_charts.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    if st.button('Download PDF summary'):
        summary_lines = [f'Rows: {df.shape[0]}', f'Columns: {df.shape[1]}']
        if value_col:
            summary_lines.append(f'Total {value_col}: {float(df[value_col].sum())}')
        # Add a chart image or two
        chart_images = []
        try:
            if booth_col and value_col:
                pivot = make_pivot_booth(df, booth_col, value_col)
                fig = px.bar(pivot.head(20), x=booth_col, y=value_col, title='Top booths')
                img_bytes = fig.to_image(format='png')
                chart_images.append(Image.open(io.BytesIO(img_bytes)))
        except Exception as e:
            st.error('Could not create chart for PDF: '+str(e))
        pdf_bytes = export_pdf_bytes('\n'.join(summary_lines), chart_images)
        st.download_button('Download summary PDF', data=pdf_bytes, file_name='summary.pdf', mime='application/pdf')

    st.markdown('---')
    st.subheader('Live refresh')
    if auto_refresh:
        st.info(f'Auto-refresh enabled: polling every {refresh_interval} seconds')
        # Try to use st.experimental_rerun with time check in session_state
        if 'last_refresh' not in st.session_state:
            st.session_state['last_refresh'] = time.time()
        last = st.session_state['last_refresh']
        now = time.time()
        if now - last > refresh_interval:
            st.session_state['last_refresh'] = now
            st.experimental_rerun()
        else:
            st.write(f'Next refresh in {int(refresh_interval - (now-last))}s')
    else:
        if st.button('Refresh now'):
            st.experimental_rerun()

# Small notes and OAuth instructions
st.sidebar.markdown('---')
st.sidebar.header('Google Sheets OAuth notes')
st.sidebar.markdown(
    """
    Two practical approaches to access private Google Sheets:
    1) Service Account (server-to-server): create a service account in GCP, grant it access to the spreadsheet (share spreadsheet with service account email), download JSON key and upload it here in the sidebar. This is easiest for automated server deployments but requires keeping the JSON secret safe (use secrets manager).

    2) OAuth 2.0 user flow: implement a small OAuth backend (Flask/FastAPI) that handles the OAuth exchange and stores refresh tokens for long-lived access. The Streamlit app would redirect the user to the backend to authenticate. This is required if the spreadsheet owner must explicitly consent using their Google account.
    """
)

st.sidebar.markdown('---')
st.sidebar.info('If you want, I can also generate a small Flask backend that implements the OAuth user flow and returns an endpoint the Streamlit app can call to fetch the sheet securely.')

# End of app


# ------------------------- Deployment notes -------------------------
# - For deployment to Streamlit Cloud, include your service account JSON in Secrets (NOT in repo). Use os.environ or st.secrets to access it securely.
# - For real-time large-scale deployments, consider adding a small caching layer (Redis) and a background worker to poll sheets and push updates.
# - If you'd like, I can also produce a Dockerfile + instructions for deploying with a small OAuth backend in Flask.

# ------------------------- EOF -------------------------
