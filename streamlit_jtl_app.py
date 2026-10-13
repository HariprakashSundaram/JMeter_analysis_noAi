"""Streamlit front-end for universal_jtl_analyzer.py; rule-based, no AI."""
import io
import os
import tempfile
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import pandas as pd
import streamlit as st
import universal_jtl_analyzer as core

st.set_page_config(page_title='JTL Journey Analytics', layout='wide')
st.title('JMeter JTL · Journey & Performance Analytics')
st.caption('CSV / XML · Multiple files · Parallel processing · No AI')

with st.sidebar:
    st.header('Analysis settings')
    workers = st.slider('Parallel files', 1, 8, 4)
    gap = st.number_input('Idle gap (seconds)', min_value=1, max_value=86400, value=120)
    p95_limit = st.number_input('P95 SLA (ms)', min_value=1, value=3000)
    error_limit = st.number_input('Error SLA (%)', min_value=0.0, max_value=100.0, value=1.0)
    st.info('Journey segmentation is heuristic unless an iteration identifier is present. Parent/child samples can overlap.')

uploads = st.file_uploader('Upload one or more JTL / CSV / XML files', type=['jtl', 'csv', 'xml'], accept_multiple_files=True)


def process_one(name, payload, gap_seconds, p95_sla, error_sla):
    with tempfile.TemporaryDirectory() as temp:
        source = Path(temp) / ('input' + Path(name).suffix.lower())
        source.write_bytes(payload)
        raw, fmt = core.read_input(str(source))
        df, mapping = core.normalize(raw)
        step = core.stats(df, 'label')
        flow = core.stats(df, 'flow')
        events, journey, note = core.journeys(df, gap_seconds)
        errors = (df[df['error_count'] > 0]
                  .groupby(['label', 'code', 'message'], dropna=False)
                  .agg(Failed_Rows=('label', 'size'), Estimated_Failed_Samples=('error_count', 'sum'))
                  .reset_index())
        transitions = pd.DataFrame()
        if not events.empty:
            ordered = events.sort_values(['Journey ID', 'Step']).copy()
            ordered['Next Step'] = ordered.groupby('Journey ID')['Label'].shift(-1)
            transitions = (ordered.dropna(subset=['Next Step'])
                           .groupby(['Flow', 'Label', 'Next Step']).size()
                           .reset_index(name='Observed Count')
                           .sort_values('Observed Count', ascending=False))
        findings = core.interpretations(df, step, flow, journey, note)
        # Additional UI-configurable threshold findings; original engine uses example defaults.
        custom = []
        for _, r in step.iterrows():
            if r['P95 ms'] > p95_sla:
                custom.append({'Severity': 'HIGH', 'Area': 'Custom P95 SLA', 'Evidence': f"{r['Name']}: {r['P95 ms']} ms", 'Interpretation': f'P95 exceeds {p95_sla} ms', 'Suggested Action': 'Investigate latency and dependency timing.'})
            if r['Error %'] > error_sla:
                custom.append({'Severity': 'HIGH', 'Area': 'Custom Error SLA', 'Evidence': f"{r['Name']}: {r['Error %']}%", 'Interpretation': f'Error rate exceeds {error_sla}%', 'Suggested Action': 'Review failed samples and application logs.'})
        if custom:
            findings = pd.concat([findings, pd.DataFrame(custom)], ignore_index=True)
        total = int(df['sample_count'].sum())
        duration = (df['end'].max() - df['timestamp'].min()) / 1000 if df['timestamp'].notna().any() else float('nan')
        summary = pd.DataFrame([
            ('Input file', name), ('Detected format', fmt), ('Recorded rows', len(df)),
            ('Estimated sample count', total), ('Distinct labels', df['label'].nunique()),
            ('Distinct threads', df['thread'].nunique()), ('Recorded failures', int(df['error_count'].sum())),
            ('Error %', round(100 * df['error_count'].sum() / total, 2)),
            ('Test wall duration seconds', round(duration, 2) if pd.notna(duration) else 'Unavailable'),
            ('Overall observed samples/sec', round(total / duration, 3) if pd.notna(duration) and duration > 0 else 'Unavailable'),
            ('Journey confidence', note), ('Configured P95 SLA ms', p95_sla), ('Configured Error SLA %', error_sla)
        ], columns=['Metric', 'Value'])
        quality = pd.DataFrame([{'Canonical Field': k, 'Source Column': mapping.get(k, 'NOT PRESENT')} for k in core.ALIASES])
        tables = {'Executive Summary': summary, 'Findings': findings, 'Step Analysis': step,
                  'Flow Analysis': flow, 'Journey Summary': journey, 'Journey Steps': events,
                  'Step Transitions': transitions, 'Error Analysis': errors, 'Data Mapping': quality}
        output = Path(temp) / 'analysis.xlsx'
        core.write_excel(tables, str(output))
        return {'name': name, 'tables': tables, 'xlsx': output.read_bytes(), 'rows': len(df),
                'errors': int(df['error_count'].sum()), 'samples': total, 'duration': duration}

if uploads and st.button('Analyze uploaded JTL files', type='primary'):
    # Snapshot uploaded bytes before concurrent workers; each worker uses its own temp directory.
    jobs = [(f.name, f.getvalue()) for f in uploads]
    results, failures = [], []
    progress = st.progress(0, text='Processing files concurrently...')
    with ThreadPoolExecutor(max_workers=min(workers, len(jobs))) as pool:
        futures = {pool.submit(process_one, name, data, gap, p95_limit, error_limit): name for name, data in jobs}
        for i, future in enumerate(as_completed(futures), 1):
            try:
                results.append(future.result())
            except Exception as exc:
                failures.append((futures[future], str(exc)))
            progress.progress(i / len(jobs), text=f'Processed {i} / {len(jobs)} files')
    st.session_state['jtl_results'] = results
    st.session_state['jtl_failures'] = failures

results = st.session_state.get('jtl_results', [])
failures = st.session_state.get('jtl_failures', [])
for name, reason in failures:
    st.error(f'{name}: {reason}')

if results:
    st.success(f'{len(results)} JTL file(s) analyzed')
    a, b, c = st.columns(3)
    a.metric('Recorded rows', f"{sum(r['rows'] for r in results):,}")
    b.metric('Estimated samples', f"{sum(r['samples'] for r in results):,}")
    c.metric('Recorded failures', f"{sum(r['errors'] for r in results):,}")
    overview = pd.DataFrame([{'File': r['name'], 'Rows': r['rows'], 'Samples': r['samples'], 'Failures': r['errors'], 'Wall duration (s)': round(r['duration'], 2) if pd.notna(r['duration']) else None} for r in results])
    st.subheader('Cross-file overview')
    st.dataframe(overview, use_container_width=True, hide_index=True)
    st.caption('Files are analyzed independently. Their timestamps, load levels, and sample populations may differ; aggregated numbers are not a single test-run TPS.')

    archive = io.BytesIO()
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
        for idx, r in enumerate(results, 1):
            safe = Path(r['name']).stem.replace('/', '_').replace('\\', '_')
            z.writestr(f'{idx:02d}_{safe}_analysis.xlsx', r['xlsx'])
    st.download_button('Download all Excel reports (ZIP)', archive.getvalue(), 'JTL_Analysis_Reports.zip', 'application/zip')

    for i, r in enumerate(results):
        with st.expander(f"{r['name']} · {r['rows']:,} rows", expanded=i == 0):
            st.download_button('Download Excel report', r['xlsx'], file_name=f"{Path(r['name']).stem}_analysis.xlsx", key=f'dl_{i}')
            tables = r['tables']
            tab_names = ['Findings', 'Step Analysis', 'Flow Analysis', 'Journey Summary', 'Journey Steps', 'Step Transitions', 'Error Analysis', 'Data Mapping']
            tabs = st.tabs(tab_names)
            for tab, key in zip(tabs, tab_names):
                with tab:
                    frame = tables[key]
                    if frame.empty:
                        st.info('No records available.')
                    else:
                        st.dataframe(frame, use_container_width=True, hide_index=True)
                    if key == 'Step Analysis' and not frame.empty:
                        chart = frame.set_index('Name')['P95 ms'].sort_values(ascending=False).head(20)
                        st.bar_chart(chart)
