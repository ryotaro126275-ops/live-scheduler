import csv
import hashlib
import io
import subprocess
import sys
import tempfile
from pathlib import Path
import pandas as pd
import streamlit as st
import scheduler_v6 as v3
from sample_data import sample_bytes
import json
import altair as alt

ROOT = Path(__file__).resolve().parent
st.set_page_config(page_title='Live Scheduler', page_icon='🎸', layout='wide')
st.title('🎸 Live Scheduler')
st.caption('軽音サークル向け日割り自動生成ツール｜Ver.7')
st.info('出演順・休憩・掛け持ちは対象外です。')
st.subheader('開催条件')
defaults=pd.DataFrame({'開催日':['2026-12-08','2026-12-12','2026-12-13','2026-12-14','2026-12-19'],'開始':['18:00','18:00','13:00','13:00','13:00'],'終了':['22:00']*5})
days=st.data_editor(defaults,num_rows='dynamic',hide_index=True,key='days')
cols=st.columns(3)
durations=[int(c.number_input(label,min_value=1,max_value=1440,value=n,step=1)) for c,label,n in zip(cols,['1曲（分）','2曲（分）','3曲以上（分）'],[15,20,25])]
try:
    from datetime import date
    hours={}
    for d,a,z in days.itertuples(index=False,name=None):
        d=date.fromisoformat(str(d)).isoformat(); a,z=v3.minutes(a),v3.minutes(z)
        if d in hours or a>=z: raise ValueError('日付重複または開始・終了時刻が不正です')
        hours[d]=[a,z]
    if not hours: raise ValueError('開催日を追加してください')
except Exception as e:
    st.error(f'開催条件: {e}'); st.stop()
config={'hours':hours,'durations':durations}; config_bytes=json.dumps(config,sort_keys=True).encode()
with st.expander('入力Excelの形式・サンプル'):
    st.markdown('先頭シートの1行目は見出し。A列：メンバー名、B列：学年（1〜4）、G列：曲数、H列：バンド名、J〜P列：出演メンバー、R列：出演可能日。サンプルは架空のデータです。')
    st.download_button('サンプルExcelをダウンロード', sample_bytes(), 'sample_data.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
source = st.file_uploader('1. 元のExcelファイル', type=['xlsx'])
if source is None:
    st.stop()
data = source.getvalue()
key = hashlib.sha256(data+config_bytes).hexdigest()
try:
    bands = v3.load_bands(io.BytesIO(data), durations)
    if not bands:
        raise ValueError('バンドがありません')
except Exception as e:
    st.error(f'Excel読み込みエラー: {e}')
    st.stop()
if st.session_state.get('source_key') != key:
    st.session_state.source_key = key
    st.session_state.review = pd.DataFrame([dict(row=b['row'], band=b['name'], original=b['raw'], windows=v3.suggest(b['raw'], hours)[0], confirmed='yes' if v3.suggest(b['raw'], hours)[1] else 'no') for b in bands])
    st.session_state.pop('result', None)
    st.session_state.pop('csv_key', None)
    st.session_state.revision = 0
st.subheader('2. 出演可能条件の確認')
review_file = st.file_uploader('確認済みCSVがあれば読み込む（任意）', type=['csv'], key='csv_'+key)
if review_file:
    raw = review_file.getvalue()
    ck = hashlib.sha256(raw).hexdigest()
    if st.session_state.get('csv_key') != ck:
        try:
            df = pd.read_csv(io.BytesIO(raw), encoding='utf-8-sig', dtype=str, keep_default_na=False)
            cols = ['row','band','original','windows','confirmed']
            df = df[cols]; df['row'] = df['row'].astype(int)
            expected = {b['row']:(b['name'],b['raw']) for b in bands}
            actual = {r.row:(r.band,r.original) for r in df.itertuples()}
            if len(df)!=len(bands) or actual!=expected:
                raise ValueError('元ExcelとCSVが一致しません')
            st.session_state.review = df
            st.session_state.csv_key = ck
            st.session_state.revision += 1
        except Exception as e:
            st.error(f'CSV読み込みエラー: {e}')
            st.stop()
st.caption('windows の例：2026-12-14;2026-12-19@19:00-22:00（旧CSVの日番号も一意なら使用可）。確認後 confirmed を yes に変更してください。')
edited = st.data_editor(st.session_state.review, disabled=['row','band','original'], hide_index=True, column_config={'confirmed':st.column_config.SelectboxColumn('confirmed',options=['yes','no'],required=True)}, key=f'editor_{key}_{st.session_state.revision}')
confirmed = int((edited['confirmed']=='yes').sum())
a,b = st.columns(2); a.metric('登録バンド',len(bands)); b.metric('確認済み',confirmed)
review_bytes = edited.to_csv(index=False).encode('utf-8-sig')
st.download_button('確認用CSVを保存',review_bytes,'availability_review.csv','text/csv')
st.subheader('3. 日割り生成・検証')
fingerprint = hashlib.sha256(data+review_bytes+config_bytes).hexdigest()
if st.button('日割りを自動生成する', type='primary', disabled=confirmed!=len(bands)):
    st.session_state.pop('result',None)
    try:
        with st.spinner('日割りを計算・検証しています（数分かかる場合があります）…'):
            with tempfile.TemporaryDirectory() as tmp:
                p=Path(tmp); (p/'config.json').write_bytes(config_bytes); (p/'input.xlsx').write_bytes(data); (p/'availability_review.csv').write_bytes(review_bytes)
                def run(script):
                    return subprocess.run([sys.executable,str(ROOT/script),'solve','input.xlsx'] if script=='scheduler_v6.py' else [sys.executable,str(ROOT/script),'input.xlsx'],cwd=p,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=360,env={**__import__('os').environ,'PYTHONIOENCODING':'utf-8'})
                result=run('scheduler_v6.py')
                if result.returncode or not (p/'schedule_result.xlsx').exists():
                    raise RuntimeError(result.stdout+'\n'+result.stderr)
                verify=run('verify_schedule_v6.py')
                st.session_state.result=(fingerprint,(p/'schedule_result.xlsx').read_bytes(),result.stdout,verify.stdout+'\n'+verify.stderr,verify.returncode)
    except Exception as e:
        st.error(str(e))
result=st.session_state.get('result')
if result and result[0]==fingerprint:
    _, xlsx, log, verification, code=result
    st.subheader('4. 日割り結果')
    if code==0 and verification.strip().endswith('PASS'):
        st.success('PASS：検証成功')
    else:
        st.warning('検証未通過：未割当や出演条件の矛盾、検証時間切れがないかログを確認してください。')
    summary=pd.read_excel(io.BytesIO(xlsx),sheet_name='集計')
    st.dataframe(summary,hide_index=True)
    for fields, title in [(['使用分数','残り分数'],'開催時間（分）'),(['1年','2年','3年','4年'],'学年別延べ人数')]:
        chart=summary.melt(id_vars='開催日',value_vars=fields,var_name='項目',value_name='値')
        st.altair_chart(alt.Chart(chart).mark_bar().encode(x=alt.X('開催日:N',sort=sorted(hours)),y=alt.Y('値:Q',title=title),color='項目:N',tooltip=['開催日','項目','値']),use_container_width=True)
    detail=pd.read_excel(io.BytesIO(xlsx),sheet_name='日割り')
    unassigned=int((detail['開催日']=='未割当').sum())
    if unassigned: st.warning(f'{unassigned}組が未割当です。開催時間や出演可能条件を見直してください。')
    choice=st.selectbox('開催日で絞り込む',['すべて']+sorted(hours)+['未割当'])
    st.dataframe(detail if choice=='すべて' else detail[detail['開催日']==choice],hide_index=True)
    with st.expander('計算・検証ログ'):
        st.code(log+'\n'+verification)
    st.download_button('日割り結果をExcelで保存',xlsx,'schedule_result.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.caption('出力の行順は出演順ではありません。')
