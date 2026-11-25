"""Excelからライブ日割りを作成。実データは公開しないこと。"""
import csv
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path
from openpyxl import load_workbook, Workbook
from ortools.sat.python import cp_model

HOURS = {8:(1080,1320),12:(1080,1320),13:(780,1320),14:(780,1320),19:(780,1320)}
DURATIONS = (15,20,25)
REVIEW = Path('availability_review.csv')

def clock(n):
    return f'{n//60:02}:{n%60:02}'

def minutes(s):
    h,m = map(int,s.split(':'))
    if not (0 <= h <= 23 and 0 <= m < 60):
        raise ValueError('時刻が不正です')
    return h*60+m

def normalize(s):
    return unicodedata.normalize('NFKC',str(s or '')).replace('、',',').replace(' ','')

def resolve_day(token, hours=None):
    hours = HOURS if hours is None else hours
    if token in hours: return token
    if token.isdigit():
        matches=[d for d in hours if int(str(d)[-2:])==int(token)]
        if len(matches)==1: return matches[0]
    raise ValueError(f'開催日が不明または曖昧: {token}')

def suggest(raw, hours=None):
    s=normalize(raw)
    if re.fullmatch(r'\d+(?:[,と]\d+)*',s):
        try: return ';'.join(str(resolve_day(x, hours)) for x in re.split(r'[,と]',s)),True
        except ValueError: pass
    return '',False

def load_bands(path, durations=None):
    durations = DURATIONS if durations is None else durations
    rows=list(load_workbook(path,data_only=True).active.values)
    grades={str(r[0]).strip():int(r[1]) for r in rows[1:] if r[0] and r[1] in (1,2,3,4)}
    bands=[]
    for row,r in enumerate(rows[1:],2):
        if not r[7]: continue
        m=re.match(r'\d+',normalize(r[6]))
        if not m or int(m[0])<1: raise ValueError(f'Excel {row}行: 曲数が不正')
        songs=int(m[0]); duration=durations[min(songs,3)-1]
        members={str(v).strip() for v in r[9:16] if v}
        bands.append(dict(row=row,name=str(r[7]),songs=songs,duration=duration,raw=str(r[17] or ''),grades=Counter(grades[x] for x in members if x in grades),unknown=sum(x not in grades for x in members)))
    return bands

def prepare(bands):
    with REVIEW.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f); w.writerow(['row','band','original','windows','confirmed'])
        for b in bands:
            guess,ok=suggest(b['raw']); w.writerow([b['row'],b['name'],b['raw'],guess,'yes' if ok else 'no'])
    print(f'{REVIEW} を作成しました。no の行を確認し、windowsを修正してyesにしてください。')

def read_windows(bands):
    with REVIEW.open(encoding='utf-8-sig',newline='') as f:
        rows={int(r['row']):r for r in csv.DictReader(f)}
    result={}
    for b in bands:
        r=rows.get(b['row'])
        if not r or r['band']!=b['name'] or r['original']!=b['raw'] or r['confirmed'].strip().lower()!='yes':
            raise ValueError(f"{b['row']}行: 未確認または元データ変更。確認CSVを修正してください")
        windows=[]
        for item in r['windows'].split(';'):
            item=item.strip()
            if not item: continue
            d,_,span=item.partition('@'); d=resolve_day(d)
            if d not in HOURS: raise ValueError(f'{d}日は開催日ではありません')
            a,z=HOURS[d]
            if span:
                start,end=span.split('-'); x,y=minutes(start),minutes(end)
                if x>=y: raise ValueError(f'{item}: 時間範囲が不正')
                a,z=max(a,x),min(z,y)
            if z-a>=b['duration']: windows.append((d,a,z))
        result[b['row']]=windows
    return result

def solve(bands,windows):
    model=cp_model.CpModel(); intervals={d:[] for d in HOURS}; choices={}; assigned=[]; by_day={d:[] for d in HOURS}
    for b in bands:
        opts=[]; length=b['duration']; row=b['row']
        for j,(d,a,z) in enumerate(windows[row]):
            p=model.NewBoolVar(f'p{row}_{j}'); start=model.NewIntVar(a,z-length,f's{row}_{j}')
            intervals[d].append(model.NewOptionalFixedSizeIntervalVar(start,length,p,f'i{row}_{j}'))
            opts.append((d,p)); by_day[d].append((b,p))
        flag=model.NewBoolVar(f'assigned{row}'); model.Add(sum(p for d,p in opts)==flag); assigned.append(flag); choices[row]=opts
    for d in HOURS: model.AddNoOverlap(intervals[d])
    # 第一段階：出演できるバンド数を最大化
    model.Maximize(sum(assigned)); solver=cp_model.CpSolver(); solver.parameters.max_time_in_seconds=60
    status=solver.Solve(model)
    if status not in (cp_model.OPTIMAL,cp_model.FEASIBLE): raise RuntimeError('解が見つかりませんでした')
    best=sum(solver.Value(p) for p in assigned); first_status=solver.StatusName(status)
    selected={row:next((d for d,p in opts if solver.Value(p)),None) for row,opts in choices.items()}
    # 第二段階：開催時間に比例した演奏時間・学年別延べ人数の配分
    model.Add(sum(assigned)==best); penalties=[]
    capacity=sum(z-a for a,z in HOURS.values()); total=sum(b['duration'] for b in bands)
    for d,(a,z) in HOURS.items():
        cap=z-a; load=sum(b['duration']*p for b,p in by_day[d]); target=round(total*cap/capacity)
        delta=model.NewIntVar(0,total+capacity,f'load{d}'); model.AddAbsEquality(delta,load-target); penalties.append(delta)
        for g in range(1,5):
            n=sum(b['grades'][g] for b in bands); count=sum(b['grades'][g]*p for b,p in by_day[d]); target=round(n*cap/capacity)
            dev=model.NewIntVar(0,n,f'grade{d}_{g}'); model.AddAbsEquality(dev,count-target); penalties.append(5*dev)
    model.Minimize(sum(penalties)); status=solver.Solve(model)
    if status in (cp_model.OPTIMAL,cp_model.FEASIBLE):
        selected={row:next((d for d,p in opts if solver.Value(p)),None) for row,opts in choices.items()}
    print(f'割当 {best}/{len(bands)}組; 第1段階 {first_status}; 第2段階 {solver.StatusName(status)}')
    return selected

def export(bands,selected):
    wb=Workbook(); s=wb.active; s.title='日割り'; s.append(['開催日','バンド名','曲数','出演分数','1年','2年','3年','4年','学年不明','元の出演可能日','元Excel行'])
    for b in sorted(bands,key=lambda b:(selected[b['row']] or '9999',b['row'])):
        d=selected[b['row']]; s.append([d if d else '未割当',b['name'],b['songs'],b['duration'],*[b['grades'][g] for g in range(1,5)],b['unknown'],b['raw'],b['row']])
    t=wb.create_sheet('集計'); t.append(['開催日','バンド数','使用分数','上限分数','残り分数','1年','2年','3年','4年'])
    for d,(a,z) in HOURS.items():
        group=[b for b in bands if selected[b['row']]==d]; used=sum(b['duration'] for b in group); t.append([d,len(group),used,z-a,z-a-used,*[sum(b['grades'][g] for b in group) for g in range(1,5)]])
    for sheet in wb:
        sheet.freeze_panes='A2'; sheet.auto_filter.ref=sheet.dimensions
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width=min(55,max(12,max(len(str(c.value or '')) for c in col)*1.6))
    wb.save('schedule_result.xlsx'); print('schedule_result.xlsx を出力しました（行順は出演順ではありません）')

def configure(config):
    global HOURS,DURATIONS
    HOURS={d:tuple(t) for d,t in sorted(config['hours'].items())}
    DURATIONS=tuple(config['durations'])

if __name__=='__main__':
    import json
    configure(json.loads(Path('config.json').read_text(encoding='utf-8')))
    if len(sys.argv)!=3 or sys.argv[1] not in ('prepare','solve'):
        sys.exit('使い方: python scheduler_v6.py prepare "input.xlsx" または solve "input.xlsx"')
    try:
        bands=load_bands(sys.argv[2])
        if sys.argv[1]=='prepare':
            if REVIEW.exists(): sys.exit('確認CSVが既にあります。再作成する場合は退避してください。')
            prepare(bands)
        else: export(bands,solve(bands,read_windows(bands)))
    except (ValueError,KeyError,FileNotFoundError) as e: sys.exit(str(e))
