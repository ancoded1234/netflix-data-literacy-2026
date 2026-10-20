"""Reproduce stage 3: descriptive concentration analysis and validated charts."""
from pathlib import Path
from fractions import Fraction
import hashlib
import json
import math
import textwrap

import numpy as np
import pandas as pd
import openpyxl
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.ticker import PercentFormatter

ROOT = Path(__file__).resolve().parents[1]
RESULTS = ROOT / 'results'
CHARTS = RESULTS / 'charts'
TABLES = RESULTS / 'chart_data'
PAGE = 'https://about.netflix.com/en/news/what-we-watched-the-first-half-of-2026'
GROUPS = ['Movies', 'Shows', 'All']
LABEL = {'Movies': '영화', 'Shows': '시리즈·방송', 'All': '전체 개별 항목'}
COLORS = {'Movies': '#2563EB', 'Shows': '#E34252', 'All': '#6D4BD1'}
CHART_CHECKS = []


def check(condition, message):
    if not condition:
        raise ValueError(message)


def gini(values):
    x = sorted(int(v) for v in values)
    n, total = len(x), sum(x)
    # Exact integer numerator avoids floating point accumulation in the formula.
    numerator = sum((2 * i - n - 1) * v for i, v in enumerate(x, 1))
    return float(Fraction(numerator, n * total))


def ranked(frame, metric):
    # Equal published values: original sheet and row order, no inferred precision.
    return frame.sort_values([metric, 'source_sheet', 'source_row'], ascending=[False, True, True])


def md_table(frame, columns=None):
    if columns:
        frame = frame[columns]
    def cell(value):
        if isinstance(value, float):
            return f'{value:,.6f}'
        if isinstance(value, (int, np.integer)):
            return f'{value:,}'
        return str(value).replace('|', '\\|').replace('\n', ' ')
    return '\n'.join(['| ' + ' | '.join(frame.columns) + ' |',
                      '| ' + ' | '.join(['---'] * len(frame.columns)) + ' |'] +
                     ['| ' + ' | '.join(cell(v) for v in row) + ' |' for row in frame.itertuples(index=False, name=None)])


def main():
    for directory in [RESULTS, CHARTS, TABLES, ROOT / 'reports']:
        directory.mkdir(parents=True, exist_ok=True)
    check((ROOT / 'reports/cleaning_report.md').exists(), 'Missing cleaning report')
    paths = [ROOT / 'data/processed/netflix_cleaned.csv', ROOT / 'data/processed/netflix_other.csv',
             ROOT / 'data/raw/Netflix_What_We_Watched_2026_H1.xlsx']
    hashes = {str(p.relative_to(ROOT)): hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
    d = pd.read_csv(paths[0], dtype={'hours_viewed': 'int64', 'views': 'int64'})
    o = pd.read_csv(paths[1], dtype={'hours_viewed': 'int64', 'views': 'Int64'})
    check(d.groupby('content_type').size().to_dict() == {'Movies': 9177, 'Shows': 8193}, 'Counts changed')
    check(len(o) == 2 and set(o.title) == {'Other Movies', 'Other Shows'}, 'Unexpected Other')
    check(d.record_id.is_unique and not d.is_other.any() and o.is_other.all(), 'Invalid split or IDs')
    check(o.views.isna().all(), 'Other Views must remain missing')
    check(d[['hours_viewed', 'views']].notna().all().all() and (d[['hours_viewed', 'views']] > 0).all().all(), 'Invalid metrics')
    check(d.release_date.isna().sum() == 11324 and d.runtime_has_asterisk.sum() == 310, 'Rows filtered')
    # Independent cell-by-cell reconciliation: official Views, never reconstructed.
    wb = openpyxl.load_workbook(paths[2], read_only=True, data_only=True)
    official = {}
    for ws in wb:
        for rownum, row in enumerate(ws.iter_rows(min_row=7, min_col=2, max_col=7, values_only=True), 7):
            if row[3] is not None:
                official[f'{ws.title}:{rownum}'] = (ws.title, rownum, row[0], int(row[3]), row[5])
    wb.close()
    combined = pd.concat([d, o], ignore_index=True)
    check(set(combined.record_id) == set(official), 'Original IDs not preserved')
    for r in combined.itertuples():
        sheet, rownum, title, hours, views = official[r.record_id]
        check((r.source_sheet, r.content_type, r.source_row, r.title, r.hours_viewed)
              == (sheet, sheet, rownum, title, hours), f'Source mismatch {r.record_id}')
        check((views == '*' and pd.isna(r.views)) or views == r.views, f'Official Views mismatch {r.record_id}')
    check(o.set_index('content_type').title.to_dict() == {'Shows': 'Other Shows', 'Movies': 'Other Movies'}, 'Other type attribution unclear')
    # Semantics verified against the original sheet names and Other footnote.
    summary, concentration, tails, lorenz_rows, top_rows, ranking_rows, segments = [], [], [], [], [], [], []
    stats = {'period': '2026-01-01/2026-06-30', 'source_url': PAGE,
             'input_sha256': hashes, 'denominators': {
                 'A': 'Sum of Hours Viewed of disclosed individual rows within group; Other excluded',
                 'B': 'A plus same-group Other Hours Viewed; Other excluded from ranking and counts'},
             'top_n_rule': 'ceil(n*p/100); descending Hours Viewed, ties by original sheet and row',
             'bottom_half_rule': 'floor(n/2) rows sorted ascending Hours Viewed',
             'groups': {}}
    for group in GROUPS:
        frame = d if group == 'All' else d[d.content_type == group]
        other = o if group == 'All' else o[o.content_type == group]
        n = len(frame)
        A, oh = int(frame.hours_viewed.sum()), int(other.hours_viewed.sum())
        B = A + oh
        ordered = ranked(frame, 'hours_viewed')
        info = dict(n=n, hours_individual=A, hours_other=oh, hours_including_other=B,
                    views_individual=int(frame.views.sum()), views_including_other=None,
                    hours_mean=float(frame.hours_viewed.mean()), hours_median=float(frame.hours_viewed.median()),
                    hours_max=int(frame.hours_viewed.max()), gini_hours=gini(frame.hours_viewed),
                    release_date_missing=int(frame.release_date.isna().sum()),
                    runtime_asterisk_rows=int(frame.runtime_has_asterisk.sum()))
        summary.append(dict(group=group, **info))
        info['top_shares'] = {}
        for pct in [1, 5, 10, 20]:
            k = (n * pct + 99) // 100
            h = int(ordered.head(k).hours_viewed.sum())
            # Independent Python sort + sum verifies pandas selections.
            check(h == sum(sorted(frame.hours_viewed.tolist(), reverse=True)[:math.ceil(n * pct / 100)]), 'Top sum mismatch')
            boundary = int(ordered.iloc[k - 1].hours_viewed)
            row = dict(group=group, n=n, top_percent=pct, selected_rows=k, selected_actual_percent=100*k/n,
                       top_hours=h, denominator_A_hours=A, denominator_B_hours=B,
                       share_A=h/A, share_B=h/B, cutoff_hours=boundary,
                       tied_rows_at_cutoff=int((frame.hours_viewed == boundary).sum()))
            concentration.append(row)
            info['top_shares'][str(pct)] = row
            if pct in [1, 5, 10]:
                rest = A - h
                tails.append(dict(group=group, n=n, category=f'remainder_after_top_{pct}pct', rows=n-k,
                                  hours=rest, denominator_A_hours=A, denominator_B_hours=B,
                                  share_A=rest/A, share_B=rest/B, share_B_plus_other=(rest+oh)/B))
                check(abs(h/A + rest/A - 1) < 1e-12, 'Complement mismatch')
        kbottom = n // 2
        bottom = sum(sorted(frame.hours_viewed.tolist())[:kbottom])
        bottomrow = dict(group=group, n=n, category='bottom_50pct', rows=kbottom, hours=bottom,
                         denominator_A_hours=A, denominator_B_hours=B, share_A=bottom/A,
                         share_B=bottom/B, share_B_plus_other=None)
        tails.append(bottomrow)
        info['bottom_50pct'] = bottomrow
        info['longtail'] = [dict(row) for row in tails if row['group'] == group]
        x = np.sort(frame.hours_viewed.to_numpy(dtype=np.float64))
        L = np.concatenate(([0.], np.cumsum(x) / A))
        pop = np.arange(n+1) / n
        check(L[0] == 0 and abs(L[-1]-1) < 1e-12 and (np.diff(L) >= 0).all(), 'Lorenz endpoints')
        check((L <= pop + 1e-12).all(), 'Lorenz above equality')
        check(abs((1 - 2 * np.trapezoid(L, pop)) - info['gini_hours']) < 1e-12, 'Independent Gini mismatch')
        lorenz_rows.extend(dict(group=group, n=n, ascending_rows=i, population_share=float(px),
                                cumulative_hours_share_A=float(ly), denominator_A_hours=A)
                           for i, (px, ly) in enumerate(zip(pop, L)))
        info['top10'] = {}
        for metric in ['hours_viewed', 'views']:
            top = ranked(frame, metric).head(10).copy()
            records = []
            for rank, r in enumerate(top.itertuples(), 1):
                rec = dict(group=group, ranked_by=metric, rank=rank, record_id=r.record_id,
                           title=r.title, release_date=None if pd.isna(r.release_date) else r.release_date,
                           source_sheet=r.source_sheet, source_row=r.source_row,
                           hours_viewed=int(r.hours_viewed), views=int(r.views),
                           hours_share_A=r.hours_viewed/A, hours_share_B=r.hours_viewed/B,
                           views_share_disclosed=r.views/info['views_individual'])
                top_rows.append(rec)
                records.append(rec)
            info['top10'][metric] = records
        overlap = set(r['record_id'] for r in info['top10']['hours_viewed']) & set(r['record_id'] for r in info['top10']['views'])
        info['top10_ranking_overlap'] = len(overlap)
        info['top10_hours_share_A'] = sum(r['hours_viewed'] for r in info['top10']['hours_viewed'])/A
        info['top10_hours_share_B'] = sum(r['hours_viewed'] for r in info['top10']['hours_viewed'])/B
        ranks = frame[['record_id', 'title', 'hours_viewed', 'views']].copy()
        ranks['hours_rank_min'] = frame.hours_viewed.rank(method='min', ascending=False).astype(int)
        ranks['views_rank_min'] = frame.views.rank(method='min', ascending=False).astype(int)
        for rec in ranks.to_dict(orient='records'):
            ranking_rows.append(dict(group=group, **rec))
        previous = 0
        for pct, name in [(1, '상위 1%'), (5, '다음 4%'), (10, '다음 5%')]:
            h = info['top_shares'][str(pct)]['top_hours']
            segments.append(dict(group=group, segment=name, hours=h-previous, share_B=(h-previous)/B, denominator_B_hours=B, n=n))
            previous = h
        segments += [dict(group=group, segment='나머지 개별 항목', hours=A-previous, share_B=(A-previous)/B, denominator_B_hours=B, n=n),
                     dict(group=group, segment='Other', hours=oh, share_B=oh/B, denominator_B_hours=B, n=n)]
        stats['groups'][group] = info
    check(gini([1, 1, 1]) == 0 and abs(gini([0, 0, 1])-2/3) < 1e-12, 'Gini sanity check')
    tables = dict(summary=pd.DataFrame(summary), concentration=pd.DataFrame(concentration),
                  longtail=pd.DataFrame(tails), lorenz=pd.DataFrame(lorenz_rows),
                  top10=pd.DataFrame(top_rows), rankings=pd.DataFrame(ranking_rows),
                  composition=pd.DataFrame(segments))
    for name, table in tables.items():
        table.to_csv(TABLES / f'{name}.csv', index=False, encoding='utf-8', float_format='%.15g')
    allinfo = stats['groups']['All']
    total = allinfo['hours_including_other']
    stats['official_total_check'] = dict(official_statement='more than 97 billion hours',
        threshold_hours=97000000000, individual_hours=allinfo['hours_individual'],
        other_hours=allinfo['hours_other'], calculated_hours=total, above_threshold=total>97000000000,
        excess_over_97b_hours=total-97000000000,
        note='97B is a lower-bound narrative, not an exact total; no contradiction. Public values are rounded.')
    (RESULTS / 'statistics.json').write_text(json.dumps(stats, ensure_ascii=False, indent=2, allow_nan=False)+'\n', encoding='utf-8')
    configure_style()
    make_charts(tables, stats)
    for path, sha in hashes.items():
        check(hashlib.sha256((ROOT/path).read_bytes()).hexdigest() == sha, 'Input modified')
    # Every plotted series checked against the persisted source CSV, not rounded labels.
    (TABLES / 'chart_validation.json').write_text(json.dumps(CHART_CHECKS, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    write_reports(stats, tables, hashes)
    print(json.dumps({'groups': {g: {k: v for k, v in stats['groups'][g].items() if k in ['n', 'gini_hours', 'hours_individual', 'hours_including_other', 'hours_mean', 'hours_median', 'top10_ranking_overlap']} for g in GROUPS},
                      'all_top_shares': stats['groups']['All']['top_shares'],
                      'chart_count': 8, 'chart_series_validated': len(CHART_CHECKS)}, ensure_ascii=True))


def configure_style():
    font = Path('C:/Windows/Fonts/malgun.ttf')
    if font.exists():
        font_manager.fontManager.addfont(str(font))
        plt.rcParams['font.family'] = font_manager.FontProperties(fname=str(font)).get_name()
    plt.rcParams.update({'font.size': 14, 'axes.titlesize': 19, 'axes.labelsize': 14,
                         'axes.spines.top': False, 'axes.spines.right': False,
                         'axes.unicode_minus': False, 'svg.fonttype': 'path',
                         'figure.facecolor': '#FAFAFC', 'axes.facecolor': '#FAFAFC',
                         'grid.color': '#DDE1E8', 'grid.alpha': .7})


def figure(title, subtitle, footnote, ncols=1):
    fig, axes = plt.subplots(1, ncols, figsize=(16, 9), squeeze=False)
    fig.subplots_adjust(left=.10, right=.94, bottom=.23, top=.77, wspace=.33)
    fig.text(.055, .93, title, fontsize=25, fontweight='bold', color='#182235')
    fig.text(.055, .865, subtitle, fontsize=14, color='#4B5563')
    fig.text(.055, .055, footnote, fontsize=12, color='#4B5563', linespacing=1.7)
    return fig, axes[0]


def finish(fig, stem):
    fig.savefig(CHARTS/f'{stem}.png', dpi=160)  # 2560 x 1440, exact 16:9
    fig.savefig(CHARTS/f'{stem}.svg')
    plt.close(fig)


def validated_bars(ax, values, source_file, selector, column, scale=1, horizontal=False, **kwargs):
    disk = pd.read_csv(TABLES/source_file)
    expected = selector(disk)[column].to_numpy(dtype=float) / scale
    check(np.allclose(values, expected, rtol=1e-12, atol=1e-12), 'Chart input differs from CSV')
    positions = kwargs.pop('positions', np.arange(len(values)))
    bars = (ax.barh if horizontal else ax.bar)(positions, values, **kwargs)
    actual = np.array([b.get_width() if horizontal else b.get_height() for b in bars])
    check(np.allclose(actual, expected, rtol=1e-12, atol=1e-12), 'Rendered bars differ from CSV')
    CHART_CHECKS.append(dict(source=source_file, column=column, scale=scale, plotted_values=actual.tolist(), passed=True))
    return bars


def make_charts(t, stats):
    s, c, tails = t['summary'], t['concentration'], t['longtail']
    types = s[s.group != 'All']
    note = '2026년 1~6월 | n=영화 9,177 · 시리즈·방송 8,193개 공개 행 | Other 제외\n단위: 왼쪽 십억 시간, 오른쪽 십억 조회 | 합계 비교(점유율 분모 없음); 조회수는 고유 시청자 수가 아님'
    fig, axs = figure('시청 시간과 조회수는 서로 다른 모습을 보여준다', '시리즈·방송은 시즌·방송 행 단위이며 영화 한 편과 같은 단위가 아닙니다.', note, 2)
    for ax, col, title in zip(axs, ['hours_individual', 'views_individual'], ['시청 시간 합계', '공식 조회수 합계']):
        vals = types[col].to_numpy()/1e9
        bars = validated_bars(ax, vals, 'summary.csv', lambda z:z[z.group != 'All'], col, 1e9, color=[COLORS[g] for g in types.group])
        ax.set_xticks([0,1], ['영화', '시리즈·방송'])
        ax.set_title(title, pad=16)
        ax.set_ylabel('십억 시간' if col == 'hours_individual' else '십억 조회')
        ax.set_ylim(0, max(vals)*1.23)
        ax.bar_label(bars, labels=[f'{v:.2f}' for v in vals], padding=8, fontsize=18)
        ax.grid(axis='y'); ax.set_axisbelow(True)
    finish(fig, '01_hours_views_comparison')
    for group, stem in [('Movies','02_movies_top10_hours'), ('Shows','03_shows_top10_hours')]:
        table = t['top10'][(t['top10'].group == group)&(t['top10'].ranked_by == 'hours_viewed')]
        info = stats['groups'][group]
        note = f'2026년 1~6월 | 대상 n={info["n"]:,}, 표시 10행 | 단위: 백만 시간 | Other 제외\n막대: 공식 Hours Viewed(점유율 분모 없음) | 원제·행 ID·공개일: chart_data/top10.csv; 긴 제목은 // 앞 표기 사용'
        fig, axs = figure(f'{LABEL[group]} 시청 시간 TOP 10', f'상위 10행 합계는 개별 항목 시청 시간의 {info["top10_hours_share_A"]:.2%}입니다. (분모 A)', note)
        fig.subplots_adjust(left=.35, top=.78, bottom=.23)
        ax = axs[0]
        vals = table.hours_viewed.to_numpy()/1e6
        selector = lambda z, g=group:z[(z.group==g)&(z.ranked_by=='hours_viewed')]
        bars = validated_bars(ax, vals, 'top10.csv', selector, 'hours_viewed', 1e6, horizontal=True, color=COLORS[group], height=.65)
        labels = ['\n'.join(textwrap.wrap(f'{i}. '+str(title).split(' // ')[0], width=38)) for i,title in zip(table['rank'],table.title)]
        ax.set_yticks(np.arange(10), labels, fontsize=12)
        ax.invert_yaxis(); ax.set_xlim(0,max(vals)*1.17)
        ax.set_xlabel('시청 시간 (백만 시간)'); ax.grid(axis='x'); ax.set_axisbelow(True)
        ax.bar_label(bars, labels=[f'{v:,.1f}' for v in vals], padding=7, fontsize=12)
        finish(fig, stem)
    fig, axs = figure('상위 콘텐츠가 차지하는 시청 시간', 'A. 공개된 개별 항목 사이의 집중도 — Other 제외 분모',
        '2026년 1~6월 | n=영화 9,177 · 시리즈·방송 8,193 · 전체 17,370 | 단위: %\n분모 A=유형별 개별 항목 Hours Viewed 합계; Other 제외 | 상위 행 수=ceil(n×비율); 막대 위 n은 선택 행 수')
    ax=axs[0]
    for i,g in enumerate(GROUPS):
        sub=c[c.group==g]; pos=np.arange(4)+(i-1)*.25
        bars=validated_bars(ax,sub.share_A.to_numpy()*100,'concentration.csv',lambda z,g=g:z[z.group==g], 'share_A',.01,positions=pos,width=.23,color=COLORS[g],label=LABEL[g])
        ax.bar_label(bars,labels=[f'{v:.1%}\nn={k:,}' for v,k in zip(sub.share_A,sub.selected_rows)],padding=6,fontsize=10)
    ax.set_xticks(np.arange(4),['상위 1%','상위 5%','상위 10%','상위 20%'])
    ax.set_ylim(0,112); ax.set_yticks([0,20,40,60,80,100]); ax.yaxis.set_major_formatter(PercentFormatter(100))
    ax.set_ylabel('개별 항목 시청 시간 점유율'); ax.legend(loc='upper left',frameon=False,ncols=3);ax.grid(axis='y');ax.set_axisbelow(True)
    finish(fig,'04_top_shares_A')
    fig,axs=figure('로렌츠 곡선으로 보는 시청 시간의 불균등', '대각선에서 멀어질수록 시청 시간이 일부 항목에 더 집중됩니다.',
        '2026년 1~6월 | n=영화 9,177 · 시리즈·방송 8,193 | 단위: 누적 비율(%)\n분모 A=각 유형 개별 항목 시청 시간 합계; Other 제외 | 가로축: 시청 시간 오름차순 행 누적 비율')
    ax=axs[0]
    for g in GROUPS[:2]:
        sub=t['lorenz'][t['lorenz'].group==g]
        line,=ax.plot(sub.population_share,sub.cumulative_hours_share_A,lw=3,color=COLORS[g],label=LABEL[g])
        disk=pd.read_csv(TABLES/'lorenz.csv');disk=disk[disk.group==g]
        check(np.allclose(line.get_xdata(),disk.population_share) and np.allclose(line.get_ydata(),disk.cumulative_hours_share_A),'Lorenz chart differs')
        CHART_CHECKS.append(dict(chart='05_lorenz_comparison',group=g,source='lorenz.csv',points=len(disk),passed=True))
    ax.plot([0,1],[0,1],ls='--',color='#8993A2',label='완전 균등선')
    ax.set(xlim=(0,1),ylim=(0,1),xlabel='누적 콘텐츠 행 비율',ylabel='누적 시청 시간 점유율')
    ax.xaxis.set_major_formatter(PercentFormatter(1));ax.yaxis.set_major_formatter(PercentFormatter(1));ax.legend(frameon=False);ax.grid()
    finish(fig,'05_lorenz_comparison')
    fig,axs=figure('영화와 시리즈·방송의 지니계수', '0은 균등 분포, 1에 가까울수록 집중된 분포를 뜻합니다.',
        '2026년 1~6월 | n=영화 9,177 · 시리즈·방송 8,193 | 단위: 지니계수(0~1)\n개별 행의 Hours Viewed 분포와 유형별 합계로 계산; Other 제외 | 시즌·방송과 영화의 단위 차이 유의')
    ax=axs[0];vals=types.gini_hours.to_numpy()
    bars=validated_bars(ax,vals,'summary.csv',lambda z:z[z.group!='All'],'gini_hours',color=[COLORS[g] for g in types.group],width=.5)
    ax.set_xticks([0,1],['영화','시리즈·방송']);ax.set_ylim(0,1);ax.set_ylabel('지니계수')
    ax.bar_label(bars,labels=[f'{v:.4f}' for v in vals],padding=8,fontsize=23);ax.grid(axis='y');ax.set_axisbelow(True)
    finish(fig,'06_gini_comparison')
    fig,axs=figure('상위 항목 밖에도 시청은 남아 있다', '상위 1%·5%·10% 제외 후 나머지와 하위 50%의 시청 시간 점유율',
        '2026년 1~6월 | n=영화 9,177 · 시리즈·방송 8,193 · 전체 17,370 | 단위: %\n분모 A=각 그룹 개별 항목 시청 시간 합계; Other 제외 | 하위 50% 행 수=floor(n/2); 네 범주는 서로 겹침')
    ax=axs[0]
    for i,g in enumerate(GROUPS):
        sub=tails[tails.group==g]
        bars=validated_bars(ax,sub.share_A.to_numpy()*100,'longtail.csv',lambda z,g=g:z[z.group==g],'share_A',.01,positions=np.arange(4)+(i-1)*.25,width=.23,color=COLORS[g],label=LABEL[g])
        ax.bar_label(bars,labels=[f'{v:.1%}' for v in sub.share_A],padding=6,fontsize=12)
    ax.set_xticks(np.arange(4),['상위 1% 제외','상위 5% 제외','상위 10% 제외','하위 50%'])
    ax.set_ylim(0,100);ax.yaxis.set_major_formatter(PercentFormatter(100));ax.set_ylabel('개별 항목 시청 시간 점유율');ax.legend(frameon=False,ncols=3);ax.grid(axis='y');ax.set_axisbelow(True)
    finish(fig,'07_longtail_shares_A')
    fig,axs=figure('Other를 포함한 전체 시청 시간의 구성', 'B. 개별 항목의 상위 구간과 Other를 분리해 전체 분모에서 비교합니다.',
        '2026년 1~6월 | n=영화 9,177 · 시리즈·방송 8,193 · 전체 17,370 (+유형별 Other 1행, 전체 2행)\n단위: % | 분모 B=개별 항목+해당 유형 Other 시청 시간; Other는 순위·콘텐츠 수에서 제외 | 구간 행 수는 올림값 차이')
    ax=axs[0];left=np.zeros(3)
    seg_colors=['#B91C3C','#F05C6C','#F9A4A6','#BCCBE0','#525C70']
    for name,color in zip(t['composition'].segment.unique(),seg_colors):
        sub=t['composition'][t['composition'].segment==name]
        values=sub.share_B.to_numpy()*100
        bars=validated_bars(ax,values,'composition.csv',lambda z,name=name:z[z.segment==name],'share_B',.01,horizontal=True,positions=np.arange(3),left=left,color=color,label=name,height=.55)
        for j,(l,v) in enumerate(zip(left,values)):
            if v>5:ax.text(l+v/2,j,f'{v:.1f}%',ha='center',va='center',fontsize=14,color='white' if name=='상위 1%' else '#182235')
        left+=values
    check(np.allclose(left,100),'Composition not 100 percent')
    ax.set_yticks(np.arange(3),[LABEL[g] for g in GROUPS]);ax.invert_yaxis();ax.set_xlim(0,100);ax.xaxis.set_major_formatter(PercentFormatter(100));ax.set_xlabel('Other 포함 전체 시청 시간 점유율');ax.legend(loc='upper center',bbox_to_anchor=(.5,-.12),ncols=5,frameon=False,fontsize=11)
    finish(fig,'08_total_composition_B')


def write_reports(stats,t,hashes):
    info=stats['groups'];a=info['All'];m=info['Movies'];s=info['Shows']
    most='영화' if m['gini_hours']>s['gini_hours'] else '시리즈·방송'
    def pct(g,p,den='A'):
        return info[g]['top_shares'][str(p)]['share_'+den]
    findings=[
        f'공개된 전체 개별 항목에서 상위 1%({a["top_shares"]["1"]["selected_rows"]:,}행)가 시청 시간의 {pct("All",1):.2%}, 상위 10%가 {pct("All",10):.2%}를 차지한다. 균등 분포와 비교하면 소수 항목 집중이 뚜렷하다. 분모는 Other 제외 개별 합계(A)이다.',
        f'{most}의 집중도가 더 높다. 영화 지니계수 {m["gini_hours"]:.4f}, 시리즈·방송 {s["gini_hours"]:.4f}; 상위 10% 점유율은 각각 {pct("Movies",10):.2%}, {pct("Shows",10):.2%}(A)이다. 유형별 관측 행 단위가 달라 동일한 작품 단위의 비교는 아니다.',
        f'평균은 영화 {m["hours_mean"]:,.0f}시간(중앙값 {m["hours_median"]:,.0f}), 시리즈·방송 {s["hours_mean"]:,.0f}시간(중앙값 {s["hours_median"]:,.0f})이다. 평균이 중앙값의 각각 {m["hours_mean"]/m["hours_median"]:.2f}배, {s["hours_mean"]/s["hours_median"]:.2f}배로, 평균 하나로 전형적인 항목을 설명하기 어렵다.',
        f'상위 10%를 제외한 나머지 90%의 전체 개별 시청 시간 점유율은 {1-pct("All",10):.2%}, 하위 50%는 {a["bottom_50pct"]["share_A"]:.2%}이다(A). 롱테일 소비는 존재하지만 개별 행에 균등하게 배분되지는 않는다.',
        f'인기 순위는 지표에 따라 다르다. 시청 시간·조회수 TOP 10의 겹치는 항목은 영화 {m["top10_ranking_overlap"]}개, 시리즈·방송 {s["top10_ranking_overlap"]}개다. 시청 시간은 소비 시간 총량, 공식 Views는 길이를 반영한 지표이며 고유 시청자 수가 아니다.'
    ]
    (ROOT/'reports/key_findings.md').write_text('# 중요한 발견 5가지\n\n'+'\n\n'.join(f'{i}. {v}' for i,v in enumerate(findings,1))+'\n\n분모 A=공개 개별 항목의 시청 시간 합계. 전체 카탈로그의 모든 콘텐츠가 개별 공개된 것은 아니다. 2026년 상반기 단면 관측이며 인과·장기 추세를 뜻하지 않는다.\n',encoding='utf-8')
    lines=['# 넷플릭스의 인기는 소수의 콘텐츠에 집중될까?', '',
           f'공개된 개별 콘텐츠 행 기준으로 시청 시간은 소수 항목에 집중된다. 전체 상위 1%가 {pct("All",1):.2%}, 상위 10%가 {pct("All",10):.2%}를 차지하며 지니계수는 {a["gini_hours"]:.4f}이다. 이는 원인 추정이 아닌 2026년 상반기 공개 분포의 기술적 설명이다.', '',
           '## 대상과 분모', '',
           '- 개별 17,370행: Movies 9,177행, Shows 8,193행. Other 2행은 개별 순위와 콘텐츠 수에서 제외.',
           '- Shows는 시즌·리미티드 시리즈·개별 방송 등이 섞인 보고서 행 단위. 시리즈 전체 또는 영화 한 편과 같은 단위라고 주장하지 않음.',
           '- A(개별 항목 사이 점유율): 해당 그룹의 개별 항목 시청 시간 합계. 지니계수·로렌츠 곡선도 이 분포만 사용.',
           '- B(전체 시청 시간 대비 비중): A + 해당 그룹 Other 시청 시간. Other Movies/Other Shows의 귀속은 원본 시트·제목·각주로 확인.',
           '- Runtime 별표 310행, 공개일 결측 11,324행 모두 유지. 공식 Hours Viewed와 Views 사용; 조회수 재계산·보간 없음.', '',
           '## 기초 통계', '',
           '| 유형 | 행 수 | 시청 시간 A | Other 시청 시간 | 시청 시간 B | 공개 항목 조회수 | 평균 시간 | 중앙값 시간 | 최대 시간 |',
           '|---|---:|---:|---:|---:|---:|---:|---:|---:|']
    for g in GROUPS:
        z=info[g]
        lines.append(f'| {g} | {z["n"]:,} | {z["hours_individual"]:,} | {z["hours_other"]:,} | {z["hours_including_other"]:,} | {z["views_individual"]:,} | {z["hours_mean"]:,.2f} | {z["hours_median"]:,.0f} | {z["hours_max"]:,} |')
    lines+=['', '- 시간 단위는 시간, 조회수는 공식 Views의 합계. 평균·중앙값·최대는 Other 제외 개별 행 기준. Other 조회수는 미공개이므로 B에 대응하는 전체 조회수는 계산 불가.',
            '- 공식 Views 합계는 고유 시청자·계정 합계가 아니다. 서로 다른 항목의 반복 시청이나 동일 이용자 중복을 제거한 수치도 아니다.', '',
            '## 시청 시간·조회수 상위 10개', '', '- 순위는 해당 지표 내림차순. 공개값 동률이면 원본 시트명·행 번호 순으로 정확히 10행 선택. 실제 인기의 미세한 우열을 뜻하지 않음.']
    for g in GROUPS[:2]:
        for metric,name in [('hours_viewed','시청 시간'),('views','조회수')]:
            top=t['top10'][(t['top10'].group==g)&(t['top10'].ranked_by==metric)].copy()
            top['시청 시간 점유율 A(%)']=top.hours_share_A*100
            lines+=['',f'### {g}: {name} TOP 10','',md_table(top,['rank','record_id','title','release_date','hours_viewed','views','시청 시간 점유율 A(%)'])]
    lines+=['', '## 집중도: 상위 비율·지니계수', '',
            '- 상위 행 수 k=ceil(n×p/100). 시청 시간 내림차순 정렬 후 k행 합계 H_top 사용. A=H_top/개별 시청 시간 합계, B=H_top/(개별+Other 시청 시간).', '',
            '| 그룹 | 상위 비율 | 선택 행 수 | 시청 시간 합계 | 점유율 A | 점유율 B |', '|---|---:|---:|---:|---:|---:|']
    for r in t['concentration'].itertuples():
        lines.append(f'| {r.group} | {r.top_percent}% | {r.selected_rows:,} | {r.top_hours:,} | {r.share_A:.4%} | {r.share_B:.4%} |')
    lines+=['', '| 그룹 | 지니계수(Other 제외) | 상위 10행 시간 점유율 A | 상위 10행 시간 점유율 B |', '|---|---:|---:|---:|']
    for g in GROUPS:
        z=info[g];lines.append(f'| {g} | {z["gini_hours"]:.6f} | {z["top10_hours_share_A"]:.4%} | {z["top10_hours_share_B"]:.4%} |')
    lines+=['', findings[1], '',
            f'영화의 상위 1행 {m["top10"]["hours_viewed"][0]["title"]}은 유형별 개별 시청 시간의 {m["top10"]["hours_viewed"][0]["hours_share_A"]:.2%}, 시리즈·방송의 {s["top10"]["hours_viewed"][0]["title"]}은 {s["top10"]["hours_viewed"][0]["hours_share_A"]:.2%}이다. 특정 한 작품만으로 전체 소비를 설명할 수는 없지만 여러 상위 항목을 묶으면 집중도가 커진다.', '',
            '합친 전체 분포는 유형 내부 집중도와 유형 사이 시청 시간 규모 차이를 함께 반영한다. 전체 지니계수가 두 유형 계수의 단순 평균일 필요는 없다.', '',
            '## 평균·중앙값과 지표별 순위 차이', '', findings[2], '',findings[4], '']
    for g in GROUPS[:2]:
        z=info[g]
        lines.append(f'- {g}: 시청 시간 1위 **{z["top10"]["hours_viewed"][0]["title"]}**, 공식 조회수 1위 **{z["top10"]["views"][0]["title"]}**. 공식 조회수는 시청 시간 총량과 다른 기준이다. 긴 항목은 같은 조회수에서도 시청 시간이 커질 수 있으나, 이번 비교만으로 길이가 순위 차이의 유일한 원인임을 입증하지는 않는다.')
    lines+=['', '## 롱테일', '',
            '- 상위 1%·5%·10% 제외 시 남은 개별 행의 비중을 계산. B에서는 Other를 제외한 나머지 개별 항목과 Other를 더한 나머지를 분리. 하위 50%는 floor(n/2)행(오름차순).', '',
            '| 그룹 | 범주 | 남은/선택 행 수 | 점유율 A | 개별 나머지 점유율 B | Other까지 포함한 나머지 B |', '|---|---|---:|---:|---:|---:|']
    for r in t['longtail'].itertuples():
        incl='해당 없음' if pd.isna(r.share_B_plus_other) else f'{r.share_B_plus_other:.4%}'
        lines.append(f'| {r.group} | {r.category} | {r.rows:,} | {r.share_A:.4%} | {r.share_B:.4%} | {incl} |')
    lines+=['', findings[3], '',
            '하위 50%는 공개 항목만을 정렬한 결과이다. Other 내부 항목을 풀어 전체 카탈로그의 하위 절반을 구성할 수 없고, 하위 50%에 Other를 임의 배분하지 않는다.', '',
            '## 공식 발표와 합계 대조', '',
            f'- 개별 시청 시간 {a["hours_individual"]:,}시간 + Other {a["hours_other"]:,}시간 = **{a["hours_including_other"]:,}시간(976.645억 시간)**.',
            f'- [공식 발표]({PAGE})의 970억 시간 초과라는 설명과 일치한다. 970억보다 {a["hours_including_other"]-97000000000:,}시간 많지만, 발표는 정확히 970억이라고 한 것이 아니므로 불일치가 아니다.',
            '- 합계는 공개 반올림 값의 합산이며 내부 비반올림 총계와 정확히 동일하다고 보장할 수 없다.', '',
            '## 한계와 다음 단계', '',
            '- 공개 반올림 값의 동률과 임계값 주변 순위는 정밀한 실제 순위와 다를 수 있음. 선택 행 수·실제 선택 비율·경계 동률 수는 concentration.csv에 보존.',
            '- Other 내부의 제목 수·개별 분포·Views를 알 수 없어 전체 카탈로그의 지니계수와 전체 조회수 집중도를 산출하지 않음.',
            '- Runtime 별표의 정확한 의미 미확인. 공식 시청 시간·조회수 자체를 사용했으므로 해당 행을 제외할 근거로 삼지 않음.',
            '- 본 결과는 2026년 상반기 공개 보고서 항목의 기술통계. 추천 알고리즘의 효과, 장기 추세, 국가·언어·장르 원인을 입증하지 않음. 표본 추론이나 유의성 검정을 시행하지 않음.',
            '- 4단계 슬라이드 제작 진행 가능. A/B 분모와 시즌·방송 단위를 유지해 설명해야 함. 이번 단계에서는 슬라이드·웹·배포 작업 없음.', '',
            '## 결과 파일', '', '- 주요 수치: results/statistics.json', '- 그래프: results/charts/ (PNG 2560×1440 및 SVG)',
            '- 원본 계산값: results/chart_data/summary.csv, concentration.csv, longtail.csv, lorenz.csv, top10.csv, rankings.csv, composition.csv',
            '- 계산·그래프 검증: reports/analysis_validation.md 및 results/chart_data/chart_validation.json',
            '- 재현: `.venv/Scripts/python scripts/analyze_data.py`']
    (ROOT/'reports/analysis.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    charts=[('01_hours_views_comparison','영화·시리즈 시청 시간 및 조회수 비교','summary.csv'),
            ('02_movies_top10_hours','영화 시청 시간 TOP 10','top10.csv'),
            ('03_shows_top10_hours','시리즈·방송 시청 시간 TOP 10','top10.csv'),
            ('04_top_shares_A','상위 1%·5%·10%·20% 점유율 A','concentration.csv'),
            ('05_lorenz_comparison','영화·시리즈 로렌츠 곡선','lorenz.csv'),
            ('06_gini_comparison','영화·시리즈 지니계수','summary.csv'),
            ('07_longtail_shares_A','롱테일·하위 50% 점유율 A','longtail.csv'),
            ('08_total_composition_B','Other 포함 전체 시청 시간 구성 B','composition.csv')]
    v=['# 분석 계산식과 검증 결과', '', '## 입력 무결성', '',
       '- 원본 Excel과 정제 CSV의 모든 17,372개 ID·제목·유형·원본 행 번호·Hours Viewed·Views를 행별 대조: 통과.',
       '- 정제 데이터 17,370행, Other 2행. 공개일 결측 11,324행 및 Runtime 별표 310행 유지: 통과.',
       '- 주요 개별 수치 결측·음수·0 없음. Other 조회수 미공개 결측 보존: 통과.',
       '- Other 귀속은 원본 시트 및 Other Movies/Other Shows 제목으로 확정; 원본 각주로 시청 시간 집계 의미 확인. 임의 귀속 없음.', '',
       '입력 SHA-256:', '']+[f'- `{p}`: `{sha}`' for p,sha in hashes.items()]
    v+=['', '## 계산식 및 독립 검증', '',
        '- 상위 p%: 정수 연산 k=(n×p+99)//100로 올림. pandas 내림차순 상위 k행 합계를 독립 Python sorted+sum과 비교: 전체 12개 조합 통과.',
        '- 공개 수치 동률: 원본 시트명·행 번호 순으로 k행만 선택. 같은 시간 동률 선택은 시간 합계를 바꾸지 않음. 경계 동률 개수 저장.',
        '- A=Σ개별 시청 시간. B=A+Σ동일 유형 Other. 상위 점유율은 동일 분자 H_top을 각 분모로 나눔.',
        '- 하위 50%: floor(n/2)행. 홀수 n의 정확한 절반을 만들기 위한 분할 관측치는 생성하지 않음.',
        '- 지니계수: x_i는 시청 시간 오름차순, G=Σ(2i−n−1)x_i / (nΣx_i), i=1..n. 유한 모집단 보정 n/(n−1)은 사용하지 않음.',
        '- 지니 분자는 Python 정수, 비율은 Fraction으로 계산 후 float 저장. 독립 로렌츠 사다리꼴 면적식 G=1−2∫L(p)dp와 1e−12 이내 일치: 세 그룹 통과.',
        '- 지니 검산: [1,1,1]→0; [0,0,1]→2/3. 로렌츠 (0,0)/(1,1), 단조 증가, 균등선 이하: 통과.',
        '- 롱테일 A=1−상위 점유율 A; B의 나머지 개별 + Other + 상위 =100%. 100% 누적 막대 합계 검사: 통과.',
        '- 평균·중앙값·최대·행 수는 개별 항목만 사용. 조회수는 원본 공식 Views와 직접 대조; 재계산 조회수 컬럼 미사용.',
        '- 공식 발표 970억 초과 대조: 합계 97,664,500,000시간으로 조건 충족. 정확히 970억과의 동등 비교는 부적절.', '',
        '## 그래프 검증', '',
        '- 16:9 PNG(2560×1440), 확장 가능한 SVG. 한국어 글꼴 사용. 모든 그래프에 기간·표본 수·단위·분모·Other 포함 여부 명시.',
        f'- 막대의 실제 높이/너비와 로렌츠 선 좌표를 저장된 CSV에서 다시 읽어 대조: {len(CHART_CHECKS)}개 시리즈 통과. 허용오차 rtol=1e−12(막대), 부동소수점 CSV 직렬화 오차만 허용.',
        '- 그래프 표시값은 가독성을 위해 소수 1~4자리 반올림; 원본 CSV·JSON은 계산 정밀도 유지. 누적 구성의 5% 이하 구간(Other 등)은 숫자 생략하되 범례·원본값 보존.', '',
        '| 파일명(PNG/SVG) | 그래프 | 원본 계산 CSV |', '|---|---|---|']
    v += [f'| {stem} | {title} | {source} |' for stem,title,source in charts]
    v += ['', '## 분석상의 한계', '',
          '- 반올림 공개 데이터에 대한 기술통계이며 내부 실제 시청량의 정밀 추정·신뢰구간을 제공하지 않음.',
          '- Other 내부 행 개수와 분포 불명: 전체 카탈로그 Gini/Lorenz와 전체 조회수 합계 계산 보류. B의 시간 합계와 점유율 계산은 가능.',
          '- Runtime 별표 의미 미확인 유지. 이를 분석 제외 또는 임의 길이·조회수 보정 근거로 사용하지 않음.',
          '- Shows의 시즌·방송 행 단위와 영화 단위가 다름. 전체 합산 계수는 그 혼합 분포의 계수.',
          '- 검증 중 발견된 미해결 계산 오류 없음. 장기 변화·추천 시스템의 인과 효과는 분석 범위 밖.', '',
          '## 실행 환경', '', f'- pandas {pd.__version__}; numpy {np.__version__}; openpyxl {openpyxl.__version__}; matplotlib {matplotlib.__version__}.',
          '- 입력 파일 해시 실행 전후 동일. 모든 검증 실패 시 코드가 ValueError를 발생시킴.',
          f'- 공식 발표: {PAGE} (2026-10-20 확인).']
    (ROOT/'reports/analysis_validation.md').write_text('\n'.join(v)+'\n',encoding='utf-8')


if __name__=='__main__':
    main()
