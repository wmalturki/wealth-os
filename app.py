import io
from pathlib import Path
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title='WEALTH OS | Financial Command Center', page_icon='◆', layout='wide', initial_sidebar_state='expanded')
st.markdown('''<style>
.stApp {background:#0B111B;color:#E8EDF6} [data-testid="stSidebar"]{background:#111C2B}
[data-testid="stMetric"]{background:#142235;border:1px solid #263C53;border-radius:15px;padding:17px}
[data-testid="stMetricLabel"]{color:#A6B8C9} h1,h2,h3{letter-spacing:-.03em}
.stButton>button,.stDownloadButton>button{background:#C6A46A;color:#101A26;border:0;border-radius:10px;font-weight:700}
</style>''',unsafe_allow_html=True)
BASE=Path(__file__).parent/'financial_goals_tracker_2028_2040.xlsx'
DEFAULT=BASE.read_bytes()
# The workbook is the source of truth. Map UI fields to actual Inputs row labels,
# rather than assuming the dashboard and Excel have identical row numbers.
FIELD_NAMES={
2:'Starting cash Jan 2028',
3:'Jan 2028 bonus',
4:'2028 pre-switch monthly net',
5:'New gross salary July 2028',
6:'Net factor',
7:'Annual salary raise',
8:'Promotion year',
9:'Promotion raise',
10:'Bonus in months',
11:'First annual bonus year',
12:'Living expenses pre-marriage',
13:'Living expenses after marriage',
14:'Marriage year',
15:'Marriage cost',
16:'Car 1 price',
17:'Car 1 down payment',
18:'Car 1 balloon percentage',
19:'Car 1 financing annual rate',
20:'Car 1 term in months',
21:'Car 1 purchase year',
22:'Car 1 sale year',
23:'Car 1 resale value',
24:'Car 2 purchase year',
25:'Car 2 cash price',
26:'Villa year',
27:'Villa price',
28:'Villa down payment',
29:'Villa fees and fitout',
30:'Mortgage rate',
31:'Mortgage years',
32:'Emergency reserve',
33:'Annual portfolio return',
34:'Portfolio seed after villa',
35:'Monthly investing after villa',
36:'Bonus invested after villa',
37:'Withdrawal per year'}

LABELS={i:name for i,name in FIELD_NAMES.items()}
PCT={6,7,9,18,19,30,33,36}
YEARS={8,11,14,20,21,22,24,26,31}

def workbook_rows(w):
    ws=w['Inputs']
    mapping={str(ws.cell(row,1).value).strip():row for row in range(2,ws.max_row+1)
             if ws.cell(row,1).value is not None}
    missing=set(FIELD_NAMES.values())-set(mapping)
    if missing:raise ValueError('Missing required Excel inputs: '+', '.join(sorted(missing)))
    return mapping

def read_inputs(data):
    w=load_workbook(io.BytesIO(data),read_only=False,data_only=False)
    mapping=workbook_rows(w)
    d={i:float(w['Inputs'].cell(mapping[name],2).value or 0) for i,name in FIELD_NAMES.items()}
    w.close()
    return d

def amort_payment(principal,annual,months,balloon=0):
 if months<=0:return 0
 r=annual/12
 return (principal-balloon/(1+r)**months)*r/(1-(1+r)**(-months)) if r else (principal-balloon)/months

def balance(principal,annual,payment,n):
 r=annual/12
 return principal*(1+r)**n-payment*((1+r)**n-1)/r if r else principal-payment*n

EARLY_DEFAULTS={'opening_cash':0.0,'opening_portfolio':0.0,'salary_2026':12500.0,'salary_2027':12500.0,'living':6500.0,'tuition':3534.0,'visa':1800.0,'bonus_2027':0.0,'invest_early':1000.0,'invest_pre_villa':3000.0,'cash_reserve':10000.0}
EARLY_LABELS={'opening_cash':'Cash at October 2026','opening_portfolio':'Portfolio at October 2026','salary_2026':'Net salary Oct–Dec 2026','salary_2027':'Net salary 2027','living':'Monthly living expenses (estimate)','tuition':'Monthly university fees (estimate)','visa':'Monthly credit card payment (estimate)','bonus_2027':'January 2027 bonus','invest_early':'Monthly investment 2026–27 (estimate)','invest_pre_villa':'Monthly investment 2028–2035 (estimate)','cash_reserve':'Minimum cash reserve before investing'}
def read_early(data):
 out=dict(EARLY_DEFAULTS)
 w=load_workbook(io.BytesIO(data),read_only=False,data_only=True)
 if 'Early Years' in w:
  for key,val in w['Early Years'].iter_rows(min_row=2,max_col=2,values_only=True):
   if key in out and val is not None:out[key]=float(val)
 w.close()
 return out
def simulate(d,early=None):
 early=early or EARLY_DEFAULTS
 p=d[16]-d[17]; balloon=d[16]*d[18]; term=int(d[20]); carpay=amort_payment(p,d[19],term,balloon)
 mortgage=amort_payment(max(0,d[27]-d[28]),d[30],int(d[31])*12)
 cash=early['opening_cash']; portfolio=early['opening_portfolio']; rows=[]; lowest=1e18; monthly=[]
 for year in [2026,2027]:
  income=spending=bonus_total=0
  for month in range(10 if year==2026 else 1,13):
   net=early['salary_2026'] if year==2026 else early['salary_2027']
   bonus=early['bonus_2027'] if (year,month)==(2027,1) else 0
   costs=early['living']+early['tuition']+(early['visa'] if (year,month)<=(2027,1) else 0)
   flow=net+bonus-costs
   invest=min(max(0,cash+flow-early['cash_reserve']),max(0,early['invest_early']))
   flow-=invest;portfolio+=invest
   portfolio*=((1+d[33])**(1/12))
   cash+=flow;income+=net+bonus;bonus_total+=bonus;spending+=costs
   lowest=min(lowest,cash)
   monthly.append({'year':year,'month':month,'cash':cash,'portfolio':portfolio,'net_cash_flow':flow})
  rows.append({'السنة':year,'الدخل':income,'البونص':bonus_total,'المعيشة والأقساط':spending,'أحداث مالية صافية':0,'استثمار جديد':0,'السيولة':cash,'المحفظة':portfolio,'الثروة السائلة':cash+portfolio})
 cash+=d[2] # Optional additional cash injected at Jan 2028; early-year savings carry forward.
 purchase1=(int(d[21]),7); sale1=(int(d[22]),7); purchase2=(int(d[24]),7); marriage=(int(d[14]),1); villa=(int(d[26]),12)
 car_months=0; loan_balance=p; mortgage_started=False
 for year in range(2028,2041):
  income=spending=goalcost=bonus_total=investment_total=0
  for month in range(1,13):
   if year==2028: gross=d[5]; net=d[4] if month<=6 else gross*d[6]
   else: gross=d[5]*(1+d[7])**(year-2028)*(1+d[9] if year>=int(d[8]) else 1);net=gross*d[6]
   bonus=(d[3] if (year,month)==(2028,1) else gross*d[10]*d[6] if year>=int(d[11]) and month==12 else 0)
   living=d[12] if (year,month)<marriage else d[13]
   flow=net+bonus-living; income+=net+bonus; bonus_total+=bonus; spending+=living
   if (year,month)==purchase1:flow-=d[17];goalcost+=d[17]
   if purchase1<=(year,month)<sale1 and car_months<term:
    flow-=carpay; spending+=carpay; car_months+=1
   if (year,month)==sale1:
    payoff=max(0,balance(p,d[19],carpay,car_months)); flow+=d[23]-payoff;goalcost+=payoff-d[23]
   if (year,month)==purchase2:flow-=d[25];goalcost+=d[25]
   if (year,month)==marriage:flow-=d[15];goalcost+=d[15]
   if (year,month)==villa:
    flow-=d[28]+d[29];goalcost+=d[28]+d[29];mortgage_started=True
    seed=min(max(0,cash+flow-d[32]),d[34]);flow-=seed;portfolio+=seed;investment_total+=seed
   elif mortgage_started:
    flow-=mortgage;spending+=mortgage
    invest=min(max(0,cash+flow-d[32]),d[35]+(bonus*d[36] if bonus else 0))
    flow-=invest;portfolio+=invest;investment_total+=invest
   if not mortgage_started:
    invest=min(max(0,cash+flow-early['cash_reserve']),max(0,early['invest_pre_villa']))
    flow-=invest;portfolio+=invest;investment_total+=invest
   # Draw from invested assets only when a major planned outflow would otherwise
   # create negative cash; the remaining shortfall is explicitly shown as a deficit.
   if cash+flow<0 and portfolio>0:
    withdrawal=min(portfolio,-(cash+flow))
    portfolio-=withdrawal;flow+=withdrawal
   portfolio*=((1+d[33])**(1/12))
    if month==12 and year>int(d[26]):portfolio-=min(max(0,portfolio),d[37])
   cash+=flow;lowest=min(lowest,cash)
   monthly.append({'year':year,'month':month,'cash':cash,'portfolio':portfolio,'net_cash_flow':flow})
  rows.append({'السنة':year,'الدخل':income,'البونص':bonus_total,'المعيشة والأقساط':spending,'أحداث مالية صافية':goalcost,'استثمار جديد':investment_total,'السيولة':cash,'المحفظة':portfolio,'الثروة السائلة':cash+portfolio})
 return pd.DataFrame(rows),pd.DataFrame(monthly),carpay,mortgage,lowest

st.markdown('## ◆ WEALTH OS')
st.caption('PERSONAL FINANCIAL COMMAND CENTER  ·  OCT 2026—2040  ·  SAR')
upload=st.sidebar.file_uploader('📂 ارفع ملف الإكسل المحدّث',type=['xlsx'])
source=upload.getvalue() if upload else DEFAULT
try: values=read_inputs(source); early_values=read_early(source)
except Exception as e:st.error(f'تعذّر قراءة الملف: {e}');st.stop()
st.sidebar.markdown('### ⚙️ محاكي الافتراضات')
if 'last_file' not in st.session_state or st.session_state.last_file!=hash(source):
 st.session_state.last_file=hash(source)
 for i,v in values.items():st.session_state[f'input_{i}']=float(v)
 for k,v in early_values.items():st.session_state[f'early_{k}']=float(v)
with st.sidebar.expander('🗓️ 2026–2027 | مرحلة التأسيس',expanded=True):
 for k,label in EARLY_LABELS.items():st.number_input(label,key=f'early_{k}',step=100.0,format='%.0f')
 st.caption('الاستثمار الشهري تقديري، ويتوقف إذا كانت السيولة أقل من الاحتياطي. العائد السنوي افتراض وليس ضمانًا. بونص يناير 2027 يبدأ من صفر.')
with st.sidebar.expander('الدخل والبونص',expanded=False):
 for i in [5,6,7,8,9,10,11,4,3]:
  step=.01 if i in PCT else 1.0 if i in YEARS else 500.0
  st.number_input(LABELS[i],key=f'input_{i}',step=step,format='%.3f' if i in PCT else '%.0f')
with st.sidebar.expander('السيارات والزواج'):
 for i in [12,13,14,15,16,17,18,19,20,21,22,23,24,25]:
  st.number_input(LABELS[i],key=f'input_{i}',step=.01 if i in PCT else 1.0 if i in YEARS else 1000.0,format='%.3f' if i in PCT else '%.0f')
with st.sidebar.expander('الفيلا والمحفظة'):
 for i in [26,27,28,29,30,31,32,33,34,35,36,37,2]:
  st.number_input(LABELS[i],key=f'input_{i}',step=.01 if i in PCT else 1.0 if i in YEARS else 1000.0,format='%.3f' if i in PCT else '%.0f')
d={i:float(st.session_state[f'input_{i}']) for i in LABELS}
early={k:float(st.session_state[f'early_{k}']) for k in EARLY_DEFAULTS}
try:annual,monthly,carpay,mortgage,lowest=simulate(d,early)
except Exception as e:st.error(f'تحقق من الافتراضات: {e}');st.stop()

def sar(x):return f'{x:,.0f} ر.س'
selected=st.select_slider('السنة المعروضة',options=list(range(2026,2041)),value=2035)
r=annual[annual['السنة']==selected].iloc[0]
a,b,c,e=st.columns(4)
a.metric('الثروة السائلة المتوقعة',sar(r['الثروة السائلة']))
b.metric('السيولة النقدية',sar(r['السيولة']))
c.metric('المحفظة الاستثمارية',sar(r['المحفظة']))
e.metric('الدخل السنوي',sar(r['الدخل']))
if lowest<0:st.error(f'⚠️ الخطة تُظهر عجزًا نقديًا مؤقتًا يصل إلى {sar(-lowest)}. يلزم تعديل توقيت الأهداف أو التمويل.')
else:st.success('✓ لا يظهر عجز نقدي في التدفقات الشهرية وفق الافتراضات المدخلة. هذا لا يغني عن اختبار المخاطر.')
t1,t2,t3,t4,t5,t6=st.tabs(['📊 نظرة عامة','🎯 الأهداف','💸 التدفقات النقدية','🧪 اختبار السيناريوهات','📅 Actual vs Forecast','💎 Net Worth & Portfolio'])
with t1:
 fig=go.Figure()
 for col,color in [('السيولة','#4DB9B4'),('المحفظة','#D3AD71'),('الثروة السائلة','#8D9FF2')]:
  fig.add_trace(go.Scatter(x=annual['السنة'],y=annual[col],mode='lines+markers',name=col,line=dict(color=color,width=3)))
 fig.update_layout(template='plotly_dark',paper_bgcolor='#0B111B',plot_bgcolor='#0B111B',title='تطور السيولة والمحفظة والثروة السائلة',height=420,legend=dict(orientation='h'),yaxis_title='ريال')
 st.plotly_chart(fig,use_container_width=True)
 c1,c2=st.columns(2);c1.metric('قسط سيارة 2028 التقديري',sar(carpay));c2.metric('القسط العقاري التقديري',sar(mortgage))
 st.caption('الثروة السائلة = النقد + المحفظة فقط. لا تشمل القيمة السوقية للفيلا والسيارات أو رصيد التمويل العقاري.')
with t2:
 goals=[('🚘 السيارة الأولى',int(d[21]),d[16]),('💍 الزواج',int(d[14]),d[15]),('🚘 السيارة الثانية',int(d[24]),d[25]),('🏡 الفيلا',int(d[26]),d[27]),('📈 محفظة مليون',2038,1000000)]
 for title,yr,amt in goals:
  st.markdown(f'**{title}**  ·  {yr}  ·  {sar(amt)}')
  if title.startswith('📈'):
   actual=float(annual.loc[annual['السنة']==yr,'المحفظة'].iloc[0]);st.progress(min(1,max(0,actual/amt)));st.caption(f'المحفظة المتوقعة: {sar(actual)}')
  else:st.caption('موعد وتكلفة مخططان — تحقق من الرصيد المتاح وقت التنفيذ')
with t3:
 st.markdown('#### سجل الأرقام الفعلية من Excel — Update Log')
 try:
  histwb=load_workbook(io.BytesIO(source),read_only=True,data_only=True)
  histws=histwb['Update Log']
  hist=list(histws.iter_rows(min_row=4,max_col=7,values_only=True))
  hist=[r for r in hist if r[0] is not None]
  if hist:
   st.dataframe(pd.DataFrame(hist,columns=['التاريخ','صافي الراتب الفعلي','البونص الفعلي','الكاش الفعلي','الاستثمارات الفعلية','الديون','ملاحظات']),hide_index=True,use_container_width=True)
  else:st.info('أضف أرقامك الفعلية في تبويب Update Log في الإكسل ثم ارفعه.')
  histwb.close()
 except Exception as exc:st.warning(f'تعذر قراءة Update Log: {exc}')
 fig2=go.Figure()
 fig2.add_bar(x=annual['السنة'],y=annual['الدخل'],name='الدخل',marker_color='#4DB9B4')
 fig2.add_bar(x=annual['السنة'],y=-annual['المعيشة والأقساط'],name='المعيشة والأقساط',marker_color='#DB7B76')
 fig2.update_layout(template='plotly_dark',paper_bgcolor='#0B111B',plot_bgcolor='#0B111B',barmode='relative',height=370)
 st.plotly_chart(fig2,use_container_width=True)
 st.dataframe(annual.style.format({col:'{:,.0f}' for col in annual.columns if col!='السنة'}),hide_index=True,use_container_width=True)
with t4:
 st.markdown('#### أثر العائد الاستثماري على محفظتك')
 comparison=[]
 for rate in [.05,.08,.10]:
  scenario=dict(d);scenario[33]=rate;out,_,_,_,_=simulate(scenario,early)
  comparison.append({'العائد المفترض':f'{rate:.0%}','محفظة 2038':out.loc[out['السنة']==2038,'المحفظة'].iloc[0],'محفظة 2040':out.loc[out['السنة']==2040,'المحفظة'].iloc[0]})
 st.dataframe(pd.DataFrame(comparison).style.format({'محفظة 2038':'{:,.0f}','محفظة 2040':'{:,.0f}'}),hide_index=True,use_container_width=True)
 st.warning('العوائد افتراضية وليست مضمونة. المحاكاة لا تشمل التضخم، الرسوم، ضرائب محتملة، الأطفال، أو هبوط الأسواق المفاجئ.')


# Monthly actuals come from the workbook's Update Log. No personal data is persisted on the server.
def actual_history(raw):
    book=load_workbook(io.BytesIO(raw),read_only=False,data_only=True)
    try:
        sheet=book['Update Log']
        records=[]
        for vals in sheet.iter_rows(min_row=4,max_col=7,values_only=True):
            if vals[0] is None:continue
            dt=pd.to_datetime(vals[0],errors='coerce')
            if pd.isna(dt):continue
            def num(v):
                try:return float(v) if v is not None and v!='' else None
                except (ValueError,TypeError):return None
            records.append({'date':dt,'salary':num(vals[1]),'bonus':num(vals[2]),
                            'cash':num(vals[3]),'portfolio':num(vals[4]),
                            'debt':num(vals[5]),'notes':str(vals[6] or '')})
        if not records:return pd.DataFrame()
        result=pd.DataFrame(records).sort_values('date').drop_duplicates('date',keep='last')
        result['period']=result['date'].dt.to_period('M')
        return result.drop_duplicates('period',keep='last').reset_index(drop=True)
    finally:book.close()

try:
    actuals=actual_history(source)
    actual_error=None
except Exception as err:
    actuals=pd.DataFrame()
    actual_error=str(err)

with t5:
    st.subheader('Actual vs Forecast | الفعلي مقابل الخطة')
    st.caption('أدخل بيانات كل شهر في Update Log داخل ملف Excel، ثم ارفع الملف مجددًا. المقارنة تبدأ من أكتوبر 2026؛ مصاريف المرحلة الأولى تقديرية.')
    if actual_error:st.warning(f'تعذر قراءة سجل الأرقام الفعلية: {actual_error}')
    elif actuals.empty:st.info('ما فيه بيانات شهرية فعلية حتى الآن. افتح تبويب Update Log في الإكسل، وأدخل التاريخ والكاش والمحفظة والديون ثم ارفعه.')
    else:
        actuals['forecast_cash']=actuals.apply(lambda z: monthly.loc[(monthly['year']==z['period'].year)&(monthly['month']==z['period'].month),'cash'].iloc[0] if ((monthly['year']==z['period'].year)&(monthly['month']==z['period'].month)).any() else float('nan'),axis=1)
        actuals['forecast_portfolio']=actuals.apply(lambda z: monthly.loc[(monthly['year']==z['period'].year)&(monthly['month']==z['period'].month),'portfolio'].iloc[0] if ((monthly['year']==z['period'].year)&(monthly['month']==z['period'].month)).any() else float('nan'),axis=1)
        actuals['actual_liquid']=actuals['cash']+actuals['portfolio']
        actuals['forecast_liquid']=actuals['forecast_cash']+actuals['forecast_portfolio']
        valid=actuals.dropna(subset=['actual_liquid','forecast_liquid'])
        if not valid.empty:
            last=valid.iloc[-1]
            x,y,z=st.columns(3)
            x.metric('الثروة السائلة الفعلية',sar(last['actual_liquid']))
            y.metric('المخطط لنفس الشهر',sar(last['forecast_liquid']))
            z.metric('الفرق عن الخطة',sar(last['actual_liquid']-last['forecast_liquid']))
            graph=go.Figure()
            graph.add_trace(go.Scatter(x=valid['date'],y=valid['actual_liquid'],name='Actual',mode='lines+markers',line=dict(color='#4DB9B4',width=3)))
            graph.add_trace(go.Scatter(x=valid['date'],y=valid['forecast_liquid'],name='Forecast',mode='lines+markers',line=dict(color='#D3AD71',width=3,dash='dash')))
            graph.update_layout(template='plotly_dark',paper_bgcolor='#0B111B',plot_bgcolor='#0B111B',height=360,yaxis_title='SAR')
            st.plotly_chart(graph,use_container_width=True)
        else:st.info('أدخل النقد والمحفظة معًا في سجل Update Log للأشهر الواقعة بين أكتوبر 2026 و2040 لإظهار المقارنة.')
        st.dataframe(actuals[['date','salary','bonus','cash','portfolio','debt','notes']].rename(columns={'date':'الشهر','salary':'الراتب','bonus':'البونص','cash':'النقد','portfolio':'المحفظة','debt':'الديون','notes':'ملاحظات'}),hide_index=True,use_container_width=True)

with t6:
    st.subheader('Net Worth & Portfolio | صافي الثروة والمحفظة')
    st.caption('صافي الثروة هنا = النقد + المحفظة − الديون المسجلة. لا يشمل قيمة العقارات والسيارات إلا إذا أضفتها لاحقًا ضمن نموذج أصول منفصل.')
    if actuals.empty:st.info('سجّل النقد والمحفظة والديون في Update Log حتى يظهر تطور صافي ثروتك.')
    else:
        tracked=actuals.dropna(subset=['cash','portfolio','debt']).copy()
        if tracked.empty:st.info('أدخل القيم الثلاث: الكاش، المحفظة، والديون. اكتب 0 عندما لا توجد ديون.')
        else:
            tracked['net_worth']=tracked['cash']+tracked['portfolio']-tracked['debt']
            latest=tracked.iloc[-1]
            a,b,c=st.columns(3)
            a.metric('صافي الثروة المسجل',sar(latest['net_worth']))
            b.metric('رصيد المحفظة',sar(latest['portfolio']))
            c.metric('إجمالي الديون',sar(latest['debt']))
            graph=go.Figure()
            for col,label,color in [('net_worth','Net Worth','#8D9FF2'),('portfolio','Portfolio','#D3AD71'),('cash','Cash','#4DB9B4')]:
                graph.add_trace(go.Scatter(x=tracked['date'],y=tracked[col],mode='lines+markers',name=label,line=dict(color=color,width=3)))
            graph.update_layout(template='plotly_dark',paper_bgcolor='#0B111B',plot_bgcolor='#0B111B',height=380,yaxis_title='SAR')
            st.plotly_chart(graph,use_container_width=True)
            st.markdown('#### Portfolio balance trend | نمو رصيد المحفظة')
            st.caption('هذا يعرض تغير قيمة المحفظة، وليس العائد الاستثماري الحقيقي؛ لأن الإيداعات والسحوبات غير مسجلة بصورة مستقلة.')
            changes=tracked[['date','portfolio']].copy()
            changes['change_SAR']=changes['portfolio'].diff()
            st.dataframe(changes.rename(columns={'date':'الشهر','portfolio':'رصيد المحفظة','change_SAR':'تغير الرصيد (ريال)'}),hide_index=True,use_container_width=True)
    st.markdown('#### Goal readiness | جاهزية الأهداف')
    if actuals.empty or actuals['cash'].dropna().empty:
        st.caption('أدخل الكاش الفعلي حتى نحسب جاهزيتك النقدية للأهداف القادمة.')
    else:
        current_cash=float(actuals['cash'].dropna().iloc[-1])
        goal_rows=[('زواج',int(d[14]),d[15]),('السيارة الثانية',int(d[24]),d[25]),('دفعة الفيلا والرسوم',int(d[26]),d[28]+d[29])]
        for name,yr,cost in goal_rows:
            target=max(0,float(cost))
            st.write(f'**{name} ({yr})** — المطلوب {sar(target)}')
            st.progress(min(1.0,max(0.0,current_cash/target)) if target else 1.0)
            st.caption(f'التغطية بالنقد المسجل: {sar(current_cash)} · الفجوة: {sar(max(0,target-current_cash))}. لا تعني التغطية الحالية القدرة على تحمل الالتزامات الأخرى.')


# Download the updated source workbook: input cells are synced, formulas preserved.
wb=load_workbook(io.BytesIO(source));ws=wb['Inputs']
ews=wb['Early Years'] if 'Early Years' in wb else wb.create_sheet('Early Years')
ews.cell(1,1,'Key');ews.cell(1,2,'Value');ews.cell(1,3,'Description')
for idx,(k,v) in enumerate(early.items(),2):
 ews.cell(idx,1,k);ews.cell(idx,2,v);ews.cell(idx,3,EARLY_LABELS[k])
for i,v in d.items():
    row=workbook_rows(wb)[FIELD_NAMES[i]]
    ws.cell(row,2).value=int(v) if i in YEARS else v
try:wb.calculation.fullCalcOnLoad=True;wb.calculation.forceFullCalc=True
except Exception:pass
buf=io.BytesIO();wb.save(buf)
st.sidebar.download_button('⬇️ تنزيل الإكسل بالافتراضات الجديدة',buf.getvalue(),file_name='wealth_plan_updated.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
st.sidebar.caption('Excel Inputs ↔ Dashboard: القيم تُقرأ مباشرة من تبويب Inputs وتُكتب إلى نفس الخلايا عند التنزيل. لا يوجد حفظ تلقائي على خادم.')
st.caption('مرتبط بملف Excel: Inputs للتحكم، Update Log للأرقام الفعلية، وMonthly Plan / Annual Summary للتوقعات. قد توجد فروقات طفيفة في توقيت الأقساط؛ راجع شروط التمويل الفعلية.')
