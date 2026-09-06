# -*- coding: utf-8 -*-
"""تحديثات حالة شفهية من المستخدمة — بتتقدّم على اللي في ملف أوامر التصنيع"""
import openpyxl
from collections import defaultdict
U='/root/.claude/uploads/fdb1b8fa-74fd-56bd-9c54-782073111ccf/'
def num(v): return v if isinstance(v,(int,float)) else 0
_wr=openpyxl.load_workbook(U+'7b215713-_________________....xlsx',data_only=True)['Sheet1']
RT=defaultdict(lambda: defaultdict(list))
for r in [x for x in _wr.iter_rows(values_only=True)][1:]:
    if r[5] and r[1]!='التغليف':
        RT[str(r[5]).strip()][r[1]].append((r[6] if isinstance(r[6],int) else 999,str(r[7]).strip(),num(r[8])))
for p in RT:
    for d in RT[p]: RT[p][d].sort()

def from_op(prod, dept, opname, frac_done=0.0):
    """باقي دقائق القسم من عملية معيّنة (مع نسبة إنجاز جوه العملية نفسها)"""
    l=RT[prod][dept]; i=0
    for j,(_,n,_m) in enumerate(l):
        if opname in n or n in opname: i=j; break
    return l[i][2]*(1-frac_done) + sum(x[2] for x in l[i+1:])
def full(prod,dept): return sum(x[2] for x in RT[prod][dept])

# ---- متوسط "فوتيه (اريكة)" للمنتجات اللي مالهاش مسار ----
_ref=[p for p in RT if 'فوتيه' in p and 'اريكة' in p]
def _avg(dept, skip_name=None):
    vals=[]
    for p in _ref:
        if dept not in RT[p]: continue
        vals.append(sum(x[2] for x in RT[p][dept] if not (skip_name and skip_name in x[1])))
    return sum(vals)/len(vals) if vals else 0.0

# ---- تشطيب التنجيد ناقص من ملف المسارات لمنتجات هاربر: متوسط الركن/الكنب ----
_fin=[sum(x[2] for x in RT[p]['تشطيب التنجيد']) for p in RT
      if 'تشطيب التنجيد' in RT[p] and ('ركن' in p or 'كنب' in p)]
FIN_AVG=sum(_fin)/len(_fin) if _fin else 35.0     # ≈35 دقيقة ("تقفيل")

# ================= الحالة المحدّثة =================
STATUS={}
# ١. هاربر ركنة حرف U — في الكسوة، مرحلة التفصيل، نص المرحلة
p='هاربر ركنة حرف U (اريكة)'
STATUS[p]=dict(qty=7,
    rem={'التفصيل والكسوة': from_op(p,'التفصيل والكسوة','تفصيل',0.5),
         'الدهانات': full(p,'الدهانات'), 'تشطيب التنجيد': FIN_AVG},
    cur='التفصيل والكسوة', stage='تفصيل (نص المرحلة)',
    note='خلصت نجارة وسفنجة | التشطيب وقت تقديري (ناقص من ملف المسارات)')
# ٢. هاربر ركنة حرف L — خلصت كسوة، فاضل دهانات وتشطيب بس
p='هاربر ركنة حرف L (اريكة)'
STATUS[p]=dict(qty=1,
    rem={'الدهانات': full(p,'الدهانات'), 'تشطيب التنجيد': FIN_AVG},
    cur='الدهانات', stage='الدهانات (من أولها)',
    note='خلصت كسوة | التشطيب وقت تقديري (ناقص من ملف المسارات)')
# ٢. ماتشا فوتيه — فوتيتين بس، قماشهم متحدد، خلصت سفنجة وقشرة، هتبدأ دهانات وكسوة
p='ماتشا فوتيه (اريكة)'
STATUS[p]=dict(qty=2, rem={'التفصيل والكسوة':full(p,'التفصيل والكسوة'),'الدهانات':full(p,'الدهانات')},
               cur='التفصيل والكسوة', stage='هتبدأ دهانات وكسوة',
               note='خلصت سفنجة وقشرة | القماش متحدد')
# ٣. تورين فوتيه — في السفنجة، والدهانات خلصت "خدمة". مالهاش مسار → متوسط فوتيه (اريكة)
STATUS['تورين فوتيه (اريكة)']=dict(qty=8,
    rem={'السفنجة':_avg('السفنجة'), 'التفصيل والكسوة':_avg('التفصيل والكسوة'),
         'الدهانات':_avg('الدهانات','خدمة')},
    cur='السفنجة', stage='السفنجة (من أولها)',
    note='وقت تقديري — متوسط فوتيه (اريكة). الدهانات بعد خصم "خدمة"')

if __name__=='__main__':
    print(f"{'المنتج':<30}{'كمية':>6}{'س/وحدة':>9}{'إجمالي':>9}   التوزيع")
    for p,s in STATUS.items():
        t=sum(s['rem'].values())/60
        print(f"{p:<30}{s['qty']:>6}{t:>9.2f}{t*s['qty']:>9.1f}   "+
              " | ".join(f"{d} {m/60:.2f}س" for d,m in s['rem'].items()))
        print(f"{'':30}  ↳ {s['note']}")

# ================= تحديثات إضافية =================
import datetime as _dt
# ٤. بابلي فوتيه — خلصت نجارة، شغالة كسوة بس. 7 منهم القماش يتحدد بكرة.
_kes=_avg('التفصيل والكسوة')
STATUS['بابلي فوتيه (اريكة)']=dict(qty=9, rem={'التفصيل والكسوة':_kes},
    cur='التفصيل والكسوة', stage='الكسوة (من أولها)', ready=_dt.date(2026,9,7),
    note='خلصت نجارة | وقت تقديري — متوسط فوتيه (اريكة)')
STATUS['بابلي فوتيه (اريكة) — قماش بكرة']=dict(qty=7, rem={'التفصيل والكسوة':_kes},
    cur='التفصيل والكسوة', stage='الكسوة (من أولها)', ready=_dt.date(2026,9,8),
    match='بابلي فوتيه (اريكة)',
    note='⚠ القماش يتحدد بكرة 8/9 — مش هيبدأ قبل كده')
# ٥. كرسي ستانلس أريكة — في السفنجة
STATUS['كرسي ستانلس اريكة']=dict(qty=1, rem=None, cur='السفنجة', stage='السفنجة (من أولها)',
    note='وقت تقديري — متوسط كرسي سفرة')

# ٦. نيست فوتيه بتوع مصطفي عبد الهادي (MO-2026-00038/39) — في الكسوة
#    بكود الأمر لأن في نيست فوتيه لعملاء تانيين لسه في السفنجة والنجارة
STATUS_BY_ID={
  'MO-2026-00038': dict(cur='التفصيل والكسوة', stage='الكسوة (من أولها)',
                        note='مصطفي عبد الهادي — في الكسوة | وقت تقديري'),
  'MO-2026-00039': dict(cur='التفصيل والكسوة', stage='الكسوة (من أولها)',
                        note='مصطفي عبد الهادي — في الكسوة | وقت تقديري'),
}

# ---- أوردرات خلصت خالص ----
FINISHED={'اوردين ترابيزة كبيرة (اريكة)'}
# ---- اتشالت من الإنتاج بأمر المستخدمة (بكود الأمر — في كراسي زين لعملاء تانيين) ----
# كراسي جنا حازم: 8 على فاتورة NF-003308 + 1 على FXN-000499
EXCLUDE_IDS={f'MO-2026-00{n}' for n in range(117,125)} | {'MO-2026-00208'}
# أوامر كودها BPO مش هتشتغل — ماعدا BPO-000055 (الـ128 كرسي سبايدر بتوع عمرو حسن)
BPO_KEEP={'BPO-000055'}
def is_bpo_excluded(order_id):
    t=str(order_id or '').upper()
    if not t.startswith('BPO'): return False
    return t.split('-')[0]+'-'+ (t.split('-')[1] if '-' in t else '') not in BPO_KEEP
# ---- مؤجّلة: مش هتشتغل الأسبوع ده ----
DEFERRED={'نيو سيول كرسي سفرة': _dt.date(2026,9,12)}
# ---- كل دول شغل نجارة النوم والسفرة مش نجارة التنجيد ----
NW_CATS={'ترابيزة رئيسية','ترابيزة جانبية','ترابيزة سفرة','مرايا','كومود','تسريحة','بوفيه','كونصول'}
NW_KEYWORDS=('عاصم',)
NW_EXCEPT={'اتووم كومود'}
def to_nw(prod, category):
    """هل شغل النجارة للمنتج ده بيتعمل في نجارة النوم والسفرة؟"""
    if prod in NW_EXCEPT: return False
    if any(k in prod for k in NW_KEYWORDS): return True
    return category in NW_CATS
def remap(work, prod, category):
    if not to_nw(prod, category) or 'نجارة التنجيد' not in work: return work
    w=dict(work); w['نجارة النوم والسفرة']=w.get('نجارة النوم والسفرة',0)+w.pop('نجارة التنجيد')
    return w
