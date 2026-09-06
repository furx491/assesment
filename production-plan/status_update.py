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

# ================= الحالة المحدّثة =================
STATUS={}
# ١. الركن هاربر — في الكسوة، مرحلة التفصيل، نص المرحلة
for p,q in [('هاربر ركنة حرف U (اريكة)',7), ('هاربر ركنة حرف L (اريكة)',1)]:
    rem={'التفصيل والكسوة': from_op(p,'التفصيل والكسوة','تفصيل',0.5)}
    if 'الدهانات' in RT[p]: rem['الدهانات']=full(p,'الدهانات')
    STATUS[p]=dict(qty=q, rem=rem, cur='التفصيل والكسوة', stage='تفصيل (نص المرحلة)',
                   note='خلصت نجارة وسفنجة وقشرة')
# ٢. ماتشا فوتيه — خلصت سفنجة وقشرة، هتبدأ دهانات وكسوة. الكمية بقت 7.
p='ماتشا فوتيه (اريكة)'
STATUS[p]=dict(qty=7, rem={'التفصيل والكسوة':full(p,'التفصيل والكسوة'),'الدهانات':full(p,'الدهانات')},
               cur='التفصيل والكسوة', stage='هتبدأ دهانات وكسوة',
               note='⚠ القماش لسه متحددش — الكسوة ممكن تقف')
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
