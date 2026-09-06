# -*- coding: utf-8 -*-
"""تاب مستقل لكل قسم — ينزل لرئيس القسم"""
import runpy, math, datetime as dt, openpyxl
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
g=runpy.run_path('sched.py')
ORDERS=g['ORDERS']; DAYS=g['DAYS']; W=g['WORKERS']; NORM=g['NORM_H']; OT=g['OT_H']; run=g['run']
done,hist,deptday,plan=run(False)          # واقعي — ساعات عادية
doneO,histO,deptdayO,planO=run(True)       # بالسهر
fin={}
for o in ORDERS:
    k=o['rank']; last=-1
    for d,Wd in o['work'].items():
        if hist[k][d]: last=max(last,max(t for t,_ in hist[k][d]))
    fin[k]=DAYS[last] if last>=0 else None
AR=['الاتنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت','الأحد']
byrank={o['rank']:o for o in ORDERS}

wb=openpyxl.Workbook(); wb.remove(wb.active)
HF=Font(bold=True,color='FFFFFF',size=11); HB=PatternFill('solid',fgColor='1F4E79')
DAYF=Font(bold=True,size=12); DAYB=PatternFill('solid',fgColor='BDD7EE')
P1=PatternFill('solid',fgColor='FFC7CE'); P2=PatternFill('solid',fgColor='FFF2CC')
FRI=PatternFill('solid',fgColor='E2EFDA'); OTF=PatternFill('solid',fgColor='F8CBAD')
thin=Border(*[Side(style='thin',color='B0B0B0')]*4)

def dept_tab(d, title):
    ws=wb.create_sheet(title); ws.sheet_view.rightToLeft=True
    ws.append([f'خطة شغل قسم: {d}',None,None,None,None,f'عدد العمال: {W[d]}',
               f'الطاقة اليومية: {W[d]*NORM:.0f} ساعة عادي + {W[d]*OT:.0f} ساعة سهر'])
    ws['A1'].font=Font(bold=True,size=14); ws.append([])
    hdr=['التاريخ','اليوم','الأولوية','كود الأمر','المنتج','الكمية','العميل',
         'يبدأ من مرحلة','ساعات الشغل','عدد العمال المقترح','سهر؟','آخر يوم للأمر']
    ws.append(hdr)
    for c in ws[3]: c.font=HF;c.fill=HB;c.alignment=Alignment(horizontal='center',wrap_text=True)
    row=4
    for t,day in enumerate(DAYS):
        items=plan[t][d]
        if not items: continue
        if day>dt.date(2026,10,10): break
        tot=sum(h for _,h in items)
        need=min(math.ceil(tot/NORM),W[d])
        exc=max(0.0, deptdayO[t][d]-W[d]*NORM)          # الزيادة لو شغّلنا سهر
        otp=min(W[d], math.ceil(exc/OT)) if exc>0.01 else 0
        ws.append([str(day),AR[day.weekday()],'— ملخص اليوم —','',
                   f'إجمالي {tot:.1f} ساعة',len(items),
                   f'الطاقة {W[d]*NORM:.0f} س',f'التحميل {100*tot/(W[d]*NORM):.0f}%','',
                   f'{need} عامل',
                   (f'{otp} يسهروا لتقديم الشغل' if otp else 'مش محتاج سهر'),''])
        for c in ws[row]:
            c.font=DAYF; c.fill=FRI if day.weekday()==4 else DAYB; c.border=thin
        row+=1
        for k,h in sorted(items,key=lambda x:(byrank[x[0]]['P'],byrank[x[0]]['sub'],-x[1])):
            o=byrank[k]
            stg = str(o['stage']) if (o['cur']==d and o['stage']) else 'من أول القسم'
            ws.append([str(day),AR[day.weekday()],f"P{o['P']}",o['id'],o['prod'],o['qty'],
                       (o['trk'] or '—')[:36],stg,round(h,2),
                       max(1,round(h/NORM,1)),'',str(fin[k]) if fin[k] else '—'])
            f=P1 if o['P']==1 else P2
            for c in ws[row]: c.fill=f; c.border=thin
            if otp: ws.cell(row=row,column=11).value='ممكن سهر'
            row+=1
    for i,x in enumerate([12,10,8,17,30,7,36,26,12,17,12,13],1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=x
    ws.freeze_panes='A4'
    return ws

ORDER_TABS=[('نجارة التنجيد','1-نجارة التنجيد'),('التفصيل والكسوة','2-الكسوة'),
            ('الدهانات','3-الدهانات'),('نجارة النوم والسفرة','4-نجارة النوم والسفرة'),
            ('السفنجة','5-السفنجة'),('القشرة','6-القشرة'),('الاستانلس','7-الاستانلس'),
            ('تشطيب التنجيد','8-تشطيب التنجيد'),('تشطيب نوم وسفرة','9-تشطيب نوم وسفرة')]

# ===== تاب ملخص لكل المشرفين =====
ws=wb.create_sheet('0-ملخص كل الأقسام'); ws.sheet_view.rightToLeft=True
ws.append(['ملخص يومي — كل قسم يشتغل كام ساعة وكام عامل']); ws['A1'].font=Font(bold=True,size=14)
ws.append([])
hdr=['التاريخ','اليوم']+[d for d,_ in ORDER_TABS]
ws.append(hdr)
for c in ws[3]: c.font=HF;c.fill=HB;c.alignment=Alignment(horizontal='center',wrap_text=True)
r=4
for t,day in enumerate(DAYS):
    if not deptday[t]: continue
    if day>dt.date(2026,10,10): break
    line=[str(day),AR[day.weekday()]]
    for d,_ in ORDER_TABS:
        h=deptday[t][d]
        line.append(f"{h:.0f} س / {min(math.ceil(h/NORM),W[d]) if h>0 else 0} عامل" if h>0.05 else '—')
    ws.append(line)
    for c in ws[r]:
        c.border=thin
        if day.weekday()==4: c.fill=FRI
    r+=1
for i,x in enumerate([12,10]+[20]*len(ORDER_TABS),1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=x
ws.freeze_panes='C4'

for d,title in ORDER_TABS: dept_tab(d,title)
wb.move_sheet('0-ملخص كل الأقسام', offset=-len(ORDER_TABS))
wb.save('خطة_المشرفين.xlsx')
print("تم: خطة_المشرفين.xlsx")
print("التابات:", wb.sheetnames)
for d,_ in ORDER_TABS:
    days=[t for t in range(len(DAYS)) if plan[t][d]]
    if days: print(f"  {d:<22} {len(days):>3} يوم شغل | من {DAYS[min(days)]} لـ {DAYS[max(days)]}")
