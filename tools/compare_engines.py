from model import *
def price_control(catn,adults,kids,pl,plan,day,promo):
    c=plcol[pl]; info=cat[catn]; tot=adults+kids
    base=val(basecats[info['base']],c); mult=val(matrows[catn],c); wm=val(matrows[catn],c+1)
    if tot<info['mn'] or tot>info['mx'] or base==0 or mult==0: return None
    if day=='Вихідні' and wm==0: return None
    if plan in ('RO','ROSPA'): pa=pk=0
    else:
        key='BB|'+day if plan=='ROSPABB' else plan+'|'+day
        pa=val(pkgr[key],c); pk=val(pkgr[key],c+1)
        if pa==0: return None
    spa=0
    if plan in ('ROSPA','ROSPABB'):
        spa=val(spar[info['spa']],c)
        if spa==0: return None
    disc=G8 if plan=='BBSPA' else tariff.get(promo,0)
    acc=base*mult*(wm if day=='Вихідні' else 1)+max(0,tot-info['incl'])*G6
    return rnd(acc*(1-disc)+adults*pa+kids*pk+spa)
usable=['Осінь 2025','Осінь 2026 База ','Осінь 2026 (00-10%)','Осінь 2026 (Канікули) 0-10%','Осінь 2026 (10-20%)','Осінь 2026 (20-30%)']
plans=['RO','BB','ROSPA','ROSPABB']; days=['Будні','Вихідні']; promos=[t for t in tariff if t]
d_pc=d_ck=d_pk=t_=0
for pl in usable:
 for plan in plans:
  for day in days:
   for promo in promos:
    for catn,info in cat.items():
     for tot in range(1,int(info['mx'])+1):
      for kids in range(0,tot+1):
        a=price_sheet(catn,tot,pl,plan,day,promo)
        b=price_calc(catn,tot-kids,kids,pl,plan,day,promo)
        cc=price_control(catn,tot-kids,kids,pl,plan,day,promo)
        if a is None and b is None and cc is None: continue
        t_+=1
        if a!=cc: d_pc+=1
        if b!=cc: d_ck+=1
        if a!=b: d_pk+=1
print(f'усього сценаріїв: {t_}')
print(f'  Прайс ≠ Контроль:      {d_pc} ({d_pc/t_:.1%})')
print(f'  Калькулятор ≠ Контроль:{d_ck} ({d_ck/t_:.1%})')
print(f'  Прайс ≠ Калькулятор:   {d_pk} ({d_pk/t_:.1%})')
print()
print('Приклад: Котедж комфорт класу, Осінь 2026 (00-10%), BB, Будні, Rack Rate, 2 дор.+2 діт.:')
print('  Прайс      =',price_sheet('Котедж комфорт класу',4,'Осінь 2026 (00-10%)','BB','Будні','Rack Rate'))
print('  Калькулятор=',price_calc('Котедж комфорт класу',2,2,'Осінь 2026 (00-10%)','BB','Будні','Rack Rate'))
print('  Контроль   =',price_control('Котедж комфорт класу',2,2,'Осінь 2026 (00-10%)','BB','Будні','Rack Rate'))
