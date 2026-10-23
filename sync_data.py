#!/usr/bin/env python3
from pathlib import Path
from datetime import datetime,date,timedelta
import zipfile, xml.etree.ElementTree as ET, json,re
H=Path(__file__).resolve().parent
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def read_sheet(z,name,shared):
 w=ET.fromstring(z.read('xl/workbook.xml'))
 rel=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
 links={r.get('Id'):r.get('Target') for r in rel}
 target=None
 for sh in w.findall('.//m:sheet',ns):
  if sh.get('name')==name:
   target=links[sh.get('{'+ns['r']+'}id')];break
 if not target:raise ValueError('Missing '+name)
 target=target.lstrip('/')
 if not target.startswith('xl/'):target='xl/'+target
 x=ET.fromstring(z.read(target));output=[]
 for row in x.findall('.//m:sheetData/m:row',ns):
  d={}
  for cell in row.findall('m:c',ns):
   letters=re.match(r'[A-Z]+',cell.get('r')).group(0);c=0
   for k in letters:c=c*26+ord(k)-64
   t=cell.get('t','');v=cell.find('m:v',ns);inline=cell.find('m:is',ns)
   if t=='s' and v is not None:value=shared[int(v.text)]
   elif t=='inlineStr' and inline is not None:value=''.join(n.text or '' for n in inline.findall('.//m:t',ns))
   elif v is not None:value=v.text or ''
   else:value=''
   d[c-1]=value
  if d:output.append([d.get(i,'') for i in range(max(d)+1)])
 return output
def xldate(value):
 v=str(value).strip()
 if not v:return ''
 try:
  if re.fullmatch(r'\d+(?:\.\d+)?',v):
   return (date(1899,12,30)+timedelta(days=float(v))).isoformat()
  return datetime.fromisoformat(v[:10]).date().isoformat()
 except (ValueError,OverflowError):raise ValueError('Invalid date: '+v)
with zipfile.ZipFile(H/'PACT_100_JOURS_MASTER.xlsx') as z:
 shared=[]
 if 'xl/sharedStrings.xml' in z.namelist():
  ss=ET.fromstring(z.read('xl/sharedStrings.xml'))
  shared=[''.join(t.text or '' for t in si.findall('.//m:t',ns)) for si in ss.findall('m:si',ns)]
 simple=read_sheet(z,'Suivi simplifié',shared)
 settingsrows=read_sheet(z,'Paramètres',shared)
 procrows=read_sheet(z,'Marchés prioritaires',shared)
if not simple:raise SystemExit('Feuille Suivi simplifié vide')
fields=['ID','Groupe','Activité','Responsable','Statut','Avancement (%)','Jour du pacte','Date réelle de réalisation','Blocage / décision attendue','Prochaine étape']
cols={name:i for i,name in enumerate(simple[0])}
missing=[k for k in fields if k not in cols]
if missing:raise SystemExit('Colonnes absentes : '+', '.join(missing))
GROUP_INTERNAL={
 '1. Entrée en vigueur':'1. Effectiveness',
 '2. Mise en place et personnel de l’UGP':'2. PIU establishment & staffing',
 '3. Gestion fiduciaire, planification et opérations':'3. Fiduciary, planning & ops',
 '4. Administration et logistique':'4. Administration & logistics',
 '5. Lancement de la mise en œuvre':'5. Implementation launch',
}
STATUS={'À confirmer':'To confirm','En bonne voie':'On track','À risque':'At risk','En retard':'Overdue','Terminé':'Completed',
        'To confirm':'To confirm','On track':'On track','At risk':'At risk','Overdue':'Overdue','Completed':'Completed'}
settings={str(r[0]).strip():str(r[1]).strip() for r in settingsrows[1:] if len(r)>1 and r[0]}
start=xldate(settings.get('Date de début du pacte','2026-10-15'))
duration=int(float(settings.get('Durée (jours)','100')))
if not 1<=duration<=10000:raise SystemExit('Durée du pacte invalide')
acts=[];ids=set()
for raw in simple[1:]:
 row=raw+['']*(len(simple[0])-len(raw))
 v=lambda key:str(row[cols[key]] or '').strip()
 id=v('ID')
 if not id and not v('Activité'):continue
 if not id or not v('Activité'):raise SystemExit('ID ou activité manquant')
 if id in ids:raise SystemExit('ID en double : '+id)
 ids.add(id)
 group=v('Groupe')
 if group not in GROUP_INTERNAL:raise SystemExit('Groupe invalide pour '+id+': '+group)
 status=v('Statut') or 'À confirmer'
 if status not in STATUS:raise SystemExit('Statut invalide pour '+id+': '+status)
 try: progress=int(float(v('Avancement (%)') or 0))
 except ValueError:raise SystemExit('Avancement invalide pour '+id)
 if not 0<=progress<=100:raise SystemExit('Avancement hors limites pour '+id)
 rawday=v('Jour du pacte')
 if rawday:
  try:
   daynum=int(float(rawday))
   if float(rawday)!=daynum or not 1<=daynum<=duration:raise ValueError()
  except ValueError:raise SystemExit('Jour du pacte invalide pour '+id+': '+rawday)
 else:daynum=''
 completed=xldate(v('Date réelle de réalisation'))
 acts.append(dict(id=id,axis=GROUP_INTERNAL[group],title=v('Activité'),owner=v('Responsable'),status=STATUS[status],progress=progress,pactDay=daynum,deadlineType='relative' if daynum!='' else '',targetDate='',completedDate=completed,blocker=v('Blocage / décision attendue'),next=v('Prochaine étape')))
proc=[]
if procrows:
 headers=procrows[0]
 for raw in procrows[1:]:
  row=raw+['']*(len(headers)-len(raw)); d={k:str(row[i] or '').strip() for i,k in enumerate(headers)}
  if not d.get('Marché / lot',''):continue
  dtype=d.get('Mode échéance (fixe/relatif)','fixed') or 'fixed'
  day=d.get('Jour du pacte','')
  proc.append(dict(id=d.get('ID',''),title=d.get('Marché / lot',''),owner=d.get('Responsable',''),stage=d.get('Étape',''),amount=d.get('Montant',''),deadlineType=dtype,pactDay=int(float(day)) if day else '',targetDate=xldate(d.get('Date cible fixe','')),completedDate=xldate(d.get('Date de réalisation',''))))
(H/'data.js').write_text('const PACT_SETTINGS = '+json.dumps(dict(startDate=start,durationDays=duration,projectName=settings.get('Nom du projet','Projet à personnaliser'),projectCode=settings.get('Code du projet',''),projectTeam=settings.get('Équipe / auteurs',''),organization=settings.get('Organisation / pied de page','')),ensure_ascii=False,indent=2)+';\nconst PRIORITY_PROCUREMENTS = '+json.dumps(proc,ensure_ascii=False,indent=2)+';\nconst MOCK_ACTIONS = '+json.dumps(acts,ensure_ascii=False,indent=2)+';\nconst INFO_GAPS = [];\n',encoding='utf8')
print('Synchronisation réussie :',len(acts),'activités, source exclusive : Suivi simplifié')
