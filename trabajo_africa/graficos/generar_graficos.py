from pathlib import Path
from html import escape
import json
p=Path(__file__).resolve().parent.parent; out=p/'graficos'
BG='#FAFBFC'; INK='#17313D'; MUT='#536873'; TEAL='#087F8C'; GOLD='#B97832'
def text(x,y,s,size=20,color=INK,weight='400',anchor='start'):
 return f'<text x="{x}" y="{y}" font-family="Helvetica, Arial, sans-serif" font-size="{size}" fill="{color}" font-weight="{weight}" text-anchor="{anchor}">{escape(s)}</text>'
def rect(x,y,w,h,fill,rx=14):return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="{fill}"/>'
def line(x,y,xx,yy,color='#DCE5E9',width=1):return f'<path d="M{x},{y} L{xx},{yy}" fill="none" stroke="{color}" stroke-width="{width}"/>'
def save(name,title,subtitle,parts,foot):
 svg=f'<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="680" viewBox="0 0 1200 680" role="img"><title>{escape(title)}</title>'+rect(0,0,1200,680,BG,0)+text(55,60,title,30,weight='700')+text(55,98,subtitle,18,MUT)+''.join(parts)+line(55,592,1145,592)+''.join(text(55,625+i*23,s,15,MUT) for i,s in enumerate(foot))+'</svg>'
 (out/(name+'.svg')).write_text(svg)
# Selected years only, separated into two pairs; no interpolated trajectory.
a=[]; baseline=492; scale=38
for v in [0,2,4,6,8]:
 y=baseline-v*scale;a += [line(105,y,1145,y),text(85,y+6,str(v),17,MUT,anchor='end')]
a += [text(55,145,'Variación anual del PIB real (%)',17,MUT)]
for x,year,v in [(245,2020,-.3),(425,2021,7.6),(805,2024,4.7),(985,2025,4.6)]:
 y=baseline-v*scale;color=GOLD if v<0 else TEAL
 a += [rect(x-52,min(y,baseline),104,abs(v)*scale,color,3),text(x,(y-15 if v>=0 else y+30),str(v).replace('.',',')+' %',23,color,'700','middle'),text(x,555,str(year),21,INK,'600','middle')]
a += [line(615,180,615,560,'#CCD8DF',1),text(335,162,'Pandemia y rebote',18,MUT,anchor='middle'),text(895,162,'Años recientes',18,MUT,anchor='middle')]
save('01_crecimiento_pib','Figura 1. Crecimiento de Kenia: crisis y recuperación','Años seleccionados · No se representan 2022 y 2023',a,['Fuente: elaboración propia con KNBS, Economic Survey 2026 [2], y CBK, Annual GDP [4].','La serie revisada sitúa 2021 en 7,6 %. Ambas instituciones reproducen la serie nacional.'])
# Store exact small input series, without suggesting a new external dataset download.
(p/'datos/elaborados/graficos/datos_grafico_pib.json').write_text(json.dumps({'variable':'variación anual del PIB real','unidad':'porcentaje','serie':'nacional revisada, edición 2026','observaciones':[{'año':y,'valor':v} for y,v in [(2020,-.3),(2021,7.6),(2024,4.7),(2025,4.6)]],'fuentes_referencias':[2,4],'nota':'Transcripción de las cifras documentadas en el informe; selección de años, no serie completa.'},ensure_ascii=False,indent=2)+'\n')
