# -*- coding: utf-8 -*-
"""
Uppdatera inventeringslager: skriver resultatet fran ledning_analys.py (ledningsresultat.json) till
falt pa de lokala ledningslagren (fran Exportera omrade), sa att kartan kan symboliseras och
etiketteras efter schaktgrupp, djupklass, metod osv. och sa att Kartor for omradet kan satta
etiketter med ledningens id.

Ett GIS-objekt kan vara delat i flera ledningar i inventeringen (brunn mitt pa objektet); da
summeras langder och kostnader, medeldjupet viktas med langden och schaktgrupp/djupklass/schaktsatt
tas fran den langsta delen. Objekt som inte finns i resultatet far tomma falt (aldre varden tas bort).

Faltnamnen ar hogst 10 tecken (shapefil). Featureklasser med klasstillagg (SVOA:s geopipe) uppdateras
i en redigeringssession om det kravs. Python 2.7 + arcpy (ArcMap 10.x); testas med latsas-arcpy.
"""
from __future__ import unicode_literals

import io
import json
import os

import arcpy

from skapa_ledningslager import txt, logg, hitta_lager, kalla, _tal

# (falt, typ, langd, alias)
FALT = [
    ('LEDN_ID', 'TEXT', 80, 'Ledning (id)'),
    ('SYS', 'TEXT', 4, 'System'),
    ('DIM_MM', 'DOUBLE', None, 'Dimension (mm)'),
    ('LANGD_M', 'DOUBLE', None, 'Längd (m)'),
    ('SCHAKT_M', 'DOUBLE', None, 'Schaktlängd (m)'),
    ('DJUP_MEDEL', 'DOUBLE', None, 'Medeldjup (m)'),
    ('DJUPKLASS', 'TEXT', 12, 'Djupklass (m)'),
    ('SCHAKTGRP', 'LONG', None, 'Schaktgrupp'),
    ('I_GRAV', 'LONG', None, 'Ledningar i graven'),
    ('ALT_OMLAGG', 'DOUBLE', None, 'Alt. lägg om (kr)'),
    ('ALT_SKFRI', 'DOUBLE', None, 'Alt. schaktfritt (kr)'),
    ('METOD_SKFR', 'TEXT', 40, 'Metod alt. schaktfritt'),
    ('SCHAKTSATT', 'TEXT', 30, 'Schaktsätt alt. lägg om'),
    ('ANMARKNING', 'TEXT', 254, 'Anmärkning'),
]
SCHAKTSATT_NAMN = {'ingen': 'ingen spont', 'spontlada': 'spontlåda', 'slagen': 'slagen spont'}


def las_resultat(path):
    """ledningsresultat.json fran ledning_analys.py."""
    with io.open(txt(path), encoding='utf-8') as f:
        data = json.load(f)
    if data.get('kalla') != 'ledning_analys.py' or not isinstance(data.get('ledningar'), list):
        raise RuntimeError('%s ar inte en ledningsresultat.json fran ledning_analys.py' % txt(path))
    return data


def per_objekt(ledningar):
    """{(lagernamn, oid): [ledningar]} - ett GIS-objekt kan vara delat i flera ledningar."""
    per = {}
    for l in ledningar:
        if not l.get('lager') or l.get('oid') is None:
            continue
        try:
            oid = int(l['oid'])
        except (TypeError, ValueError):
            continue
        per.setdefault((txt(l['lager']), oid), []).append(l)
    return per


def _summa(delar, nyckel):
    v = [_tal(l.get(nyckel)) for l in delar]
    if not v or any(x is None for x in v):
        return None
    return round(sum(v), 2)


def _unika(delar, nyckel, sep):
    ut = []
    for l in delar:
        v = txt(l.get(nyckel) or '').strip()
        if v and v not in ut:
            ut.append(v)
    return sep.join(ut) or None


def grund_id(delar):
    """Objektets id: ledningens id, utan '/del' nar objektet ar delat."""
    i = txt(delar[0].get('id') or '')
    if len(delar) > 1 and '/' in i:
        return i.rsplit('/', 1)[0]
    return i or None


def varden(delar):
    """Faltvardena (i FALT-ordning) for ett objekt."""
    delar = sorted(delar, key=lambda l: l.get('del') or 0)
    langst = max(delar, key=lambda l: _tal(l.get('langd_m')) or 0)
    L = _summa(delar, 'langd_m')
    djup = None
    vikt = [(_tal(l.get('djup_medel')), _tal(l.get('langd_m')) or 0) for l in delar]
    if all(d is not None for d, _ in vikt) and sum(w for _, w in vikt) > 0:
        djup = round(sum(d * w for d, w in vikt) / sum(w for _, w in vikt), 2)
    grav = [l.get('ledningar_i_grav') for l in delar if l.get('ledningar_i_grav') is not None]
    grp = langst.get('schaktgrupp')
    schakt = _summa(delar, 'schakt_langd_m')
    if schakt is None and L is not None and all(l.get('schaktgrupp') is None for l in delar):
        schakt = L                  # aldre resultatfil: utan schaktgrupp ar schaktet ledningens langd
    ss = langst.get('schaktsatt')
    return [
        grund_id(delar),
        txt(langst.get('system') or '') or None,
        _tal(langst.get('dimension')),
        L,
        schakt,
        djup,
        txt(langst.get('djupklass') or '') or None,
        int(grp) if grp is not None else None,
        max(grav) if grav else None,
        _summa(delar, 'alt_lagg_om_kr'),
        _summa(delar, 'alt_schaktfritt_kr'),
        _unika(delar, 'alt_schaktfritt_metod', ' + '),
        txt(ss) or None,
        (_unika(delar, 'anmarkning', '; ') or '')[:254] or None,
    ]


def arbetsyta(src):
    """Geodatabasen (eller mappen for en shapefil) som featureklassen ligger i."""
    p = txt(src)
    while p:
        if p.lower().endswith(('.gdb', '.mdb', '.sde')):
            return p
        upp = os.path.dirname(p)
        if upp == p:
            break
        p = upp
    return os.path.dirname(txt(src))


def med_redigering(src, funktion, namn):
    """Kor funktion() (som uppdaterar src med en UpdateCursor). Kraver featureklassen en
    redigeringssession (klasstillagg) provas en sadan."""
    try:
        return funktion()
    except RuntimeError as e:
        if 'edit session' not in txt(e).lower():
            raise
    ed = arcpy.da.Editor(arbetsyta(src))
    try:
        ed.startEditing(False, False)
        ed.startOperation()
        ut = funktion()
        ed.stopOperation()
        ed.stopEditing(True)
        logg('  %s: uppdaterat i en redigeringssession (klasstillagg i featureklassen)' % namn)
        return ut
    except Exception:
        try:
            ed.stopEditing(False)
        except Exception:
            pass
        raise


def lagg_till_falt(src, falt):
    """Lagger till falten som saknas. falt: [(namn, typ, langd, alias)]. Returnerar {NAMN: verkligt namn}."""
    finns = dict((f.name.upper(), f.name) for f in arcpy.ListFields(src))
    for namn, typ, langd, alias in falt:
        if namn.upper() in finns:
            continue
        if langd:
            arcpy.AddField_management(src, namn, typ, field_length=langd, field_alias=alias)
        else:
            arcpy.AddField_management(src, namn, typ, field_alias=alias)
    return dict((f.name.upper(), f.name) for f in arcpy.ListFields(src))


def uppdatera(resultat_json):
    """Skriver FALT till de ledningslager som ledningsresultat.json namner. Returnerar
    {lagernamn: antal uppdaterade objekt}."""
    data = las_resultat(resultat_json)
    per = per_objekt(data['ledningar'])
    lagernamn = sorted(set(k[0] for k in per))
    logg('%d ledningar i %s (%d GIS-objekt i %d lager)' % (len(data['ledningar']), os.path.basename(txt(resultat_json)),
                                                           len(per), len(lagernamn)))
    ut = {}
    saknas = []
    for namn in lagernamn:
        try:
            lyr = hitta_lager(namn)
            src = kalla(lyr)[0]
            arcpy.Describe(src)
        except Exception as e:
            saknas.append(namn)
            logg('  OBS: lagret "%s" finns inte i kartan (%s) - hoppas over' % (namn, txt(e)))
            continue
        try:
            namnmap = lagg_till_falt(src, FALT)
        except Exception as e:
            logg('  OBS: kunde inte lagga till falten i %s (%s) - ar lagret skrivskyddat (SDE) eller last? '
                 'Kor mot de lokala lagren fran Exportera omrade.' % (namn, txt(e)))
            continue
        kolumner = ['OID@'] + [namnmap[f[0].upper()] for f in FALT]
        mina = dict((oid, delar) for (l, oid), delar in per.items() if l == namn)

        def skriv(src=src, kolumner=kolumner, mina=mina):
            n = rensade = 0
            with arcpy.da.UpdateCursor(src, kolumner) as cur:
                for rad in cur:
                    delar = mina.get(rad[0])
                    if delar:
                        cur.updateRow([rad[0]] + varden(delar))
                        n += 1
                    elif any(v not in (None, '') for v in rad[1:]):
                        cur.updateRow([rad[0]] + [None] * len(FALT))
                        rensade += 1
            return n, rensade
        try:
            n, rensade = med_redigering(src, skriv, namn)
        except Exception as e:
            logg('  OBS: kunde inte uppdatera %s: %s' % (namn, txt(e)))
            continue
        ut[namn] = n
        logg('  %s: %d objekt uppdaterade%s%s' % (namn, n,
                                                  ', %d i resultatet hittades inte' % (len(mina) - n) if len(mina) > n else '',
                                                  ', %d andra objekt tomda' % rensade if rensade else ''))
    if saknas:
        logg('  %d lager i resultatet saknas i kartan - lagg in de lokala lagren (samma namn som i Ledningsinventering)'
             % len(saknas))
    logg('Falt: %s. Symbolisera t.ex. pa SCHAKTGRP, DJUPKLASS eller METOD_SKFR och etikettera med LEDN_ID.'
         % ', '.join(f[0] for f in FALT))
    logg('KLART')
    return ut
