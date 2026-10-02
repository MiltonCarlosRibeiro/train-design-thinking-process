"""Animações e transições para os slides (python-pptx não tem API para isso: geramos o XML <p:timing>).

Uso dentro de uma função de slide:

    A.begin(slide)                      # (feito por new_slide)
    with A.beat("rise"):                # tudo criado aqui entra junto
        box(...); text(...)
    with A.beat("zoom", gap=0.1):       # entra 0,1 s depois do início do beat anterior
        circle(...)
    with A.beat("fade", click=True):    # espera um clique do apresentador
        ...

Shapes criados fora de um beat ficam estáticos (cabeçalho, rodapé, logo).
Por padrão os beats começam sozinhos ao entrar no slide, em cascata: o instrutor não precisa clicar.
"""
from contextlib import contextmanager

from lxml import etree
from pptx.oxml.ns import qn

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# efeito -> (presetID, presetSubtype, duração padrão em ms)
EFFECTS = {
    "fade": (10, 0, 500),
    "rise": (42, 0, 650),       # Float In: surge subindo levemente
    "zoom": (53, 16, 450),      # Zoom com um pequeno "pop" no final
    "wipe": (22, 8, 600),       # Wipe a partir da esquerda
    "wipe_up": (22, 4, 600),    # Wipe a partir de baixo
    "words": (10, 0, 420),      # Fade palavra por palavra
}


class _Rec:
    slide = None
    groups = None       # lista de grupos (0 = automático, 1+ = por clique); cada grupo é lista de beats
    cur = None
    t = 0.0             # início do último beat (s), relativo ao grupo
    t_end = 0.0         # fim do último beat (s)
    transition = "fade"


R = _Rec()


def begin(slide, transition="fade"):
    R.slide, R.groups, R.cur, R.t, R.t_end, R.transition = slide, [[]], None, 0.0, 0.0, transition


def track(shape):
    if R.cur is not None:
        R.cur["shapes"].append(shape)
    return shape


def transition(kind):
    R.transition = kind


@contextmanager
def beat(effect="rise", gap=0.15, after=False, dur=None, click=False, words_pct=35, stagger=0.0):
    """gap: segundos após o INÍCIO do beat anterior (after=True: após o FIM).
    stagger: atraso extra entre os shapes do mesmo beat."""
    if R.slide is None:
        yield
        return
    if click:
        R.groups.append([])
        start = 0.0
    elif not R.groups[-1]:
        start = 0.0
    else:
        start = (R.t_end if after else R.t) + gap
    pid, sub, d = EFFECTS[effect]
    d = int(dur * 1000) if dur else d
    b = {"effect": effect, "start": start, "dur": d, "shapes": [], "pct": words_pct, "stagger": stagger}
    prev, R.cur = R.cur, b
    try:
        yield b
    finally:
        R.cur = prev
    if not b["shapes"]:
        return
    R.groups[-1].append(b)
    total = d / 1000 + stagger * (len(b["shapes"]) - 1)
    if effect == "words":
        nw = max(_word_count(s) for s in b["shapes"])
        total += (nw - 1) * d / 1000 * words_pct / 100
    R.t, R.t_end = start, max(R.t_end if not click else 0, start + total)


def _word_count(shape):
    if not shape.has_text_frame:
        return 1
    return max(1, len(shape.text_frame.text.split()))


# ---------------------------------------------------------------- XML
class _Ids:
    def __init__(self):
        self.n = 0

    def __call__(self):
        self.n += 1
        return self.n


def _bhvr(ids, spid, dur, attr=None, extra=""):
    an = f"<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst>" if attr else ""
    return (f'<p:cBhvr{extra}><p:cTn id="{ids()}" dur="{dur}" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>{an}</p:cBhvr>')


def _set_visible(ids, spid):
    return (f'<p:set><p:cBhvr><p:cTn id="{ids()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
            f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>')


def _anim(ids, spid, dur, attr, keys):
    tav = "".join(f'<p:tav tm="{tm}"><p:val><p:strVal val="{v}"/></p:val></p:tav>' for tm, v in keys)
    return (f'<p:anim calcmode="lin" valueType="num">{_bhvr(ids, spid, dur, attr, " additive=\"base\"")}'
            f'<p:tavLst>{tav}</p:tavLst></p:anim>')


def _fx(ids, spid, dur, filt):
    return f'<p:animEffect transition="in" filter="{filt}">{_bhvr(ids, spid, dur)}</p:animEffect>'


def _behaviours(effect, ids, spid, dur):
    out = [_set_visible(ids, spid)]
    if effect in ("fade", "words"):
        out.append(_fx(ids, spid, dur, "fade"))
    elif effect == "rise":
        out.append(_fx(ids, spid, dur, "fade"))
        out.append(_anim(ids, spid, dur, "ppt_x", [(0, "#ppt_x"), (100000, "#ppt_x")]))
        out.append(_anim(ids, spid, dur, "ppt_y", [(0, "#ppt_y+.06"), (100000, "#ppt_y")]))
    elif effect == "zoom":
        for a in ("ppt_w", "ppt_h"):
            out.append(_anim(ids, spid, dur, a, [(0, "0"), (70000, f"#{a}*1.08"), (100000, f"#{a}")]))
        out.append(_fx(ids, spid, int(dur * 0.6), "fade"))
    elif effect == "wipe":
        out.append(_fx(ids, spid, dur, "wipe(left)"))
    elif effect == "wipe_up":
        out.append(_fx(ids, spid, dur, "wipe(down)"))
    return "".join(out)


def _effect_par(ids, b, shape, delay_ms, node):
    pid, sub, _ = EFFECTS[b["effect"]]
    spid = shape.shape_id
    it = f'<p:iterate type="wd"><p:tmPct val="{b["pct"] * 1000}"/></p:iterate>' if b["effect"] == "words" else ""
    ease = ' decel="100000"' if b["effect"] in ("rise", "zoom") else ""
    return (f'<p:par><p:cTn id="{ids()}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" fill="hold" grpId="0" '
            f'nodeType="{node}"{ease}><p:stCondLst><p:cond delay="{delay_ms}"/></p:stCondLst>{it}'
            f'<p:childTnLst>{_behaviours(b["effect"], ids, spid, b["dur"])}</p:childTnLst></p:cTn></p:par>')


def _timing_xml():
    ids = _Ids()
    root_id = ids()
    seq_id = ids()
    groups = []
    for gi, beats in enumerate(R.groups):
        if not beats:
            continue
        effects = []
        for b in beats:
            for k, shp in enumerate(b["shapes"]):
                node = "clickEffect" if gi > 0 and not effects else "withEffect"
                delay = int(round((b["start"] + k * b["stagger"]) * 1000))
                effects.append(_effect_par(ids, b, shp, delay, node))
        auto = gi == 0
        cond = (f'<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="{seq_id}"/></p:cond>'
                if auto else '<p:cond delay="indefinite"/>')
        groups.append(f'<p:par><p:cTn id="{ids()}" fill="hold"><p:stCondLst>{cond}</p:stCondLst><p:childTnLst>'
                      f'<p:par><p:cTn id="{ids()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                      f'<p:childTnLst>{"".join(effects)}</p:childTnLst></p:cTn></p:par>'
                      f'</p:childTnLst></p:cTn></p:par>')
    if not groups:
        return None
    seen, bld = set(), []
    for beats in R.groups:
        for b in beats:
            for shp in b["shapes"]:
                if shp.shape_id in seen or shp._element.tag != qn("p:sp"):
                    continue
                seen.add(shp.shape_id)
                bld.append(f'<p:bldP spid="{shp.shape_id}" grpId="0" animBg="1"/>')
    return (f'<p:timing xmlns:p="{P_NS}"><p:tnLst><p:par><p:cTn id="{root_id}" dur="indefinite" restart="never" nodeType="tmRoot">'
            f'<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="{seq_id}" dur="indefinite" nodeType="mainSeq">'
            f'<p:childTnLst>{"".join(groups)}</p:childTnLst></p:cTn>'
            f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
            f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
            f'<p:bldLst>{"".join(bld)}</p:bldLst></p:timing>')


TRANSITIONS = {
    "fade": '<p:transition xmlns:p="%s" spd="med"><p:fade/></p:transition>' % P_NS,
    "black": '<p:transition xmlns:p="%s" spd="slow"><p:fade thruBlk="1"/></p:transition>' % P_NS,
    "push": '<p:transition xmlns:p="%s" spd="med"><p:push dir="u"/></p:transition>' % P_NS,
    None: None,
}


def finish():
    """Grava transição + animações no slide atual."""
    if R.slide is None:
        return
    sld = R.slide._element
    for tag in ("p:transition", "p:timing"):
        for el in sld.findall(qn(tag)):
            sld.remove(el)
    tr = TRANSITIONS[R.transition]
    if tr:
        sld.append(etree.fromstring(tr))
    tm = _timing_xml()
    if tm:
        sld.append(etree.fromstring(tm))
    R.slide = None
