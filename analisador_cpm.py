# -*- coding: utf-8 -*-
"""
Analisador de gargalos da sequencia de producao (CPM) - v2.

Duas familias de regras (config em regras.json), aberto para mais:
  * regras_unidade  -> complexidade por carro (opcional)
  * regras_padrao   -> gargalos posicionais na sequencia:
        - contagem_janela        : N ocorrencias do predicado numa janela de W
        - sequencia_consecutiva  : K+ posicoes consecutivas que batem o predicado
        - coocorrencia_janela    : dentro de W, todos os grupos presentes

Saida: analise.json (consumido pelo dashboard HTML).
Uso:  python analisador_cpm.py MIX.xlsx regras.json analise.json
"""
import sys, json, unicodedata, datetime
import pandas as pd

DISPLAY = ["seq","popid","chassi","modelo","comprim","susp",
           "mu","rpx","opc_fmt","barra2e","barra_tras"]

def _norm(v):
    if v is None: return ""
    s = str(v).strip().upper()
    return unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")

def _numero(valor):
    """Texto -> numero, aceitando virgula decimal. None quando nao da."""
    if valor is None: return None
    txt = str(valor).strip().replace(" ", "")
    if not txt: return None
    if "," in txt and "." in txt:
        txt = (txt.replace(".", "").replace(",", ".")
               if txt.rfind(",") > txt.rfind(".") else txt.replace(",", ""))
    else:
        txt = txt.replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return None

def _cond(u, c):
    op=c["operador"]
    # Comparacao numerica (COMPRIM > 7400 e afins): texto nao serve, "800" > "7400"
    if op in ("maior_que","maior_igual","menor_que","menor_igual"):
        atual=_numero(u.get(c["coluna"])); alvo=_numero(c.get("valor"))
        if atual is None or alvo is None: return False
        if op=="maior_que":   return atual >  alvo
        if op=="maior_igual": return atual >= alvo
        if op=="menor_que":   return atual <  alvo
        return atual <= alvo

    v=_norm(u.get(c["coluna"],"")); a=_norm(c.get("valor",""))
    if op=="contem":     return a in v
    if op=="nao_contem": return a not in v
    if op=="igual":      return v==a
    if op=="diferente":  return v!=a
    if op=="vazio":      return v==""
    if op=="nao_vazio":  return v!=""
    raise ValueError(f"operador desconhecido: {op}")

def _predicado(u, pred):
    res=[_cond(u,c) for c in pred["condicoes"]]
    return any(res) if pred.get("logica","E").upper()=="OU" else all(res)

def classifica_unidades(unidades, regras_unidade):
    for u in unidades:
        pts,ids=0,[]
        for r in regras_unidade:
            if _predicado(u,{"logica":r.get("logica","E"),"condicoes":r["condicoes"]}):
                pts+=r.get("peso",1); ids.append(r["id"])
        u["pts"]=pts; u["regras_unidade"]=ids
    return unidades

def _merge(spans):
    if not spans: return []
    spans=sorted(spans,key=lambda s:s[0]); out=[list(spans[0])]
    for a,b,m in spans[1:]:
        if a<=out[-1][1]+1:
            out[-1][1]=max(out[-1][1],b); out[-1][2]=sorted(set(out[-1][2])|set(m))
        else: out.append([a,b,sorted(set(m))])
    return out

def aplica_regra_padrao(unidades, regra):
    n=len(unidades); tipo=regra["tipo"]; spans=[]
    if tipo=="contagem_janela":
        W=regra["janela"]; minc=regra["min_ocorrencias"]; pred=regra["predicado"]
        mset={i for i,u in enumerate(unidades) if _predicado(u,pred)}
        for i in range(0,max(1,n-W+1)):
            jan=[p for p in range(i,min(n,i+W)) if p in mset]
            if len(jan)>=minc: spans.append((jan[0],jan[-1],jan))
    elif tipo=="sequencia_consecutiva":
        k=regra["min_seq"]; pred=regra["predicado"]; run=[]
        for i in range(n+1):
            hit = i<n and _predicado(unidades[i],pred)
            if hit: run.append(i)
            else:
                if len(run)>=k: spans.append((run[0],run[-1],list(run)))
                run=[]
    elif tipo=="unidade_critica":
        # O produto sozinho ja e gargalo: cada unidade que bate vira uma faixa
        # de uma posicao so. Duas seguidas viram uma faixa unica no _merge.
        pred=regra["predicado"]
        for i,u in enumerate(unidades):
            if _predicado(u,pred): spans.append((i,i,[i]))
    elif tipo=="coocorrencia_janela":
        W=regra["janela"]; grupos=regra["grupos"]
        for i in range(0,max(1,n-W+1)):
            win=list(range(i,min(n,i+W))); mpg=[]; ok=True
            for g in grupos:
                m=[p for p in win if _predicado(unidades[p],g)]
                if not m: ok=False; break
                mpg.append(m)
            if ok:
                todos=sorted({p for ms in mpg for p in ms}); spans.append((todos[0],todos[-1],todos))
    else:
        raise ValueError(f"tipo de regra desconhecido: {tipo}")
    return _merge(spans)

def _detalhe(r,a,b,membros):
    if r["tipo"]=="unidade_critica":
        if len(membros)==1:
            return ("Produto de alta complexidade. Verificar carga da estacao "
                    "e apoio no posto.")
        return (f"{len(membros)} produtos de alta complexidade em sequencia. "
                f"Intercalar produtos simples entre eles.")
    if r["tipo"]=="contagem_janela":
        return (f"{len(membros)} unidades do padrao dentro de {r['janela']} carros "
                f"(min. {r['min_ocorrencias']}). Espacar para nao sobrecarregar a estacao.")
    if r["tipo"]=="sequencia_consecutiva":
        return (f"{len(membros)} unidades consecutivas (min. {r['min_seq']}). "
                f"Intercalar unidade fora do padrao para quebrar a sequencia.")
    if r["tipo"]=="coocorrencia_janela":
        return (f"Grupos criticos coincidindo em {r['janela']} carros. "
                f"Separar os itens para reduzir a carga combinada.")
    return ""

def gera_gargalos(unidades, regras_padrao):
    gargalos=[]; gid=0
    for r in regras_padrao:
        for a,b,membros in aplica_regra_padrao(unidades,r):
            gid+=1
            gargalos.append({"id":f"g{gid}","regra_id":r["id"],"regra_nome":r["nome"],
                "tipo":r["tipo"],"severidade":r.get("severidade","ALTA"),
                "a":a,"b":b,"membros":membros,"detalhe":_detalhe(r,a,b,membros)})
    gargalos.sort(key=lambda g:(g["a"],g["b"]))
    return gargalos

def marca_participacao(unidades, gargalos):
    ordem={"ALTA":2,"MEDIA":1,"NENHUM":0}
    for u in unidades: u["gargalos"]=[]; u["heat"]="NENHUM"
    for g in gargalos:
        for p in g["membros"]:
            unidades[p]["gargalos"].append(g["id"])
            if ordem[g["severidade"]]>ordem[unidades[p]["heat"]]:
                unidades[p]["heat"]=g["severidade"]
    return unidades

def _filtrar_opcionais(v):
    filt=["HD FRONT","DISCO","SDAR","ABAFADOR","FUP","LTDA"]
    r=["HD" if x=="HD FRONT" else x for x in filt if x in str(v)]
    return "/".join(r)

def carrega(caminho):
    df=pd.read_excel(caminho,dtype=str).fillna(""); cols=df.columns
    ja_formatado = "SEQ." in cols and "BARRA TRASEIRA" in cols and \
        not any(c in cols for c in ["SCOB","CHASSIS HEIGHT"])
    unidades=[]
    for _,r in df.iterrows():
        if not str(r.get("MODELO","")).strip(): continue
        opc=r.get("OPCIONAIS","")
        opc_fmt=opc if ja_formatado else _filtrar_opcionais(opc)
        unidades.append({
            "seq": r.get("SEQ.","") or r.get("SEQ",""),
            "popid": r.get("POPID",""), "chassi": r.get("CHASSI",""),
            "modelo": r.get("MODELO",""), "comprim": r.get("COMPRIM",""),
            "susp": r.get("SUSP TRASEIRA",""), "mu": r.get("MU",""),
            "rpx": r.get("RPX",""), "opc_fmt": opc_fmt,
            "OPCIONAIS": opc, "barra2e": r.get("BARRA 2°E DIANT",""),
            "barra_tras": r.get("BARRA TRASEIRA",""),
            "MODELO": r.get("MODELO",""), "RPX": r.get("RPX",""),
            "BARRA TRASEIRA": r.get("BARRA TRASEIRA",""),
        })
    for i,u in enumerate(unidades): u["pos"]=i
    return unidades

def analisar(entrada, arq_regras):
    cfg=json.load(open(arq_regras,encoding="utf-8"))
    unidades=carrega(entrada)
    classifica_unidades(unidades,cfg.get("regras_unidade",[]))
    gargalos=gera_gargalos(unidades,cfg.get("regras_padrao",[]))
    marca_participacao(unidades,gargalos)
    saida_unid=[]
    for u in unidades:
        d={k:u.get(k,"") for k in ["pos"]+DISPLAY}
        d["gargalos"]=u["gargalos"]; d["heat"]=u["heat"]; d["pts"]=u.get("pts",0)
        saida_unid.append(d)
    regras_meta=[{"id":r["id"],"nome":r["nome"],"tipo":r["tipo"],
                  "severidade":r.get("severidade","ALTA")} for r in cfg.get("regras_padrao",[])]
    return {"meta":{"arquivo":entrada.split("/")[-1],
                    "gerado_em":datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "n_unidades":len(unidades),"n_gargalos":len(gargalos)},
            "regras":regras_meta,"unidades":saida_unid,"gargalos":gargalos}

def main(entrada, arq_regras, saida_json):
    dados=analisar(entrada,arq_regras)
    json.dump(dados,open(saida_json,"w",encoding="utf-8"),ensure_ascii=False,indent=1)
    print(f"OK -> {saida_json}")
    print(f"  unidades : {dados['meta']['n_unidades']}")
    print(f"  gargalos : {dados['meta']['n_gargalos']}")
    for r in dados["regras"]:
        c=sum(1 for g in dados["gargalos"] if g["regra_id"]==r["id"])
        print(f"    - {r['nome']:48s} {c} faixa(s)")
    return dados

if __name__=="__main__":
    ent=sys.argv[1] if len(sys.argv)>1 else "MIX.xlsx"
    reg=sys.argv[2] if len(sys.argv)>2 else "regras.json"
    out=sys.argv[3] if len(sys.argv)>3 else "analise.json"
    main(ent,reg,out)


# ============================================================================
# v3 — analise a partir de unidades ja montadas (ex: API) + logica de SETOR
# ============================================================================
def _serializa(unidades):
    saida = []
    for u in unidades:
        d = {k: u.get(k, "") for k in ["pos"] + DISPLAY}
        if "posto" in u:
            d["posto"] = u["posto"]
        d["gargalos"] = u.get("gargalos", [])
        d["heat"] = u.get("heat", "NENHUM")
        d["pts"] = u.get("pts", 0)
        saida.append(d)
    return saida


def enriquecer_setor(dados, setor):
    """Adiciona status/ETA de cada gargalo em relacao aos postos do setor.
    setor = {postos:[...], lookahead:int, takt_seg:float|None}
    Assume que o carro flui no sentido crescente de 'posto' (entra por min, sai por max).
    """
    postos = setor["postos"]
    look = setor.get("lookahead", 8)
    takt = setor.get("takt_seg")
    sentido = setor.get("sentido", "crescente")   # crescente: entra no menor posto
    setor_set = set(postos)
    upos = {u["pos"]: u.get("posto") for u in dados["unidades"]}
    cresc = sentido != "decrescente"
    entrada = min(postos) if cresc else max(postos)
    saida = max(postos) if cresc else min(postos)

    def a_montante(mpost):   # ainda vai chegar no setor?
        return (max(mpost) < entrada) if cresc else (min(mpost) > entrada)

    def a_jusante(mpost):    # ja passou
        return (min(mpost) > saida) if cresc else (max(mpost) < saida)

    def eta_de(mpost):       # distancia em postos ate a entrada do setor
        return (entrada - max(mpost)) if cresc else (min(mpost) - entrada)

    chegando, no_setor = [], []
    for g in dados["gargalos"]:
        mpost = [upos[p] for p in g["membros"] if upos.get(p) is not None]
        if not mpost:
            g["setor_status"] = "SEM_POSTO"; g["eta_postos"] = None; continue
        if any(p in setor_set for p in mpost):
            g["setor_status"] = "NO_SETOR"; g["eta_postos"] = 0
            no_setor.append(g)
        elif a_montante(mpost):
            eta = eta_de(mpost)
            g["setor_status"] = "CHEGANDO"; g["eta_postos"] = eta
            g["eta_min"] = round(eta * takt / 60, 1) if takt else None
            if eta <= look:
                chegando.append(g)
        elif a_jusante(mpost):
            g["setor_status"] = "PASSOU"; g["eta_postos"] = None
        else:
            g["setor_status"] = "PARCIAL"; g["eta_postos"] = 0
            no_setor.append(g)

    chegando.sort(key=lambda g: g["eta_postos"])
    dados["meta"]["setor"] = {"postos": postos, "lookahead": look, "takt_seg": takt}
    dados["setor_resumo"] = {
        "no_setor": [g["id"] for g in no_setor],
        "chegando": [{"id": g["id"], "eta_postos": g["eta_postos"],
                      "eta_min": g.get("eta_min")} for g in chegando],
    }
    return dados


def analisar_unidades(unidades, cfg, setor=None):
    """Pipeline a partir de uma lista de unidades ja montada (API/mock)."""
    for i, u in enumerate(unidades):
        u["pos"] = i
    classifica_unidades(unidades, cfg.get("regras_unidade", []))
    gargalos = gera_gargalos(unidades, cfg.get("regras_padrao", []))
    marca_participacao(unidades, gargalos)
    dados = {
        "meta": {
            "arquivo": "AO VIVO",
            "gerado_em": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "n_unidades": len(unidades),
            "n_gargalos": len(gargalos),
        },
        "regras": [{"id": r["id"], "nome": r["nome"], "tipo": r["tipo"],
                    "severidade": r.get("severidade", "ALTA")}
                   for r in cfg.get("regras_padrao", [])],
        "unidades": _serializa(unidades),
        "gargalos": gargalos,
    }
    if setor:
        enriquecer_setor(dados, setor)
    return dados
