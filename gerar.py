# -*- coding: utf-8 -*-
"""Gera o painel comercial (site/index.html) lendo a planilha da nuvem (link no config.json).
Uso:  python gerar.py                 (baixa a planilha do link do config.json)
      python gerar.py arquivo.xlsx    (teste local, sem baixar)
Se algo estiver errado, termina com erro e NAO gera o painel (o painel antigo continua no ar).
Requer: pip install openpyxl"""
import sys, re, json, os, io, urllib.request, urllib.parse, datetime as dt, unicodedata
from zoneinfo import ZoneInfo
import openpyxl

PASTA = os.path.dirname(os.path.abspath(__file__))
ETAPAS = [  # id, chave no título, nome, cor
    (0, 'radar',      'Radar',                  '#6b7280'),
    (1, 'solicit',    'Solicitação de orçamento','#d6204a'),
    (2, 'orcamento',  'Orçamento enviado',       '#c99a0a'),
    (3, 'visita',     'Visitação',               '#2b8bd6'),
    (4, 'contrato',   'Contrato / Aditivo',      '#e8631a'),
    (5, 'assinado',   'Contrato assinado',       '#1f9d55'),
    (6, 'check',      'Check-in realizados',     '#0f7a3f'),
]
def norm(s):
    s = unicodedata.normalize('NFD', str(s or '')).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'\s+', ' ', s).strip()

def titulo_de(txt):
    t = norm(txt)
    if t.startswith('clientes perdidos'): return 'perdidos'
    for i, k, n, c in ETAPAS:
        if k == 'radar' and t.startswith('clientes no radar'): return k
        if k == 'solicit' and t.startswith('solicitacao'): return k
        if k == 'orcamento' and t.startswith('orcamento enviado'): return k
        if k == 'visita' and t.startswith('visita'): return k
        if k == 'contrato' and t.startswith('contrato') and 'assinado' not in t: return k
        if k == 'assinado' and t.startswith('contrato assinado'): return k
        if k == 'check' and t.startswith('check'): return k
    return None

def pessoas(txt):
    """Número de pessoas a partir do texto (veja aba Critérios)."""
    if txt is None: return None
    if isinstance(txt, (int, float)): return int(txt)
    s = norm(txt)
    m = re.match(r'^(\d+)\s*a\s*(\d+)', s)
    if m: return int(m.group(1))
    nums = [int(x) for x in re.findall(r'\d+', s)]
    return sum(nums) if nums else None

def fdata(v):
    if isinstance(v, (dt.datetime, dt.date)):
        d = v.date() if isinstance(v, dt.datetime) else v
        return d.strftime('%d/%m/%Y'), d.isoformat()
    return (str(v).strip() if v not in (None, '') else ''), ''

def txt(v):
    if v is None: return ''
    if isinstance(v, (dt.datetime, dt.date)): return fdata(v)[0]
    return str(v).strip()

def grupo_motivo(m):
    t = norm(m)
    if 'sem retorno' in t: return 'Sem retorno'
    if 'preco' in t: return 'Preço'
    if 'fechou' in t: return 'Fechou com concorrente'
    if 'desalinhamento' in t or 'desacordo' in t: return 'Desalinhamento / desacordo comercial'
    return 'Demanda parada ou sem demanda'

def ler(caminho, aba=None):
    wb = openpyxl.load_workbook(caminho, data_only=True)
    if aba:
        if aba not in wb.sheetnames: sys.exit(f'ERRO: aba "{aba}" nao encontrada. Abas existentes: {wb.sheetnames}')
        ws = wb[aba]
    else:
        ws = wb.active
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    # localiza títulos na coluna B (índice 1), seguidos de linha de cabeçalho
    marcas = []
    for i, r in enumerate(rows):
        b = r[1] if len(r) > 1 else None
        k = titulo_de(b) if b else None
        if k and i + 1 < len(rows) and norm(rows[i+1][1]) in ('cliente', 'motivo'):
            marcas.append((i, k))
    blocos = {}
    for n, (i, k) in enumerate(marcas):
        fim = marcas[n+1][0] if n + 1 < len(marcas) else len(rows)
        cab = [norm(c) for c in rows[i+1]]
        blocos[k] = (cab, rows[i+2:fim])
    return blocos

def col(cab, *nomes, fallback=None):
    for j, c in enumerate(cab):
        if any(c.startswith(n) for n in nomes): return j
    return fallback


def baixar(url):
    """Baixa a planilha do link de compartilhamento (SharePoint 'qualquer pessoa com o link')."""
    if not url or 'COLE_AQUI' in url: sys.exit('ERRO: preencha planilha_url no config.json (ou o segredo PLANILHA_URL).')
    p = urllib.parse.urlparse(url); q = urllib.parse.parse_qs(p.query); q['download'] = ['1']
    url = urllib.parse.urlunparse(p._replace(query=urllib.parse.urlencode(q, doseq=True)))
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        dados = urllib.request.urlopen(req, timeout=60).read()
    except Exception as e:
        sys.exit(f'ERRO ao baixar a planilha: {e}')
    if dados[:2] != b'PK':
        sys.exit('ERRO: o link nao devolveu uma planilha (provavelmente pede login ou o compartilhamento mudou). Painel nao atualizado.')
    return io.BytesIO(dados)

def validar(blocos):
    obrig = ['solicit', 'orcamento', 'contrato', 'assinado', 'check', 'perdidos']
    falta = [k for k in obrig if k not in blocos]
    if falta: sys.exit(f'ERRO: blocos nao encontrados na planilha: {falta}. Confira os titulos na coluna B. Painel nao atualizado.')
    if sum(len(v[1]) for v in blocos.values()) < 5: sys.exit('ERRO: planilha com poucas linhas. Painel nao atualizado.')

def agora_br():
    return dt.datetime.now(ZoneInfo('America/Sao_Paulo'))

def main():
    cfg = json.load(open(os.path.join(PASTA, 'config.json'), encoding='utf-8'))
    if len(sys.argv) > 1:
        arq = sys.argv[1]
    else:
        arq = baixar(os.environ.get('PLANILHA_URL') or cfg.get('planilha_url', ''))
    blocos = ler(arq, cfg.get('aba'))
    validar(blocos)
    etapas, tarefas, alertas = [], [], []
    for i, k, nome, cor in ETAPAS:
        linhas = []
        if k in blocos:
            cab, dados = blocos[k]
            jc, jq, jt, jl, jd = (col(cab, 'cliente', fallback=1), col(cab, 'qtd', fallback=2), col(cab, 'tempo', fallback=3),
                                  col(cab, 'local', fallback=4), col(cab, 'data', fallback=5))
            ja = col(cab, 'acao', 'observ', fallback=6)
            vistos = set()
            for r in dados:
                g = lambda j: r[j] if j is not None and j < len(r) else None
                cli = txt(g(jc))
                if not cli: continue
                if norm(cli).startswith('fazer busca'):
                    tarefas.append({'cliente': 'Tarefa', 'qtd': '', 'acao': (txt(g(ja)) + ' — ' + cli) if txt(g(ja)) else cli}); continue
                d_txt, d_iso = fdata(g(jd))
                chave = (norm(cli), norm(txt(g(jq))))
                dup = chave in vistos and chave[1] != ''
                vistos.add(chave)
                linhas.append(dict(cliente=cli, qtd=txt(g(jq)), pessoas=pessoas(g(jq)), tempo=txt(g(jt)), local=txt(g(jl)),
                                   data=d_txt, data_iso=d_iso, acao=txt(g(ja)), dup=dup))
                if dup: alertas.append(f'Etapa {i} ({nome}): "{cli}" aparece repetido com a mesma quantidade. Conferir se é duplicidade na planilha.')
        etapas.append(dict(id=i, nome=nome, cor=cor, n=len(linhas),
                           vagas=sum(r['pessoas'] or 0 for r in linhas if not r['dup']), linhas=linhas))
    E = {e['id']: e for e in etapas}
    # perdidos
    perd, ult = [], ''
    if 'perdidos' in blocos:
        cab, dados = blocos['perdidos']
        jm, jc, jq = col(cab, 'motivo', fallback=1), col(cab, 'cliente', fallback=2), col(cab, 'qtd', fallback=3)
        for r in dados:
            cli = txt(r[jc]) if jc < len(r) else ''
            if not cli: continue
            m = txt(r[jm]) if jm < len(r) else ''
            if m: ult = m
            q = r[jq] if jq < len(r) else None
            perd.append(dict(motivo=m or ult, grupo=grupo_motivo(m or ult), cliente=cli, qtd=txt(q), pessoas=pessoas(q) or 0))
    grupos = {}
    for p in perd:
        g = grupos.setdefault(p['grupo'], dict(grupo=p['grupo'], n=0, pessoas=0)); g['n'] += 1; g['pessoas'] += p['pessoas']
    grupos = sorted(grupos.values(), key=lambda x: -x['n'])
    # KPIs
    neg = [E[i] for i in (1, 2, 3, 4)]
    fech = E[5]['n'] + E[6]['n']
    sem = sum(1 for r in E[0]['linhas'] if not r['qtd'] and not r['tempo'])
    kpi = dict(neg_n=sum(e['n'] for e in neg), neg_vagas=sum(e['vagas'] for e in neg),
               contr_n=E[4]['n'], contr_vagas=E[4]['vagas'], chk_n=E[6]['n'], chk_vagas=E[6]['vagas'],
               radar_n=E[0]['n'], radar_sem=sem, fech_n=fech, perd_n=len(perd),
               conv=round(100 * fech / (fech + len(perd))) if fech + len(perd) else 0)
    # alertas extras
    for e in etapas:
        if e['id'] in (1, 2, 3, 4):
            s = sum(1 for r in e['linhas'] if not r['qtd'])
            if s: alertas.append(f'Etapa {e["id"]} ({e["nome"]}): {s} cliente(s) sem quantidade de pessoas.')
    # ---------- oportunidades ----------
    hoje = agora_br().date().isoformat()
    def it(e, r, det): return dict(cliente=r['cliente'], etapa=e['nome'], qtd=r['qtd'], detalhe=det)
    def card(titulo, texto, itens): return dict(titulo=titulo, texto=texto, n=len(itens),
        pessoas=sum(i.get('p', 0) for i in itens), itens=[{k: v for k, v in i.items() if k != 'p'} for i in itens])
    def com_p(e, r, det): d = it(e, r, det); d['p'] = 0 if r['dup'] else (r['pessoas'] or 0); return d
    PEND = ('recontato', 'entrar em contato', 'mandar mensagem', 'irá mandar', 'contato daqui', 'entender com')
    ops = []
    ops.append(card('Fechar contrato: etapa Contrato / Aditivo', 'Já passaram por orçamento e estão em finalização. É o ganho mais rápido.',
        [com_p(E[4], r, r['local'] + (' · entrada ' + r['data'] if r['data'] else '')) for r in E[4]['linhas']]))
    ops.append(card('Contatos a fazer', 'Linhas cuja Ação pede recontato ou mensagem (radar, solicitação, orçamento e visitação). O responsável aparece no texto.',
        [com_p(e, r, r['acao']) for e in etapas if e['id'] in (0, 1, 2, 3) for r in e['linhas'] if any(x in norm(r['acao']) for x in PEND)]))
    ops.append(card('Orçamento enviado sem resposta', 'Clientes da etapa 2, com a maior quantidade de pessoas primeiro. Quanto mais tempo parado, menor a chance de fechar.',
        [com_p(E[2], r, ' · '.join(x for x in (r['tempo'], r['local'], r['acao']) if x)) for r in sorted(E[2]['linhas'], key=lambda r: -(r['pessoas'] or 0))]))
    ops.append(card('Sem dados para orçar', 'Clientes do radar, solicitação, orçamento e visitação sem quantidade de pessoas. Sem isso não há orçamento nem resposta.',
        [com_p(e, r, r['acao']) for e in etapas if e['id'] in (0, 1, 2, 3) for r in e['linhas'] if r['pessoas'] is None]))
    ops.append(card('Entradas com data vencida', 'Data de entrada anterior a hoje em etapa que não é check-in. Conferir se entrou, mudou ou caiu.',
        [com_p(e, r, 'entrada prevista ' + r['data']) for e in etapas if e['id'] in (0, 1, 2, 3, 4, 5) for r in e['linhas'] if r['data_iso'] and r['data_iso'] < hoje]))
    ops.append(card('Reativar: perdidos por "sem retorno"', 'Uma última tentativa de contato costuma valer mais do que abrir cliente novo.',
        [dict(cliente=q['cliente'], etapa='Perdido', qtd=q['qtd'], detalhe=q['motivo'], p=q['pessoas']) for q in perd if q['grupo'] == 'Sem retorno']))
    # ---------- tendências (textos recalculados a cada rodada) ----------
    tend = []
    e2 = [r for r in E[2]['linhas'] if not r['dup'] and r['pessoas']]
    if e2 and E[2]['vagas']:
        top = max(e2, key=lambda r: r['pessoas']); pc = round(100 * top['pessoas'] / E[2]['vagas'])
        tend.append(dict(titulo='O orçamento enviado depende de poucos clientes',
            texto=f"<b>{top['cliente']}</b> responde por <b>{top['pessoas']}</b> das {E[2]['vagas']} pessoas em orçamento enviado (<b>{pc}%</b>). Sem ele, o volume em orçamento cai para {E[2]['vagas'] - top['pessoas']} pessoas."))
    if perd:
        g0 = grupos[0]; sr = next((g for g in grupos if g['grupo'] == 'Sem retorno'), dict(n=0)); pr = next((g for g in grupos if g['grupo'] == 'Preço'), dict(n=0))
        tend.append(dict(titulo='Perde-se mais por falta de retorno e demanda do que por preço',
            texto=f"De <b>{len(perd)}</b> perdidos, <b>{sr['n']}</b> foram por sem retorno e só <b>{pr['n']}</b> por preço. O maior grupo é “{g0['grupo']}” (<b>{g0['n']}</b>). A conversão está em <b>{kpi['conv']}%</b> ({fech} fechados contra {len(perd)} perdidos)."))
    tend.append(dict(titulo='O radar ainda está sem informação',
        texto=f"<b>{sem}</b> dos <b>{E[0]['n']}</b> clientes do radar não têm pessoas nem prazo. O radar soma <b>{E[0]['vagas']}</b> pessoas só com os que informaram."))
    em_op, ass, ctr = E[6]['vagas'], E[5]['vagas'], E[4]['vagas']
    if em_op:
        tend.append(dict(titulo='O que já está garantido para entrar',
            texto=f"Hoje há <b>{em_op}</b> pessoas em operação (check-in). Com contrato assinado vêm mais <b>{ass}</b>, e em Contrato/Aditivo mais <b>{ctr}</b>. Se tudo fechar, a operação vai para <b>{em_op + ass + ctr}</b> pessoas (<b>+{round(100 * (ass + ctr) / em_op)}%</b>)."))
    # ---------- histórico ----------
    hp = os.path.join(PASTA, 'historico.json')
    hist = json.load(open(hp, encoding='utf-8')) if os.path.exists(hp) else []
    agora0 = agora_br()
    snap = dict(data=agora0.date().isoformat(), hora=agora0.strftime('%d/%m/%Y %H:%M'),
                etapas={str(e['id']): dict(n=e['n'], vagas=e['vagas']) for e in etapas}, perdidos=len(perd))
    hist = ([h for h in hist if h['data'] != snap['data']] + [snap])[-400:]
    json.dump(hist, open(hp, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    if len(hist) >= 2:
        a, b = hist[-2], hist[-1]
        partes = []
        for e in etapas:
            d = b['etapas'][str(e['id'])]['n'] - a['etapas'].get(str(e['id']), {}).get('n', 0)
            if d: partes.append(f"{e['nome']} {'+' if d > 0 else ''}{d}")
        dp = b['perdidos'] - a['perdidos']
        if dp: partes.append(f"Perdidos {'+' if dp > 0 else ''}{dp}")
        tend.append(dict(titulo=f"O que mudou desde {a['hora']}",
            texto=('Variação de clientes por etapa: <b>' + '; '.join(partes) + '</b>.') if partes else 'Nenhuma mudança na quantidade de clientes por etapa.'))
    else:
        tend.append(dict(titulo='Evolução ao longo do tempo',
            texto=f"O histórico começou em <b>{snap['hora']}</b>. A comparação com a leitura anterior aparece a partir da próxima atualização com data diferente."))
    agora = agora_br()
    data = dict(arquivo='planilha na nuvem', gerado=agora.strftime('%d/%m/%Y %H:%M'), hoje_iso=agora.date().isoformat(),
                etapas=etapas, tarefas=tarefas, alertas=alertas, kpi=kpi,
                perdidos=dict(total=len(perd), linhas=perd, grupos=grupos), oportunidades=ops, tendencias=tend, historico=hist)
    html = open(os.path.join(PASTA, 'modelo.html'), encoding='utf-8').read()
    html = html.replace('__DATA__', json.dumps(data, ensure_ascii=False))
    os.makedirs(os.path.join(PASTA, 'site'), exist_ok=True)
    saida = os.path.join(PASTA, 'site', 'index.html')
    open(saida, 'w', encoding='utf-8').write(html)
    print('Painel gerado:', saida, '|', agora.strftime('%d/%m/%Y %H:%M'), '(Brasilia)')

if __name__ == '__main__':
    main()
