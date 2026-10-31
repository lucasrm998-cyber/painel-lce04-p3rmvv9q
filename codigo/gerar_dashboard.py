# -*- coding: utf-8 -*-
"""Gera o Dashboard Shopee Inventário a partir da planilha do Google Sheets.

Uso: dois cliques em 'atualizar.bat' (ou: python gerar_dashboard.py)

Todo dia basta preencher a linha do dia no Google Sheets e rodar de novo. A planilha é
baixada e salva em 'Dados Inventário Shopee.xlsx' (cópia local); sem internet, usa essa cópia.
O resultado é 'Dashboard Shopee Inventário.html'.
Dias sem nada lançado depois do último dia preenchido são ignorados.
"""
import base64
import io
import json
import os
import re
import sys
import unicodedata
import urllib.request
from datetime import datetime
from pathlib import Path

import pandas as pd

PASTA = Path(__file__).resolve().parent
ARQUIVO_DADOS = PASTA / "Dados Inventário Shopee.xlsx"
# Planilha compartilhada como "Qualquer pessoa com o link" (Leitor). O ID não fica no código
# (o script vai para o GitHub público): vem do segredo ID_SHEETS no GitHub ou de sheets_id.txt no PC.
ARQUIVO_ID_SHEETS = PASTA / "sheets_id.txt"
ID_SHEETS = os.environ.get("ID_SHEETS") or (
    ARQUIVO_ID_SHEETS.read_text(encoding="utf-8").strip() if ARQUIVO_ID_SHEETS.exists() else "")
URL_SHEETS = f"https://docs.google.com/spreadsheets/d/{ID_SHEETS}/export?format=xlsx"
MODELO = PASTA / "modelo" / "dashboard_modelo.html"
SAIDA = PASTA / "Dashboard Shopee Inventário.html"
PASTA_IMAGENS = PASTA / "imagens"

NOME_HUB = "LCE-04"

# Farol diário (total de ocorrências do dia) — mesma regra do Data Studio
LIMITE_CRITICO = 500   # acima disso: Crítico
LIMITE_ATENCAO = 300   # acima disso: Atenção; até isso: Normal

# Famílias: agrupam as ocorrências nos gráficos empilhados (cor fixa por família)
FAMILIAS = [
    ("sobras", "Sobras"),
    ("avarias", "Avarias"),
    ("onhold", "ON Holds"),
    ("expedicao", "Expedição & operação"),
    ("risco", "Risco de perda"),
]

# Cabeçalho do Excel (sem acento, maiúsculo, com _) -> (nome, nome curto, família)
CATEGORIAS = {
    "SOBRAS_LM": ("Sobras LM", "Sobras LM", "sobras"),
    "SOBRAS_XPT": ("Sobras XPT", "Sobras XPT", "sobras"),
    "FORA_DE_ROTA_RETIRADOS_NA_BANCADA": ("Fora de rota (bancada)", "Fora de rota", "expedicao"),
    "VOLUMOSOS_ESQUECIDOS": ("Volumosos esquecidos", "Volumosos", "expedicao"),
    "ERROS_OPERACIONAIS": ("Erros operacionais", "Erros op.", "expedicao"),
    "ON_HOLDS": ("ON Holds", "ON Holds", "onhold"),
    "PROJECAO_PARA_LOST": ("Projeção para Lost", "Proj. Lost", "risco"),
    "NAO_COUBE_NO_VEICULO": ("Não coube no veículo", "Não coube", "expedicao"),
    "AVARIAS_HUB": ("Avarias HUB", "Av. HUB", "avarias"),
    "AVARIAS_EXTERNAS": ("Avarias externas", "Av. externas", "avarias"),
}

# Colunas de apoio de quem preenche a planilha: não entram no dashboard
COLUNAS_IGNORADAS = {"COLUNA_DE_CONFERENCIA"}

# Insucesso de Auto Add: fica fora do total de ocorrências e do farol (aba própria).
# Cada insucesso tem um Status. Vários no mesmo dia vêm na mesma célula separados por "|"
# (ex.: "Não Coube no Veículo | Não Deu Tempo"); linhas sem DATA logo abaixo de um dia
# também contam para esse dia.
COL_INSUCESSO = "AUTO_ADD"
NOMES_ANTIGOS = {"INSUCESSO_DE_AUTO_ADD": COL_INSUCESSO}
COL_STATUS = "STATUS"
STATUS_AUTO_ADD = [  # a ordem define a ordem e a cor no dashboard
    "Não Coube no Veículo",
    "Não Deu Tempo",
    "Removidos Como Fora de Rota",
    "Erro Operacional",
    "Realocado",
]

# Ocorrências Fleet: driver que devolveu pedidos e a quantidade, na mesma ordem.
# Ex.: "LEANDRO SOBREIRA SILVA | OSMAR HENRIQUE" com Qtd "11 | 15".
COL_DRIVER = "DRIVER_OFENSOR_DO_DIA"
COL_QTD_DRIVER = "QTD"

# Flutuantes XPT: quantidade e base, na mesma ordem. Ex.: "3 | 5" e "XPT - PICOS | XPT - TAUÁ".
# Também aceita só o total e uma base por item. Ex.: "3" e "XPT - PICOS | XPT - PICOS | XPT - TAUÁ".
COL_FLUTUANTES = "FLUTUANTES_XPT"
COL_BASE = "BASE_XPT"
BASES_XPT = ["Araripina", "Barro", "Campos Sales", "Itaporanga", "Ouricuri", "Picos", "Pombal", "Tauá"]

COLUNAS_TEXTO = [COL_STATUS, COL_DRIVER, COL_QTD_DRIVER, COL_FLUTUANTES, COL_BASE]


def normalizar(texto):
    """'Sobras _XPT ' -> 'SOBRAS_XPT' (sem acento, maiúsculo, com um _)."""
    texto = unicodedata.normalize("NFKD", str(texto).strip().upper())
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"_+", "_", "_".join(texto.split()))


def ler_planilha(caminho):
    bruto = pd.read_excel(caminho, header=None)

    # Procura a linha do cabeçalho (a que tem a coluna DATA)
    linha_cab = next(
        (i for i in range(min(50, len(bruto)))
         if "DATA" in {normalizar(v) for v in bruto.iloc[i] if pd.notna(v)}),
        None,
    )
    if linha_cab is None:
        raise ValueError("Não encontrei o cabeçalho (coluna DATA) na planilha.")

    df = bruto.iloc[linha_cab + 1:].copy()
    df.columns = [normalizar(c) if pd.notna(c) else f"_vazia{i}" for i, c in enumerate(bruto.iloc[linha_cab])]
    df = df.loc[:, [c for c in df.columns if not c.startswith("_vazia") and c not in COLUNAS_IGNORADAS]]
    df = df.rename(columns=NOMES_ANTIGOS)

    faltando = [CATEGORIAS[k][0] for k in CATEGORIAS if k not in df.columns]
    if faltando:
        print(f"  Aviso: colunas não encontradas (ficam zeradas): {', '.join(faltando)}")
    if COL_INSUCESSO not in df.columns:
        print("  Aviso: coluna 'Auto Add' não encontrada (fica zerada).")
    extras = [c for c in df.columns if c not in {"DATA", COL_INSUCESSO, *COLUNAS_TEXTO, *CATEGORIAS}]
    if extras:
        print(f"  Aviso: colunas novas ignoradas (cadastre em CATEGORIAS): {', '.join(extras)}")

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
    status = ler_status(df)
    fleet = ler_pares(df, COL_DRIVER, COL_QTD_DRIVER, "Driver ofensor", nome_driver)
    flut = ler_pares(df, COL_BASE, COL_FLUTUANTES, "Flutuantes XPT", nome_base)
    sem_data = df["DATA"].isna() & df.drop(columns=["DATA", *COLUNAS_TEXTO], errors="ignore").notna().any(axis=1)
    if sem_data.any():
        print(f"  Aviso: {sem_data.sum()} linha(s) sem data válida foram ignoradas.")
    df = df[df["DATA"].notna()]

    numericas = [*CATEGORIAS, COL_INSUCESSO]
    for k in numericas:
        df[k] = pd.to_numeric(df[k], errors="coerce").fillna(0) if k in df else 0
    df = df[["DATA", *numericas]]

    repetidas = df["DATA"].duplicated(keep=False)
    if repetidas.any():
        print(f"  Aviso: {df.loc[repetidas, 'DATA'].nunique()} data(s) repetida(s) — os valores foram somados.")
    df = df.groupby("DATA", as_index=False).sum().sort_values("DATA")

    # Dias futuros já criados na planilha vêm zerados: corta depois do último dia lançado
    total = df[list(CATEGORIAS)].sum(axis=1)
    if not (total > 0).any():
        raise ValueError("A planilha não tem nenhuma ocorrência lançada.")
    lancado = ((total > 0) | (df[COL_INSUCESSO] > 0) | df["DATA"].isin(status.index)
               | df["DATA"].isin(fleet["DATA"]) | df["DATA"].isin(flut["DATA"]))
    df = df[df["DATA"] <= df.loc[lancado, "DATA"].max()]

    # Dias que faltam no meio do período entram zerados (o calendário fica contínuo)
    dias = pd.date_range(df["DATA"].min(), df["DATA"].max(), freq="D")
    faltam = len(dias) - len(df)
    if faltam:
        print(f"  Aviso: {faltam} dia(s) sem linha na planilha entraram como zero.")
    no_periodo = lambda t: t[t["DATA"].isin(dias)]
    return (df.set_index("DATA").reindex(dias, fill_value=0), status.reindex(dias, fill_value=0),
            no_periodo(fleet), no_periodo(flut))


def partes(valor):
    """'A | B' -> ['A', 'B']; vazio -> []."""
    if pd.isna(valor):
        return []
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return [p.strip() for p in str(valor).split("|") if p.strip()]


def nome_driver(texto):
    return " ".join(texto.upper().split())


def nome_base(texto):
    """'XPT - PICOS' -> 'Picos' (usa a grafia de BASES_XPT quando reconhece)."""
    sem_xpt = re.sub(r"^\s*XPT\s*[-–:]?\s*", "", texto, flags=re.I)
    conhecidas = {normalizar(b): b for b in BASES_XPT}
    return conhecidas.get(normalizar(sem_xpt), " ".join(sem_xpt.split()).title())


def ler_pares(df, col_nome, col_qtd, rotulo, limpa):
    """Lê colunas casadas por '|' (nome ↔ quantidade) -> DataFrame DATA, NOME, QTD."""
    vazio = pd.DataFrame(columns=["DATA", "NOME", "QTD"])
    if col_nome not in df.columns or col_qtd not in df.columns:
        print(f"  Aviso: colunas de '{rotulo}' não encontradas (aba fica vazia).")
        return vazio
    linhas = []
    for data, nomes, qtds in zip(df["DATA"].ffill(), df[col_nome], df[col_qtd]):
        nomes, qtds = partes(nomes), partes(qtds)
        qtds = [q for q in qtds if q not in {"0", "0.0"}] if not nomes else qtds
        if not nomes and not qtds:
            continue
        if pd.isna(data):
            print(f"  Aviso {rotulo}: lançamento acima da primeira data ignorado.")
            continue
        if len(qtds) == 1 and len(nomes) > 1:
            # Um nome por item e só o total na quantidade (ex.: 60 e "XPT - BARRO | XPT - BARRO | ..."):
            # cada nome conta 1
            total = pd.to_numeric(qtds[0].replace(",", "."), errors="coerce")
            if total != len(nomes):
                print(f"  Aviso {rotulo} {data:%d/%m}: total {qtds[0]} e {len(nomes)} nome(s) lançado(s) — "
                      "contei 1 por nome; confira a planilha.")
            linhas += [(data, limpa(n), 1.0) for n in nomes]
            continue
        if len(nomes) != len(qtds):
            print(f"  Aviso {rotulo} {data:%d/%m}: {len(nomes)} nome(s) e {len(qtds)} quantidade(s) — confira a planilha.")
        for k in range(max(len(nomes), len(qtds))):
            nome = limpa(nomes[k]) if k < len(nomes) else "Não informado"
            qtd = pd.to_numeric(qtds[k].replace(",", "."), errors="coerce") if k < len(qtds) else 0
            if pd.isna(qtd):
                print(f"  Aviso {rotulo} {data:%d/%m}: quantidade '{qtds[k]}' não é número (entra 0).")
                qtd = 0
            linhas.append((data, nome, float(qtd)))
    return pd.DataFrame(linhas, columns=["DATA", "NOME", "QTD"]) if linhas else vazio


def ler_status(df):
    """Conta, por dia, quantos insucessos de Auto Add têm cada Status."""
    if COL_STATUS not in df.columns:
        print("  Aviso: coluna 'Status' não encontrada (status ficam zerados).")
        return pd.DataFrame(columns=STATUS_AUTO_ADD, dtype=int)
    st = pd.DataFrame({"DATA": df["DATA"].ffill(), "S": df[COL_STATUS].map(partes)}).explode("S")
    st = st[st["S"].notna()]
    sem_dia = st["DATA"].isna().sum()
    if sem_dia:
        print(f"  Aviso: {sem_dia} status acima da primeira data foram ignorados.")
    st = st[st["DATA"].notna()]

    conhecidos = {normalizar(s): s for s in STATUS_AUTO_ADD}
    st["S"] = st["S"].map(lambda s: conhecidos.get(normalizar(s), str(s).strip()))
    novos = sorted(set(st["S"]) - set(STATUS_AUTO_ADD))
    if novos:
        print(f"  Aviso: status fora da lista (entram no fim): {', '.join(novos)}")
    contagem = st.groupby(["DATA", "S"]).size().unstack(fill_value=0)
    return contagem.reindex(columns=[*STATUS_AUTO_ADD, *novos], fill_value=0).astype(int)


def baixar_planilha():
    """Baixa o Google Sheets e atualiza a cópia local. Sem internet, usa a cópia que já existe."""
    print("Baixando a planilha do Google Sheets ...")
    try:
        with urllib.request.urlopen(URL_SHEETS, timeout=60) as resp:
            conteudo = resp.read()
        if not conteudo.startswith(b"PK"):  # xlsx é um zip; outra coisa = página de login do Google
            raise ValueError("o Google não devolveu a planilha (confira se o link está como "
                             "'Qualquer pessoa com o link')")
    except Exception as erro:
        print(f"  Aviso: não consegui baixar ({erro}).")
        if not ARQUIVO_DADOS.exists():
            raise FileNotFoundError(f"Sem planilha online e sem cópia local: {ARQUIVO_DADOS}")
        print(f"  Usando a cópia local de {datetime.fromtimestamp(ARQUIVO_DADOS.stat().st_mtime):%d/%m/%Y %H:%M}.")
        return ARQUIVO_DADOS
    try:
        ARQUIVO_DADOS.write_bytes(conteudo)
        print("  Cópia local atualizada:", ARQUIVO_DADOS.name)
        return ARQUIVO_DADOS
    except PermissionError:
        print(f"  Aviso: '{ARQUIVO_DADOS.name}' está aberto no Excel; a cópia local não foi atualizada.")
        return io.BytesIO(conteudo)


def imagem(nome):
    """Usa imagens/<nome>.png|jpg|svg|webp se existir; senão, None."""
    tipos = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
             ".svg": "image/svg+xml", ".webp": "image/webp"}
    for ext, mime in tipos.items():
        arq = PASTA_IMAGENS / f"{nome}{ext}"
        if arq.exists():
            return f"data:{mime};base64,{base64.b64encode(arq.read_bytes()).decode()}"
    return None


def conferir_auto_add(df, status):
    """Avisa os dias em que a quantidade de status não bate com a de insucessos."""
    lancados = status.sum(axis=1)
    for dia in df.index[(df[COL_INSUCESSO] != lancados)]:
        ins, qtd = int(df.at[dia, COL_INSUCESSO]), int(lancados[dia])
        falta = f" — faltam {ins - qtd} status" if ins > qtd else " — status a mais"
        print(f"  Aviso Auto Add {dia:%d/%m}: {ins} insucesso(s) e {qtd} status lançado(s){falta}.")


def montar_dados(df, status, fleet, flut):
    numero = lambda v: int(v) if float(v).is_integer() else round(float(v), 2)
    pos = {d: i for i, d in enumerate(df.index)}
    pares = lambda t: [[pos[d], n, numero(q)] for d, n, q in t.itertuples(index=False)]
    novas = sorted(set(flut["NOME"]) - set(BASES_XPT))
    if novas:
        print(f"  Aviso Flutuantes XPT: base(s) fora da lista (entram no fim): {', '.join(novas)}")
    return {
        "auto": {
            "status": list(status.columns),
            "ins": [numero(v) for v in df[COL_INSUCESSO]],
            "st": status.values.tolist(),
        },
        "fleet": pares(fleet),
        "flut": pares(flut),
        "bases": [*BASES_XPT, *novas],
        "gerado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "hub": NOME_HUB,
        "limites": {"critico": LIMITE_CRITICO, "atencao": LIMITE_ATENCAO},
        "familias": [{"k": k, "nome": n} for k, n in FAMILIAS],
        "cats": [{"k": k, "nome": n, "curto": c, "fam": f} for k, (n, c, f) in CATEGORIAS.items()],
        "dias": [d.strftime("%Y-%m-%d") for d in df.index],
        "v": [[numero(x) for x in linha] for linha in df[list(CATEGORIAS)].itertuples(index=False)],
    }


def main():
    df, status, fleet, flut = ler_planilha(baixar_planilha())
    conferir_auto_add(df, status)
    dados = montar_dados(df, status, fleet, flut)

    modelo = MODELO.read_text(encoding="utf-8")
    dados_json = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = (
        modelo.replace("__DADOS__", dados_json)
        .replace("__LOGO__", json.dumps(imagem("logo_shopee")))
        .replace("__MASCOTE__", json.dumps(imagem("mascote")))
    )
    SAIDA.write_text(html, encoding="utf-8")

    # Resumo para conferir com o Data Studio
    ultimo = df.index.max()
    mes = df[df.index >= ultimo.replace(day=1)]
    total_mes = int(mes[list(CATEGORIAS)].sum().sum())
    print(f"  Período: {df.index.min():%d/%m/%Y} a {ultimo:%d/%m/%Y} ({len(df)} dias)")
    print(f"  Mês atual ({ultimo:%m/%Y}): {total_mes:,} ocorrências em {len(mes)} dia(s)".replace(",", "."))
    print(f"  Último dia ({ultimo:%d/%m}): {int(df.loc[ultimo, list(CATEGORIAS)].sum()):,} ocorrências".replace(",", "."))
    print(f"  Auto Add no mês: {int(mes[COL_INSUCESSO].sum())} insucesso(s), "
          f"{int(status[status.index >= ultimo.replace(day=1)].values.sum())} status lançado(s)")
    no_mes = lambda t: t[t["DATA"] >= ultimo.replace(day=1)]
    print(f"  Ocorrências Fleet no mês: {int(no_mes(fleet)['QTD'].sum())} devolução(ões), "
          f"{no_mes(fleet)['NOME'].nunique()} driver(s)")
    print(f"  Flutuantes XPT no mês: {int(no_mes(flut)['QTD'].sum())}")
    print("Dashboard gerado:", SAIDA.name)
    if "--abrir" in sys.argv:
        os.startfile(SAIDA)


if __name__ == "__main__":
    try:
        main()
    except Exception as erro:
        print("\nERRO:", erro)
        sys.exit(1)
