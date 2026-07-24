#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cobre Seu Político — Coletor de contatos oficiais de deputados e senadores
==========================================================================

Coleta os contatos INSTITUCIONAIS (públicos, publicados pelos próprios órgãos
nos portais de Dados Abertos) de todos os parlamentares federais em exercício
e grava numa planilha do Google Sheets.

Fontes oficiais:
  - Câmara:  https://dadosabertos.camara.leg.br/api/v2
  - Senado:  https://legis.senado.leg.br/dadosabertos

IMPORTANTE (LGPD / ética):
  Este script coleta APENAS dados de contato funcional já tornados públicos
  pelos próprios órgãos (e-mail de gabinete @camara.leg.br / @senado.leg.br,
  telefone e endereço do gabinete). Não coleta celular pessoal, e-mail
  particular nem endereço residencial.

Autor: (seu nome)
Licença de uso dos dados: domínio público (Dados Abertos gov.br)
"""

import os
import sys
import time
import json
import logging
from datetime import datetime, timezone

import requests
import gspread
from google.oauth2.service_account import Credentials

# --------------------------------------------------------------------------- #
# Configuração
# --------------------------------------------------------------------------- #

# Nome da planilha no Google Drive (será criada/aberta por esse nome)
NOME_PLANILHA = os.environ.get("NOME_PLANILHA", "Cobre Seu Político - Contatos")

# Caminho do arquivo de credenciais da conta de serviço.
# No GitHub Actions, o conteúdo vem do Secret e é escrito neste arquivo.
CREDS_PATH = os.environ.get("GOOGLE_CREDS_PATH", "credenciais.json")

# Pausa entre chamadas à API da Câmara (respeitar o servidor / rate limit)
PAUSA_SEGUNDOS = float(os.environ.get("PAUSA_SEGUNDOS", "0.15"))

# Timeout padrão das requisições HTTP
TIMEOUT = 30

CAMARA_BASE = "https://dadosabertos.camara.leg.br/api/v2"
SENADO_BASE = "https://legis.senado.leg.br/dadosabertos"

HEADERS = {
    "Accept": "application/json",
    "User-Agent": "CobreSeuPolitico/1.0 (projeto cívico de transparência)",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("cobre-politico")


# --------------------------------------------------------------------------- #
# Utilidades HTTP
# --------------------------------------------------------------------------- #

def get_json(url, params=None, tentativas=3):
    """GET com retry simples e backoff. Retorna dict/list ou None."""
    for i in range(1, tentativas + 1):
        try:
            r = requests.get(url, params=params, headers=HEADERS, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json()
        except requests.RequestException as e:
            espera = 2 * i
            log.warning("Falha em %s (tentativa %d/%d): %s — aguardando %ds",
                        url, i, tentativas, e, espera)
            time.sleep(espera)
    log.error("Desisti de %s após %d tentativas", url, tentativas)
    return None


# --------------------------------------------------------------------------- #
# Coleta — Câmara dos Deputados
# --------------------------------------------------------------------------- #

def coletar_deputados():
    """
    Retorna lista de dicts com os deputados em exercício e seus contatos
    de gabinete. Faz uma chamada de listagem paginada + 1 chamada de
    detalhe por deputado (para pegar telefone e gabinete).
    """
    log.info("=== CÂMARA: listando deputados em exercício ===")
    deputados = []
    pagina = 1
    while True:
        dados = get_json(
            f"{CAMARA_BASE}/deputados",
            params={"ordem": "ASC", "ordenarPor": "nome",
                    "itens": 100, "pagina": pagina},
        )
        if not dados or not dados.get("dados"):
            break
        deputados.extend(dados["dados"])
        # A API traz links de navegação; paramos quando não há "next"
        links = {l["rel"]: l["href"] for l in dados.get("links", [])}
        if "next" not in links:
            break
        pagina += 1

    log.info("Encontrados %d deputados. Buscando detalhes de contato...",
             len(deputados))

    resultado = []
    total = len(deputados)
    for n, dep in enumerate(deputados, 1):
        dep_id = dep["id"]
        detalhe = get_json(f"{CAMARA_BASE}/deputados/{dep_id}")
        gab = {}
        email = dep.get("email") or ""
        if detalhe and detalhe.get("dados"):
            d = detalhe["dados"]
            email = d.get("email") or email
            gab = (d.get("ultimoStatus", {}) or {}).get("gabinete", {}) or {}

        resultado.append({
            "casa": "Câmara",
            "nome": dep.get("nome", ""),
            "partido": dep.get("siglaPartido", ""),
            "uf": dep.get("siglaUf", ""),
            "email": email,
            "telefone": _formatar_telefone_camara(gab),
            "gabinete": _formatar_gabinete(gab),
            "endereco": "Câmara dos Deputados - Praça dos Três Poderes, Brasília/DF - 70160-900",
            "url_perfil": f"https://www.camara.leg.br/deputados/{dep_id}",
            "foto": dep.get("urlFoto", ""),
        })

        if n % 50 == 0 or n == total:
            log.info("  ... %d/%d deputados processados", n, total)
        time.sleep(PAUSA_SEGUNDOS)

    return resultado


def _formatar_telefone_camara(gab):
    tel = (gab or {}).get("telefone", "")
    if tel:
        return f"(61) 3215-{tel}" if len(str(tel)) == 4 else str(tel)
    return ""


def _formatar_gabinete(gab):
    if not gab:
        return ""
    predio = gab.get("predio", "")
    andar = gab.get("andar", "")
    sala = gab.get("sala", "")
    partes = []
    if predio:
        partes.append(f"Anexo {predio}")
    if andar:
        partes.append(f"{andar}º andar")
    if sala:
        partes.append(f"gab. {sala}")
    return ", ".join(partes)


# --------------------------------------------------------------------------- #
# Coleta — Senado Federal
# --------------------------------------------------------------------------- #

def coletar_senadores():
    """
    Retorna lista de dicts com os senadores em exercício e seus contatos
    institucionais. A lista atual já traz e-mail e telefone do gabinete.
    """
    log.info("=== SENADO: listando senadores em exercício ===")
    dados = get_json(f"{SENADO_BASE}/senador/lista/atual.json")
    if not dados:
        log.error("Não consegui obter a lista de senadores.")
        return []

    # Navegação do JSON do Senado (estrutura aninhada)
    try:
        parlamentares = (dados["ListaParlamentarEmExercicio"]
                              ["Parlamentares"]["Parlamentar"])
    except (KeyError, TypeError):
        log.error("Estrutura inesperada no JSON do Senado.")
        return []

    resultado = []
    for p in parlamentares:
        ident = p.get("IdentificacaoParlamentar", {}) or {}
        resultado.append({
            "casa": "Senado",
            "nome": ident.get("NomeParlamentar", ""),
            "partido": ident.get("SiglaPartidoParlamentar", ""),
            "uf": ident.get("UfParlamentar", ""),
            "email": ident.get("EmailParlamentar", ""),
            "telefone": _telefone_senador(p),
            "gabinete": _gabinete_senador(p),
            "endereco": "Senado Federal - Praça dos Três Poderes, Brasília/DF - 70165-900",
            "url_perfil": ident.get("UrlPaginaParlamentar", ""),
            "foto": ident.get("UrlFotoParlamentar", ""),
        })

    log.info("Encontrados %d senadores.", len(resultado))
    return resultado


def _telefone_senador(p):
    # Alguns registros trazem telefones em Telefones/Telefone
    tels = (p.get("Telefones", {}) or {}).get("Telefone", [])
    if isinstance(tels, dict):
        tels = [tels]
    for t in tels:
        num = t.get("NumeroTelefone", "")
        if num:
            return num
    return ""


def _gabinete_senador(p):
    # O Senado nem sempre traz o gabinete estruturado na lista atual
    return ""


# --------------------------------------------------------------------------- #
# Google Sheets
# --------------------------------------------------------------------------- #

def abrir_planilha():
    escopos = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(CREDS_PATH, scopes=escopos)
    cliente = gspread.authorize(creds)

    try:
        planilha = cliente.open(NOME_PLANILHA)
        log.info("Planilha '%s' aberta.", NOME_PLANILHA)
    except gspread.SpreadsheetNotFound:
        planilha = cliente.create(NOME_PLANILHA)
        log.info("Planilha '%s' criada.", NOME_PLANILHA)
        # Se você quiser abrir para o público automaticamente, descomente:
        # planilha.share(None, perm_type="anyone", role="reader")
    return planilha


def escrever_dados(planilha, linhas):
    aba = planilha.sheet1
    aba.clear()

    cabecalho = [
        "Casa", "Nome", "Partido", "UF", "E-mail oficial",
        "Telefone gabinete", "Gabinete", "Endereço", "Página oficial",
    ]
    matriz = [cabecalho]
    for r in linhas:
        matriz.append([
            r["casa"], r["nome"], r["partido"], r["uf"], r["email"],
            r["telefone"], r["gabinete"], r["endereco"], r["url_perfil"],
        ])

    aba.update(matriz, "A1")

    # Formatação leve: cabeçalho em negrito e congelado
    aba.freeze(rows=1)
    aba.format("A1:I1", {
        "textFormat": {"bold": True},
        "backgroundColor": {"red": 0.15, "green": 0.20, "blue": 0.30},
        "horizontalAlignment": "CENTER",
    })
    aba.format("A1:I1", {"textFormat": {
        "bold": True,
        "foregroundColor": {"red": 1, "green": 1, "blue": 1},
    }})

    # Aba de metadados (data da última atualização)
    _escrever_metadados(planilha, len(linhas))

    log.info("Planilha atualizada com %d parlamentares.", len(linhas))


def _escrever_metadados(planilha, total):
    try:
        meta = planilha.worksheet("Info")
    except gspread.WorksheetNotFound:
        meta = planilha.add_worksheet("Info", rows=10, cols=2)
    agora = datetime.now(timezone.utc).astimezone().strftime("%d/%m/%Y %H:%M")
    meta.clear()
    meta.update([
        ["Última atualização", agora],
        ["Total de parlamentares", total],
        ["Fonte", "Dados Abertos - Câmara e Senado"],
        ["Observação", "Somente contatos institucionais públicos (LAI)"],
    ], "A1")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def main():
    inicio = time.time()
    log.info("Iniciando coleta de contatos parlamentares...")

    deputados = coletar_deputados()
    senadores = coletar_senadores()

    todos = deputados + senadores
    # Ordena por UF depois por nome, para leitura amigável
    todos.sort(key=lambda r: (r["uf"], r["casa"], r["nome"]))

    if not todos:
        log.error("Nenhum dado coletado. Abortando sem tocar na planilha.")
        sys.exit(1)

    planilha = abrir_planilha()
    escrever_dados(planilha, todos)

    dur = time.time() - inicio
    log.info("Concluído em %.1fs. %d deputados + %d senadores.",
             dur, len(deputados), len(senadores))
    log.info("URL da planilha: %s", planilha.url)


if __name__ == "__main__":
    main()
