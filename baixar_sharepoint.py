"""
baixar_sharepoint.py — Baixa a planilha do SharePoint para pasta local.
Execute este script manualmente ou agende via Agendador de Tarefas do Windows.

Dependências: pip install msal requests
"""

import msal
import requests
import os
from pathlib import Path

# ─────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────

SHAREPOINT_URL = "https://animaeducacao.sharepoint.com"
SITE_PATH      = "/sites/InputdeNovasAdmisses"
FILE_ID        = "BD3809A8-52B2-4DB4-BBBC-A5953DB7AB36"

# Caminho local onde o arquivo será salvo
# O app.py vai ler daqui — mantenha o mesmo nome
PASTA_LOCAL    = Path(__file__).parent  # mesma pasta do projeto
NOME_LOCAL     = "Controle_Admissoes.xlsx"

CLIENT_ID   = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
AUTHORITY   = "https://login.microsoftonline.com/common"
SCOPES      = ["https://animaeducacao.sharepoint.com/.default"]
TOKEN_CACHE = str(Path(__file__).parent / "token_cache.json")

# ─────────────────────────────────────────────


def carregar_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE):
        cache.deserialize(open(TOKEN_CACHE, "r").read())
    return cache


def salvar_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE, "w") as f:
            f.write(cache.serialize())


def obter_token():
    cache = carregar_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    contas = app.get_accounts()
    if contas:
        resultado = app.acquire_token_silent(SCOPES, account=contas[0])
        if resultado and "access_token" in resultado:
            salvar_cache(cache)
            print("✅ Token reutilizado do cache.")
            return resultado["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    print("\n" + "="*60)
    print("🔐 AUTENTICAÇÃO NECESSÁRIA")
    print(f"1. Acesse: {flow['verification_uri']}")
    print(f"2. Digite o código: {flow['user_code']}")
    print("="*60 + "\n")

    resultado = app.acquire_token_by_device_flow(flow)

    if "access_token" not in resultado:
        raise Exception(f"Falha na autenticação: {resultado.get('error_description')}")

    salvar_cache(cache)
    print("✅ Autenticado com sucesso!")
    return resultado["access_token"]


def baixar_planilha() -> Path:
    PASTA_LOCAL.mkdir(parents=True, exist_ok=True)
    destino = PASTA_LOCAL / NOME_LOCAL

    token = obter_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{SHAREPOINT_URL}{SITE_PATH}/_api/web/GetFileById('{FILE_ID}')/$value"

    print("⬇️  Baixando arquivo do SharePoint...")
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        with open(destino, "wb") as f:
            f.write(response.content)
        print(f"✅ Arquivo salvo em: {destino}")
        return destino
    else:
        raise Exception(f"Erro ao baixar ({response.status_code}): {response.text}")


if __name__ == "__main__":
    baixar_planilha()
