# =========================
# SYNC CHUVA -> GOOGLE SHEETS
# =========================
# Busca a chuva dos últimos 15 dias (D-15 até D-1) para os municípios do IBGE
# usando as mesmas APIs do app (IBGE + Open-Meteo) e grava no Google Sheets.
#
# Uso:
#   python sync_sheets.py             -> grava no Google Sheets
#   python sync_sheets.py --dry-run   -> só gera CSV em dados_chuva/ (sem Google)
#
# Variáveis de ambiente:
#   GOOGLE_SERVICE_ACCOUNT_JSON  conteúdo do JSON da service account
#   GOOGLE_APPLICATION_CREDENTIALS  (alternativa) caminho do arquivo JSON
#   SPREADSHEET_ID               id da planilha (padrão: planilha do rain tracker)
#   UFS                          filtro opcional de UFs, ex: "SP,RJ" (vazio = todas)
#   DIAS                         quantidade de dias (padrão: 15)

import argparse
import json
import os
import sys
import time
import traceback
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import requests


# =========================
# CONFIGURAÇÕES GERAIS
# =========================

IBGE_MUNICIPIOS_URL = "https://servicodados.ibge.gov.br/api/v1/localidades/municipios?orderBy=nome"
OPEN_METEO_GEOCODING_URL = "https://geocoding-api.open-meteo.com/v1/search"
OPEN_METEO_HISTORICAL_URL = "https://archive-api.open-meteo.com/v1/archive"

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID") or "1HGGh0EWn97Wo3ZB18Bx7cbgq6VO_S8xT-y3yivpH3IE"
ABA_CHUVA = "base_chuva"
ABA_COORDENADAS = "coordenadas"

OUTPUT_DIR = Path("dados_chuva")
COORDENADAS_LOCAL_PATH = OUTPUT_DIR / "coordenadas.csv"

DIAS = int(os.environ.get("DIAS") or 15)

BATCH_SIZE_CHUVA = 40
SLEEP_GEOCODING = 0.2
# A Open-Meteo gratuita limita ~600 chamadas/minuto e cada município do batch conta
# como uma chamada; 5s entre batches de 40 mantém abaixo do limite.
SLEEP_CHUVA = 5
SHEETS_CHUNK_ROWS = 10000

COLUNAS_COORDENADAS = [
    "codigo_ibge",
    "latitude",
    "longitude",
    "timezone",
    "geocoding_name",
    "geocoding_admin1"
]

COLUNAS_CHUVA = [
    "codigo_ibge",
    "municipio",
    "uf",
    "estado",
    "regiao",
    "latitude",
    "longitude",
    "data",
    "chuva_mm",
    "chuva_liquida_mm",
    "horas_com_chuva"
]


# =========================
# FUNÇÕES AUXILIARES
# =========================

def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def normalizar_texto(texto):
    if pd.isna(texto):
        return ""

    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join([c for c in texto if not unicodedata.combining(c)])

    return texto


def request_with_retry(url, params=None, max_retries=5, timeout=90):
    for tentativa in range(1, max_retries + 1):
        try:
            response = requests.get(url, params=params, timeout=timeout)

            if response.status_code == 200:
                return response.json()

            log(f"Tentativa {tentativa} falhou. Status: {response.status_code} | {response.text[:300]}")

            # Limite de requisições: espera o minuto virar antes de tentar de novo
            if response.status_code == 429:
                time.sleep(65)
                continue

        except Exception as e:
            log(f"Tentativa {tentativa} com erro: {e}")

        time.sleep(tentativa * 3)

    raise RuntimeError(f"Falha após múltiplas tentativas na URL: {url}")


def get_nested(dicionario, caminho, default=None):
    atual = dicionario

    for chave in caminho:
        if isinstance(atual, dict) and chave in atual:
            atual = atual[chave]
        else:
            return default

    return atual


def chunk_dataframe(df, size):
    for start in range(0, len(df), size):
        yield df.iloc[start:start + size].copy()


# =========================
# 1. BUSCAR MUNICÍPIOS NO IBGE
# =========================

def buscar_municipios_ibge():
    dados = request_with_retry(IBGE_MUNICIPIOS_URL)

    linhas = []

    for item in dados:
        uf = get_nested(item, ["microrregiao", "mesorregiao", "UF"], {})

        linhas.append({
            "codigo_ibge": str(item.get("id")),
            "municipio": str(item.get("nome")),
            "uf": str(uf.get("sigla")),
            "estado": str(uf.get("nome")),
            "regiao": str(get_nested(uf, ["regiao", "nome"]))
        })

    return pd.DataFrame(linhas)


# =========================
# 2. GEOCODING (COM CACHE)
# =========================

def escolher_melhor_resultado_geocoding(resultados, municipio, estado):
    if not resultados:
        return None

    municipio_norm = normalizar_texto(municipio)
    estado_norm = normalizar_texto(estado)

    candidatos_br = [
        r for r in resultados
        if r.get("country_code") == "BR"
    ]

    if not candidatos_br:
        candidatos_br = resultados

    for r in candidatos_br:
        if normalizar_texto(r.get("name")) == municipio_norm and normalizar_texto(r.get("admin1")) == estado_norm:
            return r

    for r in candidatos_br:
        if normalizar_texto(r.get("admin1")) == estado_norm:
            return r

    for r in candidatos_br:
        if normalizar_texto(r.get("name")) == municipio_norm:
            return r

    return candidatos_br[0]


def geocodificar_municipio(municipio, uf, estado):
    termos_busca = [
        f"{municipio}, {uf}, Brasil",
        f"{municipio}, {estado}, Brasil",
        municipio
    ]

    for termo in termos_busca:
        params = {
            "name": termo,
            "count": 20,
            "language": "pt",
            "format": "json",
            "countryCode": "BR"
        }

        data = request_with_retry(
            OPEN_METEO_GEOCODING_URL,
            params=params,
            max_retries=3,
            timeout=30
        )

        melhor = escolher_melhor_resultado_geocoding(
            resultados=data.get("results", []),
            municipio=municipio,
            estado=estado
        )

        if melhor:
            return {
                "latitude": melhor.get("latitude"),
                "longitude": melhor.get("longitude"),
                "timezone": melhor.get("timezone"),
                "geocoding_name": melhor.get("name"),
                "geocoding_admin1": melhor.get("admin1")
            }

        time.sleep(SLEEP_GEOCODING)

    return {
        "latitude": None,
        "longitude": None,
        "timezone": None,
        "geocoding_name": None,
        "geocoding_admin1": None
    }


def adicionar_coordenadas(df_municipios, cache):
    """Usa o cache de coordenadas e geocodifica só os municípios que faltam.
    Retorna (df_com_coordenadas, cache_atualizado, houve_novos)."""

    cache = cache.copy()
    cache["latitude"] = pd.to_numeric(cache["latitude"], errors="coerce")
    cache["longitude"] = pd.to_numeric(cache["longitude"], errors="coerce")
    cache = cache.dropna(subset=["latitude", "longitude"])

    faltantes = df_municipios[~df_municipios["codigo_ibge"].isin(cache["codigo_ibge"])]
    total = len(faltantes)

    log(f"Coordenadas em cache: {len(cache)} | A geocodificar: {total}")

    novos = []

    for i, row in enumerate(faltantes.itertuples(index=False), start=1):
        # Se estourar o limite da API, guarda o que já foi feito e completa no próximo run
        try:
            geo = geocodificar_municipio(row.municipio, row.uf, row.estado)
        except RuntimeError as e:
            log(f"Geocodificação interrompida em {i}/{total}: {e}")
            break

        geo["codigo_ibge"] = row.codigo_ibge
        novos.append(geo)

        if i % 100 == 0 or i == total:
            log(f"Geocodificados {i}/{total}")

        time.sleep(SLEEP_GEOCODING)

    if novos:
        cache = pd.concat([cache, pd.DataFrame(novos)], ignore_index=True)
        cache["latitude"] = pd.to_numeric(cache["latitude"], errors="coerce")
        cache["longitude"] = pd.to_numeric(cache["longitude"], errors="coerce")

    cache = cache[COLUNAS_COORDENADAS].drop_duplicates("codigo_ibge", keep="last")

    df_geo = df_municipios.merge(cache, on="codigo_ibge", how="left")
    df_geo["timezone"] = df_geo["timezone"].replace("", pd.NA).fillna("America/Sao_Paulo")

    return df_geo, cache, bool(novos)


# =========================
# 3. BUSCAR CHUVA NA OPEN-METEO
# =========================

def fetch_chuva_batch(df_batch, start_date, end_date, timezone_name):
    params = {
        "latitude": ",".join(df_batch["latitude"].astype(str).tolist()),
        "longitude": ",".join(df_batch["longitude"].astype(str).tolist()),
        "start_date": start_date,
        "end_date": end_date,
        "daily": "precipitation_sum,rain_sum,precipitation_hours",
        "timezone": timezone_name,
        "precipitation_unit": "mm"
    }

    data = request_with_retry(
        OPEN_METEO_HISTORICAL_URL,
        params=params,
        max_retries=5,
        timeout=120
    )

    if isinstance(data, dict):
        data = [data]

    resultados = []

    for i, item in enumerate(data):
        if i >= len(df_batch):
            continue

        daily = item.get("daily", {})

        if not daily:
            continue

        temp = pd.DataFrame(daily)

        if temp.empty:
            continue

        municipio_info = df_batch.iloc[i]

        for col in ["codigo_ibge", "municipio", "uf", "estado", "regiao", "latitude", "longitude"]:
            temp[col] = municipio_info[col]

        resultados.append(temp)

    if resultados:
        return pd.concat(resultados, ignore_index=True)

    return pd.DataFrame()


def buscar_chuva_municipios(df_municipios_geo, start_date, end_date):
    df_ok = df_municipios_geo.dropna(subset=["latitude", "longitude"]).copy()

    total_municipios = len(df_ok)

    if total_municipios == 0:
        raise RuntimeError("Nenhum município com latitude/longitude para buscar chuva.")

    bases = []
    processados = 0

    for timezone_name, df_fuso in df_ok.groupby("timezone"):
        for batch in chunk_dataframe(df_fuso, BATCH_SIZE_CHUVA):
            df_chuva_batch = fetch_chuva_batch(batch, start_date, end_date, timezone_name)

            if not df_chuva_batch.empty:
                bases.append(df_chuva_batch)

            processados += len(batch)
            log(f"Chuva | {timezone_name} | {processados}/{total_municipios} municípios")

            time.sleep(SLEEP_CHUVA)

    if not bases:
        raise RuntimeError("Nenhum dado de chuva foi retornado pela API.")

    chuva = pd.concat(bases, ignore_index=True)

    chuva = chuva.rename(columns={
        "time": "data",
        "precipitation_sum": "chuva_mm",
        "rain_sum": "chuva_liquida_mm",
        "precipitation_hours": "horas_com_chuva"
    })

    for col in ["chuva_mm", "chuva_liquida_mm", "horas_com_chuva"]:
        chuva[col] = pd.to_numeric(chuva[col], errors="coerce")

    return chuva[COLUNAS_CHUVA].sort_values(["uf", "municipio", "data"]).reset_index(drop=True)


# =========================
# 4. GOOGLE SHEETS
# =========================

def conectar_planilha():
    import gspread

    credenciais_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

    if credenciais_json:
        client = gspread.service_account_from_dict(json.loads(credenciais_json))
    else:
        client = gspread.service_account(filename=os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"))

    return client.open_by_key(SPREADSHEET_ID)


def obter_aba(planilha, nome, linhas, colunas):
    import gspread

    try:
        return planilha.worksheet(nome)
    except gspread.WorksheetNotFound:
        return planilha.add_worksheet(title=nome, rows=linhas, cols=colunas)


def ler_coordenadas_sheets(planilha):
    aba = obter_aba(planilha, ABA_COORDENADAS, 100, len(COLUNAS_COORDENADAS))
    valores = aba.get_all_values()

    if len(valores) < 2:
        return pd.DataFrame(columns=COLUNAS_COORDENADAS)

    df = pd.DataFrame(valores[1:], columns=valores[0])

    return df.reindex(columns=COLUNAS_COORDENADAS)


def valores_para_sheets(df):
    df = df.astype(object).where(pd.notna(df), "")

    return [df.columns.tolist()] + df.values.tolist()


def escrever_aba(planilha, nome, df, value_input_option):
    valores = valores_para_sheets(df)

    aba = obter_aba(planilha, nome, len(valores), len(df.columns))
    aba.clear()
    aba.resize(rows=len(valores), cols=len(df.columns))

    for inicio in range(0, len(valores), SHEETS_CHUNK_ROWS):
        bloco = valores[inicio:inicio + SHEETS_CHUNK_ROWS]

        aba.update(
            bloco,
            range_name=f"A{inicio + 1}",
            value_input_option=value_input_option
        )

        log(f"Aba '{nome}': {min(inicio + SHEETS_CHUNK_ROWS, len(valores))}/{len(valores)} linhas")


# =========================
# 5. EXECUÇÃO
# =========================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Gera CSV local sem gravar no Google Sheets")
    args = parser.parse_args()

    d_1 = date.today() - timedelta(days=1)
    data_inicio = d_1 - timedelta(days=DIAS - 1)

    start_date = data_inicio.isoformat()
    end_date = d_1.isoformat()

    log(f"Período: {start_date} até {end_date}")

    municipios = buscar_municipios_ibge()

    ufs = [u.strip().upper() for u in os.environ.get("UFS", "").split(",") if u.strip()]

    if ufs:
        municipios = municipios[municipios["uf"].isin(ufs)].copy()

    log(f"Municípios: {len(municipios)}")

    OUTPUT_DIR.mkdir(exist_ok=True)

    if args.dry_run:
        planilha = None

        if COORDENADAS_LOCAL_PATH.exists():
            cache = pd.read_csv(COORDENADAS_LOCAL_PATH, dtype={"codigo_ibge": str})
        else:
            cache = pd.DataFrame(columns=COLUNAS_COORDENADAS)
    else:
        planilha = conectar_planilha()
        cache = ler_coordenadas_sheets(planilha)

    municipios_geo, cache, houve_novos = adicionar_coordenadas(municipios, cache)

    # Salva o cache antes de buscar a chuva, para não perder a geocodificação se algo falhar depois
    if houve_novos:
        if args.dry_run:
            cache.to_csv(COORDENADAS_LOCAL_PATH, index=False, encoding="utf-8-sig")
        else:
            escrever_aba(planilha, ABA_COORDENADAS, cache, value_input_option="RAW")

    chuva = buscar_chuva_municipios(municipios_geo, start_date, end_date)

    log(f"Linhas de chuva: {len(chuva)}")

    if args.dry_run:
        csv_path = OUTPUT_DIR / f"sync_chuva_{start_date}_a_{end_date}.csv"
        chuva.to_csv(csv_path, index=False, encoding="utf-8-sig")
        log(f"Dry run: CSV salvo em {csv_path}")
        return

    # USER_ENTERED para a planilha reconhecer as datas; números vão como número
    escrever_aba(planilha, ABA_CHUVA, chuva, value_input_option="USER_ENTERED")

    log("Sync finalizado.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        # Traceback completo: alguns erros do gspread (ex: PermissionError) vêm sem mensagem
        traceback.print_exc()
        sys.exit(1)
