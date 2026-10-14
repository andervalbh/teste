#!/usr/bin/env python3
"""Extrai o extrato de pontos do Azul Fidelidade e salva em CSV/Excel.

Fluxo:
  1. Abre o Google Chrome ou o Microsoft Edge já instalados no computador
     (ou o Chromium do Playwright, se nenhum dos dois existir), visível (perfil persistente em ./.perfil-azul,
     então o login fica salvo entre execuções).
  2. Você faz login manualmente (captcha / código por SMS / e-mail) e abre a
     página de extrato, ajustando o período desejado.
  3. Enquanto isso o script escuta as respostas JSON do site e guarda as que
     parecem conter movimentações de pontos.
  4. Ao pressionar ENTER no terminal, o script consolida as movimentações
     (JSON capturado; se não houver, lê as tabelas/listas da página) e grava
     em ./saida/extrato_<data>.csv e .xlsx, além do JSON bruto para conferência.

Uso:
  python -m pip install -r requirements.txt
  python azul_extrato.py
"""

import argparse
import csv
import json
import re
import sys
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

URL_PADRAO = "https://www.voeazul.com.br/br/pt/azul-fidelidade"

# Palavras que indicam que uma resposta/registro tem a ver com extrato de pontos.
PALAVRAS_URL = re.compile(
    r"extrato|statement|transac|transaction|movimenta|accrual|redemption|points|pontos|history|historico",
    re.I,
)
CHAVES_DATA = re.compile(r"date|data|dt|created|posted|transaction.?date", re.I)
CHAVES_PONTOS = re.compile(r"point|ponto|amount|quantidade|qtd|valor|value|miles|balance|saldo", re.I)


def achatar(obj, prefixo=""):
    """Transforma um dict aninhado em um dict de um nível (chaves com '.')."""
    plano = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            plano.update(achatar(v, f"{prefixo}{k}."))
    elif isinstance(obj, list):
        if all(not isinstance(i, (dict, list)) for i in obj):
            plano[prefixo[:-1]] = "; ".join(map(str, obj))
        else:
            plano[prefixo[:-1]] = json.dumps(obj, ensure_ascii=False)
    else:
        plano[prefixo[:-1]] = obj
    return plano


def parece_movimentacao(item):
    """Um registro de extrato normalmente tem ao menos uma data e um valor de pontos."""
    if not isinstance(item, dict):
        return False
    chaves = list(achatar(item).keys())
    return any(CHAVES_DATA.search(k) for k in chaves) and any(CHAVES_PONTOS.search(k) for k in chaves)


def listas_de_movimentacoes(obj):
    """Procura recursivamente listas de dicts que pareçam movimentações."""
    if isinstance(obj, list):
        dicts = [i for i in obj if isinstance(i, dict)]
        if dicts and sum(parece_movimentacao(i) for i in dicts) >= max(1, len(dicts) // 2):
            yield dicts
            return
        for i in obj:
            yield from listas_de_movimentacoes(i)
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from listas_de_movimentacoes(v)


def ler_tabelas_da_pagina(page):
    """Plano B: lê tabelas HTML e linhas de listas visíveis em todos os frames."""
    script = """
    () => {
      const linhas = [];
      for (const tabela of document.querySelectorAll('table')) {
        const cab = [...tabela.querySelectorAll('thead th, tr:first-child th')]
          .map(th => th.innerText.trim());
        for (const tr of tabela.querySelectorAll('tbody tr, tr')) {
          const cels = [...tr.querySelectorAll('td')].map(td => td.innerText.trim());
          if (!cels.length) continue;
          const reg = {};
          cels.forEach((c, i) => reg[cab[i] || `coluna_${i + 1}`] = c);
          linhas.push(reg);
        }
      }
      if (!linhas.length) {
        // Layouts sem <table>: linhas com data dd/mm/aaaa e algum número.
        const vistos = new Set();
        for (const el of document.querySelectorAll('li, [role=row], div')) {
          const t = el.innerText ? el.innerText.trim() : '';
          if (!t || t.length > 300 || vistos.has(t)) continue;
          if (!/\\b\\d{2}\\/\\d{2}\\/\\d{2,4}\\b/.test(t) || !/\\d/.test(t.replace(/\\d{2}\\/\\d{2}\\/\\d{2,4}/, ''))) continue;
          // Ignora contêineres que englobam várias linhas já capturadas.
          if ((t.match(/\\b\\d{2}\\/\\d{2}\\/\\d{2,4}\\b/g) || []).length > 1) continue;
          vistos.add(t);
          const partes = t.split(/\\n+/).map(s => s.trim()).filter(Boolean);
          const reg = {};
          partes.forEach((p, i) => reg[`campo_${i + 1}`] = p);
          linhas.push(reg);
        }
      }
      return linhas;
    }
    """
    linhas = []
    for frame in page.frames:
        try:
            linhas.extend(frame.evaluate(script))
        except Exception:
            pass
    return linhas


def salvar(linhas, destino: Path, nome: str):
    colunas = []
    for linha in linhas:
        for k in linha:
            if k not in colunas:
                colunas.append(k)

    csv_path = destino / f"{nome}.csv"
    # utf-8-sig + ';' para abrir direto no Excel em português.
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=colunas, delimiter=";", extrasaction="ignore")
        w.writeheader()
        w.writerows(linhas)
    print(f"  CSV:   {csv_path}")

    try:
        from openpyxl import Workbook
    except ImportError:
        print("  (openpyxl não instalado; pulando .xlsx)")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Extrato"
    ws.append(colunas)
    for linha in linhas:
        ws.append([linha.get(c) if not isinstance(linha.get(c), (dict, list)) else json.dumps(linha.get(c), ensure_ascii=False) for c in colunas])
    xlsx_path = destino / f"{nome}.xlsx"
    wb.save(xlsx_path)
    print(f"  Excel: {xlsx_path}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--url", default=URL_PADRAO, help="página inicial (padrão: Azul Fidelidade)")
    ap.add_argument("--saida", default="saida", help="pasta de saída (padrão: ./saida)")
    ap.add_argument("--perfil", default=".perfil-azul", help="pasta do perfil do navegador (guarda o login)")
    ap.add_argument("--headless", action="store_true", help="sem janela (só funciona se o login já estiver salvo no perfil)")
    ap.add_argument("--auto", type=float, metavar="SEG",
                    help="em vez de esperar ENTER, extrai após SEG segundos (útil com --headless)")
    ap.add_argument("--navegador", choices=["auto", "chrome", "msedge", "chromium"], default="auto",
                    help="auto (padrão) tenta Chrome, depois Edge, depois o Chromium do Playwright")
    ap.add_argument("--chrome", metavar="CAMINHO", help="usar outro executável do Chrome/Chromium")
    args = ap.parse_args()

    destino = Path(args.saida)
    (destino / "bruto").mkdir(parents=True, exist_ok=True)
    capturas = []  # (url, json)

    def ao_responder(resp):
        try:
            if "json" not in (resp.headers.get("content-type") or ""):
                return
            corpo = resp.json()
        except Exception:
            return
        tem_mov = any(True for _ in listas_de_movimentacoes(corpo))
        if tem_mov or PALAVRAS_URL.search(resp.url):
            capturas.append((resp.url, corpo))
            if tem_mov:
                print(f"  [capturado] {resp.url[:120]}")

    with sync_playwright() as p:
        opcoes = dict(headless=args.headless, locale="pt-BR", viewport={"width": 1280, "height": 900})
        if args.chrome:
            tentativas = [dict(executable_path=args.chrome)]
        elif args.navegador == "auto":
            tentativas = [dict(channel="chrome"), dict(channel="msedge"), {}]
        elif args.navegador == "chromium":
            tentativas = [{}]
        else:
            tentativas = [dict(channel=args.navegador)]
        ctx = None
        for extra in tentativas:
            try:
                ctx = p.chromium.launch_persistent_context(args.perfil, **opcoes, **extra)
                break
            except Exception as e:
                erro = e
        if ctx is None:
            sys.exit(
                "Não consegui abrir nenhum navegador. Instale o Google Chrome ou o Microsoft Edge,\n"
                "ou rode: python -m playwright install chromium\n\n"
                f"Último erro: {str(erro).splitlines()[0]}"
            )
        ctx.on("response", ao_responder)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(args.url, wait_until="domcontentloaded")

        if args.auto is not None:
            page.wait_for_timeout(args.auto * 1000)
        else:
            print(
                "\n1) Faça login no Azul Fidelidade na janela que abriu.\n"
                "2) Vá em 'Extrato' e escolha o período (role a página / clique em 'ver mais'\n"
                "   até aparecerem todas as movimentações que você quer).\n"
                "3) Volte aqui e pressione ENTER para extrair.\n"
            )
            input()

        # A página ativa pode ter mudado (nova aba etc.): usa a última aberta.
        page = ctx.pages[-1]
        linhas_html = ler_tabelas_da_pagina(page)
        ctx.close()

    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Guarda tudo que foi capturado para conferência / ajuste do script.
    bruto = destino / "bruto" / f"respostas_{carimbo}.json"
    bruto.write_text(json.dumps([{"url": u, "json": j} for u, j in capturas], ensure_ascii=False, indent=2), encoding="utf-8")

    linhas, vistos = [], set()
    for url, corpo in capturas:
        for lista in listas_de_movimentacoes(corpo):
            for item in lista:
                plano = achatar(item)
                chave = json.dumps(plano, sort_keys=True, ensure_ascii=False, default=str)
                if chave not in vistos:  # a mesma página pode ser carregada mais de uma vez
                    vistos.add(chave)
                    linhas.append(plano)

    print()
    if linhas:
        print(f"{len(linhas)} movimentações encontradas nas respostas do site.")
    elif linhas_html:
        print(f"Nenhum JSON de extrato reconhecido; usando {len(linhas_html)} linhas lidas da página.")
        linhas = linhas_html
    else:
        print(
            "Não encontrei movimentações. Confira se a página de extrato estava aberta ao pressionar ENTER.\n"
            f"As respostas capturadas estão em {bruto} — me envie (sem dados sensíveis) para ajustar o script."
        )
        sys.exit(1)

    salvar(linhas, destino, f"extrato_{carimbo}")
    print(f"  Bruto: {bruto}")


if __name__ == "__main__":
    main()
