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

Se o site bloquear o navegador automatizado ("comportamento incomum vindo do
seu IP"), use o modo arquivo: abra o extrato no seu navegador normal, exporte
um HAR pelas ferramentas de desenvolvedor (F12 > Rede) ou salve a página, e rode
  python azul_extrato.py --arquivo extrato.har
"""

import argparse
import base64
import csv
import json
import re
import sys
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path

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


def ler_har(caminho):
    """Lê as respostas JSON de um arquivo HAR exportado pelo navegador (F12 > Rede)."""
    har = json.loads(Path(caminho).read_text(encoding="utf-8-sig"))
    capturas = []
    for entrada in har.get("log", {}).get("entries", []):
        conteudo = entrada.get("response", {}).get("content", {})
        texto = conteudo.get("text")
        if not texto:
            continue
        if conteudo.get("encoding") == "base64":
            try:
                texto = base64.b64decode(texto).decode("utf-8")
            except Exception:
                continue
        try:
            corpo = json.loads(texto)
        except ValueError:
            continue
        url = entrada.get("request", {}).get("url", "")
        if any(True for _ in listas_de_movimentacoes(corpo)) or PALAVRAS_URL.search(url):
            capturas.append((url, corpo))
    return capturas


class _LeitorHTML(HTMLParser):
    """Extrai as linhas de tabelas e o texto, linha a linha, de uma página salva."""

    BLOCOS = {"p", "div", "li", "tr", "br", "h1", "h2", "h3", "h4", "section", "article"}

    def __init__(self):
        super().__init__()
        self.tabelas, self.texto = [], []
        self._ignorar = 0
        self._linha = self._celula = None

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "noscript"):
            self._ignorar += 1
        elif tag == "table":
            self.tabelas.append([])
        elif tag == "tr" and self.tabelas:
            self._linha = []
        elif tag in ("td", "th") and self._linha is not None:
            self._celula = []
        if tag in self.BLOCOS:
            self.texto.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript"):
            self._ignorar = max(0, self._ignorar - 1)
        elif tag in ("td", "th") and self._celula is not None:
            self._linha.append((tag, " ".join("".join(self._celula).split())))
            self._celula = None
        elif tag == "tr" and self._linha is not None:
            self.tabelas[-1].append(self._linha)
            self._linha = None
        if tag in self.BLOCOS:
            self.texto.append("\n")

    def handle_data(self, data):
        if self._ignorar:
            return
        if self._celula is not None:
            self._celula.append(data)
        self.texto.append(data)


def ler_html(caminho):
    """Plano B para página salva (Ctrl+S): tabelas ou linhas com data dd/mm/aaaa."""
    leitor = _LeitorHTML()
    leitor.feed(Path(caminho).read_text(encoding="utf-8", errors="replace"))
    linhas = []
    for tabela in leitor.tabelas:
        cab = [t for tag, t in tabela[0]] if tabela and all(tag == "th" for tag, _ in tabela[0]) else []
        for tr in tabela:
            cels = [t for tag, t in tr if tag == "td"]
            if cels:
                linhas.append({(cab[i] if i < len(cab) and cab[i] else f"coluna_{i + 1}"): c for i, c in enumerate(cels)})
    if linhas:
        return linhas
    data = re.compile(r"\b\d{2}/\d{2}/\d{2,4}\b")
    for t in "".join(leitor.texto).split("\n"):
        t = " ".join(t.split())
        if len(data.findall(t)) == 1 and re.search(r"\d", data.sub("", t)):
            linhas.append({"linha": t})
    return linhas


def capturar_no_navegador(args):
    """Abre o navegador, espera o usuário chegar ao extrato e devolve (capturas, linhas_html)."""
    from playwright.sync_api import sync_playwright

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
                "   Se aparecer 'comportamento incomum vindo do seu IP', feche tudo e use o\n"
                "   modo --arquivo (veja o README).\n"
            )
            input()

        # A página ativa pode ter mudado (nova aba etc.): usa a última aberta.
        page = ctx.pages[-1]
        linhas_html = ler_tabelas_da_pagina(page)
        ctx.close()
    return capturas, linhas_html


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
    ap.add_argument("--arquivo", nargs="+", metavar="ARQ",
                    help="não abre navegador: lê arquivo(s) .har, .json ou .html salvos do seu navegador")
    args = ap.parse_args()

    destino = Path(args.saida)
    (destino / "bruto").mkdir(parents=True, exist_ok=True)
    if args.arquivo:
        capturas, linhas_html = [], []
        for arq in args.arquivo:
            if arq.lower().endswith(".har"):
                capturas += ler_har(arq)
            elif arq.lower().endswith(".json"):
                corpo = json.loads(Path(arq).read_text(encoding="utf-8-sig"))
                capturas.append((arq, corpo))
            else:
                linhas_html += ler_html(arq)
    else:
        capturas, linhas_html = capturar_no_navegador(args)

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
            "Não encontrei movimentações. Confira se a página de extrato estava aberta\n"
            "(ao pressionar ENTER, ou ao exportar o HAR / salvar a página).\n"
            f"As respostas capturadas estão em {bruto} — me envie (sem dados sensíveis) para ajustar o script."
        )
        sys.exit(1)

    salvar(linhas, destino, f"extrato_{carimbo}")
    print(f"  Bruto: {bruto}")


if __name__ == "__main__":
    main()
