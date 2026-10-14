# Extrato Azul Fidelidade

Script que extrai o extrato de pontos do **Azul Fidelidade** e salva em CSV e Excel.

## Instalação (uma vez)

```bash
python -m pip install -r requirements.txt
```

O script usa o **Google Chrome** ou o **Microsoft Edge** que já estão no computador (o Edge vem
com o Windows), então não precisa baixar navegador. Só se nenhum dos dois existir rode
`python -m playwright install chromium`. Para forçar um deles: `--navegador msedge` ou `--navegador chrome`.

> **Windows:** use `python -m playwright ...` (e não só `playwright ...`). O comando solto
> só funciona se a pasta `Scripts` do Python estiver no PATH. Se `python` não for reconhecido,
> troque por `py` (ex.: `py azul_extrato.py`).

## Uso

```bash
python azul_extrato.py
```

1. Abre uma janela do Chromium no site da Azul.
2. Faça login normalmente (captcha, código por SMS/e-mail etc.).
3. Abra a página **Extrato** e escolha o período. Se tiver botão "ver mais", clique até carregar tudo.
4. Volte ao terminal e pressione **ENTER**.

Os arquivos vão para `saida/`:

- `extrato_AAAAMMDD_HHMMSS.csv`: separado por `;`, abre direto no Excel.
- `extrato_AAAAMMDD_HHMMSS.xlsx`
- `bruto/respostas_*.json`: respostas do site capturadas, para conferência.

O login fica salvo em `.perfil-azul/`, então nas próximas vezes você normalmente não precisa logar de novo.
Com a sessão salva dá até para rodar sem janela:

```bash
python azul_extrato.py --headless --auto 20 --url "<URL da página de extrato>"
```

## Como funciona

O script não usa sua senha. Você faz o login no navegador e o script só escuta as respostas JSON
que o próprio site carrega na página de extrato, reconhecendo listas de movimentações
(registros com data e pontos). Se nenhuma resposta for reconhecida, ele lê as tabelas/linhas
visíveis na página.

> As pastas `.perfil-azul/` e `saida/` contêm sua sessão e seus dados. Já estão no `.gitignore`,
> não as compartilhe.
