# Simulador de Modelagem Econômico-Financeira

Aplicação web estática baseada na aba `SIMULADOR` da planilha `Anexo I – Simulador de Modelagem Econômico-Financeira-Apresentação 1.xlsm`.

## Como usar localmente

Abra `index.html` em um servidor estático local:

```powershell
python -m http.server 8000
```

Depois acesse:

```text
http://localhost:8000
```

## Publicação no GitHub Pages

1. Suba os arquivos deste diretório para um repositório GitHub.
2. Em `Settings > Pages`, selecione a branch principal e a pasta `/root`.
3. Aguarde a publicação do GitHub Pages.

O app é 100% estático: HTML, CSS e JavaScript no navegador.

## Manutenção

Se a planilha original mudar, gere novamente os arquivos de dados:

```powershell
python scripts\generate_workbook_model.py
```
