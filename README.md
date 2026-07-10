# Simulador de Modelagem Economico-Financeira

Aplicacao web estatica com formulario baseado na aba `SIMULADOR` da planilha original. Os calculos rodam no proprio navegador com o modelo extraido da planilha, sem backend, sem PowerShell e sem Excel instalado no servidor.

## Arquitetura

- Frontend: HTML, CSS e JavaScript estaticos.
- Hospedagem: compativel com GitHub Pages.
- Calculo: `worker.js` usa HyperFormula e `assets/workbook-model.js`.
- Saida: arquivo `.xls` com resumo dos resultados, dados gerais, centros de custo e ajustes informados.

## Como usar

Abra a pagina publicada no GitHub Pages, preencha os campos e clique em:

```text
Gerar planilha calculada
```

O navegador executa a simulacao localmente e baixa uma planilha de resultados.

## Publicacao no GitHub Pages

1. Envie os arquivos para o repositorio.
2. No GitHub, acesse `Settings > Pages`.
3. Use `Deploy from branch`.
4. Selecione a branch `main` e a pasta `/root`.
5. Aguarde a publicacao.

Comandos usuais:

```powershell
git add .
git commit -m "Converte simulador para versao estatica"
git push origin main
```

## Desenvolvimento local

Por usar Web Worker, prefira abrir por um servidor local:

```powershell
python -m http.server 8000
```

Depois acesse:

```text
http://127.0.0.1:8000
```

## Manutencao

Se a planilha original mudar, gere novamente os dados do modelo:

```powershell
python scripts\generate_workbook_model.py
```

Depois valide os scripts:

```powershell
node --check app.js
node --check worker.js
```

## Observacao

A versao estatica nao entrega o arquivo `.xlsm` original com macros recalculadas pelo Excel. Ela entrega uma planilha `.xls` de resultados calculados no navegador, adequada para publicacao gratuita no GitHub Pages.
