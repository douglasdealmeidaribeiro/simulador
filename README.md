# Simulador de Modelagem Economico-Financeira

Aplicacao web com formulario baseado na aba `SIMULADOR` da planilha original. O calculo roda no navegador (modo gratuito) e gera arquivo de resultados localmente.

## Arquitetura

- Frontend: HTML, CSS e JavaScript estaticos.
- Calculo: `worker.js` usa HyperFormula e `assets/workbook-model.js`.
- Saida: arquivo `.xls` de resultados gerado no navegador.
- Opcional avancado: backend Excel (`backend/server.js`) para cenarios de alta aderencia.

## Como usar

Abra a pagina, preencha os campos e clique em:

```text
Gerar planilha calculada
```

O app devolve uma planilha `.xls` com resultados simulados.

## Execucao

Abra o frontend e use normalmente. O backend nao e necessario no modo gratuito.

## Producao no GitHub Pages

Modo recomendado para custo zero: publicar somente o frontend no GitHub Pages.

### 1) Publicar frontend no Pages

Este repositorio ja inclui o workflow:

- [.github/workflows/deploy-pages.yml](.github/workflows/deploy-pages.yml)

Passos:

1. Envie para a branch `main`.
2. No GitHub, abra `Settings > Pages`.
3. Em `Build and deployment`, selecione `GitHub Actions`.
4. Aguarde o workflow `Deploy GitHub Pages` concluir.

### 2) (Opcional) Backend Excel

Se no futuro voce quiser maior aderencia ao Excel, pode habilitar backend depois em `assets/api-config.js`.

Comandos usuais:

```powershell
git add .
git commit -m "Converte simulador para versao estatica"
git push origin main
```

## Desenvolvimento local

Por usar Web Worker, prefira abrir por servidor local:

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

No modo gratuito/local pode haver diferencas em alguns cenarios em relacao ao Excel original.
