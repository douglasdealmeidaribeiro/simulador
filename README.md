# Simulador de Modelagem Economico-Financeira

Aplicacao web com formulario baseado na aba `SIMULADOR` da planilha original. A simulacao usa exclusivamente o backend com Excel para evitar divergencias do calculo no navegador.

## Arquitetura

- Frontend: HTML, CSS e JavaScript estaticos.
- Frontend: HTML, CSS e JavaScript estaticos.
- Backend exato: `backend/server.js` chama `backend/simulate-excel.ps1` no Excel.
- Saida final: arquivo `.xlsx` da aba `Orçamento (Mensal)` atualizado com as premissas da simulacao.

## Como usar

Abra a pagina, preencha os campos e clique em:

```text
Gerar planilha calculada
```

O app devolve uma planilha `.xlsx` gerada no Excel contendo a aba `Orçamento (Mensal)` atualizada.

## Execucao

O backend Excel e obrigatorio para executar a simulacao.

## Producao no GitHub Pages

O frontend pode ser publicado no GitHub Pages, mas a simulacao depende do backend Excel acessivel pela internet.

### 1) Publicar frontend no Pages

Este repositorio ja inclui o workflow:

- [.github/workflows/deploy-pages.yml](.github/workflows/deploy-pages.yml)

Passos:

1. Envie para a branch `main`.
2. No GitHub, abra `Settings > Pages`.
3. Em `Build and deployment`, selecione `GitHub Actions`.
4. Aguarde o workflow `Deploy GitHub Pages` concluir.

### 2) Modo gratis com resultado exato no proprio PC

Se voce quer manter custo mensal zero e ainda usar o Excel real para gerar o arquivo final, rode a API no seu proprio computador e exponha via Cloudflare Tunnel.

Guia completo:

- [docs/cloudflare-self-hosted-exact-mode.md](docs/cloudflare-self-hosted-exact-mode.md)

Arquivos auxiliares incluidos:

- [scripts/start-selfhosted-exact-mode.ps1](scripts/start-selfhosted-exact-mode.ps1)
- [scripts/start-selfhosted-exact-background.ps1](scripts/start-selfhosted-exact-background.ps1)
- [scripts/stop-selfhosted-exact-background.ps1](scripts/stop-selfhosted-exact-background.ps1)
- [scripts/set-selfhosted-backend-mode.ps1](scripts/set-selfhosted-backend-mode.ps1)
- [scripts/install-selfhosted-startup-task.ps1](scripts/install-selfhosted-startup-task.ps1)
- [scripts/uninstall-selfhosted-startup-task.ps1](scripts/uninstall-selfhosted-startup-task.ps1)
- [cloudflare/config.example.yml](cloudflare/config.example.yml)

Fluxo resumido:

1. Coloque o DNS do dominio na Cloudflare.
2. Instale `cloudflared` no Windows.
3. Crie um tunnel nomeado para `api.seudominio.com`.
4. Rode `scripts/set-selfhosted-backend-mode.ps1` para apontar o frontend para a URL publica da API.
5. Rode `scripts/start-selfhosted-exact-mode.ps1` para subir backend + tunnel no seu PC.

Exemplo de configuracao do frontend:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\set-selfhosted-backend-mode.ps1 -PublicApiUrl https://api.seudominio.com
```

Exemplo para subir backend + tunnel:

```powershell
npm run exact:selfhosted:start -- -TunnelName simulador-api -PublicHostname api.seudominio.com -FrontendOrigin https://SEU_USUARIO.github.io
```

Para desabilitar o apontamento para API publica e voltar para configuracao vazia:

```powershell
npm run exact:selfhosted:disable
```

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

Sem backend Excel ativo, a aplicacao nao executa simulacao.
