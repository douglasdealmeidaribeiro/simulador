# Simulador de Modelagem Econômico-Financeira

Aplicação web com formulário baseado na aba `SIMULADOR` da planilha original. Os cálculos são executados pelo Excel em um backend Windows, e o usuário baixa a planilha `.xlsm` já calculada.

## Arquitetura

- Frontend: HTML, CSS e JavaScript estáticos, compatíveis com GitHub Pages.
- Backend recomendado: PowerShell em servidor Windows com Microsoft Excel instalado.
- Cálculo: o backend abre a planilha, preenche os inputs, executa a macro `SimularESDigital` ou GoalSeek equivalente, salva e devolve o `.xlsm`.

## Frontend

O endereço da API fica em:

```text
assets/api-config.js
```

Por padrão:

```js
window.SIMULADOR_API_URL = 'http://127.0.0.1:3000';
```

Em produção, altere esse valor para a URL pública HTTPS do backend.

## Backend local

Requisitos:

- Windows
- Microsoft Excel instalado
- Permissão para automação COM do Excel
- Sessão interativa de usuário logada no Windows; automação COM do Office normalmente falha quando executada como serviço ou sessão não interativa

Execute:

```powershell
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File backend\server.ps1
```

O `package.json` também mantém `npm run backend` com o mesmo comando, mas em ambientes com restrição de COM prefira executar o PowerShell diretamente. O backend alternativo em Node foi mantido em `npm run backend:node`, porém a automação COM do Excel é mais confiável quando roda diretamente em PowerShell na sessão do usuário.

Se aparecer o erro `80070520`, o Windows bloqueou a criação do `Excel.Application` nessa sessão. Rode o backend em uma sessão Windows interativa com Excel instalado, ou publique atrás de um serviço/reverse proxy que encaminhe para esse processo interativo.

Teste:

```text
http://127.0.0.1:3000/health
```

## Publicação

1. Publique o frontend no GitHub Pages.
2. Publique o backend em um servidor Windows com Excel instalado.
3. Atualize `assets/api-config.js` com a URL pública do backend.
4. Faça novo commit/push do frontend.

## Manutenção

Se a planilha original mudar, gere novamente os dados do formulário:

```powershell
python scripts\generate_workbook_model.py
```
