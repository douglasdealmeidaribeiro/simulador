# Modo Exato Gratis no Proprio PC com Cloudflare Tunnel

Este guia publica a API local do simulador na internet sem custo mensal, usando:

- seu PC Windows
- Excel instalado localmente
- API local do projeto
- Cloudflare Tunnel
- seu dominio registrado na Hostinger

## Visao geral

Arquitetura:

1. Frontend no GitHub Pages.
2. API rodando localmente no seu PC em `http://127.0.0.1:3000`.
3. Cloudflare Tunnel expondo a API em `https://api.seudominio.com`.
4. Frontend configurado para usar `https://api.seudominio.com`.

Limites importantes:

- seu PC precisa ficar ligado
- Excel precisa estar instalado
- a API so funciona enquanto o backend e o tunnel estiverem rodando

## Requisito importante sobre dominio Hostinger

Ter o dominio registrado na Hostinger e suficiente, mas para usar `api.seudominio.com` com Cloudflare Tunnel voce precisa colocar o DNS do dominio sob a Cloudflare.

Isso significa:

1. manter a Hostinger como registradora do dominio
2. trocar os nameservers do dominio para os nameservers da Cloudflare

Sem isso, voce ainda pode usar Quick Tunnel da Cloudflare, mas a URL sera aleatoria e nao o seu dominio.

## Etapa 1 - Colocar o DNS na Cloudflare

1. Crie uma conta gratuita na Cloudflare.
2. Adicione seu dominio na Cloudflare.
3. A Cloudflare vai importar os registros DNS existentes.
4. No painel da Hostinger, abra a administracao do dominio.
5. Troque os nameservers atuais pelos nameservers informados pela Cloudflare.
6. Aguarde a propagacao.

Quando terminar, o dominio continua sendo seu na Hostinger, mas o DNS passa a ser administrado pela Cloudflare.

## Etapa 2 - Instalar o cloudflared no Windows

Opcoes comuns:

1. Baixar o binario/instalador em:
   https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/
2. Instalar via `winget`:

```powershell
winget install Cloudflare.cloudflared
```

Depois valide:

```powershell
cloudflared --version
```

## Etapa 3 - Autenticar no Cloudflare

No PowerShell:

```powershell
cloudflared tunnel login
```

Isso abre o navegador para autorizar a conta e escolher o dominio.

## Etapa 4 - Criar um tunnel nomeado

Exemplo:

```powershell
cloudflared tunnel create simulador-api
```

Guarde o nome `simulador-api`.

## Etapa 5 - Apontar o subdominio da API

Crie o hostname publico do tunnel:

```powershell
cloudflared tunnel route dns simulador-api api.seudominio.com
```

Substitua `seudominio.com` pelo seu dominio real.

## Etapa 6 - Criar configuracao local do tunnel

Copie o arquivo exemplo:

- [cloudflare/config.example.yml](../cloudflare/config.example.yml)

para um arquivo local, por exemplo:

- `%USERPROFILE%\\.cloudflared\\config.yml`

Conteudo esperado:

```yml
tunnel: simulador-api
credentials-file: C:\\Users\\SEU_USUARIO\\.cloudflared\\TUNNEL_ID.json

ingress:
  - hostname: api.seudominio.com
    service: http://127.0.0.1:3000
  - service: http_status:404
```

Ajuste:

1. `tunnel` com o nome ou ID correto
2. `credentials-file` com o caminho real do arquivo gerado pela Cloudflare
3. `hostname` com seu dominio real

## Etapa 7 - Configurar o frontend para usar a API publica

No projeto, execute:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\set-selfhosted-backend-mode.ps1 -PublicApiUrl https://api.seudominio.com
```

Esse comando atualiza `assets/api-config.js` para:

1. usar `https://api.seudominio.com`
2. manter o app em modo exato via backend

Depois publique novamente o frontend no GitHub Pages.

## Etapa 8 - Subir backend e tunnel no seu PC

Use o script auxiliar do projeto:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\start-selfhosted-exact-mode.ps1 -TunnelName simulador-api -FrontendOrigin https://SEU_USUARIO.github.io
```

Se preferir informar um hostname especifico no log:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\start-selfhosted-exact-mode.ps1 -TunnelName simulador-api -PublicHostname api.seudominio.com -FrontendOrigin https://SEU_USUARIO.github.io
```

O script:

1. inicia a API local
2. inicia `cloudflared tunnel run`
3. mostra instrucoes de parada

## Rodar em segundo plano (sem janela aberta)

Para iniciar em segundo plano:

```powershell
npm run exact:selfhosted:bg:start
```

Para parar os processos de segundo plano:

```powershell
npm run exact:selfhosted:bg:stop
```

## Iniciar automaticamente com o Windows (Task Scheduler + fallback sem admin)

Instalar tarefa de inicializacao no logon do seu usuario:

```powershell
npm run exact:selfhosted:task:install
```

Remover tarefa de inicializacao:

```powershell
npm run exact:selfhosted:task:uninstall
```

Observacao: a tarefa inicia no logon do usuario atual (modo interativo), que e o modo mais compativel para automacao do Excel COM.
Se o Windows negar permissao para criar Scheduled Task, o script cria automaticamente um fallback em `HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run` com o mesmo comando de inicializacao.

## Etapa 9 - Teste fim a fim

1. Abra `https://api.seudominio.com/health`
2. Deve retornar `{"ok":true}`
3. Abra o GitHub Pages do frontend
4. Rode uma simulacao
5. Confirme o download do arquivo mensal gerado pela API

## Parar os processos

Se usou o script do projeto, pressione `Ctrl+C` na janela onde ele esta rodando.

## Modo gratuito com dominio proprio

Esse fluxo tem custo mensal zero, mas depende de:

1. PC ligado
2. Excel instalado
3. internet ativa
4. tunnel em execucao

Sem backend acessivel, o botao de simulacao fica indisponivel no frontend.

## Alternativa temporaria sem trocar DNS

Se voce ainda nao quiser mover o DNS para a Cloudflare, pode usar Quick Tunnel apenas para teste:

```powershell
cloudflared tunnel --url http://127.0.0.1:3000
```

Nesse caso a URL sera aleatoria e mudara a cada execucao.
