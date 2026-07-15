# Deploy no Windows Server

Este projeto pode rodar inteiro no servidor Windows: tela web, API e calculo Excel.

## Requisitos

- Windows Server com Microsoft Excel instalado e ativado.
- Git instalado.
- Usuario logado via RDP. A automacao COM do Excel deve rodar em sessao interativa.
- Permissao de administrador para abrir firewall e registrar tarefa agendada.

## Instalacao

No servidor RDP, abra PowerShell como administrador e execute:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
git clone https://github.com/douglasdealmeidaribeiro/simulador.git C:\simulador
cd C:\simulador
.\deploy\install-server.ps1 -Port 80 -BindAddress 0.0.0.0 -PublicHost 131.72.140.23
```

Depois acesse:

```text
http://131.72.140.23/
```

Se a porta 80 nao estiver liberada no provedor/firewall externo, use outra porta:

```powershell
.\deploy\install-server.ps1 -Port 3000 -BindAddress 0.0.0.0 -PublicHost 131.72.140.23
```

E acesse:

```text
http://131.72.140.23:3000/
```

## GitHub Pages

Para usar a tela no GitHub Pages, o backend precisa de URL publica HTTPS. Sem HTTPS, o navegador bloqueia a chamada por mixed content.
