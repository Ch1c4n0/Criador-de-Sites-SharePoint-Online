# Criador de Sites — SharePoint Online

<p align="center">
  <img src="https://img.shields.io/badge/SharePoint%20Online-0078D4?style=for-the-badge&logo=microsoftsharepoint&logoColor=white" alt="SharePoint Online" />
  <img src="https://img.shields.io/badge/PowerShell%207+-5391FE?style=for-the-badge&logo=powershell&logoColor=white" alt="PowerShell" />
  <img src="https://img.shields.io/badge/PnP%20PowerShell-512BD4?style=for-the-badge&logo=powershell&logoColor=white" alt="PnP PowerShell" />
</p>

<p align="center">
  <img src="https://raw.githubusercontent.com/pnp/media/master/parker/pnp/300w/parker.png" alt="PnP" height="90" />
</p>

Script PowerShell para criar sites no SharePoint Online de **duas formas**:

1. **Em lote** — a partir de um arquivo CSV.
2. **Individual** — informando os dados de um único site.

Em ambos os modos é possível definir **owners**, **membros** e a **privacidade**
(Público / Privado), além de escolher o **tipo de site**:

- **Team Site** — Site de Equipe conectado a um Grupo do Microsoft 365
  (tem o conceito de Público/Privado, owners e membros do grupo).
- **Communication Site** — Site de Comunicação (sem grupo M365; owners viram
  administradores do site e membros entram no grupo "Membros" do SharePoint).

---

## Arquivos

| Arquivo                  | Descrição                                                        |
|--------------------------|------------------------------------------------------------------|
| `Criar-Sites.ps1`        | Script principal (**PnP.PowerShell**) — recomendado.             |
| `modelo-sites.csv`       | Modelo de CSV do script PnP.                                     |
| `Criar-Sites-SPO.ps1`    | Script alternativo (**SPO Management Shell**) — mais limitado.   |
| `modelo-sites-spo.csv`   | Modelo de CSV do script SPO.                                     |
| `README.md`              | Este guia.                                                       |

> Existem **dois scripts**. O principal usa **PnP.PowerShell** (Team Site com
> Grupo M365, público/privado, owners e membros). O alternativo usa o módulo
> oficial **SPO Management Shell** — veja a seção *"Script alternativo (SPO)"*
> no final para saber quando usar e suas limitações.

---

## Pré-requisitos

- **PowerShell 7+** (recomendado) ou Windows PowerShell 5.1.
- Módulo **PnP.PowerShell** — o script instala automaticamente se não existir.
  Para instalar manualmente:
  ```powershell
  Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
  ```
- Permissão de **Administrador do SharePoint** (ou Administrador Global) no tenant.
- A URL de administração do SharePoint, no formato:
  `https://SEU_TENANT-admin.sharepoint.com`

### Permitir execução de scripts (uma vez por máquina)

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### Registro do app para login interativo (PnP.PowerShell 2.x)

As versões recentes do PnP.PowerShell **não têm mais** um app padrão para login
interativo — por isso, sem um `-ClientId`, você verá o erro
`Specified method is not supported` e o aviso *"Please specify a valid client id"*.
É preciso registrar um app no Entra ID **uma vez** por tenant.

> ⚠️ **Quem pode registrar o app:** a conta usada neste registro precisa ser
> **Administrador Global** ou **Administrador do Entra ID** (Application
> Administrator / Cloud Application Administrator) do tenant — é ela que vai
> **criar o app e dar o consentimento** das permissões no navegador.
> **Não** é necessário ser administrador da máquina (não precisa abrir o
> PowerShell "como Administrador").

Rode no **PowerShell 7 (`pwsh`)**:

```powershell
Import-Module PnP.PowerShell

Register-PnPEntraIDAppForInteractiveLogin `
    -ApplicationName "PnP-CriarSites" `
    -Tenant "SEU_TENANT.onmicrosoft.com" `
    -SharePointDelegatePermissions "AllSites.FullControl" `
    -GraphDelegatePermissions "Group.ReadWrite.All","User.ReadWrite.All"
```

- Abre o navegador para você **logar e consentir** as permissões.
- Ao final, mostra um **ClientId** (um GUID). Copie e use no script com `-ClientId`.
- O app pode levar alguns minutos para propagar após o consentimento.

#### Permissões da API (o que o app precisa)

O comando acima já registra o app com este conjunto **delegado**, que cobre tudo
o que o script faz (criar sites, criar grupos M365, adicionar owners e membros):

| API              | Permissão (delegada)   | Para quê                                              |
|------------------|------------------------|-------------------------------------------------------|
| SharePoint       | `AllSites.FullControl` | Criar/gerenciar sites e grupos do SharePoint.         |
| Microsoft Graph  | `Group.ReadWrite.All`  | Criar o Grupo M365 do Team Site; add owners/membros.  |
| Microsoft Graph  | `User.ReadWrite.All`   | Resolver e adicionar usuários (owners/membros).       |

> São permissões **delegadas**: o app age **em nome do usuário** que faz login.
> Ou seja, além do consentimento do app, a conta que executa o script ainda
> precisa ser **Administrador do SharePoint / Admin Global** para criar sites.
>
> Se você omitir os parâmetros `-SharePointDelegatePermissions` e
> `-GraphDelegatePermissions`, o cmdlet aplica um conjunto padrão um pouco maior
> (inclui também `TermStore.ReadWrite.All`) — também funciona.

#### Permissões para o modo App-Only (`-AuthMode AppOnly`)

No App-Only não há usuário logado, então as permissões precisam ser do tipo
**Application** (com **consentimento de administrador**):

| API              | Permissão (application) |
|------------------|--------------------------|
| SharePoint       | `Sites.FullControl.All`  |
| Microsoft Graph  | `Group.ReadWrite.All`    |
| Microsoft Graph  | `User.ReadWrite.All`     |

Use os parâmetros `-SharePointApplicationPermissions` e
`-GraphApplicationPermissions` (e `-CertificatePath`) ao registrar, ou configure
manualmente no portal do Entra ID e conceda o consentimento de admin.

---

## O que o script faz (passo a passo)

Ao ser executado, o `Criar-Sites.ps1` segue este fluxo:

1. **Verifica o PowerShell**: exige PS7 e **interrompe** se for o Windows PowerShell 5.1.
2. **Cria um log** da execução em `.\logs\Criar-Sites_<data>_<hora>.log`.
3. **Garante o módulo**: instala o `PnP.PowerShell` se faltar (com TLS 1.2 + NuGet) e o carrega.
4. **Conecta** ao tenant com `Connect-PnPOnline` (interativo com `-ClientId`, ou App-Only).
5. **Escolhe o modo**:
   - Se você passar `-CsvPath`, vai **direto para o lote**.
   - Senão, mostra o **menu** com as duas opções (detalhadas abaixo).
6. **Para cada site** (em qualquer modo), o núcleo `New-SharePointSite`:
   - **Team Site** → cria a coleção + **Grupo M365** com `New-PnPSite -Type TeamSite`,
     define **Público/Privado**, espera o grupo provisionar (até `-GroupWaitSeconds`)
     e adiciona **owners e membros** do grupo via Microsoft Graph.
   - **Communication Site** → cria com `New-PnPSite -Type CommunicationSite`, define
     os **owners como admins do site** e os **membros no grupo "Membros"** do SharePoint.
   - Aplica **Sensitivity Label/Classificação** se informados; repete operações com
     throttling (429/503) com **backoff** (`-MaxRetries`).
7. **Ao final do lote**: exibe um **resumo** (Criado / ERRO) e grava um `*.resultado.csv`.

### As duas opções

O script tem **dois modos de criar sites** — é a essência dele:

#### 🔹 Opção 1 — Em LOTE (a partir de um CSV)

Cria **vários sites de uma vez** lendo o arquivo [modelo-sites.csv](modelo-sites.csv)
(ou outro que você indicar). Cada linha do CSV vira um site. Antes de criar, o script
**valida o arquivo inteiro** e, se houver erro, **não cria nada**. Ideal para provisionar
muitos sites com Owners/Membros/privacidade já definidos na planilha. É o que você usa
para trabalho em volume.

#### 🔹 Opção 2 — Site ÚNICO (interativo)

Cria **um site por vez**, perguntando os dados na tela (Título, Tipo, Alias/URL,
Privacidade, Owners, Membros, Descrição). Não precisa de CSV. Ideal para **um site
avulso** ou para **testar** o fluxo antes de rodar um lote grande.

> Os dois modos usam **a mesma lógica de criação** por baixo — a diferença é só a
> **origem dos dados** (planilha CSV × perguntas na tela).

---

## Como usar

> **Importante:** rode sempre no **PowerShell 7 (`pwsh`)**, não no Windows
> PowerShell 5.1. E o `-ClientId` (gerado no registro do app) é **obrigatório**
> no login interativo — sem ele você recebe `Specified method is not supported`.

Abra o PowerShell 7 na pasta do script:

```powershell
pwsh
cd "D:\sharepoint\Criar Sites"
```

### Opção A — Menu interativo (login no navegador)

```powershell
.\Criar-Sites.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -ClientId "00000000-0000-0000-0000-000000000000"
```

Aparecerá um menu:

```
  1) Criar sites EM LOTE (a partir de um CSV)
  2) Criar um site UNICO
  3) Sair
```

- **Opção 1** pede o caminho do CSV (padrão: `.\modelo-sites.csv`).
- **Opção 2** pergunta, passo a passo, os dados de um único site.

### Opção B — Lote direto (sem menu)

```powershell
.\Criar-Sites.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -ClientId "00000000-0000-0000-0000-000000000000" `
    -CsvPath ".\modelo-sites.csv"
```

### Opção C — Autenticação App-Only (automação / agendado)

Sem interação, usando ClientId + certificado:

```powershell
.\Criar-Sites.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -CsvPath ".\modelo-sites.csv" `
    -AuthMode AppOnly `
    -ClientId "00000000-0000-0000-0000-000000000000" `
    -Tenant "contoso.onmicrosoft.com" `
    -Thumbprint "ABCDEF0123456789ABCDEF0123456789ABCDEF01"
```

Ou usando um arquivo `.pfx` em vez do thumbprint:

```powershell
$pwd = Read-Host "Senha do PFX" -AsSecureString
.\Criar-Sites.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -CsvPath ".\modelo-sites.csv" `
    -AuthMode AppOnly `
    -ClientId "00000000-0000-0000-0000-000000000000" `
    -Tenant "contoso.onmicrosoft.com" `
    -CertificatePath ".\certificado.pfx" `
    -CertificatePassword $pwd
```

### Parâmetros opcionais de ajuste (lote)

Para lotes grandes ou tenants com regras específicas, há parâmetros extras:

| Parâmetro            | Padrão | Para quê                                                                 |
|----------------------|--------|--------------------------------------------------------------------------|
| `-DelaySeconds`      | `5`    | Pausa entre a criação de cada site, reduz **throttling** em lotes grandes. |
| `-MaxRetries`        | `5`    | Tentativas com **backoff exponencial** quando o SharePoint retorna 429/503. |
| `-GroupWaitSeconds`  | `120`  | Tempo máximo de espera pelo provisionamento do Grupo M365 (Team Site).   |
| `-SensitivityLabel`  | —      | Rótulo de confidencialidade aplicado a **todos** os sites da execução.   |
| `-LogPath`           | `.\logs` | Pasta onde o log e o resumo `.resultado.csv` são gravados.             |

Exemplo para um lote grande e mais tolerante:

```powershell
.\Criar-Sites.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -ClientId "00000000-0000-0000-0000-000000000000" `
    -CsvPath ".\modelo-sites.csv" `
    -DelaySeconds 10 -MaxRetries 8 -GroupWaitSeconds 180
```

### Logs

Cada execução gera um arquivo em `.\logs\Criar-Sites_<data>_<hora>.log` com todos
os passos, e no modo lote também um `*.resultado.csv` com o status de cada site
(`Criado` ou `ERRO: ...`).

---

## Formato do CSV

O delimitador é **ponto e vírgula (`;`)**. Veja `modelo-sites.csv`.

| Coluna        | Obrigatório            | Descrição                                                                 |
|---------------|------------------------|---------------------------------------------------------------------------|
| `Titulo`      | Sim                    | Nome de exibição do site.                                                 |
| `Tipo`        | Sim                    | `Team` ou `Communication`.                                                |
| `Alias`       | Sim (apenas Team)      | Nome curto do grupo M365 (sem espaços/acentos). Ex.: `marketing-equipe`.  |
| `Url`         | Sim (apenas Comm.)     | URL completa **ou** apenas o sufixo. Ex.: `rh-comunicados`.               |
| `Privacidade` | Não (Team; padrão Private) | `Public` ou `Private`. Ignorado em Communication Site.                |
| `Owners`      | Recomendado            | E-mails separados por `;` (ou `,`). Ex.: `ana@contoso.com;bruno@contoso.com`. |
| `Members`     | Não                    | E-mails separados por `;` (ou `,`).                                       |
| `Descricao`   | Não                    | Descrição do site.                                                        |
| `SensitivityLabel` | Não               | Id ou nome de um rótulo de confidencialidade a aplicar no site/grupo.     |
| `Classificacao`    | Não               | Classificação do grupo (se o tenant usar classificações herdadas).        |

> **Atenção ao separar e-mails dentro de uma coluna:** como o delimitador do CSV
> é `;`, use **vírgula** para separar múltiplos e-mails na mesma célula se for
> editar no Excel — o script aceita tanto `;` quanto `,`. No modelo em texto puro,
> os e-mails estão separados por `;` em colunas próprias.
>
> As colunas `SensitivityLabel` e `Classificacao` podem ficar **vazias**. Se o seu
> tenant **exige** rótulo de confidencialidade na criação de grupos, preencha a
> coluna (ou use o parâmetro global `-SensitivityLabel`), senão a criação falha.

### Validação automática

Antes de criar qualquer site, o script **valida o CSV inteiro** e **aborta sem
criar nada** se encontrar problemas, listando cada um: `Tipo` inválido, `Titulo`
vazio, `Alias`/`Url` faltando ou **duplicados no arquivo**, e e-mails malformados.

### Exemplo (conteúdo de `modelo-sites.csv`)

```csv
Titulo;Tipo;Alias;Url;Privacidade;Owners;Members;Descricao;SensitivityLabel;Classificacao
Marketing Equipe;Team;marketing-equipe;;Private;ana@contoso.com;bruno@contoso.com;Site da equipe de marketing;;
Comunicados RH;Communication;;rh-comunicados;;maria@contoso.com;;Portal de comunicados do RH;;
Projeto Alpha;Team;projeto-alpha;;Public;ana@contoso.com;bruno@contoso.com;Colaboracao do Projeto Alpha;;
```

---

## Notas e comportamento

- **PowerShell 7 obrigatório**: o script verifica a versão no início e **interrompe**
  com mensagem clara se for executado no Windows PowerShell 5.1.
- **Team Site**: o Grupo M365 é provisionado de forma assíncrona; o script aguarda
  até `-GroupWaitSeconds` (padrão **120s**) antes de adicionar membros. Owners
  informados já entram na criação e são reforçados depois.
- **Communication Site**: não possui grupo M365. Os **owners** são definidos como
  **administradores do site** e os **membros** são adicionados ao grupo associado
  "Membros" do SharePoint.
- **Throttling**: operações que retornam 429/503 são **repetidas automaticamente**
  com backoff exponencial (até `-MaxRetries`). Entre sites do lote há uma pausa de
  `-DelaySeconds`.
- **Validação prévia**: o CSV é checado por inteiro antes de criar — se houver erro,
  **nada é criado**.
- **Tolerância a falhas**: passada a validação, se uma linha falhar na criação, o
  script registra o erro e continua com as demais. Ao final, exibe um **resumo** e
  grava um `*.resultado.csv`.
- **URLs**: para Communication Site você pode informar a URL completa
  (`https://contoso.sharepoint.com/sites/rh`) ou só o sufixo (`rh`).

---

## Solução de problemas

| Problema                                          | Solução                                                                 |
|---------------------------------------------------|-------------------------------------------------------------------------|
| `requer PowerShell 7+`                            | Abra o **`pwsh`** (PowerShell 7), não o Windows PowerShell 5.1.         |
| `Specified method is not supported` + *"Please specify a valid client id"* | Falta o `-ClientId`. Registre o app (acima) e passe o ClientId ao script. |
| `Connect-PnPOnline ... AADSTS65001` (sem consentimento) | Execute `Register-PnPEntraIDAppForInteractiveLogin` e use o `-ClientId`. |
| Script não executa (`...não pode ser carregado`)  | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.                  |
| `Corrija o CSV e tente novamente`                 | A validação achou problemas (listados acima dela). Ajuste o CSV.        |
| Erros `429` / *throttling* constantes             | Aumente `-DelaySeconds` e `-MaxRetries`; rode em horários de menor uso. |
| Alias já em uso                                   | Escolha outro `Alias` (único no tenant; aliases excluídos ficam ~30 dias na lixeira). |
| Membros não aparecem no Team Site                 | Grupo demorou a provisionar; aumente `-GroupWaitSeconds` ou rode de novo. |
| Falha exigindo rótulo/classificação               | Preencha `SensitivityLabel`/`Classificacao` no CSV ou use `-SensitivityLabel`. |
| Sem permissão                                     | Confirme que a conta é Admin do SharePoint / Admin Global.              |

---

## Script alternativo (SPO) — `Criar-Sites-SPO.ps1`

Versão que usa **apenas o módulo oficial** `Microsoft.Online.SharePoint.PowerShell`
(SPO Management Shell), sem PnP. Útil quando você quer ficar no módulo da Microsoft
e **não precisa** de Grupo M365.

### O que o script faz (passo a passo)

Ao ser executado, ele:

1. **Cria um log** da execução em `.\logs\Criar-Sites-SPO_<data>_<hora>.log`.
2. **Avisa as limitações** do modo SPO (sem público/privado, sem membros, 1 owner).
3. **Garante o módulo**: instala `Microsoft.Online.SharePoint.PowerShell` se faltar
   e o carrega (no PowerShell 7, em modo de compatibilidade `-UseWindowsPowerShell`).
4. **Conecta** ao tenant com `Connect-SPOService` (login no navegador, sem ClientId).
5. **Escolhe o modo**:
   - Se você passar `-CsvPath`, vai **direto para o lote**.
   - Senão, mostra um **menu** com: `1) Lote (CSV)`, `2) Site único`, `3) Sair`.
6. **No lote**: lê o CSV, **valida o arquivo inteiro antes de criar** (tipo, URL,
   owner, duplicados, e-mails) e **aborta sem criar nada** se houver erro; passando
   na validação, cria cada site com uma **pausa** (`-DelaySeconds`) entre eles.
7. **Para cada site**:
   - Cria a coleção com `New-SPOSite` (template conforme o `Tipo`), definindo
     **Título**, **owner principal** e **cota de armazenamento**.
   - Adiciona os **admins extras** (owners adicionais) com `Set-SPOUser
     -IsSiteCollectionAdmin`.
   - Operações com erro transitório (throttling 429/503) são **repetidas** com
     backoff (`-MaxRetries`).
8. **Ao final do lote**: exibe um **resumo** (Criado / ERRO) e grava um
   `*.resultado.csv` ao lado do log.

### Vantagens

- **Login mais simples**: `Connect-SPOService` abre o navegador e **não exige
  registrar app/ClientId** no Entra ID.
- Permite definir **cota de armazenamento** (`StorageQuota`) por site.

### ⚠️ Limitações (importantes)

| Recurso                        | PnP (`Criar-Sites.ps1`) | SPO (`Criar-Sites-SPO.ps1`) |
|--------------------------------|-------------------------|------------------------------|
| Público / Privado              | ✅ Sim                  | ❌ Não (não há Grupo M365)   |
| Team Site com Grupo M365       | ✅ Sim                  | ❌ Não (Team Site **sem** grupo) |
| Múltiplos owners               | ✅ Sim                  | ⚠️ 1 owner + extras como **admins** |
| Membros                        | ✅ Sim                  | ❌ **Não suportado** (sem cmdlet; coluna nem existe no CSV do SPO) |
| Descrição do site              | ✅ Sim                  | ❌ Não suportado             |
| Cota de armazenamento          | ➖                      | ✅ Sim                       |

> Por isso o **PnP continua sendo o recomendado** para o requisito de owners +
> membros + público/privado. Use o SPO só para criação simples de coleções.

### Tipos de site (templates)

| `Tipo`          | Template usado          |
|-----------------|-------------------------|
| `Team`          | `STS#3` (Team Site sem grupo) |
| `Communication` | `SITEPAGEPUBLISHING#0`  |

### Instalação do módulo (automática e manual)

O módulo `Microsoft.Online.SharePoint.PowerShell` é baseado em **.NET Framework** e
roda no **Windows PowerShell 5.1**. No **PowerShell 7** ele é carregado em modo de
compatibilidade (`-UseWindowsPowerShell`), que abre uma sessão do Windows PowerShell
por baixo — por isso o módulo precisa estar instalado **para o Windows PowerShell**,
e não para o PS7.

**Automático:** o script já cuida disso. Quando rodado no PS7, ele verifica e instala
o módulo **dentro do contexto do `powershell.exe`** (Windows PowerShell), no lugar
certo. Na primeira execução isso leva ~1 min (mostra barra de progresso).

> ⚠️ Se você instalar com `Install-Module` **dentro do PowerShell 7**, o módulo vai
> para o caminho do PS7 e o import com `-UseWindowsPowerShell` **não o encontra**,
> resultando em: *"no valid module file was found in any module directory"*. Por isso
> a instalação tem que ser feita no Windows PowerShell (o script faz isso por você).

**Manual (opcional):** para adiantar e não esperar o download na primeira execução,
abra o **Windows PowerShell** (procure por *"Windows PowerShell"* no menu Iniciar —
**não** o "PowerShell 7"/`pwsh`) e rode:

```powershell
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
```

Para conferir se ficou instalado no contexto certo:

```powershell
# No Windows PowerShell (5.1):
Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell
```

Se preferir instalar para **todos os usuários** da máquina (requer abrir o Windows
PowerShell **como Administrador**):

```powershell
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope AllUsers -Force
```

### Como usar

```powershell
# Pode rodar no PowerShell 7 (importa em modo de compatibilidade) ou no 5.1.
cd "D:\sharepoint\Criar Sites"

# Menu interativo
.\Criar-Sites-SPO.ps1 -AdminUrl "https://contoso-admin.sharepoint.com"

# Lote direto
.\Criar-Sites-SPO.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -CsvPath ".\modelo-sites-spo.csv"
```

### Formato do CSV (`modelo-sites-spo.csv`)

Delimitador **`;`**. Este CSV é **específico do modo SPO** e **não tem coluna
`Members`** (o SPO Management Shell não adiciona membros). Para sites com membros,
use o `Criar-Sites.ps1` (PnP) e o `modelo-sites.csv`.

| Coluna         | Obrigatório | Descrição                                                       |
|----------------|-------------|------------------------------------------------------------------|
| `Titulo`       | Sim         | Nome de exibição do site.                                        |
| `Tipo`         | Sim         | `Team` ou `Communication`.                                       |
| `Url`          | Sim         | URL completa **ou** apenas o sufixo (ex.: `marketing`).          |
| `Owner`        | Sim         | **Um** e-mail (owner principal).                                 |
| `Admins`       | Não         | Owners adicionais → viram **admins da coleção** (`;` ou `,`).    |
| `StorageQuota` | Não         | Cota em MB (padrão `-DefaultStorageQuota`, 1024).               |

```csv
Titulo;Tipo;Url;Owner;Admins;StorageQuota
Marketing Equipe;Team;marketing-equipe;ana@contoso.com;bruno@contoso.com;1024
Comunicados RH;Communication;rh-comunicados;maria@contoso.com;;2048
Projeto Alpha;Team;projeto-alpha;ana@contoso.com;bruno@contoso.com,carla@contoso.com;1024
```

### Parâmetros úteis

`-DefaultStorageQuota` (MB, padrão 1024), `-DelaySeconds` (pausa entre sites),
`-MaxRetries` (retry de throttling), `-LogPath` (pasta de logs).

---

## Qual script usar? (PnP × SPO)

### Decisão rápida

| Sua necessidade                                            | Use                         |
|------------------------------------------------------------|-----------------------------|
| Precisa de **owners E membros**                            | **PnP** (`Criar-Sites.ps1`) |
| Precisa de site **Público ou Privado**                     | **PnP**                     |
| **Team Site conectado a Grupo M365 / Teams**               | **PnP**                     |
| Precisa de **rótulo de confidencialidade / classificação** | **PnP**                     |
| **Vários owners** por site                                 | **PnP**                     |
| Só **Communication Site / coleção simples**, sem grupo     | SPO (`Criar-Sites-SPO.ps1`) |
| Precisa definir **cota de armazenamento** por site         | SPO                         |
| **Não pode/quer registrar app** no Entra ID (login simples)| SPO                         |
| Ambiente preso ao **Windows PowerShell 5.1**               | SPO                         |

> **Regra geral:** na dúvida, use o **PnP** — ele cobre tudo que o seu requisito
> original pede (owners + membros + público/privado). O **SPO** é para cenários
> mais simples ou restrições de ambiente/permissão.

### Casos de exemplo

**Caso 1 — Onboarding de equipes (owners + membros + privacidade) → PnP**
> "Preciso criar 30 sites de equipe; cada um com 2 owners e vários membros, alguns
> privados e outros públicos."

```powershell
.\Criar-Sites.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -ClientId "SEU_CLIENT_ID" `
    -CsvPath ".\modelo-sites.csv" `
    -DelaySeconds 10 -GroupWaitSeconds 180
```

**Caso 2 — Portais de comunicação da intranet com cota → SPO**
> "Quero criar 10 Communication Sites para a intranet, cada um com um responsável
> e uma cota de armazenamento definida. Não preciso de membros nem de grupo."

```powershell
.\Criar-Sites-SPO.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -CsvPath ".\modelo-sites-spo.csv" `
    -DefaultStorageQuota 2048
```

**Caso 3 — Sem permissão para registrar app no Entra ID → SPO**
> "Sou Admin do SharePoint, mas não consigo registrar um app no Entra ID (o login
> interativo do PnP exige isso). Quero algo que só peça meu login."

```powershell
# Connect-SPOService abre o navegador, sem ClientId.
.\Criar-Sites-SPO.ps1 -AdminUrl "https://contoso-admin.sharepoint.com"
```

**Caso 4 — Site único de equipe com governança (rótulo) → PnP**
> "Só um site de equipe novo, privado, com um rótulo de confidencialidade da empresa."

```powershell
.\Criar-Sites.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -ClientId "SEU_CLIENT_ID" `
    -SensitivityLabel "Confidencial"
# depois escolha a opção 2 (site único) no menu
```

**Caso 5 — Provisionamento rápido de coleções no Windows PowerShell 5.1 → SPO**
> "Estou numa máquina antiga só com Windows PowerShell 5.1 e preciso criar várias
> coleções de site rapidamente."

```powershell
# O script SPO roda no 5.1 nativamente (o PnP exigiria PS7).
.\Criar-Sites-SPO.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -CsvPath ".\modelo-sites-spo.csv"
```

---

## Autor

Desenvolvido por **Marcelo Gonçalves**.
