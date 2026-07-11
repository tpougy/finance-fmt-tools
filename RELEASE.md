# Runbook de Release — Finance Fmt Tools

Guia para criar uma nova release do add-in "Finance Fmt Tools". Destinado a uma pessoa ou a um
agente de IA que precisa publicar uma nova versão — todos os comandos abaixo são executáveis via
terminal, sem depender do GitHub Actions.

> **Atenção — repositório público:** `tpougy/finance-fmt-tools` é um repositório **público** no
> GitHub. O passo `git push origin main` deste runbook publica ali, pela primeira vez, todo o
> histórico local da migração VBA → C# (incluindo commits internos de bookkeeping do fluxo GSD).
> Isso é esperado e aceito para este projeto (nenhum segredo/credencial existe nesse histórico —
> apenas código-fonte e documentação de planejamento) — mas não é reversível sem reescrever
> histórico público, então confira se é isso mesmo que você quer antes de rodar o passo 5 abaixo.

---

## Ferramentas necessárias

| Ferramenta | Versão mínima | Finalidade |
|---|---|---|
| **GitHub CLI (`gh`)** | 2.0+ | Criar/consultar releases via linha de comando |
| **.NET 8 SDK** | 8.0+ | Compilar e testar o projeto (`dotnet restore`/`build`/`test`) |
| **PowerShell** | 5.1+ | Empacotamento (`Compress-Archive`) e scripts de instalação |
| **Windows** | — | Necessário para compilar o add-in COM (.NET Framework 4.8, net48) |

Autenticação do `gh` (uma vez por máquina):

```powershell
gh auth login
gh auth status
gh --version
```

---

## Fluxo de release

1. **Rodar os testes automatizados** (devem passar antes de qualquer coisa):

   ```powershell
   dotnet test src\FinanceFmtTools.Engine.Tests\FinanceFmtTools.Engine.Tests.csproj -c Release
   ```

2. **Atualizar o changelog** — edite `RELEASE_NOTES.md` com o conteúdo da nova versão (esse arquivo
   é o corpo publicado da release, tanto pelo workflow automático quanto pelo fallback manual deste
   runbook).

3. **Atualizar os dois campos de versão hardcoded** para o novo `vX.Y.Z`:
   - `<Version>` em `src\FinanceFmtTools.ComAddin\FinanceFmtTools.ComAddin.csproj`
   - a constante `AddinVersion` em `src\FinanceFmtTools.ComAddin\AddInHost.cs` (usada pelo diálogo
     "Sobre" da Ribbon)

   Esta é uma etapa **manual** desta checklist, não um parâmetro de build automático
   (`-p:Version=`) — mantém o mesmo comportamento hoje já usado pelo projeto irmão
   `outlook-classic-delay-send`.

4. **Compilar e empacotar localmente** (mesma sequência usada pelo workflow de CI):

   ```powershell
   $ErrorActionPreference = 'Stop'

   dotnet restore src\FinanceFmtTools.sln
   dotnet build src\FinanceFmtTools.sln -c Release --no-restore
   dotnet test src\FinanceFmtTools.Engine.Tests\FinanceFmtTools.Engine.Tests.csproj -c Release --no-build

   $binSrc = "src\FinanceFmtTools.ComAddin\bin\Release\net48"
   New-Item -ItemType Directory -Path staging -Force | Out-Null

   # Cada entrada precisa existir -- Copy-Item sozinho falha "soft" (escreve no stream
   # de erro sem lançar) quando a origem não existe, o que geraria um
   # FinanceFmtTools.zip incompleto sem nenhum sinal de falha.
   $required = @(
     "$binSrc\FinanceFmtTools.ComAddin.dll",
     "$binSrc\FinanceFmtTools.Engine.dll",
     "$binSrc\Microsoft.Office.Interop.Excel.dll",
     "$binSrc\office.dll",
     "scripts\install.ps1",
     "scripts\uninstall.ps1",
     "scripts\verify-environment.ps1"
   )
   foreach ($f in $required) {
     if (-not (Test-Path -LiteralPath $f)) { throw "Arquivo obrigatório ausente para o pacote: $f" }
     Copy-Item -LiteralPath $f -Destination staging\ -Force
   }

   # Nome FIXO e literal -- scripts/install.ps1 depende de "FinanceFmtTools.zip"
   # nunca mudar de nome entre releases (URL .../releases/latest/download/ fixa).
   Compress-Archive -Path staging\* -DestinationPath FinanceFmtTools.zip -Force

   # Confere que o zip resultante contém exatamente os 7 arquivos esperados antes de
   # publicar (manual ou automático) -- não confie apenas em "o script rodou sem erro".
   $zipEntries = (Get-ChildItem staging -File -Name) | Sort-Object
   $expected   = ($required | ForEach-Object { Split-Path $_ -Leaf }) | Sort-Object
   if (Compare-Object $zipEntries $expected) { throw "staging\ não contém exatamente os arquivos esperados." }
   ```

5. **Criar a tag git e publicar** — nesta ordem exata (`main` primeiro, tag depois):

   ```powershell
   git tag -a vX.Y.Z -m "vX.Y.Z — descrição curta da versão"
   git push origin main
   git push origin vX.Y.Z
   ```

   > `git push origin main` deve rodar **antes** de `git push origin vX.Y.Z`. Se a tag for
   > enviada sem `main` ter sido enviado primeiro, o commit apontado pela tag fica órfão de
   > qualquer branch publicado em `origin` — o push da tag ainda funciona e ainda dispara o
   > workflow (uma tag carrega seus próprios objetos de commit), mas o GitHub mostra a release
   > como "sem branch correspondente", o que é confuso para quem for auditar o histórico depois.

### Caminho automático (recomendado)

Ao rodar `git push origin vX.Y.Z` (passo 5 acima), o GitHub Actions dispara automaticamente
`.github/workflows/release.yml` em `windows-latest`, que roda a mesma sequência
restore/build/test/package do passo 4 e publica a release com o `FinanceFmtTools.zip` gerado,
usando `RELEASE_NOTES.md` como corpo da release. Nenhuma ação adicional é necessária além do push
da tag.

### Caminho manual — `gh release create` (REL-02, zero dependência de CI)

Se o GitHub Actions estiver indisponível, ou se você quiser publicar a release imediatamente sem
esperar/disparar um workflow, rode o comando abaixo diretamente — ele fala com a API de Releases do
GitHub sem nenhuma dependência do Actions:

```powershell
gh release create vX.Y.Z FinanceFmtTools.zip --title "Finance Fmt Tools vX.Y.Z" -F RELEASE_NOTES.md
```

- `vX.Y.Z` — a mesma tag criada/enviada no passo 5.
- `FinanceFmtTools.zip` — o zip gerado no passo 4 (nome fixo, sem sufixo de versão).
- `-F RELEASE_NOTES.md` — usa o conteúdo do arquivo como corpo da release (mesmo arquivo lido pelo
  `body_path` do workflow automático).

## Verificação

```powershell
gh release view vX.Y.Z
```

Ou acesse: `https://github.com/tpougy/finance-fmt-tools/releases/tag/vX.Y.Z`

Confirme que o asset `FinanceFmtTools.zip` está anexado e que o corpo da release corresponde ao
conteúdo de `RELEASE_NOTES.md`.

---

## Comandos úteis

```powershell
# Listar releases existentes
gh release list

# Ver detalhes de uma release específica
gh release view vX.Y.Z

# Baixar o asset de uma release
gh release download vX.Y.Z

# Deletar uma release (não apaga a tag)
gh release delete vX.Y.Z

# Apagar uma tag localmente e remotamente
git tag -d vX.Y.Z
git push origin :refs/tags/vX.Y.Z
```

---

## Checklist rápido de release

```
[ ] Testes passando: dotnet test src\FinanceFmtTools.Engine.Tests\FinanceFmtTools.Engine.Tests.csproj -c Release
[ ] RELEASE_NOTES.md atualizado com o changelog da nova versão
[ ] <Version> atualizado em FinanceFmtTools.ComAddin.csproj
[ ] AddinVersion atualizado em AddInHost.cs
[ ] Build + empacotamento local gerou FinanceFmtTools.zip
[ ] Tag criada: git tag -a vX.Y.Z -m "..."
[ ] Push do branch: git push origin main (ANTES da tag)
[ ] Push da tag: git push origin vX.Y.Z
[ ] Caminho automático: workflow do GitHub Actions concluído sem erros -- OU
    Caminho manual: gh release create vX.Y.Z FinanceFmtTools.zip --title "Finance Fmt Tools vX.Y.Z" -F RELEASE_NOTES.md
[ ] Verificado: gh release view vX.Y.Z mostra o asset e as notas corretas
[ ] Testado o instalador: irm https://raw.githubusercontent.com/tpougy/finance-fmt-tools/main/scripts/install.ps1 | iex
```
