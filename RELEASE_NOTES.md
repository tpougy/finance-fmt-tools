# Release Notes

Changelog mantido manualmente. Cada release publicada (automática via `.github/workflows/release.yml`
ou manual via `gh release create`, ver `RELEASE.md`) usa o conteúdo deste arquivo como corpo da
release (`body_path`/`-F RELEASE_NOTES.md`). Este arquivo é sobrescrito com a entrada da próxima
versão antes de cada nova tag — o histórico de versões anteriores só fica preservado na página de
Releases do GitHub (editável de lá via `gh release edit`, não neste arquivo).

---

## Finance Fmt Tools v2.1.0

### Migração automática da versão VBA legada

O instalador (`scripts/install.ps1`) agora detecta sozinho uma instalação anterior da versão VBA
(`FinanceFmtTools.xlam` em `%APPDATA%\Microsoft\AddIns`), desregistra-a do Excel via automação COM
(`AddIns.Installed = $false`, evitando deixar o Excel com uma referência quebrada a um arquivo que
será apagado) e remove o arquivo do disco — tudo isso **antes** de instalar a versão C#, sem
nenhuma ação manual do usuário. Isso substitui a orientação manual publicada na v2.0.0 ("remova o
`.xlam` antes de instalar"). Se nenhuma instalação VBA for encontrada, essa etapa é pulada
silenciosamente (não abre o Excel à toa). Validado de ponta a ponta contra Excel real (Office 16.0):
detecção, desregistro, remoção do arquivo e instalação da versão C# em sequência, sem deixar
resíduo de nenhuma das duas versões.

### Ajuste no formato contábil

O token de padding "_-" (espaço do tamanho de um hífen), usado no final de cada seção do formato
contábil (`Fin 0D/2D/4D/8D`), foi removido — resta apenas o padding "_(" / "_)" (espaço do tamanho
de um parêntese) já usado para alinhar os dígitos entre as seções positiva/negativa/zero. Efeito
visual: uma célula formatada com `Fin 2D`, por exemplo, ganha um caractere a menos de espaço em
branco à direita. Comportamento interno reorganizado em `FormatTokens.cs`, um arquivo dedicado de
constantes nomeadas para cada peça combinável do formato — facilita ajustes futuros desse tipo sem
precisar mexer na lógica de montagem em `AccountingFormatBuilder.cs`.

### Correção de bug (encontrado em teste ao vivo)

Durante a validação da migração automática de VBA contra um Excel real, foi encontrada e corrigida
uma race condition: a rotina que desregistra o `.xlam` legado abre e fecha sua própria instância do
Excel via COM, mas não esperava o processo `EXCEL.EXE` correspondente realmente terminar antes de
devolver o controle — o que podia fazer a checagem seguinte ("Excel está aberto?") falhar
intermitentemente. Corrigido com uma espera limitada (até 15s) pelo término do processo.

---

## Finance Fmt Tools v2.0.1

**Correção crítica**: a v2.0.0 publicava um add-in que **nunca carregava de verdade no Excel**,
apesar de toda a suíte de testes automatizados (`dotnet test`) passar 100% e o pacote/instalador
funcionarem sem erro aparente.

### O que estava quebrado

O shim `Extensibility.IDTExtensibility2` (declarado à mão em
`src/FinanceFmtTools.ComAddin/Extensibility.cs`, já que não existe um pacote NuGet oficial e leve
para essa interface COM clássica) estava incompleto: faltavam os atributos `[DispId]` em cada
método e `[MarshalAs]`/`[In]` nos parâmetros (em especial `ref Array custom`, que precisa de
`UnmanagedType.SafeArray`), que a interface COM real do Office define. Isso quebrava o layout de
vtable que o carregador nativo de add-ins do Excel espera:

- `CoCreateInstance`/`QueryInterface` funcionavam normalmente (o add-in aparecia na lista de
  Suplementos COM);
- mas a chamada real a `OnConnection` nunca chegava ao código gerenciado — o Excel silenciosamente
  tratava isso como falha de carregamento e rebaixava `LoadBehavior` de `3` para `2` no primeiro
  uso, sem gerar exceção .NET nem entrada no Visualizador de Eventos do Windows.

Resultado prático para quem instalou a v2.0.0: o add-in aparecia em **Arquivo > Opções >
Suplementos**, mas a aba "Finance Fmt" nunca aparecia na Ribbon.

### Correção

`Extensibility.cs` agora replica byte-a-byte a assinatura da interface COM real (`DispId(1..5)`,
`MarshalAs(UnmanagedType.IDispatch)` para os parâmetros `object`, `MarshalAs(UnmanagedType.SafeArray)`
para `ref Array custom`), verificada por reflection contra uma cópia real de `Extensibility.dll`.
Reproduzido e confirmado corrigido em Excel real (Office 16.0, Click-to-Run x64): `LoadBehavior`
permanece em `3` e o add-in conecta (`Connect=True`) de forma estável após reinstalação.

### Ação recomendada

Se você instalou a v2.0.0, rode o instalador novamente (mesmo comando, sem nenhum parâmetro extra)
para atualizar para a v2.0.1 — o fluxo é idempotente:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force; irm https://raw.githubusercontent.com/tpougy/finance-fmt-tools/main/scripts/install.ps1 | iex
```

---

## Finance Fmt Tools v2.0.0

Esta é a primeira release da migração completa do add-in "Finance Fmt Tools" de VBA (`.xlam`) para
um add-in COM em C#. A experiência do usuário final na Ribbon do Excel é preservada — mesma aba
"Finance Fmt", mesmos grupos, mesmos botões, mesmos atalhos de formatação — mas toda a implementação
por trás foi reescrita em C#/.NET Framework 4.8, com cobertura de testes automatizados e um fluxo de
build/release 100% via terminal (`dotnet` CLI + GitHub Actions), sem depender de Visual Studio
completo.

> **Nota**: esta versão tinha um bug crítico que impedia o add-in de carregar de verdade no Excel —
> corrigido na v2.0.1 acima. Recomendamos atualizar diretamente para v2.0.1.

### O que mudou

- **Format engine portado para C#**: todos os botões de formatação (contábil `Fin 0D/2D/4D/8D`,
  percentual `% 2D`/`% 4D`, `Spread bps`, datas `ISO`/`BR`/`BR Extenso` e `Texto`) foram portados 1:1
  a partir do VBA original, com cobertura de testes xUnit automatizados (`dotnet test`) para as 16
  combinações do formato contábil (decimais × alinhamento × zero contábil).
- **Novo instalador/desinstalador HKCU**: `scripts/install.ps1` e `scripts/uninstall.ps1` substituem
  o antigo `Install-FinanceFmtTools.ps1` (que usava automação COM do Excel para instalar o `.xlam`
  em `%APPDATA%\Microsoft\AddIns`). O novo fluxo registra o add-in COM inteiramente em `HKCU`, sem
  exigir administrador e sem `regasm`.
- **Pipeline de release automatizado**: um push de tag `v*.*.*` agora dispara um workflow do GitHub
  Actions que compila, testa, empacota e publica a release automaticamente — nenhuma montagem manual
  de `.xlam` é mais necessária.

### Mudança de comportamento

As preferências dos checkboxes **"Alinhar à direita"** e **"Zero contábil"** não persistem mais
entre sessões do Excel. Na versão VBA anterior, esses valores eram salvos dentro do próprio arquivo
`.xlam` (via `CustomXMLPart`) e recarregados na próxima abertura. Nesta migração, os dois checkboxes
sempre iniciam nos valores padrão documentados a cada abertura do Excel — "Alinhar à direita"
desligado e "Zero contábil" ligado — independentemente do que foi selecionado na sessão anterior.
Esta é uma decisão deliberada de simplificação de escopo para a migração, não uma regressão ou bug.

### Atualizando a partir da versão VBA

Se você já tem o add-in antigo (`FinanceFmtTools.xlam`) instalado, remova-o **antes** de instalar
esta versão, para evitar que duas abas "Finance Fmt" apareçam simultaneamente na Ribbon:

- No Excel: **Arquivo > Opções > Suplementos > Gerenciar: Suplementos COM > Ir...**, desmarque/remova
  o add-in antigo; ou
- Feche o Excel e apague o arquivo `FinanceFmtTools.xlam` de `%APPDATA%\Microsoft\AddIns`.

### Instalação

Uma linha no PowerShell (recomendado):

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force; irm https://raw.githubusercontent.com/tpougy/finance-fmt-tools/main/scripts/install.ps1 | iex
```

### Notas técnicas

- O código-fonte VBA original (`.bas`, instaladores legados) foi preservado integralmente no branch
  `archive/vba-legacy` e removido do fluxo ativo de build/release do `main`.
- Para desinstalar, rode `scripts/uninstall.ps1` — ele reverte tudo o que `scripts/install.ps1`
  registrou (chaves HKCU e arquivos instalados).
