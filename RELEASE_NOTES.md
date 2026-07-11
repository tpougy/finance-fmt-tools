# Release Notes

Changelog mantido manualmente. Cada release publicada (automática via `.github/workflows/release.yml`
ou manual via `gh release create`, ver `RELEASE.md`) usa o conteúdo deste arquivo como corpo da
release (`body_path`/`-F RELEASE_NOTES.md`). Depois de uma release ser criada no GitHub, o texto
publicado é permanente — este arquivo é sobrescrito com a entrada da próxima versão antes de cada
nova tag.

---

## Finance Fmt Tools v2.0.0

Esta é a primeira release da migração completa do add-in "Finance Fmt Tools" de VBA (`.xlam`) para
um add-in COM em C#. A experiência do usuário final na Ribbon do Excel é preservada — mesma aba
"Finance Fmt", mesmos grupos, mesmos botões, mesmos atalhos de formatação — mas toda a implementação
por trás foi reescrita em C#/.NET Framework 4.8, com cobertura de testes automatizados e um fluxo de
build/release 100% via terminal (`dotnet` CLI + GitHub Actions), sem depender de Visual Studio
completo.

### O que mudou

- **Format engine portado para C#**: todos os botões de formatação (contábil `Fin 0D/2D/4D/8D`,
  percentual `Pct 0,00%`/`Pct 0,0000%`, `Spread (bps)`, datas `Date ISO`/`Date BR`/`Date BR Longa`,
  `Integer` e `Text`) foram portados 1:1 a partir do VBA original, com cobertura de testes xUnit
  automatizados (`dotnet test`) para as 16 combinações do formato contábil (decimais × alinhamento ×
  zero contábil).
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
