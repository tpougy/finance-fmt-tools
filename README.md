# Finance Format Tools

> Add-in Excel (.xlam) para formatação padronizada de dados financeiros em mercado de capitais.

**Versão:** 1.3.0 · **Plataforma:** Excel para Windows (Office 2016+)

---

## Introdução

O **Finance Format Tools** é um add-in para Excel desenvolvido para padronizar a formatação de dados em planilhas de renda fixa e mercado de capitais. Ele resolve um problema recorrente em análises de debêntures, CRI/CRA, NTN-B e FIIs: a ausência de um padrão consistente de exibição para taxas, preços unitários, spreads e datas — que muitas vezes resulta em células com formatações inconsistentes, números difíceis de comparar visualmente e erros silenciosos de leitura.

O add-in se integra ao Excel por meio de uma aba customizada no ribbon chamada **Finance Fmt**, que agrupa todos os formatos por categoria. A aplicação é feita com um clique: selecione o intervalo e clique no botão correspondente.

### O que o add-in faz

| Categoria | Formatos disponíveis |
|---|---|
| **Numérico / Financeiro** | Fin 0D, Fin 2D, Fin 4D, Fin 8D |
| **Percentual** | % 2D, % 4D |
| **Spread** | Spread em bps |
| **Data** | ISO 8601, BR (dd/mm/yyyy), BR Extenso |
| **Texto** | Forçar formato texto |

Dois checkboxes globais no ribbon modificam o comportamento de todos os formatos da família Fin:

- **Forçar à direita** — preenche o espaço vazio da célula com espaços, alinhando os dígitos à direita da coluna.
- **Zero contábil ("-")** — exibe células com valor exatamente zero como um traço (`-`) no lugar de zeros decimais.

As preferências são persistidas dentro do próprio arquivo `.xlam` via `CustomXMLPart`, portanto sobrevivem ao fechamento e reabertura do Excel.

---

## Família Fin xD

A família **Fin xD** é o núcleo do add-in. São quatro formatos numéricos com estilo contábil — negativos entre parênteses, separador de milhar, alinhamento por vírgula decimal — diferindo apenas no número de casas decimais.

### Quando usar cada formato

| Botão | Casas decimais | Uso típico |
|---|---|---|
| **Fin 0D** | 0 | Quantidades em unidades, notional arredondado, contagem de dias |
| **Fin 2D** | 2 | Preços de mercado, cotações, valores em R$ para relatórios |
| **Fin 4D** | 4 | Spreads em decimal, taxas resumidas, Duration |
| **Fin 8D** | 8 | Taxas internas (CDI%, IPCA+), PU de debêntures, fator de correção |

### Exemplos de exibição

Valor de entrada: `1234567.8901234`

| Formato | Positivo | Negativo | Zero (padrão) | Zero (contábil) |
|---|---|---|---|---|
| Fin 0D | `1.234.568` | `(1.234.568)` | `0` | `-` |
| Fin 2D | `1.234.567,89` | `(1.234.567,89)` | `0,00` | `-` |
| Fin 4D | `1.234.567,8901` | `(1.234.567,8901)` | `0,0000` | `-` |
| Fin 8D | `1.234.567,89012340` | `(1.234.567,89012340)` | `0,00000000` | `-` |

> Os exemplos acima assumem **Alinhar à direita (force)** desativado. Com o checkbox ativo, o Excel insere espaços à esquerda para que os dígitos sempre colem na margem direita da coluna.

### Comportamento de alinhamento (Force Align)

O caractere `*` nas format strings instrui o Excel a repetir o caractere seguinte até preencher a largura da célula. Nos formatos Fin, o caractere repetido é o espaço (` `), o que empurra os dígitos para a direita sem alterar o valor ou a aparência dos números em si.

```
Com Force Align OFF:  _(#,##0.00_)_-
Com Force Align ON:    * _(#,##0.00_)_-
                      ^^^
                      espaços preenchem até a margem
```

Isso é especialmente útil em tabelas onde colunas de diferentes larguras precisam ter os algarismos alinhados entre si visualmente.

### Comportamento do zero como traço (Zero Dash)

Quando ativado, a terceira seção do formato — que controla exclusivamente células com valor zero — é substituída por:

```
_(-_)_-
```

Decomposição:

| Token | Significado |
|---|---|
| `_(` | Reserva um espaço equivalente ao `(` do negativo |
| `-` | Traço literal exibido no lugar do zero |
| `_)` | Reserva um espaço equivalente ao `)` do negativo |
| `_-` | Reserva um espaço equivalente ao `-` final do positivo |

O resultado é um traço perfeitamente alinhado com a coluna de dígitos, sem deslocar o layout da linha.

### Strings de formato geradas

A tabela abaixo mostra as format strings completas em cada combinação de configuração. O separador `;` divide as três seções: `positivo ; negativo ; zero`.

#### Force Align OFF · Zero Dash OFF

| Formato | String |
|---|---|
| Fin 0D | `_(#,##0_)_-;(#,##0)_-;_(#,##0_)_-` |
| Fin 2D | `_(#,##0.00_)_-;(#,##0.00)_-;_(#,##0.00_)_-` |
| Fin 4D | `_(#,##0.0000_)_-;(#,##0.0000)_-;_(#,##0.0000_)_-` |
| Fin 8D | `_(#,##0.00000000_)_-;(#,##0.00000000)_-;_(#,##0.00000000_)_-` |

#### Force Align OFF · Zero Dash ON

| Formato | String |
|---|---|
| Fin 0D | `_(#,##0_)_-;(#,##0)_-;_(-_)_-` |
| Fin 2D | `_(#,##0.00_)_-;(#,##0.00)_-;_(-_)_-` |
| Fin 4D | `_(#,##0.0000_)_-;(#,##0.0000)_-;_(-_)_-` |
| Fin 8D | `_(#,##0.00000000_)_-;(#,##0.00000000)_-;_(-_)_-` |

#### Force Align ON · Zero Dash OFF

| Formato | String |
|---|---|
| Fin 0D | ` * _(#,##0_)_-; * (#,##0)_-; * _(#,##0_)_-` |
| Fin 2D | ` * _(#,##0.00_)_-; * (#,##0.00)_-; * _(#,##0.00_)_-` |
| Fin 4D | ` * _(#,##0.0000_)_-; * (#,##0.0000)_-; * _(#,##0.0000_)_-` |
| Fin 8D | ` * _(#,##0.00000000_)_-; * (#,##0.00000000)_-; * _(#,##0.00000000_)_-` |

#### Force Align ON · Zero Dash ON

| Formato | String |
|---|---|
| Fin 0D | ` * _(#,##0_)_-; * (#,##0)_-; * _(-_)_-` |
| Fin 2D | ` * _(#,##0.00_)_-; * (#,##0.00)_-; * _(-_)_-` |
| Fin 4D | ` * _(#,##0.0000_)_-; * (#,##0.0000)_-; * _(-_)_-` |
| Fin 8D | ` * _(#,##0.00000000_)_-; * (#,##0.00000000)_-; * _(-_)_-` |

---

## Outros formatos

### Percentual

| Botão | String de formato | Exemplo (valor `0.12345`) |
|---|---|---|
| % 2D | `0.00%` | `12,35%` |
| % 4D | `0.0000%` | `12,3450%` |

O Excel multiplica o valor por 100 automaticamente ao detectar `%` na format string — armazene sempre o valor em forma decimal (ex.: `0.1235` para 12,35%).

### Spread em bps

| Botão | String de formato | Exemplo (valor `0.0125`) |
|---|---|---|
| Spread bps | `#,##0.0" bps"` | `125,0 bps` |

O valor esperado é o spread já em pontos-base (ex.: `125.0` exibe `125,0 bps`). O sufixo `" bps"` é texto literal embutido na format string.

### Datas

| Botão | String de formato | Exemplo |
|---|---|---|
| ISO | `yyyy-mm-dd` | `2025-03-15` |
| BR | `[$-pt-BR]dd/mm/yyyy` | `15/03/2025` |
| BR Extenso | `[$-pt-BR]dd/mmm/yyyy` | `15/mar/2025` |

O prefixo `[$-pt-BR]` instrui o Excel a usar o locale pt-BR para abreviações de meses, independentemente do locale configurado no sistema operacional.

### Texto

Aplica a format string `@`, que força o Excel a tratar o conteúdo da célula como texto — útil para CNPJs, códigos CETIP, CEPs e outros identificadores que começam com zero ou contêm caracteres que o Excel interpretaria como número ou data.

---

## Instalação

> 🚧 _Seção em construção._

---

## Referência rápida do ribbon

```
Finance Fmt
├── Numérico
│   ├── Fin 8D          → Financeiro 8 casas decimais
│   ├── Fin 4D          → Financeiro 4 casas decimais
│   ├── Fin 2D          → Financeiro 2 casas decimais
│   ├── Fin 0D          → Financeiro 0 casas decimais (inteiro)
│   ├── ─────────────
│   ├── ☐ Alinhar à direita (force)
│   └── ☐ Zero como "-"
├── Percentual
│   ├── % 4D            → 0.0000%
│   ├── % 2D            → 0.00%
│   └── Spread bps      → #,##0.0 bps
├── Data
│   ├── ISO             → yyyy-mm-dd
│   ├── BR              → dd/mm/yyyy
│   └── BR Extenso      → dd/mmm/yyyy
├── Texto
│   └── Texto           → @
└── Info
    ├── Guia Fin        → abre esta documentação
    └── Sobre           → versão do add-in
```

---

## Persistência de configurações

As preferências dos checkboxes (**Force Align** e **Zero Dash**) são salvas dentro do próprio arquivo `.xlam` usando um `CustomXMLPart` com namespace `urn:finance-fmt-tools`:

```xml
<FmtConfig xmlns="urn:finance-fmt-tools">
  <ForceAlign>true</ForceAlign>
  <ZeroDash>false</ZeroDash>
</FmtConfig>
```

Isso significa que as configurações sobrevivem ao fechamento do Excel e são carregadas automaticamente na próxima sessão, sem depender de registro do Windows, arquivos `.ini` externos ou outras fontes.

---

## Arquitetura do projeto

```
RBR Finance Tools.xlam
│
├── customUI14.xml      Ribbon XML — define a aba "Finance Fmt" e seus controles
│
├── modConfig.bas       Constantes globais, chaves de formato e estado dos checkboxes
├── modFormatEngine.bas Motor central — GetFormatDef(), ApplyFormat(), AccountingFmt()
├── modRibbon.bas       Callbacks do ribbon — wrappers de uma linha para cada botão
└── modUtils.bas        Log, SafeSelection, ShowAbout, persistência em CustomXMLPart
```

**Princípios de design:**

- Todo acesso a `Selection` passa por `SafeSelection()` em `modUtils` — nenhum outro módulo toca `Selection` diretamente.
- Cada callback do ribbon tem exatamente uma linha de lógica; toda a lógica real vive em `modFormatEngine` ou `modUtils`.
- Para adicionar um novo formato: crie uma constante em `modConfig`, adicione um `Case` em `GetFormatDef()` em `modFormatEngine`, e adicione um botão em `customUI14.xml` com o callback em `modRibbon`. Nenhum outro arquivo precisa ser modificado.

---

## Licença

<!-- Adicionar licença aqui -->