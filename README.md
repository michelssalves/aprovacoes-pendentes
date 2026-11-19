# aprovacoes-pendentes — Aprovações pendentes (WebChannel)

Painel em **Angular 19 + PO-UI** aberto **dentro do TOTVS Protheus** por `FwCallApp`, que reúne numa
única visão os **pedidos de compra, notas fiscais e solicitações de pagamento** parados na alçada de
aprovação. A comunicação com o AdvPL/TLPP é pelo **WebChannel**, sem API REST.

**Demonstração (dados fictícios):** https://michelssalves.github.io/aprovacoes-pendentes/

## O que o painel entrega

- Pedidos (PC), notas (NF) e solicitações de pagamento (SP) com status, nível atual, aprovadores,
  valor e há quanto tempo estão com o aprovador atual
- **Fluxo por nível** de cada documento (aprovado, rejeitado, bloqueado, pendente, automático...)
- Indicadores clicáveis que filtram a tabela: pendentes (quantidade e valor), idade média, documento
  mais antigo, tempo médio de aprovação e **tempo até classificar a nota em dias úteis** (sem fim de
  semana e feriados nacionais/municipais)
- Gráficos por tipo, nível e aprovador que também filtram a tabela
- Filtros por período, empresa, filial, tipo, classificação da nota e aprovador
- Excel no modelo da planilha que a área já usava, com **ordem das colunas escolhida pelo usuário**
- Atalho para abrir o pedido, a solicitação ou a SP no Salesforce

## Contrato das mensagens

| Ida (`jsToAdvpl`) | Volta (`advplToJs`) |
|---|---|
| `jstoadvplPar1` (uma vez, ao conectar) | `aprovacoesPendentesInicio` → empresas e filiais disponíveis |
| `consultarPendencias` `{dataDe, dataAte}` | `aprovacoesPendentesConsulta` → `{sucesso, mensagem, itens[]}` |
| `abrirSalesforce` `{tipo, id}` | — (o AdvPL abre no navegador padrão) |
| `close` | — (tratado pelo `FwCallApp`) |

Os tipos estão em [`src/main.ts`](src/main.ts), junto da ponte do WebChannel.

## Builds

| Configuração | Uso |
|---|---|
| `protheus` | conecta o WebChannel de verdade; `npm run protheus:build` gera `Protheus/Resource/aprovacoes-pendentes.app` (fora do Git) |
| `demo` | troca `modo.ts` por `modo.demo.ts` e carrega `src/demo/protheus-simulado.ts`, que responde às mesmas mensagens com dados fictícios |

```bash
npm ci
npm start                   # demonstração em http://localhost:4200
npm run protheus:build      # pacote .app para o RPO
npm run deploy              # publica a demonstração no GitHub Pages (branch gh-pages)
```

No Protheus, o launcher chama `FwCallApp("aprovacoes-pendentes")` e trata as mensagens acima numa
`Static Function JsToAdvpl` no mesmo fonte `.prw`.

---

Michel Alves · Desenvolvedor TOTVS Protheus (AdvPL/TLPP) e Angular
