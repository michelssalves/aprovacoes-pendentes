import { CommonModule } from '@angular/common';
import { Component, ElementRef, OnDestroy, OnInit, ViewChild } from '@angular/core';
import { FormsModule } from '@angular/forms';
import {
  PoAccordionItemComponent,
  PoAccordionModule,
  PoButtonModule,
  PoChartModule,
  PoChartType,
  PoFieldModule,
  PoLoadingModule,
  PoModalModule,
  PoNotificationService,
  PoSelectOption,
  PoTableColumn,
  PoTableComponent,
  PoTableLiterals,
  PoTableModule,
  PoTagType
} from '@po-ui/ng-components';
import type { PoChartOptions, PoChartSerie, PoDatepickerRange } from '@po-ui/ng-components';
import type {
  AprovacoesPendentesConsultaEnvelope,
  AprovacoesPendentesEmpresa,
  AprovacoesPendentesInicioEnvelope,
  NivelFluxo,
  Pendencia,
  SituacaoNivel,
  StatusDocumento,
  WebChannelStatus
} from '../main';
import { diasUteisEntre } from './dias-uteis';
import { ColunaExportacao, LinhaExportacao, exportarExcel } from './exportar-excel';

/** Tempo sem resposta da abertura da tela até avisar que o WebChannel provavelmente se perdeu. */
const ESPERA_INICIO_MS = 30_000;

/** Tempo sem resposta da consulta até desistir. */
const ESPERA_CONSULTA_MS = 120_000;

/** O filtro abre nos últimos 90 dias contando hoje (hoje menos 89). */
const DIAS_PADRAO = 89;

const TODOS = 'todos';

/** Prefixo das preferências no localStorage; a chave termina no código do usuário Protheus. */
const CHAVE_COLUNAS = 'aprovacoes-pendentes.colunas.v2';
const CHAVE_ORDEM_EXCEL = 'aprovacoes-pendentes.excel.v2';

/** Marcador da coluna do Excel que se abre em "aprovador + status" por etapa do fluxo. */
const CAMPO_FLUXO = 'fluxo';

/**
 * Colunas do Excel na ordem da planilha que o cliente já usa (relatório de status de aprovação, aba 2-Status).
 * É a ordem padrão; cada usuário pode reordenar e tirar colunas (janela "Colunas do Excel").
 */
const COLUNAS_EXCEL: ColunaExportacao[] = [
  { campo: 'pedido', titulo: 'Pedido', tipo: 'C' },
  { campo: 'valor', titulo: 'Total', tipo: 'N' },
  { campo: 'fornecedor', titulo: 'Nome fornecedor', tipo: 'C' },
  { campo: 'codFornecedor', titulo: 'Fornecedor', tipo: 'C' },
  { campo: 'lojaFornecedor', titulo: 'Loja', tipo: 'C' },
  { campo: 'cnpjFornecedor', titulo: 'CNPJ/CPF', tipo: 'C' },
  { campo: 'contrato', titulo: 'Contrato', tipo: 'C' },
  { campo: 'tipoDescricao', titulo: 'Tipo', tipo: 'C' },
  { campo: 'dataCriacao', titulo: 'Criado em', tipo: 'D' },
  { campo: 'vencimentoPrevisto', titulo: 'Vencimento previsto', tipo: 'D' },
  { campo: 'codComprador', titulo: 'Comprador', tipo: 'C' },
  { campo: 'comprador', titulo: 'Nome comprador', tipo: 'C' },
  { campo: 'requisitante', titulo: 'Requisitante', tipo: 'C' },
  { campo: 'aprovadores', titulo: 'Responsável pendente', tipo: 'C' },
  { campo: 'pendenteDesde', titulo: 'Data do envio ao resp.', tipo: 'D' },
  { campo: 'totalNiveis', titulo: 'Total de etapas', tipo: 'I' },
  { campo: 'niveisAprovados', titulo: 'Etapas aprovadas', tipo: 'I' },
  { campo: CAMPO_FLUXO, titulo: 'Aprovador e status por etapa', tipo: 'C' },
  { campo: 'statusDescricao', titulo: 'Status', tipo: 'C' },
  { campo: 'dataStatus', titulo: 'Data do status', tipo: 'D' },
  { campo: 'empresaNome', titulo: 'Empresa', tipo: 'C' },
  { campo: 'filial', titulo: 'Filial', tipo: 'C' },
  { campo: 'solicitacao', titulo: 'Solicitação', tipo: 'C' },
  { campo: 'medicao', titulo: 'Medição', tipo: 'C' },
  { campo: 'nota', titulo: 'Nota', tipo: 'C' },
  { campo: 'serie', titulo: 'Série', tipo: 'C' },
  { campo: 'classificacao', titulo: 'Classificação da nota', tipo: 'C' },
  { campo: 'sp', titulo: 'SP', tipo: 'C' },
  { campo: 'diasPendente', titulo: 'Dias pendente', tipo: 'I' },
  { campo: 'diasTotalPendente', titulo: 'Dias total pendente', tipo: 'I' },
  { campo: 'classificadaEm', titulo: 'Classificada em', tipo: 'D' },
  { campo: 'diasAteClassificar', titulo: 'Dias úteis até classificar', tipo: 'I' },
  { campo: 'nivel', titulo: 'Nível atual', tipo: 'C' },
  { campo: 'andamento', titulo: 'Andamento', tipo: 'C' }
];

/** Status de cada etapa no Excel, no vocabulário do relatório anterior. */
const STATUS_EXCEL: Record<SituacaoNivel, string> = {
  APROVADO: 'Aprovado',
  REJEITADO: 'Rejeitado',
  BLOQUEADO: 'Bloqueado',
  PENDENTE: 'Aguardando aprovação deste usuário',
  AUTOMATICO: 'Automático (FSSC)',
  DISPENSADO: 'Dispensado',
  AGUARDANDO: 'Aguardando aprovação de outros usuários'
};

const DESCRICAO_SITUACAO: Record<SituacaoNivel, string> = {
  APROVADO: 'Aprovado',
  REJEITADO: 'Rejeitado',
  BLOQUEADO: 'Bloqueado',
  PENDENTE: 'Pendente',
  AUTOMATICO: 'Automático (FSSC)',
  DISPENSADO: 'Dispensado',
  AGUARDANDO: 'Aguardando'
};

const DESCRICAO_STATUS: Record<StatusDocumento, string> = {
  PENDENTE: 'Pendente',
  APROVADO: 'Aprovado',
  REJEITADO: 'Rejeitado',
  BLOQUEADO: 'Bloqueado'
};

/**
 * Cores de status nos gráficos (paleta de status da skill de visualização). A ordem de empilhamento
 * Aprovado, Pendente, Rejeitado, Bloqueado passou no validador de daltonismo; o amarelo do pendente
 * é claro por definição e sempre aparece com legenda e visão em tabela.
 */
const COR_STATUS: Record<StatusDocumento, string> = {
  APROVADO: '#0ca30c',
  PENDENTE: '#fab219',
  REJEITADO: '#d03b3b',
  BLOQUEADO: '#ec835a'
};

const ORDEM_EMPILHAMENTO: StatusDocumento[] = ['APROVADO', 'PENDENTE', 'REJEITADO', 'BLOQUEADO'];

/** Cor única dos gráficos de uma série (azul do Protheus, o mesmo da tela). */
const COR_SERIE = '#0c9abe';

/** Faixas de idade das pendências, em dias desde que chegaram no aprovador atual. */
const FAIXAS_IDADE: { rotulo: string; ate: number }[] = [
  { rotulo: 'Até 2 dias', ate: 2 },
  { rotulo: '3 a 7 dias', ate: 7 },
  { rotulo: '8 a 15 dias', ate: 15 },
  { rotulo: '16 a 30 dias', ate: 30 },
  { rotulo: 'Mais de 30 dias', ate: Number.POSITIVE_INFINITY }
];

/** Quantos aprovadores aparecem no gráfico; o resto vira "Outros". */
const MAX_APROVADORES = 10;

/** Ordem dos botões de status acima da tabela. */
const ORDEM_STATUS: StatusDocumento[] = ['PENDENTE', 'APROVADO', 'REJEITADO', 'BLOQUEADO'];

const DESCRICAO_TIPO: Record<Pendencia['tipo'], string> = {
  PC: 'Pedido de compra',
  NF: 'Nota fiscal',
  SP: 'Solicitação de pagamento'
};

/** A tag de status diz de que fluxo é a aprovação: 3-WAY na nota, P.O. no pedido, S.P. na SP. */
const SUFIXO_TIPO: Record<Pendencia['tipo'], string> = {
  PC: 'P.O.',
  NF: '3-WAY',
  SP: 'S.P.'
};

const TAG_STATUS: Record<StatusDocumento, PoTagType> = {
  PENDENTE: PoTagType.Warning,
  APROVADO: PoTagType.Success,
  REJEITADO: PoTagType.Danger,
  BLOQUEADO: PoTagType.Neutral
};

/** Nível do fluxo como o detalhe da linha mostra. */
interface NivelLinha extends NivelFluxo {
  situacaoDescricao: string;
  dataTexto: string;
}

/** Pendência como a tabela mostra: os campos do AdvPL mais os calculados na tela. */
interface LinhaPendencia extends Pendencia {
  empresaNome: string;
  tipoDescricao: string;
  statusDescricao: string;
  classificacao: string;
  diasPendente: number | null;
  /** Da criação até a aprovação final (ou rejeição/bloqueio); no pendente, até hoje. */
  diasTotalPendente: number | null;
  /** NF aprovada e classificada: dias úteis da aprovação final até a classificação (mesmo dia = 0). */
  diasAteClassificar: number | null;
  /** Chave da tag de status: status + tipo, para a tag dizer o fluxo (3-WAY, P.O., S.P.). */
  statusRotulo: string;
  niveis: NivelLinha[];
}

/** Coluna na janela de ordem do Excel. */
interface ColunaOrdemExcel {
  campo: string;
  titulo: string;
  incluir: boolean;
}

/** Indicadores da faixa acima dos gráficos (pendentes no recorte atual). */
interface Indicadores {
  pendentes: number;
  valorPendente: number;
  idadeMedia: number | null;
  maisAntiga: number | null;
  tempoAprovacao: number | null;
  aprovados: number;
  tempoClassificacao: number | null;
  classificadas: number;
}

/** Indicador clicável: cada um abre na tabela os documentos que entraram na conta dele. */
type TipoIndicador = 'pendentes' | 'idadeMedia' | 'maisAntiga' | 'tempoAprovacao' | 'tempoClassificacao';

/** Recorte escolhido com um clique num gráfico ou indicador: aplicado à tabela até ser limpo. */
interface FiltroGrafico {
  descricao: string;
  aceita: (linha: LinhaPendencia) => boolean;
}

/** Botão de status acima da tabela, com a quantidade nos filtros atuais. */
interface ContagemStatus {
  status: StatusDocumento | typeof TODOS;
  rotulo: string;
  quantidade: number;
}

@Component({
  selector: 'app-root',
  imports: [
    CommonModule,
    FormsModule,
    PoTableModule,
    PoFieldModule,
    PoButtonModule,
    PoLoadingModule,
    PoModalModule,
    PoAccordionModule,
    PoChartModule
  ],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent implements OnInit, OnDestroy {
  protected status: WebChannelStatus = {
    connected: false,
    message: 'Aguardando inicialização do WebChannel.'
  };

  protected iniciando = true;
  protected erroInicio: string | null = null;
  protected consultando = false;
  protected exportando = false;

  protected dataMinima = '2026-01-01';
  protected periodo: PoDatepickerRange = { start: '', end: '' };
  protected statusFiltro: StatusDocumento | typeof TODOS = TODOS;
  protected empresa = TODOS;
  protected filial = TODOS;
  protected tipo = TODOS;
  protected classificacaoFiltro = TODOS;
  protected aprovador = '';

  protected opcoesEmpresa: PoSelectOption[] = [];
  protected opcoesFilial: PoSelectOption[] = [{ label: 'Todas as filiais', value: TODOS }];
  protected readonly opcoesTipo: PoSelectOption[] = [
    { label: 'Todos os tipos', value: TODOS },
    { label: DESCRICAO_TIPO.PC, value: 'PC' },
    { label: DESCRICAO_TIPO.NF, value: 'NF' },
    { label: DESCRICAO_TIPO.SP, value: 'SP' }
  ];

  /** Classificação só existe na nota: escolher uma delas deixa na lista apenas as notas. */
  protected readonly opcoesClassificacao: PoSelectOption[] = [
    { label: 'Todas', value: TODOS },
    { label: 'Classificada', value: 'Classificada' },
    { label: 'Não classificada', value: 'Não classificada' }
  ];

  private readonly colunasPadrao: PoTableColumn[] = [
    { property: 'empresaNome', label: 'Empresa', type: 'string' },
    { property: 'filial', label: 'Filial', type: 'string', width: '95px' },
    { property: 'tipoDescricao', label: 'Tipo', type: 'string' },
    {
      property: 'statusRotulo', label: 'Status', type: 'label', width: '150px',
      labels: ORDEM_STATUS.flatMap(status =>
        (Object.keys(SUFIXO_TIPO) as Pendencia['tipo'][]).map(tipo => ({
          value: `${status}_${tipo}`,
          label: `${DESCRICAO_STATUS[status]} ${SUFIXO_TIPO[tipo]}`,
          type: TAG_STATUS[status]
        }))
      )
    },
    {
      property: 'pedido', label: 'Pedido', type: 'link',
      action: (_valor: string, linha: LinhaPendencia) => this.abrirSalesforce('PEDIDO', linha.idSalesPedido),
      disabled: (linha: LinhaPendencia) => !linha.idSalesPedido,
      tooltip: 'Abrir o pedido no Salesforce'
    },
    {
      property: 'solicitacao', label: 'Solicitação', type: 'link',
      action: (_valor: string, linha: LinhaPendencia) => this.abrirSalesforce('SOLICITACAO', linha.idSalesSolicitacao),
      disabled: (linha: LinhaPendencia) => !linha.idSalesSolicitacao,
      tooltip: 'Abrir a solicitação no Salesforce'
    },
    { property: 'contrato', label: 'Contrato', type: 'string' },
    { property: 'medicao', label: 'Medição', type: 'string' },
    { property: 'nota', label: 'Nota', type: 'string' },
    { property: 'serie', label: 'Série', type: 'string', width: '95px' },
    { property: 'classificacao', label: 'Classificação', type: 'string' },
    {
      property: 'sp', label: 'SP', type: 'link',
      action: (_valor: string, linha: LinhaPendencia) => this.abrirSalesforce('SP', linha.idSalesSp),
      disabled: (linha: LinhaPendencia) => !linha.idSalesSp,
      tooltip: 'Abrir a SP no Salesforce'
    },
    { property: 'codFornecedor', label: 'Cód. fornecedor', type: 'string', visible: false },
    { property: 'lojaFornecedor', label: 'Loja', type: 'string', width: '95px', visible: false },
    { property: 'fornecedor', label: 'Fornecedor', type: 'string' },
    { property: 'cnpjFornecedor', label: 'CNPJ/CPF', type: 'string', visible: false },
    { property: 'codComprador', label: 'Cód. comprador', type: 'string', visible: false },
    { property: 'comprador', label: 'Comprador', type: 'string', visible: false },
    { property: 'requisitante', label: 'Requisitante', type: 'string' },
    { property: 'valor', label: 'Valor', type: 'number', format: '1.2-2' },
    { property: 'dataCriacao', label: 'Criado em', type: 'date', format: 'dd/MM/yyyy' },
    { property: 'vencimentoPrevisto', label: 'Vencimento previsto', type: 'date', format: 'dd/MM/yyyy' },
    { property: 'pendenteDesde', label: 'Pendente desde', type: 'date', format: 'dd/MM/yyyy' },
    { property: 'diasPendente', label: 'Dias pendente', type: 'number' },
    { property: 'diasTotalPendente', label: 'Dias total pendente', type: 'number' },
    { property: 'dataStatus', label: 'Data do status', type: 'date', format: 'dd/MM/yyyy' },
    { property: 'classificadaEm', label: 'Classificada em', type: 'date', format: 'dd/MM/yyyy' },
    { property: 'diasAteClassificar', label: 'Dias úteis até classificar', type: 'number' },
    { property: 'nivel', label: 'Nível', type: 'string', width: '95px' },
    { property: 'andamento', label: 'Andamento', type: 'string' },
    { property: 'aprovadores', label: 'Aprovador atual', type: 'string' }
  ];

  /** A seta de abrir o fluxo só aparece na linha que tem níveis na SCR. */
  protected readonly temFluxo = (linha: LinhaPendencia): boolean => linha.niveis.length > 0;

  protected colunasTabela: PoTableColumn[] = [...this.colunasPadrao];
  protected colunasExcel: ColunaOrdemExcel[] = [];
  protected usuario = '';

  /** Até a primeira consulta a tabela orienta o usuário; depois, vazio é falta de documento. */
  protected literaisTabela: PoTableLiterals = {
    noData: 'Escolha o período e clique em Consultar.'
  };

  protected linhasFiltradas: LinhaPendencia[] = [];
  protected contagens: ContagemStatus[] = [];
  protected indicadores: Indicadores = {
    pendentes: 0, valorPendente: 0, idadeMedia: null, maisAntiga: null, tempoAprovacao: null, aprovados: 0,
    tempoClassificacao: null, classificadas: 0
  };

  // Gráficos: o po-chart criado com o acordeão fechado mede largura 0 e não desenha, por isso só
  // existe enquanto a seção está aberta (mesma solução do relatório de notas).
  protected graficosVisiveis = false;
  /** Quando o Protheus devolveu a última consulta (os números valem para esse momento). */
  protected ultimaLeitura: Date | null = null;
  protected filtroGrafico: FiltroGrafico | null = null;
  protected readonly tipoBarra = PoChartType.Bar;
  protected readonly tipoColuna = PoChartType.Column;
  protected categoriasAprovador: string[] = [];
  protected seriesAprovador: PoChartSerie[] = [];
  protected readonly categoriasIdade = FAIXAS_IDADE.map(faixa => faixa.rotulo);
  protected seriesIdade: PoChartSerie[] = [];
  protected readonly categoriasTipo = [DESCRICAO_TIPO.PC, DESCRICAO_TIPO.NF, DESCRICAO_TIPO.SP];
  protected seriesStatusTipo: PoChartSerie[] = [];
  protected readonly opcoesUmaSerie: PoChartOptions = { legend: false };
  protected readonly opcoesEmpilhado: PoChartOptions = { stacked: true, legendVerticalPosition: 'bottom' };

  /** Os filtros abrem expandidos; o acordeão só existe depois que a carga inicial responde. */
  @ViewChild('itemFiltros')
  private set itemFiltros(item: PoAccordionItemComponent | undefined) {
    if (item && !this.filtrosAbertos) {
      this.filtrosAbertos = true;
      setTimeout(() => item.expand());
    }
  }

  private filtrosAbertos = false;

  /**
   * O fluxo aberto na linha acompanha a rolagem lateral da tabela (o sticky não funciona dentro da
   * célula do po-table): a rolagem horizontal vira a variável --rolagem-x do contêiner.
   */
  @ViewChild('areaTabela')
  private set areaTabela(area: ElementRef<HTMLElement> | undefined) {
    this.areaTabelaElemento?.removeEventListener('scroll', this.aoRolarTabela, true);
    this.areaTabelaElemento = area?.nativeElement;
    this.areaTabelaElemento?.addEventListener('scroll', this.aoRolarTabela, true);
  }

  private areaTabelaElemento?: HTMLElement;

  private readonly aoRolarTabela = (evento: Event): void => {
    const alvo = evento.target as HTMLElement;

    if (alvo.scrollWidth > alvo.clientWidth + 1) {
      this.areaTabelaElemento?.style.setProperty('--rolagem-x', `${alvo.scrollLeft}px`);
    }
  };

  private empresas: AprovacoesPendentesEmpresa[] = [];
  private linhas: LinhaPendencia[] = [];
  private tabelaObservada?: PoTableComponent;

  /**
   * O gerenciador de colunas do po-table entrega a lista inteira (ordem, visível, fixa) ao
   * onVisibleColumnsChange da tabela, sem evento para fora: a tela intercepta esse método para
   * salvar o layout do usuário. Vale para mostrar/ocultar, mover, fixar e "Restaurar padrão".
   */
  @ViewChild(PoTableComponent)
  private set tabela(tabela: PoTableComponent | undefined) {
    if (!tabela || tabela === this.tabelaObservada) {
      return;
    }

    const original = tabela.onVisibleColumnsChange.bind(tabela);

    this.tabelaObservada = tabela;
    tabela.onVisibleColumnsChange = (colunas: PoTableColumn[]): void => {
      original(colunas);
      this.colunasTabela = colunas;
      this.gravarLayoutColunas(colunas);
    };
  }
  private esperaInicio?: ReturnType<typeof setTimeout>;
  private esperaConsulta?: ReturnType<typeof setTimeout>;
  private arrasto?: { elemento: HTMLElement; inicioX: number; scrollInicial: number; moveu: boolean };

  private readonly statusListener = (event: Event): void => {
    this.status = (event as CustomEvent<WebChannelStatus>).detail;
    this.processarMensagem();
  };

  constructor(private readonly notificacao: PoNotificationService) {}

  ngOnInit(): void {
    if (window.aprovacoesPendentesStatus) {
      this.status = window.aprovacoesPendentesStatus;
      this.processarMensagem();
    }

    window.addEventListener('aprovacoesPendentesStatus', this.statusListener);

    // A primeira resposta do AdvPL às vezes não chega (ver main.ts). Sem este aviso a tela
    // ficaria em "Carregando..." para sempre.
    this.esperaInicio = setTimeout(() => {
      if (this.iniciando) {
        this.iniciando = false;
        this.erroInicio = 'O Protheus não respondeu à abertura da tela. Feche a tela e abra novamente.';
      }
    }, ESPERA_INICIO_MS);
  }

  ngOnDestroy(): void {
    window.removeEventListener('aprovacoesPendentesStatus', this.statusListener);
    this.areaTabelaElemento?.removeEventListener('scroll', this.aoRolarTabela, true);
    clearTimeout(this.esperaInicio);
    clearTimeout(this.esperaConsulta);
    this.aoSoltar();
  }

  protected consultar(): void {
    const dataDe = this.paraAaaammdd(this.periodo?.start);
    const dataAte = this.paraAaaammdd(this.periodo?.end);

    if (dataDe.length !== 8 || dataAte.length !== 8) {
      this.notificacao.warning('Informe o período: data inicial e data final.');
      return;
    }

    if (dataDe < this.dataMinima.replace(/-/g, '')) {
      this.notificacao.warning('A data inicial não pode ser anterior a 01/01/2026.');
      return;
    }

    if (this.diasEntre(dataDe, dataAte) > 366) {
      this.notificacao.warning('Consulte no máximo 1 ano por vez.');
      return;
    }

    const enviado = window.jsToAdvpl?.('consultarPendencias', JSON.stringify({ dataDe, dataAte })) ?? false;

    if (!enviado) {
      this.notificacao.error('WebChannel indisponível para consultar as pendências.');
      return;
    }

    this.consultando = true;

    clearTimeout(this.esperaConsulta);
    this.esperaConsulta = setTimeout(() => {
      this.consultando = false;
      this.notificacao.error('O Protheus não respondeu à consulta. Tente novamente.');
    }, ESPERA_CONSULTA_MS);
  }

  protected aoMudarEmpresa(empresa: string): void {
    this.empresa = empresa;
    this.filial = TODOS;
    this.montarOpcoesFilial();
    this.aplicarFiltros();
  }

  protected aoMudarFilial(filial: string): void {
    this.filial = filial;
    this.aplicarFiltros();
  }

  protected aoMudarTipo(tipo: string): void {
    this.tipo = tipo;
    this.aplicarFiltros();
  }

  protected aoMudarClassificacao(classificacao: string): void {
    this.classificacaoFiltro = classificacao;
    this.aplicarFiltros();
  }

  /** Clique numa barra de aprovador: mostra na tabela as pendências daquele aprovador. */
  protected aoClicarAprovador(evento: { category?: string }): void {
    const nome = evento?.category ?? '';

    if (!nome) {
      return;
    }

    // "Outros (n)" junta quem ficou fora dos maiores: filtra por exclusão dos nomes do gráfico.
    const visiveis = new Set(this.categoriasAprovador.filter(categoria => !categoria.startsWith('Outros (')));
    const aceita = nome.startsWith('Outros (')
      ? (linha: LinhaPendencia) => linha.status === 'PENDENTE' && !visiveis.has(linha.aprovadores || 'Sem aprovador')
      : (linha: LinhaPendencia) => linha.status === 'PENDENTE' && (linha.aprovadores || 'Sem aprovador') === nome;

    this.aplicarFiltroGrafico({ descricao: `Pendências de ${nome}`, aceita });
  }

  /** Clique numa faixa de idade: mostra as pendências daquela faixa. */
  protected aoClicarIdade(evento: { category?: string }): void {
    const indice = FAIXAS_IDADE.findIndex(faixa => faixa.rotulo === evento?.category);

    if (indice < 0) {
      return;
    }

    const minimo = indice === 0 ? Number.NEGATIVE_INFINITY : FAIXAS_IDADE[indice - 1].ate;
    const maximo = FAIXAS_IDADE[indice].ate;

    this.aplicarFiltroGrafico({
      descricao: `Pendências de ${FAIXAS_IDADE[indice].rotulo.toLowerCase()}`,
      aceita: linha => linha.status === 'PENDENTE' && (linha.diasPendente ?? 0) > minimo && (linha.diasPendente ?? 0) <= maximo
    });
  }

  /** Clique num pedaço da coluna empilhada: tipo do documento + status. */
  protected aoClicarStatusTipo(evento: { label?: string; category?: string }): void {
    const tipo = (Object.keys(DESCRICAO_TIPO) as Pendencia['tipo'][]).find(chave => DESCRICAO_TIPO[chave] === evento?.category);
    const status = ORDEM_STATUS.find(chave => DESCRICAO_STATUS[chave] === evento?.label);

    if (!tipo || !status) {
      return;
    }

    this.aplicarFiltroGrafico({
      descricao: `${DESCRICAO_TIPO[tipo]} · ${DESCRICAO_STATUS[status].toLowerCase()}`,
      aceita: linha => linha.tipo === tipo && linha.status === status
    });
  }

  protected limparFiltroGrafico(): void {
    this.filtroGrafico = null;
    this.aplicarFiltros();
  }

  /** O recorte do gráfico já carrega o status; o botão de status volta para Todos. */
  /** Clique num indicador: a tabela mostra exatamente os documentos usados na conta dele. */
  protected aoClicarIndicador(tipo: TipoIndicador): void {
    const pendente = (linha: LinhaPendencia) => linha.status === 'PENDENTE';

    switch (tipo) {
      case 'pendentes':
        this.aplicarFiltroGrafico({ descricao: 'Documentos pendentes', aceita: pendente });
        break;
      case 'idadeMedia':
        this.aplicarFiltroGrafico({
          descricao: 'Pendentes usados na idade média',
          aceita: linha => pendente(linha) && linha.diasPendente !== null
        });
        break;
      case 'maisAntiga': {
        const maisAntiga = this.indicadores.maisAntiga;

        if (maisAntiga === null) {
          return;
        }

        this.aplicarFiltroGrafico({
          descricao: 'Pendência mais antiga',
          aceita: linha => pendente(linha) && linha.diasPendente === maisAntiga
        });
        break;
      }
      case 'tempoAprovacao':
        this.aplicarFiltroGrafico({
          descricao: 'Aprovados usados no tempo até aprovar',
          aceita: linha => this.entraNoTempoAprovacao(linha)
        });
        break;
      case 'tempoClassificacao':
        this.aplicarFiltroGrafico({
          descricao: 'Notas usadas no tempo até classificar',
          aceita: linha => linha.diasAteClassificar !== null
        });
        break;
    }
  }

  protected get textoUltimaLeitura(): string {
    if (!this.ultimaLeitura) {
      return '';
    }

    const data = this.ultimaLeitura.toLocaleDateString('pt-BR');
    const hora = this.ultimaLeitura.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

    return `Dados lidos do Protheus em ${data} às ${hora}`;
  }

  /** Aprovado com data de criação e de aprovação, sem data invertida: entra na média do tempo. */
  private entraNoTempoAprovacao(linha: LinhaPendencia): boolean {
    return linha.status === 'APROVADO' && !!linha.dataStatus && !!linha.dataCriacao
      && this.diasEntre(this.paraAaaammdd(linha.dataCriacao), this.paraAaaammdd(linha.dataStatus)) >= 0;
  }

  private aplicarFiltroGrafico(filtro: FiltroGrafico): void {
    this.filtroGrafico = filtro;
    this.statusFiltro = TODOS;
    this.aplicarFiltros();

    setTimeout(() => document.querySelector('.barra-tabela')?.scrollIntoView({ behavior: 'smooth', block: 'start' }));
  }

  protected aoMudarStatus(status: StatusDocumento | typeof TODOS): void {
    this.statusFiltro = status;
    this.aplicarFiltros();
  }

  /** Empresa, filial, tipo e aprovador filtram a lista; os botões de status contam sobre esse recorte. */
  protected aplicarFiltros(): void {
    const busca = this.aprovador.trim().toLowerCase();
    const recorte = this.linhas.filter(linha =>
      (this.empresa === TODOS || linha.empresa === this.empresa)
      && (this.filial === TODOS || linha.filial === this.filial)
      && (this.tipo === TODOS || linha.tipo === this.tipo)
      && (this.classificacaoFiltro === TODOS || linha.classificacao === this.classificacaoFiltro)
      && (!busca || (linha.aprovadores ?? '').toLowerCase().includes(busca))
    );

    this.linhasFiltradas = recorte.filter(linha =>
      (this.statusFiltro === TODOS || linha.status === this.statusFiltro)
      && (!this.filtroGrafico || this.filtroGrafico.aceita(linha))
    );

    this.montarIndicadores(recorte);
    this.montarGraficos(recorte);

    this.contagens = [
      { status: TODOS, rotulo: 'Todos', quantidade: recorte.length },
      ...ORDEM_STATUS.map(status => ({
        status,
        rotulo: DESCRICAO_STATUS[status],
        quantidade: recorte.filter(linha => linha.status === status).length
      }))
    ];
  }

  protected async exportar(): Promise<void> {
    const hoje = new Date().toISOString().substring(0, 10).replace(/-/g, '');
    const etapas = this.linhasFiltradas.reduce((maior, linha) => Math.max(maior, linha.niveis.length), 0);

    this.exportando = true;

    try {
      await exportarExcel({
        parametros: this.parametrosExcel(),
        colunas: this.colunasParaExportar(etapas),
        linhas: this.linhasFiltradas.map(linha => this.paraLinhaExcel(linha, etapas)),
        nomeArquivo: `aprovacoes-pendentes_${hoje}.xlsx`
      });
    } catch (erro) {
      console.error('[aprovacoes-pendentes] falha ao gerar o Excel', erro);
      this.notificacao.error('Não foi possível gerar o Excel.');
    } finally {
      this.exportando = false;
    }
  }

  protected prepararOrdemExcel(): void {
    this.colunasExcel = this.ordenarColunasExcel();
  }

  protected moverColunaExcel(indice: number, destino: number): void {
    if (destino < 0 || destino >= this.colunasExcel.length) {
      return;
    }

    const [coluna] = this.colunasExcel.splice(indice, 1);
    this.colunasExcel.splice(destino, 0, coluna);
  }

  protected restaurarOrdemExcel(): void {
    this.colunasExcel = COLUNAS_EXCEL.map(coluna => ({ campo: coluna.campo, titulo: coluna.titulo, incluir: true }));
  }

  protected salvarOrdemExcel(): boolean {
    if (!this.usuario) {
      this.notificacao.warning('Usuário não identificado. Abra o relatório pelo Protheus para salvar as colunas.');
      return false;
    }

    if (!this.colunasExcel.some(coluna => coluna.incluir)) {
      this.notificacao.warning('Marque pelo menos uma coluna para sair no Excel.');
      return false;
    }

    const preferencia = this.colunasExcel.map(coluna => `${coluna.incluir ? '+' : '-'}${coluna.campo}`);

    if (!this.gravarPreferencia(CHAVE_ORDEM_EXCEL, preferencia)) {
      this.notificacao.error('Não foi possível salvar a ordem das colunas.');
      return false;
    }

    this.notificacao.success('Colunas do Excel salvas para o seu usuário.');
    return true;
  }

  /** Fecha a tela pelo FwCallApp (tipo "close", tratado pelo próprio framework). */
  protected sair(): void {
    const enviado = window.jsToAdvpl?.('close', '') ?? false;

    if (!enviado) {
      this.notificacao.error('WebChannel indisponível para fechar a tela.');
    }
  }

  /** O AdvPL abre o link no navegador da estação (ShellExecute), como a RCAP faz. */
  private abrirSalesforce(tipo: 'PEDIDO' | 'SOLICITACAO' | 'SP', id?: string): void {
    if (!id) {
      return;
    }

    const enviado = window.jsToAdvpl?.('abrirSalesforce', JSON.stringify({ tipo, id })) ?? false;

    if (!enviado) {
      this.notificacao.error('WebChannel indisponível para abrir o Salesforce.');
    }
  }

  private processarMensagem(): void {
    if (!this.status.receivedContent) {
      return;
    }

    switch (this.status.receivedType) {
      case 'aprovacoesPendentesInicio':
        this.processarInicio(this.status.receivedContent);
        break;
      case 'aprovacoesPendentesConsulta':
        this.processarConsulta(this.status.receivedContent);
        break;
      default:
        break;
    }
  }

  private processarInicio(conteudo: string): void {
    clearTimeout(this.esperaInicio);

    // A mensagem de status fica guardada na window e é reprocessada na abertura do componente.
    if (!this.iniciando && !this.erroInicio) {
      return;
    }

    this.iniciando = false;

    try {
      const envelope = JSON.parse(conteudo) as AprovacoesPendentesInicioEnvelope;

      if (!envelope.sucesso || !envelope.dados) {
        this.erroInicio = envelope.mensagem || 'O AdvPL não devolveu as empresas.';
        return;
      }

      const dataMinima = envelope.dados.dataMinima;

      this.erroInicio = null;
      this.empresas = envelope.dados.empresas;
      this.usuario = (envelope.dados.usuario ?? '').trim();
      this.aplicarColunasSalvas();
      this.dataMinima = `${dataMinima.substring(0, 4)}-${dataMinima.substring(4, 6)}-${dataMinima.substring(6, 8)}`;
      this.periodo = { start: this.dataPadrao(), end: this.hojeIso() };
      this.opcoesEmpresa = [
        { label: 'Todas as empresas', value: TODOS },
        ...this.empresas.map(empresa => ({ label: empresa.nome, value: empresa.codigo }))
      ];
      this.montarOpcoesFilial();

      // A tela abre só com os filtros: o usuário escolhe o período e clica em Consultar.
    } catch {
      this.erroInicio = 'Resposta do AdvPL não é um JSON válido.';
    }
  }

  private processarConsulta(conteudo: string): void {
    clearTimeout(this.esperaConsulta);

    if (!this.consultando) {
      return;
    }

    this.consultando = false;

    try {
      const envelope = JSON.parse(conteudo) as AprovacoesPendentesConsultaEnvelope;

      if (!envelope.sucesso) {
        this.notificacao.error(envelope.mensagem || 'O AdvPL não devolveu as pendências.');
        return;
      }

      this.linhas = (envelope.itens ?? []).map(item => this.paraLinha(item));
      this.ultimaLeitura = new Date();
      this.literaisTabela = { noData: 'Nenhum documento com estes filtros.' };
      this.aplicarFiltros();
      console.log(`[aprovacoes-pendentes] ${this.linhas.length} documentos em ${envelope.tempos?.total ?? '?'}s`);
    } catch {
      this.notificacao.error('Resposta do AdvPL não é um JSON válido.');
    }
  }

  private paraLinha(item: Pendencia): LinhaPendencia {
    const empresa = this.empresas.find(e => e.codigo === item.empresa);

    return {
      ...item,
      empresaNome: empresa?.nome ?? item.empresa,
      tipoDescricao: DESCRICAO_TIPO[item.tipo] ?? item.tipo,
      statusRotulo: `${item.status}_${item.tipo}`,
      statusDescricao: `${DESCRICAO_STATUS[item.status] ?? item.status} ${SUFIXO_TIPO[item.tipo] ?? ''}`.trim(),
      classificacao: this.descricaoClassificacao(item),
      aprovadores: item.aprovadores ?? '',
      cnpjFornecedor: this.formatarDocumento(item.cnpjFornecedor),
      diasPendente: this.diasDesde(item.pendenteDesde),
      diasTotalPendente: this.diasTotal(item),
      diasAteClassificar: this.diasAteClassificar(item),
      niveis: (item.niveis ?? []).map(nivel => ({
        ...nivel,
        aprovadores: nivel.aprovadores || '—',
        situacaoDescricao: DESCRICAO_SITUACAO[nivel.situacao] ?? nivel.situacao,
        dataTexto: this.formatarDataHora(nivel.data)
      }))
    };
  }

  /** Pendente conta até hoje; concluído conta até a data do status. Sem data de criação, nada. */
  private diasTotal(item: Pendencia): number | null {
    if (!item.dataCriacao || item.dataCriacao.length < 10) {
      return null;
    }

    if (item.status === 'PENDENTE') {
      return this.diasDesde(item.dataCriacao);
    }

    if (!item.dataStatus) {
      return null;
    }

    return Math.max(0, this.diasEntre(this.paraAaaammdd(item.dataCriacao), this.paraAaaammdd(item.dataStatus)));
  }

  /**
   * Dias úteis entre a aprovação final da nota e a classificação (sem sábado, domingo e feriados,
   * calendário do RCAP); no mesmo dia, 0. Nota classificada antes de terminar a aprovação (poucos
   * casos no QA) fica sem medida, para não entrar negativa na média.
   */
  private diasAteClassificar(item: Pendencia): number | null {
    if (item.tipo !== 'NF' || item.status !== 'APROVADO' || !item.dataStatus || !item.classificadaEm) {
      return null;
    }

    const dias = diasUteisEntre(this.paraAaaammdd(item.dataStatus), this.paraAaaammdd(item.classificadaEm));

    return dias < 0 ? null : dias;
  }

  private descricaoClassificacao(item: Pendencia): string {
    if (item.tipo !== 'NF') {
      return '';
    }

    return item.notaClassificada ? 'Classificada' : 'Não classificada';
  }

  /** AAAA-MM-DD ou AAAA-MM-DDTHH:MM para DD/MM/AAAA HH:MM. */
  protected formatarDataHora(data: string): string {
    if (!data || data.length < 10) {
      return '';
    }

    const [ano, mes, dia] = data.substring(0, 10).split('-');
    const hora = data.length >= 16 ? ` ${data.substring(11, 16)}` : '';

    return `${dia}/${mes}/${ano}${hora}`;
  }

  /** A2_CGC vem só com dígitos: 14 = CNPJ, 11 = CPF; o resto (estrangeiro) fica como veio. */
  private formatarDocumento(documento?: string): string {
    const digitos = documento ?? '';

    if (/^\d{14}$/.test(digitos)) {
      return digitos.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
    }

    if (/^\d{11}$/.test(digitos)) {
      return digitos.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4');
    }

    return digitos;
  }

  private montarOpcoesFilial(): void {
    const empresas = this.empresa === TODOS
      ? this.empresas
      : this.empresas.filter(empresa => empresa.codigo === this.empresa);
    const codigos = new Map<string, string>();

    for (const empresa of empresas) {
      for (const filial of empresa.filiais) {
        if (!codigos.has(filial.codigo)) {
          codigos.set(filial.codigo, empresas.length > 1 ? filial.codigo : filial.nome);
        }
      }
    }

    this.opcoesFilial = [
      { label: 'Todas as filiais', value: TODOS },
      ...[...codigos].map(([codigo, nome]) => ({ label: nome, value: codigo }))
    ];
  }

  /** Período resumido no cabeçalho do acordeão de filtros. */
  protected get rotuloPeriodo(): string {
    const inicio = this.formatarDataHora(String(this.periodo?.start ?? ''));
    const fim = this.formatarDataHora(String(this.periodo?.end ?? ''));

    return inicio && fim ? `${inicio} a ${fim}` : '';
  }

  /** Dias com no máximo 1 casa, no padrão brasileiro (1,8); traço quando não há dado. */
  protected formatarDias(dias: number | null): string {
    return dias === null ? '–' : dias.toLocaleString('pt-BR', { maximumFractionDigits: 1 });
  }

  protected formatarMoeda(valor: number): string {
    return valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 });
  }

  private montarIndicadores(recorte: LinhaPendencia[]): void {
    const pendentes = recorte.filter(linha => linha.status === 'PENDENTE');
    const idades = pendentes.map(linha => linha.diasPendente).filter((dias): dias is number => dias !== null);
    const tempos = recorte
      .filter(linha => this.entraNoTempoAprovacao(linha))
      .map(linha => this.diasEntre(this.paraAaaammdd(linha.dataCriacao), this.paraAaaammdd(linha.dataStatus)));
    const diasClassificacao = recorte
      .map(linha => linha.diasAteClassificar)
      .filter((dias): dias is number => dias !== null);

    this.indicadores = {
      pendentes: pendentes.length,
      valorPendente: pendentes.reduce((soma, linha) => soma + (linha.valor || 0), 0),
      idadeMedia: idades.length ? Math.round(idades.reduce((a, b) => a + b, 0) / idades.length) : null,
      maisAntiga: idades.length ? Math.max(...idades) : null,
      tempoAprovacao: tempos.length ? Math.round(tempos.reduce((a, b) => a + b, 0) / tempos.length * 10) / 10 : null,
      aprovados: tempos.length,
      tempoClassificacao: diasClassificacao.length
        ? Math.round(diasClassificacao.reduce((a, b) => a + b, 0) / diasClassificacao.length * 10) / 10
        : null,
      classificadas: diasClassificacao.length
    };
  }

  private montarGraficos(recorte: LinhaPendencia[]): void {
    const pendentes = recorte.filter(linha => linha.status === 'PENDENTE');

    // Pendências por aprovador atual: os maiores primeiro, o resto somado em "Outros".
    const porAprovador = new Map<string, number>();

    for (const linha of pendentes) {
      const nome = linha.aprovadores || 'Sem aprovador';
      porAprovador.set(nome, (porAprovador.get(nome) ?? 0) + 1);
    }

    const ordenados = [...porAprovador.entries()].sort((a, b) => b[1] - a[1]);
    const principais = ordenados.slice(0, MAX_APROVADORES);
    const outros = ordenados.slice(MAX_APROVADORES).reduce((soma, [, qtd]) => soma + qtd, 0);

    if (outros > 0) {
      principais.push([`Outros (${ordenados.length - MAX_APROVADORES})`, outros]);
    }

    // O gráfico de barras desenha a primeira categoria embaixo: invertido, o maior fica no topo e
    // "Outros" no pé.
    principais.reverse();

    this.categoriasAprovador = principais.map(([nome]) => nome);
    this.seriesAprovador = [{ label: 'Documentos pendentes', data: principais.map(([, qtd]) => qtd), color: COR_SERIE }];

    // Idade das pendências por faixa de dias.
    const idade = FAIXAS_IDADE.map(() => 0);

    for (const linha of pendentes) {
      const dias = linha.diasPendente ?? 0;
      const indice = FAIXAS_IDADE.findIndex(faixa => dias <= faixa.ate);
      idade[indice]++;
    }

    this.seriesIdade = [{ label: 'Documentos pendentes', data: idade, color: COR_SERIE }];

    // Status por tipo de documento, empilhado.
    const tipos: Pendencia['tipo'][] = ['PC', 'NF', 'SP'];

    this.seriesStatusTipo = ORDEM_EMPILHAMENTO.map(status => ({
      label: DESCRICAO_STATUS[status],
      color: COR_STATUS[status],
      data: tipos.map(tipo => recorte.filter(linha => linha.tipo === tipo && linha.status === status).length)
    }));
  }

  /** "AAAA-MM-DD" (ou Date) para "AAAAMMDD"; vazio se não houver data. */
  private paraAaaammdd(data?: string | Date): string {
    if (data instanceof Date) {
      return this.paraIso(data).replace(/-/g, '');
    }

    return (data ?? '').replace(/-/g, '').substring(0, 8);
  }

  private diasEntre(de: string, ate: string): number {
    const inicio = Date.UTC(+de.substring(0, 4), +de.substring(4, 6) - 1, +de.substring(6, 8));
    const fim = Date.UTC(+ate.substring(0, 4), +ate.substring(4, 6) - 1, +ate.substring(6, 8));

    return Math.round((fim - inicio) / 86_400_000);
  }

  private paraIso(data: Date): string {
    return [
      data.getFullYear(),
      String(data.getMonth() + 1).padStart(2, '0'),
      String(data.getDate()).padStart(2, '0')
    ].join('-');
  }

  private hojeIso(): string {
    return this.paraIso(new Date());
  }

  /** Abre nos últimos 14 dias (pedido do usuário), sem passar da data mínima do relatório. */
  private dataPadrao(): string {
    const data = new Date();

    data.setDate(data.getDate() - DIAS_PADRAO);

    const iso = this.paraIso(data);

    return iso < this.dataMinima ? this.dataMinima : iso;
  }

  /**
   * Arrastar a tabela com o mouse rola para o lado. Só vira arrasto depois de 5px, para o clique
   * nos links (pedido, solicitação, SP) continuar funcionando.
   */
  protected iniciarArrasto(evento: MouseEvent): void {
    const elemento = evento.button === 0
      ? this.areaRolavel(evento.target as HTMLElement, evento.currentTarget as HTMLElement)
      : null;

    if (!elemento) {
      return;
    }

    this.arrasto = { elemento, inicioX: evento.clientX, scrollInicial: elemento.scrollLeft, moveu: false };
    document.addEventListener('mousemove', this.aoArrastar);
    document.addEventListener('mouseup', this.aoSoltar);
  }

  private readonly aoArrastar = (evento: MouseEvent): void => {
    if (!this.arrasto) {
      return;
    }

    const deslocamento = evento.clientX - this.arrasto.inicioX;

    if (!this.arrasto.moveu && Math.abs(deslocamento) < 5) {
      return;
    }

    if (!this.arrasto.moveu) {
      this.arrasto.moveu = true;
      this.arrasto.elemento.classList.add('tabela-arrastando');
      window.getSelection()?.removeAllRanges();
    }

    evento.preventDefault();
    this.arrasto.elemento.scrollLeft = this.arrasto.scrollInicial - deslocamento;
  };

  private readonly aoSoltar = (): void => {
    document.removeEventListener('mousemove', this.aoArrastar);
    document.removeEventListener('mouseup', this.aoSoltar);

    if (this.arrasto?.moveu) {
      this.arrasto.elemento.classList.remove('tabela-arrastando');

      // O clique que fecha o arrasto não pode abrir o link que estava embaixo do mouse.
      const bloquearClique = (evento: Event): void => evento.stopPropagation();

      window.addEventListener('click', bloquearClique, { capture: true, once: true });
      setTimeout(() => window.removeEventListener('click', bloquearClique, { capture: true }), 0);
    }

    this.arrasto = undefined;
  };

  /** Primeiro elemento, do alvo até a área da tabela, que tem rolagem horizontal de verdade. */
  private areaRolavel(alvo: HTMLElement, limite: HTMLElement): HTMLElement | null {
    let elemento: HTMLElement | null = alvo;

    while (elemento && elemento !== limite.parentElement) {
      const overflowX = getComputedStyle(elemento).overflowX;

      if (elemento.scrollWidth > elemento.clientWidth + 1 && (overflowX === 'auto' || overflowX === 'scroll')) {
        return elemento;
      }

      elemento = elemento.parentElement;
    }

    return null;
  }

  /** Dias inteiros entre a data (AAAA-MM-DD...) e hoje. */
  private diasDesde(data: string): number | null {
    if (!data || data.length < 10) {
      return null;
    }

    const [ano, mes, dia] = data.substring(0, 10).split('-').map(Number);
    const inicio = new Date(ano, mes - 1, dia);
    const hoje = new Date();
    const hojeSemHora = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());

    return Math.round((hojeSemHora.getTime() - inicio.getTime()) / 86_400_000);
  }

  /** Colunas do Excel na ordem do usuário; a de fluxo vira um par aprovador/status por etapa. */
  private colunasParaExportar(etapas: number): ColunaExportacao[] {
    const tipos = new Map(COLUNAS_EXCEL.map(coluna => [coluna.campo, coluna.tipo]));

    return this.ordenarColunasExcel()
      .filter(coluna => coluna.incluir)
      .flatMap((coluna): ColunaExportacao[] => {
        if (coluna.campo !== CAMPO_FLUXO) {
          return [{ campo: coluna.campo, titulo: coluna.titulo, tipo: tipos.get(coluna.campo) ?? 'C' }];
        }

        return Array.from({ length: etapas }, (_, indice): ColunaExportacao[] => [
          { campo: `aprovadorEtapa${indice + 1}`, titulo: `Aprovador etapa ${indice + 1}`, tipo: 'C' },
          { campo: `statusEtapa${indice + 1}`, titulo: `Status etapa ${indice + 1}`, tipo: 'C' }
        ]).flat();
      });
  }

  private paraLinhaExcel(linha: LinhaPendencia, etapas: number): LinhaExportacao {
    const excel: LinhaExportacao = {
      ...(linha as unknown as LinhaExportacao),
      totalNiveis: linha.niveis.length,
      niveisAprovados: linha.niveis.filter(nivel => nivel.situacao === 'APROVADO').length,
    };

    // Etapa que o documento não tem sai com "---", como no relatório anterior.
    for (let indice = 0; indice < etapas; indice++) {
      const nivel = linha.niveis[indice];

      excel[`aprovadorEtapa${indice + 1}`] = nivel ? (nivel.aprovadores || '---') : '---';
      excel[`statusEtapa${indice + 1}`] = nivel ? (STATUS_EXCEL[nivel.situacao] ?? nivel.situacao) : '---';
    }

    return excel;
  }

  /** Aba "1-Parametros": os filtros com que o Excel foi gerado. */
  private parametrosExcel(): [string, string][] {
    const rotulo = (opcoes: PoSelectOption[], valor: string): string =>
      String(opcoes.find(opcao => opcao.value === valor)?.label ?? valor);

    return [
      ['Período', `${this.formatarDataHora(String(this.periodo?.start ?? ''))} a ${this.formatarDataHora(String(this.periodo?.end ?? ''))}`],
      ['Empresa', rotulo(this.opcoesEmpresa, this.empresa)],
      ['Filial', rotulo(this.opcoesFilial, this.filial)],
      ['Tipo', rotulo(this.opcoesTipo, this.tipo)],
      ['Classificação da nota', rotulo(this.opcoesClassificacao, this.classificacaoFiltro)],
      ['Status', this.statusFiltro === TODOS ? 'Todos' : DESCRICAO_STATUS[this.statusFiltro]],
      ['Aprovador', this.aprovador.trim() || 'Todos'],
      ['Gerado em', new Date().toLocaleString('pt-BR')]
    ];
  }

  /**
   * Ordem e colunas salvas pelo usuário ("+campo" sai no Excel, "-campo" não sai). Coluna nova,
   * que não existia quando ele salvou, entra no fim e marcada.
   */
  private ordenarColunasExcel(): ColunaOrdemExcel[] {
    const disponiveis = new Map(COLUNAS_EXCEL.map(coluna => [coluna.campo, coluna]));
    const ordenadas: ColunaOrdemExcel[] = [];

    for (const item of this.lerPreferencia(CHAVE_ORDEM_EXCEL) ?? []) {
      const campo = item.substring(1);
      const coluna = disponiveis.get(campo);

      if (coluna && (item.startsWith('+') || item.startsWith('-'))) {
        ordenadas.push({ campo, titulo: coluna.titulo, incluir: item.startsWith('+') });
        disponiveis.delete(campo);
      }
    }

    return [
      ...ordenadas,
      ...[...disponiveis.values()].map(coluna => ({ campo: coluna.campo, titulo: coluna.titulo, incluir: true }))
    ];
  }

  /** Layout salvo: um item por coluna, na ordem; "+" visível, "-" oculta, "!" depois do sinal = fixa. */
  private gravarLayoutColunas(colunas: PoTableColumn[]): void {
    this.gravarPreferencia(CHAVE_COLUNAS, colunas.map(coluna =>
      `${coluna.visible === false ? '-' : '+'}${coluna.fixed ? '!' : ''}${String(coluna.property)}`));
  }

  /** Aplica o layout de colunas que o usuário deixou da última vez (ou o padrão). */
  private aplicarColunasSalvas(): void {
    const disponiveis = new Map(this.colunasPadrao.map(coluna => [String(coluna.property), coluna]));
    const colunas: PoTableColumn[] = [];

    for (const item of this.lerPreferencia(CHAVE_COLUNAS) ?? []) {
      const fixa = item.charAt(1) === '!';
      // Layout salvo antes da tag por tipo guardava a coluna como "status".
      const salvo = item.substring(fixa ? 2 : 1);
      const campo = salvo === 'status' ? 'statusRotulo' : salvo;
      const coluna = disponiveis.get(campo);

      if (coluna && (item.startsWith('+') || item.startsWith('-'))) {
        colunas.push({ ...coluna, visible: coluna.type === 'detail' || item.startsWith('+'), fixed: fixa });
        disponiveis.delete(campo);
      }
    }

    // Coluna nova, que não existia quando o usuário salvou, entra no fim com o padrão dela.
    this.colunasTabela = [...colunas, ...disponiveis.values()];
  }

  /**
   * Preferência do usuário no localStorage do navegador embutido do Protheus (mesmo padrão do RCAP,
   * rcap-impostos): a chave leva o código do usuário. Sem usuário, nada é lido nem gravado.
   */
  private lerPreferencia(prefixo: string): string[] | null {
    if (!this.usuario) {
      return null;
    }

    try {
      const salvo: unknown = JSON.parse(window.localStorage.getItem(`${prefixo}:${this.usuario}`) ?? 'null');

      return Array.isArray(salvo) ? salvo.map(String) : null;
    } catch {
      return null;
    }
  }

  private gravarPreferencia(prefixo: string, valor: string[]): boolean {
    if (!this.usuario) {
      return false;
    }

    try {
      window.localStorage.setItem(`${prefixo}:${this.usuario}`, JSON.stringify(valor));
      return true;
    } catch {
      return false;
    }
  }
}
