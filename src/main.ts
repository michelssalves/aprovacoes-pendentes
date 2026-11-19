import { bootstrapApplication } from '@angular/platform-browser';
import '@totvs/twebchannel-js';
import { AppComponent } from './app/app.component';
import { appConfig } from './app/app.config';
import { MODO_DEMO } from './app/modo';

// Ponte WebChannel no padrao validado nas telas FwCallApp: conecta e dispara jstoadvplPar1 uma
// unica vez dentro do callback de connect() (padrao do TDN, sem retry/setInterval). A primeira
// resposta do AdvPL pode se perder de forma nao deterministica; a tela avisa por tempo limite.

export interface WebChannelStatus {
  connected: boolean;
  message: string;
  receivedType?: string;
  receivedContent?: string;
}

export type TipoDocumento = 'PC' | 'NF' | 'SP';

export type SituacaoNivel = 'APROVADO' | 'REJEITADO' | 'BLOQUEADO' | 'PENDENTE' | 'AUTOMATICO' | 'DISPENSADO' | 'AGUARDANDO';

/** Status do documento como um todo (AprPndStatus). */
export type StatusDocumento = 'PENDENTE' | 'APROVADO' | 'REJEITADO' | 'BLOQUEADO';

/** Um nível do fluxo de aprovação do documento (AprPndFechaFluxo). */
export interface NivelFluxo {
  nivel: string;
  situacao: SituacaoNivel;
  aprovadores: string;
  /** Liberação do nível, AAAA-MM-DDTHH:MM; vazio enquanto não aprovado. */
  data: string;
}

export interface AprovacoesPendentesEmpresa {
  codigo: string;
  nome: string;
  filiais: { codigo: string; nome: string }[];
}

export interface AprovacoesPendentesInicioEnvelope {
  sucesso: boolean;
  mensagem: string;
  dados?: {
    empresas: AprovacoesPendentesEmpresa[];
    dataMinima: string;
    /** Código do usuário Protheus: chave das preferências de colunas. */
    usuario?: string;
  };
}

/** Um documento com aprovação pendente, como o AdvPL devolve (U_AprPndCon). */
export interface Pendencia {
  empresa: string;
  filial: string;
  tipo: TipoDocumento;
  status: StatusDocumento;
  /** Data da aprovação, rejeição ou bloqueio; vazio no pendente. */
  dataStatus?: string;
  nivel: string;
  valor: number;
  aprovadores: string;
  dataCriacao: string;
  pendenteDesde: string;
  fornecedor?: string;
  codFornecedor?: string;
  lojaFornecedor?: string;
  cnpjFornecedor?: string;
  codComprador?: string;
  comprador?: string;
  pedido?: string;
  idSalesPedido?: string;
  solicitacao?: string;
  idSalesSolicitacao?: string;
  contrato?: string;
  medicao?: string;
  nota?: string;
  serie?: string;
  /** Só na NF: F1_STATUS = "A" (todas as linhas com TES). */
  notaClassificada?: boolean;
  /** Só na NF classificada: F1_DTDIGIT + F1_XHORCLA (AAAA-MM-DDTHH:MM). */
  classificadaEm?: string;
  /** Solicitante da SC (pedido e nota; sem SC, quem criou o pedido) ou da SP. */
  requisitante?: string;
  /** SP: ZA0_VENCTO; contrato: CND_XVENC; NF: vencimento da pré-nota (F1_FINJSON) ou o do contrato. */
  vencimentoPrevisto?: string;
  sp?: string;
  idSalesSp?: string;
  andamento?: string;
  niveis?: NivelFluxo[];
}

export interface AprovacoesPendentesConsultaEnvelope {
  sucesso: boolean;
  mensagem: string;
  dataDe?: string;
  dataAte?: string;
  itens?: Pendencia[];
  tempos?: { total: number };
}

declare global {
  interface Window {
    advplToJs?: (type: string, content: string) => void;
    jsToAdvpl?: (type: string, content: string) => boolean;
    aprovacoesPendentesStatus?: WebChannelStatus;
    twebchannel?: {
      advplToJs?: (type: string, content: string) => void;
      connect?: (callback: () => void) => void;
      gotConnection?: boolean;
      jsToAdvpl?: (type: string, content: string) => void;
    };
  }
}

function publicarStatus(status: WebChannelStatus): void {
  window.aprovacoesPendentesStatus = status;
  window.dispatchEvent(new CustomEvent<WebChannelStatus>('aprovacoesPendentesStatus', {
    detail: status
  }));
}

function configurarBridge(): void {
  const w = window;
  const advplAnterior = typeof w.advplToJs === 'function' ? w.advplToJs.bind(w) : undefined;

  w.advplToJs = (type: string, content: string): void => {
    publicarStatus({
      connected: w.twebchannel?.gotConnection === true,
      message: 'Mensagem recebida do AdvPL.',
      receivedType: type,
      receivedContent: content
    });

    if (advplAnterior) {
      advplAnterior(type, content);
    }
  };

  w.jsToAdvpl = (type: string, content: string): boolean => {
    const canal = w.twebchannel;

    if (typeof canal?.jsToAdvpl !== 'function') {
      publicarStatus({
        connected: false,
        message: 'WebChannel indisponível para enviar JavaScript → AdvPL.'
      });
      return false;
    }

    canal.jsToAdvpl(type, content);
    return true;
  };
}

function conectarWebChannel(): void {
  const canal = window.twebchannel;

  if (typeof canal?.connect !== 'function') {
    publicarStatus({
      connected: false,
      message: 'Pacote @totvs/twebchannel-js não expôs twebchannel.connect().'
    });
    return;
  }

  canal.advplToJs = (type: string, content: string): void => {
    window.advplToJs?.(type, content);
  };

  try {
    canal.connect((): void => {
      canal.gotConnection = true;
      publicarStatus({
        connected: true,
        message: 'WebChannel conectado. Carregando os filtros...'
      });

      // Disparo unico dentro do callback de connect() (padrao oficial TDN,
      // sem retry/setInterval) - o que provou funcionar de forma consistente
      // na investigacao original. Nao reintroduzir retry sem reler a nota.
      dispararCargaInicial();
    });
  } catch (error) {
    publicarStatus({
      connected: false,
      message: `Falha ao conectar o WebChannel: ${String(error)}`
    });
  }
}

function dispararCargaInicial(): void {
  const enviado = window.jsToAdvpl?.('jstoadvplPar1', 'content') ?? false;

  publicarStatus({
    connected: window.twebchannel?.gotConnection === true,
    message: enviado
      ? 'jstoadvplPar1 enviado (padrao oficial TDN) - aguardando resposta do AdvPL...'
      : 'Falha ao enviar jstoadvplPar1: WebChannel indisponivel.'
  });
}

configurarBridge();

if (MODO_DEMO) {
  // Chunk separado: o build do Protheus nem carrega o simulador.
  import('./demo/protheus-simulado').then(simulador => simulador.instalar());
} else {
  conectarWebChannel();
}

bootstrapApplication(AppComponent, appConfig)
  .catch((error: unknown) => console.error(error));
