import type {
  AprovacoesPendentesConsultaEnvelope,
  AprovacoesPendentesInicioEnvelope,
  NivelFluxo,
  Pendencia,
  StatusDocumento,
  TipoDocumento
} from '../main';

/**
 * Simulador do Protheus para a demonstração (GitHub Pages): responde às mesmas mensagens do
 * WebChannel que o AdvPL responderia, com documentos fictícios gerados aqui.
 */
export function instalar(): void {
  window.jsToAdvpl = (tipo: string, conteudo: string): boolean => {
    setTimeout(() => responder(tipo, conteudo), 400);
    return true;
  };

  responder('jstoadvplPar1', '');
}

function responder(tipo: string, conteudo: string): void {
  if (tipo === 'jstoadvplPar1') {
    window.advplToJs?.('aprovacoesPendentesInicio', JSON.stringify(inicio()));
  } else if (tipo === 'consultarPendencias') {
    const filtro = JSON.parse(conteudo) as { dataDe: string; dataAte: string };
    window.advplToJs?.('aprovacoesPendentesConsulta', JSON.stringify(consultar(filtro.dataDe, filtro.dataAte)));
  }
  // abrirSalesforce e close não têm resposta.
}

function inicio(): AprovacoesPendentesInicioEnvelope {
  return {
    sucesso: true,
    mensagem: '',
    dados: {
      empresas: [
        { codigo: '01', nome: 'Empresa Demonstração S.A.', filiais: [{ codigo: '01', nome: 'Matriz' }, { codigo: '02', nome: 'Filial Sul' }] },
        { codigo: '02', nome: 'Logística Demonstração Ltda', filiais: [{ codigo: '01', nome: 'Matriz' }] }
      ],
      dataMinima: '20260101',
      usuario: '000001'
    }
  };
}

const APROVADORES = ['Ana Souza', 'Carlos Lima', 'Bianca Rocha', 'João Santos', 'Marina Costa', 'Paulo Reis'];
const COMPRADORES = ['Rita Alves', 'Diego Martins', 'Laura Pinto'];
const FORNECEDORES = [
  'Alfa Serviços Portuários Ltda', 'Beta Manutenção Industrial Ltda', 'Gama Tecnologia S.A.',
  'Delta Transportes Ltda', 'Épsilon Engenharia Ltda', 'Zeta Suprimentos Ltda'
];
const CNPJS = ['11222333000181', '11444777000161', '22333444000181', '33444555000172'];
const TIPOS: TipoDocumento[] = ['PC', 'NF', 'SP'];

function consultar(dataDe: string, dataAte: string): AprovacoesPendentesConsultaEnvelope {
  const inicio = paraData(dataDe);
  const fim = paraData(dataAte);
  const dias = Math.max(1, Math.round((fim.getTime() - inicio.getTime()) / 86_400_000) + 1);
  const quantidade = Math.min(300, 20 + dias * 3);

  return {
    sucesso: true,
    mensagem: '',
    dataDe,
    dataAte,
    itens: Array.from({ length: quantidade }, (_, idx) => gerar(idx, inicio, dias)),
    tempos: { total: 0.4 }
  };
}

function gerar(idx: number, inicio: Date, dias: number): Pendencia {
  const tipo = TIPOS[idx % TIPOS.length];
  const criacao = somarDias(inicio, (idx * 7) % dias);
  const sorteio = (idx * 37) % 10;
  const status: StatusDocumento = sorteio < 5 ? 'PENDENTE' : sorteio < 8 ? 'APROVADO' : sorteio < 9 ? 'REJEITADO' : 'BLOQUEADO';
  const qtdNiveis = 1 + (idx % 3);
  const niveis = gerarNiveis(idx, qtdNiveis, criacao, status);
  const pendenteDesde = status === 'PENDENTE' ? iso(somarDias(criacao, idx % 4)) : '';
  const dataStatus = status === 'PENDENTE' ? '' : iso(somarDias(criacao, 1 + (idx % 6)));
  const fornecedor = FORNECEDORES[idx % FORNECEDORES.length];
  const empresa = idx % 4 === 3 ? '02' : '01';
  const numero = String(310000 + idx).padStart(6, '0');

  const item: Pendencia = {
    empresa,
    filial: empresa === '01' && idx % 5 === 4 ? '02' : '01',
    tipo,
    status,
    dataStatus,
    nivel: niveis.find(n => n.situacao === 'PENDENTE')?.nivel ?? niveis[niveis.length - 1].nivel,
    valor: Math.round((800 + ((idx * 1931.7) % 95000)) * 100) / 100,
    aprovadores: niveis.find(n => n.situacao === 'PENDENTE')?.aprovadores ?? '',
    dataCriacao: iso(criacao),
    pendenteDesde,
    requisitante: COMPRADORES[(idx + 1) % COMPRADORES.length],
    andamento: status === 'PENDENTE' ? `Aguardando nível ${niveis.find(n => n.situacao === 'PENDENTE')?.nivel ?? ''}` : '',
    niveis
  };

  if (tipo === 'SP') {
    item.sp = numero;
    item.idSalesSp = 'a2S000000DEMO01AAA';
    item.vencimentoPrevisto = iso(somarDias(criacao, 20));
    return item;
  }

  item.fornecedor = fornecedor;
  item.codFornecedor = String(100100 + (idx % FORNECEDORES.length)).padStart(6, '0');
  item.lojaFornecedor = '01';
  item.cnpjFornecedor = CNPJS[idx % CNPJS.length];
  item.codComprador = String(1 + (idx % COMPRADORES.length)).padStart(3, '0');
  item.comprador = COMPRADORES[idx % COMPRADORES.length];
  item.pedido = numero;
  item.idSalesPedido = 'a1v000000DEMO03AAA';
  item.solicitacao = String(720000 + idx).padStart(6, '0');
  item.idSalesSolicitacao = 'a1B000000DEMO01AAA';

  if (idx % 3 === 0) {
    item.contrato = `CTR-${100 + (idx % 20)}`;
    item.medicao = String(52000 + idx).padStart(6, '0');
  }

  if (tipo === 'NF') {
    item.nota = String(53190 + idx).padStart(9, '0');
    item.serie = '1';
    item.vencimentoPrevisto = iso(somarDias(criacao, 28));
    item.notaClassificada = status === 'APROVADO' && idx % 2 === 0;

    if (item.notaClassificada) {
      item.classificadaEm = `${iso(somarDias(criacao, 2 + (idx % 6) + (idx % 3)))}T${String(9 + (idx % 8)).padStart(2, '0')}:15`;
    }
  }

  return item;
}

function gerarNiveis(idx: number, quantidade: number, criacao: Date, status: StatusDocumento): NivelFluxo[] {
  const niveis: NivelFluxo[] = [];
  const parado = idx % quantidade;

  for (let n = 0; n < quantidade; n++) {
    const nivel = String(n + 1).padStart(2, '0');
    const aprovadores = `${APROVADORES[(idx + n) % APROVADORES.length]}, ${APROVADORES[(idx + n + 2) % APROVADORES.length]}`;
    const data = `${iso(somarDias(criacao, n + 1))}T${String(10 + n).padStart(2, '0')}:30`;

    if (status === 'APROVADO' || n < parado) {
      niveis.push({ nivel, situacao: 'APROVADO', aprovadores, data });
    } else if (n === parado) {
      const situacao = status === 'PENDENTE' ? 'PENDENTE' : status === 'REJEITADO' ? 'REJEITADO' : 'BLOQUEADO';
      niveis.push({ nivel, situacao, aprovadores, data: status === 'PENDENTE' ? '' : data });
    } else {
      niveis.push({ nivel, situacao: 'AGUARDANDO', aprovadores, data: '' });
    }
  }

  return niveis;
}

function paraData(aaaammdd: string): Date {
  return new Date(Number(aaaammdd.slice(0, 4)), Number(aaaammdd.slice(4, 6)) - 1, Number(aaaammdd.slice(6, 8)));
}

function somarDias(data: Date, dias: number): Date {
  const nova = new Date(data.getTime());
  nova.setDate(nova.getDate() + dias);
  return nova;
}

/** AAAA-MM-DD, formato das datas que o AdvPL devolve. */
function iso(data: Date): string {
  return `${data.getFullYear()}-${String(data.getMonth() + 1).padStart(2, '0')}-${String(data.getDate()).padStart(2, '0')}`;
}
