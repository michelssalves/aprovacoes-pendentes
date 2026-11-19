export type ValorCelula = string | number;

export interface ColunaExportacao {
  campo: string;
  titulo: string;
  /** C = texto, N = número com 2 decimais, I = inteiro, D = data (AAAA-MM-DD...). */
  tipo: 'C' | 'N' | 'I' | 'D';
}

export type LinhaExportacao = Record<string, ValorCelula | undefined>;

export interface ExportacaoStatus {
  /** Aba "1-Parametros": pares pergunta/resposta dos filtros usados. */
  parametros: [string, string][];
  colunas: ColunaExportacao[];
  linhas: LinhaExportacao[];
  nomeArquivo: string;
}

/*
 * Visual da planilha que o cliente usa hoje (exportação do TReport anterior): cabeçalho azul com
 * letra branca, linhas zebradas em dois tons de azul, Courier New 8.
 */
const FONTE = 'Courier New';
const AZUL_CABECALHO = 'FF4F81BD';
const ZEBRA = ['FFB8CCE4', 'FFDCE6F1'];

/**
 * Gera o .xlsx no modelo do relatório anterior: aba de parâmetros e aba de status, uma linha por
 * documento. Os números de documento (pedido, nota, SP) vão como texto para manter os zeros à esquerda.
 */
export async function exportarExcel(exportacao: ExportacaoStatus): Promise<void> {
  const ExcelJS = (await import('exceljs')).default;
  const workbook = new ExcelJS.Workbook();

  montarParametros(workbook.addWorksheet('1-Parametros'), exportacao.parametros);
  montarStatus(workbook.addWorksheet('2-Status', { views: [{ state: 'frozen', ySplit: 1 }] }), exportacao);
  workbook.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 1, visibility: 'visible' }];

  const buffer = await workbook.xlsx.writeBuffer();
  baixarArquivo(new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  }), exportacao.nomeArquivo);
}

type Planilha = import('exceljs').Worksheet;

function montarParametros(planilha: Planilha, parametros: [string, string][]): void {
  planilha.columns = [{ width: 32 }, { width: 28 }];
  planilha.mergeCells('A1:B1');
  planilha.getCell('A1').value = 'Parametros';
  planilha.getCell('A1').font = { name: FONTE, size: 8, bold: true };

  estilizarCabecalho(planilha.addRow(['Pergunta', 'Resposta']));

  parametros.forEach(([pergunta, resposta], indice) => {
    estilizarLinha(planilha.addRow([pergunta, resposta]), indice);
  });
}

function montarStatus(planilha: Planilha, exportacao: ExportacaoStatus): void {
  const { colunas, linhas } = exportacao;

  planilha.columns = colunas.map(coluna => ({
    key: coluna.campo,
    width: larguraDaColuna(coluna, linhas),
    style: { numFmt: formatoDaColuna(coluna.tipo) }
  }));

  estilizarCabecalho(planilha.addRow(colunas.map(coluna => coluna.titulo)));
  planilha.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: Math.max(colunas.length, 1) } };

  linhas.forEach((linha, indice) => {
    const linhaExcel = planilha.addRow(colunas.map(coluna => valorParaCelula(linha[coluna.campo], coluna.tipo)));

    estilizarLinha(linhaExcel, indice);
    colunas.forEach((coluna, posicao) => {
      const formato = formatoDaColuna(coluna.tipo);

      if (formato) {
        linhaExcel.getCell(posicao + 1).numFmt = formato;
      }
    });
  });
}

function estilizarCabecalho(linha: import('exceljs').Row): void {
  linha.eachCell(celula => {
    celula.font = { name: FONTE, size: 8, bold: true, color: { argb: 'FFFFFFFF' } };
    celula.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_CABECALHO } };
  });
}

function estilizarLinha(linha: import('exceljs').Row, indice: number): void {
  linha.eachCell({ includeEmpty: true }, celula => {
    celula.font = { name: FONTE, size: 8 };
    celula.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ZEBRA[indice % 2] } };
  });
}

function formatoDaColuna(tipo: ColunaExportacao['tipo']): string | undefined {
  if (tipo === 'N') {
    return '#,##0.00';
  }

  if (tipo === 'I') {
    return '0';
  }

  if (tipo === 'D') {
    return 'dd/mm/yyyy';
  }

  return undefined;
}

/** Largura pelo maior conteúdo (como o "Auto ajuste" do Excel), entre 10 e 60 caracteres. */
function larguraDaColuna(coluna: ColunaExportacao, linhas: LinhaExportacao[]): number {
  if (coluna.tipo === 'D') {
    return Math.max(coluna.titulo.length + 2, 12);
  }

  const maior = linhas.reduce((tamanho, linha) => Math.max(tamanho, String(linha[coluna.campo] ?? '').length), 0);

  return Math.min(Math.max(coluna.titulo.length, maior, 8) + 2, 60);
}

/** Data chega como AAAA-MM-DD (ou AAAA-MM-DDTHH:MM) e vai para o Excel como data real. */
function valorParaCelula(valor: ValorCelula | undefined, tipo: ColunaExportacao['tipo']): ValorCelula | Date | null {
  if (valor === undefined || valor === '') {
    return null;
  }

  if (tipo === 'D' && typeof valor === 'string' && valor.length >= 10) {
    const [ano, mes, dia] = valor.substring(0, 10).split('-').map(Number);
    return new Date(Date.UTC(ano, mes - 1, dia));
  }

  return valor;
}

function baixarArquivo(blob: Blob, nomeArquivo: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');

  link.href = url;
  link.download = nomeArquivo;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}
