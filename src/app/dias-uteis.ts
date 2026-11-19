import feriadosBrasil from './feriados/feriados-br.json';
import feriadosMunicipais from './feriados/feriados-municipais.json';

/**
 * Dias úteis com o mesmo calendário do relatório RCAP (rcap-impostos): fora sábado, domingo,
 * feriados nacionais (feriados-br.json, por ano) e os municipais, que se repetem todo ano
 * (feriados-municipais.json). Os dois arquivos são cópia dos assets do RCAP.
 */
const nacionais = new Set(
  Object.values(feriadosBrasil as Record<string, { date: string }[]>).flat().map(feriado => feriado.date)
);

const municipais = new Set(
  (feriadosMunicipais.recorrentes ?? []).map(feriado =>
    `${String(feriado.month).padStart(2, '0')}-${String(feriado.day).padStart(2, '0')}`)
);

export function ehDiaUtil(data: Date): boolean {
  const diaSemana = data.getDay();

  if (diaSemana === 0 || diaSemana === 6) {
    return false;
  }

  const mes = String(data.getMonth() + 1).padStart(2, '0');
  const dia = String(data.getDate()).padStart(2, '0');

  return !nacionais.has(`${data.getFullYear()}-${mes}-${dia}`) && !municipais.has(`${mes}-${dia}`);
}

/**
 * Dias úteis depois de "de" até "ate" (AAAAMMDD), contando "ate" e não "de": no mesmo dia dá 0;
 * aprovado na sexta e classificado na segunda dá 1. Negativo quando "ate" vem antes de "de".
 */
export function diasUteisEntre(de: string, ate: string): number {
  const inicio = new Date(+de.substring(0, 4), +de.substring(4, 6) - 1, +de.substring(6, 8));
  const fim = new Date(+ate.substring(0, 4), +ate.substring(4, 6) - 1, +ate.substring(6, 8));

  if (fim < inicio) {
    return -diasUteisEntre(ate, de);
  }

  let dias = 0;
  const cursor = new Date(inicio);

  cursor.setDate(cursor.getDate() + 1);

  while (cursor <= fim) {
    if (ehDiaUtil(cursor)) {
      dias++;
    }

    cursor.setDate(cursor.getDate() + 1);
  }

  return dias;
}
