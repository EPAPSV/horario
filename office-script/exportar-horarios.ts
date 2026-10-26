// Office Script: lê as folhas com nome dd_MM_aaaa (ex.: 26_10_2026)
// e devolve os horários organizados por data.
//
// Estrutura esperada em cada folha:
//   Linha 2: Equipa | Horário de entrada | Espaço de Trabalho
//   Linha 3 em diante: os dados
//
// Alterações de última hora (opcional): pintar a linha no Excel.
//   - Amarelo, laranja, vermelho... -> "alterado"
//   - Cinzento                       -> "desmarcado"
//   - Verdes, branco ou sem cor      -> normal

interface Linha {
  equipa: string;
  entrada: string;
  espaco: string;
  estado: string;
}

interface Saida {
  dias: { [data: string]: Linha[] };
}

// Quantos dias para trás e para a frente se enviam (mantém o ficheiro pequeno)
const DIAS_ATRAS = 1;
const DIAS_FRENTE = 14;

function main(workbook: ExcelScript.Workbook): Saida {
  const saida: Saida = { dias: {} };

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const inicio = new Date(hoje);
  inicio.setDate(inicio.getDate() - DIAS_ATRAS);
  const fim = new Date(hoje);
  fim.setDate(fim.getDate() + DIAS_FRENTE);

  for (const folha of workbook.getWorksheets()) {
    const partes = folha.getName().trim().match(/^(\d{2})_(\d{2})_(\d{4})$/);
    if (!partes) continue; // ignora folhas que não sejam datas

    const dia = Number(partes[1]);
    const mes = Number(partes[2]);
    const ano = Number(partes[3]);
    const data = new Date(ano, mes - 1, dia);
    if (data < inicio || data > fim) continue;

    const intervalo = folha.getUsedRange();
    if (!intervalo) continue;
    const valores = intervalo.getValues();

    // Encontra a linha de cabeçalho ("Equipa") e lê a partir da seguinte
    let cabecalho = -1;
    for (let i = 0; i < valores.length; i++) {
      if (String(valores[i][0]).trim().toLowerCase() === "equipa") {
        cabecalho = i;
        break;
      }
    }
    if (cabecalho < 0) continue;

    const linhas: Linha[] = [];
    for (let i = cabecalho + 1; i < valores.length; i++) {
      const equipa = String(valores[i][0]).replace(/\u00a0/g, " ").trim();
      if (equipa === "") continue;

      const cor = folha.getCell(i, 2).getFormat().getFill().getColor();
      linhas.push({
        equipa: equipa,
        entrada: formatarHora(valores[i][1]),
        espaco: String(valores[i][2]).trim(),
        estado: classificar(cor)
      });
    }

    const chave = ano + "-" + pad(mes) + "-" + pad(dia);
    saida.dias[chave] = linhas;
  }

  return saida;
}

function pad(n: number): string {
  return n < 10 ? "0" + n : String(n);
}

// O Excel guarda horas como fração do dia (0,4375 = 10:30)
function formatarHora(v: string | number | boolean): string {
  if (typeof v === "number") {
    const minutos = Math.round(v * 1440);
    return pad(Math.floor(minutos / 60) % 24) + ":" + pad(minutos % 60);
  }
  return String(v).trim();
}

function classificar(cor: string): string {
  if (!cor || cor.length < 7) return "normal";
  const r = parseInt(cor.substr(1, 2), 16);
  const g = parseInt(cor.substr(3, 2), 16);
  const b = parseInt(cor.substr(5, 2), 16);

  if (r > 240 && g > 240 && b > 240) return "normal";            // branco
  if (g > r + 10 && g > b + 10) return "normal";                 // verdes da tabela
  if (Math.abs(r - g) < 16 && Math.abs(g - b) < 16) return "desmarcado"; // cinzentos
  return "alterado";                                             // restantes cores
}
