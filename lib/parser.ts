import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";

import { LinhaProgramacao } from "@/types/programacao";

import {
  enriquecerLinha,
} from "@/lib/programacao";

/*
|--------------------------------------------------------------------------
| PDF.JS
|--------------------------------------------------------------------------
|
| O worker é carregado diretamente pelo navegador.
| Não é necessário criar pdf.worker.min.mjs no projeto.
|
*/

if (typeof window !== "undefined") {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";
}

/*
|--------------------------------------------------------------------------
| NORMALIZAÇÃO
|--------------------------------------------------------------------------
*/

const normalize = (texto: string) =>
  texto
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");

/*
|--------------------------------------------------------------------------
| ALIASES DAS COLUNAS
|--------------------------------------------------------------------------
*/

const aliases = {
  pedido: [
    "PEDIDO",
    "NUMEROPEDIDO",
    "NRPEDIDO",
    "NROPEDIDO",
    "PED",
  ],

  of: [
    "OF",
    "ORDEMFABRICACAO",
    "ORDEMDEFABRICACAO",
    "ORDEM",
    "ORDEMFAB",
  ],

  descricao: [
    "DESCRICAO",
    "DESCRICAODOPRODUTO",
    "DESCRICAOPRODUTO",
    "PRODUTO",
    "DESC",
    "DESCRICAOITEM",
  ],

  quantidade: [
    "QUANTIDADE",
    "QTD",
    "QTDE",
    "QUANT",
    "QTDPROGRAMADA",
    "QUANTIDADEPROGRAMADA",
  ],
};

/*
|--------------------------------------------------------------------------
| BUSCA VALOR NA LINHA
|--------------------------------------------------------------------------
*/

function buscarValor(
  row: Record<string, unknown>,
  campo: keyof typeof aliases
): unknown {
  for (
    const [coluna, valor] of Object.entries(row)
  ) {
    const normalizada =
      normalize(coluna);

    if (
      aliases[campo].includes(
        normalizada
      )
    ) {
      return valor;
    }
  }

  return "";
}

/*
|--------------------------------------------------------------------------
| CONVERTE QUANTIDADE
|--------------------------------------------------------------------------
*/

function converterQuantidade(
  valor: unknown
): number {
  if (
    valor === null ||
    valor === undefined ||
    valor === ""
  ) {
    return 0;
  }

  if (
    typeof valor === "number"
  ) {
    return Number.isFinite(
      valor
    )
      ? valor
      : 0;
  }

  const texto =
    String(valor)
      .trim()
      .replace(/\s/g, "");

  if (!texto) {
    return 0;
  }

  /*
    Trata formatos como:

    10
    10,5
    10.5
    1.000
    1.000,50
  */

  if (
    texto.includes(",") &&
    texto.includes(".")
  ) {
    return (
      Number(
        texto
          .replace(/\./g, "")
          .replace(",", ".")
      ) || 0
    );
  }

  if (
    texto.includes(",")
  ) {
    return (
      Number(
        texto.replace(",", ".")
      ) || 0
    );
  }

  return Number(texto) || 0;
}

/*
|--------------------------------------------------------------------------
| CONVERTE LINHA DA PLANILHA
|--------------------------------------------------------------------------
*/

function converterLinha(
  row: Record<string, unknown>,
  fonte: string
): LinhaProgramacao {
  const descricao =
    String(
      buscarValor(
        row,
        "descricao"
      ) ?? ""
    ).trim();

  const pedido =
    String(
      buscarValor(
        row,
        "pedido"
      ) ?? ""
    ).trim();

  const of =
    String(
      buscarValor(
        row,
        "of"
      ) ?? ""
    ).trim();

  const quantidade =
    converterQuantidade(
      buscarValor(
        row,
        "quantidade"
      )
    );

  return enriquecerLinha({
    fonte,
    pedido,
    of,
    descricao,
    quantidade,
  });
}

/*
|--------------------------------------------------------------------------
| LÊ EXCEL / XLSX / XLS / CSV
|--------------------------------------------------------------------------
*/

export async function lerPlanilha(
  file: File
): Promise<LinhaProgramacao[]> {
  const buffer =
    await file.arrayBuffer();

  const workbook =
    XLSX.read(
      buffer,
      {
        type: "array",
      }
    );

  const linhas:
    LinhaProgramacao[] = [];

  for (
    const nomeAba of
    workbook.SheetNames
  ) {
    const ws =
      workbook.Sheets[
        nomeAba
      ];

    if (!ws) {
      continue;
    }

    const rows =
      XLSX.utils.sheet_to_json<
        Record<string, unknown>
      >(
        ws,
        {
          defval: "",
        }
      );

    for (
      const row of rows
    ) {
      const linha =
        converterLinha(
          row,
          `${file.name} / ${nomeAba}`
        );

      if (
        linha.descricao &&
        linha.quantidade > 0
      ) {
        linhas.push(
          linha
        );
      }
    }
  }

  return linhas;
}

/*
|--------------------------------------------------------------------------
| EXTRAI TEXTO DO PDF
|--------------------------------------------------------------------------
*/

async function extrairTextoPdf(
  file: File
): Promise<string[]> {
  const buffer =
    await file.arrayBuffer();

  const loadingTask =
    pdfjsLib.getDocument({
      data: buffer,
    });

  const pdf =
    await loadingTask.promise;

  const paginas:
    string[] = [];

  for (
    let paginaNumero = 1;
    paginaNumero <=
    pdf.numPages;
    paginaNumero++
  ) {
    const page =
      await pdf.getPage(
        paginaNumero
      );

    const content =
      await page.getTextContent();

    const textos: string[] = [];

    for (
      const item of content.items
    ) {
      if (
        "str" in item &&
        typeof item.str ===
          "string"
      ) {
        const texto =
          item.str.trim();

        if (texto) {
          textos.push(
            texto
          );
        }
      }
    }

    const textoPagina =
      textos.join(" ").trim();

    if (textoPagina) {
      paginas.push(
        textoPagina
      );
    }
  }

  return paginas;
}

/*
|--------------------------------------------------------------------------
| EXTRAI PEDIDO DO TEXTO
|--------------------------------------------------------------------------
*/

function extrairPedido(
  texto: string,
  nomeArquivo: string
): string {
  const encontrado =
    texto.match(
      /(?:PEDIDO|PED)\s*[:#-]?\s*(\d{4,})/i
    );

  if (
    encontrado &&
    encontrado[1]
  ) {
    return encontrado[1];
  }

  /*
    Se o PDF não possuir o texto
    "PEDIDO", tenta extrair do
    começo do nome do arquivo.

    Exemplo:
    26808 (NOR) - ...
  */

  const peloArquivo =
    nomeArquivo.match(
      /^(\d{4,})/
    );

  if (
    peloArquivo &&
    peloArquivo[1]
  ) {
    return peloArquivo[1];
  }

  return "";
}

/*
|--------------------------------------------------------------------------
| EXTRAI OF DO TEXTO
|--------------------------------------------------------------------------
*/

function extrairOF(
  texto: string
): string {
  const encontrado =
    texto.match(
      /\b(?:OF|O\.F\.|ORDEM\s+DE\s+FABRICA[CÇ][AÃ]O)\s*[:#-]?\s*(\d+)\b/i
    );

  if (
    encontrado &&
    encontrado[1]
  ) {
    return encontrado[1];
  }

  return "";
}

/*
|--------------------------------------------------------------------------
| EXTRAI QUANTIDADE DO TEXTO
|--------------------------------------------------------------------------
*/

function extrairQuantidade(
  texto: string
): number {
  const encontrado =
    texto.match(
      /(?:QTD|QTDE|QUANTIDADE|QUANT)\s*[:#-]?\s*([\d.,]+)/i
    );

  if (
    encontrado &&
    encontrado[1]
  ) {
    return converterQuantidade(
      encontrado[1]
    );
  }

  return 0;
}

/*
|--------------------------------------------------------------------------
| LÊ PDF
|--------------------------------------------------------------------------
*/

export async function lerPdf(
  file: File
): Promise<LinhaProgramacao[]> {
  if (
    typeof window ===
    "undefined"
  ) {
    throw new Error(
      "A leitura de PDF precisa ser executada no navegador."
    );
  }

  const paginas =
    await extrairTextoPdf(
      file
    );

  const linhas:
    LinhaProgramacao[] = [];

  for (
    let indice = 0;
    indice < paginas.length;
    indice++
  ) {
    const texto =
      paginas[indice];

    if (!texto) {
      continue;
    }

    const pedido =
      extrairPedido(
        texto,
        file.name
      );

    const of =
      extrairOF(
        texto
      );

    const quantidade =
      extrairQuantidade(
        texto
      );

    /*
      Mantemos o texto completo da
      página como descrição provisória.

      Isso permite testar o PDF mesmo
      quando o layout específico ainda
      não foi mapeado.
    */

    const descricao =
      texto.trim();

    /*
      Se não encontrou quantidade,
      usa 1 para não descartar a
      página inteira.

      Depois podemos ajustar para o
      layout exato do Pedido.
    */

    const quantidadeFinal =
      quantidade > 0
        ? quantidade
        : 1;

    const linha =
      enriquecerLinha({
        fonte:
          `${file.name} / PDF página ${indice + 1}`,

        pedido,

        of,

        descricao,

        quantidade:
          quantidadeFinal,
      });

    linhas.push(
      linha
    );
  }

  return linhas;
}

/*
|--------------------------------------------------------------------------
| MESCLA PEDIDO + USINAGEM
|--------------------------------------------------------------------------
*/

export function mesclarPedidoUsinagem(
  linhas: LinhaProgramacao[]
): LinhaProgramacao[] {
  const mapa =
    new Map<
      string,
      LinhaProgramacao
    >();

  for (
    const linha of linhas
  ) {
    const chave =
      [
        linha.pedido,
        linha.of,
        linha.descricao,
      ]
        .map(
          (valor) =>
            String(
              valor ?? ""
            )
              .trim()
              .toUpperCase()
        )
        .join("|");

    const existente =
      mapa.get(
        chave
      );

    if (!existente) {
      mapa.set(
        chave,
        linha
      );

      continue;
    }

    const quantidadeExistente =
      Number(
        existente.quantidade ||
          0
      );

    const quantidadeNova =
      Number(
        linha.quantidade ||
          0
      );

    mapa.set(
      chave,
      {
        ...existente,

        pedido:
          existente.pedido ||
          linha.pedido,

        of:
          existente.of ||
          linha.of,

        descricao:
          existente.descricao ||
          linha.descricao,

        material:
          existente.material ||
          linha.material,

        medida:
          existente.medida ||
          linha.medida,

        rebaixo:
          existente.rebaixo ||
          linha.rebaixo,

        acabamento:
          existente.acabamento ||
          linha.acabamento,

        cor:
          existente.cor ||
          linha.cor,

        quantidade:
          Math.max(
            quantidadeExistente,
            quantidadeNova
          ),

        fonte:
          `${existente.fonte} + ${linha.fonte}`,
      }
    );
  }

  return Array.from(
    mapa.values()
  );
}

/*
|--------------------------------------------------------------------------
| ENTRADA PRINCIPAL
|--------------------------------------------------------------------------
*/

export async function lerArquivo(
  file: File
): Promise<LinhaProgramacao[]> {
  const extensao =
    file.name
      .split(".")
      .pop()
      ?.toLowerCase();

  if (
    extensao === "pdf"
  ) {
    return lerPdf(
      file
    );
  }

  if (
    extensao === "xlsx" ||
    extensao === "xls" ||
    extensao === "csv"
  ) {
    return lerPlanilha(
      file
    );
  }

  throw new Error(
    `Formato não suportado: ${file.name}`
  );
}
