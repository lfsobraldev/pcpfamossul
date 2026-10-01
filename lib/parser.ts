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
) {
  for (const [coluna, valor] of Object.entries(row)) {
    const normalizada = normalize(coluna);

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
| CONVERTE LINHA DA PLANILHA
|--------------------------------------------------------------------------
*/

function converterLinha(
  row: Record<string, unknown>,
  fonte: string
): LinhaProgramacao {
  const descricao = String(
    buscarValor(
      row,
      "descricao"
    ) ?? ""
  ).trim();

  const quantidadeBruta =
    buscarValor(
      row,
      "quantidade"
    );

  const quantidade = Number(
    String(
      quantidadeBruta ?? "0"
    )
      .replace(/\./g, "")
      .replace(",", ".")
  );

  return enriquecerLinha({
    fonte,

    pedido: String(
      buscarValor(
        row,
        "pedido"
      ) ?? ""
    ).trim(),

    of: String(
      buscarValor(
        row,
        "of"
      ) ?? ""
    ).trim(),

    descricao,

    quantidade:
      Number.isFinite(
        quantidade
      )
        ? quantidade
        : 0,
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

  const buffer =
    await file.arrayBuffer();

  const pdf =
    await pdfjsLib.getDocument({
      data: buffer,
    }).promise;

  const linhas:
    LinhaProgramacao[] = [];

  for (
    let pagina = 1;
    pagina <= pdf.numPages;
    pagina++
  ) {
    const page =
      await pdf.getPage(
        pagina
      );

    const content =
      await page.getTextContent();

    const textos =
      content.items
        .map(
          (item: any) =>
            typeof item.str ===
            "string"
              ? item.str
              : ""
        )
        .filter(Boolean);

    const textoPagina =
      textos.join(" ");

    /*
    |--------------------------------------------------------------------------
    | Tenta encontrar número do pedido
    |--------------------------------------------------------------------------
    */

    let pedido = "";

    const pedidoMatch =
      textoPagina.match(
        /(?:PEDIDO|PED)\s*[:#-]?\s*(\d{4,})/i
      );

    if (pedidoMatch) {
      pedido =
        pedidoMatch[1];
    }

    /*
    |--------------------------------------------------------------------------
    | Tenta identificar OFs
    |--------------------------------------------------------------------------
    */

    const ofMatches =
      textoPagina.match(
        /\b(?:OF|O\.F\.)\s*[:#-]?\s*\d+\b/gi
      ) || [];

    /*
    |--------------------------------------------------------------------------
    | Tenta identificar quantidades
    |--------------------------------------------------------------------------
    */

    const quantidadeMatches =
      textoPagina.match(
        /\b(?:QTD|QTDE|QUANTIDADE)\s*[:#-]?\s*\d+(?:[.,]\d+)?\b/gi
      ) || [];

    /*
    |--------------------------------------------------------------------------
    | Caso o PDF tenha texto, cria uma
    | linha com o conteúdo da página.
    |
    | Isso evita perder o PDF mesmo quando
    | o layout ainda não foi mapeado.
    |--------------------------------------------------------------------------
    */

    if (
      textos.length > 0
    ) {
      const descricao =
        textos
          .join(" ")
          .trim();

      let quantidade = 0;

      if (
        quantidadeMatches.length
      ) {
        const ultimo =
          quantidadeMatches[
            quantidadeMatches.length -
              1
          ];

        const numero =
          ultimo.match(
            /(\d+(?:[.,]\d+)?)\s*$/ 
          );

        if (numero) {
          quantidade =
            Number(
              numero[1].replace(
                ",",
                "."
              )
            );
        }
      }

      let of = "";

      if (
        ofMatches.length
      ) {
        of =
          ofMatches[0]
            .replace(
              /[^0-9]/g,
              ""
            );
      }

      linhas.push(
        enriquecerLinha({
          fonte:
            `${file.name} / PDF página ${pagina}`,

          pedido,

          of,

          descricao,

          quantidade:
            quantidade > 0
              ? quantidade
              : 1,
        })
      );
    }
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
    const chave = [
      linha.pedido,
      linha.of,
      linha.descricao,
    ]
      .map((valor) =>
        String(
          valor ?? ""
        )
          .trim()
          .toUpperCase()
      )
      .join("|");

    const existente =
      mapa.get(chave);

    if (!existente) {
      mapa.set(
        chave,
        linha
      );

      continue;
    }

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
            Number(
              existente.quantidade ||
                0
            ),
            Number(
              linha.quantidade ||
                0
            )
          ),

        fonte:
          `${existente.fonte} + ${linha.fonte}`,
      }
    );
  }

  return [
    ...mapa.values(),
  ];
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
    [
      "xlsx",
      "xls",
      "csv",
    ].includes(
      extensao || ""
    )
  ) {
    return lerPlanilha(
      file
    );
  }

  if (
    extensao === "pdf"
  ) {
    return lerPdf(
      file
    );
  }

  throw new Error(
    `Formato não suportado: ${file.name}`
  );
}
