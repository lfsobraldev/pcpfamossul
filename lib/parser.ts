import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";

import { LinhaProgramacao } from "@/types/programacao";

import {
  enriquecerLinha,
} from "@/lib/programacao";

if (typeof window !== "undefined") {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "/pdf.worker.min.mjs";
}

const normalize = (texto: string) =>
  texto
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
const normalize = (texto: string) =>
  texto
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");

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
    "ITEM",
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

function buscarValor(
  row: Record<string, unknown>,
  campo: keyof typeof aliases
) {
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

  return enriquecerLinha({
    fonte,

    pedido:
      String(
        buscarValor(
          row,
          "pedido"
        ) ?? ""
      ).trim(),

    of:
      String(
        buscarValor(
          row,
          "of"
        ) ?? ""
      ).trim(),

    descricao,

    quantidade:
      Number(
        buscarValor(
          row,
          "quantidade"
        ) || 0
      ),
  });
}

/*
|--------------------------------------------------------------------------
| LÊ EXCEL / XLS / CSV
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
    const nomeAba of workbook.SheetNames
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
| CONVERSÃO DE NÚMERO
|--------------------------------------------------------------------------
*/

function converterNumero(
  valor: string
): number {
  const texto =
    String(valor || "")
      .trim()
      .replace(/\./g, "")
      .replace(",", ".");

  const numero =
    Number(texto);

  return Number.isFinite(
    numero
  )
    ? numero
    : 0;
}

/*
|--------------------------------------------------------------------------
| EXTRAI CAMPOS DE UMA LINHA DE TEXTO DO PDF
|--------------------------------------------------------------------------
|
| Como PDFs não possuem necessariamente uma tabela real, fazemos uma
| interpretação tolerante:
|
| - identifica quantidade no final da linha;
| - tenta identificar Pedido e OF quando aparecem;
| - todo o restante é tratado como descrição.
|
*/

function converterLinhaPdf(
  texto: string,
  fonte: string
): LinhaProgramacao | null {
  const linhaOriginal =
    texto
      .replace(/\s+/g, " ")
      .trim();

  if (
    !linhaOriginal
  ) {
    return null;
  }

  /*
    Ignora cabeçalhos comuns
    de relatório.
  */

  const cabecalho =
    normalize(
      linhaOriginal
    );

  if (
    cabecalho === "PEDIDO" ||
    cabecalho === "DESCRICAO" ||
    cabecalho === "QUANTIDADE" ||
    cabecalho === "QTD" ||
    cabecalho.includes(
      "DESCRICAOQUANTIDADE"
    ) ||
    cabecalho.includes(
      "ORDEMFABRICACAO"
    )
  ) {
    return null;
  }

  /*
    Procura quantidade no final da linha.
  */

  const quantidadeMatch =
    linhaOriginal.match(
      /(?:^|\s)(\d+(?:[.,]\d+)?)\s*$/
    );

  if (
    !quantidadeMatch
  ) {
    return null;
  }

  const quantidade =
    converterNumero(
      quantidadeMatch[1]
    );

  if (
    quantidade <= 0
  ) {
    return null;
  }

  let restante =
    linhaOriginal
      .slice(
        0,
        quantidadeMatch.index
      )
      .trim();

  if (
    !restante
  ) {
    return null;
  }

  /*
    Tenta encontrar OF em formatos como:
      OF 12345
      OF: 12345
      12345
    quando houver uma identificação explícita.
  */

  let of = "";

  const ofMatch =
    restante.match(
      /\bOF\s*[:#-]?\s*([A-Z0-9./-]+)/i
    );

  if (
    ofMatch
  ) {
    of =
      ofMatch[1];

    restante =
      restante
        .replace(
          ofMatch[0],
          " "
        )
        .replace(
          /\s+/g,
          " "
        )
        .trim();
  }

  /*
    Tenta encontrar Pedido.
  */

  let pedido = "";

  const pedidoMatch =
    restante.match(
      /\b(?:PEDIDO|PED)\s*[:#-]?\s*([A-Z0-9./-]+)/i
    );

  if (
    pedidoMatch
  ) {
    pedido =
      pedidoMatch[1];

    restante =
      restante
        .replace(
          pedidoMatch[0],
          " "
        )
        .replace(
          /\s+/g,
          " "
        )
        .trim();
  }

  /*
    Caso o PDF tenha apenas números no começo:
      12345 67890 DESCRIÇÃO 4
    usamos os dois primeiros campos como
    possíveis Pedido e OF.
  */

  if (
    !pedido ||
    !of
  ) {
    const partes =
      restante.split(
        /\s+/
      );

    if (
      partes.length >= 3
    ) {
      const primeiro =
        partes[0];

      const segundo =
        partes[1];

      const pareceNumero =
        /^\d[\d./-]*$/.test(
          primeiro
        );

      const segundoNumero =
        /^\d[\d./-]*$/.test(
          segundo
        );

      if (
        pareceNumero &&
        segundoNumero
      ) {
        if (!pedido) {
          pedido =
            primeiro;
        }

        if (!of) {
          of =
            segundo;
        }

        restante =
          partes
            .slice(2)
            .join(" ");
      }
    }
  }

  const descricao =
    restante
      .replace(
        /\s+/g,
        " "
      )
      .trim();

  if (
    !descricao
  ) {
    return null;
  }

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
| LÊ PDF
|--------------------------------------------------------------------------
*/

export async function lerPdf(
  file: File
): Promise<LinhaProgramacao[]> {
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

    /*
      Agrupa os textos por posição vertical.
      Isso permite reconstruir aproximadamente
      as linhas visuais do PDF.
    */

    const itens =
      content.items
        .filter(
          (
            item
          ): item is typeof item & {
            str: string;
            transform: number[];
          } =>
            "str" in item &&
            "transform" in item
        )
        .map(
          (
            item
          ) => ({
            texto:
              item.str,
            x:
              item.transform[4],
            y:
              item.transform[5],
          })
        )
        .filter(
          (item) =>
            item.texto.trim()
        );

    itens.sort(
      (a, b) => {
        const diferencaY =
          b.y - a.y;

        if (
          Math.abs(
            diferencaY
          ) > 3
        ) {
          return diferencaY;
        }

        return a.x - b.x;
      }
    );

    const grupos:
      {
        y: number;
        itens: {
          texto: string;
          x: number;
          y: number;
        }[];
      }[] = [];

    for (
      const item of itens
    ) {
      let grupo =
        grupos.find(
          (g) =>
            Math.abs(
              g.y - item.y
            ) <= 3
        );

      if (
        !grupo
      ) {
        grupo = {
          y: item.y,
          itens: [],
        };

        grupos.push(
          grupo
        );
      }

      grupo.itens.push(
        item
      );
    }

    grupos.sort(
      (a, b) =>
        b.y - a.y
    );

    for (
      const grupo of grupos
    ) {
      grupo.itens.sort(
        (a, b) =>
          a.x - b.x
      );

      const texto =
        grupo.itens
          .map(
            (item) =>
              item.texto
          )
          .join(" ")
          .replace(
            /\s+/g,
            " "
          )
          .trim();

      const linha =
        converterLinhaPdf(
          texto,
          `${file.name} / página ${pagina}`
        );

      if (
        linha
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
| MESCLA PEDIDO + USINAGEM
|--------------------------------------------------------------------------
*/

export function mesclarPedidoUsinagem(
  linhas: LinhaProgramacao[]
) {
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
      .map(
        (x) =>
          String(x)
            .trim()
            .toUpperCase()
      )
      .join("|");

    const existente =
      mapa.get(
        chave
      );

    if (
      !existente
    ) {
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
            existente.quantidade,
            linha.quantidade
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
  const ext =
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
      ext || ""
    )
  ) {
    return lerPlanilha(
      file
    );
  }

  if (
    ext === "pdf"
  ) {
    return lerPdf(
      file
    );
  }

  throw new Error(
    `Formato ainda não suportado: ${file.name}`
  );
}
