import * as XLSX from "xlsx";

import { LinhaProgramacao } from "@/types/programacao";

import {
  enriquecerLinha,
} from "@/lib/programacao";

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
  row: Record<
    string,
    unknown
  >,
  campo: keyof typeof aliases
) {
  for (
    const [coluna, valor]
    of Object.entries(row)
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
  row: Record<
    string,
    unknown
  >,
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
| LÊ EXCEL
|--------------------------------------------------------------------------
*/

export async function lerPlanilha(
  file: File
): Promise<
  LinhaProgramacao[]
> {
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
    const nomeAba
    of workbook.SheetNames
  ) {
    const ws =
      workbook.Sheets[
        nomeAba
      ];

    /*
      Primeiro tenta tabela
      tradicional.
    */

    const rows =
      XLSX.utils.sheet_to_json<
        Record<
          string,
          unknown
        >
      >(
        ws,
        {
          defval: "",
        }
      );

    for (
      const row
      of rows
    ) {
      const linha =
        converterLinha(
          row,
          `${file.name} / ${nomeAba}`
        );

      if (
        linha.descricao &&
        linha.quantidade >
          0
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
| MESCLA DUPLICADOS
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
    const linha
    of linhas
  ) {
    /*
      Pedido + OF + descrição.

      Não utiliza ITEM.
    */

    const chave = [
      linha.pedido,
      linha.of,
      linha.descricao,
    ]
      .map((x) =>
        String(x)
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

    /*
      Se a mesma linha
      aparecer em mais de uma
      origem, mantém apenas uma.

      Evita dobrar quantidade
      entre Pedido e Usinagem.
    */

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
          existente
            .descricao ||
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
          existente
            .acabamento ||
          linha.acabamento,

        cor:
          existente.cor ||
          linha.cor,

        quantidade:
          Math.max(
            existente
              .quantidade,
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
) {
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
    ].includes(ext || "")
  ) {
    return lerPlanilha(
      file
    );
  }

  throw new Error(
    `Formato ainda não suportado: ${file.name}`
  );
}
