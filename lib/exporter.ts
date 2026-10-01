import ExcelJS from "exceljs";

import {
  LinhaProgramacao,
} from "@/types/programacao";

import {
  consolidarQuantidades,
} from "@/lib/programacao";

const safeName = (
  name: string
) =>
  (
    name ||
    "OUTROS"
  )
    .replace(
      /[\\/*?:[\]]/g,
      " "
    )
    .slice(
      0,
      31
    );

function titulo(
  ws: ExcelJS.Worksheet,
  texto: string,
  colunas: number
) {
  ws.mergeCells(
    1,
    1,
    1,
    colunas
  );

  const celula =
    ws.getCell(
      1,
      1
    );

  celula.value =
    texto;

  celula.font = {
    bold: true,
    size: 16,
    color: {
      argb:
        "FFFFFFFF",
    },
  };

  celula.alignment =
    {
      horizontal:
        "center",

      vertical:
        "middle",
    };

  celula.fill = {
    type: "pattern",
    pattern:
      "solid",

    fgColor: {
      argb:
        "FF17365D",
    },
  };

  ws.getRow(
    1
  ).height = 30;
}

function cabecalho(
  row: ExcelJS.Row
) {
  row.font = {
    bold: true,

    color: {
      argb:
        "FFFFFFFF",
    },
  };

  row.fill = {
    type: "pattern",
    pattern:
      "solid",

    fgColor: {
      argb:
        "FF244062",
    },
  };

  row.alignment =
    {
      horizontal:
        "center",

      vertical:
        "middle",

      wrapText: true,
    };

  row.height = 28;
}

function aplicarBordas(
  ws: ExcelJS.Worksheet,
  inicio: number,
  fim: number,
  colunas: number
) {
  for (
    let linha =
      inicio;
    linha <= fim;
    linha++
  ) {
    for (
      let coluna = 1;
      coluna <=
      colunas;
      coluna++
    ) {
      const celula =
        ws.getCell(
          linha,
          coluna
        );

      celula.border =
        {
          top: {
            style:
              "thin",

            color: {
              argb:
                "FFD9E2F3",
            },
          },

          left: {
            style:
              "thin",

            color: {
              argb:
                "FFD9E2F3",
            },
          },

          bottom: {
            style:
              "thin",

            color: {
              argb:
                "FFD9E2F3",
            },
          },

          right: {
            style:
              "thin",

            color: {
              argb:
                "FFD9E2F3",
            },
          },
        };

      celula.alignment =
        {
          vertical:
            "middle",

          wrapText: true,
        };
    }
  }
}

export async function exportarExcel(
  linhas: LinhaProgramacao[],
  dataProg: string,
  turno: string
) {
  const workbook =
    new ExcelJS.Workbook();

  workbook.creator =
    "Sobral Programação Industrial";

  /*
  |--------------------------------------------------------------------------
  | CONTROLE GERENTE
  |--------------------------------------------------------------------------
  */

  const gerente =
    workbook.addWorksheet(
      "CONTROLE GERENTE"
    );

  titulo(
    gerente,
    `CONTROLE GERENCIAL • ${dataProg} • TURNO ${turno}`,
    4
  );

  gerente.addRow(
    []
  );

  gerente.addRow([
    "Indicador",
    "Valor",
  ]);

  cabecalho(
    gerente.getRow(
      3
    )
  );

  const totalPecas =
    linhas.reduce(
      (
        total,
        linha
      ) =>
        total +
        Number(
          linha.quantidade ||
            0
        ),
      0
    );

  const apontadas =
    linhas.filter(
      (linha) =>
        linha.status ===
        "APONTADA"
    ).length;

  gerente.addRow([
    "Linhas / OFs",
    linhas.length,
  ]);

  gerente.addRow([
    "Peças programadas",
    totalPecas,
  ]);

  gerente.addRow([
    "Programadas",
    linhas.filter(
      (linha) =>
        linha.status ===
        "PROGRAMADO"
    ).length,
  ]);

  gerente.addRow([
    "Apontadas",
    apontadas,
  ]);

  gerente.addRow([
    "Pendentes",
    linhas.filter(
      (linha) =>
        linha.status ===
        "PENDENTE"
    ).length,
  ]);

  gerente.addRow([
    "Aguardando linha",
    linhas.filter(
      (linha) =>
        linha.status ===
        "AGUARDANDO LINHA"
    ).length,
  ]);

  gerente.addRow([
    "Divergências",
    linhas.filter(
      (linha) =>
        linha.status ===
        "DIVERGÊNCIA"
    ).length,
  ]);

  gerente.addRow([
    "% apontado",
    linhas.length
      ? apontadas /
        linhas.length
      : 0,
  ]);

  gerente.getCell(
    "B10"
  ).numFmt =
    "0.0%";

  gerente.columns =
    [
      {
        width: 35,
      },

      {
        width: 20,
      },
    ];

  /*
  |--------------------------------------------------------------------------
  | APONTAMENTOS
  |--------------------------------------------------------------------------
  */

  const apontamento =
    workbook.addWorksheet(
      "APONTAMENTOS"
    );

  titulo(
    apontamento,
    `CONTROLE DOS APONTADORES • ${dataProg} • TURNO ${turno}`,
    8
  );

  apontamento.addRow(
    []
  );

  apontamento.addRow(
    [
      "Pedido",
      "OF",
      "Peça",
      "Descrição",
      "Quantidade",
      "Prioridade",
      "Status",
      "Observação",
    ]
  );

  cabecalho(
    apontamento.getRow(
      3
    )
  );

  for (
    const linha
    of linhas
  ) {
    apontamento.addRow(
      [
        linha.pedido,

        linha.of,

        linha.peca,

        linha.descricao,

        linha.quantidade,

        linha.prioridade,

        linha.status,

        linha.observacao,
      ]
    );
  }

  apontamento.columns =
    [
      {
        width: 14,
      },

      {
        width: 15,
      },

      {
        width: 24,
      },

      {
        width: 55,
      },

      {
        width: 12,
      },

      {
        width: 15,
      },

      {
        width: 22,
      },

      {
        width: 30,
      },
    ];

  aplicarBordas(
    apontamento,
    3,
    apontamento.rowCount,
    8
  );

  /*
  |--------------------------------------------------------------------------
  | QUANTIDADES LÍDERES
  |--------------------------------------------------------------------------
  */

  const resumo =
    workbook.addWorksheet(
      "QUANTIDADES LIDERES"
    );

  titulo(
    resumo,
    `QUANTIDADES PARA LÍDERES • ${dataProg} • TURNO ${turno}`,
    8
  );

  resumo.addRow([]);

  resumo.addRow([
    "Grupo",
    "Peça",
    "Material",
    "Medida",
    "Rebaixo",
    "Acabamento",
    "Cor",
    "Qtd Total",
  ]);

  cabecalho(
    resumo.getRow(
      3
    )
  );

  for (
    const item
    of consolidarQuantidades(
      linhas
    )
  ) {
    resumo.addRow(
      [
        item.maquina,

        item.peca,

        item.material,

        item.medida,

        item.rebaixo,

        item.acabamento,

        item.cor,

        item.quantidade,
      ]
    );
  }

  resumo.columns =
    [
      {
        width: 20,
      },

      {
        width: 24,
      },

      {
        width: 24,
      },

      {
        width: 20,
      },

      {
        width: 15,
      },

      {
        width: 20,
      },

      {
        width: 24,
      },

      {
        width: 14,
      },
    ];

  aplicarBordas(
    resumo,
    3,
    resumo.rowCount,
    8
  );

  /*
  |--------------------------------------------------------------------------
  | ABAS DOS LÍDERES
  |--------------------------------------------------------------------------
  */

  const grupos = [
    ...new Set(
      linhas.map(
        (linha) =>
          linha.maquina ||
          "OUTROS"
      )
    ),
  ].sort();

  for (
    const grupo
    of grupos
  ) {
    const ws =
      workbook.addWorksheet(
        safeName(
          grupo
        )
      );

    titulo(
      ws,
      `${grupo} • PROGRAMAÇÃO DO DIA • ${dataProg} • TURNO ${turno}`,
      10
    );

    ws.addRow([]);

    ws.addRow([
      "Pedido",
      "OF",
      "Peça",
      "Descrição",
      "Material",
      "Medida",
      "Rebaixo",
      "Acabamento",
      "Cor",
      "Quantidade",
    ]);

    cabecalho(
      ws.getRow(
        3
      )
    );

    const linhasGrupo =
      linhas.filter(
        (linha) =>
          linha.maquina ===
          grupo
      );

    for (
      const linha
      of linhasGrupo
    ) {
      ws.addRow(
        [
          linha.pedido,

          linha.of,

          linha.peca,

          linha.descricao,

          linha.material,

          linha.medida,

          linha.rebaixo,

          linha.acabamento,

          linha.cor,

          linha.quantidade,
        ]
      );
    }

    ws.columns = [
      {
        width: 14,
      },

      {
        width: 15,
      },

      {
        width: 22,
      },

      {
        width: 50,
      },

      {
        width: 22,
      },

      {
        width: 20,
      },

      {
        width: 15,
      },

      {
        width: 20,
      },

      {
        width: 24,
      },

      {
        width: 13,
      },
    ];

    ws.pageSetup = {
      orientation:
        "landscape",

      fitToPage:
        true,

      fitToWidth:
        1,

      fitToHeight:
        0,

      paperSize:
        9,

      margins: {
        left: 0.2,
        right: 0.2,
        top: 0.4,
        bottom: 0.4,
        header: 0.2,
        footer: 0.2,
      },
    };

    aplicarBordas(
      ws,
      3,
      ws.rowCount,
      10
    );

    for (
      let i = 4;
      i <=
      ws.rowCount;
      i++
    ) {
      ws.getRow(
        i
      ).height = 35;
    }
  }

  /*
  |--------------------------------------------------------------------------
  | DOWNLOAD
  |--------------------------------------------------------------------------
  */

  const buffer =
    await workbook.xlsx.writeBuffer();

  const blob =
    new Blob(
      [buffer],
      {
        type:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }
    );

  const url =
    URL.createObjectURL(
      blob
    );

  const link =
    document.createElement(
      "a"
    );

  link.href =
    url;

  link.download =
    `Programacao_Producao_${dataProg}.xlsx`;

  link.click();

  URL.revokeObjectURL(
    url
  );
}
