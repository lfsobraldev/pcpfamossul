"use client";

import {
  useMemo,
  useState,
} from "react";

import {
  BarChart3,
  ClipboardCheck,
  Download,
  Factory,
  FileSpreadsheet,
  Trash2,
  WandSparkles,
} from "lucide-react";

import {
  LinhaProgramacao,
  PrioridadeProgramacao,
  StatusProgramacao,
  TipoPeca,
} from "@/types/programacao";

import {
  lerArquivo,
  mesclarPedidoUsinagem,
} from "@/lib/parser";

import {
  consolidarQuantidades,
  enriquecerLinha,
} from "@/lib/programacao";

import {
  exportarExcel,
} from "@/lib/exporter";

type Tab =
  | "programacao"
  | "quantidades"
  | "apontamentos";

export default function Home() {
  const [
    pedido,
    setPedido,
  ] =
    useState<File | null>(
      null
    );

  const [
    usinagem,
    setUsinagem,
  ] =
    useState<File | null>(
      null
    );

  const [
    dataProg,
    setDataProg,
  ] = useState(
    new Date()
      .toISOString()
      .slice(0, 10)
  );

  const [
    turno,
    setTurno,
  ] = useState("A");

  const [
    linhas,
    setLinhas,
  ] = useState<
    LinhaProgramacao[]
  >([]);

  const [
    tab,
    setTab,
  ] =
    useState<Tab>(
      "programacao"
    );

  const [
    busy,
    setBusy,
  ] = useState(false);

  const [
    erro,
    setErro,
  ] = useState("");

  const [
    busca,
    setBusca,
  ] = useState("");

  const resumo =
    useMemo(
      () =>
        consolidarQuantidades(
          linhas
        ),
      [linhas]
    );

  const apontadas =
    linhas.filter(
      (linha) =>
        linha.status ===
        "APONTADA"
    ).length;

  const pendentes =
    linhas.filter(
      (linha) =>
        linha.status ===
        "PENDENTE"
    ).length;

  const pecas =
    linhas.reduce(
      (total, linha) =>
        total +
        Number(
          linha.quantidade ||
            0
        ),
      0
    );

  async function gerar() {
    setErro("");

    if (
      !pedido &&
      !usinagem
    ) {
      setErro(
        "Selecione pelo menos o Pedido."
      );

      return;
    }

    setBusy(true);

    try {
      const todas:
        LinhaProgramacao[] =
        [];

      if (pedido) {
        todas.push(
          ...(await lerArquivo(
            pedido
          ))
        );
      }

      if (usinagem) {
        todas.push(
          ...(await lerArquivo(
            usinagem
          ))
        );
      }

      const resultado =
        mesclarPedidoUsinagem(
          todas
        ).map(
          enriquecerLinha
        );

      setLinhas(
        resultado
      );

      setTab(
        "programacao"
      );

      if (
        !resultado.length
      ) {
        setErro(
          "Arquivo lido, porém não encontrei as linhas de descrição e quantidade. O layout do arquivo precisa ser mapeado."
        );
      }
    } catch (
      error: any
    ) {
      setErro(
        error?.message ||
          "Falha ao processar os arquivos."
      );
    } finally {
      setBusy(false);
    }
  }

  function alterar(
    id: string,
    campo:
      keyof LinhaProgramacao,
    valor: any
  ) {
    setLinhas(
      (anteriores) =>
        anteriores.map(
          (linha) => {
            if (
              linha.id !==
              id
            ) {
              return linha;
            }

            return enriquecerLinha(
              {
                ...linha,

                [campo]:
                  campo ===
                  "quantidade"
                    ? Number(
                        valor
                      )
                    : valor,
              }
            );
          }
        )
    );
  }

  function apagar(
    id: string
  ) {
    setLinhas(
      (anteriores) =>
        anteriores.filter(
          (linha) =>
            linha.id !==
            id
        )
    );
  }

  const filtradas =
    linhas.filter(
      (linha) => {
        const texto =
          busca
            .trim()
            .toUpperCase();

        if (!texto) {
          return true;
        }

        return [
          linha.pedido,
          linha.of,
          linha.peca,
          linha.descricao,
          linha.material,
          linha.medida,
          linha.rebaixo,
          linha.acabamento,
          linha.cor,
          linha.maquina,
          linha.status,
        ]
          .join(" ")
          .toUpperCase()
          .includes(texto);
      }
    );

  return (
    <main className="shell">
      <div className="topbar">
        <div className="brand">
          SOBRAL{" "}
          <span
            style={{
              opacity: 0.72,
            }}
          >
            PROGRAMAÇÃO
            INDUSTRIAL
          </span>

          <small>
            PLANEJAMENTO DE
            PRODUÇÃO
          </small>
        </div>

        <span className="badge">
          Operação diária
        </span>
      </div>

      <div className="wrap">
        <div className="hero">
          <div>
            <h1>
              Gerador de
              Programação de
              Produção
            </h1>

            <p>
              Pedido →
              leitura →
              classificação →
              líderes →
              apontamentos →
              Excel
            </p>
          </div>
        </div>

        <section className="grid">
          <div className="card kpi">
            <Factory
              size={20}
            />

            <span className="muted">
              OFs / linhas
            </span>

            <strong>
              {
                linhas.length
              }
            </strong>
          </div>

          <div className="card kpi">
            <FileSpreadsheet
              size={20}
            />

            <span className="muted">
              Peças
              programadas
            </span>

            <strong>
              {pecas.toLocaleString(
                "pt-BR"
              )}
            </strong>
          </div>

          <div className="card kpi">
            <ClipboardCheck
              size={20}
            />

            <span className="muted">
              Apontadas
            </span>

            <strong>
              {apontadas}
            </strong>
          </div>

          <div className="card kpi">
            <BarChart3
              size={20}
            />

            <span className="muted">
              Pendentes
            </span>

            <strong>
              {pendentes}
            </strong>
          </div>
        </section>

        <section className="upload">
          <div className="drop field">
            <label>
              PEDIDO
            </label>

            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={(
                e
              ) =>
                setPedido(
                  e.target
                    .files?.[0] ||
                    null
                )
              }
            />

            <div
              className="muted"
              style={{
                marginTop: 8,
              }}
            >
              {pedido?.name ||
                "Selecione o Pedido"}
            </div>
          </div>

          <div className="drop field">
            <label>
              USINAGEM
            </label>

            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={(
                e
              ) =>
                setUsinagem(
                  e.target
                    .files?.[0] ||
                    null
                )
              }
            />

            <div
              className="muted"
              style={{
                marginTop: 8,
              }}
            >
              {usinagem?.name ||
                "Opcional nesta etapa"}
            </div>
          </div>

          <div className="field card">
            <label>
              DATA
            </label>

            <input
              type="date"
              value={
                dataProg
              }
              onChange={(
                e
              ) =>
                setDataProg(
                  e.target
                    .value
                )
              }
            />
          </div>

          <div className="field card">
            <label>
              TURNO
            </label>

            <select
              value={turno}
              onChange={(
                e
              ) =>
                setTurno(
                  e.target
                    .value
                )
              }
            >
              <option value="A">
                A
              </option>

              <option value="B">
                B
              </option>

              <option value="A + B">
                A + B
              </option>
            </select>
          </div>
        </section>

        {erro && (
          <div className="notice">
            {erro}
          </div>
        )}

        <div className="actions">
          <button
            className="btn primary"
            onClick={gerar}
            disabled={busy}
          >
            <WandSparkles
              size={16}
              style={{
                verticalAlign:
                  "middle",
                marginRight: 7,
              }}
            />

            {busy
              ? "PROCESSANDO..."
              : "GERAR PROGRAMAÇÃO"}
          </button>

          <button
            className="btn secondary"
            disabled={
              !linhas.length
            }
            onClick={() =>
              exportarExcel(
                linhas,
                dataProg,
                turno
              )
            }
          >
            <Download
              size={16}
              style={{
                verticalAlign:
                  "middle",
                marginRight: 7,
              }}
            />

            EXPORTAR EXCEL
          </button>

          <button
            className="btn danger"
            disabled={
              !linhas.length
            }
            onClick={() =>
              setLinhas([])
            }
          >
            LIMPAR
          </button>
        </div>

        <div className="tabs">
          <button
            className={`tab ${
              tab ===
              "programacao"
                ? "active"
                : ""
            }`}
            onClick={() =>
              setTab(
                "programacao"
              )
            }
          >
            Programação
          </button>

          <button
            className={`tab ${
              tab ===
              "quantidades"
                ? "active"
                : ""
            }`}
            onClick={() =>
              setTab(
                "quantidades"
              )
            }
          >
            Quantidades
            líderes
          </button>

          <button
            className={`tab ${
              tab ===
              "apontamentos"
                ? "active"
                : ""
            }`}
            onClick={() =>
              setTab(
                "apontamentos"
              )
            }
          >
            Apontamentos
          </button>
        </div>

        {tab ===
          "programacao" && (
          <>
            <div className="toolbar">
              <input
                className="search"
                placeholder="Buscar pedido, OF, medida, peça..."
                value={busca}
                onChange={(
                  e
                ) =>
                  setBusca(
                    e.target
                      .value
                  )
                }
              />

              <span className="muted">
                {
                  filtradas.length
                }{" "}
                linha(s)
              </span>
            </div>

            <div className="tableWrap">
              {!linhas.length ? (
                <div className="empty">
                  Carregue o
                  Pedido para
                  gerar a
                  programação.
                </div>
              ) : (
                <table className="table">
                  <thead>
                    <tr>
                      <th>
                        Pedido
                      </th>

                      <th>
                        OF
                      </th>

                      <th>
                        Peça
                      </th>

                      <th>
                        Descrição
                      </th>

                      <th>
                        Material
                      </th>

                      <th>
                        Medida
                      </th>

                      <th>
                        Rebaixo
                      </th>

                      <th>
                        Acabamento
                      </th>

                      <th>
                        Cor
                      </th>

                      <th>
                        Qtd
                      </th>

                      <th>
                        Prioridade
                      </th>

                      <th>
                        Status
                      </th>

                      <th />
                    </tr>
                  </thead>

                  <tbody>
                    {filtradas.map(
                      (
                        linha
                      ) => (
                        <tr
                          key={
                            linha.id
                          }
                        >
                          <td>
                            <input
                              value={
                                linha.pedido
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "pedido",
                                  e
                                    .target
                                    .value
                                )
                              }
                            />
                          </td>

                          <td>
                            <input
                              value={
                                linha.of
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "of",
                                  e
                                    .target
                                    .value
                                )
                              }
                            />
                          </td>

                          <td>
                            <select
                              value={
                                linha.peca
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "peca",
                                  e
                                    .target
                                    .value as TipoPeca
                                )
                              }
                            >
                              <option>
                                BATENTE
                              </option>

                              <option>
                                TRAVESSA BATENTE
                              </option>

                              <option>
                                PORTA
                              </option>

                              <option>
                                ALIZAR
                              </option>

                              <option>
                                PERNA ALIZAR
                              </option>

                              <option>
                                KIT CORRER
                              </option>

                              <option>
                                BAGUETE
                              </option>

                              <option>
                                OUTROS
                              </option>
                            </select>
                          </td>

                          <td
                            style={{
                              minWidth:
                                320,
                            }}
                          >
                            <input
                              value={
                                linha.descricao
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "descricao",
                                  e
                                    .target
                                    .value
                                )
                              }
                            />
                          </td>

                          <td>
                            <input
                              value={
                                linha.material
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "material",
                                  e
                                    .target
                                    .value
                                )
                              }
                            />
                          </td>

                          <td>
                            <input
                              value={
                                linha.medida
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "medida",
                                  e
                                    .target
                                    .value
                                )
                              }
                            />
                          </td>

                          <td>
                            <input
                              value={
                                linha.rebaixo
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "rebaixo",
                                  e
                                    .target
                                    .value
                                )
                              }
                            />
                          </td>

                          <td>
                            <input
                              value={
                                linha.acabamento
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "acabamento",
                                  e
                                    .target
                                    .value
                                )
                              }
                            />
                          </td>

                          <td>
                            <input
                              value={
                                linha.cor
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "cor",
                                  e
                                    .target
                                    .value
                                )
                              }
                            />
                          </td>

                          <td>
                            <input
                              type="number"
                              value={
                                linha.quantidade
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "quantidade",
                                  e
                                    .target
                                    .value
                                )
                              }
                            />
                          </td>

                          <td>
                            <select
                              value={
                                linha.prioridade
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "prioridade",
                                  e
                                    .target
                                    .value as PrioridadeProgramacao
                                )
                              }
                            >
                              <option>
                                NORMAL
                              </option>

                              <option>
                                ALTA
                              </option>

                              <option>
                                URGENTE
                              </option>
                            </select>
                          </td>

                          <td>
                            <select
                              value={
                                linha.status
                              }
                              onChange={(
                                e
                              ) =>
                                alterar(
                                  linha.id,
                                  "status",
                                  e
                                    .target
                                    .value as StatusProgramacao
                                )
                              }
                            >
                              <option>
                                PENDENTE
                              </option>

                              <option>
                                PROGRAMADO
                              </option>

                              <option>
                                APONTADA
                              </option>

                              <option>
                                AGUARDANDO LINHA
                              </option>

                              <option>
                                DIVERGÊNCIA
                              </option>
                            </select>
                          </td>

                          <td>
                            <button
                              className="btn danger"
                              style={{
                                padding: 8,
                              }}
                              onClick={() =>
                                apagar(
                                  linha.id
                                )
                              }
                            >
                              <Trash2
                                size={
                                  14
                                }
                              />
                            </button>
                          </td>
                        </tr>
                      )
                    )}
                  </tbody>
                </table>
              )}
            </div>
          </>
        )}

        {tab ===
          "quantidades" && (
          <div className="tableWrap">
            {!resumo.length ? (
              <div className="empty">
                As quantidades
                consolidadas
                aparecerão
                aqui.
              </div>
            ) : (
              <table className="table">
                <thead>
                  <tr>
                    <th>
                      Grupo
                    </th>

                    <th>
                      Peça
                    </th>

                    <th>
                      Material
                    </th>

                    <th>
                      Medida
                    </th>

                    <th>
                      Rebaixo
                    </th>

                    <th>
                      Acabamento
                    </th>

                    <th>
                      Cor
                    </th>

                    <th>
                      Qtd Total
                    </th>
                  </tr>
                </thead>

                <tbody>
                  {resumo.map(
                    (item) => (
                      <tr
                        key={
                          item.chave
                        }
                      >
                        <td>
                          {
                            item.maquina
                          }
                        </td>

                        <td>
                          {
                            item.peca
                          }
                        </td>

                        <td>
                          {
                            item.material
                          }
                        </td>

                        <td>
                          {
                            item.medida
                          }
                        </td>

                        <td>
                          {
                            item.rebaixo
                          }
                        </td>

                        <td>
                          {
                            item.acabamento
                          }
                        </td>

                        <td>
                          {
                            item.cor
                          }
                        </td>

                        <td>
                          <b>
                            {
                              item.quantidade
                            }
                          </b>
                        </td>
                      </tr>
                    )
                  )}
                </tbody>
              </table>
            )}
          </div>
        )}

        {tab ===
          "apontamentos" && (
          <div className="tableWrap">
            {!linhas.length ? (
              <div className="empty">
                Nenhuma OF
                carregada.
              </div>
            ) : (
              <table className="table">
                <thead>
                  <tr>
                    <th>
                      Pedido
                    </th>

                    <th>
                      OF
                    </th>

                    <th>
                      Peça
                    </th>

                    <th>
                      Descrição
                    </th>

                    <th>
                      Qtd
                    </th>

                    <th>
                      Prioridade
                    </th>

                    <th>
                      Status
                    </th>

                    <th>
                      Observação
                    </th>
                  </tr>
                </thead>

                <tbody>
                  {linhas.map(
                    (
                      linha
                    ) => (
                      <tr
                        key={
                          linha.id
                        }
                      >
                        <td>
                          {
                            linha.pedido
                          }
                        </td>

                        <td>
                          {
                            linha.of
                          }
                        </td>

                        <td>
                          {
                            linha.peca
                          }
                        </td>

                        <td>
                          {
                            linha.descricao
                          }
                        </td>

                        <td>
                          {
                            linha.quantidade
                          }
                        </td>

                        <td>
                          {
                            linha.prioridade
                          }
                        </td>

                        <td>
                          <select
                            value={
                              linha.status
                            }
                            onChange={(
                              e
                            ) =>
                              alterar(
                                linha.id,
                                "status",
                                e
                                  .target
                                  .value as StatusProgramacao
                              )
                            }
                          >
                            <option>
                              PENDENTE
                            </option>

                            <option>
                              PROGRAMADO
                            </option>

                            <option>
                              APONTADA
                            </option>

                            <option>
                              AGUARDANDO LINHA
                            </option>

                            <option>
                              DIVERGÊNCIA
                            </option>
                          </select>
                        </td>

                        <td>
                          <input
                            value={
                              linha.observacao
                            }
                            onChange={(
                              e
                            ) =>
                              alterar(
                                linha.id,
                                "observacao",
                                e
                                  .target
                                  .value
                              )
                            }
                          />
                        </td>
                      </tr>
                    )
                  )}
                </tbody>
              </table>
            )}
          </div>
        )}
      </div>
    </main>
  );
}
