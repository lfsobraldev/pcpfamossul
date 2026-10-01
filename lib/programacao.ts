import {
  LinhaProgramacao,
  ResumoQuantidade,
  TipoPeca,
} from "@/types/programacao";

const clean = (value: unknown) =>
  String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();

const upper = (value: unknown) => clean(value).toUpperCase();

function normalizeTexto(texto: string) {
  return upper(texto)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

/*
|--------------------------------------------------------------------------
| PEÇA
|--------------------------------------------------------------------------
*/

export function inferirPeca(descricao: string): TipoPeca {
  const d = normalizeTexto(descricao);

  if (
    d.includes("TRAVESSA BATENTE") ||
    d.includes("TRAVESSA DE BATENTE") ||
    d.includes("TRAV BATENTE")
  ) {
    return "TRAVESSA BATENTE";
  }

  if (
    d.includes("PERNA ALIZAR") ||
    d.includes("PERNA DE ALIZAR") ||
    d.includes("PERNA ALIZARES")
  ) {
    return "PERNA ALIZAR";
  }

  if (
    d.includes("KIT CORRER") ||
    d.includes("KIT DE CORRER") ||
    d.includes("KIT PORTA CORRER")
  ) {
    return "KIT CORRER";
  }

  if (d.includes("BAGUETE")) {
    return "BAGUETE";
  }

  if (d.includes("ALIZAR") || d.includes("GUARNICAO")) {
    return "ALIZAR";
  }

  if (
    d.includes("FOLHA DE PORTA") ||
    d.includes("FOLHA PORTA") ||
    d.includes("PORTA ")
  ) {
    return "PORTA";
  }

  if (d.includes("BATENTE")) {
    return "BATENTE";
  }

  return "OUTROS";
}

/*
|--------------------------------------------------------------------------
| MATERIAL
|--------------------------------------------------------------------------
*/

export function inferirMaterial(descricao: string) {
  const d = normalizeTexto(descricao);

  const regras = [
    "STD MDF ULTRA",
    "MDF ULTRA",
    "STD MDF",
    "MDF",
    "HDF",
    "PINUS",
    "EUCALIPTO",
    "MADEIRA",
    "COMPENSADO",
  ];

  for (const material of regras) {
    if (d.includes(material)) {
      return material;
    }
  }

  return "";
}

/*
|--------------------------------------------------------------------------
| MEDIDA
|--------------------------------------------------------------------------
*/

export function inferirMedida(descricao: string) {
  const d = normalizeTexto(descricao)
    .replace(/,/g, ".")
    .replace(/\s*X\s*/g, "x");

  const medidas3 = d.match(
    /(\d{2,4}(?:\.\d+)?)x(\d{2,4}(?:\.\d+)?)x(\d{1,4}(?:\.\d+)?)/
  );

  if (medidas3) {
    return `${medidas3[1]}x${medidas3[2]}x${medidas3[3]}`;
  }

  const medidas2 = d.match(
    /(\d{2,4}(?:\.\d+)?)x(\d{2,4}(?:\.\d+)?)/
  );

  if (medidas2) {
    return `${medidas2[1]}x${medidas2[2]}`;
  }

  return "";
}

/*
|--------------------------------------------------------------------------
| REBAIXO
|--------------------------------------------------------------------------
*/

export function inferirRebaixo(descricao: string) {
  const d = normalizeTexto(descricao);

  if (
    d.includes("SEM REBAIXO") ||
    d.includes("S/ REBAIXO") ||
    d.includes("S/REBAIXO") ||
    d.includes("S REBAIXO")
  ) {
    return "SEM REBAIXO";
  }

  const rb = d.match(/\bRB\s*[-:]?\s*(\d{1,3}(?:[.,]\d+)?)\b/);

  if (rb) {
    return `RB ${rb[1].replace(",", ".")}`;
  }

  const rebaixo = d.match(
    /REBAIXO\s*[-:]?\s*(\d{1,3}(?:[.,]\d+)?)/
  );

  if (rebaixo) {
    return `RB ${rebaixo[1].replace(",", ".")}`;
  }

  return "";
}

/*
|--------------------------------------------------------------------------
| ACABAMENTO
|--------------------------------------------------------------------------
*/

export function inferirAcabamento(descricao: string) {
  const d = normalizeTexto(descricao);

  const regras = [
    "PET TXT",
    "PET TEXTURIZADO",
    "PET LISO",
    "PET",
    "TXT",
    "PRIMER",
    "LAMINADO",
    "MELAMINICO",
    "PINTADO",
    "CRU",
  ];

  for (const acabamento of regras) {
    if (d.includes(acabamento)) {
      return acabamento;
    }
  }

  return "";
}

/*
|--------------------------------------------------------------------------
| COR
|--------------------------------------------------------------------------
*/

export function inferirCor(descricao: string) {
  const d = normalizeTexto(descricao);

  const cores = [
    "BRANCO CARRARA",
    "CAPUCCINO MARCHE",
    "CAPPUCCINO MARCHE",
    "BRANCO",
    "PRETO",
    "TURIM",
    "CINZA",
    "CARVALHO",
    "NOGAL",
    "NOCE",
    "MARFIM",
    "AREIA",
    "GRAFITE",
  ];

  for (const cor of cores) {
    if (d.includes(cor)) {
      return cor;
    }
  }

  return "";
}

/*
|--------------------------------------------------------------------------
| MÁQUINA / ABA
|--------------------------------------------------------------------------
*/

export function inferirMaquina(
  peca: TipoPeca,
  descricao: string
) {
  const d = normalizeTexto(descricao);

  /*
    Batente e travessa ficam juntos.
  */

  if (
    peca === "BATENTE" ||
    peca === "TRAVESSA BATENTE"
  ) {
    return "BATENTES";
  }

  if (peca === "PORTA") {
    return "PORTAS";
  }

  if (
    peca === "ALIZAR" ||
    peca === "PERNA ALIZAR"
  ) {
    return "ALIZARES";
  }

  if (peca === "KIT CORRER") {
    return "KIT CORRER";
  }

  if (peca === "BAGUETE") {
    return "BAGUETES";
  }

  /*
    Algumas descrições podem conter palavras específicas
    mesmo quando o tipo não foi identificado.
  */

  if (d.includes("BAGUETE")) return "BAGUETES";

  if (d.includes("CORRER")) return "KIT CORRER";

  return "OUTROS";
}

/*
|--------------------------------------------------------------------------
| CRIA LINHA COMPLETA
|--------------------------------------------------------------------------
*/

export function enriquecerLinha(
  base: Partial<LinhaProgramacao>
): LinhaProgramacao {
  const descricao = clean(base.descricao);

  const peca =
    base.peca ||
    inferirPeca(descricao);

  return {
    id: base.id || crypto.randomUUID(),

    fonte: clean(base.fonte),

    pedido: clean(base.pedido),

    of: clean(base.of),

    peca,

    descricao,

    material:
      clean(base.material) ||
      inferirMaterial(descricao),

    medida:
      clean(base.medida) ||
      inferirMedida(descricao),

    rebaixo:
      clean(base.rebaixo) ||
      inferirRebaixo(descricao),

    acabamento:
      clean(base.acabamento) ||
      inferirAcabamento(descricao),

    cor:
      clean(base.cor) ||
      inferirCor(descricao),

    quantidade: Number(
      base.quantidade || 0
    ),

    prioridade:
      base.prioridade ||
      "NORMAL",

    status:
      base.status ||
      "PENDENTE",

    maquina:
      clean(base.maquina) ||
      inferirMaquina(
        peca,
        descricao
      ),

    observacao:
      clean(base.observacao),
  };
}

/*
|--------------------------------------------------------------------------
| CONSOLIDA QUANTIDADES
|--------------------------------------------------------------------------
*/

export function consolidarQuantidades(
  linhas: LinhaProgramacao[]
): ResumoQuantidade[] {
  const mapa = new Map<
    string,
    ResumoQuantidade
  >();

  for (const linha of linhas) {
    const partes = [
      linha.maquina,
      linha.peca,
      linha.material,
      linha.medida,
      linha.rebaixo,
      linha.acabamento,
      linha.cor,
    ];

    const chave = partes
      .map((x) => upper(x))
      .join("|");

    const existente = mapa.get(chave);

    if (existente) {
      existente.quantidade +=
        Number(linha.quantidade || 0);

      continue;
    }

    mapa.set(chave, {
      chave,

      maquina: linha.maquina,

      peca: linha.peca,

      material: linha.material,

      medida: linha.medida,

      rebaixo: linha.rebaixo,

      acabamento:
        linha.acabamento,

      cor: linha.cor,

      quantidade:
        Number(
          linha.quantidade || 0
        ),
    });
  }

  return [...mapa.values()].sort(
    (a, b) =>
      `${a.maquina}-${a.peca}-${a.medida}`.localeCompare(
        `${b.maquina}-${b.peca}-${b.medida}`
      )
  );
}
