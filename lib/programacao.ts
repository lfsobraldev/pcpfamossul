import { LinhaProgramacao, ResumoQuantidade } from "@/types/programacao";

const clean = (v: unknown) => String(v ?? "").replace(/\s+/g, " ").trim();
const upper = (v: unknown) => clean(v).toUpperCase();

export function inferirMedida(descricao: string): string {
  const t = upper(descricao).replace(/,/g, ".");
  const m3 = t.match(/(\d{2,4}(?:\.\d+)?)\s*[X×]\s*(\d{2,4}(?:\.\d+)?)\s*[X×]\s*(\d{1,4}(?:\.\d+)?)/);
  if (m3) return `${m3[1]}x${m3[2]}x${m3[3]}`;
  const m2 = t.match(/(\d{2,4}(?:\.\d+)?)\s*[X×]\s*(\d{2,4}(?:\.\d+)?)/);
  return m2 ? `${m2[1]}x${m2[2]}` : "";
}

export function inferirTipoPeca(descricao: string): string {
  const d = upper(descricao);
  if (d.includes("TRAVESS")) return "TRAVESSA BATENTE";
  if (d.includes("PERNA") && d.includes("BATENTE")) return "PERNA BATENTE";
  if (d.includes("BATENTE")) return "BATENTE";
  if (d.includes("ALIZAR") || d.includes("GUARNI")) return "ALIZAR";
  if (d.includes("PORTA") || d.includes("FOLHA")) return "PORTA";
  if (d.includes("DOBRADI")) return "DOBRADIÇA";
  if (d.includes("CONTRATESTA")) return "CONTRATESTA";
  if (d.includes("FECHADURA") || d.includes("FERRAGEM")) return "FERRAGEM";
  return "OUTROS";
}

export function inferirMaterial(descricao: string): string {
  const d = upper(descricao);
  return ["PINUS","MDF","HDF","EUCALIPTO","MADEIRA"].find(x => d.includes(x)) ?? "";
}

export function inferirAcabamento(descricao: string): string {
  const d = upper(descricao);
  return ["PET","PRIMER","LAMINADO","MELAMÍNICO","MELAMINICO","CRU","PINTADO"].find(x => d.includes(x)) ?? "";
}

export function inferirCor(descricao: string): string {
  const d = upper(descricao);
  return ["BRANCO CARRARA","CAPUCCINO MARCHE","CAPPUCCINO MARCHE","BRANCO","PRETO","TURIM","CINZA","CARVALHO","NOCE","NOGAL"].find(x => d.includes(x)) ?? "";
}

export function inferirLado(descricao: string): string {
  const d = upper(descricao);
  if (/\bDIREIT[AO]\b|\bDIR\b/.test(d)) return "DIREITA";
  if (/\bESQUERD[AO]\b|\bESQ\b/.test(d)) return "ESQUERDA";
  return "";
}

export function inferirRebaixo(descricao: string): string {
  const m = upper(descricao).match(/REBAIXO\s*[:\-]?\s*(\d+(?:[.,]\d+)?)\s*(MM)?/);
  return m ? `${m[1].replace(",", ".")} mm` : "";
}

export function inferirMaquina(tipoPeca: string, descricao: string, processo = ""): string {
  const d = `${upper(tipoPeca)} ${upper(descricao)} ${upper(processo)}`;
  if (d.includes("PREPARA")) return "PREPARAÇÃO";
  if (d.includes("MOLDUREIRA")) return "MOLDUREIRA";
  if (d.includes("RECOBR")) return "RECOBRIDORA";
  if (d.includes("LIX")) return "LIXADEIRA";
  if (d.includes("USIN")) return "USINAGEM";
  if (d.includes("SERRA") || d.includes("CORTE")) return "SERRA / CORTE";
  if (d.includes("MONT")) return "MONTAGEM";
  if (d.includes("TRAVESSA") || d.includes("PERNA BATENTE") || d.includes("BATENTE")) return "PREPARAÇÃO";
  if (d.includes("ALIZAR")) return "RECOBRIDORA";
  if (d.includes("PORTA") || d.includes("FOLHA")) return "USINAGEM";
  if (d.includes("DOBRADIÇA") || d.includes("CONTRATESTA") || d.includes("FERRAGEM")) return "MONTAGEM";
  return "OUTROS";
}

export function enriquecerLinha(base: Partial<LinhaProgramacao>): LinhaProgramacao {
  const descricao = clean(base.descricao);
  const tipoPeca = clean(base.tipoPeca) || inferirTipoPeca(descricao);
  const processo = clean(base.processo);
  return {
    id: base.id || crypto.randomUUID(),
    fonte: clean(base.fonte),
    pedido: clean(base.pedido),
    item: clean(base.item),
    of: clean(base.of),
    descricao,
    tipoPeca,
    material: clean(base.material) || inferirMaterial(descricao),
    acabamento: clean(base.acabamento) || inferirAcabamento(descricao),
    cor: clean(base.cor) || inferirCor(descricao),
    medida: clean(base.medida) || inferirMedida(descricao),
    rebaixo: clean(base.rebaixo) || inferirRebaixo(descricao),
    lado: clean(base.lado) || inferirLado(descricao),
    processo,
    maquina: clean(base.maquina) || inferirMaquina(tipoPeca, descricao, processo),
    quantidade: Number(base.quantidade || 0),
    status: base.status || "PENDENTE",
    observacao: clean(base.observacao),
  };
}

export function consolidarQuantidades(linhas: LinhaProgramacao[]): ResumoQuantidade[] {
  const mapa = new Map<string, ResumoQuantidade>();
  for (const l of linhas) {
    const chave = [l.maquina,l.tipoPeca,l.material,l.acabamento,l.cor,l.medida,l.rebaixo].map(upper).join("|");
    const existente = mapa.get(chave);
    if (existente) existente.quantidade += Number(l.quantidade || 0);
    else mapa.set(chave, {
      chave, maquina:l.maquina, tipoPeca:l.tipoPeca, material:l.material,
      acabamento:l.acabamento, cor:l.cor, medida:l.medida, rebaixo:l.rebaixo,
      quantidade:Number(l.quantidade || 0)
    });
  }
  return [...mapa.values()].sort((a,b)=>`${a.maquina}-${a.tipoPeca}-${a.medida}`.localeCompare(`${b.maquina}-${b.tipoPeca}-${b.medida}`));
}
