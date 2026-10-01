export type StatusApontamento = "PENDENTE" | "APONTADA" | "AGUARDANDO LINHA" | "DIVERGÊNCIA";

export type LinhaProgramacao = {
  id: string;
  fonte: string;
  pedido: string;
  item: string;
  of: string;
  descricao: string;
  tipoPeca: string;
  material: string;
  acabamento: string;
  cor: string;
  medida: string;
  rebaixo: string;
  lado: string;
  processo: string;
  maquina: string;
  quantidade: number;
  status: StatusApontamento;
  observacao: string;
};

export type ResumoQuantidade = {
  chave: string;
  maquina: string;
  tipoPeca: string;
  material: string;
  acabamento: string;
  cor: string;
  medida: string;
  rebaixo: string;
  quantidade: number;
};
