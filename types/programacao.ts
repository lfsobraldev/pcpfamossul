export type StatusProgramacao =
  | "PENDENTE"
  | "PROGRAMADO"
  | "APONTADA"
  | "AGUARDANDO LINHA"
  | "DIVERGÊNCIA";

export type PrioridadeProgramacao =
  | "NORMAL"
  | "ALTA"
  | "URGENTE";

export type TipoPeca =
  | "BATENTE"
  | "TRAVESSA BATENTE"
  | "PORTA"
  | "ALIZAR"
  | "PERNA ALIZAR"
  | "KIT CORRER"
  | "BAGUETE"
  | "OUTROS";

export type LinhaProgramacao = {
  id: string;

  fonte: string;

  pedido: string;
  of: string;

  peca: TipoPeca;

  descricao: string;

  material: string;

  medida: string;

  rebaixo: string;

  acabamento: string;

  cor: string;

  quantidade: number;

  prioridade: PrioridadeProgramacao;

  status: StatusProgramacao;

  maquina: string;

  observacao: string;
};

export type ResumoQuantidade = {
  chave: string;

  maquina: string;

  peca: TipoPeca;

  material: string;

  medida: string;

  rebaixo: string;

  acabamento: string;

  cor: string;

  quantidade: number;
};
