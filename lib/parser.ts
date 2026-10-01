import * as XLSX from "xlsx";
import { LinhaProgramacao } from "@/types/programacao";
import { enriquecerLinha } from "@/lib/programacao";

const normalize = (s:string) => s.normalize("NFD").replace(/[\u0300-\u036f]/g,"").toUpperCase().replace(/[^A-Z0-9]/g,"");

const aliases:any = {
  pedido:["PEDIDO","NUMEROPEDIDO","NRPEDIDO","NROPEDIDO"],
  item:["ITEM","NRITEM","NUMEROITEM"],
  of:["OF","ORDEMFABRICACAO","ORDEMDEFABRICACAO","ORDEM"],
  descricao:["DESCRICAO","PRODUTO","DESCRICAOPRODUTO","DESCITEM","ITEMDESCRICAO"],
  tipoPeca:["TIPOPECA","TIPO","PECA"],
  material:["MATERIAL","MADEIRA"],
  acabamento:["ACABAMENTO","REVESTIMENTO"],
  cor:["COR","PADRAO","PADRAOCOR"],
  medida:["MEDIDA","DIMENSAO","DIMENSOES","BITOLA"],
  rebaixo:["REBAIXO"],
  lado:["LADO","MAO"],
  processo:["PROCESSO","OPERACAO"],
  maquina:["MAQUINA","RECURSO","CENTROTRABALHO"],
  quantidade:["QUANTIDADE","QTD","QTDE","QUANT","QTDPROGRAMADA"]
};

function findValue(row:Record<string,unknown>, field:string) {
  for (const [k,v] of Object.entries(row)) if (aliases[field].includes(normalize(k))) return v;
  return "";
}

function rowToLinha(row:Record<string,unknown>, fonte:string):LinhaProgramacao {
  return enriquecerLinha({
    fonte,
    pedido:String(findValue(row,"pedido") ?? ""),
    item:String(findValue(row,"item") ?? ""),
    of:String(findValue(row,"of") ?? ""),
    descricao:String(findValue(row,"descricao") ?? ""),
    tipoPeca:String(findValue(row,"tipoPeca") ?? ""),
    material:String(findValue(row,"material") ?? ""),
    acabamento:String(findValue(row,"acabamento") ?? ""),
    cor:String(findValue(row,"cor") ?? ""),
    medida:String(findValue(row,"medida") ?? ""),
    rebaixo:String(findValue(row,"rebaixo") ?? ""),
    lado:String(findValue(row,"lado") ?? ""),
    processo:String(findValue(row,"processo") ?? ""),
    maquina:String(findValue(row,"maquina") ?? ""),
    quantidade:Number(findValue(row,"quantidade") || 0),
  });
}

export async function lerPlanilha(file:File):Promise<LinhaProgramacao[]> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf,{type:"array"});
  const linhas:LinhaProgramacao[] = [];
  for (const nome of wb.SheetNames) {
    const ws = wb.Sheets[nome];
    const rows = XLSX.utils.sheet_to_json<Record<string,unknown>>(ws,{defval:""});
    for (const row of rows) {
      const l = rowToLinha(row, `${file.name} / ${nome}`);
      if ((l.descricao || l.of || l.pedido || l.item) && l.quantidade !== 0) linhas.push(l);
    }
  }
  return linhas;
}

export async function lerArquivo(file:File):Promise<LinhaProgramacao[]> {
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (["xlsx","xls","csv"].includes(ext || "")) return lerPlanilha(file);
  throw new Error(`Nesta primeira versão use XLSX, XLS ou CSV. PDF será habilitado na próxima etapa: ${file.name}`);
}

export function mesclarPedidoUsinagem(linhas:LinhaProgramacao[]):LinhaProgramacao[] {
  const byKey = new Map<string,LinhaProgramacao>();
  const keyOf = (l:LinhaProgramacao) => l.of ? `OF:${l.of}` : (l.pedido && l.item ? `PI:${l.pedido}:${l.item}` : `DESC:${l.descricao}:${l.medida}:${l.quantidade}`);
  for (const linha of linhas) {
    const key = keyOf(linha);
    const atual = byKey.get(key);
    if (!atual) { byKey.set(key,linha); continue; }
    byKey.set(key,{
      ...atual,
      pedido:atual.pedido || linha.pedido,
      item:atual.item || linha.item,
      of:atual.of || linha.of,
      descricao:atual.descricao.length >= linha.descricao.length ? atual.descricao : linha.descricao,
      tipoPeca:atual.tipoPeca || linha.tipoPeca,
      material:atual.material || linha.material,
      acabamento:atual.acabamento || linha.acabamento,
      cor:atual.cor || linha.cor,
      medida:atual.medida || linha.medida,
      rebaixo:atual.rebaixo || linha.rebaixo,
      lado:atual.lado || linha.lado,
      processo:atual.processo || linha.processo,
      maquina:atual.maquina || linha.maquina,
      quantidade:Math.max(atual.quantidade,linha.quantidade),
      fonte:`${atual.fonte} + ${linha.fonte}`
    });
  }
  return [...byKey.values()];
}
