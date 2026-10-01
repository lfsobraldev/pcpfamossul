import ExcelJS from "exceljs";
import { LinhaProgramacao } from "@/types/programacao";
import { consolidarQuantidades } from "@/lib/programacao";

const safe = (s:string) => (s||"OUTROS").replace(/[\\/*?:[\]]/g," ").slice(0,31);

function titulo(ws:ExcelJS.Worksheet,texto:string,cols:number){
  ws.mergeCells(1,1,1,cols);
  const c=ws.getCell(1,1);
  c.value=texto;
  c.font={bold:true,size:16,color:{argb:"FFFFFFFF"}};
  c.alignment={horizontal:"center",vertical:"middle"};
  c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF17365D"}};
  ws.getRow(1).height=28;
}
function cab(row:ExcelJS.Row){
  row.font={bold:true,color:{argb:"FFFFFFFF"}};
  row.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF244062"}};
  row.alignment={horizontal:"center",vertical:"middle",wrapText:true};
}
function bordas(ws:ExcelJS.Worksheet,ini:number,fim:number,cols:number){
  for(let r=ini;r<=fim;r++)for(let c=1;c<=cols;c++){
    ws.getCell(r,c).border={
      top:{style:"thin",color:{argb:"FFD9E2F3"}},
      left:{style:"thin",color:{argb:"FFD9E2F3"}},
      bottom:{style:"thin",color:{argb:"FFD9E2F3"}},
      right:{style:"thin",color:{argb:"FFD9E2F3"}}
    };
    ws.getCell(r,c).alignment={vertical:"middle",wrapText:true};
  }
}

export async function exportarExcel(linhas:LinhaProgramacao[],dataProg:string,turno:string){
  const wb=new ExcelJS.Workbook();
  wb.creator="Sobral Programação Industrial";

  const ger=wb.addWorksheet("CONTROLE GERENTE");
  titulo(ger,`CONTROLE GERENCIAL • ${dataProg || "SEM DATA"} • TURNO ${turno}`,6);
  ger.addRow([]);
  ger.addRow(["Indicador","Valor"]); cab(ger.getRow(3));
  const apontadas=linhas.filter(l=>l.status==="APONTADA").length;
  ger.addRow(["OFs / linhas programadas",linhas.length]);
  ger.addRow(["Peças programadas",linhas.reduce((s,l)=>s+Number(l.quantidade||0),0)]);
  ger.addRow(["Apontadas",apontadas]);
  ger.addRow(["Pendentes",linhas.filter(l=>l.status==="PENDENTE").length]);
  ger.addRow(["Aguardando linha",linhas.filter(l=>l.status==="AGUARDANDO LINHA").length]);
  ger.addRow(["Divergências",linhas.filter(l=>l.status==="DIVERGÊNCIA").length]);
  ger.addRow(["% apontado",linhas.length ? apontadas/linhas.length : 0]);
  ger.getCell("B9").numFmt="0.0%";
  ger.columns=[{width:32},{width:18}];

  const apt=wb.addWorksheet("APONTAMENTOS");
  titulo(apt,`CONTROLE DOS APONTADORES • ${dataProg || "SEM DATA"} • TURNO ${turno}`,9);
  apt.addRow([]);
  apt.addRow(["OF","Pedido","Item","Máquina","Descrição","Medida","Qtd","Status","Observação"]); cab(apt.getRow(3));
  for(const l of linhas) apt.addRow([l.of,l.pedido,l.item,l.maquina,l.descricao,l.medida,l.quantidade,l.status,l.observacao]);
  apt.columns=[{width:15},{width:14},{width:10},{width:20},{width:48},{width:20},{width:10},{width:20},{width:30}];
  bordas(apt,3,apt.rowCount,9);

  const qtd=wb.addWorksheet("QUANTIDADES LIDERES");
  titulo(qtd,`RESUMO DE QUANTIDADES PARA LÍDERES • ${dataProg || "SEM DATA"} • TURNO ${turno}`,8);
  qtd.addRow([]);
  qtd.addRow(["Máquina","Peça","Material","Acabamento","Cor","Medida","Rebaixo","Qtd Total"]); cab(qtd.getRow(3));
  for(const r of consolidarQuantidades(linhas)) qtd.addRow([r.maquina,r.tipoPeca,r.material,r.acabamento,r.cor,r.medida,r.rebaixo,r.quantidade]);
  qtd.columns=[{width:20},{width:24},{width:20},{width:20},{width:24},{width:20},{width:14},{width:12}];
  bordas(qtd,3,qtd.rowCount,8);

  const maquinas=[...new Set(linhas.map(l=>l.maquina || "OUTROS"))].sort();
  for(const maquina of maquinas){
    const ws=wb.addWorksheet(safe(maquina));
    titulo(ws,`${maquina} • PROGRAMAÇÃO DO LÍDER • ${dataProg || "SEM DATA"} • TURNO ${turno}`,7);
    ws.addRow([]);
    ws.addRow(["OF","Pedido","Item","Descrição","Medida / Rebaixo","Lado","Qtd"]); cab(ws.getRow(3));
    for(const l of linhas.filter(x=>x.maquina===maquina))
      ws.addRow([l.of,l.pedido,l.item,l.descricao,[l.medida,l.rebaixo].filter(Boolean).join(" • "),l.lado,l.quantidade]);
    ws.columns=[{width:15},{width:14},{width:10},{width:55},{width:26},{width:14},{width:12}];
    ws.pageSetup={orientation:"landscape",fitToPage:true,fitToWidth:1,fitToHeight:0,paperSize:9};
    bordas(ws,3,ws.rowCount,7);
    for(let r=4;r<=ws.rowCount;r++) ws.getRow(r).height=34;
  }

  const buf=await wb.xlsx.writeBuffer();
  const blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  a.href=url;
  a.download=`Programacao_Producao_${dataProg || "sem-data"}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}
