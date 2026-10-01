"use client";
import { useMemo,useState } from "react";
import { FileSpreadsheet,Factory,ClipboardCheck,BarChart3,Download,Trash2,WandSparkles } from "lucide-react";
import { LinhaProgramacao } from "@/types/programacao";
import { lerArquivo,mesclarPedidoUsinagem } from "@/lib/parser";
import { consolidarQuantidades,enriquecerLinha } from "@/lib/programacao";
import { exportarExcel } from "@/lib/exporter";

type Tab="programacao"|"quantidades"|"apontamentos";

export default function Home(){
  const[pedido,setPedido]=useState<File|null>(null);
  const[usinagem,setUsinagem]=useState<File|null>(null);
  const[dataProg,setDataProg]=useState(new Date().toISOString().slice(0,10));
  const[turno,setTurno]=useState("A");
  const[linhas,setLinhas]=useState<LinhaProgramacao[]>([]);
  const[tab,setTab]=useState<Tab>("programacao");
  const[busy,setBusy]=useState(false);
  const[erro,setErro]=useState("");
  const[busca,setBusca]=useState("");

  const resumo=useMemo(()=>consolidarQuantidades(linhas),[linhas]);
  const apontadas=linhas.filter(l=>l.status==="APONTADA").length;
  const pendentes=linhas.filter(l=>l.status==="PENDENTE").length;
  const pecas=linhas.reduce((s,l)=>s+Number(l.quantidade||0),0);

  async function gerar(){
    setErro("");
    if(!pedido&&!usinagem){setErro("Selecione pelo menos um arquivo de Pedido ou Usinagem.");return}
    setBusy(true);
    try{
      const all:LinhaProgramacao[]=[];
      if(pedido)all.push(...await lerArquivo(pedido));
      if(usinagem)all.push(...await lerArquivo(usinagem));
      const merged=mesclarPedidoUsinagem(all).map(enriquecerLinha);
      setLinhas(merged);setTab("programacao");
      if(!merged.length)setErro("Os arquivos foram lidos, mas não encontrei linhas com quantidade. O parser precisará ser ajustado ao layout real.");
    }catch(e:any){setErro(e?.message||"Falha ao processar os arquivos.")}finally{setBusy(false)}
  }

  function patch(id:string,key:keyof LinhaProgramacao,value:any){
    setLinhas(prev=>prev.map(l=>l.id===id?enriquecerLinha({...l,[key]:key==="quantidade"?Number(value):value}):l));
  }
  function apagar(id:string){setLinhas(prev=>prev.filter(l=>l.id!==id))}
  const filtradas=linhas.filter(l=>{const q=busca.toUpperCase().trim();return !q||[l.of,l.pedido,l.item,l.descricao,l.maquina,l.medida,l.tipoPeca].join(" ").toUpperCase().includes(q)});

  return <main className="shell">
    <div className="topbar"><div className="brand">SOBRAL <span style={{opacity:.72}}>PROGRAMAÇÃO INDUSTRIAL</span><small>PLANEJAMENTO DE PRODUÇÃO</small></div><span className="badge">Famossul • Operação diária</span></div>
    <div className="wrap">
      <div className="hero"><div><h1>Gerador de Programação de Produção</h1><p>Pedido + Usinagem → revisão → líderes → apontadores → Excel.</p></div></div>
      <section className="grid">
        <div className="card kpi"><Factory size={20}/><span className="muted">Linhas / OFs</span><strong>{linhas.length}</strong></div>
        <div className="card kpi"><FileSpreadsheet size={20}/><span className="muted">Peças programadas</span><strong>{pecas.toLocaleString("pt-BR")}</strong></div>
        <div className="card kpi"><ClipboardCheck size={20}/><span className="muted">Apontadas</span><strong>{apontadas}</strong></div>
        <div className="card kpi"><BarChart3 size={20}/><span className="muted">Pendentes</span><strong>{pendentes}</strong></div>
      </section>
      <section className="upload">
        <div className="drop field"><label>PEDIDO</label><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>setPedido(e.target.files?.[0]||null)}/><div className="muted" style={{marginTop:8}}>{pedido?.name||"XLSX, XLS ou CSV"}</div></div>
        <div className="drop field"><label>USINAGEM</label><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>setUsinagem(e.target.files?.[0]||null)}/><div className="muted" style={{marginTop:8}}>{usinagem?.name||"XLSX, XLS ou CSV"}</div></div>
        <div className="field card"><label>DATA DA PROGRAMAÇÃO</label><input type="date" value={dataProg} onChange={e=>setDataProg(e.target.value)}/></div>
        <div className="field card"><label>TURNO</label><select value={turno} onChange={e=>setTurno(e.target.value)}><option>A</option><option>B</option><option>A + B</option></select></div>
      </section>
      {erro&&<div className="notice">{erro}</div>}
      <div className="actions">
        <button className="btn primary" onClick={gerar} disabled={busy}><WandSparkles size={16} style={{verticalAlign:"middle",marginRight:7}}/>{busy?"PROCESSANDO...":"GERAR PROGRAMAÇÃO"}</button>
        <button className="btn secondary" disabled={!linhas.length} onClick={()=>exportarExcel(linhas,dataProg,turno)}><Download size={16} style={{verticalAlign:"middle",marginRight:7}}/>EXPORTAR EXCEL</button>
        <button className="btn danger" disabled={!linhas.length} onClick={()=>setLinhas([])}>LIMPAR</button>
      </div>
      <div className="tabs">
        <button className={`tab ${tab==="programacao"?"active":""}`} onClick={()=>setTab("programacao")}>Programação</button>
        <button className={`tab ${tab==="quantidades"?"active":""}`} onClick={()=>setTab("quantidades")}>Quantidades líderes</button>
        <button className={`tab ${tab==="apontamentos"?"active":""}`} onClick={()=>setTab("apontamentos")}>Apontamentos</button>
      </div>
      {tab==="programacao"&&<>
        <div className="toolbar"><input className="search" placeholder="Buscar OF, pedido, máquina, medida..." value={busca} onChange={e=>setBusca(e.target.value)}/><span className="muted">{filtradas.length} linha(s)</span></div>
        <div className="tableWrap">{!linhas.length?<div className="empty">Carregue Pedido e Usinagem para gerar a programação.</div>:<table className="table"><thead><tr><th>Máquina</th><th>OF</th><th>Pedido</th><th>Item</th><th>Peça</th><th>Descrição</th><th>Medida</th><th>Rebaixo</th><th>Lado</th><th>Qtd</th><th></th></tr></thead><tbody>{filtradas.map(l=><tr key={l.id}>
          <td><input value={l.maquina} onChange={e=>patch(l.id,"maquina",e.target.value)}/></td>
          <td><input value={l.of} onChange={e=>patch(l.id,"of",e.target.value)}/></td>
          <td><input value={l.pedido} onChange={e=>patch(l.id,"pedido",e.target.value)}/></td>
          <td><input value={l.item} onChange={e=>patch(l.id,"item",e.target.value)}/></td>
          <td><input value={l.tipoPeca} onChange={e=>patch(l.id,"tipoPeca",e.target.value)}/></td>
          <td style={{minWidth:280}}><input value={l.descricao} onChange={e=>patch(l.id,"descricao",e.target.value)}/></td>
          <td><input value={l.medida} onChange={e=>patch(l.id,"medida",e.target.value)}/></td>
          <td><input value={l.rebaixo} onChange={e=>patch(l.id,"rebaixo",e.target.value)}/></td>
          <td><input value={l.lado} onChange={e=>patch(l.id,"lado",e.target.value)}/></td>
          <td><input type="number" value={l.quantidade} onChange={e=>patch(l.id,"quantidade",e.target.value)}/></td>
          <td><button className="btn danger" style={{padding:8}} onClick={()=>apagar(l.id)}><Trash2 size={14}/></button></td>
        </tr>)}</tbody></table>}</div>
      </>}
      {tab==="quantidades"&&<div className="tableWrap">{!resumo.length?<div className="empty">As quantidades consolidadas aparecerão aqui.</div>:<table className="table"><thead><tr><th>Máquina</th><th>Peça</th><th>Material</th><th>Acabamento</th><th>Cor</th><th>Medida</th><th>Rebaixo</th><th>Qtd Total</th></tr></thead><tbody>{resumo.map(r=><tr key={r.chave}><td>{r.maquina}</td><td>{r.tipoPeca}</td><td>{r.material}</td><td>{r.acabamento}</td><td>{r.cor}</td><td>{r.medida}</td><td>{r.rebaixo}</td><td><b>{r.quantidade}</b></td></tr>)}</tbody></table>}</div>}
      {tab==="apontamentos"&&<div className="tableWrap">{!linhas.length?<div className="empty">Nenhuma OF carregada.</div>:<table className="table"><thead><tr><th>OF</th><th>Pedido</th><th>Item</th><th>Máquina</th><th>Descrição</th><th>Qtd</th><th>Status</th><th>Observação</th></tr></thead><tbody>{linhas.map(l=><tr key={l.id}><td>{l.of}</td><td>{l.pedido}</td><td>{l.item}</td><td>{l.maquina}</td><td>{l.descricao}</td><td>{l.quantidade}</td><td><select value={l.status} onChange={e=>patch(l.id,"status",e.target.value)}><option>PENDENTE</option><option>APONTADA</option><option>AGUARDANDO LINHA</option><option>DIVERGÊNCIA</option></select></td><td><input value={l.observacao} onChange={e=>patch(l.id,"observacao",e.target.value)}/></td></tr>)}</tbody></table>}</div>}
    </div>
  </main>
}
