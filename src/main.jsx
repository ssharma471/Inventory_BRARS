import React,{useMemo,useState} from 'react';
import{createRoot}from'react-dom/client';
import*as XLSX from'xlsx';
import seed from'./data/inventory.json';
import'./styles.css';

const money=n=>new Intl.NumberFormat('en-CA',{style:'currency',currency:'CAD'}).format(Number(n)||0);
const monthKey=()=>new Date().toISOString().slice(0,7);
const load=(k,d)=>{try{return JSON.parse(localStorage.getItem(k))??d}catch{return d}};

function App(){
 const CATEGORY_OPTIONS=['Finished','Packaging','Supplies','OIC'];
 const[products,setProducts]=useState(()=>load('inventory-products-master-apr2026-four-categories-v3',seed));
 const[month,setMonth]=useState(monthKey());
 const[counts,setCounts]=useState(()=>load('inventory-counts',{}));
 const[search,setSearch]=useState(''); const[cat,setCat]=useState('All'); const[tab,setTab]=useState('count');
 const[showAdd,setShowAdd]=useState(false); const[form,setForm]=useState({name:'',vendor:'',unitPrice:'',category:'Finished'});
 const monthCounts=counts[month]||{};
 const categories=['All',...CATEGORY_OPTIONS];
 const filtered=products.filter(p=>p.active!==false&&(cat==='All'||p.category===cat)&&(`${p.name} ${p.vendor}`.toLowerCase().includes(search.toLowerCase())));
 const total=products.filter(p=>p.active!==false).reduce((s,p)=>s+(Number(monthCounts[p.id]??p.quantity??0)*Number(p.unitPrice||0)),0);
 const counted=products.filter(p=>p.active!==false&&monthCounts[p.id]!==undefined).length;
 const summary=useMemo(()=>products.filter(p=>p.active!==false).reduce((a,p)=>{const k=CATEGORY_OPTIONS.includes(p.category)?p.category:'Finished';a[k]+=Number(monthCounts[p.id]??p.quantity??0)*Number(p.unitPrice||0);return a},{Finished:0,Packaging:0,Supplies:0,OIC:0}),[products,monthCounts]);
 const persistProducts=next=>{setProducts(next);localStorage.setItem('inventory-products-master-apr2026-four-categories-v3',JSON.stringify(next))};
 const setQty=(id,v)=>{const next={...counts,[month]:{...monthCounts,[id]:v===''?0:Number(v)}};setCounts(next);localStorage.setItem('inventory-counts',JSON.stringify(next))};
 const addProduct=e=>{e.preventDefault();if(!form.name.trim())return;persistProducts([...products,{id:crypto.randomUUID(),name:form.name.trim(),vendor:form.vendor.trim(),unitPrice:Number(form.unitPrice)||0,quantity:0,category:CATEGORY_OPTIONS.includes(form.category)?form.category:'Finished',active:true}]);setForm({name:'',vendor:'',unitPrice:'',category:'Finished'});setShowAdd(false)};
 const updateProduct=(id,key,value)=>persistProducts(products.map(p=>p.id===id?{...p,[key]:key==='unitPrice'?Number(value):value}:p));
 const deactivate=id=>persistProducts(products.map(p=>p.id===id?{...p,active:false}:p));
 const newMonth=()=>{if(confirm(`Start ${month} using blank quantities?`)){const next={...counts,[month]:{}};setCounts(next);localStorage.setItem('inventory-counts',JSON.stringify(next))}};
 const exportExcel=()=>{
  const activeProducts=products.filter(p=>p.active!==false);
  const rows=activeProducts.map(p=>({Name:p.name,Vendor:p.vendor,'Unit Price':Number(p.unitPrice||0),'Quantity in Stock':Number(monthCounts[p.id]??p.quantity??0),'Inventory Value':Number(p.unitPrice||0)*Number(monthCounts[p.id]??p.quantity??0),Category:CATEGORY_OPTIONS.includes(p.category)?p.category:'Finished'}));
  const ws=XLSX.utils.json_to_sheet(rows,{origin:'A4'});
  XLSX.utils.sheet_add_aoa(ws,[["BRAR'S VAUGHAN",'',`Month-End Inventory: ${month}`]],{origin:'A1'});

  // Add the grand total and category totals directly under the inventory list.
  const firstDataRow=5;
  const lastDataRow=firstDataRow+rows.length-1;
  const summaryStart=lastDataRow+4;
  const preferredOrder=['Finished','Packaging','Supplies','OIC'];
  const categoryOrder=preferredOrder;

  XLSX.utils.sheet_add_aoa(ws,[['','','','Total Inventory Value',total]],{origin:`A${summaryStart}`});
  categoryOrder.forEach((category,index)=>{
   XLSX.utils.sheet_add_aoa(ws,[['',category,'=', '',Number(summary[category]||0)]],{origin:`A${summaryStart+2+index}`});
  });

  ws['!cols']=[{wch:48},{wch:20},{wch:14},{wch:18},{wch:18},{wch:18}];
  ws['!freeze']={xSplit:0,ySplit:4,topLeftCell:'A5',activePane:'bottomLeft',state:'frozen'};
  ws['!autofilter']={ref:`A4:F${lastDataRow}`};
  ws['!merges']=[...(ws['!merges']||[]),{s:{r:0,c:0},e:{r:0,c:1}}];

  // Currency and quantity formats.
  for(let r=firstDataRow;r<=lastDataRow;r++){
   if(ws[`C${r}`])ws[`C${r}`].z='$#,##0.00';
   if(ws[`D${r}`])ws[`D${r}`].z='0.00';
   if(ws[`E${r}`])ws[`E${r}`].z='$#,##0.00';
  }
  if(ws[`E${summaryStart}`])ws[`E${summaryStart}`].z='$#,##0.00';
  categoryOrder.forEach((_,index)=>{const cell=ws[`E${summaryStart+2+index}`];if(cell)cell.z='$#,##0.00';});

  const sumRows=categoryOrder.map(k=>({Category:k,'Inventory Value':Number(summary[k]||0)}));
  const sws=XLSX.utils.json_to_sheet(sumRows);
  XLSX.utils.sheet_add_aoa(sws,[['Grand Total',total]],{origin:-1});
  sws['!cols']=[{wch:24},{wch:20}];
  for(let r=2;r<=sumRows.length+2;r++){if(sws[`B${r}`])sws[`B${r}`].z='$#,##0.00';}

  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Inventory');
  XLSX.utils.book_append_sheet(wb,sws,'Summary');
  XLSX.writeFile(wb,`${month}-month-end-inventory.xlsx`);
 };
 return <div className="app">
  <aside><div className="brand"><span>ME</span><div><b>Month-End</b><small>Inventory</small></div></div>
   <nav><button className={tab==='count'?'active':''} onClick={()=>setTab('count')}>Inventory Count</button><button className={tab==='products'?'active':''} onClick={()=>setTab('products')}>Products</button><button className={tab==='history'?'active':''} onClick={()=>setTab('history')}>Monthly History</button></nav>
   <div className="sideNote">Data is saved in this browser. Export the completed month to Excel for your accountant.</div>
  </aside>
  <main><header><div><h1>{tab==='count'?'Month-End Inventory':tab==='products'?'Product Management':'Monthly History'}</h1><p>BRAR'S VAUGHAN</p></div><div className="actions"><input type="month" value={month} onChange={e=>setMonth(e.target.value)}/><button className="secondary" onClick={newMonth}>Reset Month</button><button className="primary" onClick={exportExcel}>Export Excel</button></div></header>
   <section className="cards"><div><small>Total inventory value</small><strong>{money(total)}</strong></div><div><small>Active products</small><strong>{products.filter(p=>p.active!==false).length}</strong></div><div><small>Quantities entered</small><strong>{counted}</strong></div><div><small>Remaining</small><strong>{products.filter(p=>p.active!==false).length-counted}</strong></div></section>
   {tab==='history'?<History counts={counts} products={products} setMonth={m=>{setMonth(m);setTab('count')}}/>:<>
   <section className="toolbar"><input placeholder="Search product or vendor…" value={search} onChange={e=>setSearch(e.target.value)}/><select value={cat} onChange={e=>setCat(e.target.value)}>{categories.map(c=><option key={c}>{c}</option>)}</select>{tab==='products'&&<button className="primary" onClick={()=>setShowAdd(true)}>+ Add Product</button>}</section>
   <section className="tableWrap"><table><thead><tr><th>Product</th><th>Vendor</th><th>Category</th><th>Unit Price</th>{tab==='count'?<><th>Quantity</th><th>Inventory Value</th></>:<th>Action</th>}</tr></thead><tbody>{filtered.map(p=><tr key={p.id}><td className="product">{tab==='products'?<input value={p.name} onChange={e=>updateProduct(p.id,'name',e.target.value)}/>:p.name}</td><td>{tab==='products'?<input value={p.vendor} onChange={e=>updateProduct(p.id,'vendor',e.target.value)}/>:p.vendor||'—'}</td><td>{tab==='products'?<select value={p.category} onChange={e=>updateProduct(p.id,'category',e.target.value)}>{categories.filter(c=>c!=='All').map(c=><option key={c}>{c}</option>)}</select>:<span className="pill">{p.category}</span>}</td><td>{tab==='products'?<input type="number" step="0.01" value={p.unitPrice} onChange={e=>updateProduct(p.id,'unitPrice',e.target.value)}/>:money(p.unitPrice)}</td>{tab==='count'?<><td><input className="qty" type="number" min="0" step="0.01" value={monthCounts[p.id]??p.quantity??''} onChange={e=>setQty(p.id,e.target.value)}/></td><td className="value">{money(Number(p.unitPrice||0)*Number(monthCounts[p.id]??p.quantity??0))}</td></>:<td><button className="danger" onClick={()=>deactivate(p.id)}>Deactivate</button></td>}</tr>)}</tbody></table></section>
   {tab==='count'&&<section className="summary"><h2>Category summary</h2><div>{CATEGORY_OPTIONS.map(k=><article key={k}><span>{k}</span><b>{money(summary[k])}</b></article>)}</div></section>}</>}
  </main>
  {showAdd&&<div className="modal"><form onSubmit={addProduct}><h2>Add product</h2><label>Product name<input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} required/></label><label>Vendor<input value={form.vendor} onChange={e=>setForm({...form,vendor:e.target.value})}/></label><label>Unit price<input type="number" step="0.01" value={form.unitPrice} onChange={e=>setForm({...form,unitPrice:e.target.value})}/></label><label>Category<select value={form.category} onChange={e=>setForm({...form,category:e.target.value})}>{categories.filter(c=>c!=='All').map(c=><option key={c}>{c}</option>)}</select></label><div><button type="button" className="secondary" onClick={()=>setShowAdd(false)}>Cancel</button><button className="primary">Add Product</button></div></form></div>}
 </div>
}
function History({counts,products,setMonth}){const months=Object.keys(counts).sort().reverse();return <section className="history">{months.length?months.map(m=>{const c=counts[m];const val=products.filter(p=>p.active!==false).reduce((s,p)=>s+Number(c[p.id]||0)*Number(p.unitPrice||0),0);return <button key={m} onClick={()=>setMonth(m)}><span><b>{m}</b><small>{Object.keys(c).length} quantities saved</small></span><strong>{money(val)}</strong></button>}):<div className="empty">No saved months yet. Enter quantities in Inventory Count to create history.</div>}</section>}
createRoot(document.getElementById('root')).render(<App/>);
