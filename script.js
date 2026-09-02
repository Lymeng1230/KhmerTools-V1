"use strict";
const $ = (s, root = document) => root.querySelector(s);
const $$ = (s, root = document) => [...root.querySelectorAll(s)];

const catalog = [
  ["pdf-to-word","PDF → Word","Extract selectable text into an editable DOCX.","pdf","W","file"],
  ["pdf-to-excel","PDF → Excel","Extract tables with OCR and export a valid spreadsheet.","pdf","X","file"],
  ["pdf-to-images","PDF → Images","Render every page as PNG or JPG in a ZIP.","pdf","▧","file"],
  ["pdf-compress","PDF Compressor","Reduce PDF size using an installed Poppler engine.","pdf","↘","file"],
  ["pdf-merge","Merge PDF","Combine multiple PDFs in your chosen order.","pdf","⊕","file"],
  ["pdf-split","Split PDF","Create one PDF per page in a ZIP archive.","pdf","✂","file"],
  ["pdf-rotate","Rotate PDF","Rotate every page by 90°, 180° or 270°.","pdf","↻","file"],
  ["pdf-protect","Protect PDF","See secure qpdf integration requirements.","pdf","⌾","info"],
  ["image-text","Image → Text","Extract editable text from JPG, PNG or WEBP.","ocr","Aa","section"],
  ["khmer-ocr","Khmer OCR","Recognize Khmer Unicode from images and PDFs.","ocr","ក","section"],
  ["english-ocr","English OCR","Accurate English text recognition.","ocr","EN","section"],
  ["mixed-ocr","Khmer + English OCR","Recognize bilingual documents in one pass.","ocr","កA","section"],
  ["scanned-word","Scanned PDF → Word","Render, OCR and export all pages as DOCX.","ocr","W","file"],
  ["ocr-excel","OCR → Excel","Extract rows and columns into XLSX.","ocr","X","file"],
  ["image-convert","Image Converter","Convert between PNG, JPG and WEBP.","image","◇","file"],
  ["image-compress","Image Compressor","Reduce JPG, PNG or WEBP file size.","image","↘","file"],
  ["image-resize","Image Resizer","Resize images to exact maximum dimensions.","image","↔","file"],
  ["jpg-png","JPG → PNG","Convert a JPG image into lossless PNG.","image","PNG","file"],
  ["png-jpg","PNG → JPG","Convert PNG to a compact JPG.","image","JPG","file"],
  ["image-pdf","Image → PDF","Combine one or more images into a PDF.","image","PDF","file"],
  ["qr","QR Generator","Create a downloadable QR code from text or URL.","image","▦","text"],
  ["ai-chat","AI Chat","Ask general questions in Khmer or English.","ai","AI","section"],
  ["ai-summary","AI Summarizer","Ask AI to summarize text in your language.","ai","Σ","prompt"],
  ["ai-translate","AI Translator","Translate Khmer ↔ English with context.","ai","文","prompt"],
  ["ai-generate","AI Text Generator","Draft articles, CVs and polished copy.","ai","✦","prompt"],
  ["ai-qa","AI Question Answering","Get a clear answer from KhmerTools AI.","ai","?","prompt"],
  ["speech-text","Khmer Speech-to-Text","Requires a configured speech provider; learn more.","ai","◉","info"],
  ["text-speech","Khmer Text-to-Speech","Requires a configured speech provider; learn more.","ai","◖","info"],
  ["video-download","Video Downloader","Download a direct public video URL safely.","media","▶","section"],
  ["image-download","Image Downloader","Download a direct public image URL safely.","media","↓","section"],
  ["url-media","URL Media Downloader","Fetch public direct media without bypassing access controls.","media","↗","section"]
].map(([id,name,description,category,icon,kind])=>({id,name,description,category,icon,kind}));

const toolConfig = {
  "pdf-to-word": { endpoint:"/api/pdf-to-word", field:"pdf", accept:"application/pdf,.pdf", multiple:false },
  "pdf-to-excel": { endpoint:"/api/ocr-to-excel", field:"file", accept:"application/pdf,.pdf", extras:{language:"both"} },
  "pdf-to-images": { endpoint:"/api/pdf-to-images", field:"pdf", accept:"application/pdf,.pdf", fields:[selectField("format","Output format",[["png","PNG"],["jpg","JPG"]])] },
  "pdf-compress": { endpoint:"/api/pdf-compress", field:"pdf", accept:"application/pdf,.pdf" },
  "pdf-merge": { endpoint:"/api/pdf-merge", field:"files", accept:"application/pdf,.pdf", multiple:true },
  "pdf-split": { endpoint:"/api/pdf-split", field:"pdf", accept:"application/pdf,.pdf" },
  "pdf-rotate": { endpoint:"/api/pdf-rotate", field:"pdf", accept:"application/pdf,.pdf", fields:[selectField("angle","Rotation",[["90","90° clockwise"],["180","180°"],["270","270° clockwise"]])] },
  "scanned-word": { endpoint:"/api/pdf-ocr-to-word", field:"pdf", accept:"application/pdf,.pdf", fields:[languageField()] },
  "ocr-excel": { endpoint:"/api/ocr-to-excel", field:"file", accept:"application/pdf,image/jpeg,image/png,image/webp", fields:[languageField()] },
  "image-convert": { endpoint:"/api/image-convert", field:"image", accept:"image/jpeg,image/png,image/webp", fields:[selectField("format","Convert to",[["png","PNG"],["jpeg","JPG"],["webp","WEBP"]])] },
  "image-compress": { endpoint:"/api/image-compress", field:"image", accept:"image/jpeg,image/png,image/webp", fields:[inputField("quality","Quality (20–95)","number","72")] },
  "image-resize": { endpoint:"/api/image-resize", field:"image", accept:"image/jpeg,image/png,image/webp", fields:[inputField("width","Max width (px)","number","1600"),inputField("height","Max height (optional)","number","")] },
  "jpg-png": { endpoint:"/api/image-convert", field:"image", accept:"image/jpeg", extras:{format:"png"} },
  "png-jpg": { endpoint:"/api/image-convert", field:"image", accept:"image/png", extras:{format:"jpeg"} },
  "image-pdf": { endpoint:"/api/image-to-pdf", field:"images", accept:"image/jpeg,image/png,image/webp", multiple:true }
};

function selectField(name,label,options){return {name,label,type:"select",options};}
function inputField(name,label,type,value){return {name,label,type,value};}
function languageField(){return selectField("language","Recognition language",[["both","Khmer + English"],["khmer","Khmer"],["english","English"]]);}
function formatBytes(bytes){if(!bytes)return"0 B";const units=["B","KB","MB"];const i=Math.min(2,Math.floor(Math.log(bytes)/Math.log(1024)));return `${(bytes/1024**i).toFixed(i?1:0)} ${units[i]}`;}
function filenameFrom(response,fallback){const value=response.headers.get("content-disposition")||"";const utf=value.match(/filename\*=UTF-8''([^;]+)/i);if(utf)return decodeURIComponent(utf[1]);const plain=value.match(/filename="?([^";]+)"?/i);return plain?.[1]||fallback;}
function saveBlob(blob,name){const url=URL.createObjectURL(blob);const a=Object.assign(document.createElement("a"),{href:url,download:name});document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);}
function toast(message){const el=$("#toast");el.textContent=message;el.classList.add("show");clearTimeout(toast.timer);toast.timer=setTimeout(()=>el.classList.remove("show"),3500);}
async function apiError(response){const data=await response.json().catch(()=>({}));return new Error(data.error||`Request failed (${response.status}).`);}

const grid=$("#toolsGrid");
function renderTools(){grid.innerHTML=catalog.map(t=>`<article class="tool-card glass" data-id="${t.id}" data-category="${t.category}" data-search="${(t.name+" "+t.description).toLowerCase()}"><span class="tool-icon">${t.icon}</span><span class="cat">${t.category.toUpperCase()}</span><h3>${t.name}</h3><p>${t.description}</p><button type="button">Open Tool →</button></article>`).join("");}
renderTools();
let activeFilter="all";
function filterTools(){const q=$("#toolSearch").value.trim().toLowerCase();let count=0;$$('.tool-card').forEach(card=>{const show=(activeFilter==="all"||card.dataset.category===activeFilter)&&(!q||card.dataset.search.includes(q));card.classList.toggle("hidden",!show);if(show)count++;});$("#emptyState").hidden=count>0;}
$("#toolSearch").addEventListener("input",filterTools);
$("#searchChips").addEventListener("click",e=>{if(e.target.tagName!=="BUTTON")return;$("#toolSearch").value=e.target.textContent;filterTools();$("#tools").scrollIntoView();});
$("#categoryTabs").addEventListener("click",e=>{if(!e.target.dataset.filter)return;activeFilter=e.target.dataset.filter;$$('#categoryTabs button').forEach(b=>b.classList.toggle("active",b===e.target));filterTools();});
document.addEventListener("keydown",e=>{if(e.key==="/"&&!/INPUT|TEXTAREA/.test(e.target.tagName)){e.preventDefault();$("#toolSearch").focus();}});

const dialog=$("#toolDialog"), controls=$("#toolControls"), runTool=$("#runTool"), toolStatus=$("#toolStatus"), progress=$("#toolProgress"), progressBar=$("#toolProgress b"), progressLabel=$("#progressLabel"), progressPercent=$("#progressPercent");let selectedTool=null,selectedFiles=[];
function makeFields(fields=[]){return fields.map(f=>`<label class="dialog-field"><span>${f.label}</span>${f.type==="select"?`<select name="${f.name}">${f.options.map(o=>`<option value="${o[0]}">${o[1]}</option>`).join("")}</select>`:`<input name="${f.name}" type="${f.type}" value="${f.value}" min="1">`}</label>`).join("");}
function openFileTool(tool){const cfg=toolConfig[tool.id];selectedTool={tool,cfg};selectedFiles=[];$("#dialogIcon").textContent=tool.icon;$("#dialogCategory").textContent=tool.category.toUpperCase();$("#dialogTitle").textContent=tool.name;$("#dialogDescription").textContent=tool.description;controls.innerHTML=`<label class="dialog-upload"><input id="dialogFile" type="file" accept="${cfg.accept}" ${cfg.multiple?"multiple":""}><span>↑</span><b>${cfg.multiple?"Choose files":"Drop a file or click to browse"}</b><small id="dialogFileName">Maximum 20 MB per file</small></label>${makeFields(cfg.fields)}`;toolStatus.textContent="";toolStatus.className="status";progress.hidden=true;runTool.hidden=false;runTool.disabled=true;runTool.textContent="Process file";const input=$("#dialogFile");input.addEventListener("change",()=>{selectedFiles=[...input.files];$("#dialogFileName").textContent=selectedFiles.map(f=>`${f.name} (${formatBytes(f.size)})`).join(", ");runTool.disabled=!selectedFiles.length;});dialog.showModal();}
function openTextTool(tool){selectedTool={tool};$("#dialogIcon").textContent=tool.icon;$("#dialogCategory").textContent=tool.category.toUpperCase();$("#dialogTitle").textContent=tool.name;$("#dialogDescription").textContent=tool.description;controls.innerHTML=`<label class="dialog-field"><span>Text or URL</span><input id="qrText" type="text" maxlength="2000" placeholder="Enter text or paste a URL"></label>`;toolStatus.textContent="";runTool.hidden=false;runTool.disabled=false;runTool.textContent="Generate QR code";progress.hidden=true;dialog.showModal();}
function openTool(id){const tool=catalog.find(t=>t.id===id);if(!tool)return;if(tool.kind==="section"){const target=tool.category==="ocr"?"#ocr":tool.category==="media"?"#media":"#ai";$(target).scrollIntoView();if(tool.id==="english-ocr")$("#ocrLanguage").value="english";if(tool.id==="khmer-ocr")$("#ocrLanguage").value="khmer";if(tool.id==="mixed-ocr")$("#ocrLanguage").value="both";return;}if(tool.kind==="prompt"){$("#ai").scrollIntoView();const prompts={"ai-summary":"Summarize the following text clearly:\n\n","ai-translate":"Translate the following text between Khmer and English:\n\n","ai-generate":"Write polished content about:\n\n","ai-qa":"Answer this question clearly:\n\n"};$("#chatInput").value=prompts[id];$("#chatInput").focus();return;}if(tool.kind==="info"){toast(tool.id==="pdf-protect"?"PDF protection needs qpdf on the server; this deployment will not pretend to encrypt files.":"Speech requires a dedicated speech API and is not enabled. No fake transcription or audio will be generated.");return;}if(tool.kind==="text")return openTextTool(tool);openFileTool(tool);}
grid.addEventListener("click",e=>{const card=e.target.closest(".tool-card");if(card)openTool(card.dataset.id);});$$('[data-tool]').forEach(el=>el.addEventListener("click",()=>openTool(el.dataset.tool)));
runTool.addEventListener("click",async()=>{if(!selectedTool)return;const {tool,cfg}=selectedTool;const form=new FormData();let endpoint;if(tool.id==="qr"){endpoint="/api/qr";form.append("text",$("#qrText").value);}else{endpoint=cfg.endpoint;selectedFiles.forEach(f=>form.append(cfg.field,f));Object.entries(cfg.extras||{}).forEach(([k,v])=>form.append(k,v));$$('[name]',controls).forEach(el=>form.append(el.name,el.value));}runTool.disabled=true;progress.hidden=false;toolStatus.textContent="Uploading...";toolStatus.className="status";let pct=18;progressBar.style.width=`${pct}%`;progressPercent.textContent=`${pct}%`;const timer=setInterval(()=>{pct=Math.min(91,pct+Math.ceil(Math.random()*8));progressBar.style.width=`${pct}%`;progressPercent.textContent=`${pct}%`;progressLabel.textContent=pct>75?"Almost finished...":"Processing...";},450);try{const response=await fetch(endpoint,{method:"POST",body:form});if(!response.ok)throw await apiError(response);const blob=await response.blob();saveBlob(blob,filenameFrom(response,`${tool.id}-result`));progressBar.style.width="100%";progressPercent.textContent="100%";progressLabel.textContent="Completed";toolStatus.textContent="Completed — your download has started.";toolStatus.className="status success";}catch(error){toolStatus.textContent=error.message;toolStatus.className="status error";}finally{clearInterval(timer);runTool.disabled=false;}});

const ocrFile=$("#ocrFile"),ocrDrop=$("#ocrDrop"),ocrRun=$("#ocrRun"),ocrText=$("#ocrText"),ocrStatus=$("#ocrStatus");let ocrSelected=null;
function setOcrFile(file){ocrSelected=file;$("#ocrFileLabel").textContent=`${file.name} · ${formatBytes(file.size)}`;ocrRun.disabled=false;}
ocrFile.addEventListener("change",()=>ocrFile.files[0]&&setOcrFile(ocrFile.files[0]));["dragover","dragenter"].forEach(t=>ocrDrop.addEventListener(t,e=>{e.preventDefault();ocrDrop.classList.add("drag");}));["dragleave","drop"].forEach(t=>ocrDrop.addEventListener(t,e=>{e.preventDefault();ocrDrop.classList.remove("drag");}));ocrDrop.addEventListener("drop",e=>e.dataTransfer.files[0]&&setOcrFile(e.dataTransfer.files[0]));ocrText.addEventListener("input",()=>$("#ocrCount").textContent=`${ocrText.value.length.toLocaleString()} characters`);
ocrRun.addEventListener("click",async()=>{if(!ocrSelected)return;const form=new FormData();form.append("file",ocrSelected);form.append("language",$("#ocrLanguage").value);ocrRun.disabled=true;ocrRun.textContent="Processing...";ocrStatus.textContent=ocrSelected.type==="application/pdf"?"Rendering pages, then running OCR...":"Running OCR...";ocrStatus.className="status";try{const response=await fetch("/api/ocr",{method:"POST",body:form});if(!response.ok)throw await apiError(response);const data=await response.json();ocrText.value=data.text||"";ocrText.dispatchEvent(new Event("input"));ocrStatus.textContent=`Completed${data.pages>1?` · ${data.pages} pages`:""}. You can edit or export the text.`;ocrStatus.className="status success";}catch(error){ocrStatus.textContent=error.message;ocrStatus.className="status error";}finally{ocrRun.disabled=false;ocrRun.textContent="Extract Text";}});
$(".editor-actions").addEventListener("click",async e=>{const action=e.target.dataset.ocrAction;if(!action)return;const text=ocrText.value;if(action==="clear"){ocrText.value="";ocrText.dispatchEvent(new Event("input"));return;}if(!text.trim())return toast("There is no OCR text to export.");if(action==="copy"){await navigator.clipboard.writeText(text);return toast("Text copied.");}if(action==="txt")return saveBlob(new Blob([text],{type:"text/plain;charset=utf-8"}),"khmertools-ocr.txt");const endpoint=action==="word"?"/api/text-to-word":"/api/text-to-excel";const response=await fetch(endpoint,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({text})});if(!response.ok)return toast((await apiError(response)).message);saveBlob(await response.blob(),filenameFrom(response,`khmertools.${action==="word"?"docx":"xlsx"}`));});

const chatHistory=[];const messages=$("#chatMessages"),chatInput=$("#chatInput"),sendChat=$("#sendChat");
function addMessage(role,text,typing=false){const el=document.createElement("div");el.className=`message ${role}`;el.innerHTML=`<div class="avatar">${role==="user"?"U":"K"}</div><div><p></p>${role==="assistant"&&!typing?'<button class="copy-message">Copy</button>':""}</div>`;$("p",el).textContent=text;messages.append(el);messages.scrollTop=messages.scrollHeight;return el;}
async function sendMessage(){const text=chatInput.value.trim();if(!text)return;chatInput.value="";addMessage("user",text);chatHistory.push({role:"user",content:text});sendChat.disabled=true;const typing=addMessage("assistant","Thinking…",true);try{const response=await fetch("/api/ai/chat",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({messages:chatHistory})});if(!response.ok)throw await apiError(response);const data=await response.json();typing.remove();addMessage("assistant",data.text);chatHistory.push({role:"assistant",content:data.text});}catch(error){typing.remove();addMessage("assistant",error.message);}finally{sendChat.disabled=false;chatInput.focus();}}
sendChat.addEventListener("click",sendMessage);chatInput.addEventListener("keydown",e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendMessage();}});$(".prompts").addEventListener("click",e=>{if(e.target.tagName==="BUTTON"){chatInput.value=e.target.textContent;chatInput.focus();}});messages.addEventListener("click",async e=>{if(e.target.classList.contains("copy-message")){await navigator.clipboard.writeText($("p",e.target.parentElement).textContent);toast("Response copied.");}});$("#clearChat").addEventListener("click",()=>{chatHistory.length=0;messages.innerHTML='<div class="message assistant"><div class="avatar">K</div><div><p>Chat cleared. How can I help?</p><button class="copy-message">Copy</button></div></div>';});

$("#downloadMedia").addEventListener("click",async()=>{const url=$("#mediaUrl").value.trim(),status=$("#mediaStatus"),button=$("#downloadMedia");if(!url)return toast("Paste a direct public media URL first.");button.disabled=true;button.textContent="Fetching...";status.textContent="Validating the public URL and downloading securely...";status.className="status";try{const response=await fetch("/api/media/download",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({url})});if(!response.ok)throw await apiError(response);saveBlob(await response.blob(),filenameFrom(response,"media-download"));status.textContent="Completed — download started.";status.className="status success";}catch(error){status.textContent=error.message;status.className="status error";}finally{button.disabled=false;button.textContent="Download";}});

$$('[data-notice]').forEach(el=>el.addEventListener("click",()=>toast(el.dataset.notice)));const menu=$("#menuButton"),links=$("#navLinks");menu.addEventListener("click",()=>{const open=links.classList.toggle("open");menu.setAttribute("aria-expanded",open);menu.textContent=open?"×":"☰";});links.addEventListener("click",()=>{links.classList.remove("open");menu.setAttribute("aria-expanded","false");menu.textContent="☰";});
