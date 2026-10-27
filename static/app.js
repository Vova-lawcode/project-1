let caseData={}; let history=[];
const messages=document.getElementById('messages');
const form=document.getElementById('chatForm'); const userText=document.getElementById('userText');
const generateBtn=document.getElementById('generateBtn'); const downloadLink=document.getElementById('downloadLink');
const caseView=document.getElementById('caseView'); const progressBar=document.getElementById('progressBar'); const statusBox=document.getElementById('statusBox');
const progressCount=document.getElementById('progressCount'); const summaryCount=document.getElementById('summaryCount');
const demoBtn=document.getElementById('demoBtn');

const keys=['court_number','court_district','court_address','applicant_name','birth_date','passport','passport_issued','applicant_address','creditor_name','case_number','order_date','order_subject','awareness_date','awareness_source','missed_reason','disagreement_basis','effective_start_date','fssp_office'];
const fieldLabels={
  court_number:'Номер судебного участка',court_district:'Официальное название района и региона суда',court_address:'Адрес суда',
  applicant_name:'ФИО заявителя',birth_date:'Дата рождения',passport:'Паспорт',passport_issued:'Кем и когда выдан',
  inn:'ИНН',snils:'СНИЛС',applicant_address:'Адрес заявителя',phone:'Телефон',email:'Электронная почта',
  creditor_name:'Взыскатель',creditor_inn:'ИНН взыскателя',creditor_address:'Адрес взыскателя',case_number:'Номер дела',
  order_date:'Дата судебного приказа',order_subject:'Предмет взыскания',ip_date:'Дата исполнительного производства',
  ip_number:'Номер исполнительного производства',awareness_date:'Когда узнали о приказе',awareness_source:'Откуда узнали',
  missed_reason:'Причина пропуска срока',supporting_documents:'Подтверждающие документы',disagreement_basis:'Причина несогласия',
  effective_start_date:'Дата начала срока',fssp_office:'Отдел ФССП',correspondence_address:'Адрес для корреспонденции',composition_date:'Дата составления'
};
function add(role,text){
  const isUser=role==='user'; const row=document.createElement('div'); row.className='message-row '+(isUser?'user-row':'assistant-row');
  const avatar=document.createElement('span'); avatar.className='message-avatar'; avatar.setAttribute('aria-hidden','true'); avatar.textContent=isUser?'Вы':'И';
  const content=document.createElement('div'); content.className='message-content';
  const author=document.createElement('span'); author.className='message-author';
  const name=document.createElement('span'); name.textContent=isUser?'Вы':'Иск-Бот';
  const time=document.createElement('time'); time.textContent=new Date().toLocaleTimeString('ru-RU',{hour:'2-digit',minute:'2-digit'});
  author.append(name,time);
  const bubble=document.createElement('div'); bubble.className='msg '+(isUser?'user':'bot'); bubble.textContent=text;
  content.append(author,bubble); row.append(avatar,content); messages.appendChild(row); messages.scrollTop=messages.scrollHeight;
}
function updateView(){
  const done=keys.filter(k=>caseData[k]).length; progressBar.style.width=Math.round(done/keys.length*100)+'%';
  progressBar.parentElement.setAttribute('aria-valuenow',String(done));
  progressCount.textContent=`${done} из ${keys.length}`;
  statusBox.textContent=done===0?'Начните с описания ситуации — помощник подскажет, что понадобится.':`Заполнено ${done} из ${keys.length} обязательных сведений.`;
  summaryCount.textContent=done===0?'Пока пусто':`${done} из ${keys.length} заполнено`;
  caseView.replaceChildren();
  const entries=Object.entries(fieldLabels).filter(([key])=>{
    const value=caseData[key]; return Array.isArray(value)?value.length>0:Boolean(value);
  });
  if(!entries.length){const empty=document.createElement('p');empty.className='empty-case';empty.textContent='Здесь появятся сведения, которые мы соберём в диалоге.';caseView.appendChild(empty);return;}
  for(const [key,label] of entries){
    const row=document.createElement('div');row.className='field-row';
    const title=document.createElement('span');title.textContent=label;
    const value=document.createElement('strong');value.textContent=Array.isArray(caseData[key])?caseData[key].join(', '):caseData[key];
    row.append(title,value);caseView.appendChild(row);
  }
}
async function send(text){
  add('user',text); history.push({role:'user',content:text}); userText.value='';
  history=history.slice(-12);
  try{
    const r=await fetch('/api/turn',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({history,case:caseData,user_text:text})});
    const j=await r.json(); if(!r.ok){add('bot','Ошибка: '+(j.detail||'не удалось обработать ответ'));return;}
    caseData=j.case; history.push({role:'assistant',content:j.turn.assistant_message}); history=history.slice(-12); add('bot',j.turn.assistant_message); updateView();
    generateBtn.disabled=!j.turn.ready_to_generate;
  }catch(error){add('bot','Не удалось связаться с сервером. Проверьте подключение и попробуйте ещё раз.');}
}
form.addEventListener('submit',e=>{e.preventDefault(); if(userText.value.trim()) send(userText.value.trim())});
generateBtn.addEventListener('click',async()=>{try{const r=await fetch('/api/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(caseData)});const j=await r.json(); if(!r.ok){add('bot','Не хватает данных: '+j.detail);return;} downloadLink.href=j.download_url; downloadLink.classList.remove('hidden'); add('bot','Готово. Заявление сформировано.');}catch(error){add('bot','Не удалось сформировать документ: сервер недоступен.');}});
demoBtn.addEventListener('click',async()=>{try{const r=await fetch('/api/demo',{method:'POST'});const j=await r.json();if(!r.ok){add('bot','Не удалось открыть пример. Попробуйте ещё раз.');return;}downloadLink.href=j.download_url;downloadLink.classList.remove('hidden');caseData=j.case;updateView();generateBtn.disabled=keys.some(k=>!caseData[k]);add('bot','Открыл пример заполнения. Его можно использовать, чтобы посмотреть, как выглядит готовое заявление.');}catch(error){add('bot','Не удалось открыть пример: сервер недоступен.');}});
updateView();
