// Theme switching functionality
document.addEventListener('DOMContentLoaded', function() {
  const themeToggle = document.getElementById('theme-toggle');
  const currentTheme = localStorage.getItem('theme') || 'light';
  
  // Apply saved theme
  if (currentTheme === 'dark') {
    document.documentElement.setAttribute('data-theme', 'dark');
    themeToggle.textContent = '☀️';
  }
  
  // Toggle theme
  themeToggle.addEventListener('click', function() {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    if (currentTheme === 'dark') {
      document.documentElement.setAttribute('data-theme', 'light');
      themeToggle.textContent = '🌙';
      localStorage.setItem('theme', 'light');
    } else {
      document.documentElement.setAttribute('data-theme', 'dark');
      themeToggle.textContent = '☀️';
      localStorage.setItem('theme', 'dark');
    }
  });
});

// Helpers
function qs(s,root=document){return root.querySelector(s)}
function qsa(s,root=document){return [...root.querySelectorAll(s)]}
function toast(msg){const t=qs('#toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),1800)}
function copyToClipboard(text){navigator.clipboard?.writeText(text).then(()=>toast('Скопировано в буфер обмена')).catch(()=>{const ta=document.createElement('textarea');ta.value=text;document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);toast('Скопировано')})}

// Tabs with smooth transition
document.addEventListener('DOMContentLoaded', function() {
  const tabs = qsa('.tab');
  const tabContainer = qs('.tabs');
  
  // Add indicator element
  const indicator = document.createElement('div');
  indicator.style.position = 'absolute';
  indicator.style.bottom = '0';
  indicator.style.height = '3px';
  indicator.style.background = getComputedStyle(document.documentElement).getPropertyValue('--primary');
  indicator.style.borderRadius = '3px';
  indicator.style.transition = 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)';
  indicator.style.zIndex = '2';
  tabContainer.appendChild(indicator);
  
  // Function to update indicator position
  function updateIndicator(activeTab) {
    const tabRect = activeTab.getBoundingClientRect();
    const containerRect = tabContainer.getBoundingClientRect();
    const offset = tabRect.left - containerRect.left;
    
    indicator.style.width = `${tabRect.width}px`;
    indicator.style.transform = `translateX(${offset}px)`;
  }
  
  // Initialize indicator position
  const activeTab = qs('.tab.active');
  if (activeTab) {
    // Wait for layout to be ready
    setTimeout(() => updateIndicator(activeTab), 10);
  }
  
  tabs.forEach(tab => tab.addEventListener('click', () => {
    // Special handling for the "Исход" tab
    if (tab.dataset.target === '#tab5') {
      toast('Скоро!');
      return;
    }
    
    // Remove active class from all tabs
    tabs.forEach(t => t.classList.remove('active'));
    
    // Add active class to clicked tab
    tab.classList.add('active');
    
    // Update indicator position
    updateIndicator(tab);
    
    // Show/hide tab panels
    const target = tab.dataset.target;
    qsa('.tab-panel').forEach(p => {
      if ('#' + p.id === target) {
        p.classList.add('show');
        p.setAttribute('aria-hidden', 'false');
      } else {
        p.classList.remove('show');
        p.setAttribute('aria-hidden', 'true');
      }
    });
  }));
});

// TABS 1 logic
document.addEventListener('DOMContentLoaded', function() {
  const formFailure = qs('#form-failure');
  const addToReportBtn = qs('#addToReport');
  const reportList = qs('#reportList');
  const reportCount = qs('#reportCount');
  const downloadExcelBtn = qs('#downloadExcel');
  const clearReportBtn = qs('#clearReport');

  // Load report data from localStorage if available
  let reportData = JSON.parse(localStorage.getItem('tnlReportData')) || [];

  // Update the report display on page load
  function updateReportDisplay() {
    reportList.innerHTML = '';
    reportData.forEach(item => {
      const div = document.createElement('div');
      div.className = 'report-item';
      div.textContent = `${item.clientName || ''} • ${item.clientPhone || ''} • ${item.address || item.article || ''} • ${item.agentName || ''}`;
      reportList.appendChild(div);
    });
    reportCount.textContent = `${reportData.length} записей`;
  }

  // Initialize report display
  updateReportDisplay();

  // show/hide source-dependent fields
  const sourceSelect = qs('#sourceSelect');
  const bannerNameField = qs('#bannerNameField');
  const cityField = qs('#cityField');
  const streetField = qs('#streetField');
  const avitoLinkField = qs('#avitoLinkField');
  
  // Initialize with default values
  sourceSelect.value = 'ГородУлица';
  qs('#requestType').value = 'Покупка/аренда конкретного ОН';
  
  // Show default fields
  cityField.classList.remove('hidden');
  qs('#cityField label').textContent = 'Город и улица';
  qs('#cityField input').placeholder = 'Город и улица';
  qs('#cityField input').name = 'cityStreet';
  
  sourceSelect.addEventListener('change',()=>{
    const v=sourceSelect.value;
    // Hide all fields first
    bannerNameField.classList.add('hidden');
    cityField.classList.add('hidden');
    streetField.classList.add('hidden');
    avitoLinkField.classList.add('hidden');
    
    // Reset field names and labels to default
    qs('#cityField label').textContent = 'Город';
    qs('#cityField input').placeholder = 'Напр.: Пенза, Москва';
    qs('#cityField input').name = 'city';
    qs('#streetField label').textContent = 'Улица';
    qs('#streetField input').placeholder = 'Улица';
    qs('#streetField input').name = 'street';
    
    // Show fields based on selection
    if(v==='Баннер' || v==='Пенза баннер'){
      // Only banner name field
      bannerNameField.classList.remove('hidden');
    } else if(v==='ГородУлица'){
      // Only one field for city and street combined
      cityField.classList.remove('hidden');
      qs('#cityField label').textContent = 'Город и улица';
      qs('#cityField input').placeholder = 'Город и улица';
      qs('#cityField input').name = 'cityStreet';
    } else if(v==='Звонок с карт' || v==='Др.источники'){
      // Two fields: office and street
      cityField.classList.remove('hidden');
      streetField.classList.remove('hidden');
      qs('#cityField label').textContent = 'Офис';
      qs('#cityField input').placeholder = 'Офис';
      qs('#cityField input').name = 'office';
      qs('#streetField label').textContent = 'Улица (если не знают офис)';
      qs('#streetField input').placeholder = 'Улица';
      qs('#streetField input').name = 'streetNearOffice';
    } else if(v==='Звонок с авито'){
      // Show avito link field
      avitoLinkField.classList.remove('hidden');
    }
  })

  // request type dynamic UI
  const requestType = qs('#requestType');
  const typeDetails = qs('#typeDetails');
  // Get references to the specific fields that should only show for "Покупка/аренда конкретного ОН"
  // Select fields by their position in the form to avoid any issues with selectors
  const formFields = qsa('#form-failure .field');
  const addressField = formFields[2]; // 3rd field (0-indexed)
  const articleField = formFields[3]; // 4th field
  const priceField = formFields[4];   // 5th field
  // Get reference to the role selector field (6th field)
  const roleField = formFields[5];

  // Initialize with default state - show specific ON fields
  function updateSpecificOnFieldsVisibility() {
    const showFields = requestType.value === 'Покупка/аренда конкретного ОН';
    const showRoleField = requestType.value === 'Покупка/аренда конкретного ОН';
    
    if (showFields) {
      addressField.classList.remove('hidden');
      articleField.classList.remove('hidden');
      priceField.classList.remove('hidden');
    } else {
      addressField.classList.add('hidden');
      articleField.classList.add('hidden');
      priceField.classList.add('hidden');
    }
    
    // Show/hide role field based on request type
    if (showRoleField) {
      roleField.classList.remove('hidden');
    } else {
      roleField.classList.add('hidden');
    }
  }

  // Set initial visibility
  updateSpecificOnFieldsVisibility();

  function clearElement(el){el.innerHTML='';}
  requestType.addEventListener('change', ()=>{
    // Update visibility of specific ON fields
    updateSpecificOnFieldsVisibility();
    
    const v = requestType.value;
    
    clearElement(typeDetails);
    if(v==='Покупка.Подбор' || v==='Аренда.Подбор'){
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-4"><label>Район</label><input name="region" placeholder="Район"></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-2"><label>Комнат</label><input name="rooms" placeholder="Комнат"></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-2"><label>Площадь</label><input name="area" placeholder="м²"></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-2"><label>Этаж</label><input name="floor" placeholder="Этаж"></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-4"><label>Стоимость</label><div style="display:flex;gap:8px"><input name="targetPrice" placeholder="3 500" style="flex:1"><select name="targetPriceUnit" style="width:110px"><option value="">— ед. —</option><option value="тыс">тыс</option><option value="млн">млн</option></select></div></div>`);
      if(v==='Покупка.Подбор'){
        typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-4"><label>Первичка / Вторичка</label><select name="newOrResale"><option value="">— не выбрано —</option><option value="Первичка">Первичка</option><option value="Вторичка">Вторичка</option></select></div>`);
      }
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-12"><label>Доп. пожелания клиента</label><input name="extraWishes" placeholder="Доп. пожелания"></div>`);
    } else if(v==='Продажа/сдача.Подбор'){
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-6"><label>Адрес объекта</label><input name="saleAddress" placeholder="Адрес"></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-2"><label>Комнат</label><input name="saleRooms" placeholder="Комнат"></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-2"><label>Площадь</label><input name="saleArea" placeholder="м²"></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-2"><label>Этаж</label><input name="saleFloor" placeholder="Этаж"></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-4"><label>За сколько продать/сдать</label><div style="display:flex;gap:8px"><input name="salePrice" placeholder="Сумма или оставьте ппусто" style="flex:1"><select name="salePriceUnit" style="width:110px"><option value="">— ед. —</option><option value="тыс">тыс</option><option value="млн">млн</option></select></div></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-4"><label><input type="checkbox" name="needValuation"> Требуется оценка</label></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-12"><label>Доп. примечания</label><input name="extraNotes" placeholder="Доп. примечания"></div>`);
    } else if(v==='Ипотека'){
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-6"><label>Офис где требуется услуга</label><input name="mortOffice" placeholder="Офис"></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-6"><label>Улица (если не знают офис)</label><input name="mortStreet" placeholder="Улица"></div>`);
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-12"><label>Доп. примечания</label><input name="mortNotes" placeholder="Доп. примечания"></div>`);
    } else if(v==='Нестандартное'){
      // Add the custom text field for "Нестандартное"
      typeDetails.insertAdjacentHTML('beforeend', `<div class="field col-12"><label>Нестандартное описание</label><textarea name="customDescription" rows="4" placeholder="Опишите ситуацию подробно"></textarea></div>`);
    }
    // For "Покупка/аренда конкретного ОН" - no additional fields needed in typeDetails
  });

  // Validation: address or article required (only for specific ON request type)
  function validateAddressOrArticle(form){
    const requestTypeValue = form.requestType.value;
    // Skip validation for "Нестандартное" type
    if(requestTypeValue === 'Нестандартное') {
      // For "Нестандартное" only require client name and phone
      return true;
    }
    
    // Only validate address/article for "Покупка/аренда конкретного ОН"
    if(requestTypeValue === 'Покупка/аренда конкретного ОН') {
      const address = form.address.value.trim();
      const article = form.article.value.trim();
      if(!address && !article){
        toast('Заполните адрес или артикул');
        return false;
      }
    }
    return true;
  }

  // Custom validation for "Нестандартное" option
  function validateCustomRequest(form) {
    const requestTypeValue = form.requestType.value;
    if(requestTypeValue === 'Нестандартное') {
      const customDescription = form.customDescription?.value.trim();
      if(!customDescription) {
        toast('Заполните нестандартное описание');
        return false;
      }
    }
    return true;
  }

  // collect failure payload
  function getFailurePayload(){
    const f = formFailure;
    const data = {};
    new FormData(f).forEach((v,k)=>data[k]=v);
    return data;
  }

  function failureTemplate(p){
    // Special handling for "Нестандартное" type
    if(p.requestType === 'Нестандартное') {
      return `Сбой ТНЛ
Клиент: ${p.clientName || '—'}, ${p.clientPhone || '—'}
Нестандартное описание: ${p.customDescription || '—'}
Агент: ${p.agentName || '—'}
Источник: ${p.source || '—'}`.trim();
    }
    
    const addrOrArt = p.address ? `Адрес: ${p.address}` : (p.article ? `Артикул: ${p.article}` : '');
    const price = p.price ? `${p.price} ${p.priceUnit || ''}` : '—';
    return `Сбой ТНЛ
Клиент: ${p.clientName || '—'}, ${p.clientPhone || '—'}
${addrOrArt}
Стоимость: ${price}
Агент: ${p.agentName || '—'}
Источник: ${p.source || '—'}
Звонит: ${p.role || '—'}`.trim();
  }

  // Add to report: push, clear fields used
  addToReportBtn.addEventListener('click', ()=>{
    if(!formFailure.checkValidity()){ formFailure.reportValidity(); return; }
    if(!validateAddressOrArticle(formFailure)) return;
    if(!validateCustomRequest(formFailure)) return;
    const payload = getFailurePayload();
    reportData.push(payload);
    
    // Save to localStorage
    localStorage.setItem('tnlReportData', JSON.stringify(reportData));
    
    const div = document.createElement('div');
    div.className='report-item';
    // Update the display to show custom description for "Нестандартное" requests
    const displayText = payload.requestType === 'Нестандартное' ? 
      (payload.customDescription || '') : 
      (payload.address || payload.article || '');
    div.textContent = `${payload.clientName || ''} • ${payload.clientPhone || ''} • ${displayText} • ${payload.agentName || ''}`;
    reportList.appendChild(div);
    reportCount.textContent = `${reportData.length} записей`;
    toast('Запись добавлена в отчёт');
    // clear *all* fields in the form
    formFailure.reset();
    // hide conditional fields and clear dynamic area
    qs('#bannerNameField').classList.add('hidden'); 
    qs('#cityField').classList.add('hidden'); 
    qs('#streetField').classList.add('hidden'); 
    qs('#avitoLinkField').classList.add('hidden');
    typeDetails.innerHTML='';
    
    // Reset to default values
    sourceSelect.value = 'ГородУлица';
    requestType.value = 'Покупка/аренда конкретного ОН';
    qs('#cityField').classList.remove('hidden');
    qs('#cityField label').textContent = 'Город и улица';
    qs('#cityField input').placeholder = 'Город и улица';
    qs('#cityField input').name = 'cityStreet';
    
    // Show the specific ON fields again as default
    updateSpecificOnFieldsVisibility();
  });

  // Excel export using SheetJS
  downloadExcelBtn.addEventListener('click', ()=>{
    if(reportData.length===0){ toast('Нет записей для выгрузки'); return; }
    
    // Get current date for filename
    const now = new Date();
    const dateStr = `${String(now.getDate()).padStart(2, '0')}.${String(now.getMonth() + 1).padStart(2, '0')}.${now.getFullYear()}`;
    
    const rows = reportData.map(r=>({
      'Имя клиента': r.clientName||'',
      'Номер клиента': r.clientPhone||'',
      'Адрес': r.address||'',
      'Артикул': r.article||'',
      'Стоимость': r.price ? (r.price + ' ' + (r.priceUnit||'')) : '',
      'ФИО агента': r.agentName||'',
      'Источник звонка': r.source||'',
      'Звонит': r.role||'',
      'Тип заявки': r.requestType||'',
      // Fields for Покупка.Подбор и Аренда.Подбор
      'Район': r.region||'',
      'Комнат': r.rooms||'',
      'Площадь': r.area||'',
      'Этаж': r.floor||'',
      'Стоимость подбора': r.targetPrice ? (r.targetPrice + ' ' + (r.targetPriceUnit||'')) : '',
      'Первичка/Вторичка': r.newOrResale||'',
      'Доп. пожелания': r.extraWishes||'',
      // Fields for Продажа/сдача.Подбор
      'Адрес объекта (продажа)': r.saleAddress||'',
      'Комнат (продажа)': r.saleRooms||'',
      'Площадь (продажа)': r.saleArea||'',
      'Этаж (продажа)': r.saleFloor||'',
      'Цена продажи': r.salePrice ? (r.salePrice + ' ' + (r.salePriceUnit||'')) : '',
      'Требуется оценка': r.needValuation ? 'Да' : 'Нет',
      'Доп. примечания (продажа)': r.extraNotes||'',
      // Fields for Ипотека
      'Офис (ипотека)': r.mortOffice||'',
      'Улица (ипотека)': r.mortStreet||'',
      'Доп. примечания (ипотека)': r.mortNotes||'',
      // Fields for Нестандартное
      'Нестандартное описание': r.customDescription||'',
      // Fields for call sources
      'Название баннера': r.bannerName||'',
      'Город и улица': r.cityStreet||'',
      'Офис': r.office||'',
      'Улица (звонок)': r.streetNearOffice||'',
      'Ссылка с авито': r.avitoLink||'',
      'Дата и время': new Date().toLocaleString('ru-RU', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      })
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    
    // Add auto-highlighting to header row
    const range = XLSX.utils.decode_range(ws['!ref']);
    for(let C = range.s.C; C <= range.e.C; ++C) {
      const address = XLSX.utils.encode_cell({r: 0, c: C});
      if(!ws[address]) continue;
      ws[address].s = {
        font: { bold: true },
        fill: { fgColor: { rgb: "D3D3D3" } } // Light gray background
      };
    }
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "report");
    const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
    const blob = new Blob([wbout], {type:'application/octet-stream'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href=url; a.download=`Отчет после сбоя ТНЛ за ${dateStr}.xlsx`; a.click();
    URL.revokeObjectURL(url);
    
    // Clear report data after successful download for security
    reportData = [];
    localStorage.removeItem('tnlReportData');
    reportList.innerHTML = '';
    reportCount.textContent = '0 записей';
    toast('Excel-файл скачан. Отчёт очищен для безопасности.');
  });

  clearReportBtn.addEventListener('click', ()=>{
    reportData = [];
    localStorage.removeItem('tnlReportData');
    reportList.innerHTML = '';
    reportCount.textContent = '0 записей';
    toast('Отчёт очищен');
  });
});

// TAB 2 vacancy
document.addEventListener('DOMContentLoaded', function() {
  qs('#copyVacancy').addEventListener('click', ()=>{
    const form = qs('#form-vacancy');
    if(!form.checkValidity()){ form.reportValidity(); return; }
    const name = form.name.value.trim(); const phone = form.phone.value.trim();
    const text = `Контактный центр Самолет Плюс
Получили отклик по вакансии
${phone} - ${name}`;
    copyToClipboard(text);
    form.reset();
  });
});

// TAB 3 complaint/remove
document.addEventListener('DOMContentLoaded', function() {
  const formComplaint = qs('#form-complaint');
  const complaintField = qs('#complaintTextField');
  const addressOnField = qs('#addressOnField');
  qsa('input[name="mode"]', formComplaint).forEach(r=>r.addEventListener('change', ()=>{
    const mode = qs('input[name="mode"]:checked', formComplaint).value;
    if(mode==='complaint'){ complaintField.classList.remove('hidden'); addressOnField.classList.add('hidden'); }
    else{ addressOnField.classList.remove('hidden'); complaintField.classList.add('hidden'); }
  }));

  qs('#copyComplaint').addEventListener('click', ()=>{
    const f = formComplaint;
    const mode = qs('input[name="mode"]:checked', f).value;
    const cname = f.cname.value.trim(); const cphone = f.cphone.value.trim();
    const aname = f.aname.value.trim(); const aoffice = f.aoffice.value.trim();
    let text='';
    if(mode==='complaint'){
      const ctext = f.ctext.value.trim();
      // Copy all fields except the topic selection (complaint/remove ON)
      text = `Имя клиента: ${cname || '—'}
Телефон клиента: ${cphone || '—'}
ФИО агента: ${aname || '—'}
Офис агента: ${aoffice || '—'}
Суть жалобы: ${ctext || ''}`.trim();
    } else {
      const onaddr = f.onaddr.value.trim();
      text = `Имя клиента: ${cname || '—'}
Телефон клиента: ${cphone || '—'}
ФИО агента: ${aname || '—'}
Офис агента: ${aoffice || '—'}
Суть жалобы: Клиент просит снять ОН по адресу ${onaddr || '—'} с продажи, т.к клиент не давал(а) согласия на публикацию объявления. Просьба снять ОН с всех площадок`;
    }
    copyToClipboard(text);
    // clear form after copy
    formComplaint.reset();
    complaintField.classList.remove('hidden'); addressOnField.classList.add('hidden');
  });
});

// TAB 4 not found
document.addEventListener('DOMContentLoaded', function() {
  qs('#copyNotFound').addEventListener('click', ()=>{
    const f = qs('#form-notfound');
    if(!f.checkValidity()){ f.reportValidity(); return; }
    const addr = f.addr.value.trim(); const cost = f.cost.value.trim(); const unit = f.costUnit ? f.costUnit.value.trim() : f.costUnit; const link = f.link.value.trim();
    // Get the role (client or agent) - default to "Клиент" if nothing selected
    const role = qs('input[name="role"]:checked', f)?.value || 'Клиент';
    
    // Modify the main text based on who is calling
    const roleText = role === 'Агент' ? 'Агента интересует ОН по адресу' : 'Клиента интересует ОН по адресу';
    
    let text = `ОН не найден в ТНЛ
${roleText} ${addr || '—'} за ${cost ? (cost + ' ' + (unit||'')) : '—'}`;
    if(link) text += `
Ссылка: ${link}`;
    copyToClipboard(text);
    // clear fields
    f.reset();
    // Reset to default "Клиент" selection
    const clientRadio = qs('input[name="role"][value="Клиент"]', f);
    if(clientRadio) clientRadio.checked = true;
  });
});

// Quick action buttons
document.addEventListener('DOMContentLoaded', function() {
  // Повторное обращение
  qs('#repeat-call').addEventListener('click', () => {
    copyToClipboard('повторное обращение');
  });
  
  // Агент не связался
  qs('#agent-not-contacted').addEventListener('click', () => {
    copyToClipboard('агент не связался с клиентом, просьба связаться');
  });
});
