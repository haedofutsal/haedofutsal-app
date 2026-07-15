const supabase = require('./supabase');

function getCategoriaPadre(cId) {
    if (!cId) return '';
    cId = cId.toLowerCase();
    if (cId.includes('edefi') && cId.includes('may')) return 'edefi-mayores';
    if (cId.includes('edefi') && cId.includes('baby')) return 'edefi-baby';
    if (cId.includes('bafi') && cId.includes('fem')) return 'bafi-femenino';
    if (cId.includes('bafi') && cId.includes('masc')) return 'bafi-masculino';
    if (cId.includes('futsala') && cId.includes('prom')) return 'futsala-promocionales';
    if (cId.includes('futsala') && cId.includes('masc')) return 'futsala-masculino';
    return '';
}

async function verificarYGenerarCuotasMensualesNativo() {
  try {
    const now = new Date();
    const offset = -3;
    const argDate = new Date(now.getTime() + offset * 3600 * 1000);
    const currentYear = argDate.getUTCFullYear();
    const currentMonthNum = argDate.getUTCMonth() + 1;
    const currentDay = argDate.getUTCDate();
    const currentMonthStr = currentYear + '-' + String(currentMonthNum).padStart(2, '0');
    
    const startYear = 2026;
    const startMonth = 7;
    const monthsToGenerate = [];
    let y = startYear; let m = startMonth;
    while (true) {
      if (y > currentYear || (y === currentYear && m > currentMonthNum)) break;
      monthsToGenerate.push(y + '-' + String(m).padStart(2, '0'));
      m++; if (m > 12) { m = 1; y++; }
    }
    
    const [{ data: categorias }, { data: usuarios }, { data: pagos }] = await Promise.all([
      supabase.from('categorias').select('*'),
      supabase.from('usuarios').select('email, role, category'),
      supabase.from('pagos').select('id, payment_id, email, month, status, amount')
    ]);
    
    if (!categorias || !usuarios || !pagos) return;

    const categoryFees = {};
    const parentCategoryFees = {};
    for (const cat of categorias) {
      const id = (cat.category_id || '').toLowerCase();
      const name = (cat.name || '').toLowerCase();
      const fee = parseFloat(cat.monthly_fee) || parseFloat(cat.price_monthly) || 0;
      if (id) categoryFees[id] = fee;
      if (name) categoryFees[name] = fee;
      const parentId = getCategoriaPadre(id);
      if (parentId && fee) {
        if (!parentCategoryFees[parentId] || fee > parentCategoryFees[parentId]) {
          parentCategoryFees[parentId] = fee;
        }
      }
    }
    
    function getSocioFee(categoryStr) {
      if (!categoryStr) return 0;
      const cats = categoryStr.split(',').map(c => c.trim().toLowerCase());
      let maxFee = 0;
      cats.forEach(c => {
        let fee = categoryFees[c];
        if (fee === undefined) {
          const parentId = getCategoriaPadre(c).toLowerCase();
          fee = parentCategoryFees[parentId];
        }
        if (fee === undefined) {
          const clean = c.replace(/[^a-z0-9+]/g, '');
          for (const key in categoryFees) {
            const cleanKey = key.replace(/[^a-z0-9+]/g, '');
            if (clean === cleanKey || clean.includes(cleanKey) || cleanKey.includes(clean)) {
              fee = categoryFees[key]; break;
            }
          }
        }
        if (fee && fee > maxFee) maxFee = fee;
      });
      return maxFee;
    }
    
    const activeSocios = usuarios.filter(u => u.role === 'Deportista' && u.email);
    const existingPayments = {};
    pagos.forEach(p => { existingPayments[(p.email||'').toLowerCase() + '|' + p.month] = p; });
    
    for (const socio of activeSocios) {
      const email = socio.email.toLowerCase();
      for (const month of monthsToGenerate) {
        const key = email + '|' + month;
        const existing = existingPayments[key];
        
        let targetStatus = 'Pendiente';
        if (month < currentMonthStr) targetStatus = 'Deuda';
        else if (month === currentMonthStr) targetStatus = (currentDay > 10) ? 'Deuda' : 'Pendiente';
        
        const correctAmount = getSocioFee(socio.category);
        
        if (!existing) {
          const payId = 'PAG-' + Date.now().toString().slice(-4) + '-' + Math.floor(Math.random() * 100);
          await supabase.from('pagos').insert([{
            payment_id: payId, email: socio.email, month: month, amount: correctAmount, status: targetStatus, mp_link: '', collected_by: '', collected_at: ''
          }]);
          existingPayments[key] = { status: targetStatus, amount: correctAmount };
        } else {
          if (existing.status !== 'Pagado' && existing.amount !== correctAmount && correctAmount > 0) {
            await supabase.from('pagos').update({ amount: correctAmount }).eq('id', existing.id);
            existing.amount = correctAmount;
          }
          if (existing.status === 'Pendiente' && targetStatus === 'Deuda') {
            await supabase.from('pagos').update({ status: 'Deuda' }).eq('id', existing.id);
            existing.status = 'Deuda';
          }
        }
      }
    }
  } catch (err) {
    console.error('Error in verificarYGenerarCuotasMensualesNativo:', err);
  }
}

async function obtenerDatosSocio(email) {
  try {
    await verificarYGenerarCuotasMensualesNativo();
    const { data: users, error: errU } = await supabase.from('usuarios').select('*').eq('email', email);
    if (errU || !users || users.length === 0) throw new Error('Usuario no encontrado');
    const u = users[0];
    
    const { data: pagos, error: errP } = await supabase.from('pagos').select('*').eq('email', email).order('month', { ascending: true });
    
    const obj = {
      Email: u.email, Role: u.role, Name: u.name, Phone: u.phone, Category: u.category,
      DNI: u.dni, Birthdate: u.birthdate, Age: u.age, JoinDate: u.joindate, BloodType: u.bloodtype,
      MedicalFit: u.medicalfit, ObraSocial: u.obrasocial, EmergencyContact: u.emergencycontact,
      EmergencyPhone: u.emergencyphone, ParentName: u.parentname, ParentPhone: u.parentphone,
      Notes: u.notes, Photo: u.photo, Pagos: []
    };
    
    if (pagos && pagos.length > 0) {
      obj.Pagos = pagos.map(p => ({
        Payment_ID: p.payment_id, Month: p.month, Amount: p.amount, Status: p.status,
        MP_Link: p.mp_link, Collected_By: p.collected_by, Collected_At: p.collected_at
      }));
    }
    
    return obj;
  } catch(e) {
    console.error('obtenerDatosSocio ERROR:', e);
    throw e;
  }
}

async function obtenerDatosAdmin() {
  try {
    await verificarYGenerarCuotasMensualesNativo();
    
    const [
      { data: users }, { data: pagos }, { data: categorias }, 
      { data: torneos }, { data: finanzas }, { data: partidos }
    ] = await Promise.all([
      supabase.from('usuarios').select('*'),
      supabase.from('pagos').select('*'),
      supabase.from('categorias').select('*'),
      supabase.from('torneos').select('*'),
      supabase.from('finanzas_torneos').select('*'),
      supabase.from('partidos').select('*')
    ]);
    
    const sociosData = (users || []).filter(u => u.role === 'Deportista' && u.email);
    let recaudacionTotal = 0;
    let deudaPendiente = 0;
    
    const listaSocios = sociosData.map(u => ({
      Email: u.email,
      Name: u.name || '',
      Category: u.category || '',
      Role: u.role,
      Phone: u.phone,
      DNI: u.dni,
      Status: 'Activo',
      Deuda: 0,
      Username: u.username
    }));
    
    const todosLosPagos = (pagos || []).filter(p => p.payment_id && p.email).map(p => {
      const socio = listaSocios.find(s => s.Email.toLowerCase() === p.email.toLowerCase());
      const amount = parseFloat(p.amount) || 0;
      
      if (socio) {
        if (p.status === 'Pagado') recaudacionTotal += amount;
        else if (p.status === 'Deuda' || p.status === 'Pendiente') {
          deudaPendiente += amount;
          socio.Deuda += amount;
        }
      }
      
      return {
        Payment_ID: p.payment_id,
        Email: p.email,
        SocioName: socio ? (socio.Username || socio.Name) : 'Socio Desconocido',
        Month: p.month,
        Amount: amount,
        Status: p.status,
        MP_Link: p.mp_link,
        Collected_By: p.collected_by,
        Collected_At: p.collected_at
      };
    });
    
    const listaDeudores = listaSocios.filter(s => s.Deuda > 0);
    
    const categoriasOut = (categorias || []).map(c => ({
      Category_ID: c.category_id,
      Name: c.name,
      Coach: c.coach,
      Monthly_Fee: c.monthly_fee || c.price_monthly || 0,
      Torneos: c.torneos
    }));
    
    const torneosOut = (torneos || []).map(t => ({
      Torneo_ID: t.torneo_id, Name: t.name, Category: t.category, Year: t.year, Is_Active: t.is_active
    }));
    
    const finanzasOut = (finanzas || []).map(f => {
      const t = torneosOut.find(x => x.Torneo_ID === f.torneo_id);
      return {
        Movimiento_ID: f.movimiento_id,
        Torneo_ID: f.torneo_id,
        Torneo_Name: t ? t.Name : 'General',
        Type: f.type,
        Concept: f.concept,
        Amount: parseFloat(f.amount) || 0,
        Date: f.date
      };
    }).sort((a,b) => (b.Date || '').localeCompare(a.Date || ''));
    
    const partidosOut = (partidos || []).map(p => {
      const t = torneosOut.find(x => x.Torneo_ID === p.torneo_id);
      return {
        Partido_ID: p.partido_id,
        Torneo_ID: p.torneo_id,
        Torneo_Name: t ? t.Name : 'Desconocido',
        Date: p.date,
        Opponent: p.opponent,
        Location: p.is_home === 'Sí' ? 'Local' : 'Visitante',
        Result: p.goals_for + ' - ' + p.goals_against,
        Goals_For: p.goals_for,
        Goals_Against: p.goals_against,
        Status: p.status
      };
    });
    
    return {
      success: true,
      metrics: {
        totalSocios: listaSocios.length,
        recaudacionTotal,
        deudaPendiente
      },
      socios: listaSocios,
      deudores: listaDeudores,
      pagos: todosLosPagos,
      torneos: torneosOut,
      categorias: categoriasOut,
      finanzas: finanzasOut,
      partidos: partidosOut
    };
  } catch(e) {
    console.error('obtenerDatosAdmin ERROR:', e);
    throw e;
  }
}

module.exports = {
  obtenerDatosSocio,
  obtenerDatosAdmin
};
