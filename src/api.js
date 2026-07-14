const supabase = require('./supabase');

async function obtenerDatosSocio(email) {
  try {
    const { data: users, error: errU } = await supabase.from('usuarios').select('*').eq('email', email);
    if (errU || !users || users.length === 0) throw new Error('Usuario no encontrado');
    const u = users[0];
    
    const { data: pagos, error: errP } = await supabase.from('pagos').select('*').eq('email', email).order('month', { ascending: true });
    
    const obj = {
      Email: u.email,
      Role: u.role,
      Name: u.name,
      Phone: u.phone,
      Category: u.category,
      DNI: u.dni,
      Birthdate: u.birthdate,
      Age: u.age,
      JoinDate: u.joindate,
      BloodType: u.bloodtype,
      MedicalFit: u.medicalfit,
      ObraSocial: u.obrasocial,
      EmergencyContact: u.emergencycontact,
      EmergencyPhone: u.emergencyphone,
      ParentName: u.parentname,
      ParentPhone: u.parentphone,
      Notes: u.notes,
      Photo: u.photo,
      Pagos: []
    };
    
    if (pagos && pagos.length > 0) {
      obj.Pagos = pagos.map(p => ({
        payment_id: p.payment_id,
        month: p.month,
        amount: p.amount,
        status: p.status,
        mp_link: p.mp_link,
        collected_by: p.collected_by,
        collected_at: p.collected_at
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
    const { data: users } = await supabase.from('usuarios').select('*');
    const { data: pagos } = await supabase.from('pagos').select('*');
    const { data: categorias } = await supabase.from('categorias').select('*');
    const { data: torneos } = await supabase.from('torneos').select('*');
    const { data: finanzas } = await supabase.from('finanzas_torneos').select('*');
    const { data: partidos } = await supabase.from('partidos').select('*');
    const { data: logs } = await supabase.from('logs_audit').select('*').order('id', { ascending: false }).limit(100);
    
    return {
      Usuarios: users || [],
      Pagos: pagos || [],
      Categorias: categorias || [],
      Torneos: torneos || [],
      Finanzas_Torneos: finanzas || [],
      Partidos: partidos || [],
      Logs_Audit: logs || []
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
