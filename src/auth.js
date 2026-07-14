const jwt = require('jsonwebtoken');
const supabase = require('./supabase');

const JWT_SECRET = process.env.JWT_SECRET || 'HaedoFutsalSuperSecretKey2026';

// Middleware para proteger rutas
const authenticateToken = (req, res, next) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];
  
  if (!token) return res.status(401).json({ success: false, message: 'Token requerido' });
  
  jwt.verify(token, JWT_SECRET, (err, user) => {
    if (err) return res.status(403).json({ success: false, message: 'Token invlido o expirado' });
    req.user = user;
    next();
  });
};

const loginHandler = async (req, res) => {
  const { username, password } = req.body || {};
  if (!username) return res.status(400).json({ success: false, message: 'Falta usuario o DNI.' });
  
  try {
    const target = username.toString().trim().toLowerCase();
    
    // Buscar por DNI, Username o Email
    const { data: users, error } = await supabase
      .from('usuarios')
      .select('id, email, role, name, photo, username, dni, password')
      .or(`username.ilike.${target},dni.eq.${target},email.ilike.${target}`);
      
    if (error) {
      console.error('[LOGIN DB ERROR]', error.message);
      return res.status(500).json({ success: false, message: 'Error en la base de datos.' });
    }
    
    const user = users && users.length > 0 ? users[0] : null;
    
    if (!user) {
      return res.json({ success: false, message: 'Usuario no encontrado en el sistema.' });
    }
    
    const expectedPassword = (user.password || '1234').toString().trim();
    if (password !== null && password !== undefined && password.toString().trim() !== expectedPassword) {
      return res.json({ success: false, message: 'Clave incorrecta.' });
    }
    
    // Generar JWT Token
    const payload = {
      id: user.id,
      email: user.email,
      role: user.role,
      name: user.name,
      username: user.username,
      dni: user.dni
    };
    
    const token = jwt.sign(payload, JWT_SECRET, { expiresIn: '7d' });
    
    return res.json({
      success: true,
      token,
      email: user.email,
      role: user.role || 'Deportista',
      name: user.name || '',
      username: user.username || '',
      dni: user.dni || '',
      photo: user.photo || ''
    });
  } catch (err) {
    console.error('[LOGIN ERROR]', err.message);
    res.status(500).json({ success: false, message: 'Error interno del servidor.' });
  }
};

module.exports = { authenticateToken, loginHandler, JWT_SECRET };
