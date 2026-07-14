const { createClient } = require('@supabase/supabase-js');

const SUPABASE_URL = process.env.SUPABASE_URL || 'https://kjcnotrxxthnzpgljeus.supabase.co';
const SUPABASE_KEY = process.env.SUPABASE_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtqY25vdHJ4eHRobnpwZ2xqZXVzIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MjY5ODY3MywiZXhwIjoyMDk4Mjc0NjczfQ.2yOpoM3C9ejhxA9hY0g88bkyU6KShhSaFHfnaBOLGiU';

// Client con service_role
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: { autoRefreshToken: false, persistSession: false }
});

module.exports = supabase;
