import { createClient } from '@supabase/supabase-js';

// Thay thế bằng URL và Anon Key của thầy
const supabaseUrl = 'https://fiuwpvdqkyyatxmuwhja.supabase.co'; 
const supabaseAnonKey = 'sb_publishable_oIxLnyWObJwcUpgTCdJ_EQ_dKr6RXI8';

export const supabase = createClient(supabaseUrl, supabaseAnonKey);