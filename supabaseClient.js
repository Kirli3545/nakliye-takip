import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseKey) {
  console.error(
    "VITE_SUPABASE_URL ve VITE_SUPABASE_ANON_KEY tanımlı değil. .env dosyanızı kontrol edin."
  );
}

export const supabase = createClient(supabaseUrl, supabaseKey);
