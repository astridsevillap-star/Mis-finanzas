import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://hhvdghksbfpdzggxtukr.supabase.co";
const supabasePublishableKey = "sb_publishable_A_W60FMC96ddG04BnqEMCw_Etyc8mvE";

export const supabase = createClient(supabaseUrl, supabasePublishableKey);
