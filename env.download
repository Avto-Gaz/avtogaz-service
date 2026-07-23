import { createClient } from "@supabase/supabase-js";

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
const supabaseKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

export const supabase = createClient(supabaseUrl, supabaseKey);

const ROW_ID = "main";

// window.storage o'rniga ishlaydigan wrapper — bir xil interfeys
export const storage = {
  async get(key) {
    try {
      const { data, error } = await supabase
        .from("app_data")
        .select("data")
        .eq("id", ROW_ID)
        .single();
      if (error) throw error;
      const value = data?.data?.[key];
      if (value === undefined) return null;
      return { value: JSON.stringify(value) };
    } catch (e) {
      console.error("Supabase get xatosi:", e);
      throw e;
    }
  },

  async set(key, valueStr) {
    try {
      const { data: existing } = await supabase
        .from("app_data")
        .select("data")
        .eq("id", ROW_ID)
        .single();

      const current = existing?.data || {};
      const updated = { ...current, [key]: JSON.parse(valueStr) };

      const { error } = await supabase
        .from("app_data")
        .upsert({ id: ROW_ID, data: updated, updated_at: new Date().toISOString() });

      if (error) throw error;
      return { key, value: valueStr };
    } catch (e) {
      console.error("Supabase set xatosi:", e);
      throw e;
    }
  },
};
