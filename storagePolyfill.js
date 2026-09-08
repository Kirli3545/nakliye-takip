// Bu dosya, prototipte kullanılan window.storage API'sini gerçek bir
// Supabase veritabanına bağlar. App.jsx içinde HİÇBİR DEĞİŞİKLİK YAPMANIZA
// GEREK YOK — window.storage.get/set/delete/list çağrıları aynı şekilde çalışmaya devam eder.

import { supabase } from "./supabaseClient";

async function get(key) {
  const { data, error } = await supabase
    .from("kv_store")
    .select("value")
    .eq("key", key)
    .maybeSingle();
  if (error) {
    console.error("Storage get error:", error);
    throw error;
  }
  if (!data) return null;
  return { key, value: data.value, shared: true };
}

async function set(key, value) {
  const { error } = await supabase
    .from("kv_store")
    .upsert(
      { key, value, updated_at: new Date().toISOString() },
      { onConflict: "key" }
    );
  if (error) {
    console.error("Storage set error:", error);
    return null;
  }
  return { key, value, shared: true };
}

async function del(key) {
  const { error } = await supabase.from("kv_store").delete().eq("key", key);
  if (error) {
    console.error("Storage delete error:", error);
    return null;
  }
  return { key, deleted: true, shared: true };
}

async function list(prefix) {
  let query = supabase.from("kv_store").select("key");
  if (prefix) query = query.like("key", `${prefix}%`);
  const { data, error } = await query;
  if (error) {
    console.error("Storage list error:", error);
    return null;
  }
  return { keys: (data || []).map((r) => r.key), prefix, shared: true };
}

window.storage = { get, set, delete: del, list };
