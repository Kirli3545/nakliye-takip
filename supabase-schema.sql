-- Bu SQL kodunu Supabase panelinde "SQL Editor" bölümüne yapıştırıp çalıştırın.

create table if not exists kv_store (
  key text primary key,
  value text,
  updated_at timestamptz default now()
);

alter table kv_store enable row level security;

-- Not: Bu policy, sadece GİRİŞ YAPMIŞ (Supabase Auth ile doğrulanmış)
-- kullanıcıların okuma/yazma yapmasına izin verir. Giriş yapmayan biri
-- (uygulamanın login ekranını atlayıp API'ye direkt istek atsa bile)
-- veri göremez/değiştiremez.
create policy "Sadece giriş yapmış kullanıcılar"
on kv_store for all
using (auth.role() = 'authenticated')
with check (auth.role() = 'authenticated');
