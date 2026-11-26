create table public.career_profiles (
  id uuid primary key default gen_random_uuid(),
  name text not null check (char_length(trim(name)) between 1 and 100),
  title text not null check (char_length(trim(title)) between 1 and 120),
  location text not null check (char_length(trim(location)) between 1 and 100),
  work_type text not null check (work_type in ('On-site', 'Hybrid', 'Remote')),
  skills text[] not null check (cardinality(skills) between 1 and 15 and char_length(array_to_string(skills, ',')) <= 300),
  summary text not null check (char_length(trim(summary)) between 1 and 800),
  email text not null check (char_length(trim(email)) between 3 and 254),
  linkedin text check (linkedin is null or (char_length(linkedin) <= 300 and linkedin ~* '^https://([a-z0-9-]+\.)*linkedin\.com/')),
  approved boolean not null default false,
  created_at timestamptz not null default now()
);

create table public.career_endorsements (
  id uuid primary key default gen_random_uuid(),
  candidate_id uuid not null references public.career_profiles(id) on delete cascade,
  author text not null check (char_length(trim(author)) between 1 and 100),
  relationship text not null check (char_length(trim(relationship)) between 1 and 100),
  message text not null check (char_length(trim(message)) between 1 and 600),
  approved boolean not null default false,
  created_at timestamptz not null default now()
);

alter table public.career_profiles enable row level security;
alter table public.career_endorsements enable row level security;

revoke all on public.career_profiles, public.career_endorsements from public, anon, authenticated;
grant insert (id, name, title, location, work_type, skills, summary, email, linkedin) on public.career_profiles to anon, authenticated;
grant insert (candidate_id, author, relationship, message) on public.career_endorsements to anon, authenticated;

create policy "Submit profiles for review" on public.career_profiles for insert to anon, authenticated with check (approved = false);
create policy "Submit recommendations for review" on public.career_endorsements for insert to anon, authenticated with check (approved = false);

create view public.career_public_profiles as
  select id, name, title, location, work_type, skills, summary, linkedin
  from public.career_profiles where approved = true;

create view public.career_public_endorsements as
  select endorsement.candidate_id, endorsement.author, endorsement.relationship, endorsement.message
  from public.career_endorsements endorsement
  join public.career_profiles profile on profile.id = endorsement.candidate_id
  where endorsement.approved = true and profile.approved = true;

revoke all on public.career_public_profiles, public.career_public_endorsements from public, anon, authenticated;
grant select on public.career_public_profiles, public.career_public_endorsements to anon, authenticated;
