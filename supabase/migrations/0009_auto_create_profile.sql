-- Every authenticated user needs a profiles row (pseudonymous handle, role
-- defaulting to 'user') before they can submit or vote on edits. Rather than
-- have application code race to create it on first use, create it the
-- moment auth.users gains a row — the standard Supabase pattern.
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer set search_path = public
as $$
begin
  insert into public.profiles (id, handle)
  values (new.id, 'user-' || substr(replace(new.id::text, '-', ''), 1, 8))
  on conflict (id) do nothing;
  return new;
end;
$$;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();
