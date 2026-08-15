-- Reputation is only ever incremented as a side effect of edit approval
-- (server-side, service role), never directly by the user (see the
-- profiles_update_own RLS policy, which pins role = 'user' and says nothing
-- about reputation being user-writable via a normal update anyway — this
-- function is the sanctioned path).
create or replace function increment_reputation(p_user_id uuid, p_amount int)
returns void language sql security definer as $$
  update profiles set reputation = reputation + p_amount where id = p_user_id;
$$;
