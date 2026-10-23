-- 유통사업부 손익매출 계산기 (dist_pnl_*)
-- 원자료(고객명·연락처 등)는 저장하지 않고, 브라우저에서 계산한 "집계값"만 저장한다.

-- 1) 사용 가능한 계정 명단 -----------------------------------------------
create table if not exists public.dist_pnl_members (
  email      text primary key,
  name       text,
  created_at timestamptz not null default now()
);
comment on table public.dist_pnl_members is '유통손익 · 사용 가능 계정(이메일). 여기에 있는 Supabase 계정만 조회/저장 가능';

-- 2) 계산 회차 (월별, 여러 번 저장 가능 / 월당 확정은 1건) --------------
create table if not exists public.dist_pnl_runs (
  id           uuid primary key default gen_random_uuid(),
  period       date not null check (extract(day from period) = 1),
  status       text not null default '가정산' check (status in ('가정산','확정')),
  note         text,
  files        jsonb not null default '[]'::jsonb,
  summary      jsonb not null,
  rules        jsonb not null default '{}'::jsonb,
  created_by   text default (auth.jwt() ->> 'email'),
  created_at   timestamptz not null default now(),
  confirmed_by text,
  confirmed_at timestamptz
);
comment on table public.dist_pnl_runs is '유통손익 · 월별 계산 회차 (가정산/확정, summary = 화면 요약 숫자)';
create unique index if not exists dist_pnl_runs_one_confirmed
  on public.dist_pnl_runs (period) where status = '확정';
create index if not exists dist_pnl_runs_period_idx on public.dist_pnl_runs (period, created_at desc);

-- 3) 회차별 집계 행 (구분 × 상부점 × 상부/전체 × 항목 × 자점여부) -------
create table if not exists public.dist_pnl_lines (
  id           bigint generated always as identity primary key,
  run_id       uuid not null references public.dist_pnl_runs(id) on delete cascade,
  segment      text not null,
  upper_name   text not null,
  side         text not null check (side in ('상부','전체')),
  item         text not null check (item in ('수수료','관리수수료','정산관련','환수')),
  partner_type text not null check (partner_type in ('자점','외부')),
  included     boolean not null,
  cnt          integer not null default 0,
  amount       bigint  not null default 0,
  gift_cash    bigint  not null default 0,
  gift_goods   bigint  not null default 0,
  gift_coupon  bigint  not null default 0
);
comment on table public.dist_pnl_lines is '유통손익 · 회차별 집계 (고객 단위 데이터 없음)';
create index if not exists dist_pnl_lines_run_idx on public.dist_pnl_lines (run_id);

-- 4) 권한 확인 함수 ------------------------------------------------------
create or replace function public.dist_pnl_is_member()
returns boolean
language sql stable security definer
set search_path = ''
as $$
  select exists (
    select 1 from public.dist_pnl_members m
    where lower(m.email) = lower(coalesce((select auth.jwt()) ->> 'email', ''))
  );
$$;
revoke all on function public.dist_pnl_is_member() from public, anon;
grant execute on function public.dist_pnl_is_member() to authenticated;

-- 5) RLS ---------------------------------------------------------------
alter table public.dist_pnl_members enable row level security;
alter table public.dist_pnl_runs    enable row level security;
alter table public.dist_pnl_lines   enable row level security;

revoke all on public.dist_pnl_members, public.dist_pnl_runs, public.dist_pnl_lines from anon;
revoke all on public.dist_pnl_members from authenticated;
grant select on public.dist_pnl_members to authenticated;
grant select, insert, update, delete on public.dist_pnl_runs  to authenticated;
grant select, insert, delete         on public.dist_pnl_lines to authenticated;

drop policy if exists dist_pnl_members_self on public.dist_pnl_members;
create policy dist_pnl_members_self on public.dist_pnl_members
  for select to authenticated
  using (lower(email) = lower(coalesce((select auth.jwt()) ->> 'email', '')));

drop policy if exists dist_pnl_runs_select on public.dist_pnl_runs;
drop policy if exists dist_pnl_runs_insert on public.dist_pnl_runs;
drop policy if exists dist_pnl_runs_update on public.dist_pnl_runs;
drop policy if exists dist_pnl_runs_delete on public.dist_pnl_runs;
create policy dist_pnl_runs_select on public.dist_pnl_runs for select to authenticated using ((select public.dist_pnl_is_member()));
create policy dist_pnl_runs_insert on public.dist_pnl_runs for insert to authenticated with check ((select public.dist_pnl_is_member()));
create policy dist_pnl_runs_update on public.dist_pnl_runs for update to authenticated using ((select public.dist_pnl_is_member())) with check ((select public.dist_pnl_is_member()));
create policy dist_pnl_runs_delete on public.dist_pnl_runs for delete to authenticated using ((select public.dist_pnl_is_member()));

drop policy if exists dist_pnl_lines_select on public.dist_pnl_lines;
drop policy if exists dist_pnl_lines_insert on public.dist_pnl_lines;
drop policy if exists dist_pnl_lines_delete on public.dist_pnl_lines;
create policy dist_pnl_lines_select on public.dist_pnl_lines for select to authenticated using ((select public.dist_pnl_is_member()));
create policy dist_pnl_lines_insert on public.dist_pnl_lines for insert to authenticated with check ((select public.dist_pnl_is_member()));
create policy dist_pnl_lines_delete on public.dist_pnl_lines for delete to authenticated using ((select public.dist_pnl_is_member()));

-- 6) 저장 / 상태변경 함수 (RLS 그대로 적용되는 invoker 함수) -------------
create or replace function public.dist_pnl_save_run(
  p_period date, p_status text, p_note text, p_files jsonb, p_summary jsonb, p_rules jsonb, p_lines jsonb)
returns uuid
language plpgsql security invoker
set search_path = ''
as $$
declare v_id uuid;
begin
  if not public.dist_pnl_is_member() then raise exception '권한이 없습니다'; end if;
  if p_status = '확정' then
    update public.dist_pnl_runs set status = '가정산', confirmed_by = null, confirmed_at = null
     where period = p_period and status = '확정';
  end if;
  insert into public.dist_pnl_runs (period, status, note, files, summary, rules, confirmed_by, confirmed_at)
  values (p_period, p_status, p_note, coalesce(p_files,'[]'::jsonb), p_summary, coalesce(p_rules,'{}'::jsonb),
          case when p_status='확정' then (select auth.jwt()) ->> 'email' end,
          case when p_status='확정' then now() end)
  returning id into v_id;
  insert into public.dist_pnl_lines (run_id, segment, upper_name, side, item, partner_type, included, cnt, amount, gift_cash, gift_goods, gift_coupon)
  select v_id, l.segment, l.upper_name, l.side, l.item, l.partner_type, l.included, l.cnt, l.amount, l.gift_cash, l.gift_goods, l.gift_coupon
    from jsonb_to_recordset(p_lines) as l(segment text, upper_name text, side text, item text, partner_type text,
         included boolean, cnt integer, amount bigint, gift_cash bigint, gift_goods bigint, gift_coupon bigint);
  return v_id;
end;
$$;

create or replace function public.dist_pnl_set_status(p_run uuid, p_status text)
returns void
language plpgsql security invoker
set search_path = ''
as $$
declare v_period date;
begin
  if not public.dist_pnl_is_member() then raise exception '권한이 없습니다'; end if;
  select period into v_period from public.dist_pnl_runs where id = p_run;
  if v_period is null then raise exception '회차를 찾을 수 없습니다'; end if;
  if p_status = '확정' then
    update public.dist_pnl_runs set status = '가정산', confirmed_by = null, confirmed_at = null
     where period = v_period and status = '확정' and id <> p_run;
    update public.dist_pnl_runs set status = '확정', confirmed_by = (select auth.jwt()) ->> 'email', confirmed_at = now()
     where id = p_run;
  else
    update public.dist_pnl_runs set status = '가정산', confirmed_by = null, confirmed_at = null where id = p_run;
  end if;
end;
$$;

revoke all on function public.dist_pnl_save_run(date,text,text,jsonb,jsonb,jsonb,jsonb) from public, anon;
revoke all on function public.dist_pnl_set_status(uuid,text) from public, anon;
grant execute on function public.dist_pnl_save_run(date,text,text,jsonb,jsonb,jsonb,jsonb) to authenticated;
grant execute on function public.dist_pnl_set_status(uuid,text) to authenticated;

-- 7) 첫 사용자 -----------------------------------------------------------
insert into public.dist_pnl_members (email, name) values ('rlaekaals@010king.com', '유통사업부')
on conflict (email) do nothing;
