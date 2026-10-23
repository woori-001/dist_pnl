-- 구분별 가정산/확정 (같은 달·같은 구분 확정은 1건만)
alter table public.dist_pnl_runs add column if not exists seg_status jsonb not null default '{}'::jsonb;
comment on column public.dist_pnl_runs.seg_status is '구분별 상태 {"유통이관":"확정","SKB웨이브넷":"가정산",...}';
update public.dist_pnl_runs
   set seg_status = jsonb_build_object('유통이관', status, 'SKB웨이브넷', status, 'SKY우리정보', status, 'SK매직', status)
 where seg_status = '{}'::jsonb;

create or replace function public.dist_pnl_derive_status(p_period date)
returns void language plpgsql security invoker set search_path = ''
as $$
begin
  update public.dist_pnl_runs r set status = '가정산', confirmed_by = null, confirmed_at = null
   where r.period = p_period and r.status = '확정'
     and not (r.seg_status->>'유통이관' = '확정' and r.seg_status->>'SKB웨이브넷' = '확정' and r.seg_status->>'SKY우리정보' = '확정' and r.seg_status->>'SK매직' = '확정');
  update public.dist_pnl_runs r set status = '확정', confirmed_by = coalesce(r.confirmed_by, (select auth.jwt()) ->> 'email'), confirmed_at = coalesce(r.confirmed_at, now())
   where r.period = p_period and r.status = '가정산'
     and (r.seg_status->>'유통이관' = '확정' and r.seg_status->>'SKB웨이브넷' = '확정' and r.seg_status->>'SKY우리정보' = '확정' and r.seg_status->>'SK매직' = '확정');
end;
$$;

create or replace function public.dist_pnl_save_run2(
  p_period date, p_seg_status jsonb, p_note text, p_files jsonb, p_summary jsonb, p_rules jsonb, p_lines jsonb)
returns uuid language plpgsql security invoker set search_path = ''
as $$
declare v_id uuid; v_seg text;
begin
  if not public.dist_pnl_is_member() then raise exception '권한이 없습니다'; end if;
  for v_seg in select key from jsonb_each_text(p_seg_status) where value = '확정' loop
    update public.dist_pnl_runs set seg_status = jsonb_set(seg_status, array[v_seg], '"가정산"')
     where period = p_period and seg_status->>v_seg = '확정';
  end loop;
  perform public.dist_pnl_derive_status(p_period);
  insert into public.dist_pnl_runs (period, status, seg_status, note, files, summary, rules)
  values (p_period, '가정산', p_seg_status, p_note, coalesce(p_files,'[]'::jsonb), p_summary, coalesce(p_rules,'{}'::jsonb))
  returning id into v_id;
  insert into public.dist_pnl_lines (run_id, segment, upper_name, side, item, partner_type, included, cnt, amount, gift_cash, gift_goods, gift_coupon)
  select v_id, l.segment, l.upper_name, l.side, l.item, l.partner_type, l.included, l.cnt, l.amount, l.gift_cash, l.gift_goods, l.gift_coupon
    from jsonb_to_recordset(p_lines) as l(segment text, upper_name text, side text, item text, partner_type text,
         included boolean, cnt integer, amount bigint, gift_cash bigint, gift_goods bigint, gift_coupon bigint);
  perform public.dist_pnl_derive_status(p_period);
  return v_id;
end;
$$;

create or replace function public.dist_pnl_set_seg_status(p_run uuid, p_seg text, p_status text)
returns void language plpgsql security invoker set search_path = ''
as $$
declare v_period date;
begin
  if not public.dist_pnl_is_member() then raise exception '권한이 없습니다'; end if;
  if p_status not in ('가정산','확정') then raise exception '상태 값 오류'; end if;
  if p_seg not in ('유통이관','SKB웨이브넷','SKY우리정보','SK매직') then raise exception '구분 값 오류'; end if;
  select period into v_period from public.dist_pnl_runs where id = p_run;
  if v_period is null then raise exception '회차를 찾을 수 없습니다'; end if;
  if p_status = '확정' then
    update public.dist_pnl_runs set seg_status = jsonb_set(seg_status, array[p_seg], '"가정산"')
     where period = v_period and id <> p_run and seg_status->>p_seg = '확정';
    perform public.dist_pnl_derive_status(v_period);
  end if;
  update public.dist_pnl_runs set seg_status = jsonb_set(seg_status, array[p_seg], to_jsonb(p_status)) where id = p_run;
  perform public.dist_pnl_derive_status(v_period);
end;
$$;

revoke all on function public.dist_pnl_derive_status(date) from public, anon;
revoke all on function public.dist_pnl_save_run2(date,jsonb,text,jsonb,jsonb,jsonb,jsonb) from public, anon;
revoke all on function public.dist_pnl_set_seg_status(uuid,text,text) from public, anon;
grant execute on function public.dist_pnl_derive_status(date) to authenticated;
grant execute on function public.dist_pnl_save_run2(date,jsonb,text,jsonb,jsonb,jsonb,jsonb) to authenticated;
grant execute on function public.dist_pnl_set_seg_status(uuid,text,text) to authenticated;
