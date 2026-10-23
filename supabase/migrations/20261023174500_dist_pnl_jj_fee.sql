-- 자점운영비(자점 마진) 집계 칸 + 저장 함수 반영
alter table public.dist_pnl_lines add column if not exists jj_fee bigint not null default 0;
comment on column public.dist_pnl_lines.jj_fee is '자점운영비(자점 마진) 합계 — 코드 매출에서 차감';
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
  insert into public.dist_pnl_lines (run_id, segment, upper_name, side, item, partner_type, included, cnt, amount, gift_cash, gift_goods, gift_coupon, jj_fee)
  select v_id, l.segment, l.upper_name, l.side, l.item, l.partner_type, l.included, l.cnt, l.amount, l.gift_cash, l.gift_goods, l.gift_coupon, coalesce(l.jj_fee, 0)
    from jsonb_to_recordset(p_lines) as l(segment text, upper_name text, side text, item text, partner_type text,
         included boolean, cnt integer, amount bigint, gift_cash bigint, gift_goods bigint, gift_coupon bigint, jj_fee bigint);
  perform public.dist_pnl_derive_status(p_period);
  return v_id;
end;
$$;
