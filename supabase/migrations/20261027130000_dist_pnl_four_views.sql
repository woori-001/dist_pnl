-- 2026-10-27 4분류(코드매출·외부유통매출·유통매출·직영영업매출) 전환
-- 구분 이름을 함수에 박아두지 않고 회차의 seg_status 키를 그대로 씀 (이전 기준 회차도 계속 읽힘)
comment on column public.dist_pnl_runs.seg_status is '구분별 상태 {"코드매출":"확정","외부유통매출":"가정산",...} (10/27 전 회차는 유통이관·SKB웨이브넷·SKY우리정보·SK매직)';

create or replace function public.dist_pnl_derive_status(p_period date)
returns void language plpgsql set search_path = ''
as $$
begin
  -- 모든 구분이 확정이면 회차도 확정, 아니면 가정산 (먼저 내리고 나중에 올려서 고유 인덱스 충돌 방지)
  update public.dist_pnl_runs r set status = '가정산', confirmed_by = null, confirmed_at = null
   where r.period = p_period and r.status = '확정'
     and (r.seg_status = '{}'::jsonb or exists (select 1 from jsonb_each_text(r.seg_status) e where e.value <> '확정'));
  update public.dist_pnl_runs r set status = '확정', confirmed_by = coalesce(r.confirmed_by, (select auth.jwt()) ->> 'email'), confirmed_at = coalesce(r.confirmed_at, now())
   where r.period = p_period and r.status = '가정산'
     and r.seg_status <> '{}'::jsonb and not exists (select 1 from jsonb_each_text(r.seg_status) e where e.value <> '확정')
     and not exists (select 1 from public.dist_pnl_runs o where o.period = p_period and o.id <> r.id and o.status = '확정');
end;
$$;

create or replace function public.dist_pnl_set_seg_status(p_run uuid, p_seg text, p_status text)
returns void language plpgsql set search_path = ''
as $$
declare v_period date; v_st jsonb;
begin
  if not public.dist_pnl_is_member() then raise exception '권한이 없습니다'; end if;
  if p_status not in ('가정산','확정') then raise exception '상태 값 오류'; end if;
  select period, seg_status into v_period, v_st from public.dist_pnl_runs where id = p_run;
  if v_period is null then raise exception '회차를 찾을 수 없습니다'; end if;
  if not (v_st ? p_seg) then raise exception '구분 값 오류'; end if;
  if v_st ? '유통이관' then raise exception '이전 기준 회차는 상태를 바꿀 수 없어요. 다시 계산해서 저장하세요'; end if;
  if p_status = '확정' then
    update public.dist_pnl_runs set seg_status = jsonb_set(seg_status, array[p_seg], '"가정산"')
     where period = v_period and id <> p_run and seg_status->>p_seg = '확정';
    perform public.dist_pnl_derive_status(v_period);
  end if;
  update public.dist_pnl_runs set seg_status = jsonb_set(seg_status, array[p_seg], to_jsonb(p_status)) where id = p_run;
  perform public.dist_pnl_derive_status(v_period);
end;
$$;

-- 이전 기준 회차는 기록으로만 남기고 확정을 풀어 둠 (새 기준 회차가 확정될 수 있도록)
update public.dist_pnl_runs
   set seg_status = (select jsonb_object_agg(k, '가정산') from jsonb_object_keys(seg_status) k),
       status = '가정산', confirmed_by = null, confirmed_at = null
 where seg_status ? '유통이관' and status = '확정' or (seg_status ? '유통이관' and exists (select 1 from jsonb_each_text(seg_status) e where e.value = '확정'));
