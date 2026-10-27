-- 4분류 전환 후 하부 구분 값이 '외부' → '협력점' (이전 저장분 '외부'도 그대로 허용)
alter table public.dist_pnl_lines drop constraint if exists dist_pnl_lines_partner_type_check;
alter table public.dist_pnl_lines add constraint dist_pnl_lines_partner_type_check check (partner_type = any (array['자점','협력점','외부']));
