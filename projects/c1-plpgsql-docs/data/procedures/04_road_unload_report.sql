-- Отчёт по выгрузке дороги за период: факт против месячного плана, приведённого к дням периода.
-- Используется формой «Выгрузка по дороге» и ночной рассылкой.
CREATE OR REPLACE FUNCTION fn_road_unload_report(p_rw_code integer, p_date_from date, p_date_to date)
RETURNS TABLE (
    st_code        integer,
    st_name        text,
    fact_wagons    bigint,
    plan_wagons    numeric,
    deviation_pct  numeric
)
LANGUAGE plpgsql
STABLE
AS $$
BEGIN
    IF p_date_from > p_date_to THEN
        RAISE EXCEPTION 'Начало периода % позже конца %', p_date_from, p_date_to;
    END IF;
    IF p_date_to - p_date_from > 92 THEN
        RAISE EXCEPTION 'Период длиннее 92 дней не поддерживается';
    END IF;
    IF NOT EXISTS (SELECT 1 FROM ref_road r WHERE r.rw_code = p_rw_code) THEN
        RAISE EXCEPTION 'Дорога % не найдена', p_rw_code;
    END IF;

    RETURN QUERY
    WITH days AS (
        SELECT d::date AS d
          FROM generate_series(p_date_from, p_date_to, interval '1 day') AS d
    ),
    fact AS (
        SELECT u.st_code, sum(u.unloaded_wagons)::bigint AS fact_wagons
          FROM unload_fact u
         WHERE u.rw_code = p_rw_code
           AND u.repdate BETWEEN p_date_from AND p_date_to
         GROUP BY u.st_code
    ),
    -- месячный план делим на число дней месяца и берём столько дней, сколько попало в период
    plan AS (
        SELECT m.st_code,
               sum(m.planned_wagons::numeric
                   / extract(day FROM (m.rep_month + interval '1 month - 1 day'))
                   * (SELECT count(*) FROM days
                       WHERE date_trunc('month', days.d) = m.rep_month)) AS plan_wagons
          FROM month_plan m
         WHERE m.rw_code = p_rw_code
           AND m.rep_month BETWEEN date_trunc('month', p_date_from) AND date_trunc('month', p_date_to)
         GROUP BY m.st_code
    )
    SELECT s.st_code,
           s.st_name,
           coalesce(f.fact_wagons, 0)                                   AS fact_wagons,
           round(coalesce(p.plan_wagons, 0), 1)                         AS plan_wagons,
           CASE WHEN coalesce(p.plan_wagons, 0) = 0 THEN NULL
                ELSE round((coalesce(f.fact_wagons, 0) - p.plan_wagons) / p.plan_wagons * 100, 1)
           END                                                          AS deviation_pct
      FROM ref_station s
      LEFT JOIN fact f ON f.st_code = s.st_code
      LEFT JOIN plan p ON p.st_code = s.st_code
     WHERE s.rw_code = p_rw_code
       AND (f.fact_wagons IS NOT NULL OR p.plan_wagons IS NOT NULL)
     ORDER BY deviation_pct NULLS LAST, s.st_name;
END;
$$;
