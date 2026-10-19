-- Расчёт суточного плана подач вагонов под подтверждённые заявки.
-- Запускается ночью планировщиком на дату p_plan_date; p_st_code ограничивает расчёт одной станцией.
CREATE OR REPLACE PROCEDURE proc_build_daily_supply_plan(p_plan_date date, p_st_code integer DEFAULT NULL)
LANGUAGE plpgsql
AS $$
DECLARE
    r_order     record;
    v_available integer;
    v_take      integer;
    v_planned   integer := 0;
    v_waiting   integer := 0;
    v_deleted   bigint;
    -- заявки берём в порядке приоритета клиента
    c_orders CURSOR FOR
        SELECT o.order_id, o.st_code, o.rod_id, o.gu12_wag_cnt, o.client_id
          FROM gu12_order o
         WHERE o.status = 'CONFIRMED'
           AND o.load_date = p_plan_date
           AND (p_st_code IS NULL OR o.st_code = p_st_code)
         ORDER BY o.created_at;
BEGIN
    IF p_plan_date < current_date THEN
        RAISE EXCEPTION 'План на прошедшую дату % не строится', p_plan_date;
    END IF;

    -- Пересчёт: старый план на эту дату (и станцию) удаляем
    DELETE FROM daily_supply_plan
     WHERE plan_date = p_plan_date
       AND (p_st_code IS NULL OR st_code = p_st_code);
    GET DIAGNOSTICS v_deleted = ROW_COUNT;
    RAISE NOTICE 'Удалено строк старого плана: %', v_deleted;

    -- Свободный парк по станциям и родам
    CREATE TEMP TABLE tmp_free_park ON COMMIT DROP AS
    SELECT st_code, rod_id, count(*)::integer AS free_cnt
      FROM wagon_park
     WHERE state = 'FREE'
       AND st_code IS NOT NULL
       AND (p_st_code IS NULL OR st_code = p_st_code)
     GROUP BY st_code, rod_id;

    FOR r_order IN c_orders LOOP
        SELECT free_cnt INTO v_available
          FROM tmp_free_park
         WHERE st_code = r_order.st_code AND rod_id = r_order.rod_id;

        v_available := coalesce(v_available, 0);

        IF v_available <= 0 THEN
            -- вагонов нужного рода на станции нет: заявка ждёт
            UPDATE gu12_order SET status = 'WAITING', updated_at = now()
             WHERE order_id = r_order.order_id;
            INSERT INTO gu12_order_hist (order_id, old_status, new_status, reason)
            VALUES (r_order.order_id, 'CONFIRMED', 'WAITING', 'no free wagons at station');
            v_waiting := v_waiting + 1;
            CONTINUE;
        END IF;

        v_take := least(v_available, r_order.gu12_wag_cnt);

        INSERT INTO daily_supply_plan (plan_date, order_id, st_code, rod_id, planned_wagons)
        VALUES (p_plan_date, r_order.order_id, r_order.st_code, r_order.rod_id, v_take);

        UPDATE tmp_free_park SET free_cnt = free_cnt - v_take
         WHERE st_code = r_order.st_code AND rod_id = r_order.rod_id;

        -- частичное покрытие остаётся в статусе PLANNED, остаток заявки допланируется на следующий день
        UPDATE gu12_order SET status = 'PLANNED', updated_at = now()
         WHERE order_id = r_order.order_id;
        INSERT INTO gu12_order_hist (order_id, old_status, new_status, reason)
        VALUES (r_order.order_id, 'CONFIRMED', 'PLANNED',
                CASE WHEN v_take < r_order.gu12_wag_cnt THEN 'partial: ' || v_take || ' of ' || r_order.gu12_wag_cnt
                     ELSE 'full' END);
        v_planned := v_planned + 1;
    END LOOP;

    RAISE NOTICE 'План на %: запланировано заявок %, ожидают %', p_plan_date, v_planned, v_waiting;
END;
$$;
