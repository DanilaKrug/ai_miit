CREATE OR REPLACE FUNCTION fn_register_gu12_order(
    p_client_id   bigint,
    p_st_code     integer,
    p_rsv_st_code integer,
    p_rod_id      integer,
    p_gg_name     text,
    p_wag_cnt     integer,
    p_load_date   date
) RETURNS bigint
LANGUAGE plpgsql
AS $$
DECLARE
    v_order_id bigint;
    v_active   boolean;
BEGIN
    SELECT is_active INTO v_active FROM ref_client WHERE client_id = p_client_id;
    IF v_active IS NULL THEN
        RAISE EXCEPTION 'Клиент % не найден', p_client_id USING ERRCODE = 'P0001';
    ELSIF NOT v_active THEN
        RAISE EXCEPTION 'Клиент % заблокирован', p_client_id USING ERRCODE = 'P0002';
    END IF;
    IF p_st_code = p_rsv_st_code THEN
        RAISE EXCEPTION 'Станции отправления и назначения совпадают' USING ERRCODE = 'P0003';
    END IF;
    IF NOT EXISTS (SELECT 1 FROM ref_station WHERE st_code = p_st_code)
       OR NOT EXISTS (SELECT 1 FROM ref_station WHERE st_code = p_rsv_st_code) THEN
        RAISE EXCEPTION 'Станция не найдена: % или %', p_st_code, p_rsv_st_code USING ERRCODE = 'P0004';
    END IF;
    IF NOT EXISTS (SELECT 1 FROM ref_rod WHERE rod_id = p_rod_id) THEN
        RAISE EXCEPTION 'Род подвижного состава % не найден', p_rod_id USING ERRCODE = 'P0005';
    END IF;
    IF p_wag_cnt IS NULL OR p_wag_cnt < 1 OR p_wag_cnt > 500 THEN
        RAISE EXCEPTION 'Количество вагонов вне диапазона 1..500: %', p_wag_cnt USING ERRCODE = 'P0006';
    END IF;
    IF p_load_date < current_date THEN
        RAISE EXCEPTION 'Дата погрузки % в прошлом', p_load_date USING ERRCODE = 'P0007';
    END IF;
    IF coalesce(btrim(p_gg_name), '') = '' THEN
        RAISE EXCEPTION 'Не указана группа груза' USING ERRCODE = 'P0008';
    END IF;
    SELECT order_id INTO v_order_id
      FROM gu12_order
     WHERE client_id = p_client_id AND st_code = p_st_code AND rsv_st_code = p_rsv_st_code
       AND rod_id = p_rod_id AND load_date = p_load_date AND status NOT IN ('CANCELLED', 'DONE');
    IF v_order_id IS NOT NULL THEN
        RAISE EXCEPTION 'Дубль заявки %', v_order_id USING ERRCODE = 'unique_violation';
    END IF;
    INSERT INTO gu12_order (client_id, st_code, rsv_st_code, rod_id, gg_name, gu12_wag_cnt, load_date, status)
    VALUES (p_client_id, p_st_code, p_rsv_st_code, p_rod_id, upper(btrim(p_gg_name)), p_wag_cnt, p_load_date,
            CASE WHEN p_wag_cnt <= 50 THEN 'CONFIRMED' ELSE 'NEW' END)
    RETURNING order_id INTO v_order_id;
    INSERT INTO gu12_order_hist (order_id, old_status, new_status, reason)
    VALUES (v_order_id, NULL, CASE WHEN p_wag_cnt <= 50 THEN 'CONFIRMED' ELSE 'NEW' END,
            CASE WHEN p_wag_cnt <= 50 THEN 'auto-confirm <= 50' ELSE 'manual confirmation required' END);
    IF p_wag_cnt > 50 THEN
        PERFORM fn_notify_dispatcher(v_order_id, 'Заявка требует подтверждения');
    END IF;
    RETURN v_order_id;
END;
$$;
