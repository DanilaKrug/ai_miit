-- =====================================================================
-- Приём сообщения о вагонной операции и обновление дислокации.
-- Вызывается обработчиком очереди для каждой строки msg_inbox со статусом NEW.
-- Автор: отдел сопровождения, 2021-03. Последняя правка: 2021-04-20 (идемпотентность).
-- =====================================================================
CREATE OR REPLACE PROCEDURE proc_accept_disloc_message(p_msg_id bigint)
LANGUAGE plpgsql
AS $$
DECLARE
    v_payload     jsonb;
    v_carnum      bigint;
    v_oper_code   integer;
    v_st_code     integer;
    v_operdate    timestamp;
    v_src_msg_id  text;
    v_is_terminal boolean;
    v_prev_date   timestamp;
    v_days        numeric(8,1);
BEGIN
    -- 1. Берём сообщение и блокируем строку, чтобы два воркера не обработали его дважды
    SELECT payload INTO v_payload
      FROM msg_inbox
     WHERE msg_id = p_msg_id
       AND status = 'NEW'
       FOR UPDATE SKIP LOCKED;

    IF NOT FOUND THEN
        RAISE NOTICE 'Сообщение % уже обработано или заблокировано', p_msg_id;
        RETURN;
    END IF;

    -- 2. Разбор полей
    v_carnum     := (v_payload ->> 'carnum')::bigint;
    v_oper_code  := (v_payload ->> 'oper_code')::integer;
    v_st_code    := (v_payload ->> 'st_code')::integer;
    v_operdate   := (v_payload ->> 'operdate')::timestamp;
    v_src_msg_id := v_payload ->> 'src_msg_id';

    -- 3. Валидация
    IF v_carnum IS NULL OR v_carnum < 10000000 OR v_carnum > 99999999 THEN
        RAISE EXCEPTION 'Некорректный номер вагона: %', v_carnum
            USING ERRCODE = 'check_violation';
    END IF;

    IF v_src_msg_id IS NULL THEN
        RAISE EXCEPTION 'Нет идентификатора исходного сообщения' USING ERRCODE = 'not_null_violation';
    END IF;

    IF NOT EXISTS (SELECT 1 FROM ref_station WHERE st_code = v_st_code) THEN
        RAISE EXCEPTION 'Станция % отсутствует в справочнике', v_st_code
            USING ERRCODE = 'foreign_key_violation';
    END IF;

    SELECT is_terminal INTO v_is_terminal
      FROM ref_operation
     WHERE oper_code = v_oper_code;

    IF v_is_terminal IS NULL THEN
        RAISE EXCEPTION 'Неизвестная операция %', v_oper_code USING ERRCODE = 'foreign_key_violation';
    END IF;

    -- 4. Идемпотентность: повтор сообщения из внешней системы не должен менять дислокацию
    IF EXISTS (SELECT 1 FROM wagon_oper_log WHERE src_msg_id = v_src_msg_id) THEN
        UPDATE msg_inbox SET status = 'SKIPPED', error_text = 'duplicate src_msg_id'
         WHERE msg_id = p_msg_id;
        RETURN;
    END IF;

    -- 5. Опоздавшее сообщение: операция старее текущей дислокации — пишем в журнал, но дислокацию не трогаем
    SELECT operdate INTO v_prev_date FROM wagon_disloc WHERE carnum = v_carnum;

    INSERT INTO wagon_oper_log (carnum, oper_code, operdate, st_code, src_msg_id, msg_id)
    VALUES (v_carnum, v_oper_code, v_operdate, v_st_code, v_src_msg_id, p_msg_id);

    IF v_prev_date IS NOT NULL AND v_operdate < v_prev_date THEN
        UPDATE msg_inbox SET status = 'DONE', error_text = 'late message, dislocation kept'
         WHERE msg_id = p_msg_id;
        RETURN;
    END IF;

    -- 6. Обновление дислокации. days_disl считается от момента операции до сейчас
    v_days := round(extract(epoch FROM (now() - v_operdate)) / 86400.0, 1);

    INSERT INTO wagon_disloc (carnum, st_disl_code, oper_code, operdate, days_disl, updated_at)
    VALUES (v_carnum, v_st_code, v_oper_code, v_operdate, v_days, now())
    ON CONFLICT (carnum) DO UPDATE
       SET st_disl_code = EXCLUDED.st_disl_code,
           oper_code    = EXCLUDED.oper_code,
           operdate     = EXCLUDED.operdate,
           days_disl    = EXCLUDED.days_disl,
           updated_at   = now();

    -- 7. Терминальная операция (выгрузка, подача под выгрузку) освобождает вагон в парке
    IF v_is_terminal THEN
        UPDATE wagon_park
           SET state = 'FREE', st_code = v_st_code
         WHERE carnum = v_carnum;
    END IF;

    UPDATE msg_inbox SET status = 'DONE', error_text = NULL WHERE msg_id = p_msg_id;

EXCEPTION
    WHEN check_violation OR not_null_violation OR foreign_key_violation THEN
        UPDATE msg_inbox SET status = 'ERROR', error_text = SQLERRM WHERE msg_id = p_msg_id;
    WHEN OTHERS THEN
        UPDATE msg_inbox SET status = 'ERROR', error_text = 'unexpected: ' || SQLERRM
         WHERE msg_id = p_msg_id;
        RAISE WARNING 'proc_accept_disloc_message(%): %', p_msg_id, SQLERRM;
END;
$$;
