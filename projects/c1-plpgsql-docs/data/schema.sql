-- Схема учебной базы для проекта С1. Таблицы по мотивам витрин набора курса
-- (data/datamarts.student.json): имена колонок те же, что в витринах, чтобы
-- глоссарий из manifest можно было переиспользовать.

CREATE TABLE ref_road (
    rw_code     integer PRIMARY KEY,
    rw_name     text NOT NULL
);

CREATE TABLE ref_station (
    st_code     integer PRIMARY KEY,
    st_name     text NOT NULL,
    rw_code     integer NOT NULL REFERENCES ref_road (rw_code)
);

CREATE TABLE ref_rod (
    rod_id      integer PRIMARY KEY,
    rod_name    text NOT NULL          -- ПВ, КР, ПЛ, ЦС, ЦМВ
);

CREATE TABLE ref_client (
    client_id   bigint PRIMARY KEY,
    client_name text NOT NULL,
    is_active   boolean NOT NULL DEFAULT true
);

CREATE TABLE ref_operation (
    oper_code   integer PRIMARY KEY,
    oper_name   text NOT NULL,         -- ПРИБЫТИЕ НА СТАНЦИЮ, ПОДАЧА ВАГОНА ПОД ВЫГРУЗКУ ...
    is_terminal boolean NOT NULL DEFAULT false
);

-- Входящие сообщения о вагонных операциях (как приходят из внешней системы)
CREATE TABLE msg_inbox (
    msg_id       bigserial PRIMARY KEY,
    payload      jsonb NOT NULL,        -- {"carnum": 83196971, "oper_code": 12, "st_code": 210038, "operdate": "2021-04-13T12:12:00", "src_msg_id": "A-1"}
    received_at  timestamp NOT NULL DEFAULT now(),
    status       text NOT NULL DEFAULT 'NEW',   -- NEW | DONE | ERROR | SKIPPED
    error_text   text
);

-- Текущая дислокация: одна строка на вагон
CREATE TABLE wagon_disloc (
    carnum       bigint PRIMARY KEY,
    st_disl_code integer REFERENCES ref_station (st_code),
    oper_code    integer REFERENCES ref_operation (oper_code),
    operdate     timestamp NOT NULL,
    loaded       boolean,
    gg_name      text,
    st_dest_code integer REFERENCES ref_station (st_code),
    days_disl    numeric(8,1),
    updated_at   timestamp NOT NULL DEFAULT now()
);

-- Журнал операций по вагонам
CREATE TABLE wagon_oper_log (
    log_id       bigserial PRIMARY KEY,
    carnum       bigint NOT NULL,
    oper_code    integer NOT NULL,
    operdate     timestamp NOT NULL,
    st_code      integer,
    src_msg_id   text NOT NULL UNIQUE,  -- идемпотентность приёма
    msg_id       bigint REFERENCES msg_inbox (msg_id)
);

-- Парк вагонов: где стоит и свободен ли
CREATE TABLE wagon_park (
    carnum       bigint PRIMARY KEY,
    rod_id       integer NOT NULL REFERENCES ref_rod (rod_id),
    c_owner_id   integer NOT NULL,
    st_code      integer REFERENCES ref_station (st_code),
    state        text NOT NULL DEFAULT 'FREE'   -- FREE | ASSIGNED | LOADED | REPAIR
);

-- Заявки ГУ-12
CREATE TABLE gu12_order (
    order_id     bigserial PRIMARY KEY,
    client_id    bigint NOT NULL REFERENCES ref_client (client_id),
    st_code      integer NOT NULL REFERENCES ref_station (st_code),
    rsv_st_code  integer NOT NULL REFERENCES ref_station (st_code),
    rod_id       integer NOT NULL REFERENCES ref_rod (rod_id),
    gg_name      text NOT NULL,
    gu12_wag_cnt integer NOT NULL,
    load_date    date NOT NULL,
    status       text NOT NULL DEFAULT 'NEW',   -- NEW | CONFIRMED | WAITING | PLANNED | DONE | CANCELLED
    created_at   timestamp NOT NULL DEFAULT now(),
    updated_at   timestamp NOT NULL DEFAULT now()
);

CREATE TABLE gu12_order_hist (
    hist_id      bigserial PRIMARY KEY,
    order_id     bigint NOT NULL REFERENCES gu12_order (order_id),
    old_status   text,
    new_status   text NOT NULL,
    changed_at   timestamp NOT NULL DEFAULT now(),
    reason       text
);

-- Суточный план подач
CREATE TABLE daily_supply_plan (
    plan_date       date NOT NULL,
    order_id        bigint NOT NULL REFERENCES gu12_order (order_id),
    st_code         integer NOT NULL,
    rod_id          integer NOT NULL,
    planned_wagons  integer NOT NULL,
    created_at      timestamp NOT NULL DEFAULT now(),
    PRIMARY KEY (plan_date, order_id)
);

-- Факт выгрузки по дням и месячный план (аналоги витрин unload и month_plan)
CREATE TABLE unload_fact (
    repdate          date NOT NULL,
    st_code          integer NOT NULL,
    rw_code          integer NOT NULL,
    rod_id           integer NOT NULL,
    unloaded_wagons  integer NOT NULL
);

CREATE TABLE month_plan (
    rep_month        date NOT NULL,     -- первое число месяца
    st_code          integer NOT NULL,
    rw_code          integer NOT NULL,
    rod_id           integer NOT NULL,
    planned_wagons   integer NOT NULL
);
