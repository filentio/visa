-- Создание базы данных и пользователя
CREATE USER jobparser WITH PASSWORD 'СЮДА_ПАРОЛЬ';
CREATE DATABASE jobparser OWNER jobparser;
GRANT ALL PRIVILEGES ON DATABASE jobparser TO jobparser;

\c jobparser

-- Вакансии
CREATE TABLE vacancies (
    id SERIAL PRIMARY KEY,
    hh_id VARCHAR(50) UNIQUE NOT NULL,
    title VARCHAR(500) NOT NULL,
    company VARCHAR(500),
    salary_from INTEGER,
    salary_to INTEGER,
    salary_currency VARCHAR(10),
    experience VARCHAR(100),
    url VARCHAR(1000),
    description TEXT,
    skills TEXT[],
    city VARCHAR(200),
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- Отклики
CREATE TABLE applications (
    id SERIAL PRIMARY KEY,
    vacancy_id INTEGER REFERENCES vacancies(id),
    status VARCHAR(50) DEFAULT 'sent',  -- sent, viewed, invited, rejected
    cover_letter TEXT,
    applied_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- Анализ навыков рынка
CREATE TABLE market_skills (
    id SERIAL PRIMARY KEY,
    skill VARCHAR(200) NOT NULL,
    count INTEGER DEFAULT 1,
    role VARCHAR(200),
    period DATE DEFAULT CURRENT_DATE,
    UNIQUE(skill, role, period)
);

-- Лог парсера
CREATE TABLE parser_log (
    id SERIAL PRIMARY KEY,
    event VARCHAR(100),
    details TEXT,
    created_at TIMESTAMP DEFAULT NOW()
);

GRANT ALL PRIVILEGES ON ALL TABLES IN SCHEMA public TO jobparser;
GRANT ALL PRIVILEGES ON ALL SEQUENCES IN SCHEMA public TO jobparser;
