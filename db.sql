 -- Database: TWK

-- DROP DATABASE "TWK";

CREATE DATABASE "TWK"
    WITH 
    OWNER = postgres
    ENCODING = 'UTF8'
    LC_COLLATE = 'en_US.utf8'
    LC_CTYPE = 'en_US.utf8'
    TABLESPACE = pg_default
    CONNECTION LIMIT = -1;


-- Table: public.sites

-- DROP TABLE public.sites;

CREATE TABLE public.sites
(
    id integer NOT NULL,
    name character varying(30) COLLATE pg_catalog."default" NOT NULL,
    ip character(16) COLLATE pg_catalog."default" NOT NULL,
    CONSTRAINT sites_pkey PRIMARY KEY (id)
)
WITH (
    OIDS = FALSE
)
TABLESPACE pg_default;

ALTER TABLE public.sites
    OWNER to postgres;


-- Table: public.stats

-- DROP TABLE public.stats;

CREATE TABLE public.stats
(
    site_id integer NOT NULL,
    date date NOT NULL,
    tx numeric(6,0) NOT NULL,
    rx numeric(6,0) NOT NULL,
    CONSTRAINT stats_pkey PRIMARY KEY (site_id)
)
WITH (
    OIDS = FALSE
)
TABLESPACE pg_default;

ALTER TABLE public.stats
    OWNER to postgres;

    