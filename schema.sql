-- schema.sql
DROP TABLE IF EXISTS Reparaciones;
DROP TABLE IF EXISTS Tecnicos;
DROP TABLE IF EXISTS EstadosCatalogo;
DROP TABLE IF EXISTS EquiposCatalogo;
DROP TABLE IF EXISTS SalasCatalogo;

CREATE TABLE Tecnicos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT UNIQUE NOT NULL
);

CREATE TABLE EstadosCatalogo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT UNIQUE NOT NULL
);

CREATE TABLE EquiposCatalogo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT UNIQUE NOT NULL,
    valor TEXT DEFAULT ''
);

CREATE TABLE SalasCatalogo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT UNIQUE NOT NULL
);

CREATE TABLE Reparaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha TEXT NOT NULL,
    sala TEXT NOT NULL,
    uid TEXT,
    npu TEXT,
    parte TEXT,
    familia TEXT,
    equipo TEXT NOT NULL,
    urgente TEXT CHECK(urgente IN ('SI', 'NO')) DEFAULT 'NO',
    tecnico_id INTEGER,
    estado TEXT DEFAULT 'PEND. DE REVISION',
    observaciones TEXT,
    FOREIGN KEY (tecnico_id) REFERENCES Tecnicos (id)
);
