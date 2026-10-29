CREATE DATABASE IF NOT EXISTS mineria_facebook_db CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE mineria_facebook_db;

-- 1. Tabla de usuarios (con las mismas credenciales que puso Gus en el login)
CREATE TABLE IF NOT EXISTS usuarios (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nombre VARCHAR(100) NOT NULL,
    correo VARCHAR(120) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL,
    rol VARCHAR(20) DEFAULT 'analista',
    fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 2. Tabla de publicaciones de Facebook analizadas
CREATE TABLE IF NOT EXISTS analisis (
    id INT AUTO_INCREMENT PRIMARY KEY,
    url_post VARCHAR(500) NOT NULL,
    total_likes INT DEFAULT 0,
    total_compartidos INT DEFAULT 0,
    total_comentarios INT DEFAULT 0,
    positivos INT DEFAULT 0,
    negativos INT DEFAULT 0,
    neutros INT DEFAULT 0,
    resumen_ia TEXT,
    fecha_analisis TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 3. Tabla de comentarios analizados por la IA
CREATE TABLE IF NOT EXISTS comentarios (
    id INT AUTO_INCREMENT PRIMARY KEY,
    id_analisis INT NOT NULL,
    usuario VARCHAR(100),
    comentario TEXT NOT NULL,
    sentimiento VARCHAR(20) NOT NULL,
    FOREIGN KEY (id_analisis) REFERENCES analisis(id) ON DELETE CASCADE
);

-- Insertamos el usuario exacto que tiene Gus en su pantalla de Login
INSERT INTO usuarios (nombre, correo, password, rol) 
VALUES ('Administrador', 'admin@empresa.com', '123456', 'admin')
ON DUPLICATE KEY UPDATE id=id;