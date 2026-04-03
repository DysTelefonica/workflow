-- cleanup_hash_validation.sql
-- Teardown para: seed_hash_validation.sql

DELETE FROM TbSolicitudes WHERE ID BETWEEN 99101 AND 99110;

-- Verificar limpieza
SELECT COUNT(*) AS Restantes FROM TbSolicitudes WHERE ID BETWEEN 99101 AND 99110
