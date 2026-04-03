-- seed_hash_validation.sql
-- Fixture: HASH_EDGE_CASES
-- Proyecto: CONDOR
-- Cleanup: cleanup_hash_validation.sql
-- IDs reservados: 99101 - 99110

-- Caso 1: Hash NULL (debe fallar validacion)
INSERT INTO TbSolicitudes (ID, Referencia, Estado, HashValidacion)
VALUES (99101, 'EDGE_HASH_NULL', 'Borrador', NULL);

-- Caso 2: Hash vacio (string vacio, distinto de NULL)
INSERT INTO TbSolicitudes (ID, Referencia, Estado, HashValidacion)
VALUES (99102, 'EDGE_HASH_EMPTY', 'Borrador', '');

-- Caso 3: Hash valido estandar
INSERT INTO TbSolicitudes (ID, Referencia, Estado, HashValidacion)
VALUES (99103, 'EDGE_HASH_VALID', 'Validado', 'a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4');

-- Caso 4: Hash con longitud incorrecta
INSERT INTO TbSolicitudes (ID, Referencia, Estado, HashValidacion)
VALUES (99104, 'EDGE_HASH_SHORT', 'Borrador', 'abc123');

-- Caso 5: Hash con caracteres no hex
INSERT INTO TbSolicitudes (ID, Referencia, Estado, HashValidacion)
VALUES (99105, 'EDGE_HASH_BADHEX', 'Borrador', 'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ');

-- Caso 6: Observaciones con punto y coma (prueba del parser)
INSERT INTO TbSolicitudes (ID, Referencia, Estado, Observaciones)
VALUES (99106, 'EDGE_SEMICOLON', 'Borrador', 'Texto con; punto y coma; interno');

-- Verificacion
SELECT ID, Referencia, Estado, HashValidacion
FROM TbSolicitudes
WHERE ID BETWEEN 99101 AND 99110
